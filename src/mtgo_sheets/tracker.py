import argparse
import sqlite3
import xml.etree.ElementTree as ET  # noqa: F401 -- for reset_full_inv()
from bisect import bisect_right
from collections import defaultdict
from pathlib import Path
from statistics import stdev
from string import capwords
from typing import Any

import orjson as json
import xlwings as xw

from mtgo_sheets import util
from mtgo_sheets.config import CONFIG, XlsxFile
from mtgo_sheets.db import get_db_connection
from mtgo_sheets.models import CardDef, CardRow, ExtPriceMap, Transaction

INVALID_EXT_BUY_PRICE = 0
INVALID_EXT_SELL_PRICE = 999
DEFAULT_VENDOR = "goat"

class Logger:
    """Maintain transaction records. Update an xlsx files with new and revised transactions."""

    def __init__(self, path: Path, cols: list[str], excel_app: xw.main.App) -> None:
        self.COL_NAME_TRANSLATION_MAP = {"id": "id_",} # Alt names for cols
        self.path = path
        self.cols = cols
        if not self.path.exists():
            self.wb = xw.Book()
            self.ws = self.wb.sheets[0]
            self.ws.range((1, 1)).value = self.cols
            self.wb.save(path)
        else:
            util.assert_editable(self.path)
            self.wb = excel_app.books.open(self.path)
            self.ws = self.wb.sheets[0]

        self._load_stored_transactions()
        self.transactions: list[Transaction] = []
        self.revisions: list[tuple[int, Transaction]] = []
        self.closed = False

    def add_transaction(self, tx: Transaction) -> None:
        """Insert calculated/metadata fields into a transaction and store them until writeback on close."""
        self._check_closed()
        tx.date = tx.date or CONFIG.today
        tx.price_total = tx.price_each * tx.qty
        if tx.net_each is not None:
            tx.net_total = tx.net_each * tx.qty

        if tx.action == "revision":
            self._revise_transaction(tx)
        else:
            self.transactions.append(tx)

    def get_last_transaction(self, id_: int) -> Transaction | None:
        """Return the last transaction for this id or None."""
        self._check_closed()
        idx = self._get_last_transaction_idx(id_)
        if idx is None:
            return None
        return self.stored_data[idx]

    def close(self) -> None:
        """Commit all changes, then saves and exits the log workbook."""
        self._check_closed()
        for idx, revision in self.revisions:
            rev_dict = revision.model_dump()
            self.ws.range((idx + 2, 1)).value = [
                rev_dict.get(self.COL_NAME_TRANSLATION_MAP.get(col, col), "")
                for col in self.cols
            ]

        next_row = self.ws.range((self.ws.cells.last_cell.row, self.cols.index("id") + 1)).end("up").row + 1
        new_rows = []
        for transaction in self.transactions:
            tx_dict = transaction.model_dump()
            new_rows.append([
                tx_dict.get(self.COL_NAME_TRANSLATION_MAP.get(col, col), "")
                for col in self.cols
            ])
        self.ws.range((next_row, 1)).value = new_rows

        self.wb.save()
        self.wb.close()
        self.closed = True

    def _check_closed(self) -> None:
        if self.closed:
            raise RuntimeError("The logger is closed. Open a new instance for future usage")

    def _load_stored_transactions(self) -> None:
        """Read stored transactions from xlsx file and update self.stored_transcations."""
        last_row = self.ws.range((self.ws.cells.last_cell.row, self.cols.index("id") + 1)).end("up").row
        rows = self.ws.range((2, 1), (last_row, len(self.cols))).value
        self.stored_data = [
            Transaction(
                **{
                    self.COL_NAME_TRANSLATION_MAP.get(key, key): value
                    for key, value in zip(self.cols, row, strict=True)
                }
            )
            for row in rows
            if any(value is not None for value in row)
        ]

    def _get_last_transaction_idx(self, id_: int) -> int | None:
        """Return the index of the last transaction for this id or None.

        Index is equivalent to row (plus the offset for the header row).
        """
        return next(
            (i for i in range(len(self.stored_data) - 1, -1, -1)
            if self.stored_data[i].id_ == id_),
            None,
        )

    def _revise_transaction(self, revision: Transaction) -> None:
        """Replace a transaction that has already been comitted with a revised version."""
        self._check_closed()
        if (idx := self._get_last_transaction_idx(revision.id_)) is None:
            return
        revision.action = revision.old_action or ""

        if revision.notes is None:
            revision.notes = ""
        revision.notes += f" - Revised {CONFIG.today} - original cmd: {revision.old_cmd}"
        self.revisions.append((idx, revision))


class WishlistHandler:
    """Track wishlist additions throughout card entry processing.

    Save all updates to the correct wishlist files for each XlsxFile preference, vendor, and action when done.
    """

    def __init__(self) -> None:
        self.wishlists: dict[Path, list[tuple[int, int]]] = defaultdict(list)

    def add_to_wishlist_path(self, id_: int, qty: int, path: Path | str) -> None:
        """Store data for one wishlist line along with the path it will be written to on flush."""
        self.wishlists[Path(path)].append((id_, qty))

    def add_to_wishlist(self, base_dir: Path, id_: int, action: str, vendor: str | None = None, qty: int = 1) -> None:
        """Determine a wishlist path based on the wishlist directory, action, and vendor.

        Queue the entry for writing when the wishlist is flushed.
        """
        if vendor is not None:
            path = base_dir / f"wishlist_{CONFIG.today}_{action}_{vendor}.dek"
        else:
            path = base_dir / f"wishlist_{CONFIG.today}_{action}.dek"

        self.add_to_wishlist_path(id_, qty, path)

    def flush(self) -> list[Path]:
        """Flush all wishlists to disk. Return all wishlists that were created/modified this way."""
        modified_files = []
        for path, entries in self.wishlists.items():
            if not path.exists():
                path.parent.mkdir(parents=True, exist_ok=True)
                contents = (
                    '<?xml version="1.0" encoding="utf-8"?>\n'
                    '<Deck xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
                    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
                    '  <NetDeckID>0</NetDeckID>\n'
                    '  <PreconstructedDeckID>0</PreconstructedDeckID>\n'
                )
            else:
                contents = path.read_text(encoding="utf-8").replace("</Deck>", "")

            for id_, qty in entries:
                contents += f'  <Cards CatID="{id_:.0f}" Quantity="{qty:.0f}"/>\n'

            contents += "</Deck>"

            path.write_text(contents, encoding="utf-8")
            print(f"Updated {path.name} with {len(entries)} cards")
            modified_files.append(path)

        self.wishlists.clear()
        return modified_files

class WatchlistEntry:
    """Price history / tracking data (alerts, qty owned, etc) for one card / spreadsheet row."""

    def __init__(
        self,
        vals: CardRow,
        card_def: CardDef,
        price_hist: list[float],
        ext_prices: ExtPriceMap | None = None,
    ) -> None:
        """Initialize pricing history stats and tracking metadata."""
        self.active = True

        self.id_ = card_def.id_
        self.name = capwords(card_def.name)
        self.set_code = card_def.set_code.upper()
        self.foil = card_def.foil

        qty = vals.get("qty")
        self.qty = int(qty) if qty is not None else 0
        if self.qty == -1:
            self.active = False # Don't include suggested cards that have no modifications
        self.tag = vals.get("tag", "")
        self.notes = vals.get("notes", "")
        if self.notes is not None and "| failed" in self.notes:
            self.notes = self.notes.split("| failed")[0]
        if self.tag == "auto":
            self.tag = ""

        self.date_created = vals.get("date_created") or CONFIG.today

        # Prices
        self.pending_order = vals.get("order")
        self.price_hist = price_hist
        self.age = len(self.price_hist) - bisect_right(self.price_hist, 0)

        self._set_best_prices(ext_prices)
        self._set_price_data(vals.get("purchase_price"))
        self._set_alerts(vals)

    def buy(self, buy_qty: int, price: float | None = None, extra_notes: str = "") -> Transaction:
        """Update price, ownership, and alert fields to reflect a purchasing event."""
        if price is None:
            price = self.sell_price

        self.qty = max(0, self.qty)
        self.active = True
        total_qty = int(self.qty + buy_qty)
        avg_purchase_price = ((self.purchase_price * self.qty) + (price * buy_qty)) / total_qty
        self.purchase_price = round(avg_purchase_price, 3)
        self.qty = total_qty

        self.date_created = CONFIG.today
        if self.drop_thresh is not None and price <= self.drop_thresh:
            # clear unless it has been adjusted up
            self.drop_thresh = None
        self.hit = ""
        self.alert_msg = ""

        return Transaction(
            action="buy",
            id_=self.id_,
            name=self.name,
            price_each=price,
            qty=buy_qty,
            vendor=self.buy_from,
            holdings=total_qty,
            notes=" - ".join(note for note in [self.notes, extra_notes] if note) or None,
        )

    def sell(self, sell_qty: int, price: float | None = None, extra_notes: str = "") -> Transaction:
        """Update price and ownership fields to reflect a selling event."""
        if not price:
            price = self.buy_price
        remaining_qty = int(self.qty - sell_qty)
        self.qty = remaining_qty

        if self.rise_thresh is not None and (remaining_qty <= 0 or price >= self.rise_thresh):
            self.rise_thresh = None
        self.hit = ""
        self.alert_msg = ""

        return Transaction(
            action="sell",
            id_=self.id_,
            name=self.name,
            price_each=price,
            net_each=(price - self.purchase_price),
            qty=sell_qty,
            vendor=self.sell_to,
            holdings=remaining_qty,
            notes=" - ".join(note for note in [self.notes, extra_notes] if note) or None,
        )

    def to_dict(self) -> CardRow:
        """Return a dict containing all xlsx col names mapped to this entry's attributes."""
        data = {}

        # Card + Entry Metadata
        data["id"]              = self.id_
        data["name"]            = self.name
        data["set"]             = self.set_code
        data["notes"]           = self.notes
        data["tag"]             = self.tag
        data["foil"]            = self.foil
        data["date_created"]    = self.date_created
        data["qty"]             = self.qty

        # Prices
        data["purchase_price"]  = self.purchase_price
        data["sell_price"]      = self.sell_price
        data["buy_price"]       = self.buy_price
        data["buy_from"]        = self.buy_from
        data["sell_to"]         = self.sell_to

        # Alerts
        data["hit"]             = self.hit
        data["drop"]            = self.drop_thresh
        data["rise"]            = self.rise_thresh

        # Add price trends: min, variance, etc.
        data.update(self.price_data)

        return {
            attr_name: ("" if value is None else value)
            for attr_name, value in data.items()
        }

    def _set_best_prices(self, ext_prices: ExtPriceMap | None = None) -> None:
        """Select the best buy/sell vendors based on external price data.

        Initialize state for: self.buy_from, self.sell_to, self.buy_price, self.sell_price.
        """
        id_str = str(self.id_)
        if ext_prices and any(id_str in vendor_prices for vendor_prices in ext_prices.values()):
             # sell_price is what the vendor sells the card for. This is the vendor you should buy_from
            sell_to = max(
                ext_prices,
                key=lambda vendor: ext_prices[vendor].get(id_str, {}).get("buy", INVALID_EXT_BUY_PRICE),
            )

            buy_from = min(
                ext_prices,
                key=lambda vendor: ext_prices[vendor].get(id_str, {}).get("sell", INVALID_EXT_SELL_PRICE),
            )

            buy_price = ext_prices[sell_to][id_str]["buy"]
            if buy_price <= INVALID_EXT_BUY_PRICE:
                sell_to = DEFAULT_VENDOR

            sell_price = ext_prices[buy_from][id_str]["sell"]
            if sell_price >= INVALID_EXT_SELL_PRICE:
                sell_price = self.price_hist[-1]
                buy_from = DEFAULT_VENDOR

        else:
            buy_from = DEFAULT_VENDOR
            sell_to = DEFAULT_VENDOR
            buy_price = 0
            sell_price = self.price_hist[-1]

        self.buy_from = buy_from
        self.sell_to = sell_to
        self.buy_price = buy_price
        self.sell_price = sell_price

    def _set_alerts(self, vals: CardRow) -> None:
        """Evaluate if any alerts are triggered.

        Initialize state for: self.drop_thresh, self.rise_thresh, self.hit, self.alert_msg.
        """
        # Alerts
        self.drop_thresh = vals.get("drop")
        self.rise_thresh = vals.get("rise")
        self.drop_thresh = float(self.drop_thresh) if self.drop_thresh is not None else None
        self.rise_thresh = float(self.rise_thresh) if self.rise_thresh is not None else None
        # Activate a suggestion if triggers are provided for it
        if self.drop_thresh or self.rise_thresh:
            self.active = True
            self.qty = max(0, self.qty)

        if (self.drop_thresh or self.rise_thresh) and ((self.buy_price - self.sell_price) >= CONFIG.min_arb_thresh):
            self.hit = "arb"
            net = self.buy_price - self.sell_price
            self.alert_msg = f"{self.name:<26} | Sell: {self.sell_price} - Buy: {self.buy_price} - Net: {net:.3f}"
        elif self.drop_thresh and self.sell_price <= self.drop_thresh:
            self.hit = "dropped"
            self.alert_msg = f"{self.name:<26} | Target: {self.drop_thresh} - Actual: {self.sell_price}"
        elif self.rise_thresh and self.buy_price >= self.rise_thresh:
            self.hit = "rose"
            self.alert_msg = f"{self.name:<26} | Target: {self.rise_thresh} - Actual: {self.buy_price}"
        else:
            self.hit = ""
            self.alert_msg = ""

    def _set_price_data(self, purchase_price: float | None ) -> None:
        """Calculate price trend statistics.

        Stats include price minimum maximum, change, volatility, etc. at various intervals.
        Initialize state for: self.purchase_price, self.stats.
        """
        stats = {}

        if purchase_price is None or purchase_price <= 0:
            self.purchase_price = self.sell_price
        else:
            self.purchase_price = float(purchase_price)

        # The first few weeks are likely to have anomalous prices. Skip them if specified in the config
        skipped_days = CONFIG.skip_first_n_days
        prices = self.price_hist[-self.age + skipped_days:]

        # Historical prices are used for long-term trends instead of external prices that may be more accurate today
        today_price = prices[-1]

        stats["today_price"]    = today_price
        # stats["purchase_price"] = self.purchase_price
        stats["min"]            = min(prices)
        stats["max"]            = max(prices)

        prices_6_mo             = prices[-min(180, len(prices)):]
        stats["min_6_mo"]       = min(prices_6_mo)
        stats["max_6_mo"]       = max(prices_6_mo)

        # Note: Goatbot actually displays 8d / 1m+1d deltas on their graphs
        if self.age > 7 + skipped_days:
            stats["7d_delta"] = today_price - prices[-7]
            stats["7d_volatility"] = self._calculate_volatility(prices[-7:])
        else:
            stats["7d_delta"] = "N/A"
            stats["7d_volatility"] = "N/A"

        if self.age > 30 + skipped_days:
            stats["30d_delta"] = today_price - prices[-30]
            stats["30d_volatility"] = self._calculate_volatility(prices[-30:])
        else:
            stats["30d_delta"] = "N/A"
            stats["30d_volatility"] = "N/A"

        stats["delta"] = today_price - self.purchase_price
        stats["delta_pct"] = round((stats["delta"] / self.purchase_price), 4) # Excel will multiply by 100%

        self.price_data = stats

    def _calculate_volatility(self, prices: list[float]) -> float:
        """Calculate price volatility using standard deviation of daily returns on the given price list subset."""
        returns = [
            (prices[i] - prices[i - 1]) / prices[i - 1]
            for i in range(1, len(prices))
            if prices[i - 1] != 0
        ]
        return stdev(returns) if len(returns) > 1 else 0.0

def get_suggested_cards(
    conn: sqlite3.Connection,
    opts: dict[str, Any],
    suggested_last_time: list[int],
) -> list[CardRow]:
    """Apply filters from user defined opts and return all matching cards."""
    suggested_rows = []

    defs = build_suggestions_query(
        conn,
        skip_names=opts["skip_names"],
        modern_only=opts["modern_only"],
        allow_standard=opts["allow_standard"],
        allow_foils=False,
    )

    ids = [d.id_ for d in defs]
    prices_map = util.get_price_lists(conn, ids)

    for def_entry in defs:
        if (id_ := def_entry.id_) not in prices_map:
            continue

        name = def_entry.name
        foil = def_entry.foil

        entry = WatchlistEntry(
            vals={
                "id": id_,
                "qty": -1,
                "tag": "auto",
            },
            card_def=def_entry,
            price_hist=prices_map[id_],
            ext_prices=None,
        )

        # Range allowed for the current price. Given with min/max in either order as '[price1];[price2]'
        if (price_cmd := opts.get("price")):
            prices = [float(price.strip()) for price in price_cmd.split(";")]
            if not (min(prices) <= entry.sell_price and max(prices) >= entry.sell_price):
                continue

        if (age := opts.get("age")) and entry.age < age:
            continue

        if (lowest_max := opts.get("lowest_max")) and entry.price_data["max_6_mo"] < lowest_max:
            # highest price of last  6 months is below target (i.e. card hasn't been valuable enough to care about)
            continue

        if (dist := opts.get("dist_from_min")) and (entry.sell_price / entry.price_data["min"]) > dist:
            # current price is too far from lowest historical price
            # TODO: allow user to specify time range to check in
            continue

        if (targets := opts.get("price_change")) and not satisfies_price_change_targets(entry, targets):
            # price has not changed by the specified relative (pct)/absolute amount
            continue

        # Skip if this is not the cheapest version of the card
        cheapest_id = util.get_cheapest_variant(conn, name, foil)
        if id_ != cheapest_id:
            continue

        suggested_rows.append(entry.to_dict())

    print(f"Found {len(suggested_rows)} cards to suggest")

    return sorted(suggested_rows, key=lambda x: (x["id"] in suggested_last_time, x["name"])) # New on top

def build_suggestions_query(
    conn: sqlite3.Connection,
    skip_names: set[str] | None = None,
    modern_only: bool = True,
    allow_standard: bool = False,
    allow_foils: bool = False,
) -> list[CardDef]:
    """Query the DB for all cards matching the criteria.

    Return a card definition for each id that fulfills all requirements.
    This serves as the first layer of filtering for suggested cards.
    """
    query_params = []
    conditions = []
    skip_names = skip_names or set()
    skip_keywords = {"booster"}

    need_legalities = modern_only or not allow_standard

    if not allow_foils:
        conditions.append("foil = FALSE")

    # Skip exact card names
    conditions.append(f"LOWER(name) NOT IN ({', '.join('?' for _ in skip_names)})")
    query_params.extend(name.lower() for name in skip_names)

    # Skip keyword within names
    conditions.extend("LOWER(name) NOT LIKE ?" for _ in skip_keywords)
    query_params.extend(f"%{keyword.lower()}%" for keyword in skip_keywords)


    if need_legalities:
        try:
            set_legalities = json.loads(CONFIG.paths.set_legalities.read_bytes())
        except (FileNotFoundError, json.JSONDecodeError) as e:
            print(f"Warning: failed to update set legalities: {e}")
            set_legalities = {}

        if modern_only:
            modern_sets = set_legalities.get("modern", [])
            if not modern_sets:
                print("Warning: failed to update modern legal sets")

            conditions.append(f"LOWER(set_) IN ({', '.join('?' for _ in modern_sets)})")
            query_params.extend(modern_sets)

        blacklisted_sets: list[str] = CONFIG.lists.suggestion_set_blacklist

        if not allow_standard:
            standard_sets = set_legalities.get("standard", [])
            if not standard_sets:
                print("Warning: failed to update standard legal sets")
            blacklisted_sets.extend(standard_sets)

        conditions.append(
            f"LOWER(set_) NOT IN ({', '.join('?' for _ in blacklisted_sets)})",
        )
        query_params.extend(blacklisted_sets)

    query = "SELECT id, name, set_, foil, ver FROM defs"
    if conditions:
        query += " WHERE " + "\nAND ".join(conditions)

    rows = conn.execute(query, query_params).fetchall()
    return [
        CardDef(
            id_=row[0],
            name=row[1],
            set_code=row[2],
            foil=row[3],
            ver=row[4] if row[4] is not None else "",
        )
        for row in rows
    ]

def satisfies_price_change_targets(entry: WatchlistEntry, targets: str) -> bool:
    """Return True if the card meets all price change targets.

    Targets are supplied by the user as suggestion options like:
        '[days ago 1];[target change since then 1]<%> | [days ago 2];[target change since then 2]<%> | ...'
    """
    for raw_cmd in (cmd.strip() for cmd in targets.split("|")):
        pct = "%" in raw_cmd
        cmd = raw_cmd.rstrip("%").strip()

        days_ago_str, target_str = cmd.split(";")
        days_ago = int(days_ago_str) + 1
        target = float(target_str)
        if pct:
            target /= 100

        if days_ago > entry.age:
            continue

        old = entry.price_hist[-days_ago]
        new = entry.price_hist[-1]
        value = (new - old) / old if pct else new - old

        if target > 0:
            if value < target:
                return False
        elif value > target:
            return False

    return True

def parse_cards_from_sheet(conn: sqlite3.Connection, ws: xw.Sheet, file: XlsxFile) -> list[CardRow]:
    """Parse card data from each row of the sheet and return a CardRow for each.

    Try to determine ids for cards that don't have one listed.
    """
    last_row = ws.range((ws.cells.last_cell.row, file.cols.index("id"))).end("up").row
    rows = ws.range(
        (file.first_data_row, 1),
        (last_row, len(file.cols)),
    ).value

    parsed_vals: list[CardRow] = []

    for row_vals in rows:
        vals = {file.cols[col]: row_vals[col] for col in range(len(file.cols))}
        if not vals["name"] and not vals["id"]:
            # Not enough info to identify card
            continue

        # Resolve the id
        if vals.get("id"):
            vals["id"] = int(vals["id"])
        else:
            vals["id"] = util.name_to_id(
                conn,
                vals.get("name", ""),
                vals.get("set", ""),
                vals.get("foil") or False,
            )

        parsed_vals.append(vals)
    return parsed_vals

def create_entries_from_rows(
    conn: sqlite3.Connection,
    card_rows: list[CardRow],
    ext_prices: ExtPriceMap | None,
) -> tuple[list[WatchlistEntry], list[CardRow]]:
    """Gather price lists and card definitions for all CardRows and instantiates each as a WatchlistEntry.

    Return a list of all entry objects and a list of all CardRows that could not be converted to WatchlistEntries
    """
    entries: list[WatchlistEntry] = []
    failed_rows: list[CardRow] = []

    ids: list[int] = []
    for row in card_rows:
        if isinstance(id_ := row.get("id"), int):
            ids.append(id_)

    prices_map: dict[int, list[float]] = util.get_price_lists(conn, ids)
    defs_map: dict[int, CardDef] = {}

    if ids:
        for attr in conn.execute(
            f"""
            SELECT id, name, set_, foil, ver
            FROM defs
            WHERE id IN ({', '.join('?' for _ in ids)})
            """,
            ids,
        ).fetchall():
            defs_map[attr[0]] = CardDef(
                id_=attr[0],
                name=attr[1],
                set_code=attr[2],
                foil=bool(attr[3]),
                ver=attr[4] if attr[4] is not None else "",
            )

    for row in card_rows:
        try:
            id_ = row["id"]
            entry = WatchlistEntry(
                vals=row,
                card_def=defs_map[id_],
                price_hist=prices_map[id_],
                ext_prices=ext_prices,
            )
        except Exception as e:
            row["notes"] = (row.get("notes") or "") + f" | failed to process - {e}"
            failed_rows.append(row)
            continue

        entries.append(entry)

    return entries, failed_rows

def process_entries(
    logger: Logger,
    wishlist_handler: WishlistHandler,
    card_entries: list[WatchlistEntry],
    file: XlsxFile,
) -> tuple[list[CardRow], list[int], dict[str, list[str]]]:
    """Execute orders for each card.

    Return each card as a row, the ids of previously suggestions (for row sorting purposes), and all alerts
    """
    output_rows: list[CardRow] = []
    suggested_last_time: list[int] = []
    alerts = defaultdict(list)
    for entry in card_entries:
        process_order(entry, file, logger, wishlist_handler)
        if entry.active:
            output_rows.append(entry.to_dict())
            if entry.hit:
                alerts[entry.hit].append(entry.alert_msg)
        elif entry.qty == -1:
            suggested_last_time.append(entry.id_)

    return output_rows, suggested_last_time, alerts

def parse_suggestion_opts_from_sheet(ws: xw.Sheet) -> dict[str, Any]:
    """Initialize the suggestion options field if needed.

    Return the selected suggestion options from the sheet.
    """
    keys = CONFIG.lists.suggestion_option_keys
    num_keys = len(keys)
    opts = dict.fromkeys(keys, None)
    label_cell = ws.cells(file.suggestions_top_row - 1, file.suggestions_left_col)
    if label_cell.value != "Suggestion Args":
        # Create the suggestions box
        label_cell.value = "Suggestion Args"
        for i in range(num_keys):
            ws.cells(file.suggestions_top_row + i, file.suggestions_left_col).value = keys[i]
            ws.cells(file.suggestions_top_row + i, file.suggestions_left_col + 1).value = ""

    # Read args
    opt_range = ws.range(
        (file.suggestions_top_row , file.suggestions_left_col),
        (file.suggestions_top_row + num_keys - 1, file.suggestions_left_col + 1),
    ).value

    for name, val in opt_range:
        if val:
            opts[name] = val

    return opts

def parse_blacklist_from_sheet(file: XlsxFile, ws: xw.Sheet) -> list[str]:
    """Initialize the blacklist input field if needed.

    Return the list of cards typed by the user in the blacklist entry cell
    """
    label_cell = ws.cells(file.blacklist_row, file.blacklist_col - 1)
    blacklist_cell = ws.cells(file.blacklist_row, file.blacklist_col)
    if not label_cell.value:
        label_cell.value = "Blacklist (; sep)"

    blacklist = blacklist_cell.value # User-typed cards
    blacklist_cell.value = ""

    if blacklist is None:
        return []

    return [card.strip() for card in blacklist.split(";")]

def update_blacklist(ws: xw.Sheet, card_rows: list[CardRow], file: XlsxFile) -> tuple[list[str], list[CardRow]]:
    """Update the saved blacklist with any new cards from the sheet.

    Return all blacklisted card names and all card rows that are unaffected by the blacklist.
    """
    path = CONFIG.paths.suggestion_card_blacklist
    remaining_rows = []
    blacklist = parse_blacklist_from_sheet(file, ws)
    for row in card_rows:
        if row["add_to_blacklist"]:
            blacklist.append(row["name"])
        else:
            remaining_rows.append(row)

    blacklist = [new_name.strip() for new_name in blacklist]
    if path.exists():
        existing_blacklist = json.loads(path.read_bytes())
        blacklist.extend(existing_blacklist)
    path.write_bytes(json.dumps(blacklist))
    return blacklist, remaining_rows

def write_back_card_rows(ws: xw.Sheet, wb: xw.Book, output_rows: list[CardRow]) -> None:
    """Update spreadsheet with the new card rows."""
    last_row = ws.range((ws.cells.last_cell.row, file.cols.index("id"))).end("up").row
    ws.range(
        (file.first_data_row, 1),
        (last_row, len(file.cols)),
    ).clear_contents()

    output = [[row.get(col, "") for col in file.cols] for row in output_rows ]
    ws.range((file.first_data_row, 1)).value = output
    ws.cells(1, 2).value = f"Last updated: {CONFIG.today}"

    # Verify that the file actually saved (i.e. no other process locked the spreadsheet)
    old_mod_time = file.path.stat().st_mtime
    wb.save()
    if file.path.stat().st_mtime == old_mod_time:
        raise OSError("Unable to save the Excel file. Make sure that it is not opened in another process")

    print(f"Successfully updated {file.path}\n")

def dump_ids(files: list[XlsxFile]) -> Path:
    """Open all xlsx files and generate a wishlist containing their tracked cards.

    Return a path to the wishlist file.
    """
    excel_app = None
    wishlist_handler = WishlistHandler()

    try:
        excel_app = xw.App(visible=False)
        for file in files:
            wb = excel_app.books.open(file.path)
            ws = wb.sheets[0]
            last_row = ws.range((ws.cells.last_cell.row, file.cols.index("id") + 1)).end("up").row
            rows = ws.range(
                (file.first_data_row, 1),
                (last_row, len(file.cols)),
            ).value

            for row in rows:
                id_ = row[file.cols.index("id")]
                rise = row[file.cols.index("rise")]
                drop = row[file.cols.index("drop")]

                # Skip cards that are only suggestions
                if id_ is None or not (rise or drop):
                    continue

                wishlist_handler.add_to_wishlist(file.wishlist_dir, int(id_), "dump")

            wb.close()
    finally:
        if excel_app is not None:
            excel_app.quit()

    wishlists = wishlist_handler.flush()
    if len(wishlists) == 0:
        raise RuntimeError("Failed to dump ids")

    return wishlists[0]

def update_xlsx(file: XlsxFile) -> None:
    """Update all data in the tracker spreadsheet using today's prices."""
    print(f"=== Updating {file.path} ===")

    successfulUpdate = True
    logger = None
    excel_app = xw.App(visible=False)
    wishlist_handler = WishlistHandler()

    if not file.path.exists():
        temp_book = xw.Book()
        temp_book.save(file.path)
        temp_book.close()

    util.assert_editable(file.path)
    wb = excel_app.books.open(file.path)
    ws = wb.sheets[0]

    if CONFIG.paths.external_prices.exists():
        ext_prices = json.loads(CONFIG.paths.external_prices.read_bytes())
    else:
        ext_prices = None

    try:
        logger = Logger(file.log_path, file.log_cols, excel_app)

        with get_db_connection() as conn:
            card_rows = parse_cards_from_sheet(conn, ws, file)
            blacklist, remaining_rows = update_blacklist(ws, card_rows, file)
            card_entries, failed_rows = create_entries_from_rows(conn, remaining_rows, ext_prices)
            output_rows, suggested_last_time, alerts = process_entries(logger, wishlist_handler, card_entries, file)

            suggestion_opts = parse_suggestion_opts_from_sheet(ws)
            if suggestion_opts.get("enabled"):
                suggestion_opts["skip_names"] = {row["name"] for row in output_rows}.union(blacklist)
                suggested_rows = get_suggested_cards(conn, suggestion_opts, suggested_last_time)
            else:
                suggested_rows = []

        if file.sort_fn is not None:
            output_rows = sorted(output_rows, key=file.sort_fn)
        all_card_rows = output_rows + suggested_rows + failed_rows

        write_back_card_rows(ws, wb, all_card_rows)
        display_alerts(alerts)

    except Exception as e:
        print(f"Error occurred while updating {file.path}: {e}")
        successfulUpdate = False
    finally:
        if logger is not None:
            logger.close()


        if successfulUpdate:
            wishlist_handler.flush()
            ws.cells(1, 2).color = (255, 0, 0) # Red to indicate an update failure
        else:
            ws.cells(1, 2).color = None
        wb.close()

        excel_app.quit()

def process_order(
        entry: WatchlistEntry,
        file: XlsxFile,
        logger: Logger,
        wishlist_handler: WishlistHandler,
    ) -> None:
    """Update the entry, spreadsheet, log, and wishlist based on the result of the order command."""
    if not (order_cmd := entry.pending_order):
        return

    action, qty, prices = parse_order(order_cmd)
    if action == "a": # arbitrage
        logger.add_transaction(process_arb_order(entry, qty, prices))

    elif action == "b": # buy - b[qty];<price>
        buy_price = prices[0] if len(prices) > 0 else None
        transaction = entry.buy(qty, buy_price)
        logger.add_transaction(transaction)
        wishlist_handler.add_to_wishlist(file.wishlist_dir, entry.id_, "buy", entry.buy_from, entry.qty)

    elif action == "s": # sell - s[qty];<price>
        sell_price = prices[0] if len(prices) > 0 else None
        transaction = entry.sell(qty, sell_price)
        logger.add_transaction(transaction)
        wishlist_handler.add_to_wishlist(file.wishlist_dir, entry.id_, "sell", entry.sell_to, qty)

    elif action == "r": # revise
        old = logger.get_last_transaction(entry.id_)
        if old is None:
            raise ValueError(f"Unable to undo a past transaction for {entry.name}")
        logger.add_transaction(process_revise_order(entry, qty, prices, old))
    else:
        raise ValueError(f"Unknown action {action} for {entry.name} with order {order_cmd}")

def parse_order(order_cmd: str) -> tuple[str, int, list[float]]:
    """Parse the components of an order like 'a2;0.5;0.6'."""
    action = order_cmd[0].lower()
    args = [float(arg.strip()) for arg in order_cmd[1:].split(";")]
    qty = int(args[0])
    prices = args[1:]
    return action, qty, prices

def process_arb_order(entry: WatchlistEntry, qty: int, prices: list[float]) -> Transaction:
    """Calculate buy, sell, and profit for an arbitrage order.

    Return the arbitrage transaction.
    Used for orders like 'a[qty];<price1>;<price2>' where 0 or 2 prices are provided.
    """
    buy_price, sell_price = None, None
    if len(prices) == 2:
        buy_price = min(prices)
        sell_price = max(prices)
    elif len(prices) == 0:
        buy_price = entry.sell_price
        sell_price = entry.buy_price
    else:
        raise ValueError(f"Arbitrage order expected 2 prices but received {len(prices)}: {prices}")

    return Transaction(
        action="arb",
        id_=entry.id_,
        name=entry.name,
        price_each=buy_price,
        net_each=(sell_price - buy_price),
        qty=qty,
        vendor=f"{entry.buy_from} -> {entry.sell_to}",
        holdings=entry.qty,
    )

def process_revise_order(entry: WatchlistEntry, qty: int, prices: list[float], old: Transaction) -> Transaction:
    """Update the entry based on the order revision.

    Return the revision transaction.
    Used for orders like 'r[qty];<price1>;<price2>' where 0 to 2 prices are provided.
    """
    price_each = prices[0] if len(prices) > 0 else float(old.price_each)
    net_each = None
    old_cmd = f"{old.action[0]}{int(old.qty)};{old.price_each}"
    qty_diff = old.qty - qty

    if old.action == "sell":
        entry.qty += qty_diff
        holdings = entry.qty

        net_each = float(price_each - entry.purchase_price)

        if qty_diff > 0:
            # Didn't fulfill the entire order - Restore preference to buy more at the previously paid price
            entry.rise_thresh = old.price_each

    elif old.action == "buy":
        if qty_diff != entry.qty:
            entry.purchase_price = round(
                ((entry.purchase_price * entry.qty)
                    - (old.qty * float(old.price_each))
                    + (qty * price_each))
                    / (entry.qty - qty_diff),
                3,
            )
        entry.qty -= qty_diff
        holdings = entry.qty

        if qty_diff <= 0:
            # Successfully fulfilled the entire order - Don't need to buy any more
            entry.drop_thresh = old.price_each

    elif old.action == "arb":
        holdings = entry.qty
        if len(prices) > 1:
            sell_price = max(prices)
            buy_price = min(prices)
        else:
            sell_price = old.net_each + old.price_each # type: ignore reportOptionalOperand
            buy_price = old.price_each
        price_each = buy_price
        net_each = sell_price - buy_price
    else:
        raise ValueError(f"Unknown action {old.action}")

    revised_fields = {
        "action": "revision",
        "old_action": old.action,
        "old_cmd": old_cmd,
        "price_each": price_each,
        "net_each": net_each,
        "qty": qty,
        "holdings": holdings,
    }

    return old.model_copy(update=revised_fields)

def display_alerts(alerts: dict[str, list[str]]) -> None:
    """Print a summary of all alerts triggered while processing the spreadsheet."""
    if alerts:
        print("======== Alerts ========")
        alert_order = {"arb": 0, "rise": 1, "drop": 2}

        for type_, msgs in sorted(alerts.items(), key=lambda kv: alert_order.get(kv[0], float("inf"))):
            print("-" * 10)
            print(capwords(type_))
            for msg in msgs:
                print(f"- {msg}")
        print("=" * 24)

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument(
        "--dump",
        action="store_true",
        help="Dump all ids into a .dek file",
    )

    parser.add_argument(
        "--debug",
        action="store_true",
        help="Only touch the debug spreadsheet and log",
    )

    args = parser.parse_args()

    if args.dump:
        files = [file for file in CONFIG.xlsx.values()  if not file.ignore_for_ext_prices]
        dump_ids(files)

    elif args.debug:
        update_xlsx(CONFIG.xlsx["debug"])

    else:
        # Update all sheets
        for file in CONFIG.xlsx.values():
            if file.auto_update:
                update_xlsx(file)
