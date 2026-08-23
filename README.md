# MTGO Sheets

An automated Magic: The Gathering Online inventory manager and market tracker with an Excel spreadsheet interfeace.

## Features

### Spreadsheet-Driven Tracking

* All day-to-day controls are directly in the sheet
* Set buy/sell thresholds for each card to trigger alerts
* Highly Customizable per-sheet organization options (`config.toml`) allow you to maintain multiple spreadsheets with entirely different goals simultaneously.
* **Examples:**
  * **Full inventory sheet:** Import your entire MTGO collection. Sort cards by when you obtained them, quantity owned, or price. Only add price alerts for cards not currently in decks.
  * **Modern staples sheet:** Completion tracker for the most played Modern cards. Alerts you when cards drop so you can fill out your collection. Tag cycles of cards to be sorted together.
  * **Speculation sheet:** Use all the columns: current price, min/max prices during time intervals, price volatility, etc. Run new queries for card suggestions multiple times a day, and pick which cards to track by simply adding a "target" price.

### Automated order management

* Save full transaction history in a customizable Excel log file.
* Create orders via simple in-spreadhseet commands in the "Order" column. Each command automatically updates this sheet, the transaction log sheet, and the active order file (usable as a MTGO wishlist). Price alert thresholds are updated if this order fulfilled them. Prices are always optional and can be inferred from the sheet.
*
  **Buy:** b4; 0.25 === "Create a buy order for 4 copies. The purchase price is $0.25
  **Sell:** s2 === "Create a sell order for 2 copies. Use the current price on the sheet
  **Revise:** r3; 0.1 === "Revise this sheet and the log for the most recent order placed for this card ID. Update its quantity to 3 at a price of 0.1" Note: This is useful when a seller does not have enough copies of a card in stock, or if they change their price between trades
  **Arbitrage (external prices required):** a6; 0.15; 0.25 === "Purchase 6 copies at $0.15 and immediately sell them to another vendor for $0.25." Note: This only affects the log. The tracker sheet and order file will remain unchanged.

### Card Discovery & Recommendations

* Includes an advanced card suggestion engine to identify new specs
* The in-sheet configuration supports several options such as:
  **price_change:** 5; 0.01 | 30; -20% === "Price increased at least 1 cent in the last 5 days, but is down at least 20% in the last month"
  **dist_from_min:** Same as above, but change relative to minimum price in last N days
  **price:**: 0.5, 2 === "Between $0.5 and $2"
  **modern_only / allow_standard**: 1 / 0 === "Only include cards legal in modern but not legal in standard"
  **age:** 365 === "Only include printings at least a year old"

---

## Getting Started

### Installation

1. Clone the repository.
2. Create and activate a virtual environment: \
   `python -m venv venv` \
   Windows:     `venv\Scripts\activate` \
   macOS/Linux: `source venv/bin/activate` \

3. Install dependencies: \
	`python -m pip install .`

4. Install the package in editable mode: \
   `python -m pip install -e .`

5. Update the price history database \
   `python -m mtgo_sheets.updater`

6. Update the sample spreadsheets \
   `python -m mtgo_sheets.tracker`

### Configuration

MTGO Sheets is driven by a central `config.toml` file. Sheet layout and project structure options are available there on a per-sheet and per-project basis.

* **Sample Files Provided:** Check the `samples/` directory for a fully documented `sample_config.toml` and a matching `sample_inventory.xlsx` to immediately see how the system structures workflows.

### Automation & Daily Updates

* The system is designed to fetch and update prices silently as a daily background task.
* **Windows:** Set Task Scheduler to run `daily_sync.bat`.
* **macOS / Linux:** Set up a cron job to run `daily_sync.sh`.

---

## Acknowledgments

* **Market Data:** MTGO daily and historical price data is provided by [Goatbots](https://www.goatbots.com/).

---

## License

* MIT License
