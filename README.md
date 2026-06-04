# AI Trading Journal

AI-assisted trading journal that extracts trade details from screenshots and exports a professional Excel review workbook.

Recommended repository name: `ai-trading-journal`

## Overview

AI Trading Journal is a local desktop application for traders who want a structured workflow for trade logging, screenshot-assisted data entry, post-trade review, and Excel-based analysis. The app is designed as a portfolio project across IT support, finance technology, AI automation, data analysis, and trading review systems.

The application runs locally with Python and Tkinter. Trade records are stored on the user's machine, and the optional Claude/Anthropic integration is only used when the user chooses to run screenshot recognition.

## Features

- Manual trade entry for symbol, direction, quantity, price, P/L, fees, risk levels, tags, and notes.
- Optional AI screenshot recognition using the Claude/Anthropic Messages API.
- Review workflow for entry reason, exit reason, trade summary, lessons learned, emotion state, and trade grade.
- Searchable local trade history.
- Professional Excel export with:
  - Trade records sheet
  - Statistical analysis sheet
  - Daily P/L sheet
  - Cumulative P/L chart when enough data exists
- Local-first data storage with no default cloud database.
- VS Code launch and task configuration for portfolio-friendly development.

## Tech Stack

- Python 3
- Tkinter desktop GUI
- OpenPyXL for Excel workbook generation
- Pillow for image compatibility
- Pandas dependency included for future data analysis workflows
- Claude/Anthropic API for optional screenshot extraction
- JSON local storage

## Screenshots

Add sanitized screenshots before publishing the repository.

Suggested screenshots:

- Main trading journal dashboard
- Manual trade entry form
- AI screenshot recognition tab
- Exported Excel workbook summary sheet

Screenshot placeholders live in `docs/screenshots/`.

Important: Do not publish screenshots that show real account numbers, order IDs, balances, broker names, positions, API keys, or personally identifiable financial information.

## Project Structure

```text
ai-trading-journal/
|-- trade_journal.py              # Main Tkinter desktop application
|-- requirements.txt              # Python package dependencies
|-- start_mac_linux.sh            # Mac/Linux launcher
|-- start_windows.bat             # Windows launcher
|-- .env.example                  # Example environment variables
|-- .gitignore                    # Local data, secrets, and cache exclusions
|-- .vscode/                      # VS Code run configuration
|-- docs/screenshots/             # Placeholder folder for sanitized screenshots
|-- LICENSE                       # MIT License
`-- README.md                     # Project documentation
```

## Getting Started

### Prerequisites

- Python 3.10 or newer recommended
- Git
- Optional: VS Code with the Python extension
- Optional: Claude/Anthropic API key for screenshot recognition

### Clone

```bash
git clone https://github.com/<your-username>/ai-trading-journal.git
cd ai-trading-journal
```

### Create a Virtual Environment

Mac/Linux:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

Windows PowerShell:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

### Configure Optional AI Recognition

Copy the example environment file:

```bash
cp .env.example .env
```

Then set:

```bash
ANTHROPIC_API_KEY=your_anthropic_api_key_here
```

You can also enter the API key inside the app through the API settings dialog. The app stores that key in a local file in the user's home directory by default.

### Run

```bash
python trade_journal.py
```

Mac/Linux shortcut:

```bash
./start_mac_linux.sh
```

Windows shortcut:

```bat
start_windows.bat
```

### Run from VS Code

1. Open the project folder in VS Code.
2. Select the `.venv` interpreter.
3. Use `Run > Run Without Debugging`.
4. Or use the default build task with `Cmd+Shift+B` on macOS or `Ctrl+Shift+B` on Windows/Linux.

## Data Storage

By default, the app stores local data outside the repository:

- Trade data: `~/.trade_journal_data.json`
- API key file: `~/.trade_journal_api_key.txt`

Optional overrides are available in `.env`:

```bash
TRADE_JOURNAL_DATA_FILE=~/.trade_journal_data.json
TRADE_JOURNAL_API_KEY_FILE=~/.trade_journal_api_key.txt
```

## Security and Privacy

Trading screenshots and journals can contain sensitive financial information. Treat all local data and exported workbooks as private.

Recommended practices:

- Never commit `.env`, API keys, trade screenshots, exported Excel files, or raw broker statements.
- Use sanitized demo screenshots for GitHub and LinkedIn.
- Review screenshots for account numbers, balances, broker identifiers, order IDs, names, emails, and portfolio values before publishing.
- Rotate API keys if they are accidentally exposed.
- Keep `.trade_journal_data.json` and `.trade_journal_api_key.txt` out of the repository.
- Use local-only storage unless you intentionally add encrypted sync or a secure backend.

The `.gitignore` file is configured to exclude local environments, caches, `.env`, generated workbooks, and common sensitive data folders.

## Testing and Validation

Basic local checks:

```bash
python -m py_compile trade_journal.py
mkdir -p exports
python - <<'PY'
from trade_journal import export_to_excel, new_trade
trade = new_trade()
trade.update({
    "symbol": "AAPL",
    "quantity": "10",
    "price": "150",
    "amount": "1500",
    "pnl": "125.50",
})
export_to_excel([trade], "exports/smoke_test.xlsx")
print("Excel export smoke test passed")
PY
```

The app has also been manually verified to launch from VS Code using the project virtual environment.

## Roadmap

- Add automated unit tests for data validation and Excel export.
- Add a sanitized demo dataset for portfolio screenshots.
- Add CSV import for broker exports.
- Add dashboard charts inside the desktop app.
- Add encrypted local storage for API keys and sensitive trade data.
- Add tagging analytics for strategy, setup, emotion, and mistake patterns.
- Add packaged desktop builds for macOS and Windows.

## Portfolio Positioning

This project demonstrates:

- IT support: local setup, dependency management, VS Code workflow, and user-friendly troubleshooting.
- Finance technology: structured trade capture, review workflow, and Excel reporting.
- AI automation: screenshot-based trade detail extraction through an LLM API.
- Data analysis: daily P/L aggregation, summary statistics, and workbook export.
- Security awareness: local-first data handling and explicit sensitive-data guidance.

## License

This project is licensed under the MIT License. See `LICENSE` for details.
