# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

A personal finance tracker built on **Google Apps Script (GAS)** with Google Sheets as the database. The entire app is two files:

- [code.js](code.js) — GAS backend (server-side logic, Sheets API)
- [index.html](index.html) — Single-page frontend (vanilla HTML/CSS/JS, served via GAS `HtmlService`)

## Deployment

Pushing to `main` automatically deploys via GitHub Actions ([.github/workflows/deploy.yml](.github/workflows/deploy.yml)) — it installs clasp, authenticates using the `CLASPRC_JSON` GitHub secret, and runs `clasp push --force`.

To deploy manually: `clasp push --force` (requires clasp installed and `~/.clasprc.json` credentials).

Run `setupSpreadsheet()` once in the Apps Script editor to initialize the 8 required sheets (only needed on first setup).

## Architecture

### Backend (`code.js`)

- `doGet()` — entry point, serves the HTML UI via `HtmlService`
- `doPost(e)` — handles CSV/XLS/XLSX file uploads for bulk import
- All other functions are called from the frontend via `google.script.run.<functionName>()`

**Sheets schema** (created by `setupSpreadsheet()`):

| Sheet | Purpose |
|---|---|
| Transactions | Core log — ID, Date, Type, Category, Amount, Account, Description, Notes, Cashback, Cashback Amt, Linked Tx, Month, Created At |
| Credit Card Tracker | Card limits, outstanding, statement/due dates, rewards |
| Monthly Summary | Month-level aggregations |
| Category Summary | Per-category spending |
| Budget Settings | Monthly budget amount & alert threshold % |
| Keyword Mapping | Merchant keyword → Category auto-categorization |
| Categories | Category definitions and income/expense classification |
| Lending Tracker | Peer lending records with repayment status |

**Key constants** (top of `code.js`):
- `BUDGET_MONTHLY` — default ₹70,000 (overridden at runtime by Budget Settings sheet)
- `SHEETS` — sheet name constants
- `ACCOUNTS` — 9 predefined bank/wallet accounts (HDFC, Union, SBI, ICICI, Amazon Pay, UPI Lite, Cash, etc.)
- `CATEGORIES` — 30+ categories with type classification

**Bulk import** (`parseBankStatement()`): Bank-specific parsers for HDFC, SBI, Union Bank, ICICI CC, HDFC CC, Amazon Pay, GPay, and BHIM. Each parser handles different column layouts and date formats.

### Frontend (`index.html`)

- Self-contained — all CSS and JS are inline in the single file
- Communicates with backend exclusively via `google.script.run` async calls
- Tab navigation: Dashboard, Add Transaction, Accounts, Lending Tracker, Settings
- Sub-pages: Transactions list, Credit Cards, Monthly Summary, Category Summary, Budget Settings
- Design system uses CSS custom properties (oklch colors), Plus Jakarta Sans and DM Mono fonts
- `is-desktop` class on `<body>` toggles the responsive layout

### Data Flow

Frontend → `google.script.run.functionName(args)` → GAS function reads/writes Google Sheets → callback updates UI

## Key Behaviors to Preserve

- **Duplicate prevention**: bulk import tracks reference IDs to avoid re-importing the same transaction
- **Auto-categorization**: `getAutoCategory(description)` matches merchant keywords from the Keyword Mapping sheet
- **Date normalization**: `parseDate()` accepts DD/MM/YY, DD/MM/YYYY, DD-MM-YYYY, YYYY-MM-DD
- **XLS/XLSX handling**: files are uploaded, saved to Drive, converted to Google Sheets format, then parsed and deleted
