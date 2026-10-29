# Project Expense Entry

**GitHub description:** Windows desktop expense tracker with project-based records, searchable transactions, backups, and PDF/Excel reporting.

Project Expense Entry is designed for recording and reviewing expenses with descriptions, dates, amounts, and project context. It includes reporting and data-safety features for practical day-to-day bookkeeping.

## Features
- Add, edit, search, filter, and sort expense entries.
- Dashboard with totals, daily/monthly indicators, and a 12-month chart.
- Export selected periods to Excel or PDF, including summaries and transaction details.
- Automatic and manual backups, backup verification, and restore workflow.
- Duplicate-entry warnings, description suggestions, and keyboard shortcuts.
- Arabic, English, and French language options; light and dark themes.

## Run on Windows
1. Install Python 3.10+.
2. Install dependencies: `py -m pip install -r requirements.txt`.
3. Launch `Run_Expense_Entry.bat` or run `py app.py`.

`pywin32` is used for Windows integration. Review `CHANGES.md` for version notes. Use a copy of your workbook/data when testing restore or import features.

## Tech stack
Python · PySide6 · pywin32 · PDF/Excel export
