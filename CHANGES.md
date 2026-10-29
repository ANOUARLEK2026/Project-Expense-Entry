# UI & output polish (on top of 3.0.0)

- **LTR layout kept in every language.** `FORCE_LTR_LAYOUT = True` near the top of `app.py` (set it to `False`
  to bring back the mirrored Arabic layout). Arabic text is still shaped correctly; fields, project names,
  description cells and the status line are left-aligned; the PDF/Excel reports are LTR too.
- **Clipped buttons fixed** ("Add expense", "Edit expense"): fixed button heights and the form panel now sits in a
  scroll area, so nothing is squeezed on shorter windows.
- **Date field** always shows `dd.MM.yyyy` with Latin digits (it showed `2026.10.21` while the table showed `21.10.2026`).
- Table headers are aligned with their columns (ID centre, Description left, Date centre, Amount right).
- After an export the status bar shows the file name with **Open** / **Show in folder** links instead of a long mixed-direction path.
- **PDF report:** title banner with project and period, larger KPI cards, dark table headers, emphasised total rows,
  a *Share %* column in the monthly summary, "Amount (DZD)" header, footer rule.
- **Excel report:** coloured title band, left-aligned descriptions / right-aligned amounts, tuned column widths, frozen header,
  no gridlines, A4 fit-to-width print setup with repeating header row and page numbers, total shown in the title block,
  data labels (k / M) on the monthly chart.
- Chart: dotted guide lines.

# Project Expense Entry Pro 3.0.0

Same app, same workflow, same files. `app.py` was improved in place; nothing was rebuilt.
No new dependencies (still `PySide6` + `pywin32`). Settings keys, the saved password,
the `.xlsm` format, VBA, templates and the `~ExpenseBackups` folder are unchanged.

## What's new
**Look & feel** - consistent dark theme plus a light theme (button in the header, remembered),
clearer tables, button hierarchy (primary / secondary / ghost / danger), better spacing.

**Dashboard** - the four cards (total, today, this month, transactions) plus a 12-month bar chart.
Click a bar to filter the table to that month; click again to clear.

**Daily work**
- Click a column header to sort (ascending, descending, back to default).
- Min/Max amount filters next to the search box; a footer shows the count and total of what is shown.
- Description auto-complete from earlier entries; optional "Keep the date after saving".
- Warning before adding an identical expense (same description, date, amount).
- Ctrl+N new entry, Ctrl+E reports, Ctrl+F search (plus the existing shortcuts).

**Reports** (button "Export", Ctrl+E): choose a period (all / this month / last month / this year /
last year / custom) and an optional description filter, with a live count and total.
- *Excel (.xlsx)*: new file with a title block, a real Excel table (filters + total row),
  a monthly-summary sheet with a chart. Project files are never touched.
- *PDF*: A4 report with totals, average, largest expense, monthly chart and summary, full details,
  repeating table header and page numbers. Arabic reports are laid out right-to-left.
  (The original app had no export, so this is new rather than "improved".)

**Data safety**
- Backups: still automatic before every change, now skipped if nothing changed since the last one,
  verified after copying, and pruned (newest 20 + last of each day for 30 days). Manual backups are never deleted.
- *Backup > Restore from backup...*: pick a backup; the current file is saved first, so a restore can be undone.
- Edit/delete refuse to touch a row that was changed in Excel since it was loaded (previously the
  row *position* was trusted, which could overwrite the wrong expense). The list refreshes instead.
- Delete confirmation shows the expense; closing or switching project with an unsaved entry asks first.
- Dates: if Excel reads a typed date with day/month swapped (regional settings), the correct date is written instead.

**Languages** - Arabic now mirrors the whole layout (right-to-left); numbers, dates and the chart stay
left-to-right. Login screen is translated. French accents fixed.

## Not changed on purpose
Login/password logic, template-based new projects, global search, ID generation by the workbook's VBA.

## Tests
`py -3 -m unittest discover -s tests -v` - checks backups/restore/pruning, the stale-row protection,
filters, sorting, report totals and all three translations. They need no Excel.
