"""Logic tests for Project Expense Entry Pro (no Excel or display needed).

Run from the project folder:  py -3 -m unittest discover -s tests -v
If PySide6 is not installed (e.g. on a CI box) a tiny stub is used, because only
the plain-Python logic is exercised here.
"""
import os
import sys
import tempfile
import types
import unittest
import zipfile
from datetime import date, datetime, timedelta
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

try:  # pragma: no cover
    import PySide6  # noqa: F401
except ImportError:  # install a minimal stand-in so `import app` works
    class _Meta(type):
        def __getattr__(cls, name):  # enum-style access such as QAbstractItemView.SelectionBehavior
            if name.startswith("__"):
                raise AttributeError(name)
            return mock.MagicMock()

    class _Any(metaclass=_Meta):
        def __init__(self, *a, **k):
            pass

        def __getattr__(self, name):
            return mock.MagicMock()

    def _module(name, **attrs):
        mod = types.ModuleType(name)
        mod.__dict__.update(attrs)
        sys.modules[name] = mod
        return mod

    class _Namespace(mock.MagicMock):
        pass

    def _factory(*names):
        return {n: type(n, (_Any,), {}) for n in names}

    _module("PySide6")
    _module("PySide6.QtCore", Qt=mock.MagicMock(), Signal=lambda *a, **k: mock.MagicMock(),
            **_factory("QDate", "QMarginsF", "QObject", "QPointF", "QRectF", "QSettings", "QSizeF",
                       "QStringListModel", "QTimer", "QUrl"))
    _module("PySide6.QtGui", **_factory("QAction", "QColor", "QFont", "QIcon", "QImage", "QKeySequence",
                                         "QPageLayout", "QPageSize", "QPainter", "QPdfWriter", "QPen",
                                         "QShortcut", "QTextDocument", "QTextOption"))
    _module("PySide6.QtWidgets", **_factory(
        "QAbstractItemView", "QApplication", "QCheckBox", "QComboBox", "QCompleter", "QDateEdit", "QDialog",
        "QDialogButtonBox", "QFileDialog", "QFrame", "QGraphicsDropShadowEffect", "QGridLayout", "QHBoxLayout",
        "QHeaderView", "QLabel", "QLineEdit", "QListWidget", "QListWidgetItem", "QMainWindow", "QMenu",
        "QMessageBox", "QProgressBar", "QPushButton", "QSplitter", "QStackedWidget", "QTableWidget",
        "QTableWidgetItem", "QToolTip", "QVBoxLayout", "QWidget"))

import app  # noqa: E402


def make_rows(spec):
    """spec: [(id, description, date|None, amount)] -> prepared rows in Excel order."""
    rows = [{"excel_index": i, "id": rid, "description": desc, "date": d, "amount": amt}
            for i, (rid, desc, d, amt) in enumerate(spec, 1)]
    return app.prepare_rows(rows)


def write_xlsm(path, payload=b"x"):
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("data.txt", payload)


class BackupTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self.tmp.name)
        self.project = self.dir / "PRG-001.xlsm"
        write_xlsm(self.project, b"v1")

    def tearDown(self):
        self.tmp.cleanup()

    def test_auto_backup_created_and_reused_when_identical(self):
        first = app.make_backup(self.project)
        self.assertTrue(Path(first).exists())
        self.assertEqual(app.make_backup(self.project), first)  # nothing changed -> no new copy
        write_xlsm(self.project, b"v2-longer")
        second = app.make_backup(self.project)
        self.assertNotEqual(first, second)
        self.assertEqual(len(app.list_backups(self.project)), 2)

    def test_manual_backup_always_new_and_tagged(self):
        auto = app.make_backup(self.project)
        manual = app.make_backup(self.project, "manual")
        self.assertNotEqual(auto, manual)
        kinds = {b["kind"] for b in app.list_backups(self.project)}
        self.assertEqual(kinds, {"auto", "manual"})

    def test_legacy_backup_names_are_listed(self):
        folder = app.backup_dir_for(self.project)
        folder.mkdir()
        write_xlsm(folder / "PRG-001_20260105_101500_123456.xlsm")
        write_xlsm(folder / "OTHER_20260105_101500_123456.xlsm")      # other project: ignored
        write_xlsm(folder / "PRG-001_notes.xlsm")                      # not a backup name: ignored
        found = app.list_backups(self.project)
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0]["kind"], "auto")
        self.assertEqual(found[0]["when"], datetime(2026, 1, 5, 10, 15, 0, 123456))

    def _fake_backups(self, stamps, kind=""):
        folder = app.backup_dir_for(self.project)
        folder.mkdir(exist_ok=True)
        for s in stamps:
            write_xlsm(folder / f"PRG-001_{s}{kind}.xlsm", s.encode())

    def test_prune_keeps_recent_and_daily_and_never_manual(self):
        now = datetime(2026, 10, 21, 12, 0, 0)
        # 30 backups today, one per day for 40 older days, plus an old manual one
        today = [f"20261021_{h:02d}{m:02d}00_000000" for h in range(8, 12) for m in range(0, 60, 8)]
        older = [(now - timedelta(days=d)).strftime("%Y%m%d") + "_090000_000000" for d in range(1, 41)]
        self._fake_backups(today + older)
        self._fake_backups(["20250101_090000_000000"], "_manual")
        removed = app.prune_backups(self.project, now=now)
        left = app.list_backups(self.project)
        auto = [b for b in left if b["kind"] == "auto"]
        self.assertEqual(sum(1 for b in auto if b["when"].date() == now.date()), 20 + 0 if len(today) >= 20 else len(today))
        # one per day is kept for the last 30 days
        days = {b["when"].date() for b in auto}
        for d in range(1, 31):
            self.assertIn((now - timedelta(days=d)).date(), days)
        self.assertNotIn((now - timedelta(days=35)).date(), days)
        self.assertTrue(any(b["kind"] == "manual" for b in left))
        self.assertTrue(removed)

    def test_restore_makes_safety_copy_and_replaces_file(self):
        backup = app.make_backup(self.project, "manual")
        write_xlsm(self.project, b"v2-broken-by-mistake")
        safety = app.restore_backup(self.project, backup)
        self.assertIsNotNone(safety)
        with zipfile.ZipFile(self.project) as z:
            self.assertEqual(z.read("data.txt"), b"v1")
        with zipfile.ZipFile(safety) as z:                       # the file before the restore is kept
            self.assertEqual(z.read("data.txt"), b"v2-broken-by-mistake")
        self.assertIn("prerestore", {b["kind"] for b in app.list_backups(self.project)})
        self.assertFalse(list(self.dir.glob("*.restore.tmp")))

    def test_restore_rejects_non_excel_file(self):
        bad = app.backup_dir_for(self.project)
        bad.mkdir()
        fake = bad / "PRG-001_20260101_000000_000000.xlsm"
        fake.write_text("not a zip")
        with self.assertRaises(ValueError):
            app.restore_backup(self.project, fake)
        with zipfile.ZipFile(self.project) as z:                 # project untouched
            self.assertEqual(z.read("data.txt"), b"v1")


class RowLogicTests(unittest.TestCase):
    def setUp(self):
        self.rows = make_rows([
            (1, "Cement bags", datetime(2026, 9, 3), 12000),
            (2, "Diesel", datetime(2026, 10, 1), 4500.5),
            (3, "cement delivery", datetime(2026, 10, 15), 800),
            (4, "Broken date", "not a date", 100),
            (5, "Bad amount", datetime(2026, 10, 20), "n/a"),
        ])

    def test_filters(self):
        self.assertEqual([r["id"] for r in app.filter_rows(self.rows, terms=app.search_terms("cement"))], [1, 3])
        self.assertEqual([r["id"] for r in app.filter_rows(self.rows, month="2026-10")], [2, 3, 5])
        self.assertEqual([r["id"] for r in app.filter_rows(self.rows, min_amount=1000)], [1, 2])
        self.assertEqual([r["id"] for r in app.filter_rows(self.rows, max_amount=1000)], [3, 4])
        self.assertEqual([r["id"] for r in app.filter_rows(self.rows, min_amount=500, max_amount=5000)], [2, 3])

    def test_filter_amount_parsing(self):
        self.assertIsNone(app.parse_filter_amount(""))
        self.assertIsNone(app.parse_filter_amount("abc"))
        self.assertEqual(app.parse_filter_amount("0"), 0.0)
        self.assertEqual(app.parse_filter_amount("1 500,50"), 1500.5)

    def test_sorting(self):
        by_amount = [r["id"] for r in app.sort_rows(self.rows, "amount", True)]
        self.assertEqual(by_amount, [5, 4, 3, 2, 1])            # non-numeric counts as 0
        self.assertEqual([r["id"] for r in app.sort_rows(self.rows, "amount", False)], [1, 2, 3, 4, 5])
        by_date = [r["id"] for r in app.sort_rows(self.rows, "date", True)]
        self.assertEqual(by_date[:3], [1, 2, 3])
        self.assertEqual(by_date[-1], 4)                         # undated rows go last
        self.assertEqual([r["id"] for r in app.sort_rows(self.rows, "description", True)][0], 5)

    def test_rows_match_detects_changed_row(self):
        expected = {"description": " Diesel ", "date": datetime(2026, 10, 1), "amount": 4500.5}
        same = {"description": "Diesel", "date": datetime(2026, 10, 1, 0, 0), "amount": 4500.50}
        self.assertTrue(app.rows_match(expected, same))
        self.assertFalse(app.rows_match(expected, {**same, "description": "Gas"}))
        self.assertFalse(app.rows_match(expected, {**same, "amount": 4500}))
        self.assertFalse(app.rows_match(expected, {**same, "date": datetime(2026, 10, 2)}))
        self.assertTrue(app.rows_match({"description": "x", "date": "01.10.2026", "amount": 5},
                                       {"description": "x", "date": datetime(2026, 10, 1), "amount": 5.0}))

    def test_duplicates_and_descriptions(self):
        self.assertIsNotNone(app.find_duplicate(self.rows, "DIESEL", date(2026, 10, 1), 4500.5))
        self.assertIsNone(app.find_duplicate(self.rows, "Diesel", date(2026, 10, 2), 4500.5))
        self.assertEqual(app.unique_descriptions(self.rows)[:2], ["Bad amount", "Broken date"])

    def test_stats_and_chart_months(self):
        s = app.summarize_rows(self.rows)
        self.assertEqual(s["count"], 4)
        self.assertEqual(s["invalid"], 1)
        self.assertAlmostEqual(s["total"], 12000 + 4500.5 + 800 + 100)
        self.assertEqual(s["largest"]["id"], 1)
        data = app.chart_months(self.rows, today=date(2026, 10, 21))
        self.assertEqual(len(data), 12)
        self.assertEqual(data[-1][0], "2026-10")
        self.assertEqual(data[0][0], "2025-11")
        self.assertAlmostEqual(dict((k, t) for k, t, _c in data)["2026-10"], 4500.5 + 800)
        # old data: the window ends at the last month that has data, not at today
        old = make_rows([(1, "x", datetime(2024, 3, 5), 10)])
        self.assertEqual(app.chart_months(old, today=date(2026, 10, 21))[-1][0], "2024-03")
        self.assertEqual(app.chart_months([], today=date(2026, 10, 21)), [])

    def test_amount_parsing_unchanged(self):
        self.assertEqual(app.parse_amount("1.234,50"), 1234.5)
        self.assertEqual(app.parse_amount("1,234.50"), 1234.5)
        with self.assertRaises(ValueError):
            app.parse_amount("0")


class ReportTests(unittest.TestCase):
    def setUp(self):
        self.rows = make_rows([
            (1, "Cement <b>", datetime(2026, 9, 3), 12000),
            (2, "Diesel", datetime(2026, 10, 1), 4500.5),
            (3, "Undated", "", 50),
            (4, "Bad amount", datetime(2026, 10, 20), "n/a"),
        ])

    def test_presets(self):
        today = date(2026, 10, 21)
        self.assertEqual(app.preset_range("this_month", today, None, None), (date(2026, 10, 1), date(2026, 10, 31)))
        self.assertEqual(app.preset_range("last_month", date(2026, 1, 15), None, None), (date(2025, 12, 1), date(2025, 12, 31)))
        self.assertEqual(app.preset_range("this_year", today, None, None), (date(2026, 1, 1), date(2026, 12, 31)))
        self.assertEqual(app.preset_range("last_year", today, None, None), (date(2025, 1, 1), date(2025, 12, 31)))
        self.assertEqual(app.last_day_of_month(date(2024, 2, 10)), date(2024, 2, 29))

    def test_range_excludes_undated_and_is_inclusive(self):
        rows = app.report_rows(self.rows, date(2026, 9, 3), date(2026, 10, 1), "", False)
        self.assertEqual([r["id"] for r in rows], [1, 2])
        all_rows = app.report_rows(self.rows, all_dates=True)
        self.assertEqual([r["id"] for r in all_rows], [1, 2, 4, 3])   # oldest first, undated last

    def test_payload_and_html(self):
        payload = app.build_report_payload("PRG-001", self.rows, date(2026, 1, 1), date(2026, 12, 31), "", False, "en")
        self.assertEqual(payload["summary"]["count"], 2)
        self.assertEqual(payload["undated_excluded"], 1)
        self.assertEqual(payload["period"], "01.01.2026 - 31.12.2026")
        self.assertEqual([m[0] for m in payload["months"]], ["2026-09", "2026-10"])
        html_en = app.build_report_html(payload, "en", (500, 150))
        self.assertIn("16,500.50 DZD", html_en)
        self.assertIn("Cement &lt;b&gt;", html_en)              # user text is escaped
        self.assertIn('src="chart.png"', html_en)
        self.assertIn("without a valid date", html_en)
        # every language renders, Arabic is right-to-left with the first column on the right
        for lang in ("fr", "ar"):
            out = app.build_report_html(payload, lang)
            self.assertIn("16,500.50", out)
        html_ar = app.build_report_html(payload, "ar")
        self.assertIn("تقرير المصاريف", html_ar)
        if app.is_rtl("ar"):
            self.assertLess(html_ar.index(">المبلغ (DZD)<"), html_ar.index(">الرقم<"))   # Amount before ID in source = ID ends on the right
        else:
            self.assertLess(html_ar.index(">الرقم<"), html_ar.index(">المبلغ (DZD)<"))   # LTR layout: ID first
        self.assertIn("100%", html_en)                                   # monthly share column
        self.assertIn(">Share<", html_en)

    def test_html_is_well_formed(self):
        from html.parser import HTMLParser

        class Checker(HTMLParser):
            def __init__(self):
                super().__init__()
                self.stack = []

            def handle_starttag(self, tag, attrs):
                if tag not in ("br", "img"):
                    self.stack.append(tag)

            def handle_startendtag(self, tag, attrs):
                pass  # <br/> and <img/> need no closing tag

            def handle_endtag(self, tag):
                assert self.stack and self.stack[-1] == tag, f"mismatched </{tag}> (open: {self.stack[-3:]})"
                self.stack.pop()

        payload = app.build_report_payload("P", self.rows, None, None, "", True, "fr")
        checker = Checker()
        checker.feed(app.build_report_html(payload, "fr", (400, 100)))
        self.assertEqual(checker.stack, [])

    def test_all_translations_have_the_same_keys(self):
        en = set(app.TEXT["en"])
        for lang in ("fr", "ar"):
            self.assertFalse(en - set(app.TEXT[lang]), f"{lang} is missing: {sorted(en - set(app.TEXT[lang]))}")
        import string
        for key in en:
            fields = lambda s: {f[1] for f in string.Formatter().parse(s) if f[1]}
            for lang in ("fr", "ar"):
                self.assertEqual(fields(app.TEXT["en"][key]), fields(app.TEXT[lang][key]), f"{lang}.{key}")


# ---------------------------------------------------------------- fake Excel COM
class FakeCell:
    def __init__(self, row, col, us_dates):
        self.row, self.col, self._us = row, col, us_dates
        self.NumberFormat = "General"

    @property
    def Value(self):
        return self.row.values[self.col - 1]

    @Value.setter
    def Value(self, v):
        if isinstance(v, str) and self.col == self.row.date_col and self.row.parse_dates:
            d, m, y = (int(x) for x in v.split("."))
            v = datetime(y, d, m) if self._us else datetime(y, m, d)      # US locale swaps day/month
        self.row.values[self.col - 1] = v


class FakeRow:
    def __init__(self, values, date_col, parse_dates, us_dates):
        self.values, self.date_col, self.parse_dates, self.us = list(values), date_col, parse_dates, us_dates

    @property
    def Range(self):
        return self

    def Cells(self, _r, col):
        return FakeCell(self, col, self.us)


class FakeListRows:
    def __init__(self, table):
        self.t = table

    @property
    def Count(self):
        return len(self.t.rows)

    def Add(self):
        self.t.next_id += 1
        self.t.rows.append(FakeRow([self.t.next_id, None, None, None], 3, self.t.parse_dates, self.t.us))

    def __call__(self, i):
        row = self.t.rows[i - 1]
        row.Delete = lambda: self.t.rows.pop(i - 1)
        return row


class FakeTable:
    Name = "Expenses"

    def __init__(self, rows, parse_dates=True, us=False):
        self.parse_dates, self.us = parse_dates, us
        self.rows = [FakeRow(r, 3, parse_dates, us) for r in rows]
        self.next_id = max((r[0] for r in rows), default=0)
        self.ListRows = FakeListRows(self)
        self.Parent = types.SimpleNamespace(Name="Sheet1")
        self.names = ["ID", "Description", "Date", "Amount"]
        self.ListColumns = types.SimpleNamespace(Count=4)
        self.ListColumns.__call__ = None

    def column(self, i):
        return types.SimpleNamespace(Name=self.names[i - 1])

    @property
    def DataBodyRange(self):
        return types.SimpleNamespace(Value=tuple(tuple(r.values) for r in self.rows))


class FakeColumns:
    def __init__(self, table):
        self.t, self.Count = table, 4

    def __call__(self, i):
        return self.t.column(i)


class FakeWorkbook:
    def __init__(self, store, path, read_only):
        self.store, self.path, self.ReadOnly = store, path, read_only
        self.working = FakeTable([list(r.values) for r in store.table.rows], store.table.parse_dates, store.table.us)
        self.working.ListColumns = FakeColumns(self.working)
        self.Worksheets = [types.SimpleNamespace(ListObjects=types.SimpleNamespace(
            Count=1, __call__=None))]
        ws = self.Worksheets[0]
        table = self.working

        class LO:
            Count = 1

            def __call__(self, _i):
                return table
        ws.ListObjects = LO()

    def Save(self):
        self.store.saves += 1
        self.store.table.rows = [FakeRow(r.values, 3, r.parse_dates, r.us) for r in self.working.rows]

    def Close(self, SaveChanges=False):
        pass


class FakeExcel:
    def __init__(self, table):
        self.table = table
        self.saves = 0
        self.opened = []
        self.Workbooks = self
        self.Visible = self.DisplayAlerts = self.EnableEvents = self.ScreenUpdating = False
        self.AutomationSecurity = 2

    def Open(self, path, **kw):
        self.opened.append(path)
        return FakeWorkbook(self, path, kw.get("ReadOnly", False))

    def Calculate(self):
        pass

    def Quit(self):
        pass


class ManagerTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.path = Path(self.tmp.name) / "PRG-001.xlsm"
        write_xlsm(self.path, b"orig")
        self.excel = FakeExcel(FakeTable([
            [1, "Cement", datetime(2026, 10, 1), 100.0],
            [2, "Diesel", datetime(2026, 10, 2), 50.0],
        ]))
        fake_win32 = types.SimpleNamespace(DispatchEx=lambda _n: self.excel)
        self.patch = mock.patch.object(app, "win32", fake_win32)
        self.patch.start()
        self.mgr = app.ExpenseManager()

    def tearDown(self):
        self.patch.stop()
        self.tmp.cleanup()

    def rows(self):
        return self.excel.table.rows

    def test_load_reads_rows(self):
        meta, rows = self.mgr.load_project(str(self.path))
        self.assertEqual(meta["rows"], 2)
        self.assertEqual(rows[1]["description"], "Diesel")

    def test_add_backs_up_first_and_saves(self):
        res = self.mgr.add_expense(str(self.path), "Gravel", "05.10.2026", 75.0)
        self.assertEqual(len(self.rows()), 3)
        self.assertEqual(res["id"], 3)
        self.assertTrue(Path(res["backup"]).exists())
        self.assertEqual(self.rows()[2].values[2], datetime(2026, 10, 5))

    def test_edit_with_matching_expected_row(self):
        _, loaded = self.mgr.load_project(str(self.path))
        app.prepare_rows(loaded)
        self.mgr.edit_expense(str(self.path), 2, "Diesel 2", "02.10.2026", 60.0, expected=loaded[1])
        self.assertEqual(self.rows()[1].values[1], "Diesel 2")
        self.assertEqual(self.rows()[1].values[3], 60.0)

    def test_edit_and_delete_refuse_a_row_changed_elsewhere(self):
        _, loaded = self.mgr.load_project(str(self.path))
        app.prepare_rows(loaded)
        # someone deletes row 1 in Excel: row index 2 now holds a different expense (the old row 2 moved to 1)
        self.excel.table.rows.pop(0)
        with self.assertRaises(app.StaleRowError):
            self.mgr.edit_expense(str(self.path), 1, "X", "01.10.2026", 1.0, expected=loaded[0])
        with self.assertRaises(app.StaleRowError):
            self.mgr.delete_expense(str(self.path), 1, expected=loaded[0])
        with self.assertRaises(app.StaleRowError):                # index beyond the table
            self.mgr.delete_expense(str(self.path), 2, expected=loaded[1])
        self.assertEqual([r.values[1] for r in self.rows()], ["Diesel"])   # nothing was modified
        self.assertEqual(self.excel.saves, 0)

    def test_delete_with_matching_row(self):
        _, loaded = self.mgr.load_project(str(self.path))
        app.prepare_rows(loaded)
        res = self.mgr.delete_expense(str(self.path), 1, expected=loaded[0])
        self.assertEqual([r.values[1] for r in self.rows()], ["Diesel"])
        self.assertEqual(res["meta"]["rows"], 1)

    def test_swapped_day_and_month_is_corrected(self):
        self.excel.table.us = True                                    # Excel would read 05.10.2026 as May 10
        self.excel.table.parse_dates = True
        self.mgr.add_expense(str(self.path), "Gravel", "05.10.2026", 75.0)
        stored = self.rows()[2].values[2]
        self.assertEqual(app.parse_excel_date(stored), date(2026, 10, 5))

    def test_text_dates_stay_as_before(self):
        self.excel.table.parse_dates = False                           # column kept as text
        self.mgr.add_expense(str(self.path), "Gravel", "05.10.2026", 75.0)
        self.assertEqual(self.rows()[2].values[2], "05.10.2026")

    def test_failed_write_leaves_a_backup_but_no_changes(self):
        with mock.patch.object(app.ExpenseManager, "_wait_for_id", side_effect=RuntimeError("boom")):
            with self.assertRaises(RuntimeError):
                self.mgr.add_expense(str(self.path), "Gravel", "05.10.2026", 75.0)
        self.assertEqual(len(self.rows()), 2)
        self.assertEqual(len(app.list_backups(self.path)), 1)


if __name__ == "__main__":
    unittest.main()
