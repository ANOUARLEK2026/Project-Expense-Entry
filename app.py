import filecmp
import hashlib
import hmac
import html
import os
import re
import shutil
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor
from contextlib import suppress
from datetime import date, datetime, timedelta, timezone
from decimal import Decimal, InvalidOperation
from pathlib import Path
from string import Template
from typing import Any, cast

from PySide6.QtCore import (
    QDate,
    QLocale,
    QMarginsF,
    QObject,
    QPointF,
    QRectF,
    QSettings,
    QSizeF,
    QStringListModel,
    Qt,
    QTimer,
    QUrl,
    Signal,
)
from PySide6.QtGui import (
    QAction,
    QColor,
    QFont,
    QIcon,
    QImage,
    QKeySequence,
    QPageLayout,
    QPageSize,
    QPainter,
    QPdfWriter,
    QPalette,
    QPen,
    QShortcut,
    QTextDocument,
    QTextOption,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QCompleter,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFrame,
    QGraphicsDropShadowEffect,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStackedWidget,
    QTableWidget,
    QTableWidgetItem,
    QToolTip,
    QVBoxLayout,
    QWidget,
)

try:
    import pythoncom
    import win32com.client as win32
except ImportError:
    win32 = None
    pythoncom = None


APP_NAME = "Project Expense Entry"
APP_VERSION = "3.0.0"
REQUIRED_HEADERS = {"id", "description", "date", "amount"}
PASSWORD_HASH_KEY = "security/password_hash"
LEGACY_PASSWORD_HASH = "40063e149099fc2acbc096267ba501114be015779ce8121f8b639503648e8dc3"
SETTINGS_ORG = "BuildOffice"
SETTINGS_APP = "ProjectExpenseEntry"
CURRENCY = "DZD"


def hash_password(value: str) -> str:
    return hashlib.sha256(value.strip().encode("utf-8")).hexdigest()


def get_password_hash() -> str:
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    stored_hash = settings.value(PASSWORD_HASH_KEY, "", type=str)

    # An explicitly supplied environment password is a recovery/override path.
    # Check it BEFORE the saved value so a stale Windows registry hash cannot
    # permanently lock the user out. It updates the saved hash for later launches.
    env_password = os.getenv("PROJECT_EXPENSE_PASSWORD", "")
    if env_password.strip():
        password_hash = hash_password(env_password)
        if stored_hash != password_hash:
            settings.setValue(PASSWORD_HASH_KEY, password_hash)
            settings.sync()
        return password_hash

    if isinstance(stored_hash, str) and stored_hash:
        return stored_hash

    settings.setValue(PASSWORD_HASH_KEY, LEGACY_PASSWORD_HASH)
    settings.sync()
    return LEGACY_PASSWORD_HASH


def set_password(new_password: str) -> None:
    password = (new_password or "").strip()
    if not password:
        raise ValueError("Password cannot be empty.")
    settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
    settings.setValue(PASSWORD_HASH_KEY, hash_password(password))
    settings.sync()
    if settings.status() != QSettings.Status.NoError:
        raise OSError("Windows could not save the password setting.")


def saved_language() -> str:
    value = QSettings(SETTINGS_ORG, SETTINGS_APP).value("language", "ar", type=str)
    return value if isinstance(value, str) and value in ("ar", "en", "fr") else "ar"


def saved_theme() -> str:
    value = QSettings(SETTINGS_ORG, SETTINGS_APP).value("theme", "dark", type=str)
    return value if value in ("dark", "light") else "dark"


# The whole interface (and the exported reports) keep a left-to-right layout in every language.
# Arabic *text* is still shaped and ordered correctly inside each cell/field; only the mirroring of
# the layout is switched off. Set to False to get the old fully mirrored Arabic layout back.
FORCE_LTR_LAYOUT = True


def is_rtl(lang: str) -> bool:
    return lang == "ar" and not FORCE_LTR_LAYOUT


def layout_direction_for(lang: str):
    """Left-to-right everywhere, unless FORCE_LTR_LAYOUT is turned off (then Arabic mirrors)."""
    return Qt.LayoutDirection.RightToLeft if is_rtl(lang) else Qt.LayoutDirection.LeftToRight


def latin_locale() -> QLocale:
    """Latin digits and dd.MM.yyyy ordering regardless of the Windows regional settings."""
    return QLocale(QLocale.Language.English, QLocale.Country.UnitedKingdom)


TEXT = {
    "en": {
        "app": "Project Expense Entry",
        "projects": "Projects",
        "folder_none": "No folder selected",
        "choose_folder": "Choose folder",
        "refresh": "Refresh",
        "search": "Search projects...",
        "new_project": "New project",
        "open_excel": "Open in Excel",
        "backup": "Backup",
        "dashboard": "Dashboard",
        "recent": "Recent expenses",
        "add_expense": "Add expense",
        "save_changes": "Save Changes",
        "edit_expense": "Edit expense",
        "delete_expense": "Delete expense",
        "description": "Description",
        "date": "Date",
        "amount": "Amount",
        "id": "ID",
        "save": "Save",
        "cancel": "Cancel",
        "selected_project": "Select a project file",
        "no_project": "Select a project file first.",
        "missing_description": "Enter a description.",
        "invalid_amount": "Amount must be a valid positive number.",
        "added": "Expense added successfully.",
        "edited": "Expense updated successfully.",
        "deleted": "Expense deleted successfully.",
        "error": "Error",
        "saved": "Saved",
        "confirm_delete": "Delete this expense? This action will modify the Excel file.",
        "backup_created": "Backup created before modification.",
        "table_info": "Table: {table} • Sheet: {sheet} • {rows} expense row(s)",
        "files_found": "{count} macro-enabled project file(s) found.",
        "file_check_failed": "File check failed: {error}",
        "language": "Language",
        "total": "Total",
        "today": "Today",
        "month": "This month",
        "rows": "Rows",
        "filter_month": "Month",
        "all_months": "All months",
        "create_title": "Create new project",
        "project_name": "Project file name",
        "template": "Template",
        "select_template": "Select template",
        "created": "Project created successfully.",
        "template_help": "Choose an existing .xlsm file as the structure/template. Its table and VBA are preserved, then existing expense rows are removed from the new copy.",
        "name_required": "Enter a project file name.",
        "template_required": "Select a template .xlsm file.",
        "already_exists": "A project with this name already exists.",
        "backup_ok": "Backup saved to: {path}",
        "no_rows": "No expense rows.",
        "open_error": "Could not open the workbook.",
        "language_changed": "Language changed. Restart the application to fully refresh all text.",
        "note": "ID is generated by the workbook's existing VBA. The app never writes the ID.",
    },
    "fr": {
        "app": "Saisie des dépenses de projets",
        "projects": "Projets",
        "folder_none": "Aucun dossier sélectionné",
        "choose_folder": "Choisir le dossier",
        "refresh": "Actualiser",
        "search": "Rechercher un projet...",
        "new_project": "Nouveau projet",
        "open_excel": "Ouvrir dans Excel",
        "backup": "Sauvegarde",
        "dashboard": "Tableau de bord",
        "recent": "Dépenses récentes",
        "add_expense": "Ajouter une dépense",
        "save_changes": "Enregistrer les modifications",
        "edit_expense": "Modifier la dépense",
        "delete_expense": "Supprimer la dépense",
        "description": "Description",
        "date": "Date",
        "amount": "Montant",
        "id": "ID",
        "save": "Enregistrer",
        "cancel": "Annuler",
        "selected_project": "Sélectionnez un fichier projet",
        "no_project": "Sélectionnez d'abord un projet.",
        "missing_description": "Saisissez une description.",
        "invalid_amount": "Le montant doit être un nombre positif valide.",
        "added": "Dépense ajoutée avec succès.",
        "edited": "Dépense modifiée avec succès.",
        "deleted": "Dépense supprimée avec succès.",
        "error": "Erreur",
        "saved": "Enregistré",
        "confirm_delete": "Supprimer cette dépense ? Le fichier Excel sera modifié.",
        "backup_created": "Sauvegarde créée avant modification.",
        "table_info": "Tableau : {table} • Feuille : {sheet} • {rows} ligne(s)",
        "files_found": "{count} fichier(s) projet .xlsm trouvé(s).",
        "file_check_failed": "Vérification échouée : {error}",
        "language": "Langue",
        "total": "Total",
        "today": "Aujourd'hui",
        "month": "Ce mois",
        "rows": "Lignes",
        "filter_month": "Mois",
        "all_months": "Tous les mois",
        "create_title": "Créer un nouveau projet",
        "project_name": "Nom du fichier projet",
        "template": "Modèle",
        "select_template": "Choisir le modèle",
        "created": "Projet créé avec succès.",
        "template_help": "Choisissez un fichier .xlsm existant comme modèle. Son tableau et son VBA sont conservés, puis les lignes de dépenses sont supprimées dans la nouvelle copie.",
        "name_required": "Saisissez un nom de projet.",
        "template_required": "Sélectionnez un modèle .xlsm.",
        "already_exists": "Un projet portant ce nom existe déjà.",
        "backup_ok": "Sauvegarde enregistrée : {path}",
        "no_rows": "Aucune dépense.",
        "open_error": "Impossible d'ouvrir le classeur.",
        "language_changed": "Langue modifiée. Redémarrez l'application pour actualiser tous les textes.",
        "note": "L'ID est généré par le VBA existant du classeur. L'application n'écrit jamais l'ID.",
    },
    "ar": {
        "app": "إدخال مصاريف المشاريع",
        "projects": "المشاريع",
        "folder_none": "لم يتم اختيار مجلد",
        "choose_folder": "اختيار المجلد",
        "refresh": "تحديث",
        "search": "البحث عن مشروع...",
        "new_project": "مشروع جديد",
        "open_excel": "فتح في Excel",
        "backup": "نسخة احتياطية",
        "dashboard": "لوحة التحكم",
        "recent": "آخر المصاريف",
        "add_expense": "إضافة مصروف",
        "save_changes": "حفظ التغييرات",
        "edit_expense": "تعديل المصروف",
        "delete_expense": "حذف المصروف",
        "description": "الوصف",
        "date": "التاريخ",
        "amount": "المبلغ",
        "id": "الرقم",
        "save": "حفظ",
        "cancel": "إلغاء",
        "selected_project": "اختر ملف مشروع",
        "no_project": "اختر ملف مشروع أولًا.",
        "missing_description": "أدخل وصف المصروف.",
        "invalid_amount": "يجب أن يكون المبلغ رقمًا موجبًا صحيحًا.",
        "added": "تمت إضافة المصروف بنجاح.",
        "edited": "تم تعديل المصروف بنجاح.",
        "deleted": "تم حذف المصروف بنجاح.",
        "error": "خطأ",
        "saved": "تم الحفظ",
        "confirm_delete": "هل تريد حذف هذا المصروف؟ سيتم تعديل ملف Excel.",
        "backup_created": "تم إنشاء نسخة احتياطية قبل التعديل.",
        "table_info": "الجدول: {table} • الورقة: {sheet} • {rows} مصروف",
        "files_found": "تم العثور على {count} ملف مشروع بصيغة XLSM.",
        "file_check_failed": "فشل فحص الملف: {error}",
        "language": "اللغة",
        "total": "الإجمالي",
        "today": "اليوم",
        "month": "هذا الشهر",
        "rows": "العمليات",
        "filter_month": "الشهر",
        "all_months": "كل الأشهر",
        "create_title": "إنشاء مشروع جديد",
        "project_name": "اسم ملف المشروع",
        "template": "القالب",
        "select_template": "اختيار القالب",
        "created": "تم إنشاء المشروع بنجاح.",
        "template_help": "اختر ملف XLSM موجودًا كقالب. سيتم الحفاظ على الجدول والـ VBA، ثم حذف صفوف المصاريف من النسخة الجديدة.",
        "name_required": "أدخل اسم المشروع.",
        "template_required": "اختر ملف XLSM كقالب.",
        "already_exists": "يوجد مشروع بهذا الاسم بالفعل.",
        "backup_ok": "تم حفظ النسخة الاحتياطية في: {path}",
        "no_rows": "لا توجد مصاريف.",
        "open_error": "تعذر فتح ملف Excel.",
        "language_changed": "تم تغيير اللغة. أعد تشغيل التطبيق لتحديث كل النصوص بالكامل.",
        "note": "الـ ID يتم إنشاؤه بواسطة VBA الموجود في الملف. التطبيق لا يكتب الـ ID مطلقًا.",
    },
}


class PasswordResetDialog(QDialog):
    """Change the local app password when started with --reset-password."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Reset application password")
        self.setMinimumWidth(390)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        title = QLabel("Set a new password")
        title.setStyleSheet("font-size:20px; font-weight:800; color:#ffffff;")
        layout.addWidget(title)
        hint = QLabel("This replaces the password saved for this Windows user.")
        hint.setWordWrap(True)
        hint.setStyleSheet("color:#9fb3cf;")
        layout.addWidget(hint)

        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("New password")
        self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.password_input.setClearButtonEnabled(True)
        layout.addWidget(self.password_input)

        self.confirm_input = QLineEdit()
        self.confirm_input.setPlaceholderText("Confirm new password")
        self.confirm_input.setEchoMode(QLineEdit.EchoMode.Password)
        self.confirm_input.setClearButtonEnabled(True)
        self.confirm_input.returnPressed.connect(self._accept_if_valid)
        layout.addWidget(self.confirm_input)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(self._accept_if_valid)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)
        self.setStyleSheet(STYLE)
        self.password_input.setFocus()

    def _accept_if_valid(self):
        password = self.password_input.text()
        confirmation = self.confirm_input.text()
        if not password.strip():
            QMessageBox.warning(self, "Password required", "Enter a non-empty password.")
            self.password_input.setFocus()
            return
        if password != confirmation:
            QMessageBox.warning(self, "Passwords do not match", "Enter the same password in both fields.")
            self.confirm_input.clear()
            self.confirm_input.setFocus()
            return
        self.accept()

    def password(self) -> str:
        return self.password_input.text()


class LoginWindow(QWidget):

  def __init__(self):
    super().__init__()
    self.lang = saved_language()
    self.setWindowTitle(tr("login_window", self.lang))
    self.resize(420, 520)
    self.setMinimumWidth(360)
    self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
    self.setStyleSheet("""
        QWidget {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #0b1220, stop:1 #101b2f);
            color: #e6edf6;
            font-family: "Segoe UI", Arial, sans-serif;
        }

        #loginShell {
            background: #111b2e;
            border: 1px solid #223452;
            border-radius: 20px;
        }

        QLabel#title {
            color: #ffffff;
            font-size: 30px;
            font-weight: 800;
        }

        QLabel#subtitle {
            color: #8aa0bd;
            font-size: 12px;
            font-weight: 700;
            letter-spacing: 1.5px;
            text-transform: uppercase;
        }

        QLabel#securityNote {
            color: #9fb3cf;
            font-size: 12px;
        }

        QLineEdit {
            background: #0b1424;
            border: 1px solid #2a3d5c;
            border-radius: 10px;
            padding: 12px 12px;
            font-size: 15px;
            color: #e6edf6;
            min-height: 50px;
        }

        QLineEdit:hover { border-color: #3b5580; }
        QLineEdit:focus { border: 1px solid #34d399; background: #0d182b; }
        QLineEdit::placeholder { color: #6f86a6; }

        QPushButton {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #1a2a44, stop:1 #223e5d);
            border: 1px solid #2f4468;
            border-radius: 10px;
            min-height: 52px;
            color: #ffffff;
            font-size: 15px;
            font-weight: 700;
        }

        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #223656, stop:1 #2b4d77);
            border-color: #41598a;
        }

        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #16253c, stop:1 #1a3150);
        }
    """)
    self.init_ui()

  def init_ui(self):
    outer = QVBoxLayout(self)
    outer.setContentsMargins(28, 28, 28, 28)
    outer.setSpacing(18)

    shell = QFrame(self)
    shell.setObjectName("loginShell")
    shell.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
    shell_layout = QVBoxLayout(shell)
    shell_layout.setContentsMargins(28, 26, 28, 24)
    shell_layout.setSpacing(18)

    glow = QGraphicsDropShadowEffect(self)
    glow.setBlurRadius(28)
    glow.setOffset(0, 10)
    glow.setColor(self.palette().color(self.backgroundRole()).darker(180))
    shell.setGraphicsEffect(glow)

    brand = QLabel("Project Expense Entry")
    brand.setObjectName("subtitle")
    brand.setAlignment(Qt.AlignmentFlag.AlignCenter)
    shell_layout.addWidget(brand)

    title = QLabel(tr("login_welcome", self.lang))
    title.setObjectName("title")
    title.setAlignment(Qt.AlignmentFlag.AlignCenter)
    shell_layout.addWidget(title)

    self.label = QLabel(tr("login_prompt", self.lang))
    self.label.setObjectName("securityNote")
    self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
    self.label.setWordWrap(True)
    shell_layout.addWidget(self.label)

    self.password_input = QLineEdit(self)
    self.password_input.setPlaceholderText(tr("login_password", self.lang))
    self.password_input.setEchoMode(QLineEdit.EchoMode.Password)
    self.password_input.setClearButtonEnabled(True)
    self.password_input.returnPressed.connect(self.check_password)
    shell_layout.addWidget(self.password_input)

    self.login_button = QPushButton(tr("login_unlock", self.lang), self)
    self.login_button.clicked.connect(self.check_password)
    self.login_button.setDefault(True)
    shell_layout.addWidget(self.login_button)

    outer.addWidget(shell)
    self.password_input.setFocus()

  def check_password(self):
    entered_code = self.password_input.text().strip()
    if hmac.compare_digest(hash_password(entered_code), get_password_hash()):
      self.accept_login()
    else:
      QMessageBox.warning(
          self, tr("login_denied_title", self.lang), tr("login_denied", self.lang)
      )
      self.password_input.clear()
      self.password_input.setFocus()

  def accept_login(self):
    self.hide()
    self.main_window = MainWindow()
    self.main_window.showMaximized()

def tr(key: str, lang: str | None = None, *args: object, **kwargs: Any) -> str:
    language = lang if isinstance(lang, str) and lang in TEXT else "en"
    translations = TEXT.get(language, TEXT["en"])
    text = translations.get(key, key)
    if not isinstance(text, str):
        text = str(text)
    try:
        if kwargs:
            return text.format(**kwargs)
        if args:
            return text.format(*args)
        return text
    except (AttributeError, IndexError, KeyError, TypeError, ValueError):
        return text


def parse_amount(text: str) -> float:
    value = str(text).strip().replace(" ", "")
    if not value:
        raise ValueError("Amount is required")
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif "," in value:
        parts = value.split(",")
        if len(parts) > 2:
            value = "".join(parts)
        elif len(parts[1]) <= 2:
            value = f"{parts[0]}.{parts[1]}"
        else:
            value = "".join(parts)
    elif value.count(".") > 1:
        value = value.replace(".", "")
    try:
        amount = Decimal(value)
    except InvalidOperation as exc:
        raise ValueError("Invalid amount") from exc
    if not amount.is_finite() or amount <= 0:
        raise ValueError("Amount must be a finite positive number")
    return float(amount)


def format_amount(value):
    try:
        return f"{float(value):,.2f}"
    except (TypeError, ValueError):
        return str(value or "")


def format_id(value):
    """Show IDs as plain numbers (Excel returns 12 as 12.0 or 12.00)."""
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        text = value.strip()
        if not text:
            return ""
        if "." in text:
            whole, frac = text.split(".", 1)
            if whole.isdigit() and frac and frac.strip("0") == "":
                return whole
        return text
    try:
        number = float(value)
        if number.is_integer():
            return str(int(number))
    except (TypeError, ValueError):
        pass
    return str(value)


def search_terms(text):
    """Every word typed must appear in the description (any order, case-insensitive)."""
    return str(text or "").casefold().split()


def description_matches(row, terms):
    if not terms:
        return True
    desc = str(row.get("description") or "").casefold()
    return all(t in desc for t in terms)


def collect_matches(files, cache, terms, limit=2000):
    """Search the cached rows of every project file. Returns [(project, path, row)]."""
    out = []
    for path in files:
        entry = cache.get(path)
        if not entry:
            continue
        name = Path(path).stem
        for row in entry[1]:
            if description_matches(row, terms):
                out.append((name, path, row))
                if len(out) >= limit:
                    return out
    return out


def excel_date_text(value):
    if value is None or value == "":
        return ""
    if hasattr(value, "strftime"):
        return value.strftime("%d.%m.%Y")
    return str(value)


def parse_excel_date(value):
    if value is None or value == "":
        return None
    if hasattr(value, "date"):
        try:
            return value.date()
        except (AttributeError, TypeError, ValueError):
            pass
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            serial = float(value)
            if serial >= 0:
                return (datetime(1899, 12, 30) + timedelta(days=serial)).date()
        except (TypeError, ValueError):
            pass
    text = str(value).strip()
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()  # noqa: DTZ007
        except ValueError:
            continue
    return None


def find_table(workbook):
    tables = []
    for ws in workbook.Worksheets:
        for i in range(1, ws.ListObjects.Count + 1):
            tables.append(ws.ListObjects(i))
    if len(tables) != 1:
        raise RuntimeError(f"Expected exactly one Excel Table, but found {len(tables)} tables.")
    return tables[0]


def get_headers(table):
    return {
        str(table.ListColumns(i).Name).strip().casefold(): i
        for i in range(1, table.ListColumns.Count + 1)
    }



# Extra strings added in 2.1.0 (merged into the dictionaries above).
TEXT["en"].update({
    "loading": "Loading project...",
    "saving": "Saving...",
    "language_changed": "Language changed.",
    "hint": "Ctrl+Enter save  |  Esc clear  |  Del delete  |  F5 refresh",
})
TEXT["fr"].update({
    "loading": "Chargement du projet...",
    "saving": "Enregistrement...",
    "language_changed": "Langue modifiée.",
    "hint": "Ctrl+Entrée enregistrer  |  Échap effacer  |  Suppr supprimer  |  F5 actualiser",
})
TEXT["ar"].update({
    "loading": "جارٍ تحميل المشروع...",
    "saving": "جارٍ الحفظ...",
    "language_changed": "تم تغيير اللغة.",
    "hint": "Ctrl+Enter حفظ  |  Esc مسح  |  Del حذف  |  F5 تحديث",
})

TEXT["en"].update({
    "search_desc": "Search descriptions...", "scope_file": "This file", "scope_all": "All files",
    "col_project": "Project", "indexing": "Indexing files {done}/{total}...",
    "results": "{n} result(s) in {m} file(s)", "no_results": "No matching expenses.",
    "filtered": "{n} of {m} rows", "search_all_hint": "Type to search the descriptions of all project files.",
})
TEXT["fr"].update({
    "search_desc": "Rechercher dans les descriptions...", "scope_file": "Ce fichier", "scope_all": "Tous les fichiers",
    "col_project": "Projet", "indexing": "Indexation des fichiers {done}/{total}...",
    "results": "{n} résultat(s) dans {m} fichier(s)", "no_results": "Aucune dépense correspondante.",
    "filtered": "{n} sur {m} lignes", "search_all_hint": "Saisissez pour rechercher dans tous les fichiers projet.",
})
TEXT["ar"].update({
    "search_desc": "ابحث في الأوصاف...", "scope_file": "هذا الملف", "scope_all": "كل الملفات",
    "col_project": "المشروع", "indexing": "فهرسة الملفات {done}/{total}...",
    "results": "{n} نتيجة في {m} ملف", "no_results": "لا توجد مصاريف مطابقة.",
    "filtered": "{n} من {m} صف", "search_all_hint": "اكتب للبحث في أوصاف كل ملفات المشاريع.",
})

# Strings added in 3.0.0 (Pro edition).
TEXT["en"].update({
    "expenses": "Expenses",
    "chart_title": "Monthly spending",
    "chart_empty": "No dated expenses to chart yet.",
    "chart_tip": "{month}: {amount} DZD • {n} transaction(s)",
    "theme_light": "Light mode", "theme_dark": "Dark mode",
    "export": "Export", "export_excel": "Excel workbook (.xlsx)...", "export_pdf": "PDF report...",
    "report_title": "Reports & export",
    "report_help": "Choose a period, then export an Excel workbook or a printable PDF report for this project.",
    "period": "Period", "from": "From", "to": "To",
    "preset_all": "All dates", "preset_this_month": "This month", "preset_last_month": "Last month",
    "preset_this_year": "This year", "preset_last_year": "Last year", "preset_custom": "Custom range",
    "contains": "Description contains",
    "preview_summary": "{n} transaction(s) • Total {total} DZD",
    "btn_excel": "Export Excel", "btn_pdf": "Create PDF",
    "no_data_in_range": "No expenses match this period.",
    "save_excel_title": "Save Excel export", "save_pdf_title": "Save PDF report",
    "export_done": "Excel file saved: {path}", "pdf_done": "PDF report saved: {path}",
    "exporting": "Exporting...", "creating_pdf": "Creating PDF...",
    "report_heading": "Expense report", "report_project": "Project", "report_period": "Period",
    "report_generated": "Generated", "report_filter": "Filter", "report_total": "Total expenses",
    "report_count": "Transactions", "report_average": "Average", "report_largest": "Largest expense",
    "report_monthly": "Monthly summary", "report_details": "Expense details",
    "report_page": "Page {n} / {m}",
    "report_skipped_amounts": "{n} row(s) with a non-numeric amount are listed but not included in the totals.",
    "report_undated_excluded": "{n} row(s) without a valid date were excluded from this period.",
    "col_month": "Month", "col_count": "Count", "col_total": "Total",
    "sheet_expenses": "Expenses", "sheet_monthly": "Monthly summary",
    "backup_now": "Back up now", "restore_backup": "Restore from backup...",
    "open_backup_folder": "Open backup folder",
    "restore_title": "Restore from backup",
    "restore_help": "Select a backup to restore for this project. Your current file is saved as a safety backup first, so a restore can always be undone.",
    "restore_btn": "Restore selected",
    "backup_auto": "Automatic", "backup_manual": "Manual", "backup_prerestore": "Before restore",
    "col_when": "Date & time", "col_type": "Type", "col_size": "Size",
    "no_backups": "No backups found for this project yet.",
    "restore_confirm": "Restore this project from the backup made on {when}?\n\nYour current file will be saved as a safety backup first.",
    "restoring": "Restoring...",
    "restored": "Backup restored. A safety copy of the previous file was kept.",
    "restore_locked": "The file is open in Excel or locked by another program. Close it and try again.",
    "stale_row": "This expense was changed outside the app, so nothing was modified. The list has been refreshed - please try again.",
    "duplicate_title": "Possible duplicate",
    "duplicate_confirm": "An expense with the same description, date and amount already exists in this project.\n\nAdd it anyway?",
    "confirm_delete_detail": "Delete this expense?\n\n{description}\n{date}  •  {amount} DZD\n\nA backup is created first, so it can be restored from Backup > Restore.",
    "unsaved_title": "Unsaved entry",
    "unsaved_confirm": "The form contains an expense that has not been saved. Discard it?",
    "keep_date": "Keep the date after saving",
    "min_amount": "Min amount", "max_amount": "Max amount",
    "filtered_total": "{n} of {m} rows • Total {total} DZD",
    "results_total": "{n} result(s) in {m} file(s) • Total {total} DZD",
    "login_window": "Secure Access", "login_welcome": "Welcome back",
    "login_prompt": "Enter your secure password to continue", "login_password": "Password",
    "login_unlock": "Unlock", "login_denied_title": "Access denied",
    "login_denied": "Incorrect password. Please try again.",
    "hint": "Ctrl+Enter save  |  Esc clear  |  Del delete  |  F5 refresh  |  Ctrl+F search  |  Ctrl+E reports",
})
TEXT["fr"].update({
    "expenses": "Dépenses",
    "chart_title": "Dépenses mensuelles",
    "chart_empty": "Aucune dépense datée à afficher.",
    "chart_tip": "{month} : {amount} DZD • {n} opération(s)",
    "theme_light": "Mode clair", "theme_dark": "Mode sombre",
    "export": "Exporter", "export_excel": "Classeur Excel (.xlsx)...", "export_pdf": "Rapport PDF...",
    "report_title": "Rapports et export",
    "report_help": "Choisissez une période, puis exportez un classeur Excel ou un rapport PDF imprimable pour ce projet.",
    "period": "Période", "from": "Du", "to": "Au",
    "preset_all": "Toutes les dates", "preset_this_month": "Ce mois", "preset_last_month": "Mois dernier",
    "preset_this_year": "Cette année", "preset_last_year": "Année dernière", "preset_custom": "Période personnalisée",
    "contains": "La description contient",
    "preview_summary": "{n} opération(s) • Total {total} DZD",
    "btn_excel": "Exporter Excel", "btn_pdf": "Créer le PDF",
    "no_data_in_range": "Aucune dépense ne correspond à cette période.",
    "save_excel_title": "Enregistrer l'export Excel", "save_pdf_title": "Enregistrer le rapport PDF",
    "export_done": "Fichier Excel enregistré : {path}", "pdf_done": "Rapport PDF enregistré : {path}",
    "exporting": "Export en cours...", "creating_pdf": "Création du PDF...",
    "report_heading": "Rapport des dépenses", "report_project": "Projet", "report_period": "Période",
    "report_generated": "Généré le", "report_filter": "Filtre", "report_total": "Total des dépenses",
    "report_count": "Opérations", "report_average": "Moyenne", "report_largest": "Plus grosse dépense",
    "report_monthly": "Récapitulatif mensuel", "report_details": "Détail des dépenses",
    "report_page": "Page {n} / {m}",
    "report_skipped_amounts": "{n} ligne(s) avec un montant non numérique sont listées mais exclues des totaux.",
    "report_undated_excluded": "{n} ligne(s) sans date valide ont été exclues de cette période.",
    "col_month": "Mois", "col_count": "Nombre", "col_total": "Total",
    "sheet_expenses": "Dépenses", "sheet_monthly": "Récapitulatif mensuel",
    "backup_now": "Sauvegarder maintenant", "restore_backup": "Restaurer une sauvegarde...",
    "open_backup_folder": "Ouvrir le dossier des sauvegardes",
    "restore_title": "Restaurer une sauvegarde",
    "restore_help": "Sélectionnez la sauvegarde à restaurer pour ce projet. Le fichier actuel est d'abord sauvegardé par sécurité : une restauration peut toujours être annulée.",
    "restore_btn": "Restaurer la sélection",
    "backup_auto": "Automatique", "backup_manual": "Manuelle", "backup_prerestore": "Avant restauration",
    "col_when": "Date et heure", "col_type": "Type", "col_size": "Taille",
    "no_backups": "Aucune sauvegarde trouvée pour ce projet.",
    "restore_confirm": "Restaurer ce projet à partir de la sauvegarde du {when} ?\n\nLe fichier actuel sera d'abord sauvegardé par sécurité.",
    "restoring": "Restauration...",
    "restored": "Sauvegarde restaurée. Une copie de sécurité de l'ancien fichier a été conservée.",
    "restore_locked": "Le fichier est ouvert dans Excel ou verrouillé par un autre programme. Fermez-le et réessayez.",
    "stale_row": "Cette dépense a été modifiée en dehors de l'application : rien n'a été changé. La liste a été actualisée, veuillez réessayer.",
    "duplicate_title": "Doublon possible",
    "duplicate_confirm": "Une dépense avec la même description, la même date et le même montant existe déjà dans ce projet.\n\nL'ajouter quand même ?",
    "confirm_delete_detail": "Supprimer cette dépense ?\n\n{description}\n{date}  •  {amount} DZD\n\nUne sauvegarde est créée avant, vous pourrez la restaurer via Sauvegarde > Restaurer.",
    "unsaved_title": "Saisie non enregistrée",
    "unsaved_confirm": "Le formulaire contient une dépense non enregistrée. L'abandonner ?",
    "keep_date": "Conserver la date après l'enregistrement",
    "min_amount": "Montant min.", "max_amount": "Montant max.",
    "filtered_total": "{n} sur {m} lignes • Total {total} DZD",
    "results_total": "{n} résultat(s) dans {m} fichier(s) • Total {total} DZD",
    "login_window": "Accès sécurisé", "login_welcome": "Bon retour",
    "login_prompt": "Saisissez votre mot de passe pour continuer", "login_password": "Mot de passe",
    "login_unlock": "Déverrouiller", "login_denied_title": "Accès refusé",
    "login_denied": "Mot de passe incorrect. Veuillez réessayer.",
    "language_changed": "Langue modifiée.",
    "hint": "Ctrl+Entrée enregistrer  |  Échap effacer  |  Suppr supprimer  |  F5 actualiser  |  Ctrl+F rechercher  |  Ctrl+E rapports",
})
TEXT["en"].update({"open_file": "Open", "show_folder": "Show in folder", "col_share": "Share",
                   "amount_dzd": "Amount (DZD)"})
TEXT["fr"].update({"open_file": "Ouvrir", "show_folder": "Afficher le dossier", "col_share": "Part",
                   "amount_dzd": "Montant (DZD)"})
TEXT["ar"].update({"open_file": "فتح", "show_folder": "إظهار في المجلد", "col_share": "النسبة",
                   "amount_dzd": "المبلغ (DZD)"})
TEXT["ar"].update({
    "expenses": "المصاريف",
    "chart_title": "المصاريف الشهرية",
    "chart_empty": "لا توجد مصاريف مؤرخة لعرضها.",
    "chart_tip": "{month}: {amount} DZD • {n} عملية",
    "theme_light": "الوضع الفاتح", "theme_dark": "الوضع الداكن",
    "export": "تصدير", "export_excel": "ملف Excel (.xlsx)...", "export_pdf": "تقرير PDF...",
    "report_title": "التقارير والتصدير",
    "report_help": "اختر الفترة، ثم صدّر ملف Excel أو تقرير PDF قابلًا للطباعة لهذا المشروع.",
    "period": "الفترة", "from": "من", "to": "إلى",
    "preset_all": "كل التواريخ", "preset_this_month": "هذا الشهر", "preset_last_month": "الشهر الماضي",
    "preset_this_year": "هذه السنة", "preset_last_year": "السنة الماضية", "preset_custom": "فترة مخصصة",
    "contains": "الوصف يحتوي على",
    "preview_summary": "{n} عملية • الإجمالي {total} DZD",
    "btn_excel": "تصدير Excel", "btn_pdf": "إنشاء PDF",
    "no_data_in_range": "لا توجد مصاريف في هذه الفترة.",
    "save_excel_title": "حفظ ملف Excel", "save_pdf_title": "حفظ تقرير PDF",
    "export_done": "تم حفظ ملف Excel: {path}", "pdf_done": "تم حفظ تقرير PDF: {path}",
    "exporting": "جارٍ التصدير...", "creating_pdf": "جارٍ إنشاء PDF...",
    "report_heading": "تقرير المصاريف", "report_project": "المشروع", "report_period": "الفترة",
    "report_generated": "تاريخ الإنشاء", "report_filter": "التصفية", "report_total": "إجمالي المصاريف",
    "report_count": "العمليات", "report_average": "المتوسط", "report_largest": "أكبر مصروف",
    "report_monthly": "الملخص الشهري", "report_details": "تفاصيل المصاريف",
    "report_page": "الصفحة {n} / {m}",
    "report_skipped_amounts": "{n} صف(وف) بمبلغ غير رقمي مدرجة في القائمة لكنها غير محسوبة في الإجمالي.",
    "report_undated_excluded": "تم استبعاد {n} صف(وف) بدون تاريخ صالح من هذه الفترة.",
    "col_month": "الشهر", "col_count": "العدد", "col_total": "الإجمالي",
    "sheet_expenses": "المصاريف", "sheet_monthly": "الملخص الشهري",
    "backup_now": "نسخ احتياطي الآن", "restore_backup": "استرجاع نسخة احتياطية...",
    "open_backup_folder": "فتح مجلد النسخ الاحتياطية",
    "restore_title": "استرجاع نسخة احتياطية",
    "restore_help": "اختر النسخة الاحتياطية التي تريد استرجاعها لهذا المشروع. يتم أولًا حفظ الملف الحالي كنسخة أمان، لذلك يمكن دائمًا التراجع عن الاسترجاع.",
    "restore_btn": "استرجاع المحدد",
    "backup_auto": "تلقائية", "backup_manual": "يدوية", "backup_prerestore": "قبل الاسترجاع",
    "col_when": "التاريخ والوقت", "col_type": "النوع", "col_size": "الحجم",
    "no_backups": "لا توجد نسخ احتياطية لهذا المشروع بعد.",
    "restore_confirm": "هل تريد استرجاع هذا المشروع من النسخة المأخوذة في {when}؟\n\nسيتم أولًا حفظ الملف الحالي كنسخة أمان.",
    "restoring": "جارٍ الاسترجاع...",
    "restored": "تم استرجاع النسخة الاحتياطية. تم الاحتفاظ بنسخة أمان من الملف السابق.",
    "restore_locked": "الملف مفتوح في Excel أو مقفل من برنامج آخر. أغلقه ثم حاول مرة أخرى.",
    "stale_row": "تم تغيير هذا المصروف خارج التطبيق، لذلك لم يتم تعديل أي شيء. تم تحديث القائمة، حاول مرة أخرى.",
    "duplicate_title": "احتمال تكرار",
    "duplicate_confirm": "يوجد بالفعل مصروف بنفس الوصف والتاريخ والمبلغ في هذا المشروع.\n\nهل تريد إضافته على أي حال؟",
    "confirm_delete_detail": "هل تريد حذف هذا المصروف؟\n\n{description}\n{date}  •  {amount} DZD\n\nيتم إنشاء نسخة احتياطية أولًا، ويمكن استرجاعها من نسخ احتياطي > استرجاع.",
    "unsaved_title": "إدخال غير محفوظ",
    "unsaved_confirm": "النموذج يحتوي على مصروف لم يتم حفظه. هل تريد تجاهله؟",
    "keep_date": "الاحتفاظ بالتاريخ بعد الحفظ",
    "min_amount": "الحد الأدنى", "max_amount": "الحد الأقصى",
    "filtered_total": "{n} من {m} صف • الإجمالي {total} DZD",
    "results_total": "{n} نتيجة في {m} ملف • الإجمالي {total} DZD",
    "login_window": "دخول آمن", "login_welcome": "مرحبًا بعودتك",
    "login_prompt": "أدخل كلمة المرور للمتابعة", "login_password": "كلمة المرور",
    "login_unlock": "فتح", "login_denied_title": "تم رفض الدخول",
    "login_denied": "كلمة المرور غير صحيحة. حاول مرة أخرى.",
})


ROLE_INDEX = int(Qt.ItemDataRole.UserRole) + 1  # item data slot for the Excel row number

# False = rows are listed first to last, in the same order as the Excel table.
NEWEST_FIRST = True


def prepare_rows(rows):
    """Parse dates/amounts once, in the worker thread, so the GUI never has to."""
    for r in rows:
        r["_date"] = parse_excel_date(r.get("date"))
        try:
            r["_amount"] = float(r.get("amount") or 0)
            r["_numeric"] = True
        except (TypeError, ValueError):
            r["_amount"] = 0.0
            r["_numeric"] = False
    return rows


# ---------------------------------------------------------------------------
# Pro edition helpers. Everything below is plain Python (no Excel needed), which
# keeps it easy to test; the Qt-specific pieces are the chart painter and PDF writer.
# ---------------------------------------------------------------------------
BACKUP_DIR_NAME = "~ExpenseBackups"
KEEP_RECENT_BACKUPS = 20       # newest automatic backups always kept
KEEP_BACKUP_DAYS = 30          # ...plus the last backup of each day for this many days
KEEP_PRERESTORE_BACKUPS = 5    # safety copies made before a restore
SORT_COLUMNS = ("id", "description", "date", "amount")


class StaleRowError(RuntimeError):
    """The Excel row no longer matches what the user was looking at (file changed elsewhere)."""


def backup_dir_for(path):
    return Path(path).parent / BACKUP_DIR_NAME


def _backup_pattern(stem):
    return re.compile(rf"^{re.escape(stem)}_(\d{{8}})_(\d{{6}})(?:_(\d{{6}}))?(?:_(manual|prerestore))?$")


def list_backups(path):
    """Backups of one project file, newest first. Understands the pre-3.0 file names too."""
    p = Path(path)
    folder = backup_dir_for(p)
    if not folder.is_dir():
        return []
    rx = _backup_pattern(p.stem)
    found = []
    for f in folder.iterdir():
        if not f.is_file() or f.suffix.lower() != p.suffix.lower():
            continue
        m = rx.match(f.stem)
        if not m:
            continue
        try:
            when = datetime.strptime(m.group(1) + m.group(2), "%Y%m%d%H%M%S")  # noqa: DTZ007
            if m.group(3):
                when = when.replace(microsecond=int(m.group(3)))
            size = f.stat().st_size
        except (ValueError, OSError):
            continue
        found.append({"path": str(f), "name": f.name, "when": when, "kind": m.group(4) or "auto", "size": size})
    found.sort(key=lambda b: b["when"], reverse=True)
    return found


def prune_backups(path, now=None):
    """Keep the newest automatic backups plus the last one of every recent day.

    Manual backups are never deleted automatically. Returns the removed paths.
    """
    now = now or datetime.now()  # noqa: DTZ005
    everything = list_backups(path)
    auto = [b for b in everything if b["kind"] == "auto"]
    keep = {b["path"] for b in auto[:KEEP_RECENT_BACKUPS]}
    seen_days = set()
    for b in auto:  # newest first, so the first one seen for a day is that day's last backup
        day = b["when"].date()
        if (now.date() - day).days <= KEEP_BACKUP_DAYS and day not in seen_days:
            seen_days.add(day)
            keep.add(b["path"])
    doomed = [b["path"] for b in auto if b["path"] not in keep]
    doomed += [b["path"] for b in [x for x in everything if x["kind"] == "prerestore"][KEEP_PRERESTORE_BACKUPS:]]
    removed = []
    for victim in doomed:
        with suppress(OSError):
            os.remove(victim)
            removed.append(victim)
    return removed


def make_backup(path, kind="auto"):
    """Copy the project file into ~ExpenseBackups. Returns the backup path.

    An automatic backup is skipped (the existing one is returned) when the newest
    backup is already byte-identical, so repeated or failed attempts do not pile up.
    """
    p = Path(path)
    folder = backup_dir_for(p)
    folder.mkdir(exist_ok=True)
    if kind == "auto":
        existing = list_backups(p)
        if existing:
            newest = Path(existing[0]["path"])
            with suppress(OSError):
                if newest.stat().st_size == p.stat().st_size and filecmp.cmp(p, newest, shallow=False):
                    return str(newest)
    stamp = datetime.now(tz=timezone.utc).astimezone()
    tag = "" if kind == "auto" else f"_{kind}"
    target = folder / f"{p.stem}_{stamp:%Y%m%d_%H%M%S_%f}{tag}{p.suffix}"
    shutil.copy2(p, target)
    if target.stat().st_size != p.stat().st_size:
        with suppress(OSError):
            target.unlink()
        raise OSError("Backup verification failed: the copy has a different size than the original.")
    if kind == "auto":
        with suppress(Exception):
            prune_backups(p)
    return str(target)


def restore_backup(project_path, backup_path):
    """Replace the project file with a backup. The current file is backed up first.

    Returns the path of that safety copy (or None when there was no current file).
    """
    project, backup = Path(project_path), Path(backup_path)
    if not backup.is_file():
        raise FileNotFoundError(str(backup))
    if not zipfile.is_zipfile(backup):
        raise ValueError("The selected backup is not a valid Excel workbook.")
    safety = make_backup(project, "prerestore") if project.exists() else None
    staging = project.with_name(project.name + ".restore.tmp")
    shutil.copyfile(backup, staging)
    try:
        os.replace(staging, project)
    except OSError:
        with suppress(OSError):
            staging.unlink()
        raise
    return safety


def format_size(size):
    value = float(size)
    for unit in ("B", "KB", "MB"):
        if value < 1024:
            return f"{value:.0f} {unit}" if unit == "B" else f"{value:.1f} {unit}"
        value /= 1024
    return f"{value:.1f} GB"


def safe_filename(name):
    cleaned = re.sub(r'[\\/:*?"<>|]+', "_", str(name)).strip(" .")
    return cleaned or "report"


# --- comparing / filtering / sorting rows -----------------------------------
def _norm_text(value):
    return "" if value is None else str(value).strip()


def rows_match(expected, actual):
    """True when an Excel row still holds what the user loaded (description, date, amount)."""
    if _norm_text(expected.get("description")) != _norm_text(actual.get("description")):
        return False
    try:
        if abs(float(expected.get("amount") or 0) - float(actual.get("amount") or 0)) > 0.005:  # noqa: PLR2004
            return False
    except (TypeError, ValueError):
        if _norm_text(expected.get("amount")) != _norm_text(actual.get("amount")):
            return False
    return parse_excel_date(expected.get("date")) == parse_excel_date(actual.get("date"))


def parse_filter_amount(text):
    """Optional amount filter: None when empty/invalid, zero allowed."""
    value = str(text or "").strip()
    if not value:
        return None
    with suppress(ValueError):
        return parse_amount(value)
    return 0.0 if set(value) <= set("0.,") else None


def month_key(day):
    return day.strftime("%Y-%m")


def amount_in_range(row, min_amount=None, max_amount=None):
    if min_amount is None and max_amount is None:
        return True
    if not row.get("_numeric"):
        return False
    amount = row.get("_amount", 0.0)
    return not ((min_amount is not None and amount < min_amount) or (max_amount is not None and amount > max_amount))


def filter_rows(rows, month="all", terms=(), min_amount=None, max_amount=None):
    out = []
    for r in rows:
        if month != "all":
            d = r.get("_date")
            if not (d and month_key(d) == month):
                continue
        if not description_matches(r, terms):
            continue
        if not amount_in_range(r, min_amount, max_amount):
            continue
        out.append(r)
    return out


def sort_key(row, column):
    if column == "id":
        text = format_id(row.get("id"))
        try:
            return (0, float(text), "")
        except ValueError:
            return (1, 0.0, text.casefold())
    if column == "description":
        return (0, 0.0, _norm_text(row.get("description")).casefold())
    if column == "date":
        d = row.get("_date")
        return (0, float(d.toordinal()), "") if d else (1, 0.0, "")
    return (0, float(row.get("_amount", 0.0)), "")


def sort_rows(rows, column, ascending=True):
    return sorted(rows, key=lambda r: (sort_key(r, column), r.get("excel_index", 0)), reverse=not ascending)


def unique_descriptions(rows, limit=400):
    """Most recent descriptions first (rows are in Excel order, newest at the bottom)."""
    seen, out = set(), []
    for r in reversed(rows):
        text = _norm_text(r.get("description"))
        if text and text.casefold() not in seen:
            seen.add(text.casefold())
            out.append(text)
            if len(out) >= limit:
                break
    return out


def find_duplicate(rows, description, day, amount):
    wanted = _norm_text(description).casefold()
    for r in rows:
        if (
            _norm_text(r.get("description")).casefold() == wanted
            and r.get("_date") == day
            and r.get("_numeric")
            and abs(r.get("_amount", 0.0) - amount) < 0.005  # noqa: PLR2004
        ):
            return r
    return None


# --- statistics ---------------------------------------------------------------
def aggregate_months(rows):
    out = {}
    for r in rows:
        d = r.get("_date")
        if d and r.get("_numeric"):
            total, count = out.get(month_key(d), (0.0, 0))
            out[month_key(d)] = (total + r["_amount"], count + 1)
    return out


def chart_months(rows, today=None, span=12):
    """The last `span` months as [(YYYY-MM, total, count)], zero months included."""
    today = today or date.today()  # noqa: DTZ011
    agg = aggregate_months(rows)
    if not agg:
        return []
    latest = max(agg)
    last_idx = int(latest[:4]) * 12 + int(latest[5:7]) - 1
    cur_idx = today.year * 12 + today.month - 1
    end = cur_idx if 0 <= cur_idx - last_idx < span else last_idx
    out = []
    for i in range(end - span + 1, end + 1):
        year, month0 = divmod(i, 12)
        key = f"{year:04d}-{month0 + 1:02d}"
        total, count = agg.get(key, (0.0, 0))
        out.append((key, total, count))
    return out


def summarize_rows(rows):
    total, count, invalid, largest = 0.0, 0, 0, None
    for r in rows:
        if r.get("_numeric"):
            total += r["_amount"]
            count += 1
            if largest is None or r["_amount"] > largest["_amount"]:
                largest = r
        else:
            invalid += 1
    return {"total": total, "count": count, "invalid": invalid,
            "average": (total / count) if count else 0.0, "largest": largest}


def compact_amount(value):
    value = float(value)
    if value >= 1_000_000:  # noqa: PLR2004
        return f"{value / 1_000_000:.1f}M"
    if value >= 1_000:  # noqa: PLR2004
        return f"{value / 1_000:.1f}k"
    return f"{value:.0f}"


def month_label(key):
    return f"{key[5:7]}/{key[2:4]}"


# --- report building ----------------------------------------------------------
def last_day_of_month(day):
    return (day.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)


def preset_range(key, today, min_day, max_day):
    if key == "this_month":
        return today.replace(day=1), last_day_of_month(today)
    if key == "last_month":
        prev_last = today.replace(day=1) - timedelta(days=1)
        return prev_last.replace(day=1), prev_last
    if key == "this_year":
        return date(today.year, 1, 1), date(today.year, 12, 31)
    if key == "last_year":
        return date(today.year - 1, 1, 1), date(today.year - 1, 12, 31)
    return min_day, max_day


def report_rows(rows, day_from=None, day_to=None, text="", all_dates=True):
    """Rows for a report, oldest first. A date range never includes rows with no valid date."""
    terms = search_terms(text)
    out = []
    for r in rows:
        if not description_matches(r, terms):
            continue
        if not all_dates:
            d = r.get("_date")
            if d is None or d < day_from or d > day_to:
                continue
        out.append(r)
    out.sort(key=lambda r: (r.get("_date") is None, r.get("_date") or date.min, r.get("excel_index", 0)))
    return out


def build_report_payload(project, all_rows, day_from, day_to, text, all_dates, lang):
    rows = report_rows(all_rows, day_from, day_to, text, all_dates)
    dated = [r["_date"] for r in rows if r.get("_date")]
    if all_dates:
        period = (f"{tr('preset_all', lang)} ({excel_date_text(min(dated))} - {excel_date_text(max(dated))})"
                  if dated else tr("preset_all", lang))
    else:
        period = f"{day_from:%d.%m.%Y} - {day_to:%d.%m.%Y}"
    undated = sum(1 for r in all_rows if r.get("_date") is None and description_matches(r, search_terms(text)))
    agg = aggregate_months(rows)
    return {
        "project": project, "rows": rows, "summary": summarize_rows(rows),
        "months": [(k, agg[k][0], agg[k][1]) for k in sorted(agg)],
        "period": period, "filter": text.strip(), "all_dates": all_dates,
        "undated_excluded": 0 if all_dates else undated,
        "generated": datetime.now().strftime("%d.%m.%Y %H:%M"),  # noqa: DTZ005
    }


def _fmt(value):
    return format_amount(value)


def build_report_html(payload, lang, chart_size=None):
    """Printable report as Qt rich text (only the HTML subset Qt understands)."""
    rtl = is_rtl(lang)
    esc = html.escape
    t = lambda k, **kw: tr(k, lang, **kw)  # noqa: E731
    start, end = ("right", "left") if rtl else ("left", "right")
    s = payload["summary"]
    accent, ink, muted, band, zebra = "#16324f", "#0f172a", "#64748b", "#e6ecf5", "#f6f8fb"
    soft, total_bg = "#eef3fa", "#d9e5f3"
    parts = [f'<body style="font-family:\'Segoe UI\',Tahoma,Arial; color:{ink};">']

    # -- title banner: report name, project, period
    sub = (f'<span style="font-size:12pt; font-weight:bold; color:#ffffff;">{esc(payload["project"])}</span>'
           f'<span style="font-size:9pt; color:#b9cde4;">&nbsp;&nbsp;|&nbsp;&nbsp;{esc(payload["period"])}</span>')
    parts.append(
        f'<table width="100%" cellspacing="0" cellpadding="14"><tr><td bgcolor="{accent}" align="{start}">'
        f'<span style="font-size:8pt; color:#8fb0d3; letter-spacing:1px;">{esc(APP_NAME.upper())}</span><br/>'
        f'<span style="font-size:20pt; font-weight:bold; color:#ffffff;">{esc(t("report_heading"))}</span><br/>{sub}'
        "</td></tr></table>"
    )

    meta = [(t("report_generated"), payload["generated"])]
    if payload["filter"]:
        meta.append((t("report_filter"), f'"{payload["filter"]}"'))
    meta_html = "&nbsp;&nbsp;&nbsp;&nbsp;".join(
        f'<span style="color:{muted};">{esc(k)}:</span> <b>{esc(v)}</b>' for k, v in meta
    )
    parts.append(f'<p align="{start}" style="margin-top:6px; margin-bottom:2px; font-size:8.5pt;">{meta_html}</p>')

    # -- KPI cards
    largest = s["largest"]
    largest_text = f"{_fmt(largest['_amount'])}" if largest else "-"
    largest_sub = esc(_norm_text(largest.get("description"))[:38]) if largest else "&nbsp;"
    kpis = [
        (t("report_total"), _fmt(s["total"]), CURRENCY),
        (t("report_count"), str(s["count"]), "&nbsp;"),
        (t("report_average"), _fmt(s["average"]), CURRENCY),
        (t("report_largest"), largest_text, largest_sub),
    ]
    if rtl:
        kpis.reverse()
    cells = "".join(
        f'<td width="25%" bgcolor="{soft}" align="{start}">'
        f'<span style="font-size:8pt; color:{muted};">{esc(label)}</span><br/>'
        f'<span style="font-size:14pt; font-weight:bold; color:{accent};">{value}</span><br/>'
        f'<span style="font-size:7.5pt; color:{muted};">{sub_text}</span></td>'
        for label, value, sub_text in kpis
    )
    parts.append(f'<table width="100%" cellspacing="8" cellpadding="9"><tr>{cells}</tr></table>')

    def heading(text):
        return (f'<p align="{start}" style="margin-top:16px; margin-bottom:5px; font-size:11.5pt; '
                f'font-weight:bold; color:{accent};">{esc(text)}</p>')

    def table(head, body, widths, aligns, total_row=True):
        order = list(range(len(head)))
        if rtl:
            order.reverse()
        th = "".join(
            f'<th width="{widths[i]}" bgcolor="{accent}" align="{aligns[i]}">'
            f'<span style="color:#ffffff;">{esc(head[i])}</span></th>' for i in order
        )
        trs = []
        last = len(body) - 1
        for n, row in enumerate(body):
            if total_row and n == last:
                bg = f' bgcolor="{total_bg}"'
            else:
                bg = f' bgcolor="{zebra}"' if n % 2 else ""
            tds = "".join(f'<td{bg} align="{aligns[i]}">{row[i]}</td>' for i in order)
            trs.append(f"<tr>{tds}</tr>")
        return (f'<table width="100%" cellspacing="0" cellpadding="6"><thead><tr>{th}</tr></thead>'
                f'{"".join(trs)}</table>')

    if chart_size and payload["months"]:
        parts.append(heading(t("chart_title")))
        parts.append(f'<p align="center"><img src="chart.png" width="{chart_size[0]}" height="{chart_size[1]}"/></p>')

    if payload["months"]:
        parts.append(heading(t("report_monthly")))
        grand = s["total"] or 0.0
        body = [[esc(k), str(c), _fmt(v), f"{(v / grand * 100) if grand else 0:.1f}%"] for k, v, c in payload["months"]]
        body.append([f"<b>{esc(t('col_total'))}</b>", f"<b>{s['count']}</b>", f"<b>{_fmt(s['total'])}</b>",
                     "<b>100%</b>" if grand else ""])
        parts.append(table([t("col_month"), t("col_count"), t("col_total"), t("col_share")], body,
                           ["28%", "18%", "34%", "20%"], [start, "center", end, end]))

    parts.append(heading(t("report_details")))
    body = []
    for r in payload["rows"]:
        body.append([
            esc(format_id(r.get("id"))),
            esc(_norm_text(r.get("description"))),
            esc(excel_date_text(r.get("date"))),
            _fmt(r.get("amount")),
        ])
    body.append(["", f"<b>{esc(t('col_total'))}</b>", "", f"<b>{_fmt(s['total'])} {CURRENCY}</b>"])
    parts.append(table([t("id"), t("description"), t("date"), t("amount_dzd")], body, ["9%", "51%", "16%", "24%"],
                       ["center", start, "center", end]))

    notes = []
    if s["invalid"]:
        notes.append(t("report_skipped_amounts", n=s["invalid"]))
    if payload["undated_excluded"]:
        notes.append(t("report_undated_excluded", n=payload["undated_excluded"]))
    for note in notes:
        parts.append(f'<p align="{start}" style="font-size:8pt; color:{muted};">{esc(note)}</p>')
    parts.append("</body>")
    return "".join(parts)


def paint_monthly_chart(p, rect, data, colors, selected="all", hover=-1):
    """Draw [(YYYY-MM, total, count)] as bars inside rect. Returns one hit-rect per bar."""
    hits = []
    n = len(data)
    plot = QRectF(rect.left() + 6, rect.top() + 18, rect.width() - 12, rect.height() - 38)
    if n == 0 or plot.width() <= 0 or plot.height() <= 0:
        return hits
    peak = max(total for _k, total, _c in data)
    slot = plot.width() / n
    bar_w = max(4.0, min(slot * 0.62, 46.0))
    base_y = plot.bottom()
    line = QPen(QColor(colors["grid"]))
    line.setWidthF(1.0)
    guide = QColor(colors["grid"])
    guide.setAlpha(90)
    guide_pen = QPen(guide)
    guide_pen.setWidthF(1.0)
    guide_pen.setStyle(Qt.PenStyle.DotLine)
    p.setPen(guide_pen)
    for frac in (0.5, 1.0):
        gy = base_y - plot.height() * frac
        p.drawLine(QPointF(plot.left(), gy), QPointF(plot.right(), gy))
    p.setPen(line)
    p.drawLine(QPointF(plot.left(), base_y), QPointF(plot.right(), base_y))
    font = QFont(p.font())
    font.setPointSizeF(7.5)
    p.setFont(font)
    center = int(Qt.AlignmentFlag.AlignCenter.value)
    for i, (key, total, _count) in enumerate(data):
        cx = plot.left() + slot * (i + 0.5)
        height = 0.0 if peak <= 0 or total <= 0 else max(3.0, plot.height() * total / peak)
        bar = QRectF(cx - bar_w / 2, base_y - height, bar_w, height)
        hits.append(QRectF(plot.left() + slot * i, rect.top(), slot, rect.height()))
        color = colors["bar_sel"] if key == selected else colors["bar_hover"] if i == hover else colors["bar"]
        p.setPen(Qt.PenStyle.NoPen)
        if total > 0:
            p.setBrush(QColor(color))
            p.drawRoundedRect(bar, 3, 3)
            if slot >= 40 or key == selected or i == hover:  # noqa: PLR2004
                p.setPen(QColor(colors["text"]))
                p.drawText(QRectF(cx - slot / 2, bar.top() - 15, slot, 14), center, compact_amount(total))
        else:
            p.setBrush(QColor(colors["grid"]))
            p.drawRect(QRectF(cx - bar_w / 2, base_y - 2, bar_w, 2))
        p.setPen(QColor(colors["text"] if key == selected else colors["muted"]))
        p.drawText(QRectF(cx - slot / 2, base_y + 4, slot, 14), center, month_label(key))
    p.setBrush(Qt.BrushStyle.NoBrush)
    return hits


def render_chart_image(data, width, height):
    """Light-themed chart as an image for the PDF report (drawn at 2x for sharpness)."""
    image = QImage(width * 2, height * 2, QImage.Format.Format_ARGB32)
    image.fill(QColor("#ffffff"))
    painter = QPainter(image)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.scale(2, 2)
    paint_monthly_chart(painter, QRectF(0, 0, width, height), data, THEMES["print"])
    painter.end()
    return image


def write_pdf_report(path, payload, lang):
    """Paginated A4 PDF with repeating table headers and page numbers. Returns the page count."""
    rtl = is_rtl(lang)
    resolution = 144
    writer = QPdfWriter(str(path))
    writer.setResolution(resolution)
    writer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
    writer.setPageMargins(QMarginsF(14, 14, 14, 12), QPageLayout.Unit.Millimeter)
    writer.setTitle(f"{tr('report_heading', lang)} - {payload['project']}")
    writer.setCreator(APP_NAME)
    paint = writer.pageLayout().paintRectPixels(resolution)
    body_w, footer_h = float(paint.width()), 30.0
    body_h = float(paint.height()) - footer_h

    doc = QTextDocument()
    custom_pages = True
    try:
        doc.documentLayout().setPaintDevice(writer)
    except Exception:  # noqa: BLE001 - very old/odd Qt builds: fall back to Qt's own printing (no footer)
        custom_pages = False
    font = QFont("Segoe UI")
    font.setPointSizeF(9.5)
    doc.setDefaultFont(font)
    option = QTextOption()
    option.setTextDirection(layout_direction_for(lang))
    doc.setDefaultTextOption(option)

    chart_size = None
    if payload["months"]:
        with suppress(Exception):
            chart_w, chart_h = int(body_w * 0.96), int(body_w * 0.26)
            data = chart_months_from_payload(payload)
            doc.addResource(QTextDocument.ResourceType.ImageResource, QUrl("chart.png"),
                            render_chart_image(data, chart_w, chart_h))
            chart_size = (chart_w, chart_h)
    doc.setHtml(build_report_html(payload, lang, chart_size))
    if not custom_pages:
        doc.print_(writer)
        return max(1, doc.pageCount())
    doc.setPageSize(QSizeF(body_w, body_h))
    pages = max(1, doc.pageCount())

    painter = QPainter(writer)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    left_flag = int((Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter).value)
    right_flag = int((Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter).value)
    small = QFont("Segoe UI")
    small.setPointSizeF(8)
    for page in range(pages):
        if page:
            writer.newPage()
        painter.save()
        painter.translate(0, -page * body_h)
        doc.drawContents(painter, QRectF(0, page * body_h, body_w, body_h))
        painter.restore()
        rule = QPen(QColor("#d5dde8"))
        rule.setWidthF(1.0)
        painter.setPen(rule)
        painter.drawLine(QPointF(0, body_h + 2), QPointF(body_w, body_h + 2))
        painter.setFont(small)
        painter.setPen(QColor("#64748b"))
        footer = QRectF(0, body_h + 4, body_w, footer_h - 4)
        first = f"{APP_NAME}  •  {payload['project']}"
        second = tr("report_page", lang, n=page + 1, m=pages)
        painter.drawText(footer, right_flag if rtl else left_flag, first)
        painter.drawText(footer, left_flag if rtl else right_flag, second)
    painter.end()
    return pages


def chart_months_from_payload(payload, span=12):
    """Last `span` months of the report's own data (so the chart matches the report period)."""
    months = payload["months"]
    if not months:
        return []
    last_key = months[-1][0]
    last_idx = int(last_key[:4]) * 12 + int(last_key[5:7]) - 1
    lookup = {k: (v, c) for k, v, c in months}
    out = []
    for i in range(last_idx - span + 1, last_idx + 1):
        year, month0 = divmod(i, 12)
        key = f"{year:04d}-{month0 + 1:02d}"
        total, count = lookup.get(key, (0.0, 0))
        out.append((key, total, count))
    # drop leading empty months so short reports are not mostly blank
    while len(out) > 1 and out[0][1] == 0 and out[0][2] == 0:
        out.pop(0)
    return out


def _com_init():
    if pythoncom is not None:
        pythoncom.CoInitialize()


class ExcelWorker(QObject):
    """Runs every Excel (COM) call on ONE background thread so the window never freezes.

    COM objects must stay on the thread that created them, so a single-thread pool
    is used for the whole life of the app.
    """
    finished = Signal(int, object, object)

    def __init__(self):
        super().__init__()
        self._pool = ThreadPoolExecutor(max_workers=1, initializer=_com_init)
        self._callbacks = {}
        self._seq = 0
        self.finished.connect(self._dispatch)

    def submit(self, fn, on_ok=None, on_err=None):
        self._seq += 1
        token = self._seq
        self._callbacks[token] = (on_ok, on_err)
        future = self._pool.submit(fn)

        def _done(f, token=token):
            try:
                self.finished.emit(token, f.result(), None)
            except Exception as exc:  # noqa: BLE001
                self.finished.emit(token, None, exc)

        future.add_done_callback(_done)

    def _dispatch(self, token, result, error):
        on_ok, on_err = self._callbacks.pop(token, (None, None))
        if error is not None:
            if on_err:
                on_err(error)
        elif on_ok:
            on_ok(result)

    def shutdown(self, last_job=None):
        with suppress(Exception):
            if last_job:
                self._pool.submit(last_job).result(timeout=15)
        self._pool.shutdown(wait=False)


class ExpenseManager:
    """Excel access layer. Every method must run on the ExcelWorker thread."""

    def __init__(self):
        self.excel: Any | None = None
        self._default_security: int | None = None

    def start_excel(self):
        if win32 is None:
            raise RuntimeError("pywin32 is not installed.")
        excel = self.excel
        if excel is None:
            excel = win32.DispatchEx("Excel.Application")
            self.excel = excel
            excel.Visible = False
            excel.DisplayAlerts = False
            excel.EnableEvents = True
            excel.ScreenUpdating = False
            with suppress(Exception):
                excel.AskToUpdateLinks = False
            with suppress(Exception):
                self._default_security = int(excel.AutomationSecurity)

    def stop_excel(self):
        excel = self.excel
        if excel:
            with suppress(Exception):
                excel.DisplayAlerts = True
                excel.ScreenUpdating = True
                excel.EnableEvents = True
                excel.Quit()
            self.excel = None

    def _open(self, path, write):
        """Open a workbook.

        Reading: read-only, events off and macros disabled - nothing has to run just
        to list the rows, which is much faster and never fights over the file lock.
        Writing: events/macros as before, because the workbook's VBA generates the ID.
        """
        self.start_excel()
        excel = self.excel
        if excel is None:
            raise RuntimeError("Excel is not available.")
        with suppress(Exception):
            if write:
                if self._default_security is not None:
                    excel.AutomationSecurity = self._default_security
                excel.EnableEvents = True
            else:
                excel.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
                excel.EnableEvents = False
        wb = excel.Workbooks.Open(
            os.path.abspath(path), UpdateLinks=0, ReadOnly=(not write),
            AddToMru=False, IgnoreReadOnlyRecommended=True,
        )
        if write and bool(wb.ReadOnly):
            self._close(wb)
            raise RuntimeError("The workbook is already open or locked in read-only mode. Close it in Excel and try again.")
        return wb

    def _close(self, wb):
        if wb:
            with suppress(Exception):
                wb.Close(SaveChanges=False)
        excel = self.excel
        if excel is not None:
            with suppress(Exception):
                excel.EnableEvents = True

    def _validate(self, table):
        headers = get_headers(table)
        missing = [h.upper() for h in REQUIRED_HEADERS if h not in headers]
        if missing:
            raise RuntimeError("Missing required columns: " + ", ".join(missing))
        return headers

    @staticmethod
    def _meta(table):
        return {"sheet": str(table.Parent.Name), "table": str(table.Name), "rows": int(table.ListRows.Count)}

    @staticmethod
    def _read_rows(table, headers):
        """Read the whole table body with ONE COM call instead of 4 calls per row."""
        if int(table.ListRows.Count) == 0:
            return []
        data = table.DataBodyRange.Value
        if not isinstance(data, tuple):
            data = ((data,),)
        ci, cd, cs, ca = (headers["id"] - 1, headers["description"] - 1, headers["date"] - 1, headers["amount"] - 1)
        return [
            {"excel_index": i, "id": rec[ci], "description": rec[cd], "date": rec[cs], "amount": rec[ca]}
            for i, rec in enumerate(data, 1)
        ]

    def load_project(self, path):
        """Validate + read in a single open (the old code opened the file twice)."""
        wb = None
        try:
            wb = self._open(path, write=False)
            table = find_table(wb)
            headers = self._validate(table)
            return self._meta(table), self._read_rows(table, headers)
        finally:
            self._close(wb)

    def _backup(self, path, kind="auto"):
        return make_backup(path, kind)

    def _wait_for_id(self, row, col):
        """The workbook's VBA fills the ID; poll briefly instead of a fixed sleep."""
        excel = self.excel
        if excel is None:
            raise RuntimeError("Excel is not available.")
        deadline = time.time() + 1.0
        while True:
            excel.Calculate()
            value = row.Cells(1, col).Value
            if value not in (None, "") or time.time() >= deadline:
                return value
            time.sleep(0.05)

    def _result(self, table, headers, generated_id, backup):
        return {"id": generated_id, "backup": backup,
                "meta": self._meta(table), "rows": self._read_rows(table, headers)}

    def _open_for_write(self, path):
        """Back up first, then open for writing. Returns (workbook, table, headers, backup)."""
        backup = self._backup(path)
        wb = None
        try:
            wb = self._open(path, write=True)
            table = find_table(wb)
            headers = self._validate(table)
            return wb, table, headers, backup
        except Exception:
            self._close(wb)
            raise

    @staticmethod
    def _check_row(table, headers, excel_index, expected):
        """Refuse to touch a row that moved or changed since it was loaded."""
        if excel_index < 1 or excel_index > table.ListRows.Count:
            raise StaleRowError("The selected row no longer exists.")
        if expected is None:
            return
        rng = table.ListRows(excel_index).Range
        actual = {
            "description": rng.Cells(1, headers["description"]).Value,
            "date": rng.Cells(1, headers["date"]).Value,
            "amount": rng.Cells(1, headers["amount"]).Value,
        }
        if not rows_match(expected, actual):
            raise StaleRowError("The expense was changed outside the application.")

    @staticmethod
    def _write_date(cell, date_text):
        """Write the date as before (text Excel parses itself) but verify Excel read it correctly.

        With some regional settings Excel can swap day and month; in that case the real
        date serial is written instead, so a wrong date can never be saved silently.
        """
        cell.Value = date_text
        intended = parse_excel_date(date_text)
        if intended is None:
            return
        with suppress(Exception):
            stored = cell.Value
            if isinstance(stored, str):
                return  # kept as text, exactly like earlier versions
            got = parse_excel_date(stored)
            if got is not None and got != intended:
                cell.Value = float((intended - date(1899, 12, 30)).days)
                if str(cell.NumberFormat) == "General":
                    cell.NumberFormat = "dd.mm.yyyy"

    def add_expense(self, path, description, date_text, amount):
        wb, table, headers, backup = self._open_for_write(path)
        try:
            table.ListRows.Add()
            row = table.ListRows(table.ListRows.Count).Range
            row.Cells(1, headers["description"]).Value = description
            self._write_date(row.Cells(1, headers["date"]), date_text)
            row.Cells(1, headers["amount"]).Value = amount
            generated_id = self._wait_for_id(row, headers["id"])
            wb.Save()
            return self._result(table, headers, generated_id, backup)
        finally:
            self._close(wb)

    def edit_expense(self, path, excel_index, description, date_text, amount, expected=None):
        wb, table, headers, backup = self._open_for_write(path)
        try:
            self._check_row(table, headers, excel_index, expected)
            row = table.ListRows(excel_index).Range
            row.Cells(1, headers["description"]).Value = description
            self._write_date(row.Cells(1, headers["date"]), date_text)
            row.Cells(1, headers["amount"]).Value = amount
            generated_id = self._wait_for_id(row, headers["id"])
            wb.Save()
            return self._result(table, headers, generated_id, backup)
        finally:
            self._close(wb)

    def delete_expense(self, path, excel_index, expected=None):
        wb, table, headers, backup = self._open_for_write(path)
        try:
            self._check_row(table, headers, excel_index, expected)
            table.ListRows(excel_index).Delete()
            wb.Save()
            return self._result(table, headers, None, backup)
        finally:
            self._close(wb)

    def export_xlsx(self, target, data):
        """Write a formatted report workbook. Always a NEW file: project files are never touched."""
        self.start_excel()
        excel = self.excel
        if excel is None:
            raise RuntimeError("Excel is not available.")
        lb = data["labels"]
        wb = None
        try:
            wb = excel.Workbooks.Add()
            ws = wb.Worksheets(1)
            ws.Name = lb["sheet_expenses"]
            ws2 = wb.Worksheets.Add(After=wb.Worksheets(wb.Worksheets.Count))
            ws2.Name = lb["sheet_monthly"]
            for i in range(wb.Worksheets.Count, 0, -1):  # drop any extra default sheets
                if str(wb.Worksheets(i).Name) not in (lb["sheet_expenses"], lb["sheet_monthly"]):
                    wb.Worksheets(i).Delete()
            for sheet in (ws, ws2):
                with suppress(Exception):
                    sheet.DisplayRightToLeft = bool(data["rtl"])
                with suppress(Exception):
                    sheet.Cells.Font.Name = "Segoe UI"
                    sheet.Cells.Font.Size = 10

            # ---- expenses sheet: title block + a real Excel table (filters, totals)
            ws.Cells(1, 1).Value = lb["heading"]
            ws.Cells(1, 1).Font.Size = 16
            ws.Cells(1, 1).Font.Bold = True
            with suppress(Exception):
                ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Interior.Color = 0x4F3216  # BGR of #16324f
                ws.Range(ws.Cells(1, 1), ws.Cells(1, 4)).Font.Color = 0xFFFFFF
                ws.Rows(1).RowHeight = 28
                ws.Cells(1, 1).VerticalAlignment = -4108
            for offset, (label, value) in enumerate(data["meta"], start=2):
                ws.Cells(offset, 1).Value = label
                ws.Cells(offset, 1).Font.Bold = True
                ws.Cells(offset, 1).HorizontalAlignment = -4131  # left
                ws.Cells(offset, 2).Value = value
                ws.Cells(offset, 2).HorizontalAlignment = -4131
            head_row = len(data["meta"]) + 3
            rows = data["rows"]
            last_row = head_row + len(rows)
            block = (tuple(data["headers"]),) + tuple(tuple(r) for r in rows)
            rng = ws.Range(ws.Cells(head_row, 1), ws.Cells(last_row, 4))
            ws.Range(ws.Cells(head_row + 1, 3), ws.Cells(last_row, 3)).NumberFormat = "dd.mm.yyyy"
            ws.Range(ws.Cells(head_row + 1, 4), ws.Cells(last_row, 4)).NumberFormat = "#,##0.00"
            rng.Value = block
            amount_total = f"=SUBTOTAL(109,D{head_row + 1}:D{last_row})"
            try:
                table = ws.ListObjects.Add(1, rng, XlListObjectHasHeaders=1)
                table.Name = "ExpensesReport"
                table.TableStyle = "TableStyleMedium2"
                table.ShowTotals = True
                table.TotalsRowRange.Cells(1, 2).Value = lb["total"]
                table.ListColumns(4).TotalsCalculation = 1  # xlTotalsCalculationSum
                table.TotalsRowRange.Cells(1, 4).NumberFormat = "#,##0.00"
            except Exception:  # noqa: BLE001 - plain styled range instead of a table
                rng.Rows(1).Font.Bold = True
                rng.Rows(1).Interior.Color = 0xF5ECE6
                ws.Cells(last_row + 1, 2).Value = lb["total"]
                ws.Cells(last_row + 1, 2).Font.Bold = True
                ws.Cells(last_row + 1, 4).Formula = amount_total
                ws.Cells(last_row + 1, 4).NumberFormat = "#,##0.00"
                ws.Cells(last_row + 1, 4).Font.Bold = True
            ws.Range(ws.Cells(head_row, 3), ws.Cells(last_row, 3)).HorizontalAlignment = -4108  # centre
            ws.Range(ws.Cells(head_row, 1), ws.Cells(last_row, 1)).HorizontalAlignment = -4108
            with suppress(Exception):
                ws.Range(ws.Cells(head_row + 1, 2), ws.Cells(last_row, 2)).HorizontalAlignment = -4131  # left (LTR)
                ws.Range(ws.Cells(head_row, 4), ws.Cells(last_row + 1, 4)).HorizontalAlignment = -4152  # right
                ws.Cells(head_row, 2).HorizontalAlignment = -4131
                ws.Range(ws.Cells(head_row, 1), ws.Cells(head_row, 4)).RowHeight = 22
                ws.Range(ws.Cells(head_row, 1), ws.Cells(head_row, 4)).VerticalAlignment = -4108
            rng.Columns.AutoFit()
            ws.Columns(1).ColumnWidth = max(10, min(float(ws.Columns(1).ColumnWidth), 18))
            ws.Columns(2).ColumnWidth = max(28, min(float(ws.Columns(2).ColumnWidth), 70))
            ws.Columns(3).ColumnWidth = max(13, min(float(ws.Columns(3).ColumnWidth), 18))
            ws.Columns(4).ColumnWidth = max(18, min(float(ws.Columns(4).ColumnWidth), 26))
            with suppress(Exception):  # frozen header, print-ready page setup
                ws.Activate()
                excel.ActiveWindow.FreezePanes = False
                excel.ActiveWindow.ScrollRow = 1
                ws.Cells(head_row + 1, 1).Select()
                excel.ActiveWindow.FreezePanes = True
                excel.ActiveWindow.DisplayGridlines = False
            with suppress(Exception):
                setup = ws.PageSetup
                setup.PaperSize = 9  # A4
                setup.Orientation = 1  # portrait
                setup.Zoom = False
                setup.FitToPagesWide = 1
                setup.FitToPagesTall = False
                setup.PrintTitleRows = f"${head_row}:${head_row}"
                setup.CenterFooter = "&P / &N"
                setup.LeftFooter = lb["heading"]
                setup.LeftMargin = setup.RightMargin = excel.CentimetersToPoints(1.5)

            # ---- monthly summary sheet
            months = data["months"]
            ws2.Cells(1, 1).Value = lb["col_month"]
            ws2.Cells(1, 2).Value = lb["col_count"]
            ws2.Cells(1, 3).Value = lb["col_total"]
            if months:
                ws2.Range(ws2.Cells(2, 1), ws2.Cells(len(months) + 1, 1)).NumberFormat = "@"
                ws2.Range(ws2.Cells(2, 3), ws2.Cells(len(months) + 1, 3)).NumberFormat = "#,##0.00"
                ws2.Range(ws2.Cells(2, 1), ws2.Cells(len(months) + 1, 3)).Value = tuple(tuple(m) for m in months)
            mrng = ws2.Range(ws2.Cells(1, 1), ws2.Cells(len(months) + 1, 3))
            try:
                mt = ws2.ListObjects.Add(1, mrng, XlListObjectHasHeaders=1)
                mt.Name = "MonthlyReport"
                mt.TableStyle = "TableStyleMedium2"
                mt.ShowTotals = True
                mt.TotalsRowRange.Cells(1, 1).Value = lb["total"]
                mt.ListColumns(2).TotalsCalculation = 1
                mt.ListColumns(3).TotalsCalculation = 1
                mt.TotalsRowRange.Cells(1, 3).NumberFormat = "#,##0.00"
            except Exception:  # noqa: BLE001
                mrng.Rows(1).Font.Bold = True
            mrng.Columns.AutoFit()
            with suppress(Exception):
                ws2.Range(ws2.Cells(1, 1), ws2.Cells(len(months) + 2, 3)).HorizontalAlignment = -4108
                ws2.Range(ws2.Cells(1, 3), ws2.Cells(len(months) + 2, 3)).HorizontalAlignment = -4152
                ws2.Columns(1).ColumnWidth = 14
                ws2.Columns(2).ColumnWidth = 10
                ws2.Columns(3).ColumnWidth = 20
                ws2.Activate()
                excel.ActiveWindow.DisplayGridlines = False
            if len(months) > 1:
                with suppress(Exception):  # the chart is a bonus; never fail the export for it
                    anchor = ws2.Range("E2")
                    chart_obj = ws2.ChartObjects().Add(anchor.Left, anchor.Top, 460, 260)
                    chart = chart_obj.Chart
                    chart.ChartType = 51  # xlColumnClustered
                    series = chart.SeriesCollection().NewSeries()
                    series.Values = ws2.Range(ws2.Cells(2, 3), ws2.Cells(len(months) + 1, 3))
                    series.XValues = ws2.Range(ws2.Cells(2, 1), ws2.Cells(len(months) + 1, 1))
                    series.Name = lb["col_total"]
                    chart.HasTitle = True
                    chart.ChartTitle.Text = lb["chart_title"]
                    chart.HasLegend = False
                    with suppress(Exception):
                        series.HasDataLabels = True
                        series.DataLabels().NumberFormat = '[>=1000000]0.0,,"M";[>=1000]0.0,"k";0'
                        chart.ChartGroups(1).GapWidth = 60
                        chart.Axes(2).TickLabels.NumberFormat = '#,##0'
                        chart.Axes(2).MajorGridlines.Format.Line.ForeColor.RGB = 0xE3DCD5
                        series.Format.Fill.ForeColor.RGB = 0xB8793A  # BGR of #3a79b8

            with suppress(Exception):
                ws.Activate()
                ws.Range("A1").Select()
            wb.SaveAs(os.path.abspath(target), 51)  # xlOpenXMLWorkbook (.xlsx)
        finally:
            self._close(wb)
        return target

    def clear_table_rows(self, path):
        wb = None
        try:
            wb = self._open(path, write=True)
            table = find_table(wb)
            self._validate(table)
            while table.ListRows.Count:
                table.ListRows(1).Delete()
            wb.Save()
        finally:
            self._close(wb)

    def create_from_template(self, template, target):
        if os.path.exists(target):
            raise FileExistsError(target)
        shutil.copy2(template, target)
        try:
            self.clear_table_rows(target)
        except Exception:
            with suppress(Exception):
                os.remove(target)
            raise



class ProjectItem(QListWidgetItem):
    path: str

    def __init__(self, path: str):
        super().__init__(Path(path).stem)
        self.path = path
        self.setToolTip(path)
        self.setTextAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)


class NewProjectDialog(QDialog):
    def __init__(self, folder, manager, lang, parent=None):
        super().__init__(parent)
        self.folder = folder
        self.manager = manager
        self.lang = lang
        self.setWindowTitle(tr("create_title", lang))
        self.resize(620, 260)
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        help_label = QLabel(tr("template_help", lang))
        help_label.setWordWrap(True)
        help_label.setObjectName("muted")
        layout.addWidget(help_label)

        grid = QGridLayout()
        grid.addWidget(QLabel(tr("project_name", lang)), 0, 0)
        self.name = QLineEdit()
        self.name.setPlaceholderText("PRG-005")
        grid.addWidget(self.name, 0, 1)
        grid.addWidget(QLabel(tr("template", lang)), 1, 0)
        self.template = QLineEdit()
        self.template.setReadOnly(True)
        grid.addWidget(self.template, 1, 1)
        choose = QPushButton(tr("select_template", lang))
        choose.clicked.connect(self.choose_template)
        grid.addWidget(choose, 1, 2)
        layout.addLayout(grid)
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def choose_template(self):
        path, _ = QFileDialog.getOpenFileName(self, tr("select_template", self.lang), self.folder, "Excel Macro Workbook (*.xlsm)")
        if path:
            self.template.setText(path)

    def values(self):
        name = self.name.text().strip()
        template = self.template.text().strip()
        if not name:
            raise ValueError(tr("name_required", self.lang))
        if not template:
            raise ValueError(tr("template_required", self.lang))
        if not name.lower().endswith(".xlsm"):
            name += ".xlsm"
        return name, template



_LIGHT_CHART = {"grid": "#d5dde8", "bar": "#7aa7d9", "bar_sel": "#1f9d72", "bar_hover": "#4f8fd1",
                "text": "#0f172a", "muted": "#64748b"}

THEMES = {
    "dark": {
        "bg": "#0b1220", "panel": "#111b2e", "input": "#0b1424", "alt": "#0e182b", "border": "#223452",
        "border2": "#2a3d5c", "hover": "#16263f", "text": "#e6edf6", "strong": "#ffffff", "muted": "#8aa0bd",
        "faint": "#6f86a6", "label": "#9fb3cf", "head": "#16233a", "headtext": "#c9d6e8", "accent": "#2e8b68",
        "accent_hover": "#36a277", "accent_border": "#43ad83", "accent_text": "#34d399", "sel": "#1d4a3f",
        "sel_text": "#ffffff", "btn": "#1a2a44", "btn_hover": "#223656", "btn_border": "#2f4468",
        "btn_press": "#16253c", "danger_text": "#fca5a5", "danger_bg": "#1d1824", "danger_border": "#5a2f3a",
        "danger_hover": "#2b1c26", "ok": "#6ee7a8", "err": "#fca5a5", "header1": "#15283f", "header2": "#101b2f",
        "dis_text": "#566a89", "dis_bg": "#111b2d", "dis_border": "#1f3150", "scroll": "#2c4166",
        "scroll_hover": "#3b5580", "tip_bg": "#16233a", "c_green": "#34d399", "c_blue": "#60a5fa",
        "c_amber": "#fbbf24", "c_violet": "#a78bfa", "selection": "#2e8b68",
        "grid": "#2a3d5c", "bar": "#4a7fb8", "bar_sel": "#34d399", "bar_hover": "#6aa3dc",
    },
    "light": {
        "bg": "#f1f4f9", "panel": "#ffffff", "input": "#ffffff", "alt": "#f6f8fc", "border": "#dbe2ec",
        "border2": "#c7d1e0", "hover": "#eaf0f8", "text": "#1b2638", "strong": "#0f172a", "muted": "#5d6b82",
        "faint": "#7a889e", "label": "#4a5870", "head": "#e8eef6", "headtext": "#33435c", "accent": "#1f9d72",
        "accent_hover": "#188a63", "accent_border": "#188a63", "accent_text": "#12805c", "sel": "#d7efe6",
        "sel_text": "#0f172a", "btn": "#ffffff", "btn_hover": "#eef3fa", "btn_border": "#c7d1e0",
        "btn_press": "#e1e9f4", "danger_text": "#b42318", "danger_bg": "#fff5f4", "danger_border": "#f1b8b2",
        "danger_hover": "#ffe9e6", "ok": "#12805c", "err": "#b42318", "header1": "#ffffff", "header2": "#eef3fa",
        "dis_text": "#9aa6b8", "dis_bg": "#f1f4f9", "dis_border": "#dbe2ec", "scroll": "#c3cddc",
        "scroll_hover": "#a9b6ca", "tip_bg": "#ffffff", "c_green": "#1f9d72", "c_blue": "#3b82f6",
        "c_amber": "#f59e0b", "c_violet": "#8b5cf6", "selection": "#9bd7bf",
        "grid": "#d5dde8", "bar": "#7aa7d9", "bar_sel": "#1f9d72", "bar_hover": "#4f8fd1",
    },
    "print": _LIGHT_CHART,
}

STYLE_TEMPLATE = Template("""
* { font-family:'Segoe UI','Tahoma'; font-size:13px; }
QWidget { background:${bg}; color:${text}; }
QLabel { background:transparent; }
QToolTip { background:${tip_bg}; color:${text}; border:1px solid ${border2}; padding:4px 6px; }

QFrame#header { background:qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 ${header1}, stop:1 ${header2});
                border:1px solid ${border}; border-radius:16px; }
QLabel#appTitle { font-size:22px; font-weight:800; color:${strong}; }
QLabel#badge { background:${accent}; color:#ffffff; border-radius:8px; padding:2px 9px; font-size:11px; font-weight:800; }
QLabel#version { color:${faint}; font-size:12px; }

QFrame#panel { background:${panel}; border:1px solid ${border}; border-radius:16px; }
QFrame#card { background:${panel}; border:1px solid ${border}; border-radius:14px; min-width:140px; }
QFrame#card[accent="green"]  { border-left:4px solid ${c_green}; }
QFrame#card[accent="blue"]   { border-left:4px solid ${c_blue}; }
QFrame#card[accent="amber"]  { border-left:4px solid ${c_amber}; }
QFrame#card[accent="violet"] { border-left:4px solid ${c_violet}; }
QLabel#cardLabel { color:${muted}; font-size:12px; font-weight:600; }
QLabel#cardValue { color:${strong}; font-size:21px; font-weight:800; }

QLabel#heading { font-size:20px; font-weight:800; color:${strong}; }
QLabel#sectionTitle { font-size:16px; font-weight:700; color:${strong}; }
QLabel#fieldLabel { color:${label}; font-size:12px; font-weight:600; }
QLabel#selectedProject { color:${accent_text}; font-weight:700; font-size:14px; }
QLabel#muted { color:${muted}; }
QLabel#empty { color:${faint}; font-size:14px; }
QLabel#hint { color:${faint}; font-size:11px; padding:0 6px; }
QLabel#status { color:${label}; padding:2px 8px; }
QLabel#status[kind="ok"]  { color:${ok}; }
QLabel#status[kind="err"] { color:${err}; }
QLabel#totals { color:${label}; font-weight:700; padding:4px 6px; }

QLineEdit, QDateEdit, QComboBox { background:${input}; border:1px solid ${border2}; border-radius:10px;
                                  padding:8px 10px; selection-background-color:${selection}; min-height:20px; }
QLineEdit:hover, QDateEdit:hover, QComboBox:hover { border-color:${btn_border}; }
QLineEdit:focus, QDateEdit:focus, QComboBox:focus { border:1px solid ${accent_text}; }
QLineEdit:disabled, QDateEdit:disabled { color:${dis_text}; background:${dis_bg}; }
QComboBox::drop-down, QDateEdit::drop-down { border:0; width:24px; }
QComboBox QAbstractItemView { background:${panel}; color:${text}; border:1px solid ${border2};
                              selection-background-color:${sel}; selection-color:${sel_text}; outline:0; }
QCalendarWidget QWidget { background:${panel}; color:${text}; }
QCalendarWidget QAbstractItemView:enabled { selection-background-color:${accent}; selection-color:#fff; }

QListView { background:${panel}; color:${text}; border:1px solid ${border2}; selection-background-color:${sel};
            selection-color:${sel_text}; outline:0; }
QListWidget { background:${input}; border:1px solid ${border2}; border-radius:12px; padding:4px; outline:0; }
QListWidget::item { padding:10px 10px; border-radius:8px; margin:1px 0; }
QListWidget::item:hover { background:${hover}; }
QListWidget::item:selected { background:${sel}; color:${sel_text}; }

QTableWidget { background:${input}; alternate-background-color:${alt}; border:1px solid ${border2};
               border-radius:12px; outline:0; selection-background-color:${sel}; }
QTableWidget::item { padding:4px 10px; border:0; }
QTableWidget::item:hover { background:${hover}; }
QTableWidget::item:selected { background:${sel}; color:${sel_text}; }
QHeaderView::section { background:${head}; color:${headtext}; border:0; padding:10px; font-weight:700; }

QPushButton { background:${btn}; border:1px solid ${btn_border}; border-radius:10px; padding:8px 14px; font-weight:600; min-height:20px; }
QPushButton:hover { background:${btn_hover}; border-color:${accent_text}; }
QPushButton:pressed { background:${btn_press}; }
QPushButton:focus { border-color:${accent_text}; }
QPushButton::menu-indicator { image:none; width:0px; }
QPushButton#primary { background:${accent}; border:1px solid ${accent_border}; padding:10px 12px; font-size:14px;
                      font-weight:700; color:#ffffff; min-height:22px; }
QPushButton#primary:hover { background:${accent_hover}; }
QPushButton#danger { color:${danger_text}; border-color:${danger_border}; background:${danger_bg}; }
QPushButton#danger:hover { background:${danger_hover}; border-color:${danger_text}; }
QPushButton#ghost { background:transparent; border:1px solid ${border2}; padding:6px 12px; min-height:20px; }
QPushButton#ghost:hover { background:${hover}; }
QPushButton:disabled { color:${dis_text}; background:${dis_bg}; border-color:${dis_border}; }

QCheckBox { spacing:8px; color:${label}; background:transparent; }
QCheckBox::indicator { width:16px; height:16px; border-radius:5px; border:1px solid ${border2}; background:${input}; }
QCheckBox::indicator:hover { border-color:${accent_text}; }
QCheckBox::indicator:checked { background:${accent}; border-color:${accent_border}; }

QMenu { background:${panel}; border:1px solid ${border2}; border-radius:10px; padding:6px; }
QMenu::item { padding:8px 22px; border-radius:6px; }
QMenu::item:selected { background:${hover}; }
QMenu::item:disabled { color:${dis_text}; }
QMenu::separator { height:1px; background:${border}; margin:6px 8px; }

QProgressBar { background:${head}; border:0; border-radius:2px; max-height:4px; min-height:4px; }
QProgressBar::chunk { background:${accent_text}; border-radius:2px; }

QScrollArea#formScroll, QScrollArea#formScroll > QWidget > QWidget { background:transparent; border:0; }
QWidget#formBody { background:transparent; }
QSplitter::handle { background:transparent; }
QScrollBar:vertical { background:transparent; width:10px; margin:2px; }
QScrollBar::handle:vertical { background:${scroll}; border-radius:4px; min-height:30px; }
QScrollBar::handle:vertical:hover { background:${scroll_hover}; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height:0; }
QScrollBar:horizontal { background:transparent; height:10px; margin:2px; }
QScrollBar::handle:horizontal { background:${scroll}; border-radius:4px; min-width:30px; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal { width:0; }
""")


def build_style(theme="dark"):
    return STYLE_TEMPLATE.substitute(THEMES.get(theme, THEMES["dark"]))


STYLE = build_style("dark")  # kept for the password-reset dialog


class CardWidget(QFrame):
    def __init__(self, accent: str = "green"):
        super().__init__()
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 12, 16, 12)
        layout.setSpacing(2)
        self._label = QLabel()
        self._label.setObjectName("cardLabel")
        self._value = QLabel("0.00 DZD" if accent != "violet" else "0")
        self._value.setObjectName("cardValue")
        layout.addWidget(self._label)
        layout.addWidget(self._value)


class MonthlyChart(QWidget):
    """Small bar chart of the last 12 months. Click a bar to filter the table to that month."""

    monthClicked = Signal(str)

    def __init__(self):
        super().__init__()
        self.setMinimumHeight(110)
        self.setMouseTracking(True)
        self.setLayoutDirection(Qt.LayoutDirection.LeftToRight)  # a timeline always reads left to right
        self._data = []
        self._selected = "all"
        self._hover = -1
        self._hits = []
        self._colors = THEMES["dark"]
        self._empty_text = ""

    def set_colors(self, colors):
        self._colors = colors
        self.update()

    def set_data(self, data, empty_text=""):
        self._data = list(data)
        self._empty_text = empty_text
        self._hover = -1
        self.update()

    def set_selected(self, key):
        if key != self._selected:
            self._selected = key
            self.update()

    def _index_at(self, pos):
        for i, rect in enumerate(self._hits):
            if rect.contains(pos):
                return i
        return -1

    def paintEvent(self, _event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(self.rect())
        if not self._data:
            painter.setPen(QColor(self._colors["muted"]))
            painter.drawText(rect, int(Qt.AlignmentFlag.AlignCenter.value), self._empty_text)
        else:
            self._hits = paint_monthly_chart(painter, rect, self._data, self._colors, self._selected, self._hover)
        painter.end()

    def mouseMoveEvent(self, event):
        idx = self._index_at(event.position())
        if idx != self._hover:
            self._hover = idx
            self.update()
        if idx >= 0:
            key, total, count = self._data[idx]
            lang = self.property("lang") or "en"
            QToolTip.showText(event.globalPosition().toPoint(),
                              tr("chart_tip", lang, month=key, amount=format_amount(total), n=count), self)
        else:
            QToolTip.hideText()

    def leaveEvent(self, _event):
        if self._hover != -1:
            self._hover = -1
            self.update()

    def mousePressEvent(self, event):
        idx = self._index_at(event.position())
        if idx >= 0 and event.button() == Qt.MouseButton.LeftButton:
            self.monthClicked.emit(self._data[idx][0])


class ReportDialog(QDialog):
    """Pick a period (and optional text filter), then export to Excel or create a PDF."""

    def __init__(self, rows, project, lang, month="all", text="", parent=None):
        super().__init__(parent)
        self.rows = rows
        self.lang = lang
        self.choice = None
        self._all_dates = True
        t = lambda k, **kw: tr(k, lang, **kw)  # noqa: E731
        self.setWindowTitle(f"{t('report_title')} - {project}")
        self.setMinimumWidth(520)
        days = [r["_date"] for r in rows if r.get("_date")]
        today = datetime.now().astimezone().date()
        self.min_day = min(days) if days else today
        self.max_day = max(days) if days else today

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)
        title = QLabel(t("report_title"))
        title.setObjectName("heading")
        layout.addWidget(title)
        help_label = QLabel(t("report_help"))
        help_label.setObjectName("muted")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        layout.addWidget(self._field_label(t("period")))
        self.preset = QComboBox()
        for key in ("all", "this_month", "last_month", "this_year", "last_year", "custom"):
            self.preset.addItem(t(f"preset_{key}"), key)
        layout.addWidget(self.preset)

        dates = QHBoxLayout()
        dates.setSpacing(10)
        self.from_edit, self.to_edit = QDateEdit(), QDateEdit()
        for edit, label in ((self.from_edit, t("from")), (self.to_edit, t("to"))):
            edit.setCalendarPopup(True)
            edit.setLocale(latin_locale())
            edit.setDisplayFormat("dd.MM.yyyy")
            edit.setLayoutDirection(Qt.LayoutDirection.LeftToRight)
            box = QVBoxLayout()
            box.setSpacing(4)
            box.addWidget(self._field_label(label))
            box.addWidget(edit)
            dates.addLayout(box, 1)
        layout.addLayout(dates)

        layout.addWidget(self._field_label(t("contains")))
        self.text_edit = QLineEdit(text)
        self.text_edit.setClearButtonEnabled(True)
        layout.addWidget(self.text_edit)

        self.summary = QLabel()
        self.summary.setObjectName("totals")
        layout.addWidget(self.summary)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        self.excel_btn = QPushButton(t("btn_excel"))
        self.pdf_btn = QPushButton(t("btn_pdf"))
        self.pdf_btn.setObjectName("primary")
        cancel = QPushButton(t("cancel"))
        cancel.setObjectName("ghost")
        buttons.addStretch(1)
        for b in (cancel, self.excel_btn, self.pdf_btn):
            buttons.addWidget(b)
        layout.addLayout(buttons)
        cancel.clicked.connect(self.reject)
        self.excel_btn.clicked.connect(lambda: self._finish("xlsx"))
        self.pdf_btn.clicked.connect(lambda: self._finish("pdf"))

        if month != "all":
            first = date(int(month[:4]), int(month[5:7]), 1)
            self._set_range(first, last_day_of_month(first))
            self._set_preset("custom")
            self._all_dates = False
        else:
            self._set_range(self.min_day, self.max_day)
        self.preset.currentIndexChanged.connect(self._preset_changed)
        self.from_edit.dateChanged.connect(self._range_edited)
        self.to_edit.dateChanged.connect(self._range_edited)
        self.text_edit.textChanged.connect(self._refresh)
        self._refresh()

    @staticmethod
    def _field_label(text):
        label = QLabel(text)
        label.setObjectName("fieldLabel")
        return label

    def _set_range(self, first, last):
        for edit, day in ((self.from_edit, first), (self.to_edit, last)):
            edit.blockSignals(True)
            edit.setDate(QDate(day.year, day.month, day.day))
            edit.blockSignals(False)

    def _set_preset(self, key):
        self.preset.blockSignals(True)
        self.preset.setCurrentIndex(max(self.preset.findData(key), 0))
        self.preset.blockSignals(False)

    def _preset_changed(self):
        key = self.preset.currentData()
        if key == "custom":
            self._all_dates = False
        else:
            today = datetime.now().astimezone().date()
            self._all_dates = key == "all"
            self._set_range(*preset_range(key, today, self.min_day, self.max_day))
        self._refresh()

    def _range_edited(self):
        self._all_dates = False
        self._set_preset("custom")
        self._refresh()

    def selection(self):
        f, t_ = self.from_edit.date(), self.to_edit.date()
        first, last = date(f.year(), f.month(), f.day()), date(t_.year(), t_.month(), t_.day())
        if first > last:
            first, last = last, first
        return {"from": first, "to": last, "text": self.text_edit.text(), "all_dates": self._all_dates}

    def _refresh(self):
        sel = self.selection()
        chosen = report_rows(self.rows, sel["from"], sel["to"], sel["text"], sel["all_dates"])
        s = summarize_rows(chosen)
        if chosen:
            self.summary.setText(tr("preview_summary", self.lang, n=len(chosen), total=format_amount(s["total"])))
        else:
            self.summary.setText(tr("no_data_in_range", self.lang))
        self.excel_btn.setEnabled(bool(chosen))
        self.pdf_btn.setEnabled(bool(chosen))

    def _finish(self, kind):
        self.choice = kind
        self.accept()


class RestoreDialog(QDialog):
    """List the backups of one project; the chosen one is restored by the main window."""

    def __init__(self, project_path, lang, parent=None):
        super().__init__(parent)
        self.lang = lang
        self.project_path = project_path
        self.selected = None
        t = lambda k, **kw: tr(k, lang, **kw)  # noqa: E731
        self.setWindowTitle(f"{t('restore_title')} - {Path(project_path).stem}")
        self.resize(620, 420)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(10)
        title = QLabel(t("restore_title"))
        title.setObjectName("heading")
        layout.addWidget(title)
        help_label = QLabel(t("restore_help"))
        help_label.setObjectName("muted")
        help_label.setWordWrap(True)
        layout.addWidget(help_label)

        self.backups = list_backups(project_path)
        self.table = QTableWidget(len(self.backups), 3)
        self.table.setHorizontalHeaderLabels([t("col_when"), t("col_type"), t("col_size")])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(36)
        hh = self.table.horizontalHeader()
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        kinds = {"auto": t("backup_auto"), "manual": t("backup_manual"), "prerestore": t("backup_prerestore")}
        for i, b in enumerate(self.backups):
            items = (b["when"].strftime("%d.%m.%Y  %H:%M:%S"), kinds.get(b["kind"], b["kind"]), format_size(b["size"]))
            for c, text in enumerate(items):
                item = QTableWidgetItem(text)
                item.setToolTip(b["path"])
                if c == 0:
                    item.setTextAlignment(Qt.AlignmentFlag.AlignLeading | Qt.AlignmentFlag.AlignVCenter)
                self.table.setItem(i, c, item)
        self.empty = QLabel(t("no_backups"))
        self.empty.setObjectName("empty")
        self.empty.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.table, 1)
        layout.addWidget(self.empty)
        self.table.setVisible(bool(self.backups))
        self.empty.setVisible(not self.backups)

        buttons = QHBoxLayout()
        buttons.setSpacing(8)
        folder_btn = QPushButton(t("open_backup_folder"))
        folder_btn.setObjectName("ghost")
        folder_btn.clicked.connect(self._open_folder)
        close_btn = QPushButton(t("cancel"))
        close_btn.setObjectName("ghost")
        close_btn.clicked.connect(self.reject)
        self.restore_btn = QPushButton(t("restore_btn"))
        self.restore_btn.setObjectName("primary")
        self.restore_btn.clicked.connect(self._restore)
        buttons.addWidget(folder_btn)
        buttons.addStretch(1)
        buttons.addWidget(close_btn)
        buttons.addWidget(self.restore_btn)
        layout.addLayout(buttons)
        self.table.itemSelectionChanged.connect(self._update_buttons)
        self.table.itemDoubleClicked.connect(lambda _i: self._restore())
        if self.backups:
            self.table.selectRow(0)
        self._update_buttons()

    def _update_buttons(self):
        self.restore_btn.setEnabled(bool(self.table.selectedItems()))

    def _open_folder(self):
        folder = backup_dir_for(self.project_path)
        folder.mkdir(exist_ok=True)
        with suppress(Exception):
            os.startfile(str(folder))

    def _restore(self):
        rows = {i.row() for i in self.table.selectedItems()}
        if rows:
            self.selected = self.backups[min(rows)]
            self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.settings = QSettings(SETTINGS_ORG, SETTINGS_APP)
        self.manager = ExpenseManager()
        self.worker = ExcelWorker()
        self.files = []
        self.current_file = None
        self.file_ok = False
        self.rows = []
        self.current_row = None
        self._pending = 0
        self._load_id = 0
        self.cache = {}            # path -> (mtime, prepared rows); feeds the global search
        self.results = []
        self._focus_target = None  # (path, excel_index) to select after opening a search result
        self._indexing = False
        self._index_queue = []
        self._index_total = 0
        self._index_done = 0
        self._search_id = 0
        self.sort_column = None    # None = Excel order, newest first
        self.sort_asc = True
        self.lang = saved_language()
        self.theme = saved_theme()

        # Start Excel in the background right away: its start-up time is the biggest
        # part of the first delay, and it now overlaps with the window appearing.
        self.worker.submit(self.manager.start_excel, on_err=self._excel_start_error)

        self._load_timer = QTimer(self)
        self._load_timer.setSingleShot(True)
        self._load_timer.timeout.connect(self._start_load)
        self._search_timer = QTimer(self)
        self._search_timer.setSingleShot(True)
        self._search_timer.timeout.connect(self.run_search)

        self.setWindowTitle(f"{APP_NAME} Pro | {APP_VERSION}")
        self.resize(1280, 820)
        self.setMinimumSize(920, 620)
        self.build_ui()
        self.apply_theme()
        self.apply_layout_direction()
        self.retranslate_ui()
        self.install_actions()
        self.update_controls()
        self.load_saved_folder()

    def _excel_start_error(self, exc):
        # The callback is delivered on the GUI thread after construction.
        self.notify(f"Excel startup failed: {exc}", "err")

    # ------------------------------------------------------------------ UI
    def build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        outer = QVBoxLayout(root)
        outer.setContentsMargins(18, 16, 18, 10)
        outer.setSpacing(12)

        header = QFrame(); header.setObjectName("header")
        hl = QHBoxLayout(header); hl.setContentsMargins(20, 12, 20, 12); hl.setSpacing(12)
        icon_file = resource_path("App.ico")
        if icon_file.exists():
            self.setWindowIcon(QIcon(str(icon_file)))
            self.logo = QLabel()
            self.logo.setPixmap(QIcon(str(icon_file)).pixmap(44, 44))
            self.logo.setFixedSize(48, 48)
            self.logo.setScaledContents(False)
            self.logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
            hl.addWidget(self.logo)
        self.title = QLabel(); self.title.setObjectName("appTitle")
        hl.addWidget(self.title)
        self.badge = QLabel("PRO"); self.badge.setObjectName("badge")
        hl.addWidget(self.badge, 0, Qt.AlignmentFlag.AlignVCenter)
        self.version_label = QLabel(f"v{APP_VERSION}"); self.version_label.setObjectName("version")
        hl.addWidget(self.version_label)
        hl.addStretch(1)
        self.theme_btn = QPushButton(); self.theme_btn.setObjectName("ghost"); self.theme_btn.setFocusPolicy(Qt.FocusPolicy.NoFocus); self.theme_btn.clicked.connect(self.toggle_theme)
        hl.addWidget(self.theme_btn)
        self.lang_label = QLabel(); self.lang_label.setObjectName("muted")
        hl.addWidget(self.lang_label)
        self.lang_combo = QComboBox()
        self.lang_combo.addItem("العربية", "ar"); self.lang_combo.addItem("English", "en"); self.lang_combo.addItem("Français", "fr")
        self.lang_combo.currentIndexChanged.connect(self.language_changed)
        hl.addWidget(self.lang_combo)
        outer.addWidget(header)

        # --- dashboard band: four KPI cards (2x2) + monthly chart
        band = QHBoxLayout(); band.setSpacing(12)
        grid = QGridLayout(); grid.setSpacing(12)
        self.card_total = self.make_card("green")
        self.card_today = self.make_card("blue")
        self.card_month = self.make_card("amber")
        self.card_rows = self.make_card("violet")
        grid.addWidget(self.card_total, 0, 0); grid.addWidget(self.card_today, 0, 1)
        grid.addWidget(self.card_month, 1, 0); grid.addWidget(self.card_rows, 1, 1)
        band.addLayout(grid, 5)
        self.chart_panel = QFrame(); self.chart_panel.setObjectName("panel")
        cpl = QVBoxLayout(self.chart_panel); cpl.setContentsMargins(14, 10, 14, 8); cpl.setSpacing(2)
        self.chart_title = QLabel(); self.chart_title.setObjectName("sectionTitle"); cpl.addWidget(self.chart_title)
        self.chart = MonthlyChart(); self.chart.monthClicked.connect(self.on_chart_month); cpl.addWidget(self.chart, 1)
        band.addWidget(self.chart_panel, 6)
        outer.addLayout(band)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.setHandleWidth(12)

        # --- projects panel
        left = QFrame(); left.setObjectName("panel"); left.setMinimumWidth(210)
        ll = QVBoxLayout(left); ll.setContentsMargins(14, 14, 14, 14); ll.setSpacing(9)
        self.projects_title = QLabel(); self.projects_title.setObjectName("sectionTitle"); ll.addWidget(self.projects_title)
        self.folder_label = QLabel(); self.folder_label.setObjectName("muted"); self.folder_label.setWordWrap(True); ll.addWidget(self.folder_label)
        bar = QHBoxLayout()
        self.choose_btn = QPushButton(); self.choose_btn.clicked.connect(self.choose_folder); bar.addWidget(self.choose_btn)
        self.refresh_btn = QPushButton(); self.refresh_btn.clicked.connect(self.scan_folder); bar.addWidget(self.refresh_btn)
        ll.addLayout(bar)
        self.new_btn = QPushButton(); self.new_btn.clicked.connect(self.create_project); ll.addWidget(self.new_btn)
        self.search = QLineEdit(); self.search.setClearButtonEnabled(True); self.search.textChanged.connect(self.filter_projects); ll.addWidget(self.search)
        self.project_list = QListWidget(); self.project_list.itemSelectionChanged.connect(self.project_selected); ll.addWidget(self.project_list, 1)
        splitter.addWidget(left)

        # --- expenses panel
        center = QFrame(); center.setObjectName("panel")
        cl = QVBoxLayout(center); cl.setContentsMargins(14, 14, 14, 10); cl.setSpacing(9)
        self.dashboard_title = QLabel(); self.dashboard_title.setObjectName("sectionTitle"); cl.addWidget(self.dashboard_title)
        toolbar = QHBoxLayout(); toolbar.setSpacing(8)
        self.month_combo = QComboBox(); self.month_combo.currentIndexChanged.connect(self.on_month_changed); toolbar.addWidget(self.month_combo, 1)
        self.open_btn = QPushButton(); self.open_btn.clicked.connect(self.open_excel); toolbar.addWidget(self.open_btn)
        self.export_btn = QPushButton(); self.export_btn.clicked.connect(self.open_report_dialog); toolbar.addWidget(self.export_btn)
        self.backup_btn = QPushButton()
        self.backup_menu = QMenu(self.backup_btn)
        self.act_backup_now = self.backup_menu.addAction("", self.manual_backup)
        self.act_restore = self.backup_menu.addAction("", self.restore_from_backup)
        self.backup_menu.addSeparator()
        self.act_open_backups = self.backup_menu.addAction("", self.open_backup_folder)
        self.backup_btn.setMenu(self.backup_menu)
        toolbar.addWidget(self.backup_btn)
        cl.addLayout(toolbar)
        searchbar = QHBoxLayout(); searchbar.setSpacing(8)
        self.desc_search = QLineEdit(); self.desc_search.setClearButtonEnabled(True)
        searchbar.addWidget(self.desc_search, 1)
        self.min_amount = QLineEdit(); self.max_amount = QLineEdit()
        for box in (self.min_amount, self.max_amount):
            box.setClearButtonEnabled(True); box.setFixedWidth(104); box.textChanged.connect(self.on_search_changed)
            searchbar.addWidget(box)
        self.scope_combo = QComboBox()
        self.scope_combo.addItem("", "file"); self.scope_combo.addItem("", "all")
        searchbar.addWidget(self.scope_combo)
        cl.addLayout(searchbar)
        self.desc_search.textChanged.connect(self.on_search_changed)
        self.scope_combo.currentIndexChanged.connect(self.on_search_changed)

        self.table = QTableWidget(0, 4)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.setShowGrid(False)
        self.table.setWordWrap(False)
        self.table.verticalHeader().setVisible(False)
        self.table.verticalHeader().setDefaultSectionSize(40)
        self.table.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        hh = self.table.horizontalHeader()
        hh.setStretchLastSection(False)
        hh.setSectionsClickable(True)
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hh.sectionClicked.connect(self.on_header_clicked)
        self.table.itemSelectionChanged.connect(self.expense_selected)
        self.table.itemDoubleClicked.connect(lambda _item: (self.description.setFocus(), self.description.selectAll()))
        self.empty_label = QLabel(); self.empty_label.setObjectName("empty"); self.empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.stack = QStackedWidget()
        self.stack.addWidget(self.table)
        self.stack.addWidget(self.empty_label)
        self.results_table = self._make_results_table()
        self.stack.addWidget(self.results_table)
        cl.addWidget(self.stack, 1)
        self.totals_label = QLabel(); self.totals_label.setObjectName("totals"); self.totals_label.setMinimumWidth(1)
        cl.addWidget(self.totals_label)
        splitter.addWidget(center)

        # --- form panel
        right = QFrame(); right.setObjectName("panel"); right.setMinimumWidth(320)
        right_outer = QVBoxLayout(right); right_outer.setContentsMargins(0, 0, 0, 0); right_outer.setSpacing(0)
        form_scroll = QScrollArea(); form_scroll.setObjectName("formScroll"); form_scroll.setWidgetResizable(True)
        form_scroll.setFrameShape(QFrame.Shape.NoFrame)
        form_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        form_body = QWidget(); form_body.setObjectName("formBody")
        rl = QVBoxLayout(form_body); rl.setContentsMargins(18, 18, 18, 18); rl.setSpacing(8)
        form_scroll.setWidget(form_body)
        right_outer.addWidget(form_scroll)
        self.form_title = QLabel(); self.form_title.setObjectName("heading"); rl.addWidget(self.form_title)
        self.selected_label = QLabel(); self.selected_label.setObjectName("selectedProject"); self.selected_label.setWordWrap(True)
        self.selected_label.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter); rl.addWidget(self.selected_label)
        self.info = QLabel(); self.info.setObjectName("muted"); self.info.setWordWrap(True)
        self.info.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter); rl.addWidget(self.info)
        rl.addSpacing(6)
        self.desc_label = QLabel(); self.desc_label.setObjectName("fieldLabel"); rl.addWidget(self.desc_label)
        self.description = QLineEdit(); rl.addWidget(self.description)
        self.completer = QCompleter(self)
        self.completer.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
        self.completer.setFilterMode(Qt.MatchFlag.MatchContains)
        self.completer.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.completer.setMaxVisibleItems(8)
        self.description.setCompleter(self.completer)
        for box in (self.description, self.search, self.desc_search):
            box.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.date_label = QLabel(); self.date_label.setObjectName("fieldLabel"); rl.addWidget(self.date_label)
        self.date_edit = QDateEdit(); self.date_edit.setLocale(latin_locale()); self.date_edit.setCalendarPopup(True); self.date_edit.setDisplayFormat("dd.MM.yyyy"); self.date_edit.setDate(QDate.currentDate()); rl.addWidget(self.date_edit)
        self.amount_label = QLabel(); self.amount_label.setObjectName("fieldLabel"); rl.addWidget(self.amount_label)
        self.amount = QLineEdit(); self.amount.setPlaceholderText("0.00"); self.amount.returnPressed.connect(self.save_current_expense); rl.addWidget(self.amount)
        self.description.returnPressed.connect(self.amount.setFocus)
        self.keep_date = QCheckBox()
        self.keep_date.setChecked(self.settings.value("keep_date", False, type=bool))
        self.keep_date.toggled.connect(lambda on: self.settings.setValue("keep_date", on))
        rl.addWidget(self.keep_date)
        rl.addSpacing(6)
        self.add_btn = QPushButton(); self.add_btn.setObjectName("primary"); self.add_btn.clicked.connect(self.save_current_expense); rl.addWidget(self.add_btn)
        editbar = QHBoxLayout()
        self.edit_btn = QPushButton(); self.edit_btn.clicked.connect(self.edit_expense); editbar.addWidget(self.edit_btn)
        self.delete_btn = QPushButton(); self.delete_btn.setObjectName("danger"); self.delete_btn.clicked.connect(self.delete_expense); editbar.addWidget(self.delete_btn)
        rl.addLayout(editbar)
        self.clear_btn = QPushButton(); self.clear_btn.setObjectName("ghost"); self.clear_btn.clicked.connect(self.clear_form); rl.addWidget(self.clear_btn)
        for btn in (self.add_btn, self.edit_btn, self.delete_btn, self.clear_btn):
            btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.add_btn.setFixedHeight(46)
        for btn in (self.edit_btn, self.delete_btn, self.clear_btn):
            btn.setFixedHeight(40)
        rl.addStretch(1)
        self.note = QLabel(); self.note.setObjectName("hint"); self.note.setWordWrap(True)
        self.note.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop); rl.addWidget(self.note)
        splitter.addWidget(right)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 3)
        splitter.setStretchFactor(2, 1)
        splitter.setSizes([250, 660, 330])
        outer.addWidget(splitter, 1)

        self.progress = QProgressBar(); self.progress.setRange(0, 0); self.progress.setTextVisible(False); self.progress.setVisible(False)
        outer.addWidget(self.progress)
        footer = QHBoxLayout(); footer.setContentsMargins(0, 0, 0, 0)
        self.status = QLabel(); self.status.setObjectName("status"); self.status.setMinimumWidth(1)
        self.status.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter)
        self.status.linkActivated.connect(self._status_link); footer.addWidget(self.status, 1)
        self.hint = QLabel(); self.hint.setObjectName("hint"); self.hint.setMinimumWidth(1); footer.addWidget(self.hint)
        outer.addLayout(footer)

    def _make_results_table(self):
        t = QTableWidget(0, 5)
        t.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        t.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        t.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        t.setAlternatingRowColors(True)
        t.setShowGrid(False)
        t.setWordWrap(False)
        t.verticalHeader().setVisible(False)
        t.verticalHeader().setDefaultSectionSize(38)
        t.verticalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Fixed)
        hh = t.horizontalHeader()
        hh.setStretchLastSection(False)
        hh.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        hh.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        hh.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        t.itemActivated.connect(self.open_result)
        return t

    def make_card(self, accent):
        frame = CardWidget(accent)
        frame.setObjectName("card")
        frame.setProperty("accent", accent)
        return frame

    def install_actions(self):
        refresh = QAction(self); refresh.setShortcut("F5"); refresh.triggered.connect(self.scan_folder); self.addAction(refresh)
        for keys, fn in (("Ctrl+Return", self.save_current_expense), ("Ctrl+Enter", self.save_current_expense),
                         ("Ctrl+S", self.save_current_expense), ("Esc", self.clear_form),
                         ("Ctrl+F", lambda: (self.desc_search.setFocus(), self.desc_search.selectAll())),
                         ("Ctrl+E", self.open_report_dialog),
                         ("Ctrl+N", lambda: (self.clear_form(), self.description.setFocus()))):
            sc = QShortcut(QKeySequence(keys), self)
            sc.activated.connect(fn)
        delete_sc = QShortcut(QKeySequence(Qt.Key.Key_Delete), self.table)
        delete_sc.setContext(Qt.ShortcutContext.WidgetShortcut)
        delete_sc.activated.connect(self.delete_expense)

    def apply_layout_direction(self):
        """Arabic mirrors the whole interface (right-to-left); French/English stay left-to-right.

        Numeric inputs and the chart stay left-to-right so digits and the timeline read naturally.
        """
        QApplication.setLayoutDirection(layout_direction_for(self.lang))
        for widget in (self.date_edit, self.amount, self.min_amount, self.max_amount, self.chart):
            widget.setLayoutDirection(Qt.LayoutDirection.LeftToRight)

    def retranslate_ui(self):
        t = lambda k, **kw: tr(k, self.lang, **kw)
        self.title.setText(t("app")); self.lang_label.setText(t("language"))
        self.theme_btn.setText(t("theme_light") if self.theme == "dark" else t("theme_dark"))
        self.projects_title.setText(t("projects")); self.choose_btn.setText(t("choose_folder")); self.refresh_btn.setText(t("refresh")); self.new_btn.setText(t("new_project")); self.search.setPlaceholderText(t("search"))
        self.dashboard_title.setText(t("expenses")); self.chart_title.setText(t("chart_title")); self.chart.setProperty("lang", self.lang)
        self.open_btn.setText(t("open_excel")); self.export_btn.setText(t("export"))
        self.backup_btn.setText(t("backup") + " \u25be")
        self.act_backup_now.setText(t("backup_now")); self.act_restore.setText(t("restore_backup")); self.act_open_backups.setText(t("open_backup_folder"))
        self.selected_label.setText(t("selected_project") if not self.current_file else Path(self.current_file).stem)
        self.desc_label.setText(t("description")); self.date_label.setText(t("date")); self.amount_label.setText(t("amount"))
        self.keep_date.setText(t("keep_date"))
        self.edit_btn.setText(t("edit_expense")); self.delete_btn.setText(t("delete_expense")); self.clear_btn.setText(t("cancel")); self.note.setText(t("note")); self.hint.setText(t("hint"))
        self.card_total._label.setText(t("total")); self.card_today._label.setText(t("today")); self.card_month._label.setText(t("month")); self.card_rows._label.setText(t("rows"))
        self.empty_label.setText(t("loading") if self._pending else t("no_rows"))
        self.desc_search.setPlaceholderText(t("search_desc"))
        self.min_amount.setPlaceholderText(t("min_amount")); self.max_amount.setPlaceholderText(t("max_amount"))
        self.scope_combo.setItemText(0, t("scope_file")); self.scope_combo.setItemText(1, t("scope_all"))
        self.results_table.setHorizontalHeaderLabels([t("col_project"), t("id"), t("description"), t("date"), t("amount")])
        self.rebuild_month_combo()
        self.refresh_table_headers()
        self.update_form_title()
        self.update_cards()
        idx = self.lang_combo.findData(self.lang)
        if idx >= 0 and self.lang_combo.currentIndex() != idx:
            self.lang_combo.blockSignals(True)
            self.lang_combo.setCurrentIndex(idx)
            self.lang_combo.blockSignals(False)

    def language_changed(self):
        code = self.lang_combo.currentData()
        if isinstance(code, str) and code != self.lang:
            self.lang = code
            self.settings.setValue("language", code)
            self.apply_layout_direction()
            self.retranslate_ui()
            self.notify(tr("language_changed", self.lang), "info")

    # ------------------------------------------------------------ state helpers
    def notify(self, text, kind="info"):
        self.status.setTextFormat(Qt.TextFormat.PlainText)
        self.status.setText(text)
        self.status.setToolTip("")
        self.status.setProperty("kind", kind)
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)

    def notify_file(self, key, path):
        """Status line for a saved report: short file name plus Open / Show in folder links."""
        self._status_path = path
        name = html.escape(os.path.basename(path))
        token = "\x01"
        label = html.escape(tr(key, self.lang, path=token)).replace(token, f"<b>\u2066{name}\u2069</b>")
        links = (f'&nbsp;&nbsp;<a href="open">{html.escape(tr("open_file", self.lang))}</a>'
                 f'&nbsp;&nbsp;|&nbsp;&nbsp;<a href="folder">{html.escape(tr("show_folder", self.lang))}</a>')
        self.status.setTextFormat(Qt.TextFormat.RichText)
        self.status.setText(label + links)
        self.status.setToolTip(path)
        self.status.setProperty("kind", "ok")
        self.status.style().unpolish(self.status)
        self.status.style().polish(self.status)

    def _status_link(self, which):
        path = getattr(self, "_status_path", None)
        if not path:
            return
        if which == "folder":
            with suppress(Exception):
                os.startfile(os.path.dirname(path))
        else:
            self.open_file(path)

    def begin_busy(self):
        self._pending += 1
        self.update_controls()

    def end_busy(self):
        self._pending = max(0, self._pending - 1)
        self.update_controls()

    def update_controls(self):
        busy = self._pending > 0
        ok = self.file_ok and not busy
        for w in (self.description, self.date_edit, self.amount, self.add_btn, self.open_btn, self.backup_btn, self.export_btn):
            w.setEnabled(ok)
        has_row = ok and self.current_row is not None
        self.edit_btn.setEnabled(has_row)
        self.delete_btn.setEnabled(has_row)
        self.new_btn.setEnabled(not busy)
        self.progress.setVisible(busy or self._indexing)

    def load_saved_folder(self):
        folder_value = self.settings.value("folder", "", type=str)
        folder = folder_value if isinstance(folder_value, str) else ""
        if folder and os.path.isdir(folder):
            self.set_folder(folder)
        else:
            self.choose_folder()

    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, tr("choose_folder", self.lang))
        if folder:
            self.set_folder(folder)

    def set_folder(self, folder):
        path = folder if isinstance(folder, str) else ""
        self.settings.setValue("folder", path)
        self.folder_label.setText(path)
        self.scan_folder()

    def scan_folder(self):
        folder = self.folder_label.text()
        if not folder or not os.path.isdir(folder):
            return
        # Prevent a load from the previously selected folder repainting this view.
        self._load_timer.stop()
        self._load_id += 1
        previous = self.current_file
        self.files = sorted(
            [str(p) for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower() == ".xlsm" and not p.name.startswith("~$")],
            key=lambda x: Path(x).stem.casefold(),
        )
        self.project_list.blockSignals(True)
        self.project_list.clear()
        for path in self.files:
            self.project_list.addItem(ProjectItem(path))
        self.project_list.blockSignals(False)
        self.filter_projects(self.search.text())
        self.notify(tr("files_found", self.lang, count=len(self.files)), "info")
        if self.files:
            target = 0
            if previous:
                for i, p in enumerate(self.files):
                    if p.lower() == previous.lower():
                        target = i
                        break
            self.current_file = None  # force a (re)load of the chosen item
            self.project_list.setCurrentRow(target)
        else:
            self.current_file = None
            self.file_ok = False
            self.rows = []
            self.current_row = None
            self.results = []
            self.table.setRowCount(0)
            self.results_table.setRowCount(0)
            self.empty_label.setText(tr("no_rows", self.lang))
            self.stack.setCurrentIndex(1)
            self.rebuild_month_combo()
            self.update_cards()
            self.update_controls()

    def filter_projects(self, text):
        needle = text.strip().casefold()
        for i in range(self.project_list.count()):
            self.project_list.item(i).setHidden(needle not in self.project_list.item(i).text().casefold())

    # ------------------------------------------------------------- loading
    def project_selected(self):
        items = self.project_list.selectedItems()
        if not items:
            return
        item = cast(ProjectItem, items[0])
        if item.path == self.current_file and self.file_ok:
            return
        if self.current_file and self.file_ok and self.form_dirty() and not self.confirm_discard():
            self._reselect_current()
            return
        self.current_file = item.path
        self.file_ok = False
        self.rows = []
        self.clear_form(False)
        self.selected_label.setText(item.text())
        self.info.setText(tr("loading", self.lang))
        self.empty_label.setText(tr("loading", self.lang))
        self.rebuild_month_combo()
        self.update_cards()
        self.update_controls()
        self._load_timer.start(120)  # skip files the user is just scrolling past

    def _start_load(self):
        path = self.current_file
        if not path:
            return
        self._load_id += 1
        lid = self._load_id

        def job():
            if lid != self._load_id:  # superseded by a newer selection
                return None
            meta, rows = self.manager.load_project(path)
            return meta, prepare_rows(rows)

        def ok(res):
            self.end_busy()
            if res is None or lid != self._load_id:
                return
            meta, rows = res
            self.file_ok = True
            self.rows = rows
            self.info.setText(tr("table_info", self.lang, **meta))
            self.empty_label.setText(tr("no_rows", self.lang))
            self.cache_put(path, rows)
            self.rebuild_month_combo()
            self.update_cards()
            self.update_controls()
            self._focus_row()

        def err(exc):
            self.end_busy()
            if lid != self._load_id:
                return
            self.file_ok = False
            self.info.setText(tr("file_check_failed", self.lang, error=exc))
            self.empty_label.setText(tr("no_rows", self.lang))
            self.update_controls()

        self.begin_busy()
        self.worker.submit(job, ok, err)

    def apply_result(self, res):
        """Use the rows returned by a write operation - no second open of the file."""
        self.rows = res["rows"]
        meta = res.get("meta")
        if meta:
            self.info.setText(tr("table_info", self.lang, **meta))
        self.rebuild_month_combo()
        self.update_cards()

    # ------------------------------------------------------------- table
    def refresh_table_headers(self):
        labels = [tr(key, self.lang) for key in ("id", "description", "date", "amount")]
        if self.sort_column in SORT_COLUMNS:
            labels[SORT_COLUMNS.index(self.sort_column)] += " \u25b2" if self.sort_asc else " \u25bc"
        self.table.setHorizontalHeaderLabels(labels)
        aligns = (Qt.AlignmentFlag.AlignCenter,
                  Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter,
                  Qt.AlignmentFlag.AlignCenter,
                  Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        for col, align in enumerate(aligns):
            head = self.table.horizontalHeaderItem(col)
            if head is not None:
                head.setTextAlignment(align)

    def rebuild_month_combo(self):
        current = self.month_combo.currentData() if self.month_combo.count() else "all"
        self.month_combo.blockSignals(True)
        self.month_combo.clear()
        self.month_combo.addItem(tr("all_months", self.lang), "all")
        months = sorted({r["_date"].strftime("%Y-%m") for r in self.rows if r.get("_date")}, reverse=True)
        for m in months:
            self.month_combo.addItem(m, m)
        idx = self.month_combo.findData(current)
        self.month_combo.setCurrentIndex(max(idx, 0))
        self.month_combo.blockSignals(False)
        self.apply_table_filter()

    def on_month_changed(self):
        self.clear_form()
        self.apply_table_filter()

    def apply_table_filter(self):
        if self.scope_combo.currentData() == "all":
            return  # the global results table is shown instead
        month = self.month_combo.currentData() if self.month_combo.count() else "all"
        month = month if isinstance(month, str) else "all"
        self.chart.set_selected(month)
        terms = search_terms(self.desc_search.text())
        low, high = parse_filter_amount(self.min_amount.text()), parse_filter_amount(self.max_amount.text())
        filtered = filter_rows(self.rows, month, terms, low, high)
        narrowed = bool(terms) or low is not None or high is not None
        visible_indexes = {r.get("excel_index") for r in filtered}
        if self.current_row is not None and self.current_row.get("excel_index") not in visible_indexes:
            self.clear_form()
        if self.sort_column in SORT_COLUMNS:
            rows = sort_rows(filtered, self.sort_column, self.sort_asc)
        else:
            rows = filtered[::-1] if NEWEST_FIRST else list(filtered)
        t = self.table
        t.setUpdatesEnabled(False)
        t.blockSignals(True)
        t.clearContents()
        t.setRowCount(len(rows))
        trailing = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        leading = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        center = Qt.AlignmentFlag.AlignCenter
        for i, r in enumerate(rows):
            id_item = QTableWidgetItem(format_id(r.get("id")))
            id_item.setData(Qt.ItemDataRole.UserRole, r["excel_index"])
            id_item.setTextAlignment(center)
            desc_item = QTableWidgetItem("" if r.get("description") is None else str(r.get("description")))
            desc_item.setTextAlignment(leading)
            date_item = QTableWidgetItem(excel_date_text(r.get("date")))
            date_item.setTextAlignment(center)
            amount_item = QTableWidgetItem(format_amount(r.get("amount")))
            amount_item.setTextAlignment(trailing)
            for c, it in enumerate((id_item, desc_item, date_item, amount_item)):
                t.setItem(i, c, it)
        if self.current_row is not None:  # keep the selected expense highlighted after sorting/filtering
            for i in range(t.rowCount()):
                cell = t.item(i, 0)
                if cell is not None and cell.data(Qt.ItemDataRole.UserRole) == self.current_row.get("excel_index"):
                    t.selectRow(i)
                    break
        t.blockSignals(False)
        t.setUpdatesEnabled(True)
        if not rows:
            if narrowed and self.rows:
                self.empty_label.setText(tr("no_results", self.lang))
            elif self.file_ok:
                self.empty_label.setText(tr("no_rows", self.lang))
        self.stack.setCurrentIndex(0 if rows else 1)
        if not self.rows:
            self.totals_label.setText("")
        else:
            total = format_amount(summarize_rows(rows)["total"])
            if narrowed or month != "all":
                self.totals_label.setText(tr("filtered_total", self.lang, n=len(rows), m=len(self.rows), total=total))
            else:
                self.totals_label.setText(tr("preview_summary", self.lang, n=len(rows), total=total))
        if terms and self.rows:
            self.notify(tr("filtered", self.lang, n=len(rows), m=len(self.rows)), "info")

    def expense_selected(self):
        items = self.table.selectedItems()
        if not items:
            return
        item = self.table.item(items[0].row(), 0)
        if item is None:
            return
        excel_index = item.data(Qt.ItemDataRole.UserRole)
        self.current_row = next((r for r in self.rows if r["excel_index"] == excel_index), None)
        if not self.current_row:
            return
        self.description.setText(str(self.current_row.get("description") or ""))
        d = self.current_row.get("_date")
        if d:
            self.date_edit.setDate(QDate(d.year, d.month, d.day))
        self.amount.setText(format_amount(self.current_row.get("amount")))
        self.update_form_title()
        self.update_controls()

    def update_form_title(self):
        if self.current_row:
            self.form_title.setText(tr("edit_expense", self.lang))
            self.add_btn.setText(tr("save_changes", self.lang))
        else:
            self.form_title.setText(tr("add_expense", self.lang))
            self.add_btn.setText(tr("add_expense", self.lang))

    def clear_form(self, keep_project=True):
        self.current_row = None
        self.description.clear()
        self.amount.clear()
        self.date_edit.setDate(QDate.currentDate())
        self.table.blockSignals(True)
        self.table.clearSelection()
        self.table.blockSignals(False)
        self.update_form_title()
        self.update_controls()

    # ------------------------------------------------------------- actions
    def validate_form(self):
        if not self.current_file:
            raise ValueError(tr("no_project", self.lang))
        desc = self.description.text().strip()
        if not desc:
            raise ValueError(tr("missing_description", self.lang))
        try:
            amount = parse_amount(self.amount.text())
        except ValueError:
            raise ValueError(tr("invalid_amount", self.lang))
        return desc, self.date_edit.date().toString("dd.MM.yyyy"), amount

    def add_expense(self):
        try:
            desc, date_text, amount = self.validate_form()
        except ValueError as exc:
            QMessageBox.warning(self, tr("error", self.lang), str(exc))
            return
        if find_duplicate(self.rows, desc, parse_excel_date(date_text), amount) is not None:
            answer = QMessageBox.question(
                self, tr("duplicate_title", self.lang), tr("duplicate_confirm", self.lang),
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
            if answer != QMessageBox.StandardButton.Yes:
                return
        path = self.current_file
        self.run_action(lambda: self.manager.add_expense(path, desc, date_text, amount), tr("added", self.lang), keep_date=True)

    def save_current_expense(self):
        """Add a new expense, or update the currently selected Table row."""
        if not (self.file_ok and self._pending == 0):
            return
        if self.current_row:
            self.edit_expense()
        else:
            self.add_expense()

    def edit_expense(self):
        if not self.current_row:
            return
        try:
            desc, date_text, amount = self.validate_form()
        except ValueError as exc:
            QMessageBox.warning(self, tr("error", self.lang), str(exc))
            return
        path, idx = self.current_file, self.current_row["excel_index"]
        expected = dict(self.current_row)  # the row is only changed if it still looks like this
        self.run_action(lambda: self.manager.edit_expense(path, idx, desc, date_text, amount, expected), tr("edited", self.lang))

    def delete_expense(self):
        if not self.current_row or not (self.file_ok and self._pending == 0):
            return
        row = self.current_row
        detail = tr("confirm_delete_detail", self.lang, description=str(row.get("description") or ""),
                    date=excel_date_text(row.get("date")), amount=format_amount(row.get("amount")))
        if QMessageBox.question(self, tr("delete_expense", self.lang), detail,
                                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                                QMessageBox.StandardButton.No) != QMessageBox.StandardButton.Yes:
            return
        path, idx, expected = self.current_file, row["excel_index"], dict(row)
        self.run_action(lambda: self.manager.delete_expense(path, idx, expected), tr("deleted", self.lang))

    def run_action(self, fn, success_text, keep_date=False):
        path = self.current_file
        self.notify(tr("saving", self.lang), "info")

        def ok(res):
            self.end_busy()
            generated = res.get("id")
            res["rows"] = prepare_rows(res["rows"])
            self.cache_put(path, res["rows"])
            msg = success_text + (f"  \u2022  ID: {format_id(generated)}" if generated not in (None, "") else "")
            self.notify(msg + "  \u2022  " + tr("backup_created", self.lang), "ok")
            if path == self.current_file:
                used_date = self.date_edit.date()
                self.clear_form()
                self.apply_result(res)
                if keep_date and self.keep_date.isChecked():
                    self.date_edit.setDate(used_date)
                self.description.setFocus()

        def err(exc):
            self.end_busy()
            if isinstance(exc, StaleRowError):
                message = tr("stale_row", self.lang)
                self.notify(message, "err")
                QMessageBox.warning(self, tr("error", self.lang), message)
                self.scan_folder()  # reload so the list matches the file again
                return
            self.notify(str(exc), "err")
            QMessageBox.critical(self, tr("error", self.lang), str(exc))

        self.begin_busy()
        self.worker.submit(fn, ok, err)

    def open_excel(self):
        if self.current_file:
            try:
                os.startfile(self.current_file)
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(self, tr("error", self.lang), str(exc))

    def manual_backup(self):
        if not self.current_file or self._pending:
            return
        path = self.current_file

        def ok(target):
            self.end_busy()
            self.notify(tr("backup_ok", self.lang, path=target), "ok")

        def err(exc):
            self.end_busy()
            QMessageBox.critical(self, tr("error", self.lang), str(exc))

        self.begin_busy()
        self.worker.submit(lambda: self.manager._backup(path, "manual"), ok, err)

    def update_cards(self):
        total = today = month_total = 0.0
        now = datetime.now().astimezone().date()
        count = 0
        for r in self.rows:
            value = r.get("_amount", 0.0)
            if r.get("_numeric"):
                total += value
                count += 1
            d = r.get("_date")
            if d == now:
                today += value
            if d and d.year == now.year and d.month == now.month:
                month_total += value
        self.card_total._value.setText(f"{format_amount(total)} {CURRENCY}")
        self.card_today._value.setText(f"{format_amount(today)} {CURRENCY}")
        self.card_month._value.setText(f"{format_amount(month_total)} {CURRENCY}")
        self.card_rows._value.setText(str(count))
        self.chart.set_data(chart_months(self.rows), tr("chart_empty", self.lang))
        self.completer.setModel(QStringListModel(unique_descriptions(self.rows), self.completer))

    def cache_put(self, path, rows):
        try:
            mtime = os.path.getmtime(path)
        except OSError:
            mtime = 0
        self.cache[path] = (mtime, rows)

    def _cached(self, path):
        entry = self.cache.get(path)
        try:
            return bool(entry) and entry[0] == os.path.getmtime(path)
        except OSError:
            return False

    def on_search_changed(self, *_):
        self._search_timer.start(180)

    def run_search(self):
        scope = self.scope_combo.currentData()
        self.month_combo.setEnabled(scope == "file")
        if scope == "file":
            self._search_id += 1  # stop any background indexing
            self._indexing = False
            self.update_controls()
            self.apply_table_filter()
            if not search_terms(self.desc_search.text()):
                self.notify(tr("files_found", self.lang, count=len(self.files)), "info")
        else:
            if self.current_row is not None:
                self.clear_form()
            self.start_global_search()

    def start_global_search(self):
        self._search_id += 1
        sid = self._search_id
        terms = search_terms(self.desc_search.text())
        if not terms:
            self._indexing = False
            self.results = []
            self.results_table.setRowCount(0)
            self.empty_label.setText(tr("search_all_hint", self.lang))
            self.stack.setCurrentIndex(1)
            self.update_controls()
            return
        self._index_queue = [p for p in self.files if not self._cached(p)]
        self._index_total = len(self._index_queue)
        self._index_done = 0
        self._indexing = bool(self._index_queue)
        self.update_controls()
        self.render_results()
        if self._index_queue:
            self._index_next(sid)

    def _index_next(self, sid):
        """Index one file at a time so edits/loads are never stuck behind a long queue."""
        if sid != self._search_id:
            return
        if not self._index_queue:
            self._indexing = False
            self.update_controls()
            self.render_results()
            return
        path = self._index_queue.pop(0)

        def job():
            if sid != self._search_id:
                return None
            _, rows = self.manager.load_project(path)
            return prepare_rows(rows)

        def finish(rows):
            if rows is not None:
                self.cache_put(path, rows)
                self._index_done += 1
            if sid != self._search_id:
                return
            self.render_results()
            self._index_next(sid)

        def ok(rows):
            finish(rows)

        def err(_exc):
            finish([])  # unreadable / not a project file: skip it

        self.worker.submit(job, ok, err)

    def render_results(self):
        terms = search_terms(self.desc_search.text())
        low, high = parse_filter_amount(self.min_amount.text()), parse_filter_amount(self.max_amount.text())
        self.results = [m for m in collect_matches(self.files, self.cache, terms) if amount_in_range(m[2], low, high)]
        t = self.results_table
        t.setUpdatesEnabled(False)
        t.clearContents()
        t.setRowCount(len(self.results))
        trailing = Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter
        leading = Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter
        center = Qt.AlignmentFlag.AlignCenter
        for i, (name, path, r) in enumerate(self.results):
            proj = QTableWidgetItem(name)
            proj.setData(Qt.ItemDataRole.UserRole, path)
            proj.setData(ROLE_INDEX, r["excel_index"])
            proj.setToolTip(path)
            proj.setTextAlignment(leading)
            id_item = QTableWidgetItem(format_id(r.get("id"))); id_item.setTextAlignment(center)
            desc = QTableWidgetItem("" if r.get("description") is None else str(r.get("description")))
            desc.setTextAlignment(leading)
            date_item = QTableWidgetItem(excel_date_text(r.get("date"))); date_item.setTextAlignment(center)
            amount = QTableWidgetItem(format_amount(r.get("amount"))); amount.setTextAlignment(trailing)
            for c, it in enumerate((proj, id_item, desc, date_item, amount)):
                t.setItem(i, c, it)
        t.setUpdatesEnabled(True)
        files_hit = len({p for _n, p, _r in self.results})
        total = sum(r["_amount"] for _n, _p, r in self.results if r.get("_numeric"))
        if self._indexing:
            self.notify(tr("indexing", self.lang, done=self._index_done, total=self._index_total), "info")
            self.totals_label.setText("")
        else:
            self.notify(tr("results", self.lang, n=len(self.results), m=files_hit), "info")
            self.totals_label.setText(tr("results_total", self.lang, n=len(self.results), m=files_hit,
                                         total=format_amount(total)))
        if self.results:
            self.stack.setCurrentIndex(2)
        else:
            self.empty_label.setText(tr("indexing", self.lang, done=self._index_done, total=self._index_total)
                                     if self._indexing else tr("no_results", self.lang))
            self.stack.setCurrentIndex(1)

    def open_result(self, item):
        row = item.row()
        first = self.results_table.item(row, 0)
        if first is None:
            return
        path = first.data(Qt.ItemDataRole.UserRole)
        idx = first.data(ROLE_INDEX)
        self._focus_target = (path, idx)
        # back to the single-file view, with the search cleared so the row is visible
        self._search_id += 1
        self._indexing = False
        self.scope_combo.blockSignals(True); self.scope_combo.setCurrentIndex(0); self.scope_combo.blockSignals(False)
        self.desc_search.blockSignals(True); self.desc_search.clear(); self.desc_search.blockSignals(False)
        self.month_combo.setEnabled(True)
        self.update_controls()
        if self.file_ok and (self.current_file or "").lower() == path.lower():
            self.apply_table_filter()
            self._focus_row()
            return
        for i in range(self.project_list.count()):
            it = cast(ProjectItem, self.project_list.item(i))
            if it.path.lower() == path.lower():
                if it.isHidden():
                    self.search.clear()
                self.project_list.setCurrentRow(i)
                break
        self.apply_table_filter()

    def _focus_row(self):
        if not self._focus_target:
            return
        path, idx = self._focus_target
        if not self.file_ok or (self.current_file or "").lower() != path.lower():
            return
        self._focus_target = None
        self.month_combo.blockSignals(True); self.month_combo.setCurrentIndex(0); self.month_combo.blockSignals(False)
        self.apply_table_filter()
        for i in range(self.table.rowCount()):
            cell = self.table.item(i, 0)
            if cell is not None and cell.data(Qt.ItemDataRole.UserRole) == idx:
                self.table.selectRow(i)
                self.table.scrollToItem(cell)
                break

    def create_project(self):
        folder = self.folder_label.text()
        if not folder or not os.path.isdir(folder):
            self.choose_folder()
            folder = self.folder_label.text()
        if not folder or not os.path.isdir(folder):
            return
        dlg = NewProjectDialog(folder, self.manager, self.lang, self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        try:
            name, template = dlg.values()
        except ValueError as exc:
            QMessageBox.warning(self, tr("error", self.lang), str(exc))
            return
        target = os.path.join(folder, name)
        if os.path.exists(target):
            QMessageBox.warning(self, tr("error", self.lang), tr("already_exists", self.lang))
            return

        def ok(_res):
            self.end_busy()
            self.notify(tr("created", self.lang), "ok")
            self.current_file = target
            self.file_ok = False
            self.scan_folder()

        def err(exc):
            self.end_busy()
            QMessageBox.critical(self, tr("error", self.lang), str(exc))

        self.notify(tr("saving", self.lang), "info")
        self.begin_busy()
        self.worker.submit(lambda: self.manager.create_from_template(template, target), ok, err)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        panel = getattr(self, "chart_panel", None)
        if panel is not None:  # on short windows the table matters more than the chart
            panel.setVisible(self.height() >= 700)

    # ------------------------------------------------------------- theme / sorting / chart
    def apply_theme(self):
        self.setStyleSheet(build_style(self.theme))
        self.chart.set_colors(THEMES[self.theme])
        palette = self.status.palette()
        palette.setColor(QPalette.ColorRole.Link, QColor(THEMES[self.theme]["accent_text"]))
        self.status.setPalette(palette)
        self.theme_btn.setText(tr("theme_light", self.lang) if self.theme == "dark" else tr("theme_dark", self.lang))

    def toggle_theme(self):
        self.theme = "light" if self.theme == "dark" else "dark"
        self.settings.setValue("theme", self.theme)
        self.apply_theme()

    def on_header_clicked(self, col):
        """Click a column header: ascending, then descending, then back to the default order."""
        if col >= len(SORT_COLUMNS):
            return
        key = SORT_COLUMNS[col]
        if self.sort_column != key:
            self.sort_column, self.sort_asc = key, True
        elif self.sort_asc:
            self.sort_asc = False
        else:
            self.sort_column, self.sort_asc = None, True
        self.refresh_table_headers()
        self.apply_table_filter()

    def on_chart_month(self, key):
        if self.scope_combo.currentData() == "all":
            return
        target = "all" if self.month_combo.currentData() == key else key
        idx = self.month_combo.findData(target)
        if idx >= 0:
            self.month_combo.setCurrentIndex(idx)

    # ------------------------------------------------------------- unsaved entry guard
    def form_dirty(self):
        desc, amount = self.description.text().strip(), self.amount.text().strip()
        cur = self.current_row
        if cur is None:
            return bool(desc or amount)
        d = self.date_edit.date()
        same_date = cur.get("_date") is None or cur["_date"] == date(d.year(), d.month(), d.day())
        return not (desc == _norm_text(cur.get("description")) and amount == format_amount(cur.get("amount")) and same_date)

    def confirm_discard(self):
        answer = QMessageBox.question(
            self, tr("unsaved_title", self.lang), tr("unsaved_confirm", self.lang),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        return answer == QMessageBox.StandardButton.Yes

    def _reselect_current(self):
        for i in range(self.project_list.count()):
            if cast(ProjectItem, self.project_list.item(i)).path == self.current_file:
                self.project_list.blockSignals(True)
                self.project_list.setCurrentRow(i)
                self.project_list.blockSignals(False)
                break

    # ------------------------------------------------------------- backups
    def open_backup_folder(self):
        if not self.current_file:
            return
        folder = backup_dir_for(self.current_file)
        try:
            folder.mkdir(exist_ok=True)
            os.startfile(str(folder))
        except Exception as exc:  # noqa: BLE001
            QMessageBox.critical(self, tr("error", self.lang), str(exc))

    def restore_from_backup(self):
        if not (self.current_file and self.file_ok and self._pending == 0):
            return
        path = self.current_file
        dialog = RestoreDialog(path, self.lang, self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.selected:
            return
        backup = dialog.selected
        when = backup["when"].strftime("%d.%m.%Y %H:%M:%S")
        answer = QMessageBox.question(
            self, tr("restore_title", self.lang), tr("restore_confirm", self.lang, when=when),
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if answer != QMessageBox.StandardButton.Yes:
            return

        def job():
            restore_backup(path, backup["path"])
            meta, rows = self.manager.load_project(path)
            return meta, prepare_rows(rows)

        def ok(res):
            self.end_busy()
            meta, rows = res
            self.cache_put(path, rows)
            self.notify(tr("restored", self.lang), "ok")
            if path == self.current_file:
                self.clear_form()
                self.apply_result({"rows": rows, "meta": meta})

        def err(exc):
            self.end_busy()
            self.cache.pop(path, None)
            message = tr("restore_locked", self.lang) if isinstance(exc, PermissionError) else str(exc)
            self.notify(message, "err")
            QMessageBox.critical(self, tr("error", self.lang), message)

        self.notify(tr("restoring", self.lang), "info")
        self.begin_busy()
        self.worker.submit(job, ok, err)

    # ------------------------------------------------------------- reports / export
    def open_file(self, path):
        with suppress(Exception):
            os.startfile(path)

    def open_report_dialog(self):
        if not (self.current_file and self.file_ok and self._pending == 0):
            return
        month = self.month_combo.currentData() if self.month_combo.count() else "all"
        month = month if isinstance(month, str) else "all"
        text = self.desc_search.text() if self.scope_combo.currentData() == "file" else ""
        project = Path(self.current_file).stem
        dialog = ReportDialog(self.rows, project, self.lang, month, text, self)
        if dialog.exec() != QDialog.DialogCode.Accepted or not dialog.choice:
            return
        sel = dialog.selection()
        payload = build_report_payload(project, self.rows, sel["from"], sel["to"], sel["text"], sel["all_dates"], self.lang)
        if not payload["rows"]:
            QMessageBox.information(self, tr("report_title", self.lang), tr("no_data_in_range", self.lang))
            return
        if sel["all_dates"]:
            stem = f"{project}_{datetime.now():%Y-%m-%d}"  # noqa: DTZ005
        else:
            stem = f"{project}_{sel['from']:%Y-%m-%d}_{sel['to']:%Y-%m-%d}"
        kind = dialog.choice
        ext, name_filter, title_key = (("xlsx", "Excel Workbook (*.xlsx)", "save_excel_title") if kind == "xlsx"
                                       else ("pdf", "PDF (*.pdf)", "save_pdf_title"))
        start_dir = self.settings.value("export_dir", "", type=str)
        if not (isinstance(start_dir, str) and os.path.isdir(start_dir)):
            start_dir = os.path.dirname(self.current_file)
        target, _ = QFileDialog.getSaveFileName(self, tr(title_key, self.lang),
                                                os.path.join(start_dir, safe_filename(stem) + "." + ext), name_filter)
        if not target:
            return
        if not target.lower().endswith("." + ext):
            target += "." + ext
        self.settings.setValue("export_dir", os.path.dirname(target))
        if kind == "pdf":
            self._export_pdf(target, payload)
        else:
            self._export_xlsx(target, payload)

    def _export_pdf(self, target, payload):
        self.notify(tr("creating_pdf", self.lang), "info")
        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            QApplication.processEvents()
            write_pdf_report(target, payload, self.lang)
        except Exception as exc:  # noqa: BLE001
            self.notify(str(exc), "err")
            QMessageBox.critical(self, tr("error", self.lang), str(exc))
            return
        finally:
            QApplication.restoreOverrideCursor()
        self.notify_file("pdf_done", target)
        self.open_file(target)

    def _xlsx_data(self, payload):
        lang = self.lang
        t = lambda k: tr(k, lang)  # noqa: E731
        epoch = date(1899, 12, 30)
        rows = []
        for r in payload["rows"]:
            ident = format_id(r.get("id"))
            d = r.get("_date")
            amount = r["_amount"] if r.get("_numeric") and r.get("amount") not in (None, "") else (_norm_text(r.get("amount")) or None)
            rows.append((
                int(ident) if ident.isdigit() else (ident or None),
                _norm_text(r.get("description")),
                float((d - epoch).days) if d else (_norm_text(r.get("date")) or None),
                amount,
            ))
        meta = [(t("report_project"), payload["project"]), (t("report_period"), payload["period"]),
                (t("report_generated"), payload["generated"])]
        if payload["filter"]:
            meta.append((t("report_filter"), payload["filter"]))
        meta.append((t("report_total"), f"{_fmt(payload['summary']['total'])} {CURRENCY}  ({payload['summary']['count']})"))
        return {
            "rtl": is_rtl(lang), "rows": rows, "meta": meta,
            "headers": [t("id"), t("description"), t("date"), t("amount_dzd")],
            "months": [(k, c, v) for k, v, c in payload["months"]],
            "labels": {"heading": t("report_heading"), "total": t("col_total"), "col_month": t("col_month"),
                       "col_count": t("col_count"), "chart_title": t("chart_title"),
                       "sheet_expenses": t("sheet_expenses"), "sheet_monthly": t("sheet_monthly")},
        }

    def _export_xlsx(self, target, payload):
        data = self._xlsx_data(payload)

        def ok(_res):
            self.end_busy()
            self.notify_file("export_done", target)
            self.open_file(target)

        def err(exc):
            self.end_busy()
            self.notify(str(exc), "err")
            QMessageBox.critical(self, tr("error", self.lang), str(exc))

        self.notify(tr("exporting", self.lang), "info")
        self.begin_busy()
        self.worker.submit(lambda: self.manager.export_xlsx(target, data), ok, err)

    def closeEvent(self, event):
        if self.form_dirty() and not self.confirm_discard():
            event.ignore()
            return
        self.worker.shutdown(self.manager.stop_excel)
        event.accept()

def resource_path(name):
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
    return base / name


def main():
    if sys.platform != "win32":
        print("This application requires Windows and Microsoft Excel.")
        return 1

    # This explicit recovery mode opens a normal GUI dialog, including in a
    # windowed PyInstaller build that has no console attached.
    reset_mode = "--reset-password" in sys.argv[1:]
    if reset_mode:
        sys.argv[:] = [sys.argv[0]] + [arg for arg in sys.argv[1:] if arg != "--reset-password"]

    with suppress(Exception):
        import ctypes
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("BuildOffice.ProjectExpenseEntry")
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setLayoutDirection(layout_direction_for(saved_language()))
    icon = resource_path("App.ico")
    if icon.exists():
        app.setWindowIcon(QIcon(str(icon)))

    if reset_mode:
        dialog = PasswordResetDialog()
        if dialog.exec() == QDialog.DialogCode.Accepted:
            try:
                set_password(dialog.password())
            except Exception as exc:  # noqa: BLE001
                QMessageBox.critical(None, "Password reset failed", str(exc))
                return 1
            QMessageBox.information(None, "Password updated", "The password was saved. Restart the application normally to sign in.")
        return 0

    window = LoginWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
