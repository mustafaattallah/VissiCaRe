# piece 1 (UPDATED — full chunk)
from __future__ import annotations

import os
import sys
import json
from pathlib import Path
from typing import Any, Dict, List, Tuple, Optional
from PyQt6.QtGui import QKeySequence, QAction
from PyQt6.QtCore import Qt, QDate, QObject, QEvent
from PyQt6.QtGui import QIcon, QIntValidator, QDoubleValidator, QPalette
from PyQt6.QtGui import QPixmap, QPainter, QFont, QColor, QPainterPath, QPen
from PyQt6.QtWidgets import QSplashScreen
from PyQt6.QtCore import QRect, QTimer
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QTabWidget,
    QVBoxLayout,
    QHBoxLayout,
    QGridLayout,
    QLabel,
    QLineEdit,
    QPushButton,
    QGroupBox,
    QScrollArea,
    QSpinBox,
    QComboBox,
    QCheckBox,
    QTextEdit,
    QTextBrowser,
    QDateEdit,
    QFileDialog,
    QMessageBox,
    QFrame,
    QListWidget,
    QListWidgetItem,
    QAbstractItemView,
    QDoubleSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QHeaderView,
    QToolButton,
    QSizePolicy,
    QStyle,
    QStyleOptionButton,
    QStyleOptionGroupBox,
    QToolTip,
)

# ============================================================
# Paths / resources (works for .py and PyInstaller .exe)
# ============================================================

def app_dir() -> Path:
    # folder of the exe when frozen, else folder of this .py file
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent

def resource_path(rel_name: str) -> Path:
    # inside PyInstaller bundle => sys._MEIPASS
    base = Path(getattr(sys, "_MEIPASS", app_dir()))
    return base / rel_name

APP_DIR = app_dir()
ICON_PATH = str(resource_path("VissiCaRe_icon.ico"))


# ============================================================
# License state (placeholder - editable later)
# ============================================================

class LicenseState:
    # Default: Basic
    level: str = "Basic"  # "Basic" or "Gold"
    key: str = ""

    @classmethod
    def set_license(cls, key: str, level: str):
        cls.key = key
        cls.level = level

    @classmethod
    def is_gold(cls) -> bool:
        return cls.level.strip().lower() == "gold"


# ============================================================
# Types
# ============================================================

Scenario = Tuple[int, str, bool]
FromTo = Tuple[int, int]
NodeRowTuple = Tuple[int, str, str, bool, List[FromTo]]  # (Node#, Signalized/Unsignalized, Type, Overpass, [(From,To),...])

SegmentRow = Tuple[str, str, Tuple[int, ...]]  # (Segment Name, Segment Type, (link numbers...))

TTDirection = str
TTEntry = Tuple[str, TTDirection, Tuple[int, ...]]  # (MainLine, Direction, (TT#,...))

TPDirection = str
TPEntry = Tuple[str, TPDirection, Tuple[int, ...]]  # (MainLine, Direction, (Node#,...))

FORMAT_SCOPES = ["Overall Intersection", "Approach", "Movement"]

FORMAT_VAR_DISPLAY_TO_KEY = {
    "Delay": "Delay",
    "Vehicle": "Vehicle",
    "Queue Length": "QLen",
    "Maximum Queue": "QMax",
}
FORMAT_VAR_DISPLAY = list(FORMAT_VAR_DISPLAY_TO_KEY.keys())
WRAPPERS = ["", "[]", "{}", "{{}}", "[[]]", "()", "(())", "{| |}"]


VISSIM_PARAMS_ORDERED = [
    "SafDistFactLnChg",
    "RecovSpeed",
    "RecovSlow",
    "RecovSafDist",
    "NumInteractVeh",
    "NumInteractObj",
    "W99cc0",
    "W99cc1Distr",
    "W99cc2",
    "W99cc3",
    "W99cc4",
    "W99cc5",
    "W99cc6",
    "W99cc7",
    "W99cc8",
    "W99cc9",
    "W74ax",
    "W74bxAdd",
    "W74bxMult",
    "LookAheadDistMax",
    "LookAheadDistMin",
    "LookBackDistMax",
    "LookBackDistMin",
    "ACC_StandSafDist",
    "ACC_MinGapTime",
    "MinFrontRearClear",
    "StandDist",
    "StandDistIsFix",
    "MaxDecelOwn",
    "MaxDecelTrail",
    "AccDecelOwn",
    "AccDecelTrail",
    "DecelRedDistOwn",
    "DecelRedDistTrail",
    "CoopDecel",
    "VehRoutDecLookAhead",
    "Zipper",
    "ZipperMinSpeed",
    "CoopLnChg",
    "CoopLnChgSpeedDiff",
    "CoopLnChgCollTm",
    "ObsrvAdjLn",
    "DiamQueu",
    "ConsNextTurn",
    "MinCollTmGain",
    "MinSpeedForLat",
    "EnforcAbsBrakDist",
    "UseImplicStoch",
    "PlatoonPoss",
    "PlatoonMinClear",
    "MaxNumPlatoonVeh",
    "MaxPlatoonDesSpeed",
    "MaxPlatoonApprDist",
    "PlatoonFollowUpGapTm",
    "MesoReactTm",
    "MesoStandDist",
    "MesoMaxWaitTm",
]
BOOLEAN_DB_PARAMS = {
    "StandDistIsFix",
    "VehRoutDecLookAhead",
    "Zipper",
    "CoopLnChg",
    "ObsrvAdjLn",
    "DiamQueu",
    "ConsNextTurn",
    "EnforcAbsBrakDist",
    "UseImplicStoch",
    "PlatoonPoss",
}
def _set_combo_text(combo: QComboBox, text: str):
    """Set combo to text; if missing, add or edit depending on combo settings."""
    if text is None:
        return
    idx = combo.findText(text)
    if idx >= 0:
        combo.setCurrentIndex(idx)
    else:
        if combo.isEditable():
            combo.setEditText(text)
        else:
            combo.addItem(text)
            combo.setCurrentText(text)

def _qdate_to_str(d: QDate) -> str:
    return d.toString("MM/dd/yyyy")

def _str_to_qdate(s: str) -> QDate:
    d = QDate.fromString(s or "", "MM/dd/yyyy")
    return d if d.isValid() else QDate.currentDate()

def _force_input_bg(*widgets: QWidget) -> None:
    for w in widgets:
        w.setStyleSheet("background-color: #ffffff;")

def _force_input_bg_in_widget(root: QWidget) -> None:
    if isinstance(root, (QLineEdit, QComboBox, QSpinBox, QDoubleSpinBox)):
        _force_input_bg(root)
    for w in root.findChildren(QLineEdit):
        _force_input_bg(w)
    for w in root.findChildren(QComboBox):
        _force_input_bg(w)
    for w in root.findChildren(QSpinBox):
        _force_input_bg(w)
    for w in root.findChildren(QDoubleSpinBox):
        _force_input_bg(w)

# ============================================================
# Styles
# ============================================================

SECTION_GROUPBOX_STYLE = """
QGroupBox {
    font-weight: 800;
    border: 1px solid #c9d4e3;
    border-radius: 12px;
    margin-top: 18px;
    padding: 20px 12px 12px 12px;
    background: qlineargradient(
        x1: 0, y1: 0,
        x2: 1, y2: 0,
        stop: 0  #fff7d6,
        stop: 0.5 #ffffff,
        stop: 1  #e4f0ff
    );
}
QGroupBox QLabel { background: transparent; }
QGroupBox QLineEdit,
QGroupBox QComboBox,
QGroupBox QSpinBox,
QGroupBox QDoubleSpinBox,
QGroupBox QAbstractSpinBox {
    background-color: #ffffff;
}
QGroupBox QAbstractSpinBox::lineEdit {
    background-color: #ffffff;
}
QGroupBox QComboBox:editable {
    background-color: #ffffff;
}
QGroupBox QComboBox QAbstractItemView {
    background-color: #ffffff;
}
QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    padding: 3px 10px;
    color: #143a5d;
    background: #ffffff;
    border: 1px solid #c9d4e3;
    border-radius: 9px;
}
"""

TAB_BACKGROUND_STYLE = """
background-color: #f6f8fb;
QLabel { background: transparent; }
QLineEdit,
QComboBox,
QAbstractSpinBox,
QSpinBox,
QDoubleSpinBox {
    background-color: #ffffff;
}
QAbstractSpinBox::lineEdit {
    background-color: #ffffff;
}
QComboBox:editable {
    background-color: #ffffff;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
}
"""

TRANSPARENT_CONTAINER_STYLE = """
background: transparent;
QLineEdit,
QComboBox,
QAbstractSpinBox,
QSpinBox,
QDoubleSpinBox {
    background-color: #ffffff;
}
QAbstractSpinBox::lineEdit {
    background-color: #ffffff;
}
QComboBox:editable {
    background-color: #ffffff;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
}
"""

PRIMARY_BUTTON_STYLE = """
QPushButton {
    background: qlineargradient(
        x1: 0, y1: 0,
        x2: 1, y2: 0,
        stop: 0  #ffe066,
        stop: 0.5 #fffef5,
        stop: 1  #cfe4ff
    );
    border: 1px solid #b9c6d8;
    border-radius: 10px;
    padding: 7px 18px;
    font-weight: 900;
}
QPushButton:hover {
    background: qlineargradient(
        x1: 0, y1: 0,
        x2: 1, y2: 0,
        stop: 0  #ffd43b,
        stop: 0.5 #fff7cf,
        stop: 1  #b3d4ff
    );
}
QPushButton:pressed { background-color: #fcc419; }
"""

SOFT_BUTTON_STYLE = """
QPushButton {
    background: #ffffff;
    border: 1px solid #c9d4e3;
    border-radius: 10px;
    padding: 6px 14px;
    font-weight: 700;
}
QPushButton:hover { background: #f3f7ff; }
QPushButton:pressed { background: #e6efff; }
"""

LOG_STYLE = """
QTextEdit {
    background-color: #ffffff;
    border-radius: 10px;
    border: 1px solid #c9d4e3;
    padding: 8px;
    font-family: Consolas, monospace;
    font-size: 11pt;
}
"""

SUBTITLE_STYLE = "QLabel { font-weight: 900; color: #143a5d; padding: 4px 0px; }"

COLLAPSE_HEADER_STYLE = """
QToolButton {
    text-align: left;
    font-weight: 900;
    font-size: 13pt;
    color: #143a5d;

    background: qlineargradient(
        x1: 0, y1: 0,
        x2: 1, y2: 0,
        stop: 0  #fff7d6,
        stop: 0.5 #ffffff,
        stop: 1  #e4f0ff
    );
    border: 1px solid #c9d4e3;
    border-radius: 10px;
    padding: 10px 12px;
}
QToolButton:hover {
    background: qlineargradient(
        x1: 0, y1: 0,
        x2: 1, y2: 0,
        stop: 0  #ffeaa6,
        stop: 0.5 #ffffff,
        stop: 1  #d6ebff
    );
}
"""

HELP_BROWSER_STYLE = """
QTextBrowser {
    background-color: #ffffff;
    border-radius: 10px;
    border: 1px solid #c9d4e3;
    padding: 10px;
    font-size: 11pt;
}
QTextBrowser a {
    color: #0b5ed7;
    text-decoration: none;
    font-weight: 800;
}
QTextBrowser a:hover { text-decoration: underline; }
"""

# ============================================================
# ✅ Global modern scrollbar style (Vertical + Horizontal)
# Apply once at app startup via apply_global_scrollbar_style(app)
# ============================================================

GLOBAL_SCROLLBAR_STYLE = """
/* ScrollArea base: keep clean */
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }

/* ===== Vertical ScrollBar ===== */
QScrollBar:vertical {
    background: transparent;
    width: 12px;
    margin: 2px 2px 2px 2px;
}
QScrollBar::handle:vertical {
    background: #b9c7dd;
    min-height: 30px;
    border-radius: 6px;
}
QScrollBar::handle:vertical:hover { background: #9fb3cf; }
QScrollBar::handle:vertical:pressed { background: #86a4ca; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
    background: transparent;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: transparent;
}

/* ===== Horizontal ScrollBar ===== */
QScrollBar:horizontal {
    background: transparent;
    height: 12px;
    margin: 2px 2px 2px 2px;
}
QScrollBar::handle:horizontal {
    background: #b9c7dd;
    min-width: 30px;
    border-radius: 6px;
}
QScrollBar::handle:horizontal:hover { background: #9fb3cf; }
QScrollBar::handle:horizontal:pressed { background: #86a4ca; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
    background: transparent;
}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
    background: transparent;
}

/* Corner between horizontal+vertical */
QAbstractScrollArea::corner { background: transparent; }
"""

def apply_global_scrollbar_style(app: QApplication):
    """
    Call this ONCE right after you create QApplication.
    It appends the global scrollbar QSS without removing your other styles.
    """
    existing = app.styleSheet() or ""
    if GLOBAL_SCROLLBAR_STYLE.strip() in existing:
        return
    app.setStyleSheet(existing + "\n\n" + GLOBAL_SCROLLBAR_STYLE)


class CollapsibleSection(QWidget):
    """
    Clickable header bar with an arrow. Expands/collapses content widget.
    Not a checkbox (uses a QToolButton).
    """
    def __init__(self, title: str, content: QWidget, parent: Optional[QWidget] = None):
        super().__init__(parent)

        self.toggle_btn = QToolButton()
        self.toggle_btn.setText(title)
        self.toggle_btn.setCheckable(True)
        self.toggle_btn.setChecked(False)
        self.toggle_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.toggle_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.toggle_btn.setStyleSheet(COLLAPSE_HEADER_STYLE)

        self.content_area = QWidget()
        self.content_layout = QVBoxLayout(self.content_area)
        self.content_layout.setContentsMargins(0, 10, 0, 0)
        self.content_layout.addWidget(content)

        self.content_area.setVisible(False)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self.toggle_btn)
        lay.addWidget(self.content_area)

        self.toggle_btn.toggled.connect(self._on_toggled)
        self._layout_path = getattr(self, "_layout_path", None)  # keep if you already set it elsewhere
        self._dirty = False
        self._loading = False

        self._update_window_title()
        self._wire_dirty_tracking()

    def _on_toggled(self, checked: bool):
        self.toggle_btn.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)
        self.content_area.setVisible(checked)

    def _update_window_title(self):
        name = "VissiCaRe"
        if getattr(self, "_layout_path", None):
            base = os.path.basename(self._layout_path)
            name += f" - {base}"
        if getattr(self, "_dirty", False):
            name += " *"
        self.setWindowTitle(name)

    def _set_dirty(self, on: bool):
        self._dirty = bool(on)
        self._update_window_title()

    def _mark_dirty(self, *args, **kwargs):
        # Ignore signals while loading/applying layouts
        if getattr(self, "_loading", False):
            return
        if not getattr(self, "_dirty", False):
            self._set_dirty(True)

    def _maybe_save(self) -> bool:
        """
        Returns True if OK to continue (saved / discarded).
        Returns False if user cancelled.
        """
        if not getattr(self, "_dirty", False):
            return True

        mb = QMessageBox(self)
        mb.setIcon(QMessageBox.Icon.Warning)
        mb.setWindowTitle("Unsaved changes")
        mb.setText("You have unsaved changes. Do you want to save before continuing?")

        save_btn = mb.addButton("Save", QMessageBox.ButtonRole.AcceptRole)
        discard_btn = mb.addButton("Don't Save", QMessageBox.ButtonRole.DestructiveRole)
        cancel_btn = mb.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)

        mb.setDefaultButton(save_btn)
        mb.exec()

        clicked = mb.clickedButton()
        if clicked == save_btn:
            self._save_layout()  # uses your Save logic
            # if user cancels Save As dialog, dirty will remain True
            return not getattr(self, "_dirty", False)

        if clicked == discard_btn:
            return True

        return False

    def closeEvent(self, event):
        # Ask to save when exiting (X button, Alt+F4, File->Exit)
        if self._maybe_save():
            event.accept()
        else:
            event.ignore()

    def _wire_dirty_tracking(self):
        """
        Mark project dirty when user edits inputs.
        Uses user-only signals where possible.
        """
        # QLineEdit: user edits only
        for w in self.findChildren(QLineEdit):
            try:
                w.textEdited.connect(self._mark_dirty)
            except Exception:
                pass

        # ComboBoxes / CheckBoxes
        for w in self.findChildren(QComboBox):
            try:
                w.currentIndexChanged.connect(self._mark_dirty)
            except Exception:
                pass

        for w in self.findChildren(QCheckBox):
            try:
                w.toggled.connect(self._mark_dirty)
            except Exception:
                pass

        # SpinBoxes
        for w in self.findChildren(QSpinBox):
            try:
                w.valueChanged.connect(self._mark_dirty)
            except Exception:
                pass

        for w in self.findChildren(QDoubleSpinBox):
            try:
                w.valueChanged.connect(self._mark_dirty)
            except Exception:
                pass

        # Tables (only needed if you allow editing table items)
        for w in self.findChildren(QTableWidget):
            try:
                w.cellChanged.connect(self._mark_dirty)
            except Exception:
                pass


# ============================================================
# Tooltip helper (HOVER on the section title ONLY — not the background)
# ============================================================

class _GroupBoxTitleOnlyTooltipFilter(QObject):
    """
    Makes a QGroupBox tooltip show ONLY when hovering the groupbox TITLE area,
    not anywhere on the background/content.
    """
    def eventFilter(self, obj, event):
        if isinstance(obj, QGroupBox) and event.type() == QEvent.Type.ToolTip:
            # No tooltip? let it pass
            if not obj.toolTip():
                return False

            opt = QStyleOptionGroupBox()
            obj.initStyleOption(opt)

            # Title/label rect from style
            title_rect = obj.style().subControlRect(
                QStyle.ComplexControl.CC_GroupBox,
                opt,
                QStyle.SubControl.SC_GroupBoxLabel,
                obj,
            )

            # small padding so it's easier to hit the title
            title_rect = title_rect.adjusted(-10, -6, 10, 6)

            if title_rect.contains(event.pos()):
                return False  # allow tooltip normally (Qt will show it)

            # Otherwise, swallow tooltip event
            QToolTip.hideText()
            return True

        return super().eventFilter(obj, event)


def make_help_asterisk(help_text: str) -> QLabel:
    """
    Compatibility stub: previously returned a visible '*'.
    Now returns an invisible label so nothing changes elsewhere.
    """
    a = QLabel("")
    a.setFixedWidth(0)
    a.setFixedHeight(0)
    a.setToolTip(help_text)
    return a

def add_corner_help(group: QGroupBox, help_text: str):
    """
    OLD: tooltip showed when hovering ANYWHERE in the section background.
    NEW: tooltip shows ONLY when hovering the GROUPBOX TITLE area.
    """
    group.setToolTip(help_text)

    # install filter once and keep it alive by storing it on the group
    if not hasattr(group, "_title_only_tt_filter"):
        group._title_only_tt_filter = _GroupBoxTitleOnlyTooltipFilter(group)
        group.installEventFilter(group._title_only_tt_filter)

def hline() -> QFrame:
    line = QFrame()
    line.setFrameShape(QFrame.Shape.HLine)
    line.setFrameShadow(QFrame.Shadow.Sunken)
    line.setStyleSheet("QFrame { color: #c9d4e3; }")
    return line


from PyQt6.QtWidgets import QTableWidget  # already in your imports

def _set_table_row_count(table: QTableWidget, target_rows: int, add_row_fn=None):
    """
    Ensures `table` has exactly `target_rows` rows.

    Supports two call styles:
      1) _set_table_row_count(table, rows)
         -> uses table.insertRow(...) for new rows

      2) _set_table_row_count(table, rows, add_row_fn)
         -> uses add_row_fn to create each row (widgets/validators/etc)
            add_row_fn may be add_row_fn(table) or add_row_fn()
    """
    target_rows = max(0, int(target_rows))

    # Remove extra rows
    while table.rowCount() > target_rows:
        table.removeRow(table.rowCount() - 1)

    # Add missing rows
    while table.rowCount() < target_rows:
        if add_row_fn is None:
            table.insertRow(table.rowCount())
        else:
            try:
                add_row_fn(table)   # e.g. self._table_add_row(table)
            except TypeError:
                add_row_fn()        # e.g. self._validation_add_row()


# ============================================================
# Blue checkmark checkbox (indicator checkmark like your image)
# ============================================================

class BlueCheckBox(QCheckBox):
    """
    QCheckBox that always shows a visible checkmark when checked (✓),
    with a blue check mark (matches your provided image, but blue).
    """
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        opt = QStyleOptionButton()
        self.initStyleOption(opt)

        # Draw checkbox WITHOUT the default checked glyph (we'll draw our own)
        is_on = bool(opt.state & QStyle.StateFlag.State_On)
        if is_on:
            opt.state &= ~QStyle.StateFlag.State_On
            opt.state |= QStyle.StateFlag.State_Off

        self.style().drawControl(QStyle.ControlElement.CE_CheckBox, opt, painter, self)

        # Draw our blue checkmark when checked
        if self.isChecked():
            ind = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, opt, self)

            # slight inset for nicer geometry
            r = ind.adjusted(3, 3, -3, -3)

            # check color
            blue = QColor("#0b5ed7") if self.isEnabled() else QColor("#9fb3cf")

            pen = QPen(blue, 2.6, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
            painter.setPen(pen)

            path = QPainterPath()
            x = float(r.x()); y = float(r.y())
            w = float(r.width()); h = float(r.height())

            # classic checkmark: (left-mid) -> (mid-bottom) -> (right-top)
            path.moveTo(x + 0.10 * w, y + 0.55 * h)
            path.lineTo(x + 0.42 * w, y + 0.85 * h)
            path.lineTo(x + 0.95 * w, y + 0.15 * h)

            painter.drawPath(path)

        painter.end()


# ============================================================
# Digits-only line edit (for huge IDs safely)
# ============================================================

class DigitsLineEdit(QLineEdit):
    def __init__(self, text: str = "", parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        self.setValidator(None)
        self.textChanged.connect(self._sanitize)

    def _sanitize(self):
        t = self.text()
        if t == "":
            return
        if not t.isdigit():
            self.setText("".join(ch for ch in t if ch.isdigit()))

    def as_int(self) -> int:
        t = self.text().strip()
        if t == "":
            return 0
        try:
            return int(t)
        except ValueError:
            return 0

# ============================================================
# TEST STUBS (replace later with your core)
# ============================================================

def Simulation_Nodes(
    Vissim_Version, Path, Filename, Result_Directory, Scenario_List, Nodes_List, data, Unit,
    los_colors, Format, Mistar_Solution, Seeding_Period, eval_period, NumRuns,
    Random_Seed, Random_SeedInc, Run, Veh_Type, Time_interval, Output,
    Reporter, Date, Company, Project
):
    print("\n=== Simulation_Nodes (TEST) ===")
    for name, val in [
        ("Vissim_Version", Vissim_Version),
        ("Path", Path),
        ("Filename", Filename),
        ("Result_Directory", Result_Directory),
        ("Scenario_List", Scenario_List),
        ("Nodes_List", Nodes_List),
        ("data", data),
        ("Unit", Unit),
        ("los_colors", los_colors),
        ("Format", Format),
        ("Mistar_Solution", Mistar_Solution),
        ("Seeding_Period", Seeding_Period),
        ("eval_period", eval_period),
        ("NumRuns", NumRuns),
        ("Random_Seed", Random_Seed),
        ("Random_SeedInc", Random_SeedInc),
        ("Run", Run),
        ("Veh_Type", Veh_Type),
        ("Time_interval", Time_interval),
        ("Output", Output),
        ("Reporter", Reporter),
        ("Date", Date),
        ("Company", Company),
        ("Project", Project),
    ]:
        print(f"{name}: {val}   (type={type(val).__name__})")
    print("=== end ===\n")
    return "NODES_TEST_OK"





def Simulation_Links_Link(
    Vissim_Version, Filename, Result_Directory, Unit, Scenario_List, Segments_List,
    Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc,
    Run, Time_interval, Veh_Type, Output
):
    print("\n=== Simulation_Links_Link (TEST) ===")
    print("Segments_List:", Segments_List, type(Segments_List))
    return "LINK_TEST_OK"


def Simulation_Links_Lanes(
    Vissim_Version, Filename, Result_Directory, Unit, Scenario_List, Segments_List,
    Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc,
    Run, Time_interval, Veh_Type, Output
):
    print("\n=== Simulation_Links_Lanes (TEST) ===")
    print("Segments_List:", Segments_List, type(Segments_List))
    return "LANE_TEST_OK"


def Link_Results(
    Results_Type: str,
    Vissim_Version: int,
    Filename: str,
    Result_Directory: str,
    Unit: str,
    Scenario_List: List[Scenario],
    Segments_List: List[SegmentRow],
    Seeding_Period: int,
    eval_period: int,
    NumRuns: int,
    Random_Seed: int,
    Random_SeedInc: int,
    Run: str,
    Time_interval: str,
    Veh_Type: str,
    Output: str,
):
    if Results_Type == "Per Lane":
        res_lane = Simulation_Links_Lanes(
            Vissim_Version, Filename, Result_Directory, Unit, Scenario_List, Segments_List,
            Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc,
            Run, Time_interval, Veh_Type, Output
        )
        res_link = Simulation_Links_Link(
            Vissim_Version, Filename, Result_Directory, Unit, Scenario_List, Segments_List,
            Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc,
            Run, Time_interval, Veh_Type, Output
        )
        return {"Per Lane": res_lane, "Per Link": res_link}

    res_link = Simulation_Links_Link(
        Vissim_Version, Filename, Result_Directory, Unit, Scenario_List, Segments_List,
        Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc,
        Run, Time_interval, Veh_Type, Output
    )
    return {"Per Link": res_link}


def Simulation_TT(
    Vissim_Version,
    Filename,
    Result_Directory,
    Scenario_List,
    TT_List,
    Seeding_Period,
    eval_period,
    NumRuns,
    Random_Seed,
    Random_SeedInc,
    Run,
    Time_interval,
    Units,
    output_file
):
    print("\n=== Simulation_TT (TEST) ===")
    for name, val in [
        ("Vissim_Version", Vissim_Version),
        ("Filename", Filename),
        ("Result_Directory", Result_Directory),
        ("Scenario_List", Scenario_List),
        ("TT_List", TT_List),
        ("Seeding_Period", Seeding_Period),
        ("eval_period", eval_period),
        ("NumRuns", NumRuns),
        ("Random_Seed", Random_Seed),
        ("Random_SeedInc", Random_SeedInc),
        ("Run", Run),
        ("Time_interval", Time_interval),
        ("Units", Units),
        ("output_file", output_file),
    ]:
        print(f"{name}: {val}   (type={type(val).__name__})")
    print("=== end ===\n")
    return "TT_TEST_OK"


def Simulation_Throughput(
    Vissim_Version,
    Filename,
    Result_Directory,
    Scenario_List,
    Nodes_List,
    Seeding_Period,
    eval_period,
    NumRuns,
    Random_Seed,
    Random_SeedInc,
    Run,
    Veh_Type,
    Time_interval,
    Output
):
    print("\n=== Simulation_Throughput (TEST) ===")
    for name, val in [
        ("Vissim_Version", Vissim_Version),
        ("Filename", Filename),
        ("Result_Directory", Result_Directory),
        ("Scenario_List", Scenario_List),
        ("Nodes_List", Nodes_List),
        ("Seeding_Period", Seeding_Period),
        ("eval_period", eval_period),
        ("NumRuns", NumRuns),
        ("Random_Seed", Random_Seed),
        ("Random_SeedInc", Random_SeedInc),
        ("Run", Run),
        ("Veh_Type", Veh_Type),
        ("Time_interval", Time_interval),
        ("Output", Output),
    ]:
        print(f"{name}: {val}   (type={type(val).__name__})")
    print("=== end ===\n")
    return "THROUGHPUT_TEST_OK"


def Network_Results(
    Vissim_Version: int,
    Filename: str,
    Result_Directory: str,
    Scenario_List: List[Scenario],
    Seeding_Period: int,
    eval_period: int,
    NumRuns: int,
    Random_Seed: int,
    Random_SeedInc: int,
    Desired_Metric: List[str],
    Environment_Variables: bool,
    Veh_Type: str,
    Run: str,
    Time_interval: str,
    Output: str,
):
    print("\n=== Network_Results (TEST) ===")
    for name, val in [
        ("Vissim_Version", Vissim_Version),
        ("Filename", Filename),
        ("Result_Directory", Result_Directory),
        ("Scenario_List", Scenario_List),
        ("Seeding_Period", Seeding_Period),
        ("eval_period", eval_period),
        ("NumRuns", NumRuns),
        ("Random_Seed", Random_Seed),
        ("Random_SeedInc", Random_SeedInc),
        ("Desired_Metric", Desired_Metric),
        ("Environment_Variables", Environment_Variables),
        ("Veh_Type", Veh_Type),
        ("Run", Run),
        ("Time_interval", Time_interval),
        ("Output", Output),
    ]:
        print(f"{name}: {val}   (type={type(val).__name__})")
    print("=== end ===\n")
    return "NETWORK_TEST_OK"


def Calibration_Function(
    Vissim_Version, Path, Filename, Result_Directory, Calibrated_variables, Desired_DB, DSD_List_Values,
    distance_List, seeding, Num_Seeding_Intervals, Trials, SubTrial, Random_Slection, Variant, Variant_update, Variant_List, Step_Wise_Variant, Validation_List,
    validation_data, Penalty_Volume, Penalty_Speed, Penalty_TravelTime, Penalty_Queue, Seeding_Period, eval_period, NumRuns, Random_Seed,
    Random_SeedInc
):
    print("\n=== Calibration_Function (TEST) ===")
    for name, val in [
        ("Vissim_Version", Vissim_Version),
        ("Path", Path),
        ("Filename", Filename),
        ("Result_Directory", Result_Directory),
        ("Calibrated_variables", Calibrated_variables),
        ("Desired_DB", Desired_DB),
        ("DSD_List_Values", DSD_List_Values),
        ("distance_List", distance_List),
        ("seeding", seeding),
        ("Num_Seeding_Intervals", Num_Seeding_Intervals),
        ("Trials", Trials),
        ("SubTrial", SubTrial),
        ("Random_Slection", Random_Slection),
        ("Variant", Variant),          # ✅ fixed
        ("Variant_update", Variant_update),
        ("Variant_List", Variant_List),
        ("Step_Wise_Variant", Step_Wise_Variant),
        ("Validation_List", Validation_List),
        ("validation_data", validation_data),
        ("Penalty_Volume", Penalty_Volume),
        ("Penalty_Speed", Penalty_Speed),
        ("Penalty_TravelTime", Penalty_TravelTime),
        ("Penalty_Queue", Penalty_Queue),
        ("Seeding_Period", Seeding_Period),
        ("eval_period", eval_period),
        ("NumRuns", NumRuns),
        ("Random_Seed", Random_Seed),
        ("Random_SeedInc", Random_SeedInc),
    ]:
        print(f"{name}: {val}   (type={type(val).__name__})")
    print("=== end ===\n")
    return "CALIBRATOR_TEST_OK"


# ============================================================
# Shared: Project + Scenario
# ============================================================

class ProjectScenarioMixin:
    def _build_title(self, title_text: str):
        title = QLabel(title_text)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #fdf8e3, stop:0.5 #ffffff, stop:1 #e8f4fd);
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 10px;
                font-size: 18px;
                font-weight: 900;
                color: #143a5d;
            }
        """)
        self.main_layout.addWidget(title)

    def _build_project_setup(self):
        g = QGroupBox("Project Setup")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        add_corner_help(
            g,
            "VISSIM Version: Select the VISSIM version for this project.\n"
            "Project Path: Select the directory to the project.\n"
            "File Name: Enter the VISSIM file name without the extension (.inpx)."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        vissim_label = QLabel("VISSIM Version:")
        vissim_label.setStyleSheet("background: transparent;")
        grid.addWidget(vissim_label, 0, 0)
        self.vissim_version_combo = QComboBox()
        for v in range(21, 27):
            self.vissim_version_combo.addItem(str(v))
        self.vissim_version_combo.setCurrentText("26")
        self.vissim_version_combo.setEditable(True)
        self.vissim_version_combo.lineEdit().setValidator(QIntValidator(0, 999, self))
        self.vissim_version_combo.setFixedWidth(90)
        grid.addWidget(self.vissim_version_combo, 0, 1, Qt.AlignmentFlag.AlignLeft)

        project_label = QLabel("Project File (.inpx):")
        project_label.setStyleSheet("background: transparent;")
        grid.addWidget(project_label, 1, 0)
        self.project_file_edit = QLineEdit()
        self.project_file_edit.setReadOnly(True)
        self.project_file_edit.setFixedWidth(520)
        grid.addWidget(self.project_file_edit, 1, 1)

        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.clicked.connect(self._browse_project_path)
        grid.addWidget(browse, 1, 2)

        self.project_path_edit = QLineEdit()
        self.project_path_edit.setVisible(False)
        self.filename_edit = QLineEdit()
        self.filename_edit.setVisible(False)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_project_path(self):
        f, _ = QFileDialog.getOpenFileName(
            self,
            "Select VISSIM Project File",
            "",
            "VISSIM Project (*.inpx);;All Files (*.*)"
        )
        if f:
            p = Path(f)
            self.project_file_edit.setText(str(p))
            self.project_path_edit.setText(str(p.parent))
            self.filename_edit.setText(p.stem)

    def _build_scenario_setup(self):
        g = QGroupBox("Scenario Setup")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        add_corner_help(
            g,
            "Select the scenarios you would like to report the results for.\n"
            "If you are not using scenario manager, enter 0 as the scenario number."
        )

        top = QHBoxLayout()
        num_label = QLabel("Number of Scenarios:")
        num_label.setStyleSheet("background: transparent;")
        top.addWidget(num_label)
        self.num_scenarios_spin = QSpinBox()
        self.num_scenarios_spin.setRange(1, 999)
        self.num_scenarios_spin.setValue(1)
        self.num_scenarios_spin.setFixedWidth(80)
        _force_input_bg(self.num_scenarios_spin)
        self.num_scenarios_spin.valueChanged.connect(self._sync_scenario_rows_to_count)
        top.addWidget(self.num_scenarios_spin)
        top.addStretch()
        lay.addLayout(top)

        self.scenario_rows: List[dict] = []
        self.scenario_container = QVBoxLayout()
        self.scenario_container.setSpacing(10)
        lay.addLayout(self.scenario_container)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Scenario")
        rem_btn = QPushButton("Remove Last Scenario")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.clicked.connect(self._add_scenario_row_btn)
        rem_btn.clicked.connect(self._remove_scenario_row_btn)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(rem_btn)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        self.main_layout.addWidget(g)

        self._add_scenario_row(default_num=1, default_name="Scenario 1", default_run=False)

    def _add_scenario_row_btn(self):
        self.num_scenarios_spin.setValue(self.num_scenarios_spin.value() + 1)

    def _remove_scenario_row_btn(self):
        if self.num_scenarios_spin.value() > 1:
            self.num_scenarios_spin.setValue(self.num_scenarios_spin.value() - 1)

    def _sync_scenario_rows_to_count(self):
        target = self.num_scenarios_spin.value()
        while len(self.scenario_rows) < target:
            i = len(self.scenario_rows) + 1
            self._add_scenario_row(default_num=i, default_name=f"Scenario {i}", default_run=False)
        while len(self.scenario_rows) > target:
            self._remove_scenario_row()

        self._update_sim_visibility()

    def _add_scenario_row(self, default_num=0, default_name="", default_run=False):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        grid = QGridLayout(roww)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        num_lbl = QLabel("Scenario Number:")
        num_lbl.setAutoFillBackground(False)
        num_lbl.setStyleSheet("background: transparent;")
        grid.addWidget(num_lbl, 0, 0)
        num_spin = QSpinBox()
        num_spin.setRange(0, 999999)
        num_spin.setValue(default_num)
        num_spin.setFixedWidth(90)
        _force_input_bg(num_spin)
        grid.addWidget(num_spin, 0, 1)

        name_lbl = QLabel("Scenario Name:")
        name_lbl.setAutoFillBackground(False)
        name_lbl.setStyleSheet("background: transparent;")
        grid.addWidget(name_lbl, 0, 2)
        name_edit = QLineEdit(default_name)
        name_edit.setFixedWidth(320)
        _force_input_bg(name_edit)
        grid.addWidget(name_edit, 0, 3)

        run_check = BlueCheckBox("Run Model")
        run_check.setChecked(default_run)
        run_check.stateChanged.connect(self._update_sim_visibility)
        grid.addWidget(run_check, 0, 4)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        self.scenario_container.addLayout(wrap)

        _force_input_bg_in_widget(roww)
        self.scenario_rows.append({"widget": roww, "num": num_spin, "name": name_edit, "run": run_check})
        self._update_sim_visibility()

    def _remove_scenario_row(self):
        if not self.scenario_rows:
            return
        r = self.scenario_rows.pop()
        r["widget"].setParent(None)

    def export_project_setup(self) -> dict:
        return {
            "vissim_version": self.vissim_version_combo.currentText(),
            "project_path": self.project_path_edit.text(),
            "filename": self.filename_edit.text(),
        }

    def import_project_setup(self, d: dict):
        _set_combo_text(self.vissim_version_combo, d.get("vissim_version", "26"))
        self.project_path_edit.setText(d.get("project_path", ""))
        self.filename_edit.setText(d.get("filename", ""))
        file_path = _project_file_from_parts(self.project_path_edit.text(), self.filename_edit.text())
        self.project_file_edit.setText(file_path)

    def export_scenarios(self) -> dict:
        return {
            "count": int(self.num_scenarios_spin.value()),
            "rows": [
                {
                    "num": int(r["num"].value()),
                    "name": r["name"].text(),
                    "run": bool(r["run"].isChecked()),
                }
                for r in self.scenario_rows
            ],
        }

    def import_scenarios(self, d: dict):
        rows = d.get("rows", [])
        count = int(d.get("count", len(rows) or 1))

        # This will trigger row creation/removal via your existing sync logic
        self.num_scenarios_spin.setValue(max(1, count))

        # Fill rows
        for i, src in enumerate(rows):
            if i >= len(self.scenario_rows):
                break
            self.scenario_rows[i]["num"].setValue(int(src.get("num", 0)))
            self.scenario_rows[i]["name"].setText(src.get("name", f"Scenario {i+1}"))
            self.scenario_rows[i]["run"].setChecked(bool(src.get("run", False)))

        self._update_sim_visibility()

# ============================================================
# Nodes Tab (UPDATED — includes: veh types fixed + always 1 From/To shown when overpass True)
# ============================================================

class NodesTab(QWidget, ProjectScenarioMixin):
    FEATURE_RESULTS_BASIC = "results_basic"            # Tier 1
    FEATURE_MISTAR = "mistar_solution"                 # Tier 1_M (requires Tier 1)

    # ✅ Always keep one hidden vector pair in the background
    DEFAULT_FROM_TO_PAIR: FromTo = (1, 2)

    def __init__(self, parent=None, license_tab=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        self.license_tab = license_tab
        self.sim_groupbox: Optional[QGroupBox] = None

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        self._build_title("Nodes Results")
        self._build_project_setup()
        self._build_scenario_setup()
        self._build_intersections_block()
        self._build_simulation_parameters()
        self._build_results_specifics()
        self._build_output_metadata()
        self._build_collect()

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

        self.refresh_license_ui()

    # ------------------------------------------------------------
    # Licensing helpers (unchanged)
    # ------------------------------------------------------------
    def _find_license_tab(self):
        if self.license_tab is not None:
            return self.license_tab

        p = self.parent()
        while p is not None:
            if hasattr(p, "license_tab"):
                lt = getattr(p, "license_tab")
                if lt is not None:
                    self.license_tab = lt
                    return lt
            p = p.parent()
        return None

    def _find_tab_widget(self):
        p = self.parent()
        while p is not None:
            if isinstance(p, QTabWidget):
                return p
            p = p.parent()
        return None

    def _go_to_license_tab(self, message: str):
        lt = self._find_license_tab()
        tabs = self._find_tab_widget()

        if tabs is not None and lt is not None:
            idx = tabs.indexOf(lt)
            if idx >= 0:
                tabs.setCurrentIndex(idx)

        if lt is not None:
            if hasattr(lt, "show_status"):
                lt.show_status(message)
            if hasattr(lt, "focus_key_input"):
                lt.focus_key_input()

    def _license_snapshot(self, force: bool = False) -> dict:
        lt = self._find_license_tab()
        if lt is None:
            return {
                "app_enabled": True,
                "valid": False,
                "reason": "UNACTIVATED",
                "features": [],
            }

        if hasattr(lt, "validate_and_cache"):
            try:
                return lt.validate_and_cache(force=force)  # type: ignore
            except Exception:
                return {
                    "app_enabled": True,
                    "valid": False,
                    "reason": "NO_INTERNET",
                    "features": [],
                }

        state = {
            "app_enabled": True,
            "valid": False,
            "reason": "UNACTIVATED",
            "features": [],
        }
        if hasattr(lt, "is_app_enabled"):
            try:
                state["app_enabled"] = bool(lt.is_app_enabled())  # type: ignore
            except Exception:
                pass
        if hasattr(lt, "is_license_valid"):
            try:
                state["valid"] = bool(lt.is_license_valid())  # type: ignore
            except Exception:
                pass
        return state

    def _has_feature(self, feature_name: str) -> bool:
        lt = self._find_license_tab()
        if lt is None:
            return False
        if hasattr(lt, "has_feature"):
            try:
                return bool(lt.has_feature(feature_name))  # type: ignore
            except Exception:
                return False
        return False

    def _ensure_feature_or_redirect(self, feature_name: str, friendly_action: str) -> bool:
        snap = self._license_snapshot(force=True)

        if snap.get("app_enabled", True) is False:
            self._go_to_license_tab("Software is currently disabled by the publisher (shutdown switch is ON).")
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
            return False

        if not snap.get("valid", False):
            reason = str(snap.get("reason", "UNACTIVATED"))
            if reason == "NO_INTERNET":
                msg = "Internet is required to validate your license. Please connect and try again."
            elif reason == "EXPIRED":
                msg = "Your license has expired. Please renew to continue."
            elif reason == "SEAT_LIMIT":
                msg = "No seats available. Ask your company admin to deactivate a seat."
            else:
                msg = "Activation required. Please enter your license key to continue."

            self._go_to_license_tab(f"{friendly_action} requires activation.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        features = snap.get("features", [])
        has = (feature_name in features) if isinstance(features, list) else self._has_feature(feature_name)
        if not has:
            if feature_name == self.FEATURE_RESULTS_BASIC:
                msg = "Tier 1 is required to use Basic Results Reporter actions (Collect Results)."
            elif feature_name == self.FEATURE_MISTAR:
                msg = "Tier 1_M (Mistar Solution add-on) is required. You must have Tier 1 first."
            else:
                msg = "Your license does not include this feature."

            self._go_to_license_tab(f"{friendly_action} is locked.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        return True

    # ------------------------------------------------------------
    # UI build
    # ------------------------------------------------------------

    def _build_intersections_block(self):
        g = QGroupBox("Intersections")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        add_corner_help(
            g,
            "Define your node numbers and their types.\n"
            "If there is an overpass/underpass on a different grade, enable it and define From/To links to exclude.\n"
            "Define LOS thresholds, colors, and formatting."
        )

        title1 = QLabel("Nodes List")
        title1.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title1)

        self.node_rows: List[dict] = []
        self.nodes_container = QVBoxLayout()
        self.nodes_container.setSpacing(10)
        lay.addLayout(self.nodes_container)

        btn_row = QHBoxLayout()
        add_node = QPushButton("Add Node")
        rem_node = QPushButton("Remove Last Node")
        add_node.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_node.setStyleSheet(SOFT_BUTTON_STYLE)
        add_node.clicked.connect(self._add_node_row_btn)
        rem_node.clicked.connect(self._remove_node_row)
        btn_row.addWidget(add_node)
        btn_row.addWidget(rem_node)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        lay.addWidget(hline())

        title2 = QLabel("Level of Service (LOS)")
        title2.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title2)

        self.los_rows: List[dict] = []
        los_grid = QGridLayout()
        los_grid.setHorizontalSpacing(12)
        los_grid.setVerticalSpacing(8)

        los_grid.setColumnMinimumWidth(0, 35)
        los_grid.setColumnStretch(0, 0)

        los_lbl = QLabel("LOS")
        los_lbl.setFixedWidth(35)
        los_grid.addWidget(los_lbl, 0, 0)

        los_grid.addWidget(QLabel("Signalized Low"), 0, 1)
        los_grid.addWidget(QLabel("Signalized High"), 0, 2)
        los_grid.addWidget(QLabel("Unsignalized Low"), 0, 3)
        los_grid.addWidget(QLabel("Unsignalized High"), 0, 4)

        letters = ["A", "B", "C", "D", "E", "F"]
        s_defaults = [("0", "10"), ("10", "20"), ("20", "30"), ("30", "40"), ("40", "50"), ("50", "999999")]
        u_defaults = [("0", "10"), ("10", "20"), ("20", "30"), ("30", "40"), ("40", "50"), ("50", "999999")]

        for i, L in enumerate(letters, start=1):
            Llbl = QLabel(L)
            Llbl.setFixedWidth(35)
            los_grid.addWidget(Llbl, i, 0)

            s_lo = QLineEdit(s_defaults[i-1][0]); s_lo.setFixedWidth(110)
            s_hi = QLineEdit(s_defaults[i-1][1]); s_hi.setFixedWidth(110)
            u_lo = QLineEdit(u_defaults[i-1][0]); u_lo.setFixedWidth(110)
            u_hi = QLineEdit(u_defaults[i-1][1]); u_hi.setFixedWidth(110)

            los_grid.addWidget(s_lo, i, 1)
            los_grid.addWidget(s_hi, i, 2)
            los_grid.addWidget(u_lo, i, 3)
            los_grid.addWidget(u_hi, i, 4)

            self.los_rows.append({
                "letter": L,
                "signalized_low": s_lo,
                "signalized_high": s_hi,
                "unsignalized_low": u_lo,
                "unsignalized_high": u_hi,
            })

        los_wrap = QHBoxLayout()
        los_wrap.addLayout(los_grid)
        los_wrap.addStretch()
        lay.addLayout(los_wrap)

        lay.addWidget(hline())

        title3 = QLabel("LOS Colors")
        title3.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title3)

        self.los_color_widgets: Dict[str, QComboBox] = {}
        color_grid = QGridLayout()
        color_grid.setHorizontalSpacing(12)
        color_grid.setVerticalSpacing(8)

        color_options = ["No Color", "Dark Green", "Green", "Yellow", "Orange", "Red"]
        default_colors = {"A": "No Color", "B": "Dark Green", "C": "Green", "D": "Yellow", "E": "Orange", "F": "Red"}

        r = 0
        for L in letters:
            color_grid.addWidget(QLabel(f"LOS {L}:"), r, 0)
            cb = QComboBox()
            cb.addItems(color_options)
            cb.setCurrentText(default_colors[L])
            cb.setFixedWidth(170)
            color_grid.addWidget(cb, r, 1)
            self.los_color_widgets[L] = cb
            r += 1

        color_wrap = QHBoxLayout()
        color_wrap.addLayout(color_grid)
        color_wrap.addStretch()
        lay.addLayout(color_wrap)

        lay.addWidget(hline())

        title4 = QLabel("Results Formatting")
        title4.setStyleSheet(SUBTITLE_STYLE)
        title4.setToolTip(
            "Dummy help:\n"
            "Use the sections below to define the formatting tokens used in your reported results.\n"
            "Each scope (Overall/Approach/Movement) controls how values are labeled in the output."
        )
        lay.addWidget(title4)

        self.format_sections: Dict[str, Dict[str, Any]] = {}
        for scope in FORMAT_SCOPES:
            scope_box = self._create_format_scope(scope)
            lay.addWidget(scope_box)

        lay.addWidget(hline())

        title5 = QLabel("Units")
        title5.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title5)

        unit_row = QHBoxLayout()
        unit_row.addWidget(QLabel("Unit:"))
        self.unit_combo = QComboBox()
        self.unit_combo.addItems(["meter", "ft"])
        self.unit_combo.setCurrentText("meter")
        self.unit_combo.setFixedWidth(120)
        unit_row.addWidget(self.unit_combo)

        unit_row.addSpacing(20)

        self.mistar_check = BlueCheckBox("Enable Mistar Solution")
        self.mistar_check.setToolTip(
            "Mistar Solution helps report generic names for approaches and movements. "
            "It applies Gen AI andd mathematics to identify the approach names and their movements\n\n"
        )
        self.mistar_check.setAutoFillBackground(False)
        self.mistar_check.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, False)
        self.mistar_check.setStyleSheet(
            (self.mistar_check.styleSheet() or "") +
            """
            QCheckBox {
                background: #ffffff;
                color: #000000;
                font-weight: 700;
                border: 1px solid rgba(20, 40, 80, 0.22);
                border-radius: 10px;
                padding: 6px 10px;
                margin: 0px;
            }
            """
        )

        unit_row.addWidget(self.mistar_check)
        unit_row.addStretch()
        lay.addLayout(unit_row)

        self.main_layout.addWidget(g)

        # ✅ Default node: 1 hidden pair exists; UI shows it only when Overpass is checked.
        self._add_node_row(
            node_number=1,
            control="Signalized",
            intersection_type="Regular",
            overpass=False,
            from_to_pairs=[self.DEFAULT_FROM_TO_PAIR],
        )

    def _create_format_scope(self, scope_name: str) -> QGroupBox:
        box = QGroupBox(scope_name)
        box.setStyleSheet(SECTION_GROUPBOX_STYLE)

        dummy_scope_help = {
            "Overall Intersection": (
                "You can customize the <b>Overall Intersection</b> MoEs and their wrappers. Select the desired MoE to be reported in the final results and their format."
            ),
            "Approach": (
                "You can customize the <b>Approach</b> MoEs and their wrappers. Select the desired MoE to be reported in the final results and their format."
            ),
            "Movement": (
                "You can customize the <b>Movement</b> MoEs and their wrappers. Select the desired MoE to be reported in the final results and their format."
            ),
        }
        if scope_name in dummy_scope_help:
            add_corner_help(box, dummy_scope_help[scope_name])

        l = QVBoxLayout(box)

        top = QHBoxLayout()
        preview = QLabel("Preview: LOS")
        preview.setStyleSheet("QLabel { font-weight: 800; }")

        add_btn = QPushButton("+")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.setFixedWidth(46)

        top.addWidget(preview)
        top.addSpacing(10)
        top.addWidget(add_btn)
        top.addStretch()
        l.addLayout(top)

        container = QVBoxLayout()
        container.setSpacing(8)
        l.addLayout(container)

        self.format_sections[scope_name] = {"preview": preview, "container": container, "segments": [], "add_btn": add_btn}
        add_btn.clicked.connect(lambda: self._add_format_segment(scope_name))
        self._refresh_format_add_button_state(scope_name)
        return box

    def _current_used_format_vars(self, scope_name: str, exclude_seg=None):
        used = set()
        for seg in self.format_sections[scope_name]["segments"]:
            if seg is exclude_seg:
                continue
            used.add(seg["var_combo"].currentText())
        return used

    def _add_format_segment(self, scope_name: str):
        used = self._current_used_format_vars(scope_name)
        available = [v for v in FORMAT_VAR_DISPLAY if v not in used]
        if not available:
            return

        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        grid = QGridLayout(roww)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        grid.addWidget(QLabel("Variable:"), 0, 0)
        var_combo = QComboBox()
        var_combo.addItems(available)
        var_combo.setCurrentText(available[0])
        var_combo.setFixedWidth(190)
        grid.addWidget(var_combo, 0, 1)

        grid.addWidget(QLabel("Wrapper:"), 0, 2)
        wrap_combo = QComboBox()
        wrap_combo.addItems(WRAPPERS)
        wrap_combo.setCurrentText("()")
        wrap_combo.setFixedWidth(120)
        grid.addWidget(wrap_combo, 0, 3)

        remove_btn = QPushButton("x")
        remove_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        remove_btn.setFixedWidth(46)
        grid.addWidget(remove_btn, 0, 4)

        _force_input_bg_in_widget(roww)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()

        self.format_sections[scope_name]["container"].addLayout(wrap)

        seg = {"widget": roww, "var_combo": var_combo, "wrap_combo": wrap_combo}
        self.format_sections[scope_name]["segments"].append(seg)

        var_combo.currentTextChanged.connect(lambda _t, s=seg, scope=scope_name: self._on_format_var_changed(scope, s))
        wrap_combo.currentTextChanged.connect(lambda _t, scope=scope_name: self._update_format_preview(scope))
        remove_btn.clicked.connect(lambda: self._remove_format_segment(scope_name, seg))

        self._update_format_preview(scope_name)
        self._refresh_format_add_button_state(scope_name)

    def _remove_format_segment(self, scope_name: str, seg):
        segments = self.format_sections[scope_name]["segments"]
        if seg in segments:
            segments.remove(seg)
            seg["widget"].setParent(None)
        self._update_format_preview(scope_name)
        self._refresh_format_add_button_state(scope_name)

    def _on_format_var_changed(self, scope_name: str, seg):
        used = self._current_used_format_vars(scope_name, exclude_seg=seg)
        cur = seg["var_combo"].currentText()
        if cur in used:
            available = [v for v in FORMAT_VAR_DISPLAY if v not in used]
            if available:
                seg["var_combo"].setCurrentText(available[0])
        self._update_format_preview(scope_name)
        self._refresh_format_add_button_state(scope_name)

    def _refresh_format_add_button_state(self, scope_name: str):
        used = self._current_used_format_vars(scope_name)
        available = [v for v in FORMAT_VAR_DISPLAY if v not in used]
        self.format_sections[scope_name]["add_btn"].setDisabled(len(available) == 0)

    def _update_format_preview(self, scope_name: str):
        segs = self.format_sections[scope_name]["segments"]
        tokens = ["LOS"]
        for seg in segs:
            var_display = seg["var_combo"].currentText()
            var_key = FORMAT_VAR_DISPLAY_TO_KEY.get(var_display, var_display)
            wrap = seg["wrap_combo"].currentText()

            if wrap == "":
                tokens.append(var_key)
            elif wrap == "[]":
                tokens.extend(["[", var_key, "]"])
            elif wrap == "{}":
                tokens.extend(["{", var_key, "}"])
            elif wrap == "{{}}":
                tokens.extend(["{{", var_key, "}}"])
            elif wrap == "[[]]":
                tokens.extend(["[[", var_key, "]]"])
            elif wrap == "()":
                tokens.extend(["(", var_key, ")"])
            elif wrap == "(())":
                tokens.extend(["((", var_key, "))"])
            elif wrap == "{| |}":
                tokens.extend(["{|", var_key, "|}"])
            else:
                tokens.append(var_key)

        self.format_sections[scope_name]["preview"].setText("Preview: " + " ".join(tokens))

    def _add_node_row_btn(self):
        self._add_node_row(
            node_number=0,
            control="Signalized",
            intersection_type="Regular",
            overpass=False,
            from_to_pairs=[self.DEFAULT_FROM_TO_PAIR],  # ✅ only one
        )

    def _add_node_row(
        self,
        node_number: int,
        control: str,
        intersection_type: str,
        overpass: bool,
        from_to_pairs: List[Tuple[int, int]],
    ):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        v = QVBoxLayout(roww)
        v.setSpacing(8)

        top = QGridLayout()
        top.setHorizontalSpacing(10)
        top.setVerticalSpacing(6)

        node_num_lbl = QLabel("Node Number:")
        node_num_lbl.setAutoFillBackground(False)
        node_num_lbl.setStyleSheet("background: transparent;")
        top.addWidget(node_num_lbl, 0, 0)
        num_spin = QSpinBox()
        num_spin.setRange(0, 999999999)
        num_spin.setValue(int(node_number))
        num_spin.setFixedWidth(110)
        _force_input_bg(num_spin)
        top.addWidget(num_spin, 0, 1)

        top.addWidget(QLabel("Intersection Control:"), 0, 2)
        control_combo = QComboBox()
        control_combo.addItems(["Signalized", "Unsignalized"])
        control_combo.setCurrentText(control if control in ["Signalized", "Unsignalized"] else "Signalized")
        control_combo.setFixedWidth(170)
        _force_input_bg(control_combo)
        top.addWidget(control_combo, 0, 3)

        top.addWidget(QLabel("Intersection Type:"), 0, 4)
        type_edit = QLineEdit(intersection_type)
        type_edit.setFixedWidth(220)
        _force_input_bg(type_edit)
        top.addWidget(type_edit, 0, 5)

        over = BlueCheckBox("Overpass / Underpass")
        over.setChecked(bool(overpass))
        top.addWidget(over, 0, 6)

        top_wrap = QHBoxLayout()
        top_wrap.addLayout(top)
        top_wrap.addStretch()
        v.addLayout(top_wrap)

        ft_container = QVBoxLayout()
        ft_container.setSpacing(6)
        v.addLayout(ft_container)

        # ✅ Requirement: only one From/To row should ever appear.
        # We remove Add/Remove buttons entirely, but keep a hidden widget key for compatibility.
        ft_btns_widget = QWidget()
        ft_btns_widget.setVisible(False)
        ft_btns_widget.setEnabled(False)
        v.addWidget(ft_btns_widget)

        node_data = {
            "widget": roww,
            "num": num_spin,
            "control": control_combo,
            "type": type_edit,
            "overpass": over,
            "from_to_container": ft_container,
            "from_to_rows": [],
            "from_edits": [],
            "to_edits": [],
            "ft_btns_widget": ft_btns_widget,
        }

        over.stateChanged.connect(lambda _s, nd=node_data: self._update_from_to_visibility(nd))

        # ✅ Force exactly ONE pair in the background
        pair = from_to_pairs[0] if (from_to_pairs and len(from_to_pairs) > 0) else self.DEFAULT_FROM_TO_PAIR
        self._add_from_to_row(node_data, pair[0], pair[1])

        self._update_from_to_visibility(node_data)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        self.nodes_container.addLayout(wrap)

        _force_input_bg_in_widget(roww)
        self.node_rows.append(node_data)

    def _add_from_to_row(self, node_row, from_val=None, to_val=None):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        g = QGridLayout(roww)
        g.setHorizontalSpacing(10)
        g.setVerticalSpacing(4)

        g.addWidget(QLabel("From:"), 0, 0)
        f = DigitsLineEdit("" if from_val is None else str(from_val))
        f.setFixedWidth(160)
        g.addWidget(f, 0, 1)

        g.addWidget(QLabel("To:"), 0, 2)
        t = DigitsLineEdit("" if to_val is None else str(to_val))
        t.setFixedWidth(160)
        g.addWidget(t, 0, 3)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()

        node_row["from_to_container"].addLayout(wrap)
        _force_input_bg_in_widget(roww)
        node_row["from_to_rows"].append(roww)
        node_row["from_edits"].append(f)
        node_row["to_edits"].append(t)

        self._update_from_to_visibility(node_row)

    def _update_from_to_visibility(self, node_row):
        enabled = node_row["overpass"].isChecked()

        # Keep the (now unused) buttons widget hidden permanently
        if "ft_btns_widget" in node_row and node_row["ft_btns_widget"] is not None:
            node_row["ft_btns_widget"].setVisible(False)
            node_row["ft_btns_widget"].setEnabled(False)

        # Ensure exactly ONE row exists in the background
        if len(node_row.get("from_to_rows", [])) == 0:
            self._add_from_to_row(node_row, self.DEFAULT_FROM_TO_PAIR[0], self.DEFAULT_FROM_TO_PAIR[1])

        # If older saved layouts created extras, trim them
        while len(node_row["from_to_rows"]) > 1:
            extra = node_row["from_to_rows"].pop()
            node_row["from_edits"].pop()
            node_row["to_edits"].pop()
            extra.setParent(None)

        # Show/hide the single row
        rw = node_row["from_to_rows"][0]
        rw.setVisible(enabled)
        rw.setEnabled(enabled)

    def _remove_node_row(self):
        if not self.node_rows:
            return
        r = self.node_rows.pop()
        r["widget"].setParent(None)

    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.sim_groupbox = g
        lay = QVBoxLayout(g)

        add_corner_help(
            g,
            "Only required if you are running the model.\n"
            "Evaluation period must be greater than seeding period."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_sec = QLineEdit("1800")
        self.eval_sec = QLineEdit("3600")
        self.seed_sec.setFixedWidth(140)
        self.eval_sec.setFixedWidth(140)
        self.seed_sec.setValidator(QIntValidator(0, 999999999, self))
        self.eval_sec.setValidator(QIntValidator(0, 999999999, self))

        self.num_runs = QSpinBox()
        self.num_runs.setRange(1, 9999)
        self.num_runs.setValue(2)

        self.rand_seed = QSpinBox()
        self.rand_seed.setRange(1, 999999999)
        self.rand_seed.setValue(42)

        self.rand_seed_inc = QSpinBox()
        self.rand_seed_inc.setRange(1, 999999999)
        self.rand_seed_inc.setValue(10)

        grid.addWidget(QLabel("Seeding Period (sec):"), 0, 0)
        grid.addWidget(self.seed_sec, 0, 1)
        grid.addWidget(QLabel("Evaluation Period (sec):"), 0, 2)
        grid.addWidget(self.eval_sec, 0, 3)

        grid.addWidget(QLabel("Number of Runs:"), 1, 0)
        grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(QLabel("Random Seed:"), 1, 2)
        grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(QLabel("Random Seed Increment:"), 1, 4)
        grid.addWidget(self.rand_seed_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        self._update_sim_visibility()

    def _update_sim_visibility(self):
        if self.sim_groupbox is None:
            return
        show = any(r["run"].isChecked() for r in self.scenario_rows)
        self.sim_groupbox.setVisible(show)

    def _build_results_specifics(self):
        g = QGroupBox("Results Specifics")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        add_corner_help(
            g,
            "Run: select Avg/Min/Max/StdDev/Var or run number.\n"
            "Vehicle Type: select the vehicle class.\n"
            "Time Interval: Avg or interval number."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.run_combo = QComboBox()
        run_options = ["Avg", "Min", "Max", "StdDev", "Var"] + [str(i) for i in range(1, 11)] + ["Other"]
        self.run_combo.addItems(run_options)
        self.run_combo.setCurrentText("Avg")
        self.run_combo.setFixedWidth(140)

        self.run_other = QLineEdit()
        self.run_other.setPlaceholderText('If "Other", enter here')
        self.run_other.setFixedWidth(200)
        self.run_other.setEnabled(False)
        self.run_combo.currentTextChanged.connect(lambda t: self.run_other.setEnabled(t == "Other"))

        # ✅ Veh types fixed to your required backend strings
        self.veh_combo = QComboBox()
        self.veh_combo.addItems([
            "All",
            "Cars",
            "HGVs",
            "Buses",
            "Motorcycles",
            "Bikes",
            "Pedestrians",
        ])
        self.veh_combo.setCurrentText("All")
        self.veh_combo.setFixedWidth(210)

        self.ti_combo = QComboBox()
        self.ti_combo.addItems(["Avg"] + [str(i) for i in range(1, 11)] + ["Other"])
        self.ti_combo.setCurrentText("1")
        self.ti_combo.setFixedWidth(140)

        self.ti_other = QLineEdit()
        self.ti_other.setPlaceholderText('If "Other", enter here')
        self.ti_other.setFixedWidth(200)
        self.ti_other.setEnabled(False)
        self.ti_combo.currentTextChanged.connect(lambda t: self.ti_other.setEnabled(t == "Other"))

        grid.addWidget(QLabel("Run:"), 0, 0)
        grid.addWidget(self.run_combo, 0, 1)
        grid.addWidget(self.run_other, 0, 2)

        grid.addWidget(QLabel("Vehicle Type:"), 1, 0)
        grid.addWidget(self.veh_combo, 1, 1)

        grid.addWidget(QLabel("Time Interval:"), 2, 0)
        grid.addWidget(self.ti_combo, 2, 1)
        grid.addWidget(self.ti_other, 2, 2)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _build_output_metadata(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        add_corner_help(g, "Select where to save the output Excel file and enter report metadata.")

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.clicked.connect(self._browse_result_dir)

        self.output_file = QLineEdit("Nodes Results.xlsx")
        self.output_file.setFixedWidth(260)

        self.reporter = QLineEdit("User 1")
        self.reporter.setFixedWidth(220)

        self.date_edit = QDateEdit()
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDate(QDate.currentDate())
        self.date_edit.setFixedWidth(160)

        self.company = QLineEdit("Company A")
        self.company.setFixedWidth(220)

        self.project = QLineEdit("Project A")
        self.project.setFixedWidth(220)

        grid.addWidget(QLabel("Results Directory:"), 0, 0)
        grid.addWidget(self.result_dir, 0, 1)
        grid.addWidget(browse, 0, 2)

        grid.addWidget(QLabel("Output File Name:"), 1, 0)
        grid.addWidget(self.output_file, 1, 1)

        grid.addWidget(QLabel("Reporter:"), 2, 0)
        grid.addWidget(self.reporter, 2, 1)

        grid.addWidget(QLabel("Date:"), 2, 2)
        grid.addWidget(self.date_edit, 2, 3)

        grid.addWidget(QLabel("Company:"), 3, 0)
        grid.addWidget(self.company, 3, 1)

        grid.addWidget(QLabel("Project:"), 3, 2)
        grid.addWidget(self.project, 3, 3)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    def _build_collect(self):
        row = QHBoxLayout()
        self.collect_btn = QPushButton("Collect Nodes Results")
        self.collect_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.collect_btn.clicked.connect(self._on_collect)
        row.addStretch()
        row.addWidget(self.collect_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.main_layout.addWidget(self.log)

    def refresh_license_ui(self):
        """
        - Mistar Solution is an add-on (Tier 1_M) and requires Tier 1.
        - So we disable the checkbox unless both are present.
        """
        has_tier1 = self._has_feature(self.FEATURE_RESULTS_BASIC)
        has_mistar = self._has_feature(self.FEATURE_MISTAR)

        allowed = bool(has_tier1 and has_mistar)

        self.mistar_check.setEnabled(allowed)
        if not allowed:
            self.mistar_check.setChecked(False)

    def _on_collect(self):
        if not self._ensure_feature_or_redirect(self.FEATURE_RESULTS_BASIC, "Collect Nodes Results"):
            return

        if bool(self.mistar_check.isChecked()):
            if not self._ensure_feature_or_redirect(self.FEATURE_MISTAR, "Mistar Solution"):
                return

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Scenario_List: List[Scenario] = []
        for r in self.scenario_rows:
            Scenario_List.append((
                int(r["num"].value()),
                r["name"].text().strip(),
                bool(r["run"].isChecked()),
            ))

        Nodes_List: List[NodeRowTuple] = []
        for node in self.node_rows:
            node_number = int(node["num"].value())
            control = node["control"].currentText().strip()
            intersection_type = node["type"].text().strip()
            overpass = bool(node["overpass"].isChecked())

            # ✅ Always send a 1-element vector in the background (even if hidden)
            if node["from_edits"] and node["to_edits"]:
                pair: FromTo = (node["from_edits"][0].as_int(), node["to_edits"][0].as_int())
            else:
                pair = self.DEFAULT_FROM_TO_PAIR

            Nodes_List.append((node_number, control, intersection_type, overpass, [pair]))

        data = {"LOS": [], "Signalized": [], "Unsignalized": []}
        for row in self.los_rows:
            L = row["letter"]
            s_lo = row["signalized_low"].text().strip()
            s_hi = row["signalized_high"].text().strip()
            u_lo = row["unsignalized_low"].text().strip()
            u_hi = row["unsignalized_high"].text().strip()
            data["LOS"].append(L)
            data["Signalized"].append(f"{s_lo}-{s_hi}")
            data["Unsignalized"].append(f"{u_lo}-{u_hi}")

        los_colors = {L: cb.currentText() for L, cb in self.los_color_widgets.items()}

        Format: Dict[str, List[str]] = {}
        for scope in FORMAT_SCOPES:
            segs = self.format_sections[scope]["segments"]
            tokens = ["LOS"]
            for seg in segs:
                var_display = seg["var_combo"].currentText()
                var_key = FORMAT_VAR_DISPLAY_TO_KEY.get(var_display, var_display)
                wrap = seg["wrap_combo"].currentText()

                if wrap == "":
                    tokens.append(var_key)
                elif wrap == "[]":
                    tokens.extend(["[", var_key, "]"])
                elif wrap == "{}":
                    tokens.extend(["{", var_key, "}"])
                elif wrap == "{{}}":
                    tokens.extend(["{{", var_key, "}}"])
                elif wrap == "[[]]":
                    tokens.extend(["[[", var_key, "]]"])
                elif wrap == "()":
                    tokens.extend(["(", var_key, ")"])
                elif wrap == "(())":
                    tokens.extend(["((", var_key, "))"])
                elif wrap == "{| |}":
                    tokens.extend(["{|", var_key, "|}"])
                else:
                    tokens.append(var_key)

            Format[scope] = tokens

        Unit = self.unit_combo.currentText()
        Mistar_Solution = bool(self.mistar_check.isChecked())

        Seeding_Period = int(self.seed_sec.text().strip() or "0")
        eval_period = int(self.eval_sec.text().strip() or "0")
        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_seed_inc.value())

        Run = self.run_other.text().strip() if self.run_combo.currentText() == "Other" else self.run_combo.currentText()
        Veh_Type = self.veh_combo.currentText()
        Time_interval = self.ti_other.text().strip() if self.ti_combo.currentText() == "Other" else self.ti_combo.currentText()

        Result_Directory = self.result_dir.text().strip()
        Output = self.output_file.text().strip()
        Reporter = self.reporter.text().strip()

        qd = self.date_edit.date()
        Date = f"{qd.month():02d}/{qd.day():02d}/{qd.year()}"
        Company = self.company.text().strip()
        Project = self.project.text().strip()

        if Path_ and not os.path.isdir(Path_):
            QMessageBox.warning(self, "Project Path", "Project Path must be a directory.")
            return
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return
        if any(s[2] for s in Scenario_List):
            if eval_period <= Seeding_Period:
                QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
                return

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        log_line("VissiCaRe Nodes Results inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Result_Directory", Result_Directory),
            ("Scenario_List", Scenario_List),
            ("Nodes_List", Nodes_List),
            ("data", data),
            ("Unit", Unit),
            ("los_colors", los_colors),
            ("Format", Format),
            ("Mistar_Solution", Mistar_Solution),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
            ("Run", Run),
            ("Veh_Type", Veh_Type),
            ("Time_interval", Time_interval),
            ("Output", Output),
            ("Reporter", Reporter),
            ("Date", Date),
            ("Company", Company),
            ("Project", Project),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Simulation_Nodes (TEST)...\n")
        res = Simulation_Nodes(
            Vissim_Version=Vissim_Version,
            Path=Path_,
            Filename=Filename,
            Result_Directory=Result_Directory,
            Scenario_List=Scenario_List,
            Nodes_List=Nodes_List,
            data=data,
            Unit=Unit,
            los_colors=los_colors,
            Format=Format,
            Mistar_Solution=Mistar_Solution,
            Seeding_Period=Seeding_Period,
            eval_period=eval_period,
            NumRuns=NumRuns,
            Random_Seed=Random_Seed,
            Random_SeedInc=Random_SeedInc,
            Run=Run,
            Veh_Type=Veh_Type,
            Time_interval=Time_interval,
            Output=Output,
            Reporter=Reporter,
            Date=Date,
            Company=Company,
            Project=Project,
        )
        log_line(f"\nSimulation_Nodes finished. Return: {res}")

    def export_state(self) -> dict:
        nodes = []
        for n in self.node_rows:
            # ✅ Always persist exactly one pair (background vector)
            if n["from_edits"] and n["to_edits"]:
                pair = [n["from_edits"][0].as_int(), n["to_edits"][0].as_int()]
            else:
                pair = [self.DEFAULT_FROM_TO_PAIR[0], self.DEFAULT_FROM_TO_PAIR[1]]

            nodes.append({
                "node_number": int(n["num"].value()),
                "control": n["control"].currentText(),
                "intersection_type": n["type"].text(),
                "overpass": bool(n["overpass"].isChecked()),
                "from_to_pairs": [pair],
            })

        los = []
        for r in self.los_rows:
            los.append({
                "letter": r["letter"],
                "signalized_low": r["signalized_low"].text(),
                "signalized_high": r["signalized_high"].text(),
                "unsignalized_low": r["unsignalized_low"].text(),
                "unsignalized_high": r["unsignalized_high"].text(),
            })

        los_colors = {L: cb.currentText() for L, cb in self.los_color_widgets.items()}

        fmt = {}
        for scope in FORMAT_SCOPES:
            segs = []
            for seg in self.format_sections[scope]["segments"]:
                segs.append({
                    "var": seg["var_combo"].currentText(),
                    "wrap": seg["wrap_combo"].currentText(),
                })
            fmt[scope] = segs

        return {
            "project": self.export_project_setup(),
            "scenarios": self.export_scenarios(),
            "nodes": nodes,
            "los": los,
            "los_colors": los_colors,
            "format": fmt,
            "unit": self.unit_combo.currentText(),
            "mistar": bool(self.mistar_check.isChecked()),

            "sim": {
                "seed_sec": self.seed_sec.text(),
                "eval_sec": self.eval_sec.text(),
                "num_runs": int(self.num_runs.value()),
                "rand_seed": int(self.rand_seed.value()),
                "rand_seed_inc": int(self.rand_seed_inc.value()),
            },
            "results": {
                "run_combo": self.run_combo.currentText(),
                "run_other": self.run_other.text(),
                "veh_type": self.veh_combo.currentText(),
                "ti_combo": self.ti_combo.currentText(),
                "ti_other": self.ti_other.text(),
            },
            "output": {
                "result_dir": self.result_dir.text(),
                "output_file": self.output_file.text(),
                "reporter": self.reporter.text(),
                "date": _qdate_to_str(self.date_edit.date()),
                "company": self.company.text(),
                "project": self.project.text(),
            },
        }

    def import_state(self, s: dict):
        self.import_project_setup(s.get("project", {}))
        self.import_scenarios(s.get("scenarios", {}))

        for nd in self.node_rows:
            nd["widget"].setParent(None)
        self.node_rows = []

        for nd in s.get("nodes", []):
            pairs_raw = nd.get("from_to_pairs", []) or []
            # ✅ Use only the first pair even if older saves had multiple
            if pairs_raw:
                a, b = pairs_raw[0]
                pair = (int(a), int(b))
            else:
                pair = self.DEFAULT_FROM_TO_PAIR

            self._add_node_row(
                node_number=int(nd.get("node_number", 0)),
                control=nd.get("control", "Signalized"),
                intersection_type=nd.get("intersection_type", ""),
                overpass=bool(nd.get("overpass", False)),
                from_to_pairs=[pair],
            )

        los_rows = s.get("los", [])
        if los_rows and len(los_rows) == len(self.los_rows):
            for i, r in enumerate(los_rows):
                self.los_rows[i]["signalized_low"].setText(str(r.get("signalized_low", "")))
                self.los_rows[i]["signalized_high"].setText(str(r.get("signalized_high", "")))
                self.los_rows[i]["unsignalized_low"].setText(str(r.get("unsignalized_low", "")))
                self.los_rows[i]["unsignalized_high"].setText(str(r.get("unsignalized_high", "")))

        for L, val in (s.get("los_colors", {}) or {}).items():
            if L in self.los_color_widgets:
                _set_combo_text(self.los_color_widgets[L], val)

        for scope in FORMAT_SCOPES:
            for seg in list(self.format_sections[scope]["segments"]):
                self._remove_format_segment(scope, seg)

            for seg in s.get("format", {}).get(scope, []):
                self._add_format_segment(scope)
                new_seg = self.format_sections[scope]["segments"][-1]
                _set_combo_text(new_seg["var_combo"], seg.get("var", "Delay"))
                _set_combo_text(new_seg["wrap_combo"], seg.get("wrap", "()"))
            self._update_format_preview(scope)
            self._refresh_format_add_button_state(scope)

        _set_combo_text(self.unit_combo, s.get("unit", "meter"))
        self.mistar_check.setChecked(bool(s.get("mistar", False)))
        self.refresh_license_ui()

        sim = s.get("sim", {})
        self.seed_sec.setText(str(sim.get("seed_sec", "1800")))
        self.eval_sec.setText(str(sim.get("eval_sec", "3600")))
        self.num_runs.setValue(int(sim.get("num_runs", 2)))
        self.rand_seed.setValue(int(sim.get("rand_seed", 42)))
        self.rand_seed_inc.setValue(int(sim.get("rand_seed_inc", 10)))

        res = s.get("results", {})
        _set_combo_text(self.run_combo, res.get("run_combo", "Avg"))
        self.run_other.setText(res.get("run_other", ""))
        _set_combo_text(self.veh_combo, res.get("veh_type", "All"))
        _set_combo_text(self.ti_combo, res.get("ti_combo", "1"))
        self.ti_other.setText(res.get("ti_other", ""))

        out = s.get("output", {})
        self.result_dir.setText(out.get("result_dir", ""))
        self.output_file.setText(out.get("output_file", "Nodes Results.xlsx"))
        self.reporter.setText(out.get("reporter", ""))
        self.date_edit.setDate(_str_to_qdate(out.get("date", "")))
        self.company.setText(out.get("company", ""))
        self.project.setText(out.get("project", ""))













# ============================================================
# Links Tab (UPDATED — adds licensing gate to Collect)
# ============================================================

class LinksTab(QWidget, ProjectScenarioMixin):
    # Tier 1 feature
    FEATURE_RESULTS_BASIC = "results_basic"

    def __init__(self, parent=None, license_tab=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        # Optional direct reference to your LicenseTab (recommended)
        self.license_tab = license_tab

        self.sim_groupbox: Optional[QGroupBox] = None
        self.segment_rows: List[dict] = []

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # ✅ Apply modern global scrollbar styling (vertical/horizontal)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        self._build_title("Links Results")
        self._build_project_setup()
        self._build_scenario_setup()
        self._build_segments_section()
        self._build_simulation_parameters()
        self._build_results_specifics()
        self._build_output()
        self._build_collect()

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

    # ------------------------------------------------------------
    # Licensing helpers (same pattern as NodesTab)
    # ------------------------------------------------------------

    def _find_license_tab(self):
        if self.license_tab is not None:
            return self.license_tab

        p = self.parent()
        while p is not None:
            if hasattr(p, "license_tab"):
                lt = getattr(p, "license_tab")
                if lt is not None:
                    self.license_tab = lt
                    return lt
            p = p.parent()
        return None

    def _find_tab_widget(self):
        p = self.parent()
        while p is not None:
            if isinstance(p, QTabWidget):
                return p
            p = p.parent()
        return None

    def _go_to_license_tab(self, message: str):
        lt = self._find_license_tab()
        tabs = self._find_tab_widget()

        if tabs is not None and lt is not None:
            idx = tabs.indexOf(lt)
            if idx >= 0:
                tabs.setCurrentIndex(idx)

        if lt is not None:
            if hasattr(lt, "show_status"):
                lt.show_status(message)
            if hasattr(lt, "focus_key_input"):
                lt.focus_key_input()

    def _license_snapshot(self, force: bool = False) -> dict:
        lt = self._find_license_tab()
        if lt is None:
            return {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}

        if hasattr(lt, "validate_and_cache"):
            try:
                return lt.validate_and_cache(force=force)  # type: ignore
            except Exception:
                return {"app_enabled": True, "valid": False, "reason": "NO_INTERNET", "features": []}

        state = {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}
        if hasattr(lt, "is_app_enabled"):
            try:
                state["app_enabled"] = bool(lt.is_app_enabled())  # type: ignore
            except Exception:
                pass
        if hasattr(lt, "is_license_valid"):
            try:
                state["valid"] = bool(lt.is_license_valid())  # type: ignore
            except Exception:
                pass
        return state

    def _has_feature(self, feature_name: str) -> bool:
        lt = self._find_license_tab()
        if lt is None:
            return False
        if hasattr(lt, "has_feature"):
            try:
                return bool(lt.has_feature(feature_name))  # type: ignore
            except Exception:
                return False
        return False

    def _ensure_feature_or_redirect(self, feature_name: str, friendly_action: str) -> bool:
        snap = self._license_snapshot(force=True)

        if snap.get("app_enabled", True) is False:
            self._go_to_license_tab("Software is currently disabled by the publisher (shutdown switch is ON).")
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
            return False

        if not snap.get("valid", False):
            reason = str(snap.get("reason", "UNACTIVATED"))
            if reason == "NO_INTERNET":
                msg = "Internet is required to validate your license. Please connect and try again."
            elif reason == "EXPIRED":
                msg = "Your license has expired. Please renew to continue."
            elif reason == "SEAT_LIMIT":
                msg = "No seats available. Ask your company admin to deactivate a seat."
            else:
                msg = "Activation required. Please enter your license key to continue."

            self._go_to_license_tab(f"{friendly_action} requires activation.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        features = snap.get("features", [])
        has = (feature_name in features) if isinstance(features, list) else self._has_feature(feature_name)
        if not has:
            if feature_name == self.FEATURE_RESULTS_BASIC:
                msg = "Tier 1 is required to use Basic Results Reporter actions (Collect Results)."
            else:
                msg = "Your license does not include this feature."
            self._go_to_license_tab(f"{friendly_action} is locked.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        return True

    # ------------------------------------------------------------
    # Existing UI code
    # ------------------------------------------------------------

    def _build_segments_section(self):
        g = QGroupBox("Links")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Define segments and their link numbers.\n"
            "Start with one row and add more as needed.\n"
            "If Segment Type = Others, enter a custom type.\n"
            "Link Numbers are entered as separate boxes (like Travel Time Numbers)."
        )

        self.segments_container = QVBoxLayout()
        self.segments_container.setSpacing(10)
        lay.addLayout(self.segments_container)

        btn_row = QHBoxLayout()
        add_btn = QPushButton("Add Segment")
        rem_btn = QPushButton("Remove Last Segment")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.clicked.connect(self._add_segment_row_btn)
        rem_btn.clicked.connect(self._remove_segment_row)
        btn_row.addWidget(add_btn)
        btn_row.addWidget(rem_btn)
        btn_row.addStretch()
        lay.addLayout(btn_row)

        self.main_layout.addWidget(g)

        self._add_segment_row(default_name="Segment 1", default_type="Auxiliary", default_links=(1, 2, 3))

    def _add_segment_row_btn(self):
        i = len(self.segment_rows) + 1
        self._add_segment_row(default_name=f"Segment {i}", default_type="Auxiliary", default_links=(1,))

    def _add_segment_row(self, default_name: str, default_type: str, default_links: Tuple[int, ...]):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        v = QVBoxLayout(roww)
        v.setSpacing(8)

        # ----- top row: name + type -----
        top = QGridLayout()
        top.setHorizontalSpacing(10)
        top.setVerticalSpacing(8)

        name_edit = QLineEdit(default_name)
        name_edit.setFixedWidth(170)
        name_edit.setToolTip("Name used in the output file for this segment.")

        type_combo = QComboBox()
        type_combo.addItems(
            ["Weaving", "Merging", "Diverging", "Main Line", "Auxiliary", "Ramp", "Lane-Drop", "Transition", "Others"]
        )
        if default_type in [type_combo.itemText(i) for i in range(type_combo.count())]:
            type_combo.setCurrentText(default_type)
        else:
            type_combo.setCurrentText("Others")
        type_combo.setFixedWidth(150)
        type_combo.setToolTip("Select the segment type. Choose 'Others' for a custom type label.")

        type_other = QLineEdit()
        type_other.setPlaceholderText("Custom type")
        type_other.setFixedWidth(160)
        type_other.setToolTip("Custom type name used when Segment Type is 'Others'.")
        type_other.setVisible(type_combo.currentText() == "Others")
        type_combo.currentTextChanged.connect(lambda t: type_other.setVisible(t == "Others"))

        top.addWidget(QLabel("Segment Name:"), 0, 0)
        top.addWidget(name_edit, 0, 1)
        top.addWidget(QLabel("Segment Type:"), 0, 2)
        top.addWidget(type_combo, 0, 3)
        top.addWidget(type_other, 0, 4)

        top_wrap = QHBoxLayout()
        top_wrap.addLayout(top)
        top_wrap.addStretch()
        v.addLayout(top_wrap)

        # ----- link numbers: FIXED (no weird far-right boxes + no UI stretching) -----
        nums_title = QLabel("Link Numbers")
        nums_title.setStyleSheet(SUBTITLE_STYLE)
        nums_title.setToolTip("Add one box per VISSIM link number (digits only).")
        v.addWidget(nums_title)

        nums_row = QHBoxLayout()
        nums_row.setSpacing(8)

        add_num = QPushButton("Add Number")
        rem_num = QPushButton("Remove Last")
        add_num.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_num.setStyleSheet(SOFT_BUTTON_STYLE)

        nums_row.addWidget(add_num)
        nums_row.addWidget(rem_num)
        nums_row.addSpacing(10)

        # Scrollable area for the number boxes (prevents the whole UI from stretching)
        numbers_host = QWidget()
        numbers_layout = QHBoxLayout(numbers_host)
        numbers_layout.setContentsMargins(0, 0, 0, 0)
        numbers_layout.setSpacing(8)
        numbers_layout.addStretch()  # keep stretch LAST, always

        numbers_scroll = QScrollArea()
        numbers_scroll.setWidgetResizable(True)
        numbers_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        numbers_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # ✅ Apply modern global scrollbar styling (horizontal bar in this numbers scroller)
        numbers_scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        numbers_scroll.setWidget(numbers_host)
        numbers_scroll.setFixedHeight(48)

        nums_row.addWidget(numbers_scroll, 1)
        v.addLayout(nums_row)

        link_edits: List[DigitsLineEdit] = []

        def add_number_box(val: Optional[int] = None):
            e = DigitsLineEdit("" if val is None else str(val))
            e.setFixedWidth(90)
            e.setToolTip("Link Number (digits only).")
            _force_input_bg(e)
            link_edits.append(e)
            # ✅ insert BEFORE the stretch so boxes don’t jump to the far right
            numbers_layout.insertWidget(numbers_layout.count() - 1, e)

        def remove_number_box():
            if not link_edits:
                return
            e = link_edits.pop()
            numbers_layout.removeWidget(e)
            e.setParent(None)

        if not default_links:
            default_links = (1,)
        for n in default_links:
            add_number_box(int(n))

        add_num.clicked.connect(lambda: add_number_box(None))
        rem_num.clicked.connect(remove_number_box)

        # ----- pack row into container -----
        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        self.segments_container.addLayout(wrap)

        _force_input_bg_in_widget(roww)
        self.segment_rows.append(
            {
                "widget": roww,
                "name": name_edit,
                "type": type_combo,
                "type_other": type_other,
                "link_edits": link_edits,
            }
        )

    def _remove_segment_row(self):
        if not self.segment_rows:
            return
        r = self.segment_rows.pop()
        r["widget"].setParent(None)

    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.sim_groupbox = g
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Only required if you are running the model.\n"
            "Evaluation period must be greater than seeding period."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_sec = QLineEdit("1800")
        self.eval_sec = QLineEdit("3600")
        self.seed_sec.setFixedWidth(140)
        self.eval_sec.setFixedWidth(140)
        self.seed_sec.setValidator(QIntValidator(0, 999999999, self))
        self.eval_sec.setValidator(QIntValidator(0, 999999999, self))

        self.num_runs = QSpinBox()
        self.num_runs.setRange(1, 9999)
        self.num_runs.setValue(2)

        self.rand_seed = QSpinBox()
        self.rand_seed.setRange(1, 999999999)
        self.rand_seed.setValue(42)

        self.rand_seed_inc = QSpinBox()
        self.rand_seed_inc.setRange(1, 999999999)
        self.rand_seed_inc.setValue(10)

        grid.addWidget(QLabel("Seeding Period (sec):"), 0, 0)
        grid.addWidget(self.seed_sec, 0, 1)
        grid.addWidget(QLabel("Evaluation Period (sec):"), 0, 2)
        grid.addWidget(self.eval_sec, 0, 3)

        grid.addWidget(QLabel("Number of Runs:"), 1, 0)
        grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(QLabel("Random Seed:"), 1, 2)
        grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(QLabel("Random Seed Increment:"), 1, 4)
        grid.addWidget(self.rand_seed_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        self._update_sim_visibility()

    def _update_sim_visibility(self):
        if self.sim_groupbox is None:
            return
        show = any(r["run"].isChecked() for r in self.scenario_rows)
        self.sim_groupbox.setVisible(show)

    def _build_results_specifics(self):
        g = QGroupBox("Results Specifics")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Select the run, vehicle class, time interval, and results type: <b>Per Link</b> or <b>Per Lane</b>.\n"
            "If <b>Per Lane</b> is selected, the tool collects <b>Per Lane</b> and automatically calculated the <b>Per Link</b> results."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.run_combo = QComboBox()
        run_options = ["Avg", "Min", "Max", "StdDev", "Var"] + [str(i) for i in range(1, 11)] + ["Other"]
        self.run_combo.addItems(run_options)
        self.run_combo.setCurrentText("Avg")
        self.run_combo.setFixedWidth(140)

        self.run_other = QLineEdit()
        self.run_other.setPlaceholderText('If "Other", enter here')
        self.run_other.setFixedWidth(200)
        self.run_other.setEnabled(False)
        self.run_combo.currentTextChanged.connect(lambda t: self.run_other.setEnabled(t == "Other"))

        self.veh_combo = QComboBox()
        self.veh_combo.addItems(["All", "Cars", "Heavy Goods Vehicles", "Buses", "Motor Cycles", "Bicycles", "Pedestrians"])
        self.veh_combo.setCurrentText("All")
        self.veh_combo.setFixedWidth(210)

        self.ti_combo = QComboBox()
        self.ti_combo.addItems(["Avg"] + [str(i) for i in range(1, 11)] + ["Other"])
        self.ti_combo.setCurrentText("1")
        self.ti_combo.setFixedWidth(140)

        self.ti_other = QLineEdit()
        self.ti_other.setPlaceholderText('If "Other", enter here')
        self.ti_other.setFixedWidth(200)
        self.ti_other.setEnabled(False)
        self.ti_combo.currentTextChanged.connect(lambda t: self.ti_other.setEnabled(t == "Other"))

        self.results_type = QComboBox()
        self.results_type.addItems(["Per Link", "Per Lane"])
        self.results_type.setCurrentText("Per Link")
        self.results_type.setFixedWidth(160)

        self.unit_combo = QComboBox()
        self.unit_combo.addItems(["meter", "ft"])
        self.unit_combo.setCurrentText("ft")
        self.unit_combo.setFixedWidth(120)

        grid.addWidget(QLabel("Run:"), 0, 0)
        grid.addWidget(self.run_combo, 0, 1)
        grid.addWidget(self.run_other, 0, 2)

        grid.addWidget(QLabel("Vehicle Type:"), 1, 0)
        grid.addWidget(self.veh_combo, 1, 1)

        grid.addWidget(QLabel("Time Interval:"), 2, 0)
        grid.addWidget(self.ti_combo, 2, 1)
        grid.addWidget(self.ti_other, 2, 2)

        grid.addWidget(QLabel("Results Type:"), 3, 0)
        grid.addWidget(self.results_type, 3, 1)

        grid.addWidget(QLabel("Unit:"), 3, 2)
        grid.addWidget(self.unit_combo, 3, 3)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _build_output(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(g, "Select where to save the links results Excel file.")

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.clicked.connect(self._browse_result_dir)

        self.output_file = QLineEdit("Links Results.xlsx")
        self.output_file.setFixedWidth(260)

        grid.addWidget(QLabel("Results Directory:"), 0, 0)
        grid.addWidget(self.result_dir, 0, 1)
        grid.addWidget(browse, 0, 2)

        grid.addWidget(QLabel("Output File Name:"), 1, 0)
        grid.addWidget(self.output_file, 1, 1)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    def _build_collect(self):
        row = QHBoxLayout()
        self.collect_btn = QPushButton("Collect Links Results")
        self.collect_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.collect_btn.clicked.connect(self._on_collect)
        row.addStretch()
        row.addWidget(self.collect_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.main_layout.addWidget(self.log)

    def _on_collect(self):
        # ✅ Gate: Links Results collect requires Tier 1
        if not self._ensure_feature_or_redirect(self.FEATURE_RESULTS_BASIC, "Collect Links Results"):
            return

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Scenario_List: List[Scenario] = []
        for r in self.scenario_rows:
            Scenario_List.append((
                int(r["num"].value()),
                r["name"].text().strip(),
                bool(r["run"].isChecked()),
            ))

        Segments_List: List[SegmentRow] = []
        for i, seg in enumerate(self.segment_rows, start=1):
            name = seg["name"].text().strip() or f"Segment {i}"
            t = seg["type"].currentText()
            if t == "Others":
                custom = seg["type_other"].text().strip()
                t = custom if custom else "Others"

            links: List[int] = []
            for e in seg["link_edits"]:
                n = e.as_int()
                if n > 0:
                    links.append(n)

            if len(links) == 0:
                QMessageBox.warning(self, "Segments", f"{name}: please enter at least one link number.")
                return

            Segments_List.append((name, t, tuple(links)))

        Seeding_Period = int(self.seed_sec.text().strip() or "0")
        eval_period = int(self.eval_sec.text().strip() or "0")
        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_seed_inc.value())

        Run = self.run_other.text().strip() if self.run_combo.currentText() == "Other" else self.run_combo.currentText()
        Veh_Type = self.veh_combo.currentText()
        Time_interval = self.ti_other.text().strip() if self.ti_combo.currentText() == "Other" else self.ti_combo.currentText()

        Results_Type = self.results_type.currentText()
        Unit = self.unit_combo.currentText()

        Result_Directory = self.result_dir.text().strip()
        Output = self.output_file.text().strip()

        if Path_ and not os.path.isdir(Path_):
            QMessageBox.warning(self, "Project Path", "Project Path must be a directory.")
            return
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return
        if any(s[2] for s in Scenario_List):
            if eval_period <= Seeding_Period:
                QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
                return

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        log_line("VissiCaRe Links Results inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Scenario_List", Scenario_List),
            ("Segments_List", Segments_List),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
            ("Run", Run),
            ("Veh_Type", Veh_Type),
            ("Time_interval", Time_interval),
            ("Results_Type", Results_Type),
            ("Unit", Unit),
            ("Result_Directory", Result_Directory),
            ("Output", Output),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Link_Results (TEST)...\n")
        res = Link_Results(
            Results_Type=Results_Type,
            Vissim_Version=Vissim_Version,
            Filename=Filename,
            Result_Directory=Result_Directory,
            Unit=Unit,
            Scenario_List=Scenario_List,
            Segments_List=Segments_List,
            Seeding_Period=Seeding_Period,
            eval_period=eval_period,
            NumRuns=NumRuns,
            Random_Seed=Random_Seed,
            Random_SeedInc=Random_SeedInc,
            Run=Run,
            Time_interval=Time_interval,
            Veh_Type=Veh_Type,
            Output=Output,
        )
        log_line(f"\nLink_Results finished. Return: {res}")

    def export_state(self) -> dict:
        segments = []
        for seg in self.segment_rows:
            nums = [e.as_int() for e in seg["link_edits"]]
            segments.append({
                "name": seg["name"].text(),
                "type": seg["type"].currentText(),
                "type_other": seg["type_other"].text(),
                "links": nums,
            })

        return {
            "project": self.export_project_setup(),
            "scenarios": self.export_scenarios(),
            "segments": segments,
            "sim": {
                "seed_sec": self.seed_sec.text(),
                "eval_sec": self.eval_sec.text(),
                "num_runs": int(self.num_runs.value()),
                "rand_seed": int(self.rand_seed.value()),
                "rand_seed_inc": int(self.rand_seed_inc.value()),
            },
            "results": {
                "run_combo": self.run_combo.currentText(),
                "run_other": self.run_other.text(),
                "veh_type": self.veh_combo.currentText(),
                "ti_combo": self.ti_combo.currentText(),
                "ti_other": self.ti_other.text(),
                "results_type": self.results_type.currentText(),
                "unit": self.unit_combo.currentText(),
            },
            "output": {
                "result_dir": self.result_dir.text(),
                "output_file": self.output_file.text(),
            },
        }

    def import_state(self, s: dict):
        self.import_project_setup(s.get("project", {}))
        self.import_scenarios(s.get("scenarios", {}))

        # clear segments
        for seg in self.segment_rows:
            seg["widget"].setParent(None)
        self.segment_rows = []

        for seg in s.get("segments", []):
            raw_nums = seg.get("links", [])
            try:
                nums = tuple(int(n) for n in raw_nums if n is not None and int(n) > 0)
            except Exception:
                nums = (1,)
            if len(nums) == 0:
                nums = (1,)

            self._add_segment_row(
                default_name=seg.get("name", "Segment"),
                default_type=seg.get("type", "Auxiliary"),
                default_links=nums,
            )
            last = self.segment_rows[-1]
            last["type_other"].setText(seg.get("type_other", ""))

        sim = s.get("sim", {})
        self.seed_sec.setText(str(sim.get("seed_sec", "1800")))
        self.eval_sec.setText(str(sim.get("eval_sec", "3600")))
        self.num_runs.setValue(int(sim.get("num_runs", 2)))
        self.rand_seed.setValue(int(sim.get("rand_seed", 42)))
        self.rand_seed_inc.setValue(int(sim.get("rand_seed_inc", 10)))

        res = s.get("results", {})
        _set_combo_text(self.run_combo, res.get("run_combo", "Avg"))
        self.run_other.setText(res.get("run_other", ""))
        _set_combo_text(self.veh_combo, res.get("veh_type", "All"))
        _set_combo_text(self.ti_combo, res.get("ti_combo", "1"))
        self.ti_other.setText(res.get("ti_other", ""))
        _set_combo_text(self.results_type, res.get("results_type", "Per Link"))
        _set_combo_text(self.unit_combo, res.get("unit", "ft"))

        out = s.get("output", {})
        self.result_dir.setText(out.get("result_dir", ""))
        self.output_file.setText(out.get("output_file", "Links Results.xlsx"))









# ============================================================
# Travel Time Tab (UPDATED — adds licensing gate to Collect)
# ============================================================

class TravelTimeTab(QWidget, ProjectScenarioMixin):
    # Tier 1 feature
    FEATURE_RESULTS_BASIC = "results_basic"

    def __init__(self, parent=None, license_tab=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        # Optional direct reference to your LicenseTab (recommended)
        self.license_tab = license_tab

        self.sim_groupbox: Optional[QGroupBox] = None
        self.tt_rows: List[dict] = []

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # ✅ Apply modern global scrollbar styling (vertical/horizontal)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        self._build_title("Travel Time Results")
        self._build_project_setup()
        self._build_scenario_setup()
        self._build_tt_locations()
        self._build_simulation_parameters()
        self._build_results_specifics()
        self._build_output()
        self._build_collect()

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

    # ------------------------------------------------------------
    # Licensing helpers (same pattern as NodesTab)
    # ------------------------------------------------------------

    def _find_license_tab(self):
        if self.license_tab is not None:
            return self.license_tab

        p = self.parent()
        while p is not None:
            if hasattr(p, "license_tab"):
                lt = getattr(p, "license_tab")
                if lt is not None:
                    self.license_tab = lt
                    return lt
            p = p.parent()
        return None

    def _find_tab_widget(self):
        p = self.parent()
        while p is not None:
            if isinstance(p, QTabWidget):
                return p
            p = p.parent()
        return None

    def _go_to_license_tab(self, message: str):
        lt = self._find_license_tab()
        tabs = self._find_tab_widget()

        if tabs is not None and lt is not None:
            idx = tabs.indexOf(lt)
            if idx >= 0:
                tabs.setCurrentIndex(idx)

        if lt is not None:
            if hasattr(lt, "show_status"):
                lt.show_status(message)
            if hasattr(lt, "focus_key_input"):
                lt.focus_key_input()

    def _license_snapshot(self, force: bool = False) -> dict:
        lt = self._find_license_tab()
        if lt is None:
            return {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}

        if hasattr(lt, "validate_and_cache"):
            try:
                return lt.validate_and_cache(force=force)  # type: ignore
            except Exception:
                return {"app_enabled": True, "valid": False, "reason": "NO_INTERNET", "features": []}

        state = {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}
        if hasattr(lt, "is_app_enabled"):
            try:
                state["app_enabled"] = bool(lt.is_app_enabled())  # type: ignore
            except Exception:
                pass
        if hasattr(lt, "is_license_valid"):
            try:
                state["valid"] = bool(lt.is_license_valid())  # type: ignore
            except Exception:
                pass
        return state

    def _has_feature(self, feature_name: str) -> bool:
        lt = self._find_license_tab()
        if lt is None:
            return False
        if hasattr(lt, "has_feature"):
            try:
                return bool(lt.has_feature(feature_name))  # type: ignore
            except Exception:
                return False
        return False

    def _ensure_feature_or_redirect(self, feature_name: str, friendly_action: str) -> bool:
        snap = self._license_snapshot(force=True)

        if snap.get("app_enabled", True) is False:
            self._go_to_license_tab("Software is currently disabled by the publisher (shutdown switch is ON).")
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
            return False

        if not snap.get("valid", False):
            reason = str(snap.get("reason", "UNACTIVATED"))
            if reason == "NO_INTERNET":
                msg = "Internet is required to validate your license. Please connect and try again."
            elif reason == "EXPIRED":
                msg = "Your license has expired. Please renew to continue."
            elif reason == "SEAT_LIMIT":
                msg = "No seats available. Ask your company admin to deactivate a seat."
            else:
                msg = "Activation required. Please enter your license key to continue."

            self._go_to_license_tab(f"{friendly_action} requires activation.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        features = snap.get("features", [])
        has = (feature_name in features) if isinstance(features, list) else self._has_feature(feature_name)
        if not has:
            if feature_name == self.FEATURE_RESULTS_BASIC:
                msg = "Tier 1 is required to use Basic Results Reporter actions (Collect Results)."
            else:
                msg = "Your license does not include this feature."
            self._go_to_license_tab(f"{friendly_action} is locked.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        return True

    # ------------------------------------------------------------
    # Existing UI code
    # ------------------------------------------------------------

    def _build_tt_locations(self):
        g = QGroupBox("Travel Time Locations")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Select the main line you would like to collect travel time for.\n"
            "Define its direction, and define the travel time portal numbers.\n"
            "Use multiple number boxes (one per travel time number)."
        )

        self.tt_container = QVBoxLayout()
        self.tt_container.setSpacing(10)
        lay.addLayout(self.tt_container)

        btns = QHBoxLayout()
        add_btn = QPushButton("Add Location")
        rem_btn = QPushButton("Remove Last Location")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.clicked.connect(self._add_tt_row_btn)
        rem_btn.clicked.connect(self._remove_tt_row)
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addStretch()
        lay.addLayout(btns)

        self.main_layout.addWidget(g)

        # ✅ DEFAULT: ONLY ONE LOCATION (not two)
        # ✅ DEFAULT MAIN LINE: "Interstate 1"
        self._add_tt_row(default_mainline="Interstate 1", default_dir="South Bound", default_numbers=(1, 2, 3, 4))

    def _add_tt_row_btn(self):
        self._add_tt_row(default_mainline="", default_dir="South Bound", default_numbers=(1,))

    def _add_tt_row(self, default_mainline: str, default_dir: str, default_numbers: Tuple[int, ...]):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        v = QVBoxLayout(roww)
        v.setSpacing(8)

        top = QGridLayout()
        top.setHorizontalSpacing(10)
        top.setVerticalSpacing(6)

        top.addWidget(QLabel("Main Line:"), 0, 0)
        mainline = QLineEdit(default_mainline)
        mainline.setFixedWidth(220)
        mainline.setToolTip("Name of the mainline (used for labeling in outputs).")
        top.addWidget(mainline, 0, 1)

        top.addWidget(QLabel("Direction:"), 0, 2)
        direction = QComboBox()
        direction.addItems([
            "South Bound", "North Bound", "West Bound", "East Bound",
            "Southwest Bound", "Southeast Bound", "Northwest Bound", "Northeast Bound",
        ])
        if default_dir:
            direction.setCurrentText(default_dir)
        direction.setFixedWidth(170)
        direction.setToolTip("Direction label for this mainline entry.")
        top.addWidget(direction, 0, 3)

        top_wrap = QHBoxLayout()
        top_wrap.addLayout(top)
        top_wrap.addStretch()
        v.addLayout(top_wrap)

        nums_title = QLabel("Travel Time Numbers")
        nums_title.setStyleSheet(SUBTITLE_STYLE)
        nums_title.setToolTip("Add one box per Travel Time measurement number defined in VISSIM.")
        v.addWidget(nums_title)

        # ----- numbers: FIXED (no weird far-right boxes + no UI stretching) -----
        nums_row = QHBoxLayout()
        nums_row.setSpacing(8)

        add_num = QPushButton("Add Number")
        rem_num = QPushButton("Remove Last")
        add_num.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_num.setStyleSheet(SOFT_BUTTON_STYLE)

        nums_row.addWidget(add_num)
        nums_row.addWidget(rem_num)
        nums_row.addSpacing(10)

        numbers_host = QWidget()
        numbers_layout = QHBoxLayout(numbers_host)
        numbers_layout.setContentsMargins(0, 0, 0, 0)
        numbers_layout.setSpacing(8)
        numbers_layout.addStretch()  # keep stretch LAST, always

        numbers_scroll = QScrollArea()
        numbers_scroll.setWidgetResizable(True)
        numbers_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        numbers_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # ✅ Apply modern global scrollbar styling (horizontal bar in this numbers scroller)
        numbers_scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        numbers_scroll.setWidget(numbers_host)
        numbers_scroll.setFixedHeight(48)

        nums_row.addWidget(numbers_scroll, 1)
        v.addLayout(nums_row)

        numbers_edits: List[DigitsLineEdit] = []

        def add_number_box(val: Optional[int] = None):
            e = DigitsLineEdit("" if val is None else str(val))
            e.setFixedWidth(90)
            e.setToolTip("Travel Time Number (digits only).")
            _force_input_bg(e)
            numbers_edits.append(e)
            # ✅ insert BEFORE the stretch so boxes don’t jump to the far right
            numbers_layout.insertWidget(numbers_layout.count() - 1, e)

        def remove_number_box():
            if not numbers_edits:
                return
            e = numbers_edits.pop()
            numbers_layout.removeWidget(e)
            e.setParent(None)

        for n in default_numbers:
            add_number_box(n)

        add_num.clicked.connect(lambda: add_number_box(None))
        rem_num.clicked.connect(remove_number_box)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        self.tt_container.addLayout(wrap)

        _force_input_bg_in_widget(roww)
        self.tt_rows.append({
            "widget": roww,
            "mainline": mainline,
            "direction": direction,
            "num_edits": numbers_edits,
        })

    def _remove_tt_row(self):
        if not self.tt_rows:
            return
        r = self.tt_rows.pop()
        r["widget"].setParent(None)

    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.sim_groupbox = g
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Only required if you are running the model.\n"
            "Evaluation period must be greater than seeding period."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_sec = QLineEdit("1800")
        self.eval_sec = QLineEdit("3600")
        self.seed_sec.setFixedWidth(140)
        self.eval_sec.setFixedWidth(140)
        self.seed_sec.setValidator(QIntValidator(0, 999999999, self))
        self.eval_sec.setValidator(QIntValidator(0, 999999999, self))

        self.num_runs = QSpinBox()
        self.num_runs.setRange(1, 9999)
        self.num_runs.setValue(2)

        self.rand_seed = QSpinBox()
        self.rand_seed.setRange(1, 999999999)
        self.rand_seed.setValue(42)

        self.rand_seed_inc = QSpinBox()
        self.rand_seed_inc.setRange(1, 999999999)
        self.rand_seed_inc.setValue(10)

        grid.addWidget(QLabel("Seeding Period (sec):"), 0, 0)
        grid.addWidget(self.seed_sec, 0, 1)
        grid.addWidget(QLabel("Evaluation Period (sec):"), 0, 2)
        grid.addWidget(self.eval_sec, 0, 3)

        grid.addWidget(QLabel("Number of Runs:"), 1, 0)
        grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(QLabel("Random Seed:"), 1, 2)
        grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(QLabel("Random Seed Increment:"), 1, 4)
        grid.addWidget(self.rand_seed_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        self._update_sim_visibility()

    def _update_sim_visibility(self):
        if self.sim_groupbox is None:
            return
        show = any(r["run"].isChecked() for r in self.scenario_rows)
        self.sim_groupbox.setVisible(show)

    def _build_results_specifics(self):
        g = QGroupBox("Results Specifics")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Run: Avg/Min/Max/StdDev/Var or run number.\n"
            "Time Interval: Avg or interval number.\n"
            "Units: pick speed, distance, and time units."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.run_combo = QComboBox()
        run_options = ["Avg", "Min", "Max", "StdDev", "Var"] + [str(i) for i in range(1, 11)] + ["Other"]
        self.run_combo.addItems(run_options)
        self.run_combo.setCurrentText("1")
        self.run_combo.setFixedWidth(140)

        self.run_other = QLineEdit()
        self.run_other.setPlaceholderText('If "Other", enter here')
        self.run_other.setFixedWidth(200)
        self.run_other.setEnabled(False)
        self.run_combo.currentTextChanged.connect(lambda t: self.run_other.setEnabled(t == "Other"))

        self.ti_combo = QComboBox()
        self.ti_combo.addItems(["Avg"] + [str(i) for i in range(1, 11)] + ["Other"])
        self.ti_combo.setCurrentText("1")
        self.ti_combo.setFixedWidth(140)

        self.ti_other = QLineEdit()
        self.ti_other.setPlaceholderText('If "Other", enter here')
        self.ti_other.setFixedWidth(200)
        self.ti_other.setEnabled(False)
        self.ti_combo.currentTextChanged.connect(lambda t: self.ti_other.setEnabled(t == "Other"))

        self.speed_unit = QComboBox()
        self.speed_unit.addItems(["km/h", "Kilometer/Hour", "MPH", "Mile/Hour", "Other"])
        self.speed_unit.setCurrentText("MPH")
        self.speed_unit.setFixedWidth(160)

        self.speed_other = QLineEdit()
        self.speed_other.setPlaceholderText("Speed unit (Other)")
        self.speed_other.setFixedWidth(170)
        self.speed_other.setEnabled(False)
        self.speed_unit.currentTextChanged.connect(lambda t: self.speed_other.setEnabled(t == "Other"))

        self.dist_unit = QComboBox()
        self.dist_unit.addItems(["ft", "meter", "Other"])
        self.dist_unit.setCurrentText("ft")
        self.dist_unit.setFixedWidth(120)

        self.dist_other = QLineEdit()
        self.dist_other.setPlaceholderText("Distance unit (Other)")
        self.dist_other.setFixedWidth(170)
        self.dist_other.setEnabled(False)
        self.dist_unit.currentTextChanged.connect(lambda t: self.dist_other.setEnabled(t == "Other"))

        self.time_unit = QComboBox()
        self.time_unit.addItems(["Sec", "Other"])
        self.time_unit.setCurrentText("Sec")
        self.time_unit.setFixedWidth(120)

        self.time_other = QLineEdit()
        self.time_other.setPlaceholderText("Time unit (Other)")
        self.time_other.setFixedWidth(170)
        self.time_other.setEnabled(False)
        self.time_unit.currentTextChanged.connect(lambda t: self.time_other.setEnabled(t == "Other"))

        grid.addWidget(QLabel("Run:"), 0, 0)
        grid.addWidget(self.run_combo, 0, 1)
        grid.addWidget(self.run_other, 0, 2)

        grid.addWidget(QLabel("Time Interval:"), 1, 0)
        grid.addWidget(self.ti_combo, 1, 1)
        grid.addWidget(self.ti_other, 1, 2)

        grid.addWidget(QLabel("Speed Unit:"), 2, 0)
        grid.addWidget(self.speed_unit, 2, 1)
        grid.addWidget(self.speed_other, 2, 2)

        grid.addWidget(QLabel("Distance Unit:"), 3, 0)
        grid.addWidget(self.dist_unit, 3, 1)
        grid.addWidget(self.dist_other, 3, 2)

        grid.addWidget(QLabel("Time Unit:"), 4, 0)
        grid.addWidget(self.time_unit, 4, 1)
        grid.addWidget(self.time_other, 4, 2)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _build_output(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(g, "Select where to save the travel time results Excel file.")

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.clicked.connect(self._browse_result_dir)

        self.output_file = QLineEdit("Travel Time Results.xlsx")
        self.output_file.setFixedWidth(260)

        grid.addWidget(QLabel("Results Directory:"), 0, 0)
        grid.addWidget(self.result_dir, 0, 1)
        grid.addWidget(browse, 0, 2)

        grid.addWidget(QLabel("Output File Name:"), 1, 0)
        grid.addWidget(self.output_file, 1, 1)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    def _build_collect(self):
        row = QHBoxLayout()
        self.collect_btn = QPushButton("Collect Travel Time Results")
        self.collect_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.collect_btn.clicked.connect(self._on_collect)
        row.addStretch()
        row.addWidget(self.collect_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.main_layout.addWidget(self.log)

    def _on_collect(self):
        # ✅ Gate: Travel Time collect requires Tier 1
        if not self._ensure_feature_or_redirect(self.FEATURE_RESULTS_BASIC, "Collect Travel Time Results"):
            return

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Scenario_List: List[Scenario] = []
        for r in self.scenario_rows:
            Scenario_List.append((
                int(r["num"].value()),
                r["name"].text().strip(),
                bool(r["run"].isChecked()),
            ))

        TT_List: List[TTEntry] = []
        for row in self.tt_rows:
            main = row["mainline"].text().strip()
            if not main:
                QMessageBox.warning(self, "Travel Time Locations", "Main Line cannot be empty.")
                return
            direction = row["direction"].currentText()

            nums: List[int] = []
            for e in row["num_edits"]:
                n = e.as_int()
                if n > 0:
                    nums.append(n)

            if len(nums) == 0:
                QMessageBox.warning(self, "Travel Time Locations", f"{main} ({direction}): add at least one Travel Time Number.")
                return

            TT_List.append((main, direction, tuple(nums)))

        Seeding_Period = int(self.seed_sec.text().strip() or "0")
        eval_period = int(self.eval_sec.text().strip() or "0")
        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_seed_inc.value())

        Run = self.run_other.text().strip() if self.run_combo.currentText() == "Other" else self.run_combo.currentText()
        Time_interval = self.ti_other.text().strip() if self.ti_combo.currentText() == "Other" else self.ti_combo.currentText()

        speed = self.speed_other.text().strip() if self.speed_unit.currentText() == "Other" else self.speed_unit.currentText()
        dist = self.dist_other.text().strip() if self.dist_unit.currentText() == "Other" else self.dist_unit.currentText()
        timeu = self.time_other.text().strip() if self.time_unit.currentText() == "Other" else self.time_unit.currentText()
        Units = (speed, dist, timeu)


        Result_Directory = self.result_dir.text().strip()
        output_file = self.output_file.text().strip()

        if Path_ and not os.path.isdir(Path_):
            QMessageBox.warning(self, "Project Path", "Project Path must be a directory.")
            return
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return
        if any(s[2] for s in Scenario_List):
            if eval_period <= Seeding_Period:
                QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
                return

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        log_line("VissiCaRe Travel Time Results inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Result_Directory", Result_Directory),
            ("Scenario_List", Scenario_List),
            ("TT_List", TT_List),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
            ("Run", Run),
            ("Time_interval", Time_interval),
            ("Units", Units),
            ("output_file", output_file),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Simulation_TT (TEST)...\n")
        res = Simulation_TT(
            Vissim_Version,
            Filename,
            Result_Directory,
            Scenario_List,
            TT_List,
            Seeding_Period,
            eval_period,
            NumRuns,
            Random_Seed,
            Random_SeedInc,
            Run,
            Time_interval,
            Units,
            output_file,
        )
        log_line(f"\nSimulation_TT finished. Return: {res}")

    def export_state(self) -> dict:
        rows = []
        for r in self.tt_rows:
            nums = []
            for e in r["num_edits"]:
                nums.append(e.as_int())
            rows.append({
                "mainline": r["mainline"].text(),
                "direction": r["direction"].currentText(),
                "numbers": nums,
            })

        return {
            "project": self.export_project_setup(),
            "scenarios": self.export_scenarios(),
            "tt_rows": rows,
            "sim": {
                "seed_sec": self.seed_sec.text(),
                "eval_sec": self.eval_sec.text(),
                "num_runs": int(self.num_runs.value()),
                "rand_seed": int(self.rand_seed.value()),
                "rand_seed_inc": int(self.rand_seed_inc.value()),
            },
            "results": {
                "run_combo": self.run_combo.currentText(),
                "run_other": self.run_other.text(),
                "ti_combo": self.ti_combo.currentText(),
                "ti_other": self.ti_other.text(),
                "speed_unit": self.speed_unit.currentText(),
                "speed_other": self.speed_other.text(),
                "dist_unit": self.dist_unit.currentText(),
                "dist_other": self.dist_other.text(),
                "time_unit": self.time_unit.currentText(),
                "time_other": self.time_other.text(),
            },
            "output": {
                "result_dir": self.result_dir.text(),
                "output_file": self.output_file.text(),
            },
        }

    def import_state(self, s: dict):
        self.import_project_setup(s.get("project", {}))
        self.import_scenarios(s.get("scenarios", {}))

        for r in self.tt_rows:
            r["widget"].setParent(None)
        self.tt_rows = []

        for r in s.get("tt_rows", []):
            nums = tuple(int(n) for n in (r.get("numbers") or []) if n is not None)
            if len(nums) == 0:
                nums = (1,)
            self._add_tt_row(
                default_mainline=r.get("mainline", ""),
                default_dir=r.get("direction", "South Bound"),
                default_numbers=nums,
            )

        sim = s.get("sim", {})
        self.seed_sec.setText(str(sim.get("seed_sec", "1800")))
        self.eval_sec.setText(str(sim.get("eval_sec", "3600")))
        self.num_runs.setValue(int(sim.get("num_runs", 2)))
        self.rand_seed.setValue(int(sim.get("rand_seed", 42)))
        self.rand_seed_inc.setValue(int(sim.get("rand_seed_inc", 10)))

        res = s.get("results", {})
        _set_combo_text(self.run_combo, res.get("run_combo", "1"))
        self.run_other.setText(res.get("run_other", ""))
        _set_combo_text(self.ti_combo, res.get("ti_combo", "1"))
        self.ti_other.setText(res.get("ti_other", ""))

        _set_combo_text(self.speed_unit, res.get("speed_unit", "MPH"))
        self.speed_other.setText(res.get("speed_other", ""))
        _set_combo_text(self.dist_unit, res.get("dist_unit", "ft"))
        self.dist_other.setText(res.get("dist_other", ""))
        _set_combo_text(self.time_unit, res.get("time_unit", "Sec"))
        self.time_other.setText(res.get("time_other", ""))

        out = s.get("output", {})
        self.result_dir.setText(out.get("result_dir", ""))
        self.output_file.setText(out.get("output_file", "Travel Time Results.xlsx"))






























# ============================================================
# Throughput Tab (UPDATED — adds licensing gate to Collect)
# ============================================================

class ThroughputTab(QWidget, ProjectScenarioMixin):
    # Tier 1 feature
    FEATURE_RESULTS_BASIC = "results_basic"

    def __init__(self, parent=None, license_tab=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        # Optional direct reference to your LicenseTab (recommended)
        self.license_tab = license_tab

        self.sim_groupbox: Optional[QGroupBox] = None
        self.tp_rows: List[dict] = []

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # ✅ Apply modern global scrollbar styling (vertical/horizontal)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        # ✅ FIX: Throughput typo
        self._build_title("Throughput Results")
        self._build_project_setup()
        self._build_scenario_setup()
        self._build_throughput_nodes()
        self._build_simulation_parameters()
        self._build_results_specifics()
        self._build_output()
        self._build_collect()

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

    # ------------------------------------------------------------
    # Licensing helpers (same pattern as LinksTab / TravelTimeTab)
    # ------------------------------------------------------------

    def _find_license_tab(self):
        if self.license_tab is not None:
            return self.license_tab

        p = self.parent()
        while p is not None:
            if hasattr(p, "license_tab"):
                lt = getattr(p, "license_tab")
                if lt is not None:
                    self.license_tab = lt
                    return lt
            p = p.parent()
        return None

    def _find_tab_widget(self):
        p = self.parent()
        while p is not None:
            if isinstance(p, QTabWidget):
                return p
            p = p.parent()
        return None

    def _go_to_license_tab(self, message: str):
        lt = self._find_license_tab()
        tabs = self._find_tab_widget()

        if tabs is not None and lt is not None:
            idx = tabs.indexOf(lt)
            if idx >= 0:
                tabs.setCurrentIndex(idx)

        if lt is not None:
            if hasattr(lt, "show_status"):
                lt.show_status(message)
            if hasattr(lt, "focus_key_input"):
                lt.focus_key_input()

    def _license_snapshot(self, force: bool = False) -> dict:
        lt = self._find_license_tab()
        if lt is None:
            return {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}

        if hasattr(lt, "validate_and_cache"):
            try:
                return lt.validate_and_cache(force=force)  # type: ignore
            except Exception:
                return {"app_enabled": True, "valid": False, "reason": "NO_INTERNET", "features": []}

        state = {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}
        if hasattr(lt, "is_app_enabled"):
            try:
                state["app_enabled"] = bool(lt.is_app_enabled())  # type: ignore
            except Exception:
                pass
        if hasattr(lt, "is_license_valid"):
            try:
                state["valid"] = bool(lt.is_license_valid())  # type: ignore
            except Exception:
                pass
        return state

    def _has_feature(self, feature_name: str) -> bool:
        lt = self._find_license_tab()
        if lt is None:
            return False
        if hasattr(lt, "has_feature"):
            try:
                return bool(lt.has_feature(feature_name))  # type: ignore
            except Exception:
                return False
        return False

    def _ensure_feature_or_redirect(self, feature_name: str, friendly_action: str) -> bool:
        snap = self._license_snapshot(force=True)

        if snap.get("app_enabled", True) is False:
            self._go_to_license_tab("Software is currently disabled by the publisher (shutdown switch is ON).")
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
            return False

        if not snap.get("valid", False):
            reason = str(snap.get("reason", "UNACTIVATED"))
            if reason == "NO_INTERNET":
                msg = "Internet is required to validate your license. Please connect and try again."
            elif reason == "EXPIRED":
                msg = "Your license has expired. Please renew to continue."
            elif reason == "SEAT_LIMIT":
                msg = "No seats available. Ask your company admin to deactivate a seat."
            else:
                msg = "Activation required. Please enter your license key to continue."

            self._go_to_license_tab(f"{friendly_action} requires activation.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        features = snap.get("features", [])
        has = (feature_name in features) if isinstance(features, list) else self._has_feature(feature_name)
        if not has:
            if feature_name == self.FEATURE_RESULTS_BASIC:
                msg = "Tier 1 is required to use Basic Results Reporter actions (Collect Results)."
            else:
                msg = "Your license does not include this feature."
            self._go_to_license_tab(f"{friendly_action} is locked.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        return True

    # --------------------------
    # C - Nodes to Collect the Throughput
    # --------------------------
    def _build_throughput_nodes(self):
        g = QGroupBox("Nodes to Collect the Throughput")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Select the main line you would like to collect Throughput on.\n"
            "Define its direction, and define the node numbers to collect this data.\n"
            "The software will grab the name of the node and use it in the output, so it is a good practice to name the nodes for throughput."
        )

        self.tp_container = QVBoxLayout()
        self.tp_container.setSpacing(10)
        lay.addLayout(self.tp_container)

        btns = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        rem_btn = QPushButton("Remove Last Row")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.clicked.connect(self._add_tp_row_btn)
        rem_btn.clicked.connect(self._remove_tp_row)
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addStretch()
        lay.addLayout(btns)

        self.main_layout.addWidget(g)

        # ✅ DEFAULT: ONLY ONE ROW (same as you did in Travel Time)
        # ✅ DEFAULT MAIN LINE: "Interstate 1"
        self._add_tp_row(default_mainline="Interstate 1", default_dir="West Bound", default_numbers=(100, 101, 102))

    def _add_tp_row_btn(self):
        self._add_tp_row(default_mainline="", default_dir="West Bound", default_numbers=(1,))

    def _add_tp_row(self, default_mainline: str, default_dir: str, default_numbers: Tuple[int, ...]):
        roww = QWidget()
        roww.setStyleSheet(TRANSPARENT_CONTAINER_STYLE)
        v = QVBoxLayout(roww)
        v.setSpacing(8)

        top = QGridLayout()
        top.setHorizontalSpacing(10)
        top.setVerticalSpacing(6)

        top.addWidget(QLabel("Main Line:"), 0, 0)
        mainline = QLineEdit(default_mainline)
        mainline.setFixedWidth(220)
        mainline.setToolTip("Name of the mainline (used for labeling in outputs).")
        top.addWidget(mainline, 0, 1)

        top.addWidget(QLabel("Direction:"), 0, 2)
        direction = QComboBox()
        direction.addItems([
            "South Bound", "North Bound", "West Bound", "East Bound",
            "Southwest Bound", "Southeast Bound", "Northwest Bound", "Northeast Bound",
        ])
        if default_dir:
            direction.setCurrentText(default_dir)
        direction.setFixedWidth(170)
        direction.setToolTip("Direction label for this mainline entry.")
        top.addWidget(direction, 0, 3)

        top_wrap = QHBoxLayout()
        top_wrap.addLayout(top)
        top_wrap.addStretch()
        v.addLayout(top_wrap)

        nums_title = QLabel("Node Numbers")
        nums_title.setStyleSheet(SUBTITLE_STYLE)
        nums_title.setToolTip("Add one box per Node Number (digits only).")
        v.addWidget(nums_title)

        # ----- numbers: FIXED (no weird far-right boxes + no UI stretching) -----
        nums_row = QHBoxLayout()
        nums_row.setSpacing(8)

        add_num = QPushButton("Add Number")
        rem_num = QPushButton("Remove Last")
        add_num.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_num.setStyleSheet(SOFT_BUTTON_STYLE)
        add_num.setToolTip("Add another Node Number input box.")
        rem_num.setToolTip("Remove the last Node Number input box.")

        nums_row.addWidget(add_num)
        nums_row.addWidget(rem_num)
        nums_row.addSpacing(10)

        numbers_host = QWidget()
        numbers_layout = QHBoxLayout(numbers_host)
        numbers_layout.setContentsMargins(0, 0, 0, 0)
        numbers_layout.setSpacing(8)
        numbers_layout.addStretch()  # keep stretch LAST, always

        numbers_scroll = QScrollArea()
        numbers_scroll.setWidgetResizable(True)
        numbers_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        numbers_scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # ✅ Apply modern global scrollbar styling (horizontal bar in this numbers scroller)
        numbers_scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        numbers_scroll.setWidget(numbers_host)
        numbers_scroll.setFixedHeight(48)

        nums_row.addWidget(numbers_scroll, 1)
        v.addLayout(nums_row)

        numbers_edits: List[DigitsLineEdit] = []

        def add_number_box(val: Optional[int] = None):
            e = DigitsLineEdit("" if val is None else str(val))
            e.setFixedWidth(110)
            e.setToolTip("Node Number (digits only).")
            _force_input_bg(e)
            numbers_edits.append(e)
            numbers_layout.insertWidget(numbers_layout.count() - 1, e)

        def remove_number_box():
            if not numbers_edits:
                return
            e = numbers_edits.pop()
            numbers_layout.removeWidget(e)
            e.setParent(None)

        for n in default_numbers:
            add_number_box(n)

        add_num.clicked.connect(lambda: add_number_box(None))
        rem_num.clicked.connect(remove_number_box)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        self.tp_container.addLayout(wrap)

        _force_input_bg_in_widget(roww)
        self.tp_rows.append({
            "widget": roww,
            "mainline": mainline,
            "direction": direction,
            "num_edits": numbers_edits,
        })

    def _remove_tp_row(self):
        if not self.tp_rows:
            return
        r = self.tp_rows.pop()
        r["widget"].setParent(None)

    # --------------------------
    # D - Simulation Parameters (only if any Run Model)
    # --------------------------
    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.sim_groupbox = g
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Only required if you are running the model.\n"
            "Evaluation period must be greater than seeding period."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_sec = QLineEdit("1800")
        self.eval_sec = QLineEdit("3600")
        self.seed_sec.setFixedWidth(140)
        self.eval_sec.setFixedWidth(140)
        self.seed_sec.setValidator(QIntValidator(0, 999999999, self))
        self.eval_sec.setValidator(QIntValidator(0, 999999999, self))
        self.seed_sec.setToolTip("Seeding (warm-up) period in seconds.")
        self.eval_sec.setToolTip("Evaluation period in seconds (must be > Seeding Period).")

        self.num_runs = QSpinBox()
        self.num_runs.setRange(1, 9999)
        self.num_runs.setValue(2)
        self.num_runs.setToolTip("Number of simulation runs per scenario.")

        self.rand_seed = QSpinBox()
        self.rand_seed.setRange(1, 999999999)
        self.rand_seed.setValue(42)
        self.rand_seed.setToolTip("Base random seed for the first run.")

        self.rand_seed_inc = QSpinBox()
        self.rand_seed_inc.setRange(1, 999999999)
        self.rand_seed_inc.setValue(10)
        self.rand_seed_inc.setToolTip("Seed increment added between runs.")

        grid.addWidget(QLabel("Seeding Period (sec):"), 0, 0)
        grid.addWidget(self.seed_sec, 0, 1)
        grid.addWidget(QLabel("Evaluation Period (sec):"), 0, 2)
        grid.addWidget(self.eval_sec, 0, 3)

        grid.addWidget(QLabel("Number of Runs:"), 1, 0)
        grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(QLabel("Random Seed:"), 1, 2)
        grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(QLabel("Random Seed Increment:"), 1, 4)
        grid.addWidget(self.rand_seed_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        self._update_sim_visibility()

    def _update_sim_visibility(self):
        if self.sim_groupbox is None:
            return
        show = any(r["run"].isChecked() for r in self.scenario_rows)
        self.sim_groupbox.setVisible(show)

    # --------------------------
    # E - Results Specifics
    # --------------------------
    def _build_results_specifics(self):
        g = QGroupBox("Results Specifics")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Run: Avg/Min/Max/StdDev/Var or run number.\n"
            "Vehicle Type: select the vehicle class.\n"
            "Time Interval: Avg or interval number."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.run_combo = QComboBox()
        run_options = ["Avg", "Min", "Max", "StdDev", "Var"] + [str(i) for i in range(1, 11)] + ["Other"]
        self.run_combo.addItems(run_options)
        self.run_combo.setCurrentText("Avg")
        self.run_combo.setFixedWidth(140)
        self.run_combo.setToolTip("Which run statistic to extract (Avg/Min/Max/etc.) or a specific run number.")

        self.run_other = QLineEdit()
        self.run_other.setPlaceholderText('If "Other", enter here')
        self.run_other.setFixedWidth(200)
        self.run_other.setEnabled(False)
        self.run_other.setToolTip('Used only when Run = "Other".')
        self.run_combo.currentTextChanged.connect(lambda t: self.run_other.setEnabled(t == "Other"))

        # ✅ FIX: Vehicle types (standardized)
        self.veh_combo = QComboBox()
        self.veh_combo.addItems(["All", "Cars", "HGVs", "Buses", "Motorcycles", "Bikes", "Pedestrians"])
        self.veh_combo.setCurrentText("All")
        self.veh_combo.setFixedWidth(210)
        self.veh_combo.setToolTip("Vehicle class to extract results for.")

        self.ti_combo = QComboBox()
        self.ti_combo.addItems(["Avg"] + [str(i) for i in range(1, 11)] + ["Other"])
        self.ti_combo.setCurrentText("1")
        self.ti_combo.setFixedWidth(140)
        self.ti_combo.setToolTip("Time interval (Avg or interval number).")

        self.ti_other = QLineEdit()
        self.ti_other.setPlaceholderText('If "Other", enter here')
        self.ti_other.setFixedWidth(200)
        self.ti_other.setEnabled(False)
        self.ti_other.setToolTip('Used only when Time Interval = "Other".')
        self.ti_combo.currentTextChanged.connect(lambda t: self.ti_other.setEnabled(t == "Other"))

        grid.addWidget(QLabel("Run:"), 0, 0)
        grid.addWidget(self.run_combo, 0, 1)
        grid.addWidget(self.run_other, 0, 2)

        grid.addWidget(QLabel("Vehicle Type:"), 1, 0)
        grid.addWidget(self.veh_combo, 1, 1)

        grid.addWidget(QLabel("Time Interval:"), 2, 0)
        grid.addWidget(self.ti_combo, 2, 1)
        grid.addWidget(self.ti_other, 2, 2)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    # --------------------------
    # F - Output
    # --------------------------
    def _build_output(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # ✅ FIX: Throughput typo in tooltip
        add_corner_help(g, "Select where to save the throughput results Excel file.")

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        self.result_dir.setToolTip("Folder where the Excel output will be saved.")

        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.setToolTip("Browse for an output folder.")
        browse.clicked.connect(self._browse_result_dir)

        # ✅ FIX: Throughput typo in default filename
        self.output_file = QLineEdit("Throughput Results.xlsx")
        self.output_file.setFixedWidth(260)
        self.output_file.setToolTip("Excel file name (e.g., 'Throughput Results.xlsx').")

        grid.addWidget(QLabel("Results Directory:"), 0, 0)
        grid.addWidget(self.result_dir, 0, 1)
        grid.addWidget(browse, 0, 2)

        grid.addWidget(QLabel("Output File Name:"), 1, 0)
        grid.addWidget(self.output_file, 1, 1)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    # --------------------------
    # Collect + Log
    # --------------------------
    def _build_collect(self):
        row = QHBoxLayout()

        # ✅ FIX: Throughput typo on button
        self.collect_btn = QPushButton("Collect Throughput Results")
        self.collect_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.collect_btn.setToolTip("Run throughput collection for the selected scenarios.")
        self.collect_btn.clicked.connect(self._on_collect)
        row.addStretch()
        row.addWidget(self.collect_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.main_layout.addWidget(self.log)

    def _on_collect(self):
        # ✅ Gate: Throughput collect requires Tier 1
        # ✅ FIX: Throughput typo in friendly action
        if not self._ensure_feature_or_redirect(self.FEATURE_RESULTS_BASIC, "Collect Throughput Results"):
            return

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Scenario_List: List[Scenario] = []
        for r in self.scenario_rows:
            Scenario_List.append((
                int(r["num"].value()),
                r["name"].text().strip(),
                bool(r["run"].isChecked()),
            ))

        Nodes_List: List[TPEntry] = []
        for row in self.tp_rows:
            main = row["mainline"].text().strip()
            if not main:
                QMessageBox.warning(self, "Nodes to Collect the Throughput", "Main Line cannot be empty.")
                return
            direction = row["direction"].currentText()

            nums: List[int] = []
            for e in row["num_edits"]:
                n = e.as_int()
                if n > 0:
                    nums.append(n)

            if len(nums) == 0:
                QMessageBox.warning(
                    self,
                    "Nodes to Collect the Throughput",
                    f"{main} ({direction}): add at least one Node Number."
                )
                return

            Nodes_List.append((main, direction, tuple(nums)))

        Seeding_Period = int(self.seed_sec.text().strip() or "0")
        eval_period = int(self.eval_sec.text().strip() or "0")
        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_seed_inc.value())

        Run = self.run_other.text().strip() if self.run_combo.currentText() == "Other" else self.run_combo.currentText()
        Veh_Type = self.veh_combo.currentText()

        # ✅ Time interval variable name should be Time_interval (not Time_interval)
        Time_interval = self.ti_other.text().strip() if self.ti_combo.currentText() == "Other" else self.ti_combo.currentText()

        Result_Directory = self.result_dir.text().strip()
        Output = self.output_file.text().strip()

        if Path_ and not os.path.isdir(Path_):
            QMessageBox.warning(self, "Project Path", "Project Path must be a directory.")
            return
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return
        if any(s[2] for s in Scenario_List):
            if eval_period <= Seeding_Period:
                QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
                return

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        # ✅ FIX: Throughput typo in header
        log_line("VissiCaRe Throughput Results inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Result_Directory", Result_Directory),
            ("Scenario_List", Scenario_List),
            ("Nodes_List", Nodes_List),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
            ("Run", Run),
            ("Veh_Type", Veh_Type),
            # ✅ FIX: key name should be Time_interval
            ("Time_interval", Time_interval),
            ("Output", Output),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Simulation_Throughput (TEST)...\n")
        res = Simulation_Throughput(
            Vissim_Version,
            Filename,
            Result_Directory,
            Scenario_List,
            Nodes_List,
            Seeding_Period,
            eval_period,
            NumRuns,
            Random_Seed,
            Random_SeedInc,
            Run,
            Veh_Type,
            # ✅ FIX: pass Time_interval
            Time_interval,
            Output,
        )
        log_line(f"\nSimulation_Throughput finished. Return: {res}")

    def export_state(self) -> dict:
        rows = []
        for r in self.tp_rows:
            nums = []
            for e in r["num_edits"]:
                nums.append(e.as_int())
            rows.append({
                "mainline": r["mainline"].text(),
                "direction": r["direction"].currentText(),
                "numbers": nums,
            })

        return {
            "project": self.export_project_setup(),
            "scenarios": self.export_scenarios(),
            "tp_rows": rows,
            "sim": {
                "seed_sec": self.seed_sec.text(),
                "eval_sec": self.eval_sec.text(),
                "num_runs": int(self.num_runs.value()),
                "rand_seed": int(self.rand_seed.value()),
                "rand_seed_inc": int(self.rand_seed_inc.value()),
            },
            "results": {
                "run_combo": self.run_combo.currentText(),
                "run_other": self.run_other.text(),
                "veh_type": self.veh_combo.currentText(),
                "ti_combo": self.ti_combo.currentText(),
                "ti_other": self.ti_other.text(),
            },
            "output": {
                "result_dir": self.result_dir.text(),
                "output_file": self.output_file.text(),
            },
        }

    def import_state(self, s: dict):
        self.import_project_setup(s.get("project", {}))
        self.import_scenarios(s.get("scenarios", {}))

        for r in self.tp_rows:
            r["widget"].setParent(None)
        self.tp_rows = []

        for r in s.get("tp_rows", []):
            nums = tuple(int(n) for n in (r.get("numbers") or []) if n is not None)
            if len(nums) == 0:
                nums = (1,)
            self._add_tp_row(
                default_mainline=r.get("mainline", ""),
                default_dir=r.get("direction", "West Bound"),
                default_numbers=nums,
            )

        sim = s.get("sim", {})
        self.seed_sec.setText(str(sim.get("seed_sec", "1800")))
        self.eval_sec.setText(str(sim.get("eval_sec", "3600")))
        self.num_runs.setValue(int(sim.get("num_runs", 2)))
        self.rand_seed.setValue(int(sim.get("rand_seed", 42)))
        self.rand_seed_inc.setValue(int(sim.get("rand_seed_inc", 10)))

        res = s.get("results", {})
        _set_combo_text(self.run_combo, res.get("run_combo", "Avg"))
        self.run_other.setText(res.get("run_other", ""))
        _set_combo_text(self.veh_combo, res.get("veh_type", "All"))
        _set_combo_text(self.ti_combo, res.get("ti_combo", "1"))
        self.ti_other.setText(res.get("ti_other", ""))

        out = s.get("output", {})
        self.result_dir.setText(out.get("result_dir", ""))
        # ✅ FIX: Throughput typo in import default
        self.output_file.setText(out.get("output_file", "Throughput Results.xlsx"))




# ============================================================
# Network Results Tab (HOVER help tooltips — no corner "*")
# ============================================================

DESIRED_METRIC_OPTIONS = [
    "Delay (average)",
    "Delay (latent)",
    "Delay stopped (average)",
    "Delay stopped (total)",
    "Delay (total)",
    "Demand (latent)",
    "Distance (total)",
    "Speed (average)",
    "Stops (average)",
    "Stops (total)",
    "Travel time (total)",
    "Vehicles (arrived)",
    "Vehicles (active)",
    "Emissions CO2",
    "Emissions CO",
    "Fuel consumption",
]


class NetworkTab(QWidget, ProjectScenarioMixin):
    # Tier 1 feature
    FEATURE_RESULTS_BASIC = "results_basic"

    def __init__(self, parent=None, license_tab=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        # Optional direct reference to your LicenseTab (recommended)
        self.license_tab = license_tab

        self.sim_groupbox: Optional[QGroupBox] = None

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # ✅ Apply modern global scrollbar styling (vertical/horizontal)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        self._build_title("Network Results")
        self._build_project_setup()           # A
        self._build_scenario_setup()          # B
        self._build_simulation_parameters()   # C (conditional visibility)
        self._build_collected_variables()     # D
        self._build_results_specifics()       # E
        self._build_output()                  # F
        self._build_collect()

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

    # ------------------------------------------------------------
    # Licensing helpers (same pattern as LinksTab / TravelTimeTab)
    # ------------------------------------------------------------

    def _find_license_tab(self):
        if self.license_tab is not None:
            return self.license_tab

        p = self.parent()
        while p is not None:
            if hasattr(p, "license_tab"):
                lt = getattr(p, "license_tab")
                if lt is not None:
                    self.license_tab = lt
                    return lt
            p = p.parent()
        return None

    def _find_tab_widget(self):
        p = self.parent()
        while p is not None:
            if isinstance(p, QTabWidget):
                return p
            p = p.parent()
        return None

    def _go_to_license_tab(self, message: str):
        lt = self._find_license_tab()
        tabs = self._find_tab_widget()

        if tabs is not None and lt is not None:
            idx = tabs.indexOf(lt)
            if idx >= 0:
                tabs.setCurrentIndex(idx)

        if lt is not None:
            if hasattr(lt, "show_status"):
                lt.show_status(message)
            if hasattr(lt, "focus_key_input"):
                lt.focus_key_input()

    def _license_snapshot(self, force: bool = False) -> dict:
        lt = self._find_license_tab()
        if lt is None:
            return {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}

        if hasattr(lt, "validate_and_cache"):
            try:
                return lt.validate_and_cache(force=force)  # type: ignore
            except Exception:
                return {"app_enabled": True, "valid": False, "reason": "NO_INTERNET", "features": []}

        state = {"app_enabled": True, "valid": False, "reason": "UNACTIVATED", "features": []}
        if hasattr(lt, "is_app_enabled"):
            try:
                state["app_enabled"] = bool(lt.is_app_enabled())  # type: ignore
            except Exception:
                pass
        if hasattr(lt, "is_license_valid"):
            try:
                state["valid"] = bool(lt.is_license_valid())  # type: ignore
            except Exception:
                pass
        return state

    def _has_feature(self, feature_name: str) -> bool:
        lt = self._find_license_tab()
        if lt is None:
            return False
        if hasattr(lt, "has_feature"):
            try:
                return bool(lt.has_feature(feature_name))  # type: ignore
            except Exception:
                return False
        return False

    def _ensure_feature_or_redirect(self, feature_name: str, friendly_action: str) -> bool:
        snap = self._license_snapshot(force=True)

        if snap.get("app_enabled", True) is False:
            self._go_to_license_tab("Software is currently disabled by the publisher (shutdown switch is ON).")
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
            return False

        if not snap.get("valid", False):
            reason = str(snap.get("reason", "UNACTIVATED"))
            if reason == "NO_INTERNET":
                msg = "Internet is required to validate your license. Please connect and try again."
            elif reason == "EXPIRED":
                msg = "Your license has expired. Please renew to continue."
            elif reason == "SEAT_LIMIT":
                msg = "No seats available. Ask your company admin to deactivate a seat."
            else:
                msg = "Activation required. Please enter your license key to continue."

            self._go_to_license_tab(f"{friendly_action} requires activation.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        features = snap.get("features", [])
        has = (feature_name in features) if isinstance(features, list) else self._has_feature(feature_name)
        if not has:
            if feature_name == self.FEATURE_RESULTS_BASIC:
                msg = "Tier 1 is required to use Basic Results Reporter actions (Collect Results)."
            else:
                msg = "Your license does not include this feature."
            self._go_to_license_tab(f"{friendly_action} is locked.\n\n{msg}")
            QMessageBox.warning(self, "License", msg)
            return False

        return True

    # --------------------------
    # C - Simulation Parameters (only if any Run Model)
    # --------------------------
    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.sim_groupbox = g
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Only required if you are running the model.\n"
            "Evaluation period must be greater than seeding period."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_sec = QLineEdit("1800")
        self.eval_sec = QLineEdit("3600")
        self.seed_sec.setFixedWidth(140)
        self.eval_sec.setFixedWidth(140)
        self.seed_sec.setValidator(QIntValidator(0, 999999999, self))
        self.eval_sec.setValidator(QIntValidator(0, 999999999, self))
        self.seed_sec.setToolTip("Seeding (warm-up) period in seconds.")
        self.eval_sec.setToolTip("Evaluation period in seconds (must be > Seeding Period).")

        self.num_runs = QSpinBox()
        self.num_runs.setRange(1, 9999)
        self.num_runs.setValue(2)
        self.num_runs.setToolTip("Number of simulation runs per scenario.")

        self.rand_seed = QSpinBox()
        self.rand_seed.setRange(1, 999999999)
        self.rand_seed.setValue(42)
        self.rand_seed.setToolTip("Base random seed for the first run.")

        self.rand_seed_inc = QSpinBox()
        self.rand_seed_inc.setRange(1, 999999999)
        self.rand_seed_inc.setValue(10)
        self.rand_seed_inc.setToolTip("Seed increment added between runs.")

        grid.addWidget(QLabel("Seeding Period (sec):"), 0, 0)
        grid.addWidget(self.seed_sec, 0, 1)
        grid.addWidget(QLabel("Evaluation Period (sec):"), 0, 2)
        grid.addWidget(self.eval_sec, 0, 3)

        grid.addWidget(QLabel("Number of Runs:"), 1, 0)
        grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(QLabel("Random Seed:"), 1, 2)
        grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(QLabel("Random Seed Increment:"), 1, 4)
        grid.addWidget(self.rand_seed_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        self._update_sim_visibility()

    def _update_sim_visibility(self):
        if self.sim_groupbox is None:
            return
        show = any(r["run"].isChecked() for r in self.scenario_rows)
        self.sim_groupbox.setVisible(show)

    # --------------------------
    # D - Collected Variables
    # --------------------------
    def _build_collected_variables(self):
        g = QGroupBox("Collected Variables")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Desired Output: Select the variables you wish to collect results for in the order you want them to appear in the final output.\n"
            "If you wish to select the environment variables, make sure you have their results in your VISSIM, otherwise the software will produce an error.\n"
            "Environment Variables: If you have Emissions CO2, Emissions CO, or Fuel consumption and wish to collect these results, select Yes; otherwise leave No."
        )

        top = QGridLayout()
        top.setHorizontalSpacing(10)
        top.setVerticalSpacing(10)

        top.addWidget(QLabel("Desired Output:"), 0, 0)
        self.metric_combo = QComboBox()
        self.metric_combo.addItems(DESIRED_METRIC_OPTIONS)
        self.metric_combo.setFixedWidth(260)
        self.metric_combo.setToolTip("Pick a metric then click Add to include it in the output list.")
        top.addWidget(self.metric_combo, 0, 1, 1, 2)

        add_btn = QPushButton("Add")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.setFixedWidth(90)
        add_btn.setToolTip("Add selected metric to the list (duplicates are ignored).")

        rem_btn = QPushButton("Remove")
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setFixedWidth(90)
        rem_btn.setToolTip("Remove selected metric(s) from the list.")

        up_btn = QPushButton("Move Up")
        up_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        up_btn.setFixedWidth(110)
        up_btn.setToolTip("Move the selected metric up (changes output ordering).")

        down_btn = QPushButton("Move Down")
        down_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        down_btn.setFixedWidth(110)
        down_btn.setToolTip("Move the selected metric down (changes output ordering).")

        btns = QHBoxLayout()
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addSpacing(10)
        btns.addWidget(up_btn)
        btns.addWidget(down_btn)
        btns.addStretch()

        top_wrap = QHBoxLayout()
        top_wrap.addLayout(top)
        top_wrap.addStretch()

        lay.addLayout(top_wrap)
        lay.addLayout(btns)

        self.metrics_list = QListWidget()
        self.metrics_list.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.metrics_list.setStyleSheet("""
            QListWidget {
                background: #ffffff;
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 6px;
                font-size: 11pt;
            }
        """)
        self.metrics_list.setMinimumHeight(160)
        self.metrics_list.setToolTip("Selected metrics (top-to-bottom = output order). Select items to Remove.")
        lay.addWidget(self.metrics_list)

        lay.addWidget(hline())

        env_row = QHBoxLayout()
        env_row.addWidget(QLabel("Environment Variables:"))
        self.env_combo = QComboBox()
        self.env_combo.addItems(["No", "Yes"])
        self.env_combo.setCurrentText("No")
        self.env_combo.setFixedWidth(120)
        self.env_combo.setToolTip("Yes = include emissions/fuel variables (requires them enabled in VISSIM).")
        env_row.addWidget(self.env_combo)
        env_row.addStretch()
        lay.addLayout(env_row)

        self.main_layout.addWidget(g)

        # default: start with one metric (keeps UI usable without forcing user)
        self._add_metric_to_list("Delay (average)")

        add_btn.clicked.connect(self._add_metric_clicked)
        rem_btn.clicked.connect(self._remove_metric_clicked)
        up_btn.clicked.connect(lambda: self._move_metric(-1))
        down_btn.clicked.connect(lambda: self._move_metric(1))

    def _add_metric_to_list(self, metric: str):
        existing = [self.metrics_list.item(i).text() for i in range(self.metrics_list.count())]
        if metric in existing:
            return
        self.metrics_list.addItem(QListWidgetItem(metric))

    def _add_metric_clicked(self):
        self._add_metric_to_list(self.metric_combo.currentText())

    def _remove_metric_clicked(self):
        selected = self.metrics_list.selectedItems()
        for it in selected:
            row = self.metrics_list.row(it)
            self.metrics_list.takeItem(row)

    def _move_metric(self, delta: int):
        row = self.metrics_list.currentRow()
        if row < 0:
            return
        new_row = row + delta
        if new_row < 0 or new_row >= self.metrics_list.count():
            return
        item = self.metrics_list.takeItem(row)
        self.metrics_list.insertItem(new_row, item)
        self.metrics_list.setCurrentRow(new_row)

    # --------------------------
    # E - Results specifics
    # --------------------------
    def _build_results_specifics(self):
        g = QGroupBox("Results Specifics")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(
            g,
            "Run: select Avg/Min/Max/StdDev/Var or run number.\n"
            "Vehicle Type: select the vehicle class.\n"
            "Time Interval: Avg or interval number."
        )

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.run_combo = QComboBox()
        run_options = ["Avg", "Min", "Max", "StdDev", "Var"] + [str(i) for i in range(1, 11)] + ["Other"]
        self.run_combo.addItems(run_options)
        self.run_combo.setCurrentText("Avg")
        self.run_combo.setFixedWidth(140)
        self.run_combo.setToolTip("Which run statistic to extract (Avg/Min/Max/etc.) or a specific run number.")

        self.run_other = QLineEdit()
        self.run_other.setPlaceholderText('If "Other", enter here')
        self.run_other.setFixedWidth(200)
        self.run_other.setEnabled(False)
        self.run_other.setToolTip('Used only when Run = "Other".')
        self.run_combo.currentTextChanged.connect(lambda t: self.run_other.setEnabled(t == "Other"))

        self.veh_combo = QComboBox()
        # ✅ Updated vehicle types (matches your latest standard list)
        self.veh_combo.addItems(["All", "Cars", "HGVs", "Buses", "Motorcycles", "Bikes", "Pedestrians"])
        self.veh_combo.setCurrentText("All")
        self.veh_combo.setFixedWidth(210)
        self.veh_combo.setToolTip("Vehicle class to extract results for.")

        self.ti_combo = QComboBox()
        self.ti_combo.addItems(["Avg"] + [str(i) for i in range(1, 11)] + ["Other"])
        self.ti_combo.setCurrentText("1")
        self.ti_combo.setFixedWidth(140)
        self.ti_combo.setToolTip("Time interval (Avg or interval number).")

        self.ti_other = QLineEdit()
        self.ti_other.setPlaceholderText('If "Other", enter here')
        self.ti_other.setFixedWidth(200)
        self.ti_other.setEnabled(False)
        self.ti_other.setToolTip('Used only when Time Interval = "Other".')
        self.ti_combo.currentTextChanged.connect(lambda t: self.ti_other.setEnabled(t == "Other"))

        grid.addWidget(QLabel("Run:"), 0, 0)
        grid.addWidget(self.run_combo, 0, 1)
        grid.addWidget(self.run_other, 0, 2)

        grid.addWidget(QLabel("Vehicle Type:"), 1, 0)
        grid.addWidget(self.veh_combo, 1, 1)

        grid.addWidget(QLabel("Time Interval:"), 2, 0)
        grid.addWidget(self.ti_combo, 2, 1)
        grid.addWidget(self.ti_other, 2, 2)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    # --------------------------
    # F - Output
    # --------------------------
    def _build_output(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        # HOVER tooltip (no star)
        add_corner_help(g, "Select where to save the network results Excel file.")

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        self.result_dir.setToolTip("Folder where the Excel output will be saved.")

        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.setToolTip("Browse for an output folder.")
        browse.clicked.connect(self._browse_result_dir)

        self.output_file = QLineEdit("Network Results.xlsx")
        self.output_file.setFixedWidth(260)
        self.output_file.setToolTip("Excel file name (e.g., 'Network Results.xlsx').")

        grid.addWidget(QLabel("Results Directory:"), 0, 0)
        grid.addWidget(self.result_dir, 0, 1)
        grid.addWidget(browse, 0, 2)

        grid.addWidget(QLabel("Output File Name:"), 1, 0)
        grid.addWidget(self.output_file, 1, 1)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    # --------------------------
    # Collect + Log
    # --------------------------
    def _build_collect(self):
        row = QHBoxLayout()
        self.collect_btn = QPushButton("Collect Network Results")
        self.collect_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.collect_btn.setToolTip("Run network results collection for the selected scenarios.")
        self.collect_btn.clicked.connect(self._on_collect)
        row.addStretch()
        row.addWidget(self.collect_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.log.setToolTip("Log output.")
        self.main_layout.addWidget(self.log)

    def _on_collect(self):
        # ✅ Gate: Network Results collect requires Tier 1
        if not self._ensure_feature_or_redirect(self.FEATURE_RESULTS_BASIC, "Collect Network Results"):
            return

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Scenario_List: List[Scenario] = []
        for r in self.scenario_rows:
            Scenario_List.append((
                int(r["num"].value()),
                r["name"].text().strip(),
                bool(r["run"].isChecked()),
            ))

        Seeding_Period = int(self.seed_sec.text().strip() or "0")
        eval_period = int(self.eval_sec.text().strip() or "0")
        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_seed_inc.value())

        Desired_Metric: List[str] = [self.metrics_list.item(i).text() for i in range(self.metrics_list.count())]
        Environment_Variables = (self.env_combo.currentText() == "Yes")

        Run = self.run_other.text().strip() if self.run_combo.currentText() == "Other" else self.run_combo.currentText()
        Veh_Type = self.veh_combo.currentText()
        Time_interval = self.ti_other.text().strip() if self.ti_combo.currentText() == "Other" else self.ti_combo.currentText()

        Result_Directory = self.result_dir.text().strip()
        Output = self.output_file.text().strip()

        if Path_ and not os.path.isdir(Path_):
            QMessageBox.warning(self, "Project Path", "Project Path must be a directory.")
            return
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return
        if any(s[2] for s in Scenario_List):
            if eval_period <= Seeding_Period:
                QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
                return
        if len(Desired_Metric) == 0:
            QMessageBox.warning(self, "Collected Variables", "Please add at least one Desired Output metric.")
            return

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        log_line("VissiCaRe Network Results inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Result_Directory", Result_Directory),
            ("Scenario_List", Scenario_List),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
            ("Desired_Metric", Desired_Metric),
            ("Environment_Variables", Environment_Variables),
            ("Veh_Type", Veh_Type),
            ("Run", Run),
            ("Time_interval", Time_interval),
            ("Output", Output),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Network_Results (TEST)...\n")
        res = Network_Results(
            Vissim_Version=Vissim_Version,
            Filename=Filename,
            Result_Directory=Result_Directory,
            Scenario_List=Scenario_List,
            Seeding_Period=Seeding_Period,
            eval_period=eval_period,
            NumRuns=NumRuns,
            Random_Seed=Random_Seed,
            Random_SeedInc=Random_SeedInc,
            Desired_Metric=Desired_Metric,
            Environment_Variables=Environment_Variables,
            Veh_Type=Veh_Type,
            Run=Run,
            Time_interval=Time_interval,
            Output=Output,
        )
        log_line(f"\nNetwork_Results finished. Return: {res}")

    def export_state(self) -> dict:
        metrics = [self.metrics_list.item(i).text() for i in range(self.metrics_list.count())]
        return {
            "project": self.export_project_setup(),
            "scenarios": self.export_scenarios(),
            "sim": {
                "seed_sec": self.seed_sec.text(),
                "eval_sec": self.eval_sec.text(),
                "num_runs": int(self.num_runs.value()),
                "rand_seed": int(self.rand_seed.value()),
                "rand_seed_inc": int(self.rand_seed_inc.value()),
            },
            "collected": {
                "metrics": metrics,
                "env": self.env_combo.currentText(),
            },
            "results": {
                "run_combo": self.run_combo.currentText(),
                "run_other": self.run_other.text(),
                "veh_type": self.veh_combo.currentText(),
                "ti_combo": self.ti_combo.currentText(),
                "ti_other": self.ti_other.text(),
            },
            "output": {
                "result_dir": self.result_dir.text(),
                "output_file": self.output_file.text(),
            },
        }

    def import_state(self, s: dict):
        self.import_project_setup(s.get("project", {}))
        self.import_scenarios(s.get("scenarios", {}))

        sim = s.get("sim", {})
        self.seed_sec.setText(str(sim.get("seed_sec", "1800")))
        self.eval_sec.setText(str(sim.get("eval_sec", "3600")))
        self.num_runs.setValue(int(sim.get("num_runs", 2)))
        self.rand_seed.setValue(int(sim.get("rand_seed", 42)))
        self.rand_seed_inc.setValue(int(sim.get("rand_seed_inc", 10)))

        collected = s.get("collected", {})
        _set_combo_text(self.env_combo, collected.get("env", "No"))

        self.metrics_list.clear()
        for m in collected.get("metrics", []) or []:
            self.metrics_list.addItem(QListWidgetItem(str(m)))

        res = s.get("results", {})
        _set_combo_text(self.run_combo, res.get("run_combo", "Avg"))
        self.run_other.setText(res.get("run_other", ""))
        _set_combo_text(self.veh_combo, res.get("veh_type", "All"))
        _set_combo_text(self.ti_combo, res.get("ti_combo", "1"))
        self.ti_other.setText(res.get("ti_other", ""))

        out = s.get("output", {})
        self.result_dir.setText(out.get("result_dir", ""))
        self.output_file.setText(out.get("output_file", "Network Results.xlsx"))









# ============================================================
# CALIBRATOR TAB (DROP-IN — FULL UPDATED FILE)
#
# ✅ Updates in THIS version (only what you asked NOW):
#  1) "Seeding" (capital S) stays in Calibrated_variables as ("Seeding", False) ALWAYS
#     and is NOT shown in the UI.
#  2) "seeding" (small s) is ALWAYS the tuple: seeding = (0, 0.35, 0.99)
#     and is NOT shown in the UI.
#     Num_Seeding_Intervals is set consistently to 4 (fixed, hidden).
#
# ✅ Keeps everything else as-is.
# ============================================================

from typing import Optional, List, Tuple, Dict, Any
import os

from PyQt6.QtCore import Qt, QTimer, QEvent
from PyQt6.QtGui import QDoubleValidator, QIntValidator, QPainter, QColor, QPainterPath, QPen
from PyQt6.QtWidgets import (
    QWidget, QLineEdit, QVBoxLayout, QScrollArea, QLabel, QGroupBox,
    QGridLayout, QComboBox, QPushButton, QFileDialog, QSpinBox, QHBoxLayout,
    QCheckBox, QTableWidget, QHeaderView, QTextEdit, QMessageBox, QDoubleSpinBox,
    QSizePolicy, QStyle, QStyleOptionButton, QTabWidget
)

# ------------------------------------------------------------
# Safe fallbacks (only used if your main file didn't define them)
# ------------------------------------------------------------
try:
    SECTION_GROUPBOX_STYLE
except NameError:
    SECTION_GROUPBOX_STYLE = """
        QGroupBox { border: 1px solid #c9d4e3; border-radius: 12px; margin-top: 10px; padding: 10px; }
        QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 6px; font-weight: 900; color: #143a5d; }
    """

try:
    SUBTITLE_STYLE
except NameError:
    SUBTITLE_STYLE = "QLabel { font-size: 12pt; font-weight: 900; color: #143a5d; }"

try:
    TAB_BASE_STYLE
except NameError:
    TAB_BASE_STYLE = "background-color: #f6f8fb; QLabel { background: transparent; border: none; }"

try:
    SOFT_BUTTON_STYLE
except NameError:
    SOFT_BUTTON_STYLE = """
        QPushButton { background: #f3f7ff; border: 1px solid #c9d4e3; border-radius: 10px; padding: 8px 12px; font-weight: 800; color: #143a5d; }
        QPushButton:hover { background: #e8f1ff; }
    """

try:
    PRIMARY_BUTTON_STYLE
except NameError:
    PRIMARY_BUTTON_STYLE = """
        QPushButton { background: #2f6fed; color: white; border: 0px; border-radius: 12px; padding: 10px 18px; font-weight: 900; font-size: 12pt; }
        QPushButton:hover { background: #245fd3; }
    """

try:
    LOG_STYLE
except NameError:
    LOG_STYLE = "QTextEdit { background: #ffffff; border: 1px solid #c9d4e3; border-radius: 12px; padding: 10px; font-size: 11pt; }"

try:
    GLOBAL_SCROLLBAR_STYLE
except NameError:
    GLOBAL_SCROLLBAR_STYLE = ""


def hline():
    w = QWidget()
    w.setFixedHeight(1)
    w.setStyleSheet("background: #c9d4e3;")
    return w


# ============================================================
# ✅ Apply the same "Mistar Solution" checkbox look
#    - transparent background (no grey strip)
#    - black, bold font
#    - WHITE indicator background
# ============================================================
def _apply_mistar_checkbox_look(cb: QCheckBox):
    cb.setAutoFillBackground(False)
    cb.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
    cb.setStyleSheet(
        (cb.styleSheet() or "") +
        """
        QCheckBox {
            background: transparent;
            border: none;
            padding: 0px;
            margin: 0px;
            color: #000000;
            font-weight: 700;
        }
        QCheckBox::indicator {
            width: 18px;
            height: 18px;
            background: #ffffff;
            border: 1px solid #8aa1c1;
            border-radius: 4px;
        }
        QCheckBox::indicator:disabled {
            background: #ffffff;
            border: 1px solid #c9d4e3;
        }
        """
    )


# ============================================================
# ✅ HOVER-ONLY help (NO star widget is shown)
# ============================================================
def _help_star(help_text: str) -> QLabel:
    lbl = QLabel("")
    lbl.setFixedWidth(0)
    lbl.setToolTip(help_text or "")
    return lbl


def _label_with_help(text: str, help_text: Optional[str] = None) -> QWidget:
    lbl = QLabel(text)
    # ✅ make labels truly transparent (fixes "Hyperparameter" grey strip look)
    lbl.setAutoFillBackground(False)
    lbl.setStyleSheet("QLabel { background: transparent; border: none; color: #000000; }")
    if help_text:
        lbl.setToolTip(help_text)
    return lbl


# ============================================================
# Blue checkmark checkbox (indicator checkmark like your image)
# ============================================================
try:
    BlueCheckBox
except NameError:
    class BlueCheckBox(QCheckBox):
        """
        QCheckBox that always shows a visible checkmark when checked (✓),
        with a blue check mark (matches your provided image, but blue).
        """
        def paintEvent(self, event):
            painter = QPainter(self)
            painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

            opt = QStyleOptionButton()
            self.initStyleOption(opt)

            # Draw checkbox WITHOUT the default checked glyph (we'll draw our own)
            is_on = bool(opt.state & QStyle.StateFlag.State_On)
            if is_on:
                opt.state &= ~QStyle.StateFlag.State_On
                opt.state |= QStyle.StateFlag.State_Off

            self.style().drawControl(QStyle.ControlElement.CE_CheckBox, opt, painter, self)

            # Draw our blue checkmark when checked
            if self.isChecked():
                ind = self.style().subElementRect(QStyle.SubElement.SE_CheckBoxIndicator, opt, self)
                r = ind.adjusted(3, 3, -3, -3)

                blue = QColor("#0b5ed7") if self.isEnabled() else QColor("#9fb3cf")
                pen = QPen(blue, 2.6, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)
                painter.setPen(pen)

                path = QPainterPath()
                x = float(r.x()); y = float(r.y())
                w = float(r.width()); h = float(r.height())

                path.moveTo(x + 0.10 * w, y + 0.55 * h)
                path.lineTo(x + 0.42 * w, y + 0.85 * h)
                path.lineTo(x + 0.95 * w, y + 0.15 * h)

                painter.drawPath(path)

            painter.end()


# ============================================================
# PARAMS ORDER + BOOLEAN PARAMS
# ============================================================
VISSIM_PARAMS_ORDERED = [
    "LookAheadDistMax", "LookAheadDistMin", "LookBackDistMax", "LookBackDistMin",
    "NumInteractVeh", "NumInteractObj",
    "RecovSlow", "RecovSpeed", "RecovAcc", "RecovSafDist", "RecovDist",
    "StandDistIsFix", "StandDist",
    "W99cc0", "W99cc1Distr", "W99cc2", "W99cc3", "W99cc4", "W99cc5", "W99cc6", "W99cc7", "W99cc8", "W99cc9",
    "W74ax", "W74bxAdd", "W74bxMult",
    "ACC_StandSafDist", "ACC_MinGapTime", "MinFrontRearClear",
    "MaxDecelOwn", "MaxDecelTrail", "AccDecelOwn", "AccDecelTrail", "DecelRedDistOwn", "DecelRedDistTrail", "CoopDecel",
    "VehRoutDecLookAhead", "Zipper", "ZipperMinSpeed", "CoopLnChg", "CoopLnChgSpeedDiff", "CoopLnChgCollTm",
    "ObsrvAdjLn", "DiamQueu", "ConsNextTurn", "MinCollTmGain", "MinSpeedForLat",
    "EnforcAbsBrakDist", "UseImplicStoch", "PlatoonPoss", "PlatoonMinClear", "MaxNumPlatoonVeh",
    "MaxPlatoonDesSpeed", "MaxPlatoonApprDist", "PlatoonFollowUpGapTm",
    "MesoReactTm", "MesoStandDist", "MesoMaxWaitTm",
]

BOOLEAN_DB_PARAMS = {
    "StandDistIsFix", "VehRoutDecLookAhead", "Zipper", "CoopLnChg", "ObsrvAdjLn", "DiamQueu",
    "ConsNextTurn", "EnforcAbsBrakDist", "UseImplicStoch", "PlatoonPoss", "RecovSlow",
}


def _hover(long_name: str, default: str = "—", vmin: str = "—", vmax: str = "—", notes: str = "") -> str:
    s = f"{long_name}\nDefault: {default}\nMin: {vmin}\nMax: {vmax}"
    if notes.strip():
        s += f"\nNotes: {notes.strip()}"
    return s


VISSIM_PARAM_HOVER: Dict[str, str] = {
    "LookAheadDistMax": _hover("Look ahead distance max"),
    "LookAheadDistMin": _hover("Look ahead distance min"),
    "LookBackDistMax": _hover("Look back distance max"),
    "LookBackDistMin": _hover("Look back distance min"),
    "NumInteractVeh": _hover("Number of interaction vehicles", default="2", vmin="1", vmax="50"),
    "NumInteractObj": _hover("Number of interaction objects", default="2", vmin="1", vmax="50"),
    "RecovSlow": _hover("Recovery slow (recovery behavior enable/slow mode)"),
    "RecovSpeed": _hover("Recovery threshold speed"),
    "RecovAcc": _hover("Recovery acceleration"),
    "RecovSafDist": _hover("Recovery safety distance", default="110%"),
    "RecovDist": _hover("Recovery distance"),
    "StandDistIsFix": _hover("Standstill distance is fixed (flag)"),
    "StandDist": _hover("Standstill distance"),
    "W99cc0": _hover("CC0 Standstill distance", default="1.50", vmin="0.10", vmax="4.00"),
    "W99cc1Distr": _hover(
        "CC1 Headway time (Time Distribution)",
        default="0.90", vmin="distribution MIN", vmax="distribution MAX",
        notes=(
            "Min/Max are the time distribution numbers minimum and maximum values; "
            "user must set these distributions manually to meet the min/max required by VISSIM. "
            "The Calibrator will call the Distribution number, not the Time Value."
        ),
    ),
    "W99cc2": _hover("CC2 Following variation", default="4.00", vmin="1.00", vmax="10.00"),
    "W99cc3": _hover("CC3 Threshold for entering following", default="-8.00", vmin="-20.00", vmax="0.00"),
    "W99cc4": _hover("CC4 Negative following threshold", default="-0.35", vmin="-1.00", vmax="0.00"),
    "W99cc5": _hover("CC5 Positive following threshold", default="0.35", vmin="0.00", vmax="1.00"),
    "W99cc6": _hover("CC6 Speed dependency of oscillation", default="11.40", vmin="0.00", vmax="20.00"),
    "W99cc7": _hover("CC7 Oscillation acceleration", default="0.25", vmin="0.00", vmax="1.00"),
    "W99cc8": _hover("CC8 Desired acceleration from standstill", default="3.50", vmin="0.00", vmax="7.00"),
    "W99cc9": _hover("CC9 Desired acceleration at 80 km/h", default="1.50", vmin="0.00", vmax="7.00"),
    "W74ax": _hover("Wiedemann 74 ax (average standstill distance)", default="6.56 ft", vmin="3.28 ft", vmax="9.84 ft"),
    "W74bxAdd": _hover("Wiedemann 74 bx_add (additive part of safety distance)", default="2.00 ft", vmin="1.00 ft", vmax="3.75 ft"),
    "W74bxMult": _hover("Wiedemann 74 bx_mult (multiplicative part of safety distance)", default="3.00", vmin="2.00", vmax="4.75"),
    "ACC_StandSafDist": _hover("ACC standstill safety distance"),
    "ACC_MinGapTime": _hover("ACC minimum gap time"),
    "MinFrontRearClear": _hover("Minimum front/rear clearance"),
    "MaxDecelOwn": _hover("Maximum deceleration (own vehicle)"),
    "MaxDecelTrail": _hover("Maximum deceleration (trailing vehicle)"),
    "AccDecelOwn": _hover("Acceleration/deceleration (own vehicle)"),
    "AccDecelTrail": _hover("Acceleration/deceleration (trailing vehicle)"),
    "DecelRedDistOwn": _hover("Deceleration reduction distance (own vehicle)"),
    "DecelRedDistTrail": _hover("Deceleration reduction distance (trailing vehicle)"),
    "CoopDecel": _hover("Cooperative deceleration"),
    "VehRoutDecLookAhead": _hover("Vehicle routing decision look-ahead"),
    "Zipper": _hover("Zipper merging (flag/setting)"),
    "ZipperMinSpeed": _hover("Zipper merging minimum speed"),
    "CoopLnChg": _hover("Cooperative lane change (flag/setting)"),
    "CoopLnChgSpeedDiff": _hover("Cooperative lane change speed difference"),
    "CoopLnChgCollTm": _hover("Cooperative lane change collision time"),
    "ObsrvAdjLn": _hover("Observed adjacent lane (setting)"),
    "DiamQueu": _hover("Diameter of queue / queue diameter"),
    "ConsNextTurn": _hover("Consider next turn (flag/setting)"),
    "MinCollTmGain": _hover("Minimum collision time gain"),
    "MinSpeedForLat": _hover("Minimum speed for lateral movement"),
    "EnforcAbsBrakDist": _hover("Enforce absolute braking distance"),
    "UseImplicStoch": _hover("Use implicit stochastics (flag)"),
    "PlatoonPoss": _hover("Platooning possible (flag/setting)"),
    "PlatoonMinClear": _hover("Platoon minimum clearance"),
    "MaxNumPlatoonVeh": _hover("Maximum number of platoon vehicles"),
    "MaxPlatoonDesSpeed": _hover("Maximum platoon desired speed"),
    "MaxPlatoonApprDist": _hover("Maximum platoon approach distance"),
    "PlatoonFollowUpGapTm": _hover("Platoon follow-up gap time"),
    "MesoReactTm": _hover("Mesoscopic reaction time"),
    "MesoStandDist": _hover("Mesoscopic standstill distance"),
    "MesoMaxWaitTm": _hover("Mesoscopic maximum wait time"),
}


# -----------------------------
# Numeric-only line edit
# -----------------------------
class NumericLineEdit(QLineEdit):
    def __init__(self, text: str = "", allow_negative: bool = True, allow_float: bool = True, parent: Optional[QWidget] = None):
        super().__init__(text, parent)
        self._allow_negative = allow_negative
        self._allow_float = allow_float

        dv = QDoubleValidator(self)
        dv.setNotation(QDoubleValidator.Notation.StandardNotation)
        self.setValidator(dv)

        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.textChanged.connect(self._sanitize)

    def _sanitize(self):
        t = self.text()
        if t == "":
            return

        out = []
        dot_used = False
        sign_used = False

        for i, ch in enumerate(t):
            if ch.isdigit():
                out.append(ch)
                continue
            if ch == "-" and self._allow_negative and not sign_used and i == 0:
                out.append(ch)
                sign_used = True
                continue
            if ch == "." and self._allow_float and not dot_used:
                out.append(ch)
                dot_used = True
                continue

        cleaned = "".join(out)

        if cleaned in ("-", ".", "-."):
            self.setText(cleaned)
            return

        if cleaned != t:
            self.setText(cleaned)

    def as_number(self) -> Optional[float]:
        t = self.text().strip()
        if t == "" or t in ("-", ".", "-."):
            return None
        try:
            return float(t)
        except ValueError:
            return None


# ============================================================
# Calibrator Tab
# ============================================================
class CalibratorTab(QWidget):
    # Canonical labels your server/license tab uses (NO spaces)
    CAN_TIER2 = "Tier2"
    CAN_TIER2M = "Tier2_M"
    CAN_TIER3 = "Tier3"

    # Gate-feature names (in case server returns these instead of tiers)
    FEAT_CALIBRATOR = "calibrator_basic"
    FEAT_MULTI_SCEN = "multi_scenario_calibration"
    FEAT_ALL_ACCESS = "all_access"

    def __init__(self, parent=None):
        super().__init__(parent)

        self.setStyleSheet(TAB_BASE_STYLE)
        self._autofit_tables: List[QTableWidget] = []

        outer = QVBoxLayout(self)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        # ✅ Apply your existing modern scrollbar style
        if GLOBAL_SCROLLBAR_STYLE.strip():
            scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        self.main_layout = QVBoxLayout(container)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignTop)
        self.main_layout.setSpacing(16)
        self.main_layout.setContentsMargins(12, 12, 12, 24)

        title = QLabel("Vissim Calibrator")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setStyleSheet("""
            QLabel {
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 10px;
                font-size: 18px;
                font-weight: 900;
                color: #143a5d;
            }
        """)
        self.main_layout.addWidget(title)

        self._build_project_setup()          # A
        self._build_calibrated_variables()   # B
        self._build_tuning_parameters()      # C
        self._build_validation_variables()   # D1
        self._build_validation_data()        # D2
        self._build_error_penalties()        # D3
        self._build_simulation_parameters()  # E
        self._build_output()                 # F
        self._build_run()                    # G

        self.main_layout.addStretch()
        scroll.setWidget(container)
        outer.addWidget(scroll)

        # ✅ Enforce multi-scenario gate on startup
        self._enforce_multi_scenarios_license()

    # ============================================================
    # ✅ License helpers (FIXED: NO-space / underscore-safe)
    # ============================================================
    @staticmethod
    def _norm(s: str) -> str:
        """Normalize tier/feature strings so 'Tier 2_M' == 'Tier2_M' == 'tier2m' etc."""
        s = str(s or "")
        s = s.strip().lower()
        # keep underscore but remove spaces and hyphens
        s = s.replace(" ", "").replace("-", "")
        return s

    def _find_license_tab(self):
        """
        Robustly locate your LicenseTab instance.
        Works with:
          - win.license_tab / win.lic_tab / etc
          - or any tab in QTabWidget labeled 'License'
        """
        win = self.window()

        # Common attribute names on main window
        for attr in ("license_tab", "lic_tab", "tab_license", "licenseTab", "license_widget"):
            tab = getattr(win, attr, None)
            if tab is not None:
                return tab

        # Search QTabWidgets for a tab text containing "license"
        try:
            for tabs in win.findChildren(QTabWidget):
                for i in range(tabs.count()):
                    txt = (tabs.tabText(i) or "").strip().lower()
                    if "licen" in txt:
                        w = tabs.widget(i)
                        if w is not None:
                            return w
        except Exception:
            pass

        return None

    def _license_refresh_if_possible(self):
        """
        If your license tab supports validate_and_cache(), call it to refresh from server.
        Safe if not present.
        """
        lic = self._find_license_tab()
        if lic is None:
            return
        m = getattr(lic, "validate_and_cache", None)
        if callable(m):
            try:
                m(force=True)
            except Exception:
                pass

    def _has_any_entitlement(self, candidates: List[str]) -> bool:
        """
        Checks both:
          - tier labels returned by server: Tier2, Tier2_M, Tier3
          - feature labels your gates might use: calibrator_basic, multi_scenario_calibration, all_access
        """
        lic = self._find_license_tab()
        if lic is None:
            return False

        # Prefer LicenseTab.has_feature() (best source of truth)
        hf = getattr(lic, "has_feature", None)
        if callable(hf):
            for name in candidates:
                try:
                    if bool(hf(name)):
                        return True
                except Exception:
                    pass

        # Fallback: read lic._features list directly if exists
        feats = getattr(lic, "_features", None)
        if isinstance(feats, list):
            feats_norm = {self._norm(x) for x in feats}
            for name in candidates:
                if self._norm(name) in feats_norm:
                    return True

        return False

    def _has_tier2_or_better(self) -> bool:
        return self._has_any_entitlement([
            self.CAN_TIER3, self.FEAT_ALL_ACCESS,
            self.CAN_TIER2, "tier 2", "tier2",
            self.FEAT_CALIBRATOR,
        ])

    def _has_tier2m_or_better(self) -> bool:
        return self._has_any_entitlement([
            self.CAN_TIER3, self.FEAT_ALL_ACCESS,
            self.CAN_TIER2M, "tier 2_m", "tier2_m", "tier2m",
            self.FEAT_MULTI_SCEN,
        ])

    def _goto_license_tab(self) -> bool:
        win = self.window()

        # If main window has an explicit method, use it
        for meth_name in ("show_license_tab", "go_to_license_tab", "goto_license_tab", "open_license_tab", "switch_to_license"):
            m = getattr(win, meth_name, None)
            if callable(m):
                try:
                    m()
                    return True
                except Exception:
                    pass

        # Otherwise, find a QTabWidget and select the tab containing "license"
        try:
            for tabs in win.findChildren(QTabWidget):
                for i in range(tabs.count()):
                    txt = (tabs.tabText(i) or "").strip().lower()
                    if "licen" in txt:
                        tabs.setCurrentIndex(i)
                        return True
        except Exception:
            pass

        return False

    def _send_license_message(self, msg: str):
        win = self.window()

        # direct method on main
        for meth_name in ("set_license_message", "show_license_message", "license_message", "push_license_message"):
            m = getattr(win, meth_name, None)
            if callable(m):
                try:
                    m(msg)
                    return
                except Exception:
                    pass

        # try license tab object
        lic = self._find_license_tab()
        if lic is None:
            return

        for meth_name in ("show_status", "set_message", "show_message", "set_status", "set_banner", "set_info"):
            m = getattr(lic, meth_name, None)
            if callable(m):
                try:
                    m(msg)
                    return
                except Exception:
                    pass

    def _redirect_to_license_tab(self, message: str):
        try:
            self._send_license_message(message)
        except Exception:
            pass

        QMessageBox.information(self, "License Required", message)
        self._goto_license_tab()

    def _enforce_multi_scenarios_license(self):
        """
        ✅ Rule: Unless Tier2_M (or Tier3), Multiple Scenarios is ALWAYS 'No'
        """
        if not hasattr(self, "multiple_scenarios_combo"):
            return

        # refresh first (so user just activated and came back)
        self._license_refresh_if_possible()

        if not self._has_tier2m_or_better():
            self.multiple_scenarios_combo.blockSignals(True)
            self.multiple_scenarios_combo.setCurrentText("No")
            self.multiple_scenarios_combo.blockSignals(False)
            self.multiple_scenarios_combo.setEnabled(False)
            self.multiple_scenarios_combo.setToolTip(
                "Multiple Scenarios requires Tier2_M (or Tier3).\n"
                "Activate Tier2_M in the License tab."
            )
        else:
            self.multiple_scenarios_combo.setEnabled(True)

    def showEvent(self, e):
        super().showEvent(e)
        # Whenever this tab becomes visible, re-enforce (user might have just activated)
        QTimer.singleShot(0, self._enforce_multi_scenarios_license)

    # ============================================================
    # ✅ Table auto-fit helpers
    # ============================================================
    def _configure_table_autofit(self, table: QTableWidget):
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        if table not in self._autofit_tables:
            self._autofit_tables.append(table)
            table.viewport().installEventFilter(self)

        self._refit_table_columns(table)

    def eventFilter(self, obj, event):
        if event.type() == QEvent.Type.Resize:
            for t in getattr(self, "_autofit_tables", []):
                if obj is t.viewport():
                    self._refit_table_columns(t)
                    break
        return super().eventFilter(obj, event)

    def _refit_table_columns(self, table: QTableWidget):
        def do_refit():
            if table is None:
                return

            table.resizeColumnsToContents()

            visible_cols = [c for c in range(table.columnCount()) if not table.isColumnHidden(c)]
            for c in visible_cols:
                table.setColumnWidth(c, table.columnWidth(c) + 8)

            avail = table.viewport().width()
            total = sum(table.columnWidth(c) for c in visible_cols)

            if visible_cols and total > 0 and total < avail:
                extra = avail - total

                adds = []
                used = 0
                for c in visible_cols:
                    w = table.columnWidth(c)
                    add = int(extra * (w / total))
                    adds.append(add)
                    used += add

                rem = extra - used
                i = 0
                while rem > 0:
                    adds[i % len(adds)] += 1
                    rem -= 1
                    i += 1

                for c, add in zip(visible_cols, adds):
                    table.setColumnWidth(c, table.columnWidth(c) + add)

        QTimer.singleShot(0, do_refit)

    # -------------------------
    # A) Project Setup
    # -------------------------
    def _build_project_setup(self):
        g = QGroupBox("Project Setup")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        grid.addWidget(_label_with_help("VISSIM Version:", "Select the version of VISSIM you are using for this project."), 0, 0)
        self.vissim_version_combo = QComboBox()
        for v in range(21, 27):
            self.vissim_version_combo.addItem(str(v))
        self.vissim_version_combo.setCurrentText("26")
        self.vissim_version_combo.setEditable(True)
        self.vissim_version_combo.lineEdit().setValidator(QIntValidator(0, 999, self))
        self.vissim_version_combo.setFixedWidth(90)
        grid.addWidget(self.vissim_version_combo, 0, 1, Qt.AlignmentFlag.AlignLeft)

        grid.addWidget(_label_with_help("Project File (.inpx):", "Select the VISSIM .inpx file for this project."), 1, 0)
        self.project_file_edit = QLineEdit()
        self.project_file_edit.setReadOnly(True)
        self.project_file_edit.setFixedWidth(420)
        grid.addWidget(self.project_file_edit, 1, 1)

        browse = QPushButton("Browse…")
        browse.setStyleSheet(SOFT_BUTTON_STYLE)
        browse.clicked.connect(self._browse_project_path)
        grid.addWidget(browse, 1, 2)

        self.project_path_edit = QLineEdit()
        self.project_path_edit.setVisible(False)

        self.filename_edit = QLineEdit()
        self.filename_edit.setVisible(False)

        grid.addWidget(_label_with_help(
            "Multiple Scenarios:",
            "If Yes: you can enter different scenario numbers per row in Validation Data.\n"
            "If No: all rows must use the same Scenario number."
        ), 2, 0)
        self.multiple_scenarios_combo = QComboBox()
        self.multiple_scenarios_combo.addItems(["No", "Yes"])
        self.multiple_scenarios_combo.setCurrentText("No")
        self.multiple_scenarios_combo.setFixedWidth(120)
        grid.addWidget(self.multiple_scenarios_combo, 2, 1, Qt.AlignmentFlag.AlignLeft)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    def _browse_project_path(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select VISSIM Project File",
            "",
            "VISSIM Project (*.inpx)",
        )
        if file_path:
            # keep exactly as you had it (assumes Path imported in your main file)
            p = Path(file_path)  # type: ignore[name-defined]
            self.project_file_edit.setText(str(p))
            self.project_path_edit.setText(str(p.parent))
            self.filename_edit.setText(p.stem)

    # -------------------------
    # B) Calibrated Variables
    # -------------------------
    def _build_calibrated_variables(self):
        g = QGroupBox("Calibrated Variables")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        title1 = QLabel("Driving Behaviour Parameters")
        title1.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title1)

        self.db_blocks: List[dict] = []
        self.db_container = QVBoxLayout()
        self.db_container.setSpacing(10)
        lay.addLayout(self.db_container)

        btns = QHBoxLayout()
        add_db = QPushButton("Add DB")
        rem_db = QPushButton("Remove Last DB")
        add_db.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_db.setStyleSheet(SOFT_BUTTON_STYLE)
        add_db.clicked.connect(self._add_db_block_btn)
        rem_db.clicked.connect(self._remove_last_db_block)
        btns.addWidget(add_db)
        btns.addWidget(rem_db)
        btns.addStretch()
        lay.addLayout(btns)

        lay.addWidget(hline())

        title2 = QLabel("Additional Variables")
        title2.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(title2)

        row = QHBoxLayout()
        self.chk_dsd = BlueCheckBox("Speed Distribution")
        self.chk_lcd = BlueCheckBox("Lane Change Distance")

        # ✅ "Seeding" (capital S) ALWAYS False, hidden from UI
        self.chk_seed = BlueCheckBox("Seeding")
        self.chk_seed.setChecked(False)
        self.chk_seed.setEnabled(False)
        self.chk_seed.setVisible(False)

        self.chk_dsd.setToolTip("The calibrator can calibrate the Desired Speed Distributions. You will be required to enter the information for each DSD you wish to calibrate. See the help page for more details on the process of calibrating a DSD.")
        self.chk_lcd.setToolTip("The calibrator can calibrate Lane Change Distances. You will be required to enter the information for each LCD you wish to calibrate. See the help page for more details on the process of calibrating a LCD.")
        self.chk_seed.setToolTip("")

        # ✅ Match "Mistar Solution" look
        _apply_mistar_checkbox_look(self.chk_dsd)
        _apply_mistar_checkbox_look(self.chk_lcd)
        _apply_mistar_checkbox_look(self.chk_seed)

        for cb in (self.chk_dsd, self.chk_lcd):
            cb.stateChanged.connect(self._update_connected_sections_visibility)
            row.addWidget(cb)
        row.addStretch()
        lay.addLayout(row)

        self.connected_section = QGroupBox("Connected Inputs (shows only when selected above)")
        self.connected_section.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.connected_section.setVisible(False)  # ✅ hide by default
        cs_lay = QVBoxLayout(self.connected_section)

        self.dsd_box = QGroupBox("Desired Speed Decisions (DSD)")
        self.dsd_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.dsd_box.setToolTip("Select the Desired Speed Decisions to calibrate. You should enter the number of this DSD from the distributions list, the lower bound value that the calibrator will not add or subtract below from the points values of that DSD, the upper value, and the start point where the calibrator will start by adding or subtracting on the first Trial.")
        dsd_lay = QVBoxLayout(self.dsd_box)

        self.dsd_table = self._make_triplet_table(headers=["DSD Number", "Lower Bound", "Start Point", "Upper Bound"])
        self._configure_table_autofit(self.dsd_table)
        dsd_lay.addWidget(self.dsd_table)
        dsd_lay.addLayout(self._make_table_buttons(self.dsd_table))

        self.lcd_box = QGroupBox("Lane Change Distances (LCD)")
        self.lcd_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.lcd_box.setToolTip("Select the Lane Change Distances to calibrate. You should enter the number of this LCD from the distributions list, the lower bound value that the calibrator will not add or subtract below from the points values of that LCD, the upper value, and the start point where the calibrator will start by adding or subtracting on the first Trial. The calibrator will now allow the lower value to go below zero.")
        lcd_lay = QVBoxLayout(self.lcd_box)

        self.lcd_table = self._make_triplet_table(headers=["LCD Number", "Lower Bound", "Start Point", "Upper Bound"])
        self._configure_table_autofit(self.lcd_table)
        lcd_lay.addWidget(self.lcd_table)
        lcd_lay.addLayout(self._make_table_buttons(self.lcd_table))

        # ✅ seeding (small s) fixed tuple, hidden from UI (we keep widgets for compatibility but NEVER show)
        self.seed_box = QGroupBox("seeding (fixed / hidden)")
        self.seed_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.seed_box.setToolTip("")
        seed_lay = QVBoxLayout(self.seed_box)

        sgrid = QGridLayout()
        sgrid.setHorizontalSpacing(10)
        sgrid.setVerticalSpacing(10)

        # fixed defaults that match your required tuple
        self.seed_lb = QDoubleSpinBox(); self.seed_lb.setRange(0.0, 1.0); self.seed_lb.setDecimals(6); self.seed_lb.setSingleStep(0.01); self.seed_lb.setValue(0.00)
        self.seed_sp = QDoubleSpinBox(); self.seed_sp.setRange(0.0, 1.0); self.seed_sp.setDecimals(6); self.seed_sp.setSingleStep(0.01); self.seed_sp.setValue(0.35)
        self.seed_ub = QDoubleSpinBox(); self.seed_ub.setRange(0.0, 1.0); self.seed_ub.setDecimals(6); self.seed_ub.setSingleStep(0.01); self.seed_ub.setValue(0.99)

        # fixed intervals to match your schema
        self.seed_intervals = QSpinBox(); self.seed_intervals.setRange(1, 999); self.seed_intervals.setValue(4); self.seed_intervals.setFixedWidth(90)

        # disable (still hidden anyway)
        for w in (self.seed_lb, self.seed_sp, self.seed_ub, self.seed_intervals):
            w.setEnabled(False)

        sgrid.addWidget(_label_with_help("Lower Bound:", "Hidden / fixed."), 0, 0); sgrid.addWidget(self.seed_lb, 0, 1)
        sgrid.addWidget(_label_with_help("Start Point:", "Hidden / fixed."), 0, 2); sgrid.addWidget(self.seed_sp, 0, 3)
        sgrid.addWidget(_label_with_help("Upper Bound:", "Hidden / fixed."), 0, 4); sgrid.addWidget(self.seed_ub, 0, 5)
        sgrid.addWidget(_label_with_help("Number of Seeding Intervals:", "Hidden / fixed."), 1, 0, 1, 2)
        sgrid.addWidget(self.seed_intervals, 1, 2)

        swrap = QHBoxLayout()
        swrap.addLayout(sgrid)
        swrap.addStretch()
        seed_lay.addLayout(swrap)

        cs_lay.addWidget(self.dsd_box)
        cs_lay.addWidget(self.lcd_box)
        cs_lay.addWidget(self.seed_box)

        lay.addWidget(self.connected_section)
        self.main_layout.addWidget(g)

        self._add_db_block(default_db=1, default_model="WIEDEMANN74")

        # ✅ ALWAYS hidden (both Seeding + seeding)
        self.seed_box.setVisible(False)
        self._update_connected_sections_visibility()

        self._table_add_row(self.dsd_table)
        self._table_add_row(self.lcd_table)

        self._refit_table_columns(self.dsd_table)
        self._refit_table_columns(self.lcd_table)

    def _add_db_block_btn(self):
        i = len(self.db_blocks) + 1
        self._add_db_block(default_db=i, default_model="WIEDEMANN74")

    def _add_db_block(self, default_db: int, default_model: str):
        box = QGroupBox(f"DB #{default_db}")
        box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(box)

        top = QHBoxLayout()
        top.addWidget(_label_with_help("Driving Behaviour (DB) Number:", "Select the DB number to calibrate."))
        db_spin = QSpinBox()
        db_spin.setRange(1, 999999)
        db_spin.setValue(default_db)
        db_spin.setFixedWidth(110)
        top.addWidget(db_spin)

        top.addSpacing(20)
        top.addWidget(_label_with_help("Car Following Model:", "Select the car-following model for this DB."))
        model = QComboBox()
        model.addItems(["WIEDEMANN74", "WIEDEMANN99"])
        model.setCurrentText(default_model if default_model in ["WIEDEMANN74", "WIEDEMANN99"] else "WIEDEMANN74")
        model.setFixedWidth(160)
        top.addWidget(model)

        top.addStretch()
        lay.addLayout(top)

        sub = QLabel("Hyperparameters")
        sub.setStyleSheet(SUBTITLE_STYLE)
        lay.addWidget(sub)

        params_container = QVBoxLayout()
        params_container.setSpacing(8)
        lay.addLayout(params_container)

        btns = QHBoxLayout()
        add_p = QPushButton("Add Parameter")
        rem_p = QPushButton("Remove Last Parameter")
        add_p.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_p.setStyleSheet(SOFT_BUTTON_STYLE)
        btns.addWidget(add_p)
        btns.addWidget(rem_p)
        btns.addStretch()
        lay.addLayout(btns)

        block = {
            "widget": box,
            "db": db_spin,
            "model": model,
            "params_container": params_container,
            "params": [],
        }

        add_p.clicked.connect(lambda: self._add_param_row(block))
        rem_p.clicked.connect(lambda: self._remove_param_row(block))

        self._add_param_row(block)

        wrap = QHBoxLayout()
        wrap.addWidget(box)
        wrap.addStretch()
        self.db_container.addLayout(wrap)

        self.db_blocks.append(block)

    def _remove_last_db_block(self):
        if not self.db_blocks:
            return
        b = self.db_blocks.pop()
        b["widget"].setParent(None)

    def _params_list_for_model(self, model: str) -> List[str]:
        return list(VISSIM_PARAMS_ORDERED)

    def _is_boolean_hp(self, hp_name: str) -> bool:
        return hp_name in BOOLEAN_DB_PARAMS

    def _set_bool_combo_mode(self, combo: QComboBox, mode: str):
        cur_bool = combo.currentData(Qt.ItemDataRole.UserRole)
        if cur_bool is None:
            cur_bool = (combo.currentText() in ("True", "Active"))

        combo.blockSignals(True)
        combo.clear()

        if mode == "active":
            combo.addItem("Not Active", False)
            combo.addItem("Active", True)
            combo.setCurrentIndex(1 if cur_bool else 0)
        else:
            combo.addItem("False", False)
            combo.addItem("True", True)
            combo.setCurrentIndex(1 if cur_bool else 0)

        combo.blockSignals(False)

    def _apply_boolean_param_ui(self, seg: dict, on: bool, bool_mode: str = "truefalse"):
        seg["lb"].setVisible(not on)
        seg["ub"].setVisible(not on)
        seg["lb_label"].setVisible(not on)
        seg["ub_label"].setVisible(not on)

        seg["sp"].setVisible(not on)
        seg["sp_bool"].setVisible(on)

        seg["sp_label"].setText("Value:" if on else "Start Point:")

        if on:
            self._set_bool_combo_mode(seg["sp_bool"], bool_mode)
            seg["lb"].setText("-1")
            seg["ub"].setText("1")
            seg["lb"].setEnabled(False)
            seg["ub"].setEnabled(False)
        else:
            seg["lb"].setEnabled(True)
            seg["ub"].setEnabled(True)

    def _add_param_row(self, block: dict):
        roww = QWidget()
        grid = QGridLayout(roww)
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(6)

        grid.addWidget(_label_with_help(
            "Hyperparameter:",
            "Choose the hyperparameter to calibrate from the droplist. If you wish to calibrate a parameter that does not already exist in that list, select <b>Other</b> and then specify, using <b>Other Type</b> if the new hyperparamter is <b>Numeric</b> or Boolean (<b>Active / Not Active </b>. Make sure you enter the shortname of that hyperparamter exactly as VISSIM has it in the VISSIM COM help page."
        ), 0, 0)

        hp = QComboBox()
        hp.addItems(self._params_list_for_model(block["model"].currentText()) + ["Other"])
        hp.setFixedWidth(220)
        grid.addWidget(hp, 0, 1)

        for i in range(hp.count()):
            name = hp.itemText(i)
            tip = VISSIM_PARAM_HOVER.get(name, "")
            if tip:
                hp.setItemData(i, tip, Qt.ItemDataRole.ToolTipRole)

        other = QLineEdit()
        other.setPlaceholderText("If Other, type short name")
        other.setFixedWidth(220)
        other.setVisible(False)
        grid.addWidget(other, 0, 2)

        other_type_label = QLabel("Other Type:")
        other_type_label.setVisible(False)
        other_type_label.setAutoFillBackground(False)
        other_type_label.setStyleSheet("QLabel { background: transparent; border: none; color: #000000; }")
        grid.addWidget(other_type_label, 1, 1, Qt.AlignmentFlag.AlignRight)

        other_type = QComboBox()
        other_type.addItems(["Numeric", "Active / Not Active"])
        other_type.setCurrentText("Numeric")
        other_type.setFixedWidth(180)
        other_type.setVisible(False)
        grid.addWidget(other_type, 1, 2, Qt.AlignmentFlag.AlignLeft)

        lb_label = QLabel("Lower Bound:")
        lb_label.setAutoFillBackground(False)
        lb_label.setStyleSheet("QLabel { background: transparent; border: none; color: #000000; }")
        grid.addWidget(lb_label, 0, 3)

        lb = NumericLineEdit("", allow_negative=True, allow_float=True)
        lb.setFixedWidth(120)
        grid.addWidget(lb, 0, 4)

        sp_label = QLabel("Start Point:")
        sp_label.setAutoFillBackground(False)
        sp_label.setStyleSheet("QLabel { background: transparent; border: none; color: #000000; }")
        grid.addWidget(sp_label, 0, 5)

        sp = NumericLineEdit("", allow_negative=True, allow_float=True)
        sp.setFixedWidth(120)
        grid.addWidget(sp, 0, 6)

        sp_bool = QComboBox()
        sp_bool.setFixedWidth(120)
        sp_bool.setVisible(False)
        self._set_bool_combo_mode(sp_bool, "truefalse")
        grid.addWidget(sp_bool, 0, 6)

        ub_label = QLabel("Upper Bound:")
        ub_label.setAutoFillBackground(False)
        ub_label.setStyleSheet("QLabel { background: transparent; border: none; color: #000000; }")
        grid.addWidget(ub_label, 0, 7)

        ub = NumericLineEdit("", allow_negative=True, allow_float=True)
        ub.setFixedWidth(120)
        grid.addWidget(ub, 0, 8)

        wrap = QHBoxLayout()
        wrap.addWidget(roww)
        wrap.addStretch()
        block["params_container"].addLayout(wrap)

        seg = {
            "widget": roww,
            "hp": hp,
            "other": other,
            "other_type_label": other_type_label,
            "other_type": other_type,
            "lb_label": lb_label,
            "sp_label": sp_label,
            "ub_label": ub_label,
            "lb": lb,
            "sp": sp,
            "sp_bool": sp_bool,
            "ub": ub,
        }
        block["params"].append(seg)

        def apply_for_current_selection():
            name = hp.currentText()
            is_other = (name == "Other")

            other.setVisible(is_other)
            other_type_label.setVisible(is_other)
            other_type.setVisible(is_other)

            if is_other:
                if other_type.currentText().strip().lower().startswith("active"):
                    self._apply_boolean_param_ui(seg, True, bool_mode="active")
                else:
                    self._apply_boolean_param_ui(seg, False)
                return

            self._apply_boolean_param_ui(seg, self._is_boolean_hp(name), bool_mode="truefalse")

        hp.currentTextChanged.connect(lambda _t: apply_for_current_selection())
        other_type.currentTextChanged.connect(lambda _t: apply_for_current_selection())
        apply_for_current_selection()

    def _remove_param_row(self, block: dict):
        if not block["params"]:
            return
        seg = block["params"].pop()
        seg["widget"].setParent(None)

    def _update_connected_sections_visibility(self):
        show_dsd = self.chk_dsd.isChecked()
        show_lcd = self.chk_lcd.isChecked()

        if hasattr(self, "dsd_box"):
            self.dsd_box.setVisible(show_dsd)
        if hasattr(self, "lcd_box"):
            self.lcd_box.setVisible(show_lcd)

        # ✅ hide/show the whole Connected Inputs section
        if hasattr(self, "connected_section"):
            self.connected_section.setVisible(show_dsd or show_lcd)

        # ✅ "Seeding" ALWAYS hidden + forced False
        if hasattr(self, "chk_seed"):
            self.chk_seed.blockSignals(True)
            self.chk_seed.setChecked(False)
            self.chk_seed.blockSignals(False)
            self.chk_seed.setVisible(False)

        # ✅ "seeding" ALWAYS hidden
        if hasattr(self, "seed_box"):
            self.seed_box.setVisible(False)

    # -------------------------
    # C) Tuning Parameters
    # -------------------------
    def _build_tuning_parameters(self):
        g = QGroupBox("Tuning Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        g.setToolTip("")

        lay = QVBoxLayout(g)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        grid.addWidget(_label_with_help("Variant:", "The amount of change you want to apply to each calibrated parameter. This value will be subject to random selection from a normal distribution with a mean of <b>Variant</b> and standard distribution of 0.06. one value will be used to add a change to each parameter from one sub-trial to another."), 0, 0)
        self.variant_spin = QDoubleSpinBox()
        self.variant_spin.setRange(0.0, 1000000.0)
        self.variant_spin.setDecimals(6)
        self.variant_spin.setSingleStep(0.01)
        self.variant_spin.setValue(0.3)
        self.variant_spin.setFixedWidth(160)
        grid.addWidget(self.variant_spin, 0, 1)

        grid.addWidget(_label_with_help("Variant Update During Simulation:", "If Yes: you can change Variant and a table will show up. In that table, enter the thresholds that you wish the Variant value to be changed after. Choose a Trial number and what the new Variant value to be after this trial."), 0, 2)
        self.variant_update_combo = QComboBox()
        self.variant_update_combo.addItems(["No", "Yes"])
        self.variant_update_combo.setCurrentText("No")
        self.variant_update_combo.setFixedWidth(120)
        grid.addWidget(self.variant_update_combo, 0, 3)

        self.stop_after_trials = BlueCheckBox("Stop after number of trials")
        self.stop_after_trials.setChecked(True)
        self.stop_after_trials.setToolTip(
            "If checked, calibration stops after this many trials. Otherwise, the calibration continues until all the MoEs are satisfied."
        )
        self.stop_after_trials.setAutoFillBackground(False)
        self.stop_after_trials.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground, True)
        self.stop_after_trials.setStyleSheet(
            (self.stop_after_trials.styleSheet() or "") +
            """
            QCheckBox {
                background: transparent;
                border: none;
                padding: 0px;
                margin: 0px;
                font-weight: 400;
                color: #000000;
            }
            QCheckBox::indicator {
                width: 18px;
                height: 18px;
                background: #ffffff;
                border: 1px solid #9fb3cf;
                border-radius: 4px;
            }
            QCheckBox::indicator:hover {
                border: 1px solid #6f8fbf;
            }
            QCheckBox::indicator:disabled {
                background: #ffffff;
                border: 1px solid #c9d4e3;
            }
            """
        )
        grid.addWidget(self.stop_after_trials, 1, 0, 1, 2)

        grid.addWidget(_label_with_help("Number of Trials:", "If <b>Stop after number of trials</b> is checked, then specify when do you wish to stop the calibration."), 1, 2)
        self.trials_spin = QSpinBox()
        self.trials_spin.setRange(5, 999999)
        self.trials_spin.setValue(250)
        self.trials_spin.setFixedWidth(120)
        grid.addWidget(self.trials_spin, 1, 3)

        grid.addWidget(_label_with_help("Number of Sub-Trials:", "Subtrials diversify the Trial value for finer calibration during each Trial."), 2, 0)
        self.subtrials_spin = QSpinBox()
        self.subtrials_spin.setRange(4, 999999)
        self.subtrials_spin.setValue(4)
        self.subtrials_spin.setFixedWidth(120)
        grid.addWidget(self.subtrials_spin, 2, 1)

        grid.addWidget(_label_with_help("Number of Randomly Selected Sub-Trials:", "This is the number is the assurance the calibrator uses to increase the likelihood of generating new and unique Sub-Trials and avoid stucking with the same values of the parameters."), 2, 2)
        self.random_selection_spin = QSpinBox()
        self.random_selection_spin.setRange(1, 999999)
        self.random_selection_spin.setValue(3)
        self.random_selection_spin.setFixedWidth(120)
        grid.addWidget(self.random_selection_spin, 2, 3)

        grid.addWidget(_label_with_help("Auto Stepwise Variant:", "Auto changes to the value of the variant that can be applied if the overall success of all MoE values reaches a specific threshold. Check the help page for more details on how <b>Auto Stepwise Variant</b> works."), 3, 0)
        self.stepwise_active_combo = QComboBox()
        self.stepwise_active_combo.addItem("Not Active", False)
        self.stepwise_active_combo.addItem("Active", True)
        self.stepwise_active_combo.setCurrentIndex(0)
        self.stepwise_active_combo.setFixedWidth(160)
        grid.addWidget(self.stepwise_active_combo, 3, 1)

        grid.addWidget(_label_with_help("% of Accepted Error:", "Threshold percent (1–95%)."), 3, 2)
        self.stepwise_accept_pct = QSpinBox()
        self.stepwise_accept_pct.setRange(1, 95)
        self.stepwise_accept_pct.setValue(50)
        self.stepwise_accept_pct.setSuffix("%")
        self.stepwise_accept_pct.setFixedWidth(120)
        grid.addWidget(self.stepwise_accept_pct, 3, 3)

        grid.addWidget(_label_with_help("Stepwise value (%):", "Percent multiplier applied to Variant (5–95%)."), 3, 4)
        self.stepwise_value_pct = QSpinBox()
        self.stepwise_value_pct.setRange(5, 95)
        self.stepwise_value_pct.setValue(50)
        self.stepwise_value_pct.setSuffix("%")
        self.stepwise_value_pct.setFixedWidth(120)
        grid.addWidget(self.stepwise_value_pct, 3, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.variant_list_box = QGroupBox("Variant List")
        self.variant_list_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        self.variant_list_box.setToolTip("Select the <b>Trial</b> at which you want to change the Variant and the new <b>Variant</b> you want to apply. It will dynamically apply any new <b>Variant</b>.")
        vlay = QVBoxLayout(self.variant_list_box)

        self.variant_table = QTableWidget(0, 2)
        self.variant_table.setHorizontalHeaderLabels(["Trial", "Variant"])
        self.variant_table.verticalHeader().setVisible(False)
        self.variant_table.verticalHeader().setDefaultSectionSize(40)
        self.variant_table.setMinimumHeight(160)
        self.variant_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 6px;
                font-size: 11pt;
            }
            QHeaderView::section {
                background: #f3f7ff;
                border: 1px solid #c9d4e3;
                padding: 6px;
                font-weight: 900;
                color: #143a5d;
            }
            QTableWidget QSpinBox,
            QTableWidget QDoubleSpinBox {
                font-size: 11pt;
                padding: 2px 6px;
            }
        """)
        vlay.addWidget(self.variant_table)
        self._configure_table_autofit(self.variant_table)

        btns = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        rem_btn = QPushButton("Remove Selected Row(s)")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addStretch()
        vlay.addLayout(btns)

        add_btn.clicked.connect(self._variant_add_row)
        rem_btn.clicked.connect(self._variant_remove_selected_rows)

        lay.addWidget(self.variant_list_box)

        self._variant_add_row(trial_default=0, variant_default=self.variant_spin.value())

        self.variant_update_combo.currentTextChanged.connect(self._update_variant_list_visibility)
        self._update_variant_list_visibility()

        self.variant_spin.valueChanged.connect(self._sync_variant_to_row0)
        self.subtrials_spin.valueChanged.connect(self._sync_random_selection_limit)
        self._sync_random_selection_limit()

        self.stepwise_active_combo.currentIndexChanged.connect(self._sync_stepwise_enabled)
        self._sync_stepwise_enabled()

        self.main_layout.addWidget(g)

    def _sync_stepwise_enabled(self):
        active = bool(self.stepwise_active_combo.currentData())
        self.stepwise_accept_pct.setEnabled(active)
        self.stepwise_value_pct.setEnabled(active)

    def _sync_random_selection_limit(self):
        sub = int(self.subtrials_spin.value())
        max_rs = max(1, sub - 1)

        self.random_selection_spin.blockSignals(True)
        self.random_selection_spin.setMaximum(max_rs)
        if int(self.random_selection_spin.value()) > max_rs:
            self.random_selection_spin.setValue(max_rs)
        self.random_selection_spin.blockSignals(False)

    def _update_variant_list_visibility(self):
        enabled = (self.variant_update_combo.currentText().strip().lower() == "yes")
        self.variant_list_box.setVisible(enabled)
        if enabled:
            self._refit_table_columns(self.variant_table)

    def _variant_add_row(self, trial_default: Optional[int] = None, variant_default: Optional[float] = None):
        r = self.variant_table.rowCount()
        self.variant_table.insertRow(r)
        self.variant_table.setRowHeight(r, 40)

        trial = QSpinBox()
        trial.setRange(0, 999999999)
        trial.setAlignment(Qt.AlignmentFlag.AlignCenter)
        trial.setFont(self.variant_table.font())
        trial.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        if trial_default is not None:
            trial.setValue(int(trial_default))
        self.variant_table.setCellWidget(r, 0, trial)

        varv = QDoubleSpinBox()
        varv.setRange(0.0, 1000000.0)
        varv.setDecimals(6)
        varv.setSingleStep(0.01)
        varv.setAlignment(Qt.AlignmentFlag.AlignCenter)
        varv.setFont(self.variant_table.font())
        varv.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        if variant_default is not None:
            varv.setValue(float(variant_default))
        self.variant_table.setCellWidget(r, 1, varv)

        self._refit_table_columns(self.variant_table)

    def _variant_remove_selected_rows(self):
        rows = sorted({idx.row() for idx in self.variant_table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.variant_table.removeRow(r)
        if self.variant_table.rowCount() == 0:
            self._variant_add_row(trial_default=0, variant_default=self.variant_spin.value())
        self._refit_table_columns(self.variant_table)

    def _sync_variant_to_row0(self):
        if self.variant_table.rowCount() <= 0:
            return
        w_trial = self.variant_table.cellWidget(0, 0)
        w_var = self.variant_table.cellWidget(0, 1)
        if isinstance(w_trial, QSpinBox) and isinstance(w_var, QDoubleSpinBox):
            if int(w_trial.value()) == 0:
                w_var.setValue(float(self.variant_spin.value()))
        self._refit_table_columns(self.variant_table)

    def _collect_variant_list(self) -> List[Tuple[int, float]]:
        out: List[Tuple[int, float]] = []
        for r in range(self.variant_table.rowCount()):
            wt = self.variant_table.cellWidget(r, 0)
            wv = self.variant_table.cellWidget(r, 1)
            if not isinstance(wt, QSpinBox) or not isinstance(wv, QDoubleSpinBox):
                continue
            out.append((int(wt.value()), float(wv.value())))
        return out

    # -------------------------
    # D1) Validation Variables
    # -------------------------
    def _build_validation_variables(self):
        g = QGroupBox("Validation Variables")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        row = QHBoxLayout()
        self.val_volume = BlueCheckBox("Volume"); self.val_volume.setChecked(True)
        self.val_volume.setToolTip("Validate the calibrated parameters using <b>Volume</b> observations. Enter <b>Scenario</b>, <b>Node Number (Volume)</b>, <b>Observed Volume</b>, <b>Penalty for Volume</b>, and <b>Accepted Error for Volume (%)</b> in the Validation Data table. Supports multiple <b>Nodes</b> and (if enabled) multiple <b>Scenarios</b>.")

        self.val_speed = BlueCheckBox("Speed"); self.val_speed.setChecked(False)
        self.val_speed.setToolTip("Validate the calibrated parameters using <b>Speed</b> observations. Enter <b>Scenario</b>, <b>Speed Link</b>, <b>Observed Speed</b>, <b>Penalty for Speed</b>, and <b>Accepted Error for Speed (%)</b> in the Validation Data table.")

        self.val_tt = BlueCheckBox("Travel Time"); self.val_tt.setChecked(False)
        self.val_tt.setToolTip("Validate the calibrated parameters using <b>Travel Time</b> observations. Enter <b>Scenario</b>, <b>Travel Time Number</b>, <b>Observed Travel Time</b>, <b>Penalty for Travel Time</b>, and <b>Accepted Error for Travel Time (%)</b> in the Validation Data table.")

        self.val_queue = BlueCheckBox("Queue"); self.val_queue.setChecked(False)
        self.val_queue.setToolTip("Validate the calibrated parameters using <b>Queue</b> observations. Enter <b>Scenario</b>, <b>Node Number (Queue)</b>, <b>Observed Queue</b>, <b>Penalty for Queue</b>, and <b>Accepted Error for Queue (%)</b> in the Validation Data table.")

        _apply_mistar_checkbox_look(self.val_volume)
        _apply_mistar_checkbox_look(self.val_speed)
        _apply_mistar_checkbox_look(self.val_tt)
        _apply_mistar_checkbox_look(self.val_queue)

        for cb in (self.val_volume, self.val_speed, self.val_tt, self.val_queue):
            cb.stateChanged.connect(self._update_validation_table_columns)
            row.addWidget(cb)

        row.addStretch()
        lay.addLayout(row)

        self.main_layout.addWidget(g)

    # -------------------------
    # D2) Validation Data
    # -------------------------
    def _build_validation_data(self):
        g = QGroupBox("Validation Data")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        self.validation_columns: List[Tuple[str, str]] = [
            ("Scenario", "Scenario"),
            ("Node Number (Volume)", "node_numbers"),
            ("Volume", "Volume"),
            ("Penalty for Volume", "Penalty_V"),
            ("Accepted Error for Volume (%)", "Accepted_Error_V"),
            ("Speed Link", "Speed_Link"),
            ("Speed", "Speed"),
            ("Penalty for Speed", "Penalty_S"),
            ("Accepted Error for Speed (%)", "Accepted_Error_S"),
            ("Travel Time Number", "TT_Number"),
            ("Travel Time", "TravelTime"),
            ("Penalty for Travel Time", "Penalty_T"),
            ("Accepted Error for Travel Time (%)", "Accepted_Error_T"),
            ("Node Number (Queue)", "node_numbers2"),
            ("Queue", "Queue"),
            ("Penalty for Queue", "Penalty_Q"),
            ("Accepted Error for Queue (%)", "Accepted_Error_Q"),
        ]

        self.validation_table = QTableWidget(0, len(self.validation_columns))
        self.validation_table.setHorizontalHeaderLabels([h for h, _k in self.validation_columns])
        self.validation_table.verticalHeader().setVisible(False)
        self.validation_table.setMinimumHeight(260)
        self.validation_table.setStyleSheet("""
            QTableWidget {
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 6px;
                font-size: 11pt;
            }
            QHeaderView::section {
                background: #f3f7ff;
                border: 1px solid #c9d4e3;
                padding: 6px;
                font-weight: 900;
                color: #143a5d;
            }
        """)
        lay.addWidget(self.validation_table)
        self._configure_table_autofit(self.validation_table)

        btns = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        rem_btn = QPushButton("Remove Selected Row(s)")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addStretch()
        lay.addLayout(btns)

        add_btn.clicked.connect(lambda: self._validation_add_row())
        rem_btn.clicked.connect(lambda: self._validation_remove_selected_rows())

        self.main_layout.addWidget(g)

        self._validation_add_row()
        self._update_validation_table_columns()

    def _validation_add_row(self):
        r = self.validation_table.rowCount()
        self.validation_table.insertRow(r)
        for c in range(self.validation_table.columnCount()):
            e = NumericLineEdit("", allow_negative=False, allow_float=True)
            self.validation_table.setCellWidget(r, c, e)
        self._refit_table_columns(self.validation_table)

    def _validation_remove_selected_rows(self):
        rows = sorted({idx.row() for idx in self.validation_table.selectedIndexes()}, reverse=True)
        for r in rows:
            self.validation_table.removeRow(r)
        if self.validation_table.rowCount() == 0:
            self._validation_add_row()
        self._refit_table_columns(self.validation_table)

    def _update_validation_table_columns(self):
        if not hasattr(self, "validation_table"):
            return

        show_vol = self.val_volume.isChecked()
        show_spd = self.val_speed.isChecked()
        show_tt = self.val_tt.isChecked()
        show_q = self.val_queue.isChecked()

        vol_cols = [1, 2, 3, 4]
        spd_cols = [5, 6, 7, 8]
        tt_cols  = [9, 10, 11, 12]
        q_cols   = [13, 14, 15, 16]

        for i in vol_cols: self.validation_table.setColumnHidden(i, not show_vol)
        for i in spd_cols: self.validation_table.setColumnHidden(i, not show_spd)
        for i in tt_cols:  self.validation_table.setColumnHidden(i, not show_tt)
        for i in q_cols:   self.validation_table.setColumnHidden(i, not show_q)

        self._refit_table_columns(self.validation_table)

    def _collect_validation_data(self) -> List[Dict[str, Any]]:
        out: List[Dict[str, Any]] = []
        for r in range(self.validation_table.rowCount()):
            row_dict: Dict[str, Any] = {}
            for c, (_hdr, key) in enumerate(self.validation_columns):
                w = self.validation_table.cellWidget(r, c)
                val = w.as_number() if isinstance(w, NumericLineEdit) else None
                if val is not None and float(val).is_integer():
                    val = int(val)
                row_dict[key] = val
            out.append(row_dict)
        return out

    # -------------------------
    # D3) Error Penalty (forced 1.00) + hidden
    # -------------------------
    def _build_error_penalties(self):
        g = QGroupBox("Error Penalty")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        def penalty_spin(default=1.0):
            s = QDoubleSpinBox()
            s.setRange(1.0, 5.0)
            s.setDecimals(2)
            s.setSingleStep(0.25)
            s.setValue(default)
            s.setFixedWidth(120)
            return s

        grid.addWidget(_label_with_help("Penalty for Volume:", "Higher penalty => higher weight assigned to Volume."), 0, 0); self.pen_vol = penalty_spin(1.0); grid.addWidget(self.pen_vol, 0, 1)
        grid.addWidget(_label_with_help("Penalty for Speed:", "Higher penalty => higher weight assigned to Speed."), 0, 2); self.pen_spd = penalty_spin(1.0); grid.addWidget(self.pen_spd, 0, 3)
        grid.addWidget(_label_with_help("Penalty for Travel Time:", "Higher penalty => higher weight assigned to Travel Time."), 1, 0); self.pen_tt = penalty_spin(1.0); grid.addWidget(self.pen_tt, 1, 1)
        grid.addWidget(_label_with_help("Penalty for Queue:", "Higher penalty => higher weight assigned to Queue."), 1, 2); self.pen_q = penalty_spin(1.0); grid.addWidget(self.pen_q, 1, 3)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)
        g.setVisible(False)

    # -------------------------
    # E) Simulation Parameters
    # -------------------------
    def _build_simulation_parameters(self):
        g = QGroupBox("Simulation Parameters")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        grid = QGridLayout()
        grid.setHorizontalSpacing(10)
        grid.setVerticalSpacing(10)

        self.seed_period = QLineEdit("1800")
        self.eval_period = QLineEdit("3600")
        self.seed_period.setValidator(QIntValidator(0, 999999999, self))
        self.eval_period.setValidator(QIntValidator(0, 999999999, self))
        self.seed_period.setFixedWidth(140)
        self.eval_period.setFixedWidth(140)

        self.num_runs = QSpinBox(); self.num_runs.setRange(1, 9999); self.num_runs.setValue(2)
        self.rand_seed = QSpinBox(); self.rand_seed.setRange(1, 999999999); self.rand_seed.setValue(42)
        self.rand_inc = QSpinBox(); self.rand_inc.setRange(1, 999999999); self.rand_inc.setValue(10)

        grid.addWidget(_label_with_help("Seeding Period (sec):", "Seeding period in seconds."), 0, 0); grid.addWidget(self.seed_period, 0, 1)
        grid.addWidget(_label_with_help("Evaluation Period (sec):", "Evaluation period in seconds."), 0, 2); grid.addWidget(self.eval_period, 0, 3)

        grid.addWidget(_label_with_help("Number of Runs:", "Number of simulation runs."), 1, 0); grid.addWidget(self.num_runs, 1, 1)
        grid.addWidget(_label_with_help("Random Seed:", "Random seed (integer >= 1)."), 1, 2); grid.addWidget(self.rand_seed, 1, 3)
        grid.addWidget(_label_with_help("Random Seed Increment:", "Increment added to the seed across runs."), 1, 4); grid.addWidget(self.rand_inc, 1, 5)

        wrap = QHBoxLayout()
        wrap.addLayout(grid)
        wrap.addStretch()
        lay.addLayout(wrap)

        self.main_layout.addWidget(g)

    # -------------------------
    # F) Output
    # -------------------------
    def _build_output(self):
        g = QGroupBox("Output")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        lay = QVBoxLayout(g)

        row = QHBoxLayout()
        row.addWidget(_label_with_help("Results Directory:", "Select where to save calibrator outputs (results directory)."))
        self.result_dir = QLineEdit()
        self.result_dir.setFixedWidth(420)
        row.addWidget(self.result_dir)

        b = QPushButton("Browse…")
        b.setStyleSheet(SOFT_BUTTON_STYLE)
        b.clicked.connect(self._browse_result_dir)
        row.addWidget(b)

        row.addStretch()
        lay.addLayout(row)

        self.main_layout.addWidget(g)

    def _browse_result_dir(self):
        d = QFileDialog.getExistingDirectory(self, "Select Results Directory")
        if d:
            self.result_dir.setText(d)

    # -------------------------
    # G) Run + Log
    # -------------------------
    def _build_run(self):
        row = QHBoxLayout()

        self.run_btn = QPushButton("Run Vissim Calibrator")
        self.run_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        self.run_btn.clicked.connect(self._on_run)

        row.addStretch()
        row.addWidget(self.run_btn)
        row.addStretch()
        self.main_layout.addLayout(row)

        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.setStyleSheet(LOG_STYLE)
        self.log.setMinimumHeight(420)
        self.main_layout.addWidget(self.log)

    # -------------------------
    # Helpers for triplet tables (DSD/LCD)
    # -------------------------
    def _make_triplet_table(self, headers: List[str]) -> QTableWidget:
        t = QTableWidget(0, 4)
        t.setHorizontalHeaderLabels(headers)
        t.verticalHeader().setVisible(False)
        t.setMinimumHeight(180)
        t.setStyleSheet("""
            QTableWidget {
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 6px;
                font-size: 11pt;
            }
            QHeaderView::section {
                background: #f3f7ff;
                border: 1px solid #c9d4e3;
                padding: 6px;
                font-weight: 900;
                color: #143a5d;
            }
        """)
        return t

    def _make_table_buttons(self, table: QTableWidget) -> QHBoxLayout:
        btns = QHBoxLayout()
        add_btn = QPushButton("Add Row")
        rem_btn = QPushButton("Remove Selected Row(s)")
        add_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        rem_btn.setStyleSheet(SOFT_BUTTON_STYLE)
        add_btn.clicked.connect(lambda: self._table_add_row(table))
        rem_btn.clicked.connect(lambda: self._table_remove_selected_rows(table))
        btns.addWidget(add_btn)
        btns.addWidget(rem_btn)
        btns.addStretch()
        return btns

    def _table_add_row(self, table: QTableWidget):
        r = table.rowCount()
        table.insertRow(r)

        id_edit = NumericLineEdit("", allow_negative=False, allow_float=False)
        table.setCellWidget(r, 0, id_edit)

        for c in (1, 2, 3):
            e = NumericLineEdit("", allow_negative=True, allow_float=True)
            table.setCellWidget(r, c, e)

        self._refit_table_columns(table)

    def _table_remove_selected_rows(self, table: QTableWidget):
        rows = sorted({idx.row() for idx in table.selectedIndexes()}, reverse=True)
        for r in rows:
            table.removeRow(r)
        if table.rowCount() == 0:
            self._table_add_row(table)
        self._refit_table_columns(table)

    def _collect_triplet_table(self, table: QTableWidget) -> List[Tuple[int, Tuple[float, float, float]]]:
        out: List[Tuple[int, Tuple[float, float, float]]] = []
        for r in range(table.rowCount()):
            w0 = table.cellWidget(r, 0)
            w1 = table.cellWidget(r, 1)
            w2 = table.cellWidget(r, 2)
            w3 = table.cellWidget(r, 3)

            if not isinstance(w0, NumericLineEdit):
                continue
            idv = w0.as_number()
            if idv is None:
                continue
            id_int = int(idv)

            def n(w):
                return w.as_number() if isinstance(w, NumericLineEdit) else None

            lb = n(w1); sp = n(w2); ub = n(w3)
            if lb is None or sp is None or ub is None:
                continue

            out.append((id_int, (float(lb), float(sp), float(ub))))
        return out

    # -------------------------
    # RUN (with FIXED gates)
    # -------------------------
    def _on_run(self):
        self._license_refresh_if_possible()

        if not self._has_tier2_or_better():
            self._redirect_to_license_tab(
                "Tier 2 license is required to run the Vissim Calibrator.\n\n"
                "Please go to the License tab and activate Tier2 (or Tier3), then try again."
            )
            return

        self._enforce_multi_scenarios_license()

        Vissim_Version = int(self.vissim_version_combo.currentText() or "0")
        Path_ = self.project_path_edit.text().strip()
        Filename = self.filename_edit.text().strip()
        if Filename.lower().endswith(".inpx"):
            Filename = Filename[:-5]

        Multiple_Scenarios = (self.multiple_scenarios_combo.currentText().strip().lower() == "yes")

        project_file = self.project_file_edit.text().strip()
        if project_file and not (os.path.isfile(project_file) and project_file.lower().endswith(".inpx")):
            QMessageBox.warning(self, "Project File", "Project File must be an existing .inpx file.")
            return

        Result_Directory = self.result_dir.text().strip()
        if Result_Directory and not os.path.isdir(Result_Directory):
            QMessageBox.warning(self, "Results Directory", "Results Directory must be a directory.")
            return

        # Desired_DB build
        Desired_DB_list: List[Tuple[int, str, Tuple[Tuple[str, Tuple[Any, Any, Any]], ...]]] = []
        for b in self.db_blocks:
            db_num = int(b["db"].value())
            model = b["model"].currentText().strip()
            params_out: List[Tuple[str, Tuple[Any, Any, Any]]] = []

            for p in b["params"]:
                hp = p["hp"].currentText()
                if hp == "Other":
                    hp = p["other"].text().strip() or "Other"

                if p["sp_bool"].isVisible():
                    sp_bool_val = p["sp_bool"].currentData(Qt.ItemDataRole.UserRole)
                    if sp_bool_val is None:
                        sp_bool_val = (p["sp_bool"].currentText() in ("True", "Active"))
                    params_out.append((hp, (-1, bool(sp_bool_val), 1)))
                    continue

                lb = p["lb"].as_number()
                sp = p["sp"].as_number()
                ub = p["ub"].as_number()
                if lb is None or sp is None or ub is None:
                    continue
                params_out.append((hp, (float(lb), float(sp), float(ub))))

            if len(params_out) == 0:
                QMessageBox.warning(self, "Driving Behaviour", f"DB #{db_num}: Please enter at least one hyperparameter triplet (LB, Start, UB).")
                return

            Desired_DB_list.append((db_num, model, tuple(params_out)))

        Desired_DB = tuple(Desired_DB_list)

        # ✅ "Seeding" (capital S) ALWAYS False
        Calibrated_variables = {
            ("Speed Distribution", bool(self.chk_dsd.isChecked())),
            ("Lane Change Distance", bool(self.chk_lcd.isChecked())),
            ("Seeding", False),
        }

        DSD_List_Values = self._collect_triplet_table(self.dsd_table) if self.chk_dsd.isChecked() else []
        distance_List = self._collect_triplet_table(self.lcd_table) if self.chk_lcd.isChecked() else []

        # ✅ "seeding" (small s) ALWAYS tuple and hidden
        seeding = (0, 0.35, 0.99)
        Num_Seeding_Intervals = 4

        Trials = (bool(self.stop_after_trials.isChecked()), int(self.trials_spin.value()))
        SubTrial = int(self.subtrials_spin.value())

        Random_Slection = int(self.random_selection_spin.value())
        if Random_Slection >= SubTrial:
            QMessageBox.warning(self, "Tuning Parameters", "Number of Randomly Selected Sub-Trials must be lower than Number of Sub-Trials.")
            return

        Variant = float(self.variant_spin.value())
        Variant_update = (self.variant_update_combo.currentText().strip().lower() == "yes")
        Variant_List = self._collect_variant_list() if Variant_update else [(0, Variant)]

        step_active = bool(self.stepwise_active_combo.currentData())
        accept_pct = float(self.stepwise_accept_pct.value()) / 100.0
        step_pct   = float(self.stepwise_value_pct.value()) / 100.0
        Step_Wise_Variant = (bool(step_active), (accept_pct, step_pct))

        Validation_List = {
            ("Volume", bool(self.val_volume.isChecked())),
            ("Speed", bool(self.val_speed.isChecked())),
            ("Travel Time", bool(self.val_tt.isChecked())),
            ("Queue", bool(self.val_queue.isChecked())),
        }

        validation_data = self._collect_validation_data()

        if not Multiple_Scenarios:
            scen_vals = []
            for row in validation_data:
                sv = row.get("Scenario", None)
                if sv is None:
                    continue
                try:
                    scen_vals.append(int(sv))
                except Exception:
                    continue
            uniq = sorted(set(scen_vals))
            if len(uniq) > 1:
                QMessageBox.warning(
                    self,
                    "Multiple Scenarios",
                    "Multiple Scenarios is set to NO.\n"
                    "Please use only ONE Scenario number in the Validation Data table."
                )
                return

        Penalty_Volume = 1.0
        Penalty_Speed = 1.0
        Penalty_TravelTime = 1.0
        Penalty_Queue = 1.0

        Seeding_Period = int(self.seed_period.text().strip() or "0")
        eval_period = int(self.eval_period.text().strip() or "0")
        if eval_period <= Seeding_Period:
            QMessageBox.warning(self, "Simulation Parameters", "Evaluation Period must be greater than Seeding Period.")
            return

        NumRuns = int(self.num_runs.value())
        Random_Seed = int(self.rand_seed.value())
        Random_SeedInc = int(self.rand_inc.value())

        self.log.clear()

        def log_line(s: str):
            self.log.append(s)
            print(s)

        log_line("VissiCaRe Calibrator inputs:\n")
        for k, v in [
            ("Vissim_Version", Vissim_Version),
            ("Path", Path_),
            ("Filename", Filename),
            ("Result_Directory", Result_Directory),
            ("Multiple_Scenarios", Multiple_Scenarios),
            ("Calibrated_variables", Calibrated_variables),
            ("Desired_DB", Desired_DB),
            ("DSD_List_Values", DSD_List_Values),
            ("distance_List", distance_List),
            ("seeding", seeding),
            ("Num_Seeding_Intervals", Num_Seeding_Intervals),
            ("Trials", Trials),
            ("SubTrial", SubTrial),
            ("Random_Slection", Random_Slection),
            ("Variant", Variant),
            ("Variant_update", Variant_update),
            ("Variant_List", Variant_List),
            ("Step_Wise_Variant", Step_Wise_Variant),
            ("Validation_List", Validation_List),
            ("validation_data", validation_data),
            ("Penalty_Volume", Penalty_Volume),
            ("Penalty_Speed", Penalty_Speed),
            ("Penalty_TravelTime", Penalty_TravelTime),
            ("Penalty_Queue", Penalty_Queue),
            ("Seeding_Period", Seeding_Period),
            ("eval_period", eval_period),
            ("NumRuns", NumRuns),
            ("Random_Seed", Random_Seed),
            ("Random_SeedInc", Random_SeedInc),
        ]:
            log_line(f"{k}: {v}   (type={type(v).__name__})")

        log_line("\nRunning Calibration_Function (your core)...\n")

        try:
            res = Calibration_Function(
                Vissim_Version, Path_, Filename, Result_Directory, Calibrated_variables, Desired_DB, DSD_List_Values,
                distance_List, seeding, Num_Seeding_Intervals, Trials, SubTrial, Random_Slection, Variant, Variant_update,
                Variant_List, Step_Wise_Variant, Validation_List, validation_data, Penalty_Volume, Penalty_Speed,
                Penalty_TravelTime, Penalty_Queue, Seeding_Period, eval_period, NumRuns, Random_Seed, Random_SeedInc
            )
        except Exception as e:
            log_line(f"ERROR: {e}")
            QMessageBox.critical(self, "Calibration Error", f"{e}")
            return
        else:
            log_line(f"\nCalibration_Function finished. Return: {res}")





























































# ============================================================
# ✅ Global modern scrollbar style (Vertical + Horizontal)
# Apply once at app startup via apply_global_scrollbar_style(app)
# ============================================================

GLOBAL_SCROLLBAR_STYLE = """
/* ScrollArea base: keep clean */
QScrollArea { border: none; background: transparent; }
QScrollArea > QWidget > QWidget { background: transparent; }

/* ===== Vertical ScrollBar ===== */
QScrollBar:vertical {
    background: transparent;
    width: 12px;
    margin: 2px 2px 2px 2px;
}
QScrollBar::handle:vertical {
    background: #b9c7dd;
    min-height: 30px;
    border-radius: 6px;
}
QScrollBar::handle:vertical:hover { background: #9fb3cf; }
QScrollBar::handle:vertical:pressed { background: #86a4ca; }
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
    background: transparent;
}
QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical { background: transparent; }

/* ===== Horizontal ScrollBar ===== */
QScrollBar:horizontal {
    background: transparent;
    height: 12px;
    margin: 2px 2px 2px 2px;
}
QScrollBar::handle:horizontal {
    background: #b9c7dd;
    min-width: 30px;
    border-radius: 6px;
}
QScrollBar::handle:horizontal:hover { background: #9fb3cf; }
QScrollBar::handle:horizontal:pressed { background: #86a4ca; }
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
    background: transparent;
}
QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal { background: transparent; }

/* Corner between horizontal+vertical */
QAbstractScrollArea::corner { background: transparent; }
"""

GLOBAL_LABEL_STYLE = """
QLabel { background: transparent; }
"""

GLOBAL_INPUT_STYLE = """
QLineEdit {
    background-color: #ffffff;
}
QComboBox {
    background-color: #ffffff;
}
QComboBox:editable {
    background-color: #ffffff;
}
QComboBox QAbstractItemView {
    background-color: #ffffff;
}
QAbstractSpinBox {
    background-color: #ffffff;
}
QAbstractSpinBox::lineEdit {
    background-color: #ffffff;
}
QSpinBox, QDoubleSpinBox {
    background-color: #ffffff;
}
"""

def apply_global_scrollbar_style(app: "QApplication"):
    """
    Call this ONCE right after you create QApplication.
    It appends the global scrollbar QSS without removing your other styles.
    """
    existing = app.styleSheet() or ""
    if GLOBAL_SCROLLBAR_STYLE.strip() in existing:
        return
    app.setStyleSheet(existing + "\n\n" + GLOBAL_SCROLLBAR_STYLE)


def apply_global_label_style(app: "QApplication"):
    """
    Call this ONCE right after you create QApplication.
    It appends the global label QSS without removing your other styles.
    """
    existing = app.styleSheet() or ""
    if GLOBAL_LABEL_STYLE.strip() in existing:
        return
    app.setStyleSheet(existing + "\n\n" + GLOBAL_LABEL_STYLE)


def apply_global_input_style(app: "QApplication"):
    """
    Call this ONCE right after you create QApplication.
    It appends the global input QSS without removing your other styles.
    """
    existing = app.styleSheet() or ""
    if GLOBAL_INPUT_STYLE.strip() in existing:
        return
    app.setStyleSheet(existing + "\n\n" + GLOBAL_INPUT_STYLE)
    pal = app.palette()
    pal.setColor(QPalette.ColorRole.Base, QColor("#ffffff"))
    pal.setColor(QPalette.ColorRole.Button, QColor("#ffffff"))
    app.setPalette(pal)


# ============================================================
# Help Tab (this to help open the sections completely)
# ============================================================

class AutoHeightTextBrowser(QTextBrowser):
    """
    QTextBrowser that expands vertically to fit its content,
    so the outer QScrollArea handles scrolling (no inner scrollbars).
    """
    def __init__(self, min_h: int = 160, parent=None):
        super().__init__(parent)
        self._min_h = min_h

        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        self.document().contentsChanged.connect(self._update_height)

    def resizeEvent(self, e):
        super().resizeEvent(e)
        # Ensure the document reflows to the available width
        self.document().setTextWidth(self.viewport().width())
        self._update_height()

    def _update_height(self):
        # QSizeF -> height in pixels (float)
        doc_h = self.document().size().height()
        frame = self.frameWidth() * 2
        # small safety padding so last line isn't clipped
        new_h = int(doc_h + frame + 18)
        new_h = max(self._min_h, new_h)

        # Fix the height so it doesn't create its own scrollbar
        self.setMinimumHeight(new_h)
        self.setMaximumHeight(new_h)


# ============================================================
# Help Tab (placeholder)
# ============================================================

# Reuse your existing styles:
# - SECTION_GROUPBOX_STYLE


class CollapsibleSection(QWidget):
    """
    Full-width collapsible section (MAIN only).
    IMPORTANT: does NOT call adjustSize() on parents (prevents window shrinking).
    """
    def __init__(self, title: str, parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        self.header_btn = QToolButton()
        self.header_btn.setText(title)
        self.header_btn.setCheckable(True)
        self.header_btn.setChecked(False)
        self.header_btn.setArrowType(Qt.ArrowType.RightArrow)
        self.header_btn.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        self.header_btn.setCursor(Qt.CursorShape.PointingHandCursor)
        self.header_btn.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.header_btn.setMinimumHeight(42)

        self.header_btn.setStyleSheet("""
            QToolButton {
                text-align: left;
                padding: 10px 12px;
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                font-weight: 900;
                color: #143a5d;
                background: qlineargradient(
                    x1: 0, y1: 0,
                    x2: 1, y2: 0,
                    stop: 0  #fff7d6,
                    stop: 0.5 #ffffff,
                    stop: 1  #e4f0ff
                );
            }
            QToolButton:hover {
                background: qlineargradient(
                    x1: 0, y1: 0,
                    x2: 1, y2: 0,
                    stop: 0  #ffe066,
                    stop: 0.5 #fffef5,
                    stop: 1  #cfe4ff
                );
            }
        """)

        self.content = QWidget()
        self.content.setVisible(False)
        self.content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        self.content_layout = QVBoxLayout(self.content)
        self.content_layout.setContentsMargins(0, 0, 0, 0)
        self.content_layout.setSpacing(12)

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)
        root.addWidget(self.header_btn)
        root.addWidget(self.content)

        self.header_btn.toggled.connect(self._on_toggled)

    def _on_toggled(self, checked: bool):
        self.header_btn.setArrowType(Qt.ArrowType.DownArrow if checked else Qt.ArrowType.RightArrow)
        self.content.setVisible(checked)

        # Do NOT call adjustSize() on any parent widgets (prevents window shrinking).
        self.content.updateGeometry()
        self.updateGeometry()

    def setContentWidget(self, w: QWidget):
        while self.content_layout.count():
            it = self.content_layout.takeAt(0)
            if it.widget():
                it.widget().setParent(None)
        self.content_layout.addWidget(w)


# ============================================================
# NEW: Hover-tabs + click-next image deck (node_1.jpg ... node_9.jpg)
# ============================================================

class HoverTabsImageDeck(QWidget):
    """
    - Shows an image
    - On hover: shows numbered "tabs" (buttons)
    - Left click on the image: goes to next image
    """
    def __init__(self, image_files: list[str], parent=None):
        super().__init__(parent)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        self.setAttribute(Qt.WidgetAttribute.WA_Hover, True)

        self._files = image_files[:]
        self._pix = [QPixmap(f) for f in self._files]
        self._i = 0

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        self.img_label = QLabel()
        self.img_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.img_label.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        self.img_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.img_label.setStyleSheet("""
            QLabel {
                background: #ffffff;
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 6px;
            }
        """)

        self.tabs = QWidget()
        self.tabs.setVisible(False)
        tabs_lay = QHBoxLayout(self.tabs)
        tabs_lay.setContentsMargins(0, 0, 0, 0)
        tabs_lay.setSpacing(6)
        tabs_lay.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self._btns: list[QPushButton] = []
        for k in range(len(self._files)):
            b = QPushButton(str(k + 1))
            b.setCheckable(True)
            b.setFixedSize(34, 26)
            b.setCursor(Qt.CursorShape.PointingHandCursor)
            b.setStyleSheet("""
                QPushButton {
                    border: 1px solid #c9d4e3;
                    border-radius: 12px;
                    background: #ffffff;
                    color: #143a5d;
                    font-weight: 900;
                }
                QPushButton:hover { background: #fff7d6; }
                QPushButton:checked {
                    background: #e4f0ff;
                    border: 1px solid #7aa7d9;
                }
            """)
            b.clicked.connect(lambda checked=False, idx=k: self.set_index(idx))
            self._btns.append(b)
            tabs_lay.addWidget(b)

        hint = QLabel("Hover: tabs • Left click: next")
        hint.setAlignment(Qt.AlignmentFlag.AlignCenter)
        hint.setStyleSheet("QLabel { color:#143a5d; font-weight:900; }")

        root.addWidget(self.img_label)
        root.addWidget(self.tabs)
        root.addWidget(hint)

        self._apply()

    def enterEvent(self, e):
        super().enterEvent(e)
        self.tabs.setVisible(True)

    def leaveEvent(self, e):
        super().leaveEvent(e)
        self.tabs.setVisible(False)

    def mousePressEvent(self, e):
        super().mousePressEvent(e)
        if e.button() == Qt.MouseButton.LeftButton:
            self.next()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._apply()

    def set_index(self, idx: int):
        if not self._pix:
            return
        self._i = max(0, min(idx, len(self._pix) - 1))
        self._apply()

    def next(self):
        if not self._pix:
            return
        self._i = (self._i + 1) % len(self._pix)
        self._apply()

    def _apply(self):
        for k, b in enumerate(self._btns):
            b.setChecked(k == self._i)

        if not self._pix:
            self.img_label.setText("(No slides)")
            self.img_label.setMinimumHeight(180)
            self.img_label.setMaximumHeight(180)
            return

        pm = self._pix[self._i]
        if pm.isNull():
            self.img_label.setText(f"(Missing image: {Path(self._files[self._i]).name})")
            self.img_label.setStyleSheet("QLabel { color:#b00020; font-weight:900; }")
            self.img_label.setMinimumHeight(180)
            self.img_label.setMaximumHeight(180)
            return

        # scale to a reasonable width (fits your help page)
        target_w = min(980, max(420, self.width() - 24))
        pm2 = pm.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
        self.img_label.setPixmap(pm2)
        self.img_label.setMinimumHeight(pm2.height() + 12)
        self.img_label.setMaximumHeight(pm2.height() + 12)


class HelpTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet(TAB_BACKGROUND_STYLE)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)

        # ✅ ADDED: Apply the new modern scrollbars to this HelpTab scroll area too
        # (You should still call apply_global_scrollbar_style(app) once at startup for the whole app.)
        scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)

        container = QWidget()
        container.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        lay = QVBoxLayout(container)
        lay.setAlignment(Qt.AlignmentFlag.AlignTop)
        lay.setSpacing(14)
        lay.setContentsMargins(12, 12, 12, 24)

        # Title
        title = QLabel("Help")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        title.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        title.setStyleSheet("""
            QLabel {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                    stop:0 #fdf8e3, stop:0.5 #ffffff, stop:1 #e8f4fd);
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                padding: 10px;
                font-size: 18px;
                font-weight: 900;
                color: #143a5d;
            }
        """)
        lay.addWidget(title)

        # Overview (always visible)
        overview = QGroupBox("Overview")
        overview.setStyleSheet(SECTION_GROUPBOX_STYLE)
        ov_lay = QVBoxLayout(overview)
        ov_lay.addWidget(self._make_browser("""
            <h2>VissiCaRe</h2>
            <p>
              VissiCaRe is a software devoloped by the collaboration of Traffic and Machine learning Engineers.
              Its goal is to automate the process of microsimulation as fast and efficient as possible.
              The methods used in each process is tested and validated.
              For any help needed, contact the help desk at
              <a href="mailto:hekpdesk@intree.com">hekpdesk@intree.com</a>.
            </p>
            <p>
              Below are the engines of the software. Each engine serves a purpose and aims to solve a problem by reducing time
              of calibrating the VISSIM mode or collecting results while maintaining full accuracy.
            </p>
            <p>
              For more information, visit our website at <a href="XXXXXXXX">XXXXXXXX</a> (Example website).
            </p>
        """, min_h=150))
        lay.addWidget(overview)

        # MAIN collapsibles only
        lay.addWidget(self._build_nodes_help())
        lay.addWidget(self._build_links_help())
        lay.addWidget(self._build_tt_help())
        lay.addWidget(self._build_throughput_help())
        lay.addWidget(self._build_network_help())
        lay.addWidget(self._build_calibrator_help())

        lay.addStretch(1)

        scroll.setWidget(container)
        outer.addWidget(scroll)

    # ---------- UI helpers ----------
    def _make_browser(self, html: str, min_h: int = 160) -> QTextBrowser:
        b = AutoHeightTextBrowser(min_h=min_h)
        b.setOpenExternalLinks(True)
        b.setHtml(html)
        b.setStyleSheet("""
            QTextBrowser {
                background-color: #ffffff;
                border-radius: 10px;
                border: 1px solid #c9d4e3;
                padding: 10px;
                font-family: Segoe UI, Arial;
                font-size: 11pt;
            }
        """)
        b.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
        b.setLineWrapMode(QTextBrowser.LineWrapMode.WidgetWidth)
        return b

    def _section_box(self, title: str, html: str) -> QGroupBox:
        g = QGroupBox(title)
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        g.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        v = QVBoxLayout(g)
        v.addWidget(self._make_browser(html, min_h=140))
        return g

    def _res_path(self, filename: str) -> str:
        # Works with your existing app_dir() helper if present; falls back to script folder.
        try:
            base = app_dir()
        except Exception:
            base = Path(__file__).resolve().parent
        return str((base / filename).resolve())

    def _figures_side_by_side(self) -> QGroupBox:
        g = QGroupBox("")
        g.setStyleSheet(SECTION_GROUPBOX_STYLE)
        g.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)

        root = QVBoxLayout(g)
        root.setSpacing(10)

        # Images row
        row = QHBoxLayout()
        row.setSpacing(16)

        def make_one(img_name: str, caption: str) -> QWidget:
            w = QWidget()
            v = QVBoxLayout(w)
            v.setContentsMargins(0, 0, 0, 0)
            v.setSpacing(8)

            lbl_img = QLabel()
            lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_img.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

            pm = QPixmap(self._res_path(img_name))
            if not pm.isNull():
                # Smaller + side-by-side: keep a reasonable max width
                target_w = 520  # adjust if you want even smaller (e.g., 420)
                pm2 = pm.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
                lbl_img.setPixmap(pm2)
                lbl_img.setMinimumHeight(pm2.height())
                lbl_img.setMaximumHeight(pm2.height())
            else:
                lbl_img.setText(f"(Missing image: {img_name})")
                lbl_img.setStyleSheet("color:#b00020; font-weight:900;")

            lbl_cap = QLabel(caption)
            lbl_cap.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_cap.setStyleSheet("""
                QLabel {
                    color:#143a5d;
                    font-weight: 900;
                }
            """)

            v.addWidget(lbl_img)
            v.addWidget(lbl_cap)
            return w

        row.addWidget(make_one("Figure 1.jpg", "Figure 1: Unpreferred Practice"))
        row.addWidget(make_one("Figure 2.jpg", "Figure 2: Preferred Practice"))
        root.addLayout(row)

        return g

    # --- NEW: embed image inside a section box (no extra groupbox) ---
    def _figure_widget(self, img_base_or_name: str, caption: str = "") -> QWidget:
        """
        Image + optional caption as a plain QWidget (no groupbox),
        so you can embed it inside an existing section box (like C).
        Accepts full filename OR base name without extension.
        """
        w = QWidget()
        w.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        root = QVBoxLayout(w)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(10)

        lbl_img = QLabel()
        lbl_img.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_img.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)

        # Try load with/without extension (jpeg/jpg/png)
        candidates = []
        p = Path(img_base_or_name)
        if p.suffix:
            candidates = [img_base_or_name]
        else:
            candidates = [f"{img_base_or_name}.jpg", f"{img_base_or_name}.jpeg", f"{img_base_or_name}.png"]

        pm = QPixmap()
        for name in candidates:
            pm_try = QPixmap(self._res_path(name))
            if not pm_try.isNull():
                pm = pm_try
                break

        if not pm.isNull():
            target_w = 980  # same scale you used before
            pm2 = pm.scaledToWidth(target_w, Qt.TransformationMode.SmoothTransformation)
            lbl_img.setPixmap(pm2)
            lbl_img.setMinimumHeight(pm2.height())
            lbl_img.setMaximumHeight(pm2.height())
        else:
            lbl_img.setText(f"(Missing image: {candidates[0] if candidates else img_base_or_name})")
            lbl_img.setStyleSheet("color:#b00020; font-weight:900;")

        root.addWidget(lbl_img)

        if caption:
            lbl_cap = QLabel(caption)
            lbl_cap.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl_cap.setStyleSheet("""
                QLabel {
                    color:#143a5d;
                    font-weight: 900;
                }
            """)
            root.addWidget(lbl_cap)

        return w

    def _main_section(self, title: str, intro_html: str, link: str, boxes: list[QWidget]) -> CollapsibleSection:
        sec = CollapsibleSection(title)

        content = QWidget()
        content.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        v = QVBoxLayout(content)
        v.setContentsMargins(0, 0, 0, 0)
        v.setSpacing(12)

        v.addWidget(self._make_browser(f"""
            <h2>{title}</h2>
            <p>{intro_html}</p>
            <p><b>Docs:</b> <a href="{link}">{link}</a></p>
        """, min_h=140))

        for b in boxes:
            v.addWidget(b)

        sec.setContentWidget(content)
        return sec

    # ---------- MAIN sections (no project data, no real lists/values) ----------
    def _build_nodes_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/nodes-results"
        intro = (
            "VissiCaRe helps you automate the collection of the nodes results. It organizes the results in a professional format "
            "and saves the output file in the directory you specify.<br><br>"
            "You can collect the results of multiple scenarios at the same time and have them in the same output file if you are using "
            "scenario manager in your VISSIM project. If you are not using scenario manager, you still can get the nodes results by setting "
            "the scenario umber to 0 in the scenario list.<br>"
            "if you are not using scenario manager and have more than one project separately, you should collect their results separately if you wish to. "
            "Make sure to save the files with different names so they do not overwrite each other.<br><br>"
            "<b>Good practice to draw the nodes for intersection results:</b><br><br>"
            "When coding the nodes in your project, avoid drawing a node on connectors or minor links. Try to draw on the major links that can direct MISTAR solution "
            "(If used to catch the cardinal name). Even if MISTAR solution is not active, VissiCaRe can learn better if you follow this practice. "
            "See below Figure 1 and Figure 2 for good practice. Note how the node in Figure 2 is passing over the representing links of each approach. "
            "<br><br>"
            "<b>Important:</b> When two links have the same approach name (For example two approaches are SW), the software will catch this "
            "and assign different names to them.<br>"
            "The software will add <b>(First)</b> before the first SW link, and <b>(Second)</b> before the second SW link. The numbers are assigned counter clockwise.<br>"
            "The software does the same with the movements. If, from the same approach, two movements are Left Turn, the software will assign (First) and (Second) to the movements, also in a counter clockwise manner.<br>"
            "<b>Note:</b> If an approach gets two <b>U-Turn</b> movements, they follow the counter clockwise naming as well. Therefore, your first <b>U-Turn will</b> likely be the "
            "<b>Hard Right Turn</b> and your last U-Turn is the actual <b>U-Turn</b>, unless your project is left-hand traffic (LHT), then your <b>U-Turn</b> is the <b>(First)<b>. "
        )

        # ✅ UPDATED: build C — Intersections as a real widget so we can insert the 9-slide deck inside it
        c_box = QGroupBox("C — Intersections")
        c_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        c_box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        c_layout = QVBoxLayout(c_box)
        c_layout.setSpacing(12)

        c_layout.addWidget(self._make_browser("""
                <h4>Nodes List</h4>
                <ul>
                  <li>Defines which nodes/intersections to include and how they should be treated in reporting.</li>
                  <li>Each row typically includes: <b>Node Number</b>, <b>Signalized/Unsignalized</b>, <b>Intersection Type</b>.</li>
                  <li><b>Overpass/Underpass</b> option: lets you exclude specified link pairs that are on a different grade (prevents mixing unrelated movements). See Example above, An interstate passes over a Single Point Urban Interchange (SPUI). The user may exclude the interstate from the results by checking the <b>Overpass</b> and defining the <b>From</b> and <b>To</b> as shown in the example. In this Case, The Intersection will be reported without the Interstate links, which are free movements and are not involved int he analysis in this case.</li>
                  <li>If your node is named, the software will extract the name and use it in the report, otherwise, it will use the node number.</li>
                  <li>you can add as many nodes as your project has and you can exclude any links that are not important.</li>
                </ul>

                <h4>Level of Service</h4>
                <ul>
                  <li>Controls LOS letter thresholds for <b>signalized</b> and <b>unsignalized</b> intersections.</li>
                  <li>Thresholds should be continuous and non-overlapping.</li>
                </ul>

                <h4>LOS Colors</h4>
                <ul>
                  <li>Maps LOS letters (A–F) to display colors in the output.</li>
                  <li>“No Color” keeps the cell unfilled.</li>
                </ul>

                <h4>Results Formatting</h4>
                <ul>
                  <li>Three scopes: <b>Overall Intersection</b>, <b>Approach</b>, <b>Movement</b>.</li>
                  <li>Each scope always includes <b>LOS</b>, plus up to 4 additional metrics.</li>
                  <li>Wrappers (brackets/parentheses) control how values are grouped/displayed; duplicates in the same scope are not allowed.</li>
                </ul>


                <h4>Units</h4>
                <ul>
                  <li>Choose whether distance-based outputs are interpreted/displayed in <b>meters</b> or <b>feet</b>.</li>
                </ul>

                <h4>Mistar Solution:</h4>
                <ul>
                  <li>Mistar Solution is a mathematical agent that uses analytical geometry and Generative AI to identify the cardinal name of the approach and its movements.</li>
                  <li>If Mistar Solution is not enabled, the approach name and the movements are 100 % identical to what VISSIM has in the nodes results tab. If the movement is SW-W, the software will report Northeastbound Sharp Left Turn.</li>
                  <li>On the other hand, if Mistar Solution is activated, it will check the entire intersection, the geometry of the approaches and then asign a cardinal name to each approach.</li>
                </ul>

                <p><b>Below is how the Cardinal names are assigned to each approach and movement.</b></p>
        """, min_h=520))

        # Add the 9-slide "tab deck" INSIDE section C
        node_files = [self._res_path(f"node_{i}.jpg") for i in range(1, 10)]
        c_layout.addWidget(HoverTabsImageDeck(node_files))

        boxes = [
            self._figures_side_by_side(),

            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>VISSIM Version</b>: Which VISSIM major version the tool should connect to (affects COM/API compatibility).</li>
                  <li><b>Project Path</b>: Folder containing the model files; must be an existing directory where your VISSIM project is.</li>
                  <li><b>File Name</b>: Base name of the VISSIM model file (without extension).</li>
                </ul>
                <p><i>Tip:</i> If the project cannot load, verify the version selection and that the path points to the correct folder.</p>
            """),
            self._section_box("B — Scenario Setup", """
                <ul>
                  <li><b>Scenario List</b>: Which scenarios to report results for.</li>
                  <li><b>Scenario Number</b>: ID used by your scenario manager. Select <b>0</b> if you are not using scenario manager (Note: If you select 0, you will need to do the process again for any different project you want to collect the results for)</li>
                  <li><b>Run Model</b>: If enabled, the tool will run simulation before extracting results; if disabled, it only reads existing results.</li>
                </ul>
            """),

            # ✅ ADDED (ONLY): Overpass example image between B and C
            self._figure_widget("overpass.jpg", "Overpass Example"),

            c_box,
            self._section_box("D — Simulation Parameters", """
                <ul>
                  <li><b>Seeding Period</b>: warm-up time before evaluation starts.</li>
                  <li><b>Evaluation Period</b>: time window used to compute results (must be greater than seeding).</li>
                  <li><b>Number of Runs</b>, <b>Random Seed</b>, <b>Seed Increment</b>: control stochastic runs and reproducibility.</li>
                </ul>
            """),
            self._section_box("E — Results Specifics", """
                <ul>
                  <li><b>Run</b>: which run statistic to output (average, min, max, etc.).</li>
                  <li><b>Vehicle Type</b>: filter results by class (or include all).</li>
                  <li><b>Time Interval</b>: which interval(s) to use (average across intervals or a specific interval index).</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li><b>Result Directory</b>: where to save the exported report.</li>
                  <li><b>Output File</b>: report file name.</li>
                  <li><b>Reporter / Date / Company / Project</b>: header metadata written into the report.</li>
                </ul>
            """),
        ]
        return self._main_section("Node Results", intro, link, boxes)

    def _build_links_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/link-results"
        intro = (
            "Links Results can be collected for different scenarios and reported in one output file. "
            "If the user wishes to report multiple results at the same time, they can do so if they are using scenario manger, "
            "otherwise, the number of the scenario should be set to 0 and only one scenario results will be reported. "
            "The software can process results for both <b>per link</b> and <b>per lane</b> as explained below."
        )

        # --- C box built as a real widget so the image is INSIDE the same box ---
        c_box = QGroupBox("C — Links / Segments")
        c_box.setStyleSheet(SECTION_GROUPBOX_STYLE)
        c_box.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred)
        c_layout = QVBoxLayout(c_box)
        c_layout.setSpacing(12)

        c_layout.addWidget(self._make_browser("""
            <ul>
              <li>Defines which links (or grouped segments) the report should include.</li>
              <li>Each segment typically has:
                <ul>
                  <li><b>Display Name</b> (for the report)</li>
                  <li><b>Segment Type</b> (e.g., auxiliary/weaving/merging)</li>
                  <li><b>One or more Link IDs</b> included in that segment</li>
                </ul>
              </li>
              <li>The table can contain as many segments as needed.</li>
              <li><b>The figure below shows the good practice of defining the links inside a segment.</b> The number of the lanes across all the links inside a segment should be the same.
                  If a link has more lanes than the others, the software aggregates the results on the similar lanes, and will have the fourth stand alone.</li>
            </ul>
        """, min_h=200))

        # IMPORTANT: your real file is Links_help.jpg (with the "s" in Links)
        c_layout.addWidget(self._figure_widget("Links_help", "Links_help: Good practice of defining links inside a segment"))

        boxes = [
            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>VISSIM Version</b>: Which VISSIM major version the tool should connect to (affects COM/API compatibility).</li>
                  <li><b>Project Path</b>: Folder containing the model files; must be an existing directory where your VISSIM project is.</li>
                  <li><b>File Name</b>: Base name of the VISSIM model file (without extension).</li>
                </ul>
                <p><i>Tip:</i> If the project cannot load, verify the version selection and that the path points to the correct folder.</p>
            """),
            self._section_box("B — Scenario Setup", """
                <ul>
                  <li><b>Scenario List</b>: Which scenarios to report results for.</li>
                  <li><b>Scenario Number</b>: ID used by your scenario manager. Select <b>0</b> if you are not using scenario manager
                      (Note: If you select 0, you will need to do the process again for any different project you want to collect the results for)</li>
                  <li><b>Run Model</b>: If enabled, the tool will run simulation before extracting results; if disabled, it only reads existing results.</li>
                </ul>
            """),
            c_box,
            self._section_box("D — Simulation Parameters", """
                <ul>
                  <li><b>Seeding Period</b>: warm-up time before evaluation starts.</li>
                  <li><b>Evaluation Period</b>: time window used to compute results (must be greater than seeding).</li>
                  <li><b>Number of Runs</b>, <b>Random Seed</b>, <b>Seed Increment</b>: control stochastic runs and reproducibility.</li>
                </ul>
            """),
            self._section_box("E — Results Specifics", """
                <ul>
                  <li><b>Run</b>: which run statistic to output (average, min, max, etc.).</li>
                  <li><b>Vehicle Type</b>: filter results by class (or include all).</li>
                  <li><b>Time Interval</b>: which interval(s) to use (average across intervals or a specific interval index).</li>
                  <li><b>Per Link</b>: collects the results per link only.</li>
                  <li><b>Per Lane</b>: collects the results per lane and per link.</li>
                  <li><b>Units</b>: meters or feet for distance-based values.</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li><b>Result Directory</b>: where to save the exported report.</li>
                  <li><b>Output File Name</b>: report file name.</li>
                </ul>
            """),
        ]
        return self._main_section("Link Results", intro, link, boxes)

    def _build_tt_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/travel-time-results"
        intro = (
            "Travel Time Results can be collected for different scenarios and reported in one output file. "
            "If the user wishes to report multiple results at the same time, they can do so if they are using scenario manager, "
            "otherwise, the number of the scenario should be set to 0 and only one scenario results will be reported.<br><br>"
            "Travel time is collected using Travel Time Measurement objects (portals) in VISSIM. "
            "Make sure your portal IDs are correct and match the measurement objects in your network."
        )

        boxes = [
            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>VISSIM Version</b>: Which VISSIM major version the tool should connect to (affects COM/API compatibility).</li>
                  <li><b>Project Path</b>: Folder containing the model files; must be an existing directory where your VISSIM project is.</li>
                  <li><b>File Name</b>: Base name of the VISSIM model file (without extension).</li>
                </ul>
                <p><i>Tip:</i> If the project cannot load, verify the version selection and that the path points to the correct folder.</p>
            """),
            self._section_box("B — Scenario Setup", """
                <ul>
                  <li><b>Scenario List</b>: Which scenarios to report results for.</li>
                  <li><b>Scenario Number</b>: ID used by your scenario manager. Select <b>0</b> if you are not using scenario manager
                      (Note: If you select 0, you will need to do the process again for any different project you want to collect the results for)</li>
                  <li><b>Run Model</b>: If enabled, the tool will run simulation before extracting results; if disabled, it only reads existing results.</li>
                </ul>
            """),
            self._section_box("C — Travel Time Locations", """
                <ul>
                  <li>Defines the travel-time corridors/segments you want to report.</li>
                  <li>Each corridor typically includes:
                    <ul>
                      <li><b>Mainline Name</b> (for the report)</li>
                      <li><b>Direction</b> (for the report)</li>
                      <li><b>Portal IDs</b> (the Travel Time Measurement objects used to compute travel time)</li>
                    </ul>
                  </li>
                  <li>Portal IDs must match the IDs in VISSIM exactly.</li>
                  <li>If a portal ID is missing or incorrect, the corresponding corridor may export blank results.</li>
                </ul>
            """),
            self._section_box("D — Simulation Parameters", """
                <ul>
                  <li><b>Seeding Period</b>: warm-up time before evaluation starts.</li>
                  <li><b>Evaluation Period</b>: time window used to compute results (must be greater than seeding).</li>
                  <li><b>Number of Runs</b>, <b>Random Seed</b>, <b>Seed Increment</b>: control stochastic runs and reproducibility.</li>
                </ul>
            """),
            self._section_box("E — Results Specifics", """
                <ul>
                  <li><b>Run</b>: which run statistic to output (average, min, max, etc.).</li>
                  <li><b>Time Interval</b>: which interval(s) to use (average across intervals or a specific interval index).</li>
                  <li><b>Units</b>: output units for speed/distance/time where applicable.</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li><b>Result Directory</b>: where to save the exported report.</li>
                  <li><b>Output File Name</b>: report file name.</li>
                </ul>
            """),
        ]
        return self._main_section("Travel Time Results", intro, link, boxes)

    def _build_throughput_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/throughput-results"
        intro = (
            "Throughput Results can be collected for different scenarios and reported in one output file. "
            "If the user wishes to report multiple results at the same time, they can do so if they are using scenario manager, "
            "otherwise, the number of the scenario should be set to 0 and only one scenario results will be reported.<br><br>"
            "Throughput is collected by counting vehicles passing through the nodes you define. "
            "Make sure the node IDs you enter match your VISSIM network."
        )

        boxes = [
            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>VISSIM Version</b>: Which VISSIM major version the tool should connect to (affects COM/API compatibility).</li>
                  <li><b>Project Path</b>: Folder containing the model files; must be an existing directory where your VISSIM project is.</li>
                  <li><b>File Name</b>: Base name of the VISSIM model file (without extension).</li>
                </ul>
                <p><i>Tip:</i> If the project cannot load, verify the version selection and that the path points to the correct folder.</p>
            """),
            self._section_box("B — Scenario Setup", """
                <ul>
                  <li><b>Scenario List</b>: Which scenarios to report results for.</li>
                  <li><b>Scenario Number</b>: ID used by your scenario manager. Select <b>0</b> if you are not using scenario manager
                      (Note: If you select 0, you will need to do the process again for any different project you want to collect the results for)</li>
                  <li><b>Run Model</b>: If enabled, the tool will run simulation before extracting results; if disabled, it only reads existing results.</li>
                </ul>
            """),
            self._section_box("C — Nodes to Collect Throughput", """
                <ul>
                  <li>Defines the node sets used to count vehicles (throughput).</li>
                  <li>Each row typically includes:
                    <ul>
                      <li><b>Mainline Name</b> (for the report)</li>
                      <li><b>Direction</b> (for the report)</li>
                      <li><b>Node IDs</b> where counts should be collected</li>
                    </ul>
                  </li>
                  <li>Node IDs must match the IDs in VISSIM exactly.</li>
                </ul>
            """),
            self._section_box("D — Simulation Parameters", """
                <ul>
                  <li><b>Seeding Period</b>: warm-up time before evaluation starts.</li>
                  <li><b>Evaluation Period</b>: time window used to compute results (must be greater than seeding).</li>
                  <li><b>Number of Runs</b>, <b>Random Seed</b>, <b>Seed Increment</b>: control stochastic runs and reproducibility.</li>
                </ul>
            """),
            self._section_box("E — Results Specifics", """
                <ul>
                  <li><b>Run</b>: which run statistic to output (average, min, max, etc.).</li>
                  <li><b>Vehicle Type</b>: filter results by class (or include all).</li>
                  <li><b>Time Interval</b>: which interval(s) to use (average across intervals or a specific interval index).</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li><b>Result Directory</b>: where to save the exported report.</li>
                  <li><b>Output File Name</b>: report file name.</li>
                </ul>
            """),
        ]
        return self._main_section("Throughput Results", intro, link, boxes)

    def _build_network_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/network-results"
        intro = (
            "Network Results can be collected for different scenarios and reported in one output file. "
            "If the user wishes to report multiple results at the same time, they can do so if they are using scenario manager, "
            "otherwise, the number of the scenario should be set to 0 and only one scenario results will be reported.<br><br>"
            "Network results summarize overall system performance (and optionally environmental measures if enabled). "
            "The exported report follows the exact order of the metrics you select."
        )

        boxes = [
            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>VISSIM Version</b>: Which VISSIM major version the tool should connect to (affects COM/API compatibility).</li>
                  <li><b>Project Path</b>: Folder containing the model files; must be an existing directory where your VISSIM project is.</li>
                  <li><b>File Name</b>: Base name of the VISSIM model file (without extension).</li>
                </ul>
                <p><i>Tip:</i> If the project cannot load, verify the version selection and that the path points to the correct folder.</p>
            """),
            self._section_box("B — Scenario Setup", """
                <ul>
                  <li><b>Scenario List</b>: Which scenarios to report results for.</li>
                  <li><b>Scenario Number</b>: ID used by your scenario manager. Select <b>0</b> if you are not using scenario manager
                      (Note: If you select 0, you will need to do the process again for any different project you want to collect the results for)</li>
                  <li><b>Run Model</b>: If enabled, the tool will run simulation before extracting results; if disabled, it only reads existing results.</li>
                </ul>
            """),
            self._section_box("C — Simulation Parameters", """
                <ul>
                  <li><b>Seeding Period</b>: warm-up time before evaluation starts.</li>
                  <li><b>Evaluation Period</b>: time window used to compute results (must be greater than seeding).</li>
                  <li><b>Number of Runs</b>, <b>Random Seed</b>, <b>Seed Increment</b>: control stochastic runs and reproducibility.</li>
                </ul>
            """),
            self._section_box("D — Collected Variables", """
                <ul>
                  <li><b>Desired Metrics</b>: select the metrics you wish to collect and their order in the final output file.</li>
                  <li><b>Environment Variables</b>: If your results contain Fuel Consumption, CO Emission, or CO2 Emission and you wish to collect these variables, check the Environment Variables box. If your results do not have these variables and you try to enforce them, the software will produce an error and will not report the results.</li>
                </ul>
            """),
            self._section_box("E — Results Specifics", """
                <ul>
                  <li><b>Run</b>: which run statistic to output (average, min, max, etc.).</li>
                  <li><b>Vehicle Type</b>: filter results by class (or include all).</li>
                  <li><b>Time Interval</b>: which interval(s) to use (average across intervals or a specific interval index).</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li><b>Result Directory</b>: where to save the exported report.</li>
                  <li><b>Output File Name</b>: report file name.</li>
                </ul>
            """),
        ]
        return self._main_section("Network Results", intro, link, boxes)

    def _build_calibrator_help(self) -> CollapsibleSection:
        link = "https://your-docs-site.example/vissim-calibrator"
        intro = "The calibrator of VissiCaRe is a novel framework that enables VISSIM users to calibrate their models efficiently and in a shorter time to any other framework. First, VissiCaRe applies a novel framework incorporating <b>Hall of Fame</b> algorithm. Further more, VissiCaRe enables the user to calibrate the <b>Driving Behavior</b> Hyperparameters for as ,many types as the user needs. This automates the calibration of the VISSIM models even for complex networks. Additionally, the user can calibrate the <b>Desired Speed Distributions</b> (For both Speeds and Reduced Speed Areas). Furthermore, <b>Lane Change Distances</b> can also be calibrated by VissiCaRe! VissiCaRe automates the process of globalizing your parameters to as many scenarios as you need. If you have many scenarios (AM, PM, MD, ..., etc), you can have VissiCaRe calibrate the parameters to the Measures of Effectiveness in all these scenarios. This helps you avoid any unecessary waste of time and any unecessary Trial and Error! VissiCaRe only <b>evolve</b>. VissiCaRe gives the user a room of innovation by having a set of <b>Tunning Parameters</b> where the user can make changes to speeed up or slow down the calibration process, or fine-tune the parameters when desired thresholds are achieved. Finally, VissiCaRe enables the user to use up to 4 Meausres of Effectiveness (MoEs) to validate the calibrated parameters."

        boxes = [
            self._section_box("A — Project Setup", """
                <ul>
                  <li><b>Calibrated Scenario</b>: which scenario is treated as the calibration target (or 0 if not using scenarios).</li>
                  <li><b>Multiple Scenario</b>: This option allows you to calibrate the parameters and validate (Check) the MoEs accross multiple (As many as you need) scenarios.</li>
                </ul>
            """),
            self._section_box("B — Calibrated Variables", """
                <style>
                  table.dbp {
                    width: 100%;
                    border-collapse: separate;
                    border-spacing: 0;
                    border: 1px solid #c9d4e3;
                    border-radius: 10px;
                  }
                  table.dbp th {
                    background: #e8f4fd;
                    color: #143a5d;
                    font-weight: 900;
                    padding: 8px 10px;
                    border-bottom: 1px solid #c9d4e3;
                    text-align: left;
                    white-space: nowrap;
                  }
                  table.dbp td {
                    padding: 7px 10px;
                    border-bottom: 1px solid #edf1f7;
                    vertical-align: top;
                  }
                  table.dbp tr:nth-child(even) td {
                    background: #fafcff;
                  }
                  table.dbp tr:last-child td {
                    border-bottom: none;
                  }
                  .pill {
                    display: inline-block;
                    padding: 2px 8px;
                    border-radius: 999px;
                    border: 1px solid #c9d4e3;
                    background: #ffffff;
                    font-size: 9pt;
                    color: #143a5d;
                    white-space: nowrap;
                  }
                  .muted {
                    color: #5b6b7b;
                  }
                </style>

                <h3>Driving Behavior Parameters</h3>
                <p>
                  These are the primary driving-behavior knobs typically tuned during calibration.
                  Use the <b>Short name</b> as the internal identifier in your calibration UI, and the
                  <b>Long name</b> to match what you see in (or how it is described by) VISSIM.
                </p>
                <ul>
                  <li><b>Type</b> indicates how the value is interpreted: <span class="pill">numeric</span>, <span class="pill">integer</span>, or <span class="pill">boolean</span>.</li>
                  <li><b>Lower / Upper</b> are recommended bounds for calibration search space. Keep the rule: <b>lower &lt; start &lt; upper</b>.</li>
                  <li><b>N/A</b> means “not applicable / not bounded here” (implementation may treat it as unset, model-dependent, or handled elsewhere).</li>
                </ul>

                <table class="dbp" cellpadding="0" cellspacing="0">
                  <thead>
                    <tr>
                      <th>Short name</th>
                      <th>Long name (as found / described)</th>
                      <th>Type</th>
                      <th>Lower</th>
                      <th>Upper</th>
                    </tr>
                  </thead>
                  <tbody>
                    <tr><td><b>LookBackDistMax</b></td><td>Max. look back distance [m]</td><td><span class="pill">numeric</span></td><td>50</td><td>200</td></tr>
                    <tr><td><b>LookBackDistMin</b></td><td>Min. look back distance [m]</td><td><span class="pill">numeric</span></td><td>0</td><td>200</td></tr>
                    <tr><td><b>LookAheadDistMax</b></td><td>Max. look ahead distance [m]</td><td><span class="pill">numeric</span></td><td>100</td><td>300</td></tr>
                    <tr><td><b>LookAheadDistMin</b></td><td>Min. look ahead distance [m]</td><td><span class="pill">numeric</span></td><td>0</td><td>300</td></tr>
                    <tr><td><b>NumInteractVeh</b></td><td>Number of interaction vehicles</td><td><span class="pill">integer</span></td><td>0</td><td>99</td></tr>
                    <tr><td><b>StandDist</b></td><td>Standstill distance in front of static obstacles [m]</td><td><span class="pill">numeric</span></td><td>0</td><td>3</td></tr>
                    <tr><td><b>FreeDrivTm</b></td><td>Free driving time [s]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>IncrsAccel</b></td><td>Increased acceleration [m/s²]</td><td><span class="pill">numeric</span></td><td>1</td><td>9.99</td></tr>
                    <tr><td><b>MinCollTmGain</b></td><td>Minimum collision time gain [s]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>MinFrontRearClear</b></td><td>Minimum clearance (front/rear) [m]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>SleepDur</b></td><td>Temporary lack of attention – sleep duration</td><td><span class="pill">numeric</span></td><td>0</td><td>100</td></tr>
                    <tr><td><b>DecelRedDistOwn</b></td><td>Reduction rate for leading (own) vehicle [m]</td><td><span class="pill">numeric</span></td><td>100</td><td>200</td></tr>
                    <tr><td><b>AccDecelOwn</b></td><td>Accepted deceleration for leading (own) vehicle [m/s²]</td><td><span class="pill">numeric</span></td><td>-3</td><td>-0.5</td></tr>
                    <tr><td><b>AccDecelTrail</b></td><td>Accepted deceleration for following (trailing) vehicle [m/s²]</td><td><span class="pill">numeric</span></td><td>-10</td><td>0</td></tr>
                    <tr><td><b>SafDistFactLnChg</b></td><td>Safety distance reduction factor</td><td><span class="pill">numeric</span></td><td>0.1</td><td>0.6</td></tr>
                    <tr><td><b>CoopDecel</b></td><td>Max. deceleration for cooperative lane-change/braking [m/s²]</td><td><span class="pill">numeric</span></td><td>-6</td><td>-3</td></tr>
                    <tr><td><b>MaxDecelOwn</b></td><td>Max. deceleration for leading (own) vehicle [m/s²]</td><td><span class="pill">numeric</span></td><td>-10</td><td>-0.01</td></tr>
                    <tr><td><b>MaxDecelTrail</b></td><td>Max. deceleration for following (trailing) vehicle [m/s²]</td><td><span class="pill">numeric</span></td><td>-10</td><td>-0.01</td></tr>
                    <tr><td><b>DecelRedDistTrail</b></td><td>Reduction rate for following (trailing) vehicle [m]</td><td><span class="pill">numeric</span></td><td>1</td><td>N/A</td></tr>
                    <tr><td><b>PlatoonFollowUpGapTm</b></td><td>Platooning – follow-up gap time [s]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>PlatoonMinClear</b></td><td>Platooning – minimum clearance [m]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>W74ax</b></td><td>(Wiedemann 74) Average standstill distance</td><td><span class="pill">numeric</span></td><td>0.5</td><td>2.5</td></tr>
                    <tr><td><b>W74bxAdd</b></td><td>(Wiedemann 74) Additive factor for security distance</td><td><span class="pill">numeric</span></td><td>0.7</td><td>4.7</td></tr>
                    <tr><td><b>W74bxMult</b></td><td>(Wiedemann 74) Multiplicative factor for security distance</td><td><span class="pill">numeric</span></td><td>1</td><td>8</td></tr>
                    <tr><td><b>W99cc0</b></td><td>(Wiedemann 99) Desired distance between lead &amp; following vehicle [m]</td><td><span class="pill">numeric</span></td><td>0.6</td><td>3.05</td></tr>
                    <tr><td><b>W99cc1Distr</b></td><td>(Wiedemann 99) Headway time [s]</td><td><span class="pill">numeric</span></td><td>0.5</td><td>1.5</td></tr>
                    <tr><td><b>W99cc2</b></td><td>(Wiedemann 99) Following variation [m]</td><td><span class="pill">numeric</span></td><td>1.52</td><td>6.1</td></tr>
                    <tr><td><b>W99cc3</b></td><td>(Wiedemann 99) Threshold for entering following state [s] (negative)</td><td><span class="pill">numeric</span></td><td>-15</td><td>-4</td></tr>
                    <tr><td><b>W99cc4</b></td><td>(Wiedemann 99) Negative “following threshold” [m/s]</td><td><span class="pill">numeric</span></td><td>-0.61</td><td>-0.03</td></tr>
                    <tr><td><b>W99cc5</b></td><td>(Wiedemann 99) Positive “following threshold” [m/s]</td><td><span class="pill">numeric</span></td><td>0.03</td><td>0.61</td></tr>
                    <tr><td><b>W99cc6</b></td><td>(Wiedemann 99) Speed dependency of oscillation [1/ms]</td><td><span class="pill">numeric</span></td><td>7</td><td>15</td></tr>
                    <tr><td><b>W99cc7</b></td><td>(Wiedemann 99) Oscillation acceleration [m/s²]</td><td><span class="pill">numeric</span></td><td>0.15</td><td>0.46</td></tr>
                    <tr><td><b>W99cc8</b></td><td>(Wiedemann 99) Standstill acceleration [m/s²]</td><td><span class="pill">numeric</span></td><td>2.5</td><td>5</td></tr>
                    <tr><td><b>W99cc9</b></td><td>(Wiedemann 99) Acceleration with 80 km/h [m/s²]</td><td><span class="pill">numeric</span></td><td>0.5</td><td>2.5</td></tr>
                    <tr><td><b>LatDirChgMinTm</b></td><td>Lateral direction change – minimum time [s]</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>LatDistDrivDef</b></td><td>Lateral minimum distance at 50 km/h (default)</td><td><span class="pill">numeric</span></td><td>0</td><td>N/A</td></tr>
                    <tr><td><b>MinSpeedForLat</b></td><td>Minimum longitudinal speed for lateral movement</td><td><span class="pill">numeric</span></td><td>N/A</td><td>N/A</td></tr>
                    <tr><td><b>StandDistIsFix</b></td><td>Standstill distance (in front of static obstacles) is fix</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>VehRoutDecLookAhead</b></td><td>Vehicle routing decisions look ahead</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>ObsrvAdjLn</b></td><td>Observe adjacent lane(s)</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>DiamQueu</b></td><td>Diamond queuing</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>ConsNextTurn</b></td><td>Consider next turn</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>EnforcAbsBrakDist</b></td><td>Enforce absolute braking distance</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>UseImplicStoch</b></td><td>Use implicit stochastics</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>RecovSpeed</b></td><td>Recovery threshold speed</td><td><span class="pill">numeric</span></td><td>0</td><td>1</td></tr>
                    <tr><td><b>RecovSlow</b></td><td>Recovery slow</td><td><span class="pill">boolean</span></td><td>FALSE</td><td>TRUE</td></tr>
                    <tr><td><b>RecovSafDist</b></td><td>Recovery safety distance</td><td><span class="pill">numeric</span></td><td>0.01</td><td>N/A</td></tr>
                    <tr><td><b>NumInteractObj</b></td><td>Number of interaction objects</td><td><span class="pill">integer</span></td><td>0</td><td>10</td></tr>
                  </tbody>
                </table>

                <hr />
                <h4>Optional calibration items</h4>
                <ul>
                  <li>Optional calibration items may include speed distributions, lane-change distance, seeding, etc. (only appear when enabled).</li>
                  <li>Each calibrated item typically defines:
                    <ul>
                      <li><b>Parameter name</b></li>
                      <li><b>Lower bound</b>, <b>starting value</b>, <b>upper bound</b></li>
                    </ul>
                  </li>
                  <li><b>Rule:</b> lower &lt; start &lt; upper</li>
                </ul>
            """),
            self._section_box("C — Tuning Parameters", """
                <ul>
                  <li><b>Variant</b>: step size / aggressiveness of updates. The higher the Variant, the faster the calibration is. A balanced value is recommended.</li>
                  <li><b>Variant Update During Simulation</b>: The software allows the user to update the Variant on different breaks. The user can change the Variant after a particular number of trials. This can help fine-tuning the values of the parameters and slow or speed up the changing rate of these parameters.</li>
                  <li><b>Variant List</b>: If you enable Variant Updates, this list will show up. You can enter as many breaks as they wish and the software will apply the desired Variant during the desired trials.</li>
                  <li><b>Trials</b>: Activating the Trials informs the software to stop after these many trials, even if the desired results are not achieved. If not activated, the software will stop after 9999999999999999999999998 trials.</li>
                  <li><b>SubTrials</b>: extra sampling per trial for diversity. This makes the calibration try different combinations of the parameters from the same trial.</li>
                  <li><b>Number of Randomly Selected Sub-Trials</b>: This number dectates how many randomly selected subtrials will be nominated to be pass their updated parameters to the next Trials. When these sub-trials are selected, a rank using their errors is applied, the lowest two errors will pass to the Hall of Fame. The Hall of Fame will collect all good Sub-Trials, based on the describe processs. The same number will apply here, a subset of all the subtrials will be randomly selected, then the two Sub-Trials from all trials with the lowest errors (Within that subset) will be used to generate the next Trial.</li>
                  <li><b>Auto Stepwise Variant</b>: Activating this option enables VissiCaRe to automatically shrink the Variant when the progress of the calibration reaches the <b>% of Accepted Error</b> by the value of <b>Stepwise Value (%)</b> which is a percent of the on-going <b>Variant<b> value.</li>
                </ul>
            """),
            self._section_box("D — Validation", """
                <ul>
                  <li>Select the variables you wish to use to validate the hyperparameters.</li>
                  <li>Enter the values you need for each column in that table, the software will pick them separately. the number of rows for each variable can differ from the others, they do not have to match.</li>
                  <li>if you leave empty rows for a variable, the software will exclude these values for that variable while maintaining the values in that row for the other variables, if they have any.</li>
                  <li>Each data point in that table can have weight controlled by <b>Penalty for Variable</b>. This penalty helps the optimizer focuses more (Using that weight) on that point. If no weight it to be given, enter 1.</li>
                  <li><b>Accepted Error for Variable</b>: This the accepted percentage of error for this variable/point. This will be used to assess the progress. If all Accepted errors are met, the optimizer will stop and report the results.</li>
                </ul>
            """),
            self._section_box("E — Simulation Parameters", """
                <ul>
                  <li>Warm-up, evaluation window, number of runs, and randomization settings used during calibration evaluation.</li>
                </ul>
            """),
            self._section_box("F — Output", """
                <ul>
                  <li>Where calibration outputs (logs, summaries, best parameters) are written.</li>
                </ul>
            """),
        ]
        return self._main_section("Vissim Calibrator", intro, link, boxes)



# ============================================================
# License Tab (FIXED — Tier3 truly unlocks EVERYTHING)
#
# ✅ Works with your CURRENT server tier labels:
#    Tier1, Tier1_M, Tier2, Tier2_M, Tier3  (and optionally all_access)
#
# ✅ BIG FIX:
#    We now keep TWO lists:
#      - _raw_features      = exactly what server returns
#      - _features_effective = derived/normalized features (what gates can rely on)
#
#    So even if some parts of your app accidentally check the raw list,
#    Tier3 will still behave like ALL ACCESS because we "expand" features.
#
# ✅ Keeps endpoints exactly the same:
#    /activate, /validate, /deactivate, /seats
# ============================================================

import os
import json
import time
from typing import Optional, Dict, Any, List

try:
    import requests
except Exception:
    requests = None

from PyQt6.QtCore import Qt, QUrl
from PyQt6.QtGui import QDesktopServices
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QLabel, QTextEdit, QLineEdit, QPushButton,
    QMessageBox, QGroupBox, QTableWidget, QTableWidgetItem, QHeaderView, QScrollArea
)


class LicenseTab(QWidget):
    """
    Server:
    - POST /activate   {license_key, device_name} -> {activation_token, entitlements/features, ...}
    - POST /validate   {activation_token}         -> {entitlements/features, ...}
    - POST /deactivate {activation_token}         -> {ok}
    - POST /seats      {activation_token}         -> {seats:[...]}
    """

    API_BASE_URL = "http://138.197.70.148:8000"
    GET_LICENSE_URL = "https://example.com"

    APP_DIR_NAME = "VissiCaRe"
    TOKEN_FILE_NAME = "license_token.json"

    # ------------------------------
    # Tier labels returned by server
    # ------------------------------
    TIER_1 = "Tier1"
    TIER_1_M = "Tier1_M"
    TIER_2 = "Tier2"
    TIER_2_M = "Tier2_M"
    TIER_3 = "Tier3"  # ALL ACCESS on your server today

    # ------------------------------
    # Gate feature labels in your app
    # ------------------------------
    FEATURE_RESULTS_BASIC = "results_basic"                    # Tier 1
    FEATURE_MISTAR = "mistar_solution"                         # Tier 1_M
    FEATURE_CALIBRATOR_BASIC = "calibrator_basic"              # Tier 2
    FEATURE_MULTI_SCENARIO = "multi_scenario_calibration"      # Tier 2_M
    FEATURE_ALL_ACCESS = "all_access"                          # optional future flag

    def __init__(self, on_license_changed=None, parent=None):
        super().__init__(parent)
        self._on_license_changed = on_license_changed

        self._app_enabled: bool = True
        self._valid: bool = False
        self._reason: str = "UNACTIVATED"
        self._expires_at: str = ""

        # RAW: exactly what server sends (tiers)
        self._raw_features: List[str] = []

        # EFFECTIVE: what gates can rely on (tiers + derived feature flags)
        self._features_effective: List[str] = []

        self._seats_used: Optional[int] = None
        self._seats_limit: Optional[int] = None
        self._last_validated_ts: float = 0.0
        self._cached_seats: List[Dict[str, Any]] = []

        self.setStyleSheet(TAB_BACKGROUND_STYLE)

        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QScrollArea.Shape.NoFrame)

        try:
            scroll.setStyleSheet(GLOBAL_SCROLLBAR_STYLE)
        except Exception:
            pass

        container = QWidget()
        lay = QVBoxLayout(container)
        lay.setAlignment(Qt.AlignmentFlag.AlignTop)
        lay.setSpacing(12)
        lay.setContentsMargins(12, 12, 12, 24)

        title = QLabel("License")
        title.setStyleSheet("font-size: 16px; font-weight: 900; color:#143a5d;")
        lay.addWidget(title)

        info = QTextEdit()
        info.setReadOnly(True)
        info.setMinimumHeight(185)
        try:
            info.setStyleSheet(LOG_STYLE)
        except Exception:
            info.setStyleSheet("background:white; border:1px solid #d9e2ef; border-radius:10px; padding:8px;")

        info.setText(
            "How licensing works:\n\n"
            "• You can enter and edit your project data without a license.\n"
            "• When you try to RUN / COLLECT RESULTS / CALIBRATE, the software will validate your license.\n"
            "• You enter your license key ONE time only. After activation, it is saved on this PC.\n\n"
            "Tiers / Add-ons:\n"
            "• Tier 1 unlocks Basic Results Reporter.\n"
            "• Tier 1_M unlocks Mistar Solution (requires Tier 1).\n"
            "• Tier 2 unlocks Calibrator.\n"
            "• Tier 2_M unlocks Multiple Scenario Calibration (requires Tier 2).\n"
            "• Tier 3 unlocks EVERYTHING.\n\n"
            "Network Results is always available without a license.\n"
        )
        lay.addWidget(info)

        box = QGroupBox("Activation")
        box.setStyleSheet("QGroupBox{font-weight:900; color:#143a5d;}")

        box_lay = QVBoxLayout(box)
        box_lay.setSpacing(10)

        row1 = QHBoxLayout()

        self.key_edit = QLineEdit()
        self.key_edit.setPlaceholderText("Enter your license key (one-time activation)")
        self.key_edit.setFixedWidth(420)

        self.activate_btn = QPushButton("Activate")
        try:
            self.activate_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        except Exception:
            self.activate_btn.setStyleSheet("padding:8px 14px; border-radius:10px; font-weight:900;")
        self.activate_btn.clicked.connect(self._activate_clicked)

        self.validate_btn = QPushButton("Validate")
        try:
            self.validate_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        except Exception:
            self.validate_btn.setStyleSheet("padding:8px 14px; border-radius:10px; font-weight:900;")
        self.validate_btn.clicked.connect(self._validate_clicked)

        self.deactivate_btn = QPushButton("Deactivate this PC")
        try:
            self.deactivate_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        except Exception:
            self.deactivate_btn.setStyleSheet("padding:8px 14px; border-radius:10px; font-weight:900;")
        self.deactivate_btn.clicked.connect(self._deactivate_clicked)

        self.get_license_btn = QPushButton("Get a License")
        try:
            self.get_license_btn.setStyleSheet(PRIMARY_BUTTON_STYLE)
        except Exception:
            self.get_license_btn.setStyleSheet("padding:8px 14px; border-radius:10px; font-weight:900;")
        self.get_license_btn.clicked.connect(self._get_license_clicked)

        row1.addWidget(self.key_edit)
        row1.addWidget(self.activate_btn)
        row1.addWidget(self.validate_btn)
        row1.addWidget(self.deactivate_btn)
        row1.addWidget(self.get_license_btn)
        row1.addStretch()
        box_lay.addLayout(row1)

        self.status_label = QLabel("Status: Not activated")
        self.status_label.setStyleSheet("font-weight:900; color:#143a5d;")
        box_lay.addWidget(self.status_label)

        self.details_label = QLabel("Features: —")
        self.details_label.setStyleSheet("color:#143a5d;")
        box_lay.addWidget(self.details_label)

        self.seats_label = QLabel("Seats: —")
        self.seats_label.setStyleSheet("color:#143a5d;")
        box_lay.addWidget(self.seats_label)

        self.server_label = QLabel("")
        self.server_label.setStyleSheet("color:#7a4b00; font-weight:700;")
        box_lay.addWidget(self.server_label)

        lay.addWidget(box)

        seats_box = QGroupBox("Company Seats (Admin View)")
        seats_box.setStyleSheet("QGroupBox{font-weight:900; color:#143a5d;}")
        seats_lay = QVBoxLayout(seats_box)

        self.seats_table = QTableWidget(0, 5)
        self.seats_table.setHorizontalHeaderLabels(["Device", "Activated On", "Last Seen", "Status", "Seat ID"])
        self.seats_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.seats_table.setMinimumHeight(220)
        self.seats_table.setStyleSheet("background:white; border:1px solid #d9e2ef; border-radius:10px;")
        seats_lay.addWidget(self.seats_table)

        seat_btn_row = QHBoxLayout()
        self.refresh_seats_btn = QPushButton("Refresh Seats")
        self.refresh_seats_btn.setStyleSheet("padding:8px 14px; border-radius:10px; font-weight:900;")
        self.refresh_seats_btn.clicked.connect(self._refresh_seats_clicked)

        seat_btn_row.addWidget(self.refresh_seats_btn)
        seat_btn_row.addStretch()
        seats_lay.addLayout(seat_btn_row)

        lay.addWidget(seats_box)

        scroll.setWidget(container)
        outer.addWidget(scroll)

        self._load_and_soft_validate_on_open()

    # ------------------------------------------------------------
    # UI handlers
    # ------------------------------------------------------------

    def _get_license_clicked(self):
        QDesktopServices.openUrl(QUrl(self.GET_LICENSE_URL))

    def _activate_clicked(self):
        key = self.key_edit.text().strip()
        if not key:
            QMessageBox.warning(self, "License", "Please enter your license key.")
            self.focus_key_input()
            return

        result = self._server_activate(key)
        self._apply_server_state(result, save_cache=True)

        if self._valid and self._app_enabled:
            QMessageBox.information(self, "License", "Activation successful.")
            if self._on_license_changed:
                self._on_license_changed()
        else:
            QMessageBox.warning(self, "License", self._friendly_error_message())

    def _validate_clicked(self):
        result = self.validate_and_cache(force=True)
        if (result.get("app_enabled", True) is False):
            QMessageBox.warning(self, "License", "Software is currently disabled by the publisher.")
        elif result.get("valid"):
            QMessageBox.information(self, "License", "License is valid.")
        else:
            QMessageBox.warning(self, "License", self._friendly_error_message())

        if self._on_license_changed:
            self._on_license_changed()

    def _deactivate_clicked(self):
        token = self._read_token()
        if not token:
            QMessageBox.information(self, "License", "This PC is not activated.")
            return

        _ = self._server_deactivate(token)
        self._clear_token()
        self._apply_server_state({"app_enabled": True, "valid": False, "reason": "UNACTIVATED"}, save_cache=True)

        QMessageBox.information(self, "License", "Deactivated on this PC.")
        if self._on_license_changed:
            self._on_license_changed()

    def _refresh_seats_clicked(self):
        token = self._read_token()
        if not token:
            QMessageBox.warning(self, "Seats", "Activate a license first to view seats.")
            self.focus_key_input()
            return

        seats = self._server_fetch_seats(token)
        self._cached_seats = seats or []
        self._render_seats_table(self._cached_seats)

    # ------------------------------------------------------------
    # Public helpers (gates will use these later)
    # ------------------------------------------------------------

    def focus_key_input(self):
        self.key_edit.setFocus()
        self.key_edit.selectAll()

    def show_status(self, msg: str):
        self.status_label.setText(msg)

    def is_app_enabled(self) -> bool:
        return bool(self._app_enabled)

    def is_license_valid(self) -> bool:
        return bool(self._valid)

    def has_feature(self, feature_name: str) -> bool:
        """
        ✅ This is now bulletproof:
        - Tier3 ALWAYS returns True for ANY feature check
        - Tier labels are mapped into gate features internally
        - Also accepts all_access if you ever return it
        """
        feats = set(self._features_effective or [])

        # Absolute all-access flags
        if (self.TIER_3 in feats) or (self.FEATURE_ALL_ACCESS in feats):
            return True

        return feature_name in feats

    def validate_and_cache(self, force: bool = False) -> Dict[str, Any]:
        if not force and (time.time() - self._last_validated_ts) < 120:
            return self._current_state_dict()

        result = self._server_validate()
        self._apply_server_state(result, save_cache=True)
        return self._current_state_dict()

    # ------------------------------------------------------------
    # Internals: feature normalization (THE FIX)
    # ------------------------------------------------------------

    def _compute_effective_features(self, raw_features: List[str]) -> List[str]:
        """
        Takes server tiers (raw_features) and expands them into everything
        your app might check (tiers + gate feature labels).
        This guarantees Tier3 unlocks multi-scenario even if some code
        checks the list directly later.
        """
        raw = [str(x).strip() for x in (raw_features or []) if str(x).strip()]
        raw_set = set(raw)

        eff = set(raw_set)

        # If server ever returns "all_access" directly
        if self.FEATURE_ALL_ACCESS in raw_set:
            eff.add(self.TIER_3)

        # Tier3 => everything
        if self.TIER_3 in raw_set:
            eff.update({
                # tiers
                self.TIER_1, self.TIER_1_M, self.TIER_2, self.TIER_2_M, self.TIER_3,
                # gate features
                self.FEATURE_RESULTS_BASIC,
                self.FEATURE_MISTAR,
                self.FEATURE_CALIBRATOR_BASIC,
                self.FEATURE_MULTI_SCENARIO,
                self.FEATURE_ALL_ACCESS,
            })
            return sorted(eff)

        # Tier 1 / 1_M => Results Basic
        if (self.TIER_1 in raw_set) or (self.TIER_1_M in raw_set):
            eff.add(self.FEATURE_RESULTS_BASIC)

        # Tier 1_M => Mistar
        if self.TIER_1_M in raw_set:
            eff.add(self.FEATURE_MISTAR)

        # Tier 2 / 2_M => Calibrator basic
        if (self.TIER_2 in raw_set) or (self.TIER_2_M in raw_set):
            eff.add(self.FEATURE_CALIBRATOR_BASIC)

        # Tier 2_M => Multi-scenario
        if self.TIER_2_M in raw_set:
            eff.add(self.FEATURE_MULTI_SCENARIO)

        return sorted(eff)

    # ------------------------------------------------------------
    # Internals: local storage
    # ------------------------------------------------------------

    def _token_file_path(self) -> str:
        appdata = os.getenv("APPDATA") or os.path.expanduser("~")
        folder = os.path.join(appdata, self.APP_DIR_NAME)
        os.makedirs(folder, exist_ok=True)
        return os.path.join(folder, self.TOKEN_FILE_NAME)

    def _read_token_cache_file(self) -> Optional[Dict[str, Any]]:
        p = self._token_file_path()
        if not os.path.exists(p):
            return None
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None

    def _write_token_cache_file(self, data: Dict[str, Any]) -> None:
        p = self._token_file_path()
        try:
            with open(p, "w", encoding="utf-8") as f:
                json.dump(data, f, indent=2)
        except Exception:
            pass

    def _read_token(self) -> str:
        data = self._read_token_cache_file() or {}
        return str(data.get("activation_token", "")).strip()

    def _save_token(self, token: str, cache: Optional[Dict[str, Any]] = None) -> None:
        data = self._read_token_cache_file() or {}
        data["activation_token"] = token
        if cache is not None:
            data["cached_validation"] = cache
        self._write_token_cache_file(data)

    def _clear_token(self) -> None:
        p = self._token_file_path()
        try:
            if os.path.exists(p):
                os.remove(p)
        except Exception:
            pass

    # ------------------------------------------------------------
    # Internals: server calls
    # ------------------------------------------------------------

    def _server_ready(self) -> bool:
        if not self.API_BASE_URL:
            self.server_label.setText("Server not configured yet (API_BASE_URL is empty).")
            return False
        if requests is None:
            self.server_label.setText("Missing Python package: requests")
            return False
        self.server_label.setText("")
        return True

    def _post_json(self, endpoint: str, payload: Dict[str, Any]) -> Dict[str, Any]:
        if not self._server_ready():
            return {"app_enabled": True, "valid": False, "reason": "SERVER_NOT_CONFIGURED"}

        base = self.API_BASE_URL.rstrip("/")
        ep = endpoint if endpoint.startswith("/") else ("/" + endpoint)
        url = base + ep

        try:
            r = requests.post(url, json=payload, timeout=10)

            # If server returns HTML or non-JSON, this avoids crashing
            try:
                return r.json()
            except Exception:
                return {"app_enabled": True, "valid": False, "reason": "BAD_SERVER_RESPONSE"}

        except Exception:
            return {"app_enabled": True, "valid": False, "reason": "NO_INTERNET"}

    def _server_activate(self, license_key: str) -> Dict[str, Any]:
        device_name = os.getenv("COMPUTERNAME") or "Windows-PC"
        payload = {"license_key": license_key, "device_name": device_name}
        return self._post_json("/activate", payload)

    def _server_validate(self) -> Dict[str, Any]:
        token = self._read_token()
        if not token:
            return {"app_enabled": True, "valid": False, "reason": "UNACTIVATED"}
        return self._post_json("/validate", {"activation_token": token})

    def _server_deactivate(self, token: str) -> Dict[str, Any]:
        return self._post_json("/deactivate", {"activation_token": token})

    def _server_fetch_seats(self, token: str) -> List[Dict[str, Any]]:
        resp = self._post_json("/seats", {"activation_token": token})
        seats = resp.get("seats", [])
        return seats if isinstance(seats, list) else []

    # ------------------------------------------------------------
    # Internals: apply + render state
    # ------------------------------------------------------------

    def _apply_server_state(self, data: Dict[str, Any], save_cache: bool):
        self._app_enabled = bool(data.get("app_enabled", True))
        self._valid = bool(data.get("valid", False))
        self._reason = str(data.get("reason", "UNKNOWN"))
        self._expires_at = str(data.get("expires_at", ""))

        ent = data.get("entitlements", {}) or {}

        # Server may return features in either place
        raw_feats = ent.get("features", data.get("features", []))
        self._raw_features = raw_feats if isinstance(raw_feats, list) else []

        # ✅ THE FIX: compute effective features
        self._features_effective = self._compute_effective_features(self._raw_features)

        self._seats_used = data.get("seats_used", ent.get("seats_used"))
        self._seats_limit = data.get("seats_limit", ent.get("seats_limit"))

        token = str(data.get("activation_token", "")).strip()
        if token:
            cache = {
                "app_enabled": self._app_enabled,
                "valid": self._valid,
                "reason": self._reason,
                "expires_at": self._expires_at,
                "raw_features": list(self._raw_features or []),
                "features_effective": list(self._features_effective or []),
                "seats_used": self._seats_used,
                "seats_limit": self._seats_limit,
                "ts": time.time(),
            }
            self._save_token(token, cache=cache)

        self._last_validated_ts = time.time()
        self._refresh_ui_labels()

        if not self._app_enabled:
            self.status_label.setText("Status: DISABLED (Publisher shutdown)")
            self.status_label.setStyleSheet("font-weight:900; color:#9b1c1c;")

        if self._on_license_changed:
            self._on_license_changed()

    def _refresh_ui_labels(self):
        if not self._app_enabled:
            self.status_label.setText("Status: DISABLED (Publisher shutdown)")
            self.status_label.setStyleSheet("font-weight:900; color:#9b1c1c;")
            self.details_label.setText("Features: —")
            self.seats_label.setText("Seats: —")
            return

        self.status_label.setStyleSheet("font-weight:900; color:#143a5d;")

        if self._valid:
            exp = f"  |  Expires: {self._expires_at}" if self._expires_at else ""
            self.status_label.setText(f"Status: ACTIVE{exp}")
        else:
            if self._reason == "UNACTIVATED":
                self.status_label.setText("Status: Not activated")
            elif self._reason == "NO_INTERNET":
                self.status_label.setText("Status: Internet required to validate")
            elif self._reason == "EXPIRED":
                self.status_label.setText("Status: Expired")
            elif self._reason == "SEAT_LIMIT":
                self.status_label.setText("Status: No seats available")
            elif self._reason == "SERVER_NOT_CONFIGURED":
                self.status_label.setText("Status: Server not configured")
            else:
                self.status_label.setText(f"Status: Not valid ({self._reason})")

        # Show RAW tiers to user (clean), but tag Tier3 as All Access
        if self._raw_features:
            raw = list(self._raw_features)
            if self.TIER_3 in raw:
                # make it obvious for you + users
                raw = [("Tier3 (All Access)" if x == self.TIER_3 else x) for x in raw]
            self.details_label.setText("Features: " + ", ".join(raw))
        else:
            self.details_label.setText("Features: —")

        if self._seats_used is not None and self._seats_limit is not None:
            self.seats_label.setText(f"Seats: {self._seats_used} / {self._seats_limit}")
        else:
            self.seats_label.setText("Seats: —")

    # ✅ NEW: seat-number extraction + full-seat rendering (Seat 1..Seat N always shown)
    def _seat_number_from_payload(self, seat: Dict[str, Any]) -> Optional[int]:
        """
        Tries to discover a seat index from the server payload, if your server provides one.
        Accepts common keys. Returns 1-based index if found, else None.
        """
        if not isinstance(seat, dict):
            return None

        # Common possibilities
        for k in ("seat_number", "seat_no", "seat_index", "seat_idx", "seat"):
            v = seat.get(k, None)
            try:
                if v is None or v == "":
                    continue
                n = int(v)
                if n >= 1:
                    return n
            except Exception:
                pass

        # If seat_id itself is numeric, use it
        v = seat.get("seat_id", None)
        try:
            if v is None or v == "":
                return None
            n = int(v)
            if n >= 1:
                return n
        except Exception:
            return None

    def _render_seats_table(self, seats: List[Dict[str, Any]]):
        # Determine how many seats to show (Seat 1..Seat N)
        limit = None
        try:
            if self._seats_limit is not None:
                limit = int(self._seats_limit)
        except Exception:
            limit = None

        if limit is None:
            # fallback: at least show all seats returned
            limit = max(0, len(seats or []))

        # Build N slots (Seat 1..Seat N), then fill used ones
        slots: List[Optional[Dict[str, Any]]] = [None] * limit

        # 1) Place seats that explicitly specify their seat number
        unplaced: List[Dict[str, Any]] = []
        for seat in (seats or []):
            sn = self._seat_number_from_payload(seat)
            if sn is not None and 1 <= sn <= limit and slots[sn - 1] is None:
                slots[sn - 1] = seat
            else:
                unplaced.append(seat)

        # 2) Place remaining seats in first available empty slots (stable order)
        for seat in unplaced:
            placed = False
            for i in range(limit):
                if slots[i] is None:
                    slots[i] = seat
                    placed = True
                    break
            if not placed:
                break

        # Render table rows = limit
        self.seats_table.setRowCount(0)

        for i in range(1, limit + 1):
            seat = slots[i - 1]

            # Defaults for unused seats
            device = ""
            activated_on = ""
            last_seen = ""
            status = "Unused"
            seat_id_display = f"Seat {i}"

            if isinstance(seat, dict) and seat:
                device = str(seat.get("device_name", "") or "")
                activated_on = str(seat.get("activated_on", "") or "")
                last_seen = str(seat.get("last_seen", "") or "")
                status_raw = str(seat.get("status", "") or "").strip()
                status = status_raw if status_raw else "Used"

                # Keep original seat_id if you want, but still show Seat #
                sid = str(seat.get("seat_id", "") or "").strip()
                if sid:
                    seat_id_display = f"Seat {i} ({sid})"
                else:
                    seat_id_display = f"Seat {i}"

            r = self.seats_table.rowCount()
            self.seats_table.insertRow(r)

            for c, v in enumerate([device, activated_on, last_seen, status, seat_id_display]):
                item = QTableWidgetItem(v)
                item.setFlags(item.flags() ^ Qt.ItemFlag.ItemIsEditable)
                self.seats_table.setItem(r, c, item)

    def _friendly_error_message(self) -> str:
        if not self._app_enabled:
            return "Software is currently disabled by the publisher."
        if self._reason == "UNACTIVATED":
            return "Please activate using your license key."
        if self._reason == "NO_INTERNET":
            return "Internet is required to validate your license."
        if self._reason == "EXPIRED":
            return "Your license has expired."
        if self._reason == "SEAT_LIMIT":
            return "No seats available. Ask your company admin to deactivate a seat."
        if self._reason == "SERVER_NOT_CONFIGURED":
            return "License server is not configured yet."
        if self._reason == "BAD_SERVER_RESPONSE":
            return "License server returned an invalid response."
        return "License validation failed."

    def _current_state_dict(self) -> Dict[str, Any]:
        return {
            "app_enabled": self._app_enabled,
            "valid": self._valid,
            "reason": self._reason,
            "expires_at": self._expires_at,
            "raw_features": list(self._raw_features or []),
            "features": list(self._features_effective or []),
            "seats_used": self._seats_used,
            "seats_limit": self._seats_limit,
        }

    def _load_and_soft_validate_on_open(self):
        """
        On open:
        - Load cached validation if any (so UI shows something immediately).
        - Do NOT force online check (you wanted gating on actions).
        """
        cache_file = self._read_token_cache_file() or {}
        cached = cache_file.get("cached_validation") if isinstance(cache_file, dict) else None

        if isinstance(cached, dict):
            self._app_enabled = bool(cached.get("app_enabled", True))
            self._valid = bool(cached.get("valid", False))
            self._reason = str(cached.get("reason", "UNACTIVATED"))
            self._expires_at = str(cached.get("expires_at", ""))

            self._raw_features = cached.get("raw_features", cached.get("features", []))
            if not isinstance(self._raw_features, list):
                self._raw_features = []

            # If old cache doesn't have effective, compute it
            eff = cached.get("features_effective")
            if isinstance(eff, list) and eff:
                self._features_effective = eff
            else:
                self._features_effective = self._compute_effective_features(self._raw_features)

            self._seats_used = cached.get("seats_used")
            self._seats_limit = cached.get("seats_limit")
            self._last_validated_ts = float(cached.get("ts", 0.0))

        else:
            token = self._read_token()
            if token:
                self._reason = "NEEDS_VALIDATION"
            else:
                self._reason = "UNACTIVATED"
                self._valid = False
                self._raw_features = []
                self._features_effective = []

        self._refresh_ui_labels()



# ============================================================
# Main Window
# ============================================================

class VissiCaReMainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("VissiCaRe")
        if os.path.exists(ICON_PATH):
            self.setWindowIcon(QIcon(ICON_PATH))

        self._layout_path: Optional[str] = None
        self._build_menu_bar()
        self.statusBar()

        self.tabs = QTabWidget()
        self.tabs.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #c9d4e3; border-radius: 10px; }
            QTabBar::tab {
                padding: 8px 14px;
                margin: 2px;
                border: 1px solid #c9d4e3;
                border-radius: 10px;
                background: #ffffff;
                color: #143a5d;
                font-weight: 900;
            }
            QTabBar::tab:selected {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #fff7d6, stop:1 #e4f0ff);
            }
        """)

        self.nodes_tab = NodesTab()
        self.links_tab = LinksTab()
        self.tt_tab = TravelTimeTab()
        self.throughput_tab = ThroughputTab()
        self.network_tab = NetworkTab()     # NEW
        self.calibrator_tab = CalibratorTab()
        self.help_tab = HelpTab()
        self.license_tab = LicenseTab(on_license_changed=self._refresh_license_dependent_ui)

        for tab in (
            self.nodes_tab,
            self.links_tab,
            self.tt_tab,
            self.throughput_tab,
            self.network_tab,
            self.calibrator_tab,
            self.help_tab,
            self.license_tab,
        ):
            _force_input_bg_in_widget(tab)

        self.tabs.addTab(self.nodes_tab, "Nodes Results")
        self.tabs.addTab(self.links_tab, "Links Results")
        self.tabs.addTab(self.tt_tab, "Travel Time Results")
        self.tabs.addTab(self.throughput_tab, "Troughput Results")
        self.tabs.addTab(self.network_tab, "Network Results")  # NEW (after throughput, before help)
        self.tabs.addTab(self.calibrator_tab, "Vissim Calibrator")   # NEW
        self.tabs.addTab(self.help_tab, "Help")
        self.tabs.addTab(self.license_tab, "License")

        self.tabs.currentChanged.connect(lambda _i: self._refresh_license_dependent_ui())
        self.setCentralWidget(self.tabs)
        self.resize(1200, 780)

    def _refresh_license_dependent_ui(self):
        self.nodes_tab.refresh_license_ui()
    # ============================
    # File Menu + Layout (.layx)
    # ============================

    def _build_menu_bar(self):
        mb = self.menuBar()
        file_menu = mb.addMenu("File")

        act_new = QAction("New Project", self)
        act_new.setShortcut(QKeySequence.StandardKey.New)

        act_open = QAction("Open Project…", self)
        act_open.setShortcut(QKeySequence.StandardKey.Open)

        act_save = QAction("Save Project", self)
        act_save.setShortcut(QKeySequence.StandardKey.Save)

        act_save_as = QAction("Save Project As…", self)
        act_save_as.setShortcut(QKeySequence.StandardKey.SaveAs)

        act_exit = QAction("Exit", self)
        act_exit.setShortcut(QKeySequence.StandardKey.Quit)

        file_menu.addAction(act_new)
        file_menu.addAction(act_open)
        file_menu.addSeparator()
        file_menu.addAction(act_save)
        file_menu.addAction(act_save_as)
        file_menu.addSeparator()
        file_menu.addAction(act_exit)

        act_new.triggered.connect(self._new_project)
        act_open.triggered.connect(self._open_layout_dialog)
        act_save.triggered.connect(self._save_layout)
        act_save_as.triggered.connect(self._save_layout_as)
        act_exit.triggered.connect(self.close)


    def _collect_layout(self) -> dict:
        return {
            "layout_version": 1,
            "license_state": {"level": LicenseState.level, "key": LicenseState.key},
            "tabs": {
                "nodes": self.nodes_tab.export_state(),
                "links": self.links_tab.export_state(),
                "travel_time": self.tt_tab.export_state(),
                "throughput": self.throughput_tab.export_state(),
                "network": self.network_tab.export_state(),
                "calibrator": self.calibrator_tab.export_state(),
                "license_tab": self.license_tab.export_state(),
            },
        }

    def _apply_layout(self, layout: dict):
        # restore license first (so tab locks work)
        ls = layout.get("license_state", {}) or {}
        LicenseState.set_license(key=str(ls.get("key", "")), level=str(ls.get("level", "Basic")))

        tabs = layout.get("tabs", {}) or {}
        if "nodes" in tabs:
            self.nodes_tab.import_state(tabs["nodes"])
        if "links" in tabs:
            self.links_tab.import_state(tabs["links"])
        if "travel_time" in tabs:
            self.tt_tab.import_state(tabs["travel_time"])
        if "throughput" in tabs:
            self.throughput_tab.import_state(tabs["throughput"])
        if "network" in tabs:
            self.network_tab.import_state(tabs["network"])
        if "calibrator" in tabs:
            self.calibrator_tab.import_state(tabs["calibrator"])
        if "license_tab" in tabs:
            self.license_tab.import_state(tabs["license_tab"])

        self._refresh_license_dependent_ui()

    def _save_layout_to_path(self, path: str):
        data = self._collect_layout()
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
        self._layout_path = path
        self.statusBar().showMessage(f"Saved layout: {path}", 4000)
        self._layout_path = path
        self._set_dirty(False)

    def _open_layout_from_path(self, path: str):
        self._loading = True
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)

            self._apply_layout(data)
            self._layout_path = path
            self._set_dirty(False)
        finally:
            self._loading = False


    def _save_layout(self):
        if not self._layout_path:
            self._save_layout_as()
            return
        try:
            self._save_layout_to_path(self._layout_path)
        except Exception as e:
            QMessageBox.critical(self, "Save Layout", f"Failed to save layout:\n{e}")

    def _save_layout_as(self):
        path, _ = QFileDialog.getSaveFileName(
            self, "Save Layout",
            "Layout.layx",
            "Layout Files (*.layx);;All Files (*)"
        )
        if not path:
            return
        if not path.lower().endswith(".layx"):
            path += ".layx"
        try:
            self._save_layout_to_path(path)
        except Exception as e:
            QMessageBox.critical(self, "Save Layout As", f"Failed to save layout:\n{e}")

    def _open_layout_dialog(self):
        if not self._maybe_save():
            return

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Open Project",
            "",
            "Layout Files (*.layx);;All Files (*)"
        )
        if not path:
            return

        try:
            self._open_layout_from_path(path)
        except Exception as e:
            QMessageBox.critical(self, "Open Project", f"Failed to load project:\n{e}")


    def _new_project(self):
        if not self._maybe_save():
            return
        # Rebuild tabs (keeps your original constructors + defaults)
        current = self.tabs.currentIndex()

        self.nodes_tab = NodesTab()
        self.links_tab = LinksTab()
        self.tt_tab = TravelTimeTab()
        self.throughput_tab = ThroughputTab()
        self.network_tab = NetworkTab()
        self.calibrator_tab = CalibratorTab()
        self.help_tab = HelpTab()
        self.license_tab = LicenseTab(on_license_changed=self._refresh_license_dependent_ui)

        self.tabs.clear()
        self.tabs.addTab(self.nodes_tab, "Nodes Results")
        self.tabs.addTab(self.links_tab, "Links Results")
        self.tabs.addTab(self.tt_tab, "Travel Time Results")
        self.tabs.addTab(self.throughput_tab, "Troughput Results")
        self.tabs.addTab(self.network_tab, "Network Results")
        self.tabs.addTab(self.calibrator_tab, "Vissim Calibrator")
        self.tabs.addTab(self.help_tab, "Help")
        self.tabs.addTab(self.license_tab, "License")

        self._layout_path = None
        self._refresh_license_dependent_ui()
        self.tabs.setCurrentIndex(min(current, self.tabs.count() - 1))
        self.statusBar().showMessage("New project started.", 3000)




from pathlib import Path
import sys
import math

from PyQt6.QtCore import Qt, QRect, QTimer
from PyQt6.QtGui import QPixmap, QPainter, QColor, QFont, QIcon, QImage
from PyQt6.QtWidgets import QApplication, QSplashScreen


def resource_path(filename: str) -> Path:
    here = Path(__file__).resolve().parent
    p = here / filename
    if p.exists():
        return p
    return Path(sys.argv[0]).resolve().parent / filename


def trim_transparent(pm: QPixmap) -> QPixmap:
    """Remove transparent padding around an icon/pixmap."""
    if pm.isNull():
        return pm

    img = pm.toImage().convertToFormat(QImage.Format.Format_ARGB32)
    w, h = img.width(), img.height()

    left, right = w, -1
    top, bottom = h, -1

    for y in range(h):
        for x in range(w):
            if QColor(img.pixel(x, y)).alpha() > 0:
                if x < left: left = x
                if x > right: right = x
                if y < top: top = y
                if y > bottom: bottom = y

    if right < left or bottom < top:
        return pm

    return pm.copy(QRect(left, top, right - left + 1, bottom - top + 1))


def _scaled_size_keep_aspect(pm: QPixmap, box: int) -> tuple[int, int]:
    """Size of pm when scaled to fit inside box x box with KeepAspectRatio."""
    if pm.isNull():
        return box, box
    pw, ph = pm.width(), pm.height()
    s = min(box / pw, box / ph)
    return int(pw * s), int(ph * s)





# ============================
# Start Splash (SIDE-BY-SIDE) — DYNAMIC BACKGROUND + TRAFFIC FLOW
# ✅ test logo 1.5x larger
# (implemented by multiplying test_SCALE by 1.5)
# ============================

import sys
import math

from PyQt6.QtCore import Qt, QRect, QTimer, QPointF
from PyQt6.QtGui import (
    QPixmap, QPainter, QFont, QColor, QIcon,
    QLinearGradient, QRadialGradient, QPen, QBrush, QPainterPath
)
from PyQt6.QtWidgets import QApplication, QSplashScreen


class BrandedSplash(QSplashScreen):
    # VissiCaRe scale (kept as you had)
    VissiCaRe_SCALE = 0.50

    # ✅ test 1.5x larger than the prior value (0.62 * 1.5 = 0.93)
    test_SCALE = 0.93

    # Gentle pulsing
    PULSE_AMPLITUDE = 0.020   # 2% pulse (very gentle)
    PULSE_STEP = 0.12         # animation speed

    # ===== Modern background tuning =====
    BG_GRID_ALPHA = 18
    BG_SHAPES_ALPHA = 32
    BG_PARTICLES_ALPHA = 55
    BG_VEHICLES_ALPHA = 50

    BG_GRID_SPACING = 34
    BG_PARTICLE_COUNT = 28
    BG_NODE_COUNT = 14
    BG_VEHICLE_COUNT = 7

    # Parallax speeds (subtle)
    BG_GRID_SPEED = 0.25
    BG_SHAPES_SPEED = 0.18
    BG_PARTICLES_SPEED = 0.55

    # Traffic speed (pixels per "tick unit")
    TRAFFIC_SPEED_BASE = 6.0

    def __init__(self, base_pixmap: QPixmap, big_icon: QPixmap, small_icon: QPixmap):
        super().__init__(base_pixmap)
        self.big_icon = big_icon
        self.small_icon = small_icon

        self._t = 0.0
        self._pulse_timer = QTimer(self)
        self._pulse_timer.timeout.connect(self._on_pulse)
        self._pulse_timer.start(30)  # ~33 fps

    def _on_pulse(self):
        self._t += self.PULSE_STEP
        self.update()  # triggers repaint -> drawContents runs again

    # ----------------------------
    # Modern dynamic background
    # ----------------------------
    def _draw_modern_background(self, painter: QPainter, w: int, h: int):
        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        # Soft, modern gradient base (white -> very light blue)
        base_grad = QLinearGradient(0, 0, 0, h)
        base_grad.setColorAt(0.0, QColor(255, 255, 255))
        base_grad.setColorAt(0.55, QColor(247, 250, 255))
        base_grad.setColorAt(1.0, QColor(240, 246, 255))
        painter.fillRect(QRect(0, 0, w, h), QBrush(base_grad))

        # Subtle "AI blob" radial gradients drifting slowly
        drift_x = math.sin(self._t * self.BG_SHAPES_SPEED) * 22.0
        drift_y = math.cos(self._t * self.BG_SHAPES_SPEED * 0.9) * 18.0

        def blob(cx, cy, r, c1: QColor, c2: QColor):
            g = QRadialGradient(QPointF(cx, cy), r)
            g.setColorAt(0.0, c1)
            g.setColorAt(1.0, c2)
            painter.fillRect(QRect(0, 0, w, h), QBrush(g))

        painter.setOpacity(0.85)
        blob(
            w * 0.28 + drift_x,
            h * 0.36 + drift_y,
            min(w, h) * 0.55,
            QColor(220, 238, 255, self.BG_SHAPES_ALPHA),
            QColor(220, 238, 255, 0),
        )
        blob(
            w * 0.78 - drift_x,
            h * 0.22 + drift_y * 0.7,
            min(w, h) * 0.42,
            QColor(255, 236, 214, int(self.BG_SHAPES_ALPHA * 0.85)),
            QColor(255, 236, 214, 0),
        )
        blob(
            w * 0.65 + drift_x * 0.6,
            h * 0.80 - drift_y,
            min(w, h) * 0.50,
            QColor(230, 230, 255, int(self.BG_SHAPES_ALPHA * 0.80)),
            QColor(230, 230, 255, 0),
        )
        painter.setOpacity(1.0)

        # Subtle grid
        grid_off = (self._t * self.BG_GRID_SPEED * 6.0) % self.BG_GRID_SPACING
        pen = QPen(QColor(20, 58, 93, self.BG_GRID_ALPHA))
        pen.setWidth(1)
        painter.setPen(pen)

        for x in range(-self.BG_GRID_SPACING, w + self.BG_GRID_SPACING, self.BG_GRID_SPACING):
            xx = int(x + grid_off)
            painter.drawLine(xx, 0, xx, h)
        for y in range(-self.BG_GRID_SPACING, h + self.BG_GRID_SPACING, self.BG_GRID_SPACING):
            yy = int(y + grid_off * 0.6)
            painter.drawLine(0, yy, w, yy)

        # AI network nodes + connecting lines
        painter.setOpacity(0.70)
        nodes = []
        for i in range(self.BG_NODE_COUNT):
            fx = (0.12 + 0.76 * ((i * 37) % 100) / 100.0)
            fy = (0.12 + 0.76 * ((i * 59) % 100) / 100.0)
            ox = math.sin(self._t * 0.35 + i * 0.9) * 10.0
            oy = math.cos(self._t * 0.33 + i * 1.1) * 8.0
            nodes.append(QPointF(fx * w + ox, fy * h + oy))

        line_pen = QPen(QColor(20, 58, 93, 28))
        line_pen.setWidth(2)
        painter.setPen(line_pen)
        for i in range(len(nodes)):
            j = (i * 3 + 5) % len(nodes)
            k = (i * 7 + 2) % len(nodes)
            painter.drawLine(nodes[i], nodes[j])
            painter.drawLine(nodes[i], nodes[k])

        for i, p in enumerate(nodes):
            r = 3.0 + (i % 3) * 1.2
            a = int(55 + 40 * (0.5 + 0.5 * math.sin(self._t * 0.9 + i)))
            painter.setBrush(QBrush(QColor(20, 58, 93, a)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(p, r, r)

        painter.setOpacity(1.0)

        # Floating particles
        painter.save()
        painter.setOpacity(0.85)
        for i in range(self.BG_PARTICLE_COUNT):
            bx = ((i * 83) % 1000) / 1000.0
            by = ((i * 149) % 1000) / 1000.0
            px = (bx * w + (self._t * self.BG_PARTICLES_SPEED * 22.0) + i * 7.0) % (w + 40) - 20
            py = (by * h - (self._t * self.BG_PARTICLES_SPEED * 18.0) - i * 9.0) % (h + 40) - 20
            rr = 1.2 + (i % 4) * 0.5
            aa = int(self.BG_PARTICLES_ALPHA * (0.6 + 0.4 * math.sin(self._t * 0.7 + i)))
            painter.setBrush(QBrush(QColor(20, 58, 93, aa)))
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(QPointF(px, py), rr, rr)
        painter.restore()

        # ---- TRAFFIC: vehicles moving side-to-side (loop) ----
        def draw_vehicle(cx, cy, scale, direction: int, tint: QColor, alpha: int):
            body_w = 92 * scale
            body_h = 28 * scale
            wheel_r = 6.5 * scale

            x = cx
            y = cy

            path = QPainterPath()
            r = 10 * scale
            path.addRoundedRect(x - body_w / 2, y - body_h / 2, body_w, body_h, r, r)

            cab = QPainterPath()
            if direction > 0:
                cab.addRoundedRect(x - body_w * 0.10, y - body_h * 0.95, body_w * 0.40, body_h * 0.70, 8 * scale, 8 * scale)
                sensor_x = x + body_w * 0.40
            else:
                cab.addRoundedRect(x - body_w * 0.30, y - body_h * 0.95, body_w * 0.40, body_h * 0.70, 8 * scale, 8 * scale)
                sensor_x = x - body_w * 0.40

            w1 = QPointF(x - body_w * 0.28, y + body_h * 0.45)
            w2 = QPointF(x + body_w * 0.28, y + body_h * 0.45)

            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QBrush(QColor(tint.red(), tint.green(), tint.blue(), alpha)))
            painter.drawPath(path)

            painter.setBrush(QBrush(QColor(255, 255, 255, int(alpha * 0.70))))
            painter.drawPath(cab)

            painter.setBrush(QBrush(QColor(20, 58, 93, int(alpha * 1.25))))
            painter.drawEllipse(w1, wheel_r, wheel_r)
            painter.drawEllipse(w2, wheel_r, wheel_r)

            painter.setBrush(QBrush(QColor(255, 200, 120, int(alpha * 1.05))))
            painter.drawEllipse(QPointF(sensor_x, y - body_h * 0.05), 2.3 * scale, 2.3 * scale)

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setOpacity(0.90)

        lane_base = h * 0.84
        lane_gap = 18

        travel = self._t * self.TRAFFIC_SPEED_BASE
        margin = 140

        for i in range(self.BG_VEHICLE_COUNT):
            lane = i % 3
            direction = 1 if (i % 2 == 0) else -1
            scale = 0.70 + (i % 3) * 0.10

            speed = (1.0 + (i % 4) * 0.18) * (1.0 if direction > 0 else 0.92)
            base = ((i * 211) % 1000) / 1000.0
            y = lane_base + lane * lane_gap + (0 if direction > 0 else 8)

            if direction > 0:
                x = ((base * (w + 2 * margin)) + travel * speed) % (w + 2 * margin) - margin
            else:
                x = (w + margin) - (((base * (w + 2 * margin)) + travel * speed) % (w + 2 * margin))

            tint = QColor(20, 58, 93)
            if i % 3 == 1:
                tint = QColor(35, 80, 120)
            elif i % 3 == 2:
                tint = QColor(18, 48, 78)

            draw_vehicle(x, y, scale, direction, tint, self.BG_VEHICLES_ALPHA)

        painter.restore()
        painter.restore()

    # ----------------------------
    # Main contents
    # ----------------------------
    def drawContents(self, painter: QPainter):
        pm = self.pixmap()
        if pm.isNull():
            return

        w, h = pm.width(), pm.height()

        painter.save()
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)

        self._draw_modern_background(painter, w, h)

        pad_top = 20
        pad_side_min = 30
        col_gap = 18  # close together

        label_h = 28
        gap1 = 10
        tagline_gap = 18
        tagline_h = 32

        pulse = 1.0 + (self.PULSE_AMPLITUDE * math.sin(self._t))

        left_box_base = int(w * 0.52 * self.VissiCaRe_SCALE)
        left_box = int(left_box_base * pulse)

        right_w = max(260, int(w * 0.30))
        right_icon_box = int(min(right_w * 0.98, left_box_base * 0.92) * self.test_SCALE)

        big = QPixmap()
        if not self.big_icon.isNull():
            big = self.big_icon.scaled(
                left_box, left_box,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )

        small = QPixmap()
        if not self.small_icon.isNull():
            small = self.small_icon.scaled(
                right_icon_box, right_icon_box,
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.SmoothTransformation
            )

        content_bottom = h - (tagline_gap + tagline_h + pad_top)

        # center both logos as a group
        group_w = left_box_base + col_gap + right_w
        start_x = max(pad_side_min, (w - group_w) // 2)

        left_x = start_x
        right_x = left_x + left_box_base + col_gap

        if not big.isNull():
            by = pad_top + max(0, (content_bottom - pad_top - big.height()) // 2)
            painter.drawPixmap(left_x + (left_box_base - big.width()) // 2, by, big)

        right_block_h = label_h + gap1 + (small.height() if not small.isNull() else right_icon_box)
        ry = pad_top + max(0, (content_bottom - pad_top - right_block_h) // 2)

        painter.setPen(QColor(20, 58, 93))

        painter.setFont(QFont("Segoe UI", 14, QFont.Weight.Bold))
        painter.drawText(
            QRect(right_x, ry, right_w, label_h),
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
            " "
        )
        ry += label_h + gap1

        if not small.isNull():
            painter.drawPixmap(right_x + (right_w - small.width()) // 2, ry, small)

        painter.setFont(QFont("Segoe UI", 13, QFont.Weight.Bold))
        painter.drawText(
            QRect(0, h - (tagline_h + pad_top), w, tagline_h),
            Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignVCenter,
            " "
        )

        painter.restore()


def main():
    app = QApplication(sys.argv)
    apply_global_scrollbar_style(app)

    big_icon_path = resource_path("VissiCaRe_icon_splash.ico")
    small_icon_path = resource_path("test.ico")

    big_pm = trim_transparent(QIcon(str(big_icon_path)).pixmap(1024, 1024))
    small_pm = trim_transparent(QIcon(str(small_icon_path)).pixmap(1024, 1024))

    SPLASH_W = 980

    pad_top = 20
    label_h = 28
    gap1 = 10
    tagline_gap = 18
    tagline_h = 32
    bottom_pad = 20

    left_box = int(SPLASH_W * 0.52 * BrandedSplash.VissiCaRe_SCALE)

    right_w = max(260, int(SPLASH_W * 0.30))
    right_icon_box = int(min(right_w * 0.98, left_box * 0.92) * BrandedSplash.test_SCALE)

    pulse_max = 1.0 + BrandedSplash.PULSE_AMPLITUDE
    left_box_max = int(left_box * pulse_max)

    _, big_h = _scaled_size_keep_aspect(big_pm, left_box_max)
    _, small_h = _scaled_size_keep_aspect(small_pm, right_icon_box)

    right_block_h = label_h + gap1 + small_h
    content_h = max(big_h, right_block_h)

    SPLASH_H = max(
        520,
        pad_top + content_h + tagline_gap + tagline_h + bottom_pad
    )

    base = QPixmap(SPLASH_W, SPLASH_H)
    base.fill(QColor(255, 255, 255))

    w = VissiCaReMainWindow()

    splash = BrandedSplash(base, big_pm, small_pm)
    splash.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)

    screen_geo = app.primaryScreen().availableGeometry()
    splash.move(
        screen_geo.center().x() - SPLASH_W // 2,
        screen_geo.center().y() - SPLASH_H // 2
    )

    splash.show()
    splash.raise_()
    app.processEvents()

    def show_main_and_close_splash():
        splash._pulse_timer.stop()
        w.showMaximized()
        splash.finish(w)

    QTimer.singleShot(10000, show_main_and_close_splash)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()