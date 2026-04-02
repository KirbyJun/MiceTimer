"""
MiceTimer - 实验计时计数软件
PySide6 GUI, Windows-compatible
"""
from __future__ import annotations

import json
import os
import sys
import time
from copy import deepcopy
from datetime import datetime
from typing import Any, Dict, List, Optional

from PySide6.QtCore import (
    QSize,
    Qt,
    QTimer,
    Signal,
)
from PySide6.QtGui import (
    QColor,
    QFont,
    QKeySequence,
    QPalette,
)
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QFormLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QKeySequenceEdit,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QSizePolicy,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)
from openpyxl import Workbook

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------
def _app_dir() -> str:
    """Return the directory where the application executable (or script) lives."""
    if getattr(sys, "frozen", False):
        # PyInstaller bundle
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()
DATA_DIR = os.path.join(APP_DIR, "data")
AUTOSAVE_DIR = os.path.join(DATA_DIR, "autosave")
EXPORT_DIR = os.path.join(DATA_DIR, "export")
RECOVERY_FILE = os.path.join(AUTOSAVE_DIR, "recovery.json")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
RECENT_TEMPLATE_FILE = os.path.join(DATA_DIR, "recent_template.json")
TEMPLATES_DIR = os.path.join(DATA_DIR, "templates")

for _d in (AUTOSAVE_DIR, EXPORT_DIR, TEMPLATES_DIR):
    os.makedirs(_d, exist_ok=True)

MAX_FILENAME_PART_LENGTH = 30
MAX_RECENT_TEMPLATES = 10

# ---------------------------------------------------------------------------
# Default paradigms / items
# ---------------------------------------------------------------------------
DEFAULT_PARADIGMS = [
    {
        "name": "三箱社交",
        "items": [
            {"kind": "timer", "name": "陌生鼠侧探索时间"},
            {"kind": "timer", "name": "空侧探索时间"},
            {"kind": "counter", "name": "陌生鼠侧进入次数"},
            {"kind": "counter", "name": "空侧进入次数"},
        ],
    },
    {
        "name": "自由社交",
        "items": [
            {"kind": "timer", "name": "社交时间"},
            {"kind": "timer", "name": "非社交时间"},
            {"kind": "counter", "name": "社交接触次数"},
        ],
    },
    {
        "name": "都做",
        "items": [
            {"kind": "timer", "name": "陌生鼠侧探索时间"},
            {"kind": "timer", "name": "空侧探索时间"},
            {"kind": "timer", "name": "社交时间"},
            {"kind": "counter", "name": "陌生鼠侧进入次数"},
            {"kind": "counter", "name": "空侧进入次数"},
            {"kind": "counter", "name": "社交接触次数"},
        ],
    },
]

DEFAULT_SETTINGS: Dict[str, Any] = {
    "hotkeys": {
        "start_stop": "F5",
        "reset": "F6",
        "export": "F7",
    },
}

# ---------------------------------------------------------------------------
# Data model
# ---------------------------------------------------------------------------

class Item:
    def __init__(self, kind: str, name: str):
        self.kind: str = kind          # "timer" | "counter"
        self.name: str = name
        self.elapsed: float = 0.0      # seconds (timer)
        self.count: int = 0            # (counter)
        self.running: bool = False     # timer only
        self.last_start_ts: Optional[float] = None

    def to_dict(self) -> dict:
        return {
            "kind": self.kind,
            "name": self.name,
            "elapsed": self.elapsed,
            "count": self.count,
            "running": self.running,
            "last_start_ts": self.last_start_ts,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Item":
        it = cls(d["kind"], d["name"])
        it.elapsed = float(d.get("elapsed", 0))
        it.count = int(d.get("count", 0))
        it.running = bool(d.get("running", False))
        it.last_start_ts = d.get("last_start_ts", None)
        return it

    def reset(self):
        self.elapsed = 0.0
        self.count = 0
        self.running = False
        self.last_start_ts = None

    def current_elapsed(self) -> float:
        if self.running and self.last_start_ts is not None:
            return self.elapsed + (time.perf_counter() - self.last_start_ts)
        return self.elapsed


class Experiment:
    def __init__(self):
        self.date: str = datetime.now().strftime("%Y-%m-%d")
        self.operator: str = ""
        self.mouse_id: str = ""
        self.group: str = ""
        self.paradigm: str = ""
        self.remark: str = ""
        self.items: List[Item] = []
        self.events: List[dict] = []
        self.started: bool = False
        self.start_time: Optional[str] = None

    def to_dict(self) -> dict:
        return {
            "date": self.date,
            "operator": self.operator,
            "mouse_id": self.mouse_id,
            "group": self.group,
            "paradigm": self.paradigm,
            "remark": self.remark,
            "items": [it.to_dict() for it in self.items],
            "events": self.events,
            "started": self.started,
            "start_time": self.start_time,
        }

    @classmethod
    def from_dict(cls, d: dict) -> "Experiment":
        exp = cls()
        exp.date = d.get("date", datetime.now().strftime("%Y-%m-%d"))
        exp.operator = d.get("operator", "")
        exp.mouse_id = d.get("mouse_id", "")
        exp.group = d.get("group", "")
        exp.paradigm = d.get("paradigm", "")
        exp.remark = d.get("remark", "")
        exp.items = [Item.from_dict(it) for it in d.get("items", [])]
        exp.events = d.get("events", [])
        exp.started = d.get("started", False)
        exp.start_time = d.get("start_time", None)
        return exp


# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------

def load_settings() -> dict:
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            # merge with defaults (so new keys always present)
            merged = deepcopy(DEFAULT_SETTINGS)
            for k, v in data.items():
                if k == "hotkeys" and isinstance(v, dict):
                    merged["hotkeys"].update(v)
                else:
                    merged[k] = v
            return merged
        except Exception:
            pass
    return deepcopy(DEFAULT_SETTINGS)


def save_settings(settings: dict):
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def fmt_mmssxx(seconds: float) -> str:
    """Format seconds -> MM:SS.xx"""
    if seconds < 0:
        seconds = 0.0
    total_cs = int(seconds * 100)
    cs = total_cs % 100
    total_s = total_cs // 100
    s = total_s % 60
    m = total_s // 60
    return f"{m:02d}:{s:02d}.{cs:02d}"


def safe_name_part(s: str) -> str:
    """Make a string safe for use as a filename component."""
    if not s:
        return ""
    return "".join(c for c in s if c.isalnum() or c in "-_ ")[:MAX_FILENAME_PART_LENGTH].strip()


def load_recent_templates() -> List[dict]:
    if os.path.exists(RECENT_TEMPLATE_FILE):
        try:
            with open(RECENT_TEMPLATE_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return data
        except Exception:
            pass
    return []


def save_recent_template(template: dict):
    recent = load_recent_templates()
    name = template.get("name", "")
    # remove duplicate
    recent = [t for t in recent if t.get("name") != name]
    recent.insert(0, template)
    recent = recent[:MAX_RECENT_TEMPLATES]  # keep last N
    with open(RECENT_TEMPLATE_FILE, "w", encoding="utf-8") as f:
        json.dump(recent, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# Settings Dialog
# ---------------------------------------------------------------------------

class SettingsDialog(QDialog):
    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("设置")
        self.setMinimumWidth(400)
        self._settings = deepcopy(settings)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)

        hotkey_group = QGroupBox("快捷键")
        form = QFormLayout()
        hotkey_group.setLayout(form)

        self._hk_start_stop = QKeySequenceEdit(
            QKeySequence(self._settings["hotkeys"].get("start_stop", "F5"))
        )
        self._hk_reset = QKeySequenceEdit(
            QKeySequence(self._settings["hotkeys"].get("reset", "F6"))
        )
        self._hk_export = QKeySequenceEdit(
            QKeySequence(self._settings["hotkeys"].get("export", "F7"))
        )

        form.addRow("开始/暂停:", self._hk_start_stop)
        form.addRow("重置:", self._hk_reset)
        form.addRow("导出:", self._hk_export)

        layout.addWidget(hotkey_group)

        btn_row = QHBoxLayout()
        btn_ok = QPushButton("保存")
        btn_cancel = QPushButton("取消")
        btn_ok.clicked.connect(self._on_ok)
        btn_cancel.clicked.connect(self.reject)
        btn_row.addStretch()
        btn_row.addWidget(btn_ok)
        btn_row.addWidget(btn_cancel)
        layout.addLayout(btn_row)

    def _on_ok(self):
        self._settings["hotkeys"]["start_stop"] = (
            self._hk_start_stop.keySequence().toString()
        )
        self._settings["hotkeys"]["reset"] = (
            self._hk_reset.keySequence().toString()
        )
        self._settings["hotkeys"]["export"] = (
            self._hk_export.keySequence().toString()
        )
        self.accept()

    def get_settings(self) -> dict:
        return self._settings


# ---------------------------------------------------------------------------
# Item row widget (timer or counter)
# ---------------------------------------------------------------------------

class ItemRowWidget(QWidget):
    """A single row representing one timer or counter item."""

    toggled = Signal(int)   # row index
    incremented = Signal(int)
    decremented = Signal(int)
    name_changed = Signal(int, str)
    deleted = Signal(int)   # row index

    def __init__(self, index: int, item: Item, parent=None):
        super().__init__(parent)
        self._index = index
        self._item = item
        self._build_ui()

    def _build_ui(self):
        row = QHBoxLayout(self)
        row.setContentsMargins(4, 2, 4, 2)

        # Index label (1-based)
        idx_label = QLabel(f"{self._index + 1}.")
        idx_label.setFixedWidth(24)
        idx_label.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
        font_idx = QFont()
        font_idx.setBold(True)
        idx_label.setFont(font_idx)
        row.addWidget(idx_label)

        # Name (editable)
        self._name_edit = QLineEdit(self._item.name)
        self._name_edit.setMinimumWidth(160)
        self._name_edit.textChanged.connect(self._on_name_changed)
        row.addWidget(self._name_edit)

        # Value display
        self._val_label = QLabel()
        self._val_label.setMinimumWidth(90)
        self._val_label.setAlignment(Qt.AlignCenter)
        font = QFont("Courier New", 12)
        font.setBold(True)
        self._val_label.setFont(font)
        self._update_val_label()
        row.addWidget(self._val_label)

        if self._item.kind == "timer":
            btn = QPushButton("▶/⏸")
            btn.setToolTip("开始/暂停计时")
            btn.setFixedWidth(60)
            btn.clicked.connect(lambda: self.toggled.emit(self._index))
            row.addWidget(btn)
            self._toggle_btn = btn
        else:
            btn_inc = QPushButton("+1")
            btn_inc.setToolTip("计数 +1")
            btn_inc.setFixedWidth(40)
            btn_inc.clicked.connect(lambda: self.incremented.emit(self._index))

            btn_dec = QPushButton("-1")
            btn_dec.setToolTip("计数 -1")
            btn_dec.setFixedWidth(40)
            btn_dec.clicked.connect(lambda: self.decremented.emit(self._index))

            row.addWidget(btn_inc)
            row.addWidget(btn_dec)
            self._toggle_btn = None

        row.addStretch()

        # Delete button
        btn_del = QPushButton("✕")
        btn_del.setToolTip("删除此项目")
        btn_del.setFixedWidth(30)
        btn_del.setStyleSheet("color: #cc0000;")
        btn_del.clicked.connect(lambda: self.deleted.emit(self._index))
        row.addWidget(btn_del)

    def _on_name_changed(self, text: str):
        self.name_changed.emit(self._index, text)

    def _update_val_label(self):
        if self._item.kind == "timer":
            self._val_label.setText(fmt_mmssxx(self._item.current_elapsed()))
        else:
            self._val_label.setText(str(self._item.count))

    def refresh(self):
        self._update_val_label()
        if self._item.kind == "timer" and self._toggle_btn:
            if self._item.running:
                self._toggle_btn.setStyleSheet("color: red; font-weight: bold;")
            else:
                self._toggle_btn.setStyleSheet("")


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("MiceTimer - 实验计时计数软件")
        self.setMinimumSize(QSize(800, 600))

        self._settings = load_settings()
        self._exp = Experiment()
        self._item_row_widgets: List[ItemRowWidget] = []

        self._tick_timer = QTimer(self)
        self._tick_timer.setInterval(100)  # 100ms refresh
        self._tick_timer.timeout.connect(self._on_tick)

        self._build_ui()
        self._apply_hotkeys()

        # Try to recover autosave
        self._try_recover()

        self._tick_timer.start()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(8, 8, 8, 8)

        # Status bar (red when experiment running)
        self._status_bar = QStatusBar()
        self.setStatusBar(self._status_bar)
        self._status_label = QLabel("就绪")
        self._status_bar.addWidget(self._status_label)

        # Tab widget
        tabs = QTabWidget()
        main_layout.addWidget(tabs)

        # --- Tab: Experiment ---
        exp_tab = QWidget()
        exp_layout = QVBoxLayout(exp_tab)
        tabs.addTab(exp_tab, "实验")

        # Info area
        info_group = QGroupBox("实验信息")
        info_form = QFormLayout()
        info_group.setLayout(info_form)

        self._field_date = QLineEdit(self._exp.date)
        self._field_operator = QLineEdit(self._exp.operator)
        self._field_mouse_id = QLineEdit(self._exp.mouse_id)
        self._field_group = QLineEdit(self._exp.group)
        self._field_paradigm = QLineEdit(self._exp.paradigm)
        self._field_remark = QLineEdit(self._exp.remark)

        info_form.addRow("日期:", self._field_date)
        info_form.addRow("实验员:", self._field_operator)
        info_form.addRow("实验鼠ID:", self._field_mouse_id)
        info_form.addRow("组别:", self._field_group)
        info_form.addRow("范式/实验类型:", self._field_paradigm)
        info_form.addRow("备注:", self._field_remark)
        exp_layout.addWidget(info_group)

        # Items scroll area
        items_group = QGroupBox("计时/计数项目")
        items_outer = QVBoxLayout()
        items_group.setLayout(items_outer)

        self._items_container = QWidget()
        self._items_layout = QVBoxLayout(self._items_container)
        self._items_layout.setContentsMargins(0, 0, 0, 0)
        self._items_layout.setSpacing(2)

        scroll = QScrollArea()
        scroll.setWidget(self._items_container)
        scroll.setWidgetResizable(True)
        scroll.setMinimumHeight(180)
        items_outer.addWidget(scroll)

        # Item management buttons
        item_btn_row = QHBoxLayout()
        btn_add_timer = QPushButton("+ 计时项")
        btn_add_counter = QPushButton("+ 计数项")
        btn_add_timer.clicked.connect(lambda: self._add_item("timer"))
        btn_add_counter.clicked.connect(lambda: self._add_item("counter"))
        item_btn_row.addWidget(btn_add_timer)
        item_btn_row.addWidget(btn_add_counter)
        item_btn_row.addStretch()
        items_outer.addLayout(item_btn_row)

        exp_layout.addWidget(items_group)

        # Control buttons
        ctrl_row = QHBoxLayout()
        self._btn_start_stop = QPushButton("开始实验 (F5)")
        self._btn_start_stop.setFixedHeight(40)
        self._btn_start_stop.clicked.connect(self._on_start_stop)

        self._btn_reset = QPushButton("重置 (F6)")
        self._btn_reset.setFixedHeight(40)
        self._btn_reset.clicked.connect(self._on_reset)

        ctrl_row.addWidget(self._btn_start_stop)
        ctrl_row.addWidget(self._btn_reset)
        exp_layout.addLayout(ctrl_row)

        # Export buttons
        export_row = QHBoxLayout()
        self._btn_export_default = QPushButton("导出到默认目录 (F7)")
        self._btn_export_default.clicked.connect(self.export_excel_default)

        self._btn_export_as = QPushButton("另存为...")
        self._btn_export_as.clicked.connect(self.export_excel_as)

        export_row.addWidget(self._btn_export_default)
        export_row.addWidget(self._btn_export_as)
        exp_layout.addLayout(export_row)

        # --- Tab: Templates ---
        tpl_tab = QWidget()
        tpl_layout = QVBoxLayout(tpl_tab)
        tabs.addTab(tpl_tab, "模板")

        tpl_btn_row = QHBoxLayout()
        btn_load_default = QPushButton("加载默认范式")
        btn_save_tpl = QPushButton("保存为模板...")
        btn_load_tpl = QPushButton("加载模板...")
        btn_load_default.clicked.connect(self._load_default_paradigm_dialog)
        btn_save_tpl.clicked.connect(self._save_template)
        btn_load_tpl.clicked.connect(self._load_template_dialog)
        tpl_btn_row.addWidget(btn_load_default)
        tpl_btn_row.addWidget(btn_save_tpl)
        tpl_btn_row.addWidget(btn_load_tpl)
        tpl_layout.addLayout(tpl_btn_row)

        recent_group = QGroupBox("最近模板（快速加载）")
        recent_layout = QVBoxLayout()
        recent_group.setLayout(recent_layout)
        self._recent_list_widget = QTableWidget()
        self._recent_list_widget.setColumnCount(2)
        self._recent_list_widget.setHorizontalHeaderLabels(["模板名", "项目数"])
        self._recent_list_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self._recent_list_widget.setSelectionBehavior(QAbstractItemView.SelectRows)
        self._recent_list_widget.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self._recent_list_widget.setMinimumHeight(140)
        recent_layout.addWidget(self._recent_list_widget)

        btn_load_recent = QPushButton("加载选中模板")
        btn_load_recent.clicked.connect(self._load_recent_selected)
        recent_layout.addWidget(btn_load_recent)
        tpl_layout.addWidget(recent_group)
        tpl_layout.addStretch()
        self._refresh_recent_list()

        # --- Tab: Event Log ---
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        tabs.addTab(log_tab, "事件日志")

        self._log_text = QTextEdit()
        self._log_text.setReadOnly(True)
        log_layout.addWidget(self._log_text)

        btn_clear_log = QPushButton("清空日志")
        btn_clear_log.clicked.connect(self._clear_log)
        log_layout.addWidget(btn_clear_log)

        # --- Tab: Settings ---
        settings_tab = QWidget()
        settings_layout = QVBoxLayout(settings_tab)
        tabs.addTab(settings_tab, "设置")

        btn_open_settings = QPushButton("打开快捷键设置...")
        btn_open_settings.clicked.connect(self._open_settings_dialog)
        settings_layout.addWidget(btn_open_settings)

        hk_info = QLabel(
            "当前快捷键：\n"
            + self._hotkey_summary()
        )
        hk_info.setObjectName("hk_info_label")
        settings_layout.addWidget(hk_info)
        self._hk_info_label = hk_info

        settings_layout.addStretch()

        # Load experiment fields to experiment
        self._connect_info_fields()

        # Build item rows from current exp
        self._rebuild_item_rows()

    def _hotkey_summary(self) -> str:
        hk = self._settings["hotkeys"]
        return (
            f"  开始/暂停: {hk.get('start_stop', 'F5')}\n"
            f"  重置: {hk.get('reset', 'F6')}\n"
            f"  导出: {hk.get('export', 'F7')}"
        )

    def _connect_info_fields(self):
        self._field_date.textChanged.connect(lambda v: setattr(self._exp, "date", v))
        self._field_operator.textChanged.connect(lambda v: setattr(self._exp, "operator", v))
        self._field_mouse_id.textChanged.connect(lambda v: setattr(self._exp, "mouse_id", v))
        self._field_group.textChanged.connect(lambda v: setattr(self._exp, "group", v))
        self._field_paradigm.textChanged.connect(lambda v: setattr(self._exp, "paradigm", v))
        self._field_remark.textChanged.connect(lambda v: setattr(self._exp, "remark", v))

    # ------------------------------------------------------------------
    # Item rows
    # ------------------------------------------------------------------

    def _rebuild_item_rows(self):
        # Clear existing
        while self._items_layout.count():
            child = self._items_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()
        self._item_row_widgets.clear()

        for i, item in enumerate(self._exp.items):
            w = ItemRowWidget(i, item)
            w.toggled.connect(self._toggle_timer)
            w.incremented.connect(self._increment_counter)
            w.decremented.connect(self._decrement_counter)
            w.name_changed.connect(self._on_item_name_changed)
            w.deleted.connect(self._delete_item)
            self._items_layout.addWidget(w)
            self._item_row_widgets.append(w)

        self._items_layout.addStretch()

    def _add_item(self, kind: str):
        name = "新计时项" if kind == "timer" else "新计数项"
        item = Item(kind, name)
        self._exp.items.append(item)
        self._rebuild_item_rows()
        self.save_autosave()

    def _delete_item(self, index: int):
        if index < 0 or index >= len(self._exp.items):
            return
        self._exp.items.pop(index)
        self._rebuild_item_rows()
        self.save_autosave()

    def _remove_last_item(self):
        if not self._exp.items:
            return
        self._exp.items.pop()
        self._rebuild_item_rows()
        self.save_autosave()

    # ------------------------------------------------------------------
    # Hotkeys
    # ------------------------------------------------------------------

    def _apply_hotkeys(self):
        pass  # hotkeys are handled via keyPressEvent (window-focused only)

    def keyPressEvent(self, event):
        hk = self._settings["hotkeys"]

        def _matches(key_str: str) -> bool:
            if not key_str:
                return False
            qs = QKeySequence(key_str)
            if qs.isEmpty():
                return False
            # Compare first key of sequence
            key_combo = qs[0]
            return event.keyCombination() == key_combo

        # Don't intercept if a text-editing widget has focus
        focused = QApplication.focusWidget()
        if isinstance(focused, (QLineEdit, QTextEdit, QKeySequenceEdit)):
            super().keyPressEvent(event)
            return

        if _matches(hk.get("start_stop", "F5")):
            self._on_start_stop()
            event.accept()
            return
        if _matches(hk.get("reset", "F6")):
            self._on_reset()
            event.accept()
            return
        if _matches(hk.get("export", "F7")):
            self.export_excel_default()
            event.accept()
            return

        # Delete key: delete the last item
        if event.key() == Qt.Key_Delete:
            self._remove_last_item()
            event.accept()
            return

        # Digit keys 1-9: toggle timer or increment counter for that item
        key = event.key()
        if Qt.Key_1 <= key <= Qt.Key_9:
            idx = key - Qt.Key_1  # 0-based index
            if idx < len(self._exp.items):
                item = self._exp.items[idx]
                if item.kind == "timer":
                    self._toggle_timer(idx)
                else:
                    self._increment_counter(idx)
                event.accept()
                return

        super().keyPressEvent(event)

    # ------------------------------------------------------------------
    # Experiment control
    # ------------------------------------------------------------------

    def _on_start_stop(self):
        if not self._exp.started:
            # Start experiment
            self._exp.started = True
            self._exp.start_time = datetime.now().isoformat()
            self._log_event("experiment_start")
            self._update_status_bar()
            self._btn_start_stop.setText("暂停实验 (F5)")
        else:
            # Toggle: pause all running timers
            any_running = any(it.running for it in self._exp.items if it.kind == "timer")
            if any_running:
                for it in self._exp.items:
                    if it.kind == "timer" and it.running:
                        it.elapsed += time.perf_counter() - (it.last_start_ts or time.perf_counter())
                        it.running = False
                        it.last_start_ts = None
                self._log_event("pause_all")
                self._btn_start_stop.setText("继续实验 (F5)")
            else:
                self._btn_start_stop.setText("暂停实验 (F5)")
                self._log_event("resume_all")

        self.save_autosave()

    def _on_reset(self):
        reply = QMessageBox.question(
            self,
            "确认重置",
            "确认重置所有计时/计数项目？",
            QMessageBox.Yes | QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return
        self._exp.started = False
        self._exp.start_time = None
        for it in self._exp.items:
            it.reset()
        self._btn_start_stop.setText("开始实验 (F5)")
        self._update_status_bar()
        self._log_event("reset_all")
        self._rebuild_item_rows()
        self.save_autosave()

    def _toggle_timer(self, index: int):
        if index >= len(self._exp.items):
            return
        item = self._exp.items[index]
        if item.kind != "timer":
            return

        if not self._exp.started:
            QMessageBox.information(self, "提示", '请先点击"开始实验"。')
            return

        now = time.perf_counter()
        if item.running:
            item.elapsed += now - (item.last_start_ts or now)
            item.running = False
            item.last_start_ts = None
            self._log_event("timer_stop", item_name=item.name)
        else:
            item.running = True
            item.last_start_ts = now
            self._log_event("timer_start", item_name=item.name)

        self.save_autosave()

    def _increment_counter(self, index: int):
        if index >= len(self._exp.items):
            return
        item = self._exp.items[index]
        if not self._exp.started:
            QMessageBox.information(self, "提示", '请先点击"开始实验"。')
            return
        item.count += 1
        self._log_event("counter_inc", item_name=item.name, detail=str(item.count))
        self.save_autosave()

    def _decrement_counter(self, index: int):
        if index >= len(self._exp.items):
            return
        item = self._exp.items[index]
        if item.count > 0:
            item.count -= 1
        self._log_event("counter_dec", item_name=item.name, detail=str(item.count))
        self.save_autosave()

    def _on_item_name_changed(self, index: int, name: str):
        if index < len(self._exp.items):
            self._exp.items[index].name = name

    # ------------------------------------------------------------------
    # Status bar
    # ------------------------------------------------------------------

    def _update_status_bar(self):
        if self._exp.started:
            self._status_label.setText("⚠ 实验进行中")
            self._status_bar.setStyleSheet(
                "QStatusBar { background-color: #cc0000; color: white; font-weight: bold; }"
            )
        else:
            self._status_label.setText("就绪")
            self._status_bar.setStyleSheet("")

    # ------------------------------------------------------------------
    # Tick (UI refresh)
    # ------------------------------------------------------------------

    def _on_tick(self):
        for w in self._item_row_widgets:
            w.refresh()

    # ------------------------------------------------------------------
    # Autosave / Recovery
    # ------------------------------------------------------------------

    def save_autosave(self):
        try:
            data = self._exp.to_dict()
            with open(RECOVERY_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception:
            pass

    def _try_recover(self):
        if not os.path.exists(RECOVERY_FILE):
            # Load first default paradigm
            self._apply_paradigm(DEFAULT_PARADIGMS[0])
            return
        try:
            with open(RECOVERY_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
            reply = QMessageBox.question(
                self,
                "恢复上次实验",
                "检测到上次未保存的实验数据，是否恢复？",
                QMessageBox.Yes | QMessageBox.No,
            )
            if reply == QMessageBox.Yes:
                self._exp = Experiment.from_dict(data)
                self._sync_exp_to_fields()
                self._rebuild_item_rows()
                self._update_status_bar()
                if self._exp.started:
                    self._btn_start_stop.setText("暂停实验 (F5)")
            else:
                self._apply_paradigm(DEFAULT_PARADIGMS[0])
        except Exception:
            self._apply_paradigm(DEFAULT_PARADIGMS[0])

    def _sync_exp_to_fields(self):
        self._field_date.setText(self._exp.date)
        self._field_operator.setText(self._exp.operator)
        self._field_mouse_id.setText(self._exp.mouse_id)
        self._field_group.setText(self._exp.group)
        self._field_paradigm.setText(self._exp.paradigm)
        self._field_remark.setText(self._exp.remark)

    # ------------------------------------------------------------------
    # Event log
    # ------------------------------------------------------------------

    def _log_event(self, action: str, item_name: str = "", detail: str = ""):
        ts = datetime.now().strftime("%H:%M:%S.%f")[:-3]
        entry = {
            "ts": ts,
            "action": action,
            "item": item_name,
            "detail": detail,
        }
        self._exp.events.append(entry)
        line = f"[{ts}] {action}"
        if item_name:
            line += f" | {item_name}"
        if detail:
            line += f" | {detail}"
        self._log_text.append(line)

    def _clear_log(self):
        self._exp.events.clear()
        self._log_text.clear()

    # ------------------------------------------------------------------
    # Templates
    # ------------------------------------------------------------------

    def _apply_paradigm(self, paradigm: dict):
        self._exp.items = [
            Item(it["kind"], it["name"]) for it in paradigm.get("items", [])
        ]
        self._exp.paradigm = paradigm.get("name", "")
        self._field_paradigm.setText(self._exp.paradigm)
        self._rebuild_item_rows()

    def _load_default_paradigm_dialog(self):
        names = [p["name"] for p in DEFAULT_PARADIGMS]
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getItem(
            self, "选择默认范式", "范式:", names, 0, False
        )
        if ok and name:
            for p in DEFAULT_PARADIGMS:
                if p["name"] == name:
                    self._apply_paradigm(p)
                    self.save_autosave()
                    break

    def _save_template(self):
        from PySide6.QtWidgets import QInputDialog
        name, ok = QInputDialog.getText(
            self, "保存模板", "模板名称:", text=self._exp.paradigm or "模板"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        template = {
            "name": name,
            "paradigm": self._exp.paradigm,
            "items": [{"kind": it.kind, "name": it.name} for it in self._exp.items],
        }
        # Save to file
        tpl_path = os.path.join(TEMPLATES_DIR, name + ".json")
        with open(tpl_path, "w", encoding="utf-8") as f:
            json.dump(template, f, ensure_ascii=False, indent=2)

        save_recent_template(template)
        self._refresh_recent_list()
        QMessageBox.information(self, "保存成功", f"模板 [{name}] 已保存。")

    def _load_template_dialog(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "加载模板", TEMPLATES_DIR, "JSON Files (*.json)"
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                tpl = json.load(f)
            self._apply_paradigm(tpl)
            save_recent_template(tpl)
            self._refresh_recent_list()
            self.save_autosave()
        except Exception as e:
            QMessageBox.warning(self, "加载失败", str(e))

    def _refresh_recent_list(self):
        recent = load_recent_templates()
        self._recent_list_widget.setRowCount(len(recent))
        for i, tpl in enumerate(recent):
            self._recent_list_widget.setItem(i, 0, QTableWidgetItem(tpl.get("name", "")))
            self._recent_list_widget.setItem(
                i, 1, QTableWidgetItem(str(len(tpl.get("items", []))))
            )

    def _load_recent_selected(self):
        row = self._recent_list_widget.currentRow()
        if row < 0:
            QMessageBox.information(self, "提示", "请先选择一个模板。")
            return
        recent = load_recent_templates()
        if row >= len(recent):
            return
        tpl = recent[row]
        self._apply_paradigm(tpl)
        save_recent_template(tpl)
        self._refresh_recent_list()
        self.save_autosave()

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def _open_settings_dialog(self):
        dlg = SettingsDialog(self._settings, self)
        if dlg.exec() == QDialog.Accepted:
            self._settings = dlg.get_settings()
            save_settings(self._settings)
            self._apply_hotkeys()
            self._hk_info_label.setText(
                "当前快捷键：\n" + self._hotkey_summary()
            )
            # Update button labels
            hk = self._settings["hotkeys"]
            self._btn_start_stop.setText(
                f"开始实验 ({hk.get('start_stop', 'F5')})"
            )
            self._btn_reset.setText(f"重置 ({hk.get('reset', 'F6')})")
            self._btn_export_default.setText(
                f"导出到默认目录 ({hk.get('export', 'F7')})"
            )

    # ------------------------------------------------------------------
    # Excel export
    # ------------------------------------------------------------------

    def build_export_filename(self) -> str:
        date = safe_name_part(self._exp.date) or datetime.now().strftime("%Y-%m-%d")
        mouse = safe_name_part(self._exp.mouse_id)
        group = safe_name_part(self._exp.group)
        paradigm = safe_name_part(self._exp.paradigm)

        parts = [date]
        if mouse:
            parts.append(mouse)
        if group:
            parts.append(group)
        if paradigm:
            parts.append(paradigm)

        return "_".join(parts) + ".xlsx"

    def freeze_running_timers_before_export(self):
        for item in self._exp.items:
            if item.kind == "timer" and item.running and item.last_start_ts is not None:
                item.elapsed += time.perf_counter() - item.last_start_ts
                item.running = False
                item.last_start_ts = None

    def write_excel_to_path(self, path: str):
        wb = Workbook()

        ws1 = wb.active
        ws1.title = "实验结果"
        ws1.append(["字段", "值"])
        ws1.append(["日期", self._exp.date])
        ws1.append(["实验员", self._exp.operator])
        ws1.append(["实验鼠ID", self._exp.mouse_id])
        ws1.append(["组别", self._exp.group])
        ws1.append(["范式", self._exp.paradigm])
        ws1.append(["备注", self._exp.remark])
        ws1.append([])
        ws1.append(["项目", "类型", "值"])
        for it in self._exp.items:
            value = fmt_mmssxx(it.elapsed) if it.kind == "timer" else it.count
            typ = "计时" if it.kind == "timer" else "计数"
            ws1.append([it.name, typ, value])

        ws2 = wb.create_sheet("事件日志")
        ws2.append(["时间戳", "动作", "项目", "详情"])
        for e in self._exp.events:
            ws2.append([e["ts"], e["action"], e["item"], e["detail"]])

        wb.save(path)
        QApplication.beep()
        self._log_event("export_excel", detail=path)
        self.save_autosave()

    def export_excel_default(self):
        self.freeze_running_timers_before_export()
        fname = self.build_export_filename()
        path = os.path.join(EXPORT_DIR, fname)

        # Avoid overwriting existing file
        base, ext = os.path.splitext(path)
        idx = 1
        while os.path.exists(path):
            path = f"{base}_{idx}{ext}"
            idx += 1

        try:
            self.write_excel_to_path(path)
            QMessageBox.information(
                self, "导出成功", f"已导出到默认目录：\n{path}"
            )
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))

    def export_excel_as(self):
        self.freeze_running_timers_before_export()
        fname = self.build_export_filename()
        default_path = os.path.join(EXPORT_DIR, fname)

        path, _ = QFileDialog.getSaveFileName(
            self, "另存为", default_path, "Excel Files (*.xlsx)"
        )
        if not path:
            return
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        try:
            self.write_excel_to_path(path)
            QMessageBox.information(self, "导出成功", f"已导出：\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))

    # ------------------------------------------------------------------
    # Close event
    # ------------------------------------------------------------------

    def closeEvent(self, event):
        self.save_autosave()
        self._tick_timer.stop()
        event.accept()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    app = QApplication(sys.argv)
    app.setApplicationName("MiceTimer")
    app.setApplicationVersion("1.0.0")

    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()
