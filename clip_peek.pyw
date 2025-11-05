# clip_peek.py
# Frameless, draggable, resizable, always-on-top clipboard viewer (text & images)
# Preferences dialog + header auto-clear control; JSON defaults are built-in.
# Requires: PySide6  (pip install pyside6)

import sys, os, json, time, traceback
from datetime import datetime
from PySide6.QtCore import Qt, QSize, QTimer, QPoint, QRect, QCoreApplication, QByteArray
from PySide6.QtGui import (
    QPixmap, QAction, QKeySequence, QTextOption, QIcon, QPainter, QColor,
    QCursor, QPen, QGuiApplication
)
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QTextEdit, QPushButton, QSizePolicy, QSystemTrayIcon, QMenu,
    QDialog, QFormLayout, QSlider, QSpinBox, QColorDialog, QDialogButtonBox, QCheckBox
)
from PySide6.QtNetwork import QLocalServer, QLocalSocket

APP_NAME = "ClipPeek"
SINGLETON_KEY = "com.stefan.clippeek.singleton"

# -------------------------
# Paths & storage helpers
# -------------------------
def _script_dir() -> str:
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))

def _user_config_dir(app_name: str) -> str:
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA") or os.path.expanduser("~\\AppData\\Roaming")
        return os.path.join(base, app_name)
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
        return os.path.join(base, app_name)
    else:
        base = os.environ.get("XDG_CONFIG_HOME") or os.path.expanduser("~/.config")
        return os.path.join(base, app_name)

def _user_data_dir(app_name: str) -> str:
    if sys.platform.startswith("win"):
        base = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~\\AppData\\Local")
        return os.path.join(base, app_name)
    elif sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Logs")
        return os.path.join(base, app_name)
    else:
        base = os.environ.get("XDG_DATA_HOME") or os.path.expanduser("~/.local/share")
        return os.path.join(base, app_name)

def _ensure_dir(path: str) -> None:
    os.makedirs(path, exist_ok=True)

def _atomic_write_json(path: str, data: dict) -> None:
    dir_ = os.path.dirname(path)
    _ensure_dir(dir_)
    tmp_path = os.path.join(dir_, f".{os.path.basename(path)}.tmp")
    with open(tmp_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
        f.flush()
        os.fsync(f.fileno())
    os.replace(tmp_path, path)

def _backup_bad_file(path: str) -> None:
    try:
        ts = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        bad = f"{path}.bad-{ts}.json"
        os.replace(path, bad)
    except Exception:
        pass  # best effort only

def _portable_mode_paths() -> tuple[bool, str]:
    sdir = _script_dir()
    prefs_here = os.path.join(sdir, "clip_peek_prefs.json")
    portable_flag = os.path.join(sdir, "portable.flag")
    if os.path.exists(portable_flag) or os.path.exists(prefs_here):
        return True, prefs_here
    return False, prefs_here

def _resolved_prefs_path() -> str:
    is_portable, portable_path = _portable_mode_paths()
    if is_portable:
        return portable_path
    cfgdir = _user_config_dir(APP_NAME)
    _ensure_dir(cfgdir)
    return os.path.join(cfgdir, "clip_peek_prefs.json")

PREFS_PATH = _resolved_prefs_path()
LOG_DIR = os.path.join(_user_data_dir(APP_NAME), "logs")
_ensure_dir(LOG_DIR)

def _log_path() -> str:
    day = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(LOG_DIR, f"{APP_NAME}-{day}.log")

def _log(msg: str) -> None:
    try:
        with open(_log_path(), "a", encoding="utf-8") as f:
            f.write(f"[{datetime.now().isoformat(sep=' ', timespec='seconds')}] {msg}\n")
    except Exception:
        pass

# =========================
# THEME DEFAULTS + PREFS
# =========================
THEME_DEFAULTS = {
    "config_version": 1,
    "WINDOW_OPACITY_NORMAL": 0.56,
    "WINDOW_OPACITY_HOVER":  1.00,
    "WINDOW_OPACITY_ACTIVE": 1.00,

    "COLORS": {
        "window_bg":         (11, 27, 26),
        "text_bg":           (0, 0, 0),
        "image_bg":          (107, 107, 107),
        "border":            (33, 81, 78),
        "label_text":        (162, 162, 162),
        "button_text":       (162, 162, 162),
        "button_hover_bg":   (0, 0, 0),
        "button_bg":         (255, 255, 255),
        "button_border":     (108, 108, 108),
    },
    "ALPHAS": {
        "window_bg":       0.86,
        "text_bg":         0.00,
        "image_bg":        0.00,
        "button_bg":       0.00,
        "button_hover_bg": 0.21,
        "border":          1.00,
        "label_text":      1.00,
        "button_text":     1.00,
        "button_border":   1.00,
    },
    "RADIUS": {
        "window": 14,
        "text":   8,
        "button": 10,
    },
    "PADDING": {
        "text":   "6px",
        "button": "6px 12px",
    },
    "BORDER": {
        "text":   1,
        "button": 1,
        "window": 0,
    },
    "TRAY_ICON_RGB": (162, 162, 162),

    "AUTO_CLEAR_ENABLED": False,
    "AUTO_CLEAR_MINUTES": 1,

    # New prefs
    "ALWAYS_ON_TOP": True,
    "SHOW_NOTIFICATIONS": False,

    # Geometry persistence
    "WINDOW_GEOMETRY": {"x": None, "y": None, "w": None, "h": None},

    # Stagnant clipboard notification (new)
    "STAGNANT_NOTIFY_ENABLED": False,
    "STAGNANT_NOTIFY_MINUTES": 10,
}

def _deepcopy_theme(src): return json.loads(json.dumps(src))

def _migrate_theme(loaded: dict) -> dict:
    out = _deepcopy_theme(THEME_DEFAULTS)
    if not isinstance(loaded, dict):
        return out

    for k in (
        "WINDOW_OPACITY_NORMAL","WINDOW_OPACITY_HOVER","WINDOW_OPACITY_ACTIVE",
        "AUTO_CLEAR_ENABLED","AUTO_CLEAR_MINUTES","config_version",
        "ALWAYS_ON_TOP","SHOW_NOTIFICATIONS","WINDOW_GEOMETRY",
        "STAGNANT_NOTIFY_ENABLED","STAGNANT_NOTIFY_MINUTES",
    ):
        if k in loaded:
            out[k] = loaded[k]

    for sec in ("COLORS","ALPHAS","RADIUS","PADDING","BORDER"):
        if sec in loaded and isinstance(loaded[sec], dict):
            out[sec].update(loaded[sec])

    # future migrations go here
    return out

def load_theme():
    if os.path.exists(PREFS_PATH):
        try:
            with open(PREFS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            return _migrate_theme(data)
        except Exception:
            _backup_bad_file(PREFS_PATH)
    return _deepcopy_theme(THEME_DEFAULTS)

def save_theme(theme):
    out = _migrate_theme(theme)
    try:
        _atomic_write_json(PREFS_PATH, out)
    except Exception:
        try:
            with open(PREFS_PATH, "w", encoding="utf-8") as f:
                json.dump(out, f, indent=2)
        except Exception:
            pass

THEME = load_theme()

# -------------------------
# Stylesheet helpers
# -------------------------
def _rgba(name):
    r, g, b = THEME["COLORS"][name]
    a = THEME["ALPHAS"].get(name, 1.0)
    return f"rgba({r},{g},{b},{a})"

def build_stylesheet():
    R = THEME["RADIUS"]; B = THEME["BORDER"]; P = THEME["PADDING"]
    return f"""
        QMainWindow, QDialog {{
            background: {_rgba('window_bg')};
            border-radius: {R['window']}px;
            border: {B['window']}px solid {_rgba('border')};
            color: {_rgba('label_text')};
        }}
        QWidget#PrefsRoot {{
            background: {_rgba('window_bg')};
            color: {_rgba('label_text')};
        }}
        QLabel, QCheckBox {{ color: {_rgba('label_text')}; }}
        #statusLabel {{ color: {_rgba('label_text')}; }}

        QTextEdit {{
            background: {_rgba('text_bg')};
            color: {_rgba('label_text')};
            border: {B['text']}px solid {_rgba('border')};
            border-radius: {R['text']}px;
            padding: {P['text']};
        }}
        #imageView {{
            background: {_rgba('image_bg')};
            border: 0px solid transparent;
            border-radius: {R['text']}px;
        }}

        QPushButton {{
            background: {_rgba('button_bg')};
            color: {_rgba('button_text')};
            border: {B['button']}px solid {_rgba('button_border')};
            border-radius: {R['button']}px;
            padding: {P['button']};
        }}
        QPushButton:hover {{ background: {_rgba('button_hover_bg')}; }}

        QSpinBox {{
            background: {_rgba('text_bg')};
            color: {_rgba('label_text')};
            border: {B['text']}px solid {_rgba('border')};
            border-radius: {R['text']}px;
        }}
        QSlider::groove:horizontal {{
            height: 6px; background: {_rgba('border')}; border-radius: 3px;
        }}
        QSlider::handle:horizontal {{
            width: 14px; background: {_rgba('button_bg')};
            border: 1px solid {_rgba('button_border')};
            border-radius: 7px; margin: -5px 0;
        }}

        QMenu {{
            background: {_rgba('window_bg')};
            color: {_rgba('label_text')};
            border: 1px solid {_rgba('border')};
            border-radius: {R['text']}px;
        }}
        QMenu::item:selected {{ background: {_rgba('button_hover_bg')}; }}
    """

def make_tray_icon():
    lr, lg, lb = THEME["COLORS"]["label_text"]
    br, bg, bb = THEME["COLORS"]["border"]
    pm = QPixmap(32, 32); pm.fill(Qt.transparent)
    p = QPainter(pm); p.setRenderHint(QPainter.Antialiasing, True)
    p.setBrush(QColor(br, bg, bb, int(255 * THEME["ALPHAS"].get("border", 1.0))))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(7, 8, 18, 20, 4, 4)
    p.setBrush(QColor(lr, lg, lb, 80))
    p.drawRoundedRect(9, 11, 14, 15, 3, 3)
    pen = QPen(QColor(lr, lg, lb)); pen.setWidth(2)
    p.setPen(pen); p.setBrush(Qt.NoBrush)
    p.drawRoundedRect(7, 8, 18, 20, 4, 4)
    p.drawRoundedRect(9, 11, 14, 15, 3, 3)
    p.drawRoundedRect(12, 5, 8, 6, 2, 2)
    p.drawLine(14, 8, 20, 8)
    p.end()
    return QIcon(pm)

# -------------------------
# Preferences dialog
# -------------------------
class PreferencesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("PrefsRoot")
        self.setWindowTitle("ClipPeek Preferences")
        self.setModal(True)
        self._snapshot = _deepcopy_theme(THEME)

        root = QVBoxLayout(self)
        form = QFormLayout(); root.addLayout(form)

        self.s_op_norm  = self._mk_slider_opacity(THEME["WINDOW_OPACITY_NORMAL"])
        self.s_op_hover = self._mk_slider_opacity(THEME["WINDOW_OPACITY_HOVER"])
        self.s_op_active= self._mk_slider_opacity(THEME["WINDOW_OPACITY_ACTIVE"])
        form.addRow("Window Opacity — Normal", self.s_op_norm)
        form.addRow("Window Opacity — Hover",  self.s_op_hover)
        form.addRow("Window Opacity — Active", self.s_op_active)

        self.s_alpha_text = self._mk_slider_opacity(THEME["ALPHAS"]["text_bg"])
        self.s_alpha_img  = self._mk_slider_opacity(THEME["ALPHAS"]["image_bg"])
        self.s_alpha_btn  = self._mk_slider_opacity(THEME["ALPHAS"]["button_bg"])
        self.s_alpha_btnh = self._mk_slider_opacity(THEME["ALPHAS"]["button_hover_bg"])
        form.addRow("Text Area Alpha", self.s_alpha_text)
        form.addRow("Image Area Alpha", self.s_alpha_img)
        form.addRow("Button BG Alpha", self.s_alpha_btn)
        form.addRow("Button Hover Alpha", self.s_alpha_btnh)

        # Always-on-top & notifications
        self.chk_on_top = QCheckBox("Always on top")
        self.chk_on_top.setChecked(bool(THEME.get("ALWAYS_ON_TOP", True)))
        self.chk_on_top.toggled.connect(self._on_on_top_toggled)
        form.addRow("", self.chk_on_top)

        self.chk_notify = QCheckBox("Show notifications (auto-clear)")
        self.chk_notify.setChecked(bool(THEME.get("SHOW_NOTIFICATIONS", False)))
        self.chk_notify.toggled.connect(self._on_notify_toggled)
        form.addRow("", self.chk_notify)

        # Stagnant notification controls (new)
        row = QWidget(); row_layout = QHBoxLayout(row); row_layout.setContentsMargins(0,0,0,0); row_layout.setSpacing(8)
        self.chk_stagnant = QCheckBox("Notify about stagnant clipboard after")
        self.chk_stagnant.setChecked(bool(THEME.get("STAGNANT_NOTIFY_ENABLED", False)))
        self.chk_stagnant.toggled.connect(self._on_stagnant_changed)
        self.spn_stagnant = QSpinBox(); self.spn_stagnant.setRange(1, 240)
        self.spn_stagnant.setValue(int(THEME.get("STAGNANT_NOTIFY_MINUTES", 10)))
        self.spn_stagnant.setSuffix(" min")
        self.spn_stagnant.valueChanged.connect(self._on_stagnant_changed)
        row_layout.addWidget(self.chk_stagnant, 1)
        row_layout.addWidget(self.spn_stagnant, 0)
        form.addRow("", row)

        self.btn_window_bg = self._mk_color_button("window_bg")
        self.btn_text_bg   = self._mk_color_button("text_bg")
        self.btn_image_bg  = self._mk_color_button("image_bg")
        self.btn_label     = self._mk_color_button("label_text")
        self.btn_border    = self._mk_color_button("border")
        self.btn_btn_bg    = self._mk_color_button("button_bg")
        self.btn_btn_text  = self._mk_color_button("button_text")
        self.btn_btn_bord  = self._mk_color_button("button_border")
        form.addRow("Window BG Color", self.btn_window_bg)
        form.addRow("Text BG Color", self.btn_text_bg)
        form.addRow("Image BG Color", self.btn_image_bg)
        form.addRow("Label Text Color", self.btn_label)
        form.addRow("Border Color", self.btn_border)
        form.addRow("Button BG Color", self.btn_btn_bg)
        form.addRow("Button Text Color", self.btn_btn_text)
        form.addRow("Button Border Color", self.btn_btn_bord)

        btns = QDialogButtonBox()
        self.btn_apply  = btns.addButton("Apply", QDialogButtonBox.ActionRole)
        self.btn_save   = btns.addButton("Save", QDialogButtonBox.AcceptRole)
        self.btn_reset  = btns.addButton("Reset to Defaults", QDialogButtonBox.ResetRole)
        self.btn_cancel = btns.addButton("Cancel", QDialogButtonBox.RejectRole)
        root.addWidget(btns)

        for s in (self.s_op_norm, self.s_op_hover, self.s_op_active,
                  self.s_alpha_text, self.s_alpha_img, self.s_alpha_btn, self.s_alpha_btnh):
            s.slider.valueChanged.connect(self._sliders_live_update)

        self.btn_apply.clicked.connect(self.apply_only)
        self.btn_save.clicked.connect(self.save_and_close)
        self.btn_reset.clicked.connect(self.reset_defaults)
        self.btn_cancel.clicked.connect(self.cancel_and_revert)

    def _on_on_top_toggled(self, state: bool):
        THEME["ALWAYS_ON_TOP"] = bool(state)
        self.parent().apply_window_flags()

    def _on_notify_toggled(self, state: bool):
        THEME["SHOW_NOTIFICATIONS"] = bool(state)
        save_theme(THEME)

    def _on_stagnant_changed(self, *_):
        THEME["STAGNANT_NOTIFY_ENABLED"] = bool(self.chk_stagnant.isChecked())
        THEME["STAGNANT_NOTIFY_MINUTES"] = int(self.spn_stagnant.value())
        save_theme(THEME)
        self.parent()._reset_stagnant_timer(only_if_has_content=True)

    def _sliders_live_update(self, _value):
        self._read_sliders_into_theme()
        self.parent().apply_theme_live()

    def _mk_slider_opacity(self, initial: float):
        container = QWidget(); h = QHBoxLayout(container); h.setContentsMargins(0,0,0,0)
        s = QSlider(Qt.Horizontal); s.setRange(0, 100); s.setValue(int(round(initial*100)))
        sp = QSpinBox(); sp.setRange(0, 100); sp.setValue(int(round(initial*100)))
        s.valueChanged.connect(sp.setValue); sp.valueChanged.connect(s.setValue)
        h.addWidget(s, 1); h.addWidget(sp, 0)
        container.slider = s; container.spin = sp
        return container

    def _mk_color_button(self, key):
        btn = QPushButton(); btn.setText(self._color_text(THEME["COLORS"][key]))
        btn.clicked.connect(lambda: self._pick_color_live(btn, key))
        return btn

    def _color_text(self, rgb): return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"

    def _pick_color_live(self, button: QPushButton, key: str):
        r,g,b = THEME["COLORS"][key]
        dlg = QColorDialog(QColor(r,g,b), self)
        dlg.setOption(QColorDialog.NoButtons, True)
        dlg.currentColorChanged.connect(lambda c: self._on_color_changed_live(button, key, c))
        dlg.show()

    def _on_color_changed_live(self, button: QPushButton, key: str, color: QColor):
        if not color.isValid(): return
        THEME["COLORS"][key] = (color.red(), color.green(), color.blue())
        button.setText(self._color_text(THEME["COLORS"][key]))
        self.parent().apply_theme_live()

    def _read_sliders_into_theme(self):
        THEME["WINDOW_OPACITY_NORMAL"] = self.s_op_norm.slider.value()/100.0
        THEME["WINDOW_OPACITY_HOVER"]  = self.s_op_hover.slider.value()/100.0
        THEME["WINDOW_OPACITY_ACTIVE"] = self.s_op_active.slider.value()/100.0
        THEME["ALPHAS"]["text_bg"]         = self.s_alpha_text.slider.value()/100.0
        THEME["ALPHAS"]["image_bg"]        = self.s_alpha_img.slider.value()/100.0
        THEME["ALPHAS"]["button_bg"]       = self.s_alpha_btn.slider.value()/100.0
        THEME["ALPHAS"]["button_hover_bg"] = self.s_alpha_btnh.slider.value()/100.0

    def apply_only(self):
        self._read_sliders_into_theme(); self.parent().apply_theme_live()

    def save_and_close(self):
        self._read_sliders_into_theme(); save_theme(THEME)
        self.parent().apply_theme_live(); self.accept()

    def reset_defaults(self):
        global THEME; THEME = _deepcopy_theme(THEME_DEFAULTS); save_theme(THEME)
        self.parent().apply_theme_live()
        for btn, k in [
            (self.btn_window_bg, "window_bg"), (self.btn_text_bg, "text_bg"),
            (self.btn_image_bg, "image_bg"),   (self.btn_label, "label_text"),
            (self.btn_border, "border"),       (self.btn_btn_bg, "button_bg"),
            (self.btn_btn_text, "button_text"),(self.btn_btn_bord, "button_border"),
        ]: btn.setText(self._color_text(THEME["COLORS"][k]))
        for w, val in [
            (self.s_op_norm, THEME["WINDOW_OPACITY_NORMAL"]),
            (self.s_op_hover, THEME["WINDOW_OPACITY_HOVER"]),
            (self.s_op_active, THEME["WINDOW_OPACITY_ACTIVE"]),
            (self.s_alpha_text, THEME["ALPHAS"]["text_bg"]),
            (self.s_alpha_img, THEME["ALPHAS"]["image_bg"]),
            (self.s_alpha_btn, THEME["ALPHAS"]["button_bg"]),
            (self.s_alpha_btnh, THEME["ALPHAS"]["button_hover_bg"]),
        ]: w.slider.setValue(int(round(val*100)))

    def cancel_and_revert(self):
        global THEME; THEME = _deepcopy_theme(self._snapshot); save_theme(THEME)
        self.parent().apply_theme_live(); self.reject()

# -------------------------
# Main window
# -------------------------
class ClipPeek(QMainWindow):
    RESIZE_MARGIN = 8
    EDGE_NONE, EDGE_LEFT, EDGE_TOP, EDGE_RIGHT, EDGE_BOTTOM = 0, 1, 2, 4, 8
    EDGE_TOPLEFT = EDGE_TOP | EDGE_LEFT
    EDGE_TOPRIGHT = EDGE_TOP | EDGE_RIGHT
    EDGE_BOTTOMLEFT = EDGE_BOTTOM | EDGE_LEFT
    EDGE_BOTTOMRIGHT = EDGE_BOTTOM | EDGE_RIGHT

    def __init__(self, tray_icon: QIcon, singleton_server: QLocalServer | None):
        super().__init__()
        self.setWindowTitle("ClipPeek")
        self._singleton_server = singleton_server
        self._orig_image = None  # QImage for quality-preserving rescale

        # window flags based on pref
        self._apply_initial_window_flags()
        self.setAttribute(Qt.WA_AlwaysShowToolTips, True)
        self._set_opacity_normal()
        self.resize(520, 280)
        self.setMinimumSize(320, 70)
        self._restore_geometry_safe()

        # Drag/Resize state
        self._dragging = False
        self._drag_offset = QPoint()
        self._resizing = False
        self._resize_edge = self.EDGE_NONE
        self._press_geom = QRect()

        # Auto-clear trackers
        self._autoclear_timer = QTimer(self); self._autoclear_timer.setSingleShot(True)
        self._autoclear_timer.timeout.connect(self._auto_clear_if_stale)
        self._change_seq = 0
        self._armed_seq = None

        # Stagnant notify trackers (new)
        self._stagnant_timer = QTimer(self); self._stagnant_timer.setSingleShot(True)
        self._stagnant_timer.timeout.connect(self._notify_if_stagnant)
        self._stagnant_armed_seq = None

        # ---- Central UI ----
        central = QWidget(self); self.setCentralWidget(central)
        vbox = QVBoxLayout(central); vbox.setContentsMargins(10, 10, 10, 10); vbox.setSpacing(8)

        header = QHBoxLayout(); header.setSpacing(8)
        self.status_label = QLabel("Clipboard: (watching)", self); self.status_label.setObjectName("statusLabel")
        header.addWidget(self.status_label, 1)

        self.auto_chk = QCheckBox("Auto-clear after", self)
        self.auto_chk.setChecked(bool(THEME.get("AUTO_CLEAR_ENABLED", False)))
        self.auto_chk.toggled.connect(self._autoclear_settings_changed)
        header.addWidget(self.auto_chk, 0)

        self.auto_minutes = QSpinBox(self); self.auto_minutes.setRange(1, 240)
        self.auto_minutes.setValue(int(THEME.get("AUTO_CLEAR_MINUTES", 1)))
        self.auto_minutes.setSuffix(" min")
        self.auto_minutes.valueChanged.connect(self._autoclear_settings_changed)
        header.addWidget(self.auto_minutes, 0)

        self.clear_btn = QPushButton("Clear", self)
        self.clear_btn.setToolTip("Empty the clipboard (Ctrl+Delete)")
        self.clear_btn.clicked.connect(self.clear_clipboard)
        header.addWidget(self.clear_btn, 0)
        vbox.addLayout(header)

        self.text_view = QTextEdit(self); self.text_view.setReadOnly(True)
        self.text_view.setPlaceholderText("Clipboard is empty")
        self.text_view.setWordWrapMode(QTextOption.WrapAtWordBoundaryOrAnywhere)
        self.text_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        vbox.addWidget(self.text_view, 1)

        self.image_view = QLabel(self); self.image_view.setObjectName("imageView")
        self.image_view.setAlignment(Qt.AlignCenter)
        self.image_view.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.image_view.setVisible(False)
        vbox.addWidget(self.image_view, 1)

        # Shortcuts
        hide_action = QAction(self); hide_action.setShortcut(QKeySequence(Qt.Key_Escape))
        hide_action.triggered.connect(self.hide_to_tray); self.addAction(hide_action)
        quit_action = QAction(self); quit_action.setShortcut(QKeySequence("Ctrl+Q"))
        quit_action.triggered.connect(self.safe_quit); self.addAction(quit_action)
        clear_action = QAction(self); clear_action.setShortcut(QKeySequence("Ctrl+Delete"))
        clear_action.triggered.connect(self.clear_clipboard); self.addAction(clear_action)

        # Clipboard & timers
        self.clipboard = QApplication.clipboard(); self.clipboard.dataChanged.connect(self.update_view)
        self._debounce_timer = QTimer(self); self._debounce_timer.setSingleShot(True)
        self._debounce_timer.timeout.connect(self._do_update)
        self._flash_timer = QTimer(self); self._flash_timer.setSingleShot(True)
        self._flash_timer.timeout.connect(self._end_flash)

        # Tray
        self.tray = QSystemTrayIcon(tray_icon, self)
        self.tray.setToolTip("ClipPeek — clipboard viewer")
        self.tray.activated.connect(self._on_tray_activated)
        menu = QMenu()
        menu.setMinimumWidth(120)  # set width of the tray menu
        act_show = menu.addAction("Show"); act_show.triggered.connect(self.show_from_tray)
        act_hide = menu.addAction("Hide"); act_hide.triggered.connect(self.hide_to_tray)
        menu.addSeparator()
        act_clear = menu.addAction("Clear Clipboard"); act_clear.triggered.connect(self.clear_clipboard)
        menu.addSeparator()
        # Always-on-top toggle in tray too
        self.act_on_top = menu.addAction("Always on top")
        self.act_on_top.setCheckable(True)
        self.act_on_top.setChecked(bool(THEME.get("ALWAYS_ON_TOP", True)))
        self.act_on_top.toggled.connect(self._on_top_toggled_from_tray)
        menu.addSeparator()
        act_prefs = menu.addAction("Preferences…"); act_prefs.triggered.connect(self.open_preferences)
        menu.addSeparator()
        act_quit = menu.addAction("Quit"); act_quit.triggered.connect(self.safe_quit)
        self.tray.setContextMenu(menu); self.tray.show()

        # Listen for singleton pings to raise window
        if self._singleton_server:
            self._singleton_server.newConnection.connect(self._on_new_singleton_connection)

        self.apply_theme_live(); self.update_view()

    # ----- Window flags (on-top) -----
    def _apply_initial_window_flags(self):
        flags = Qt.FramelessWindowHint | Qt.Tool
        if THEME.get("ALWAYS_ON_TOP", True):
            flags |= Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)

    def apply_window_flags(self):
        flags = Qt.FramelessWindowHint | Qt.Tool
        if THEME.get("ALWAYS_ON_TOP", True):
            flags |= Qt.WindowStaysOnTopHint
        self.setWindowFlags(flags)
        self.show()  # re-show to apply new flags properly
        self.raise_()
        self.activateWindow()
        save_theme(THEME)

    def _on_top_toggled_from_tray(self, checked: bool):
        THEME["ALWAYS_ON_TOP"] = bool(checked)
        self.apply_window_flags()

    # Preferences
    def open_preferences(self):
        dlg = PreferencesDialog(self); dlg.exec()

    def apply_theme_live(self):
        QApplication.instance().setStyleSheet(build_stylesheet())
        if self._dragging or self._resizing:
            self._set_opacity_active()
        else:
            try:
                inside = self.rect().contains(self.mapFromGlobal(QCursor.pos()))
            except Exception:
                inside = False
            self._set_opacity_hover() if inside else self._set_opacity_normal()

    # Opacity helpers
    def _set_opacity_normal(self): self.setWindowOpacity(THEME["WINDOW_OPACITY_NORMAL"])
    def _set_opacity_hover(self):  self.setWindowOpacity(THEME["WINDOW_OPACITY_HOVER"])
    def _set_opacity_active(self): self.setWindowOpacity(THEME["WINDOW_OPACITY_ACTIVE"])

    # Tray handlers
    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.Trigger:
            if self.isVisible(): self.hide_to_tray()
            else: self.show_from_tray()

    def show_from_tray(self):
        self.show(); self.raise_(); self.activateWindow(); self._set_opacity_normal()

    def hide_to_tray(self): self.hide()

    def safe_quit(self):
        THEME["AUTO_CLEAR_ENABLED"] = bool(self.auto_chk.isChecked())
        THEME["AUTO_CLEAR_MINUTES"] = int(self.auto_minutes.value())
        # persist geometry
        g = self.geometry()
        THEME["WINDOW_GEOMETRY"] = {"x": g.x(), "y": g.y(), "w": g.width(), "h": g.height()}
        save_theme(THEME)
        self.tray.hide()
        QApplication.instance().quit()

    # Auto-clear logic
    def _autoclear_settings_changed(self, *_):
        THEME["AUTO_CLEAR_ENABLED"] = bool(self.auto_chk.isChecked())
        THEME["AUTO_CLEAR_MINUTES"] = int(self.auto_minutes.value())
        save_theme(THEME); self._reset_autoclear_timer(only_if_has_content=True)

    def _reset_autoclear_timer(self, only_if_has_content=False):
        self._autoclear_timer.stop()
        if not self.auto_chk.isChecked(): return
        mins = int(self.auto_minutes.value())
        if mins <= 0: return
        if only_if_has_content:
            if not (self.image_view.isVisible() or (self.text_view.isVisible() and self.text_view.toPlainText().strip())):
                return
        self._armed_seq = self._change_seq
        self._autoclear_timer.start(mins * 60_000)

    def _auto_clear_if_stale(self):
        if self._armed_seq is not None and self._armed_seq == self._change_seq:
            self.clear_clipboard()
            if THEME.get("SHOW_NOTIFICATIONS", False):
                self.tray.showMessage("ClipPeek", "Clipboard auto-cleared.", QSystemTrayIcon.Information, 2000)

    # Stagnant notification logic (new)
    def _reset_stagnant_timer(self, only_if_has_content=False):
        self._stagnant_timer.stop()
        if not THEME.get("STAGNANT_NOTIFY_ENABLED", False): return
        mins = int(THEME.get("STAGNANT_NOTIFY_MINUTES", 10))
        if mins <= 0: return
        if only_if_has_content:
            if not (self.image_view.isVisible() or (self.text_view.isVisible() and self.text_view.toPlainText().strip())):
                return
        self._stagnant_armed_seq = self._change_seq
        self._stagnant_timer.start(mins * 60_000)

    def _notify_if_stagnant(self):
        if self._stagnant_armed_seq is not None and self._stagnant_armed_seq == self._change_seq:
            mins = int(THEME.get("STAGNANT_NOTIFY_MINUTES", 10))
            self.tray.showMessage("ClipPeek", f"Clipboard unchanged for {mins} minute(s).", QSystemTrayIcon.Information, 3000)
            # re-arm for another cycle if content is still present
            self._reset_stagnant_timer(only_if_has_content=True)

    # Clipboard handling
    def update_view(self): self._debounce_timer.start(60)

    def _do_update(self):
        mime = self.clipboard.mimeData()
        if not mime:
            self.show_empty("(no data)")
            self._reset_autoclear_timer(only_if_has_content=True)
            self._stagnant_timer.stop()
            return
        img = self.clipboard.image()
        if not img.isNull():
            self._change_seq += 1
            self._orig_image = img  # store original for quality rescale
            self._show_image_from_original()
            self._reset_autoclear_timer(only_if_has_content=True)
            self._reset_stagnant_timer(only_if_has_content=True)
            return
        text = self.clipboard.text()
        if text.strip():
            self._change_seq += 1; self._orig_image = None
            self.show_text(text)
            self._reset_autoclear_timer(only_if_has_content=True)
            self._reset_stagnant_timer(only_if_has_content=True)
            return
        self._orig_image = None
        self.show_empty("(unsupported type)")
        self._reset_autoclear_timer(only_if_has_content=True)
        self._stagnant_timer.stop()

    def _show_image_from_original(self):
        # Rescale from original QImage each time to avoid repeated scaling artifacts
        self.text_view.clear(); self.text_view.setVisible(False)
        self.image_view.setVisible(True)
        target = self.image_view.size() if self.image_view.size().width() > 0 else QSize(400, 200)
        if self._orig_image and not self._orig_image.isNull():
            scaled = self._orig_image.scaled(target, Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_view.setPixmap(QPixmap.fromImage(scaled))
            self.status_label.setText("Clipboard: image")

    def show_text(self, text: str):
        self.image_view.clear(); self.image_view.setVisible(False)
        self.text_view.setVisible(True); self.text_view.setPlainText(text)
        self.status_label.setText(f"Clipboard: text ({len(text)} chars)")

    def show_empty(self, reason: str):
        self.text_view.setVisible(True); self.image_view.setVisible(False)
        self.text_view.setPlainText(""); self.text_view.setPlaceholderText(f"Clipboard is empty {reason}")
        self.status_label.setText("Clipboard: empty")

    def clear_clipboard(self):
        ok = False
        for _ in range(5):
            try:
                self.clipboard.clear(); self.clipboard.setText("")
                ok = True; break
            except Exception:
                time.sleep(0.05)
        if ok:
            self._orig_image = None
            self.show_empty("(cleared)"); self._flash_timer.start(900)
            self._autoclear_timer.stop()
            self._stagnant_timer.stop()
        else:
            self.status_label.setText("Clipboard: failed to clear")

    def _end_flash(self): self.show_empty("")

    # --- Geometry persistence / safety ---
    def _restore_geometry_safe(self):
        g = THEME.get("WINDOW_GEOMETRY", {}) or {}
        x, y, w, h = g.get("x"), g.get("y"), g.get("w"), g.get("h")
        if all(isinstance(v, int) for v in (x, y, w, h)) and w > 0 and h > 0:
            rect = QRect(x, y, w, h)
            ar = QGuiApplication.primaryScreen().availableGeometry()
            if not ar.contains(rect.topLeft()) or rect.width() > ar.width() or rect.height() > ar.height():
                # Clamp into visible area
                x = max(ar.x(), min(rect.x(), ar.right() - 100))
                y = max(ar.y(), min(rect.y(), ar.bottom() - 60))
                w = min(rect.width(), ar.width())
                h = min(rect.height(), ar.height())
            self.setGeometry(QRect(x, y, w, h))

    # === Resize/hit-test helpers
    def _hit_test_edges(self, pos: QPoint) -> int:
        mx = pos.x(); my = pos.y()
        w = self.width(); h = self.height(); m = self.RESIZE_MARGIN
        left   = mx <= m
        right  = mx >= w - m
        top    = my <= m
        bottom = my >= h - m
        edge = self.EDGE_NONE
        if left:  edge |= self.EDGE_LEFT
        if right: edge |= self.EDGE_RIGHT
        if top:   edge |= self.EDGE_TOP
        if bottom:edge |= self.EDGE_BOTTOM
        return edge

    def _update_cursor(self, pos: QPoint):
        edge = self._hit_test_edges(pos)
        if edge in (self.EDGE_TOPLEFT, self.EDGE_BOTTOMRIGHT):
            self.setCursor(Qt.SizeFDiagCursor)
        elif edge in (self.EDGE_TOPRIGHT, self.EDGE_BOTTOMLEFT):
            self.setCursor(Qt.SizeBDiagCursor)
        elif edge in (self.EDGE_LEFT, self.EDGE_RIGHT):
            self.setCursor(Qt.SizeHorCursor)
        elif edge in (self.EDGE_TOP, self.EDGE_BOTTOM):
            self.setCursor(Qt.SizeVerCursor)
        else:
            self.setCursor(Qt.ArrowCursor)

    # === Mouse events: drag vs resize
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            edge = self._hit_test_edges(event.position().toPoint())
            if edge != self.EDGE_NONE:
                self._resizing = True
                self._resize_edge = edge
                self._press_geom = self.geometry()
                self._set_opacity_active()
                event.accept(); return
            self._dragging = True
            self._drag_offset = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
            self._set_opacity_active()
            event.accept(); return
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        gp = event.globalPosition().toPoint()
        lp = event.position().toPoint()

        if self._resizing and (event.buttons() & Qt.LeftButton):
            dx = gp.x() - self._press_geom.topLeft().x()
            dy = gp.y() - self._press_geom.topLeft().y()
            g = QRect(self._press_geom)

            if self._resize_edge & self.EDGE_LEFT:
                new_left = g.left() + dx
                max_left = g.right() - self.minimumWidth()
                g.setLeft(min(new_left, max_left))
            if self._resize_edge & self.EDGE_RIGHT:
                g.setRight(max(g.left() + self.minimumWidth(), gp.x()))
            if self._resize_edge & self.EDGE_TOP:
                new_top = g.top() + dy
                max_top = g.bottom() - self.minimumHeight()
                g.setTop(min(new_top, max_top))
            if self._resize_edge & self.EDGE_BOTTOM:
                g.setBottom(max(g.top() + self.minimumHeight(), gp.y()))

            self.setGeometry(g)
            event.accept(); return

        if self._dragging and (event.buttons() & Qt.LeftButton):
            self.move(gp - self._drag_offset)
            event.accept(); return

        self._update_cursor(lp)
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self._resizing:
                self._resizing = False
                self._resize_edge = self.EDGE_NONE
            if self._dragging:
                self._dragging = False
            if self.rect().contains(self.mapFromGlobal(event.globalPosition().toPoint())):
                self._set_opacity_hover()
            else:
                self._set_opacity_normal()
            # persist geometry on interaction end
            g = self.geometry()
            THEME["WINDOW_GEOMETRY"] = {"x": g.x(), "y": g.y(), "w": g.width(), "h": g.height()}
            save_theme(THEME)
            event.accept(); return
        super().mouseReleaseEvent(event)

    def enterEvent(self, event):
        self._set_opacity_hover()
        super().enterEvent(event)

    def leaveEvent(self, event):
        if not (self._dragging or self._resizing):
            self._set_opacity_normal()
        super().leaveEvent(event)

    def resizeEvent(self, event):
        # Rescale image from original QImage (if present) to avoid artifacts
        if self.image_view.isVisible() and self._orig_image and not self._orig_image.isNull():
            self._show_image_from_original()
        super().resizeEvent(event)

    def closeEvent(self, event):
        self.hide_to_tray()
        event.ignore()

    # --- Singleton interprocess ---
    def _on_new_singleton_connection(self):
        sock = self._singleton_server.nextPendingConnection()
        if not sock: return
        sock.readyRead.connect(lambda s=sock: self._handle_singleton_message(s))

    def _handle_singleton_message(self, sock: QLocalSocket):
        try:
            data = bytes(sock.readAll()).decode("utf-8", errors="ignore").strip()
            if data == "SHOW":
                self.show_from_tray()
        finally:
            sock.disconnectFromServer()

# -------------------------
# Exception hook
# -------------------------
def _exception_hook(exc_type, exc, tb):
    text = "".join(traceback.format_exception(exc_type, exc, tb))
    _log("UNHANDLED EXCEPTION:\n" + text)
    app = QApplication.instance()
    if app:
        for w in app.topLevelWidgets():
            if isinstance(w, ClipPeek):
                try:
                    w.tray.showMessage(APP_NAME, "An error was logged.", QSystemTrayIcon.Warning, 3000)
                except Exception:
                    pass
                break
    sys.__excepthook__(exc_type, exc, tb)

sys.excepthook = _exception_hook

# -------------------------
# Singleton helper (run or ping existing)
# -------------------------
def _create_or_ping_singleton() -> QLocalServer | None:
    sock = QLocalSocket()
    sock.connectToServer(SINGLETON_KEY)
    if sock.waitForConnected(100):
        try:
            sock.write(QByteArray(b"SHOW"))
            sock.flush()
            sock.waitForBytesWritten(100)
        finally:
            sock.disconnectFromServer()
        return None  # another instance is running
    try:
        QLocalServer.removeServer(SINGLETON_KEY)
    except Exception:
        pass
    server = QLocalServer()
    if not server.listen(SINGLETON_KEY):
        return None
    return server

# -------------------------
# App entry
# -------------------------
if __name__ == "__main__":
    # ---- Safe Qt version detection (works across PySide6 builds) ----
    try:
        from PySide6 import __version_info__ as _pyside_ver
        qt_major = _pyside_ver[0]
    except Exception:
        qt_major = 6  # assume Qt6 if uncertain

    # ---- High DPI behavior ----
    if qt_major >= 6:
        # Qt6 enables HiDPI by default; optional rounding policy helps mixed-DPI setups.
        try:
            QGuiApplication.setHighDpiScaleFactorRoundingPolicy(
                Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
            )
        except Exception:
            pass
    else:
        # Qt5 legacy support
        QCoreApplication.setAttribute(Qt.AA_EnableHighDpiScaling, True)
        QCoreApplication.setAttribute(Qt.AA_UseHighDpiPixmaps, True)

    # ---- Singleton setup ----
    singleton_server = _create_or_ping_singleton()
    if singleton_server is None:
        sys.exit(0)

    # ---- Main app ----
    app = QApplication(sys.argv)
    app.setStyleSheet(build_stylesheet())
    icon = make_tray_icon()
    w = ClipPeek(icon, singleton_server)
    w.show()
    sys.exit(app.exec())
