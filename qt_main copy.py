import json
import os
import threading
import time

import keyboard
import pydirectinput
import cv2
from PyQt6 import uic

pydirectinput.PAUSE = 0.001

from overlay_status import OverlayConfig, OverlayStatus
from myUtils import window_capture, template_match_any, TITLE, SKILLA_CROP, save_image, get_foreground_keyboard_layout

from PyQt6.QtCore import Qt, QTimer, pyqtSignal, QEvent
from PyQt6.QtGui import QIcon, QAction
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QListWidget,
    QMessageBox, QInputDialog, QDialog, QDialogButtonBox,
    QLineEdit, QSystemTrayIcon, QMenu
)

# ----------------- Settings -----------------

CONFIG_PATH = "config.json"
DNF_WINDOW_IDENTIFIERS = ["地下城与勇士：创新世纪"]  # only run macros when focused window title contains any of these
OVERLAY_X = 3470
OVERLAY_Y = 30
TOGGLE_HOTKEY = "f1"
CONFIGS_DIR = "configs"
AUTO_SWITCH_THRESHOLD = 0.8
AUTO_SWITCH_INTERVAL_DEFAULT = 0.5
CAPTURE_CROP = SKILLA_CROP
AUTO_SWITCH_CROP = SKILLA_CROP
CLICK_DELAY_DEFAULT = 0.01
KEY_PRESS_DELAY_DEFAULT = 0.01
MACRO_COOLDOWN_EXTRA = 0.01
AUTO_REPLAY_WHILE_HELD = False
DEBUG_LOG_KEYS = True


# ----------------- Window detection (Windows) -----------------

if os.name == "nt":
    import ctypes
    user32 = ctypes.windll.user32
    import winsound
    MapVirtualKeyW = user32.MapVirtualKeyW
    GetAsyncKeyState = user32.GetAsyncKeyState
    VkKeyScanW = user32.VkKeyScanW
    MapVirtualKeyW.argtypes = [ctypes.c_uint, ctypes.c_uint]
    MapVirtualKeyW.restype = ctypes.c_uint
    GetAsyncKeyState.argtypes = [ctypes.c_int]
    GetAsyncKeyState.restype = ctypes.c_short
    VkKeyScanW.argtypes = [ctypes.c_wchar]
    VkKeyScanW.restype = ctypes.c_short

    def get_active_window_title() -> str:
        hwnd = user32.GetForegroundWindow()
        if not hwnd:
            return ""
        buf = ctypes.create_unicode_buffer(512)
        user32.GetWindowTextW(hwnd, buf, 512)
        return buf.value or ""

    def is_target_window_focused() -> bool:
        title = get_active_window_title()
        return any(x in title for x in DNF_WINDOW_IDENTIFIERS)

    def is_key_pressed_fast(key_name: str) -> bool:
        if not key_name:
            return False
        try:
            scan_codes = keyboard.key_to_scan_codes(key_name)
        except Exception:
            return False
        for sc in scan_codes:
            for map_type in (1, 3):
                vk = MapVirtualKeyW(sc, map_type)
                if vk and (GetAsyncKeyState(vk) & 0x8000):
                    return True
        if len(key_name) == 1:
            vk = VkKeyScanW(key_name) & 0xFF
            if vk and (GetAsyncKeyState(vk) & 0x8000):
                return True
        return False

else:
    def get_active_window_title() -> str:
        return ""

    def is_target_window_focused() -> bool:
        return True

    def is_key_pressed_fast(key_name: str) -> bool:
        return keyboard.is_pressed(key_name)


# ----------------- Macro Engine -----------------

class MacroEngine:
    """
    Hooks keys for the current profile, runs steps on key-down.

        Step format:
            - {"type": "press",      "key": "..."}
            - {"type": "down",       "key": "..."}
            - {"type": "up",         "key": "..."}
            - {"type": "click_left"}
            - {"type": "click_right"}
            - {"type": "mouse_down", "key": "left|right"}
            - {"type": "mouse_up",   "key": "left|right"}
            - {"type": "delay",      "time": 0.3}
    """
    def __init__(self):
        self.profiles = {}
        self.active_profile = None
        self.profile_hooks = {}  # trigger_key -> hook handle
        self.running = False
        self.lock = threading.RLock()
        self.running_triggers = set()
        self.macro_cooldowns = {}  # (profile, trigger_key) -> cooldown end time
        self.loop_stop_events = {}
        self.on_toggle_loop_start = None
        self.on_toggle_loop_stop = None
        self.injecting_key_counts = {}

    def _unhook_all_unlocked(self):
        for hook in self.profile_hooks.values():
            keyboard.unhook(hook)
        self.profile_hooks.clear()
        for stop_event in self.loop_stop_events.values():
            stop_event.set()
        self.loop_stop_events.clear()
        self.running_triggers.clear()
        self.injecting_key_counts.clear()
        self.running = False

    def _set_injecting(self, key: str, active: bool):
        if not key:
            return
        if active:
            self.injecting_key_counts[key] = self.injecting_key_counts.get(key, 0) + 1
            return
        count = self.injecting_key_counts.get(key, 0) - 1
        if count > 0:
            self.injecting_key_counts[key] = count
        else:
            self.injecting_key_counts.pop(key, None)

    def _is_injecting_key(self, key: str) -> bool:
        return bool(key) and self.injecting_key_counts.get(key, 0) > 0

    def stop(self):
        with self.lock:
            if not self.running:
                return
            self._unhook_all_unlocked()

    def start(self, profiles: dict, active_profile: str):
        with self.lock:
            if self.running:
                self._unhook_all_unlocked()

            self.profiles = profiles or {}
            self.active_profile = active_profile

            if not self.active_profile or self.active_profile not in self.profiles:
                self.running = False
                return

            profile_data = self.profiles[self.active_profile]
            for trigger_key in profile_data.keys():
                hook = keyboard.hook_key(
                    trigger_key,
                    callback=lambda e, k=trigger_key: self._on_key(e, k),
                    suppress=False
                )
                self.profile_hooks[trigger_key] = hook

            self.running = True

    def reload(self, profiles: dict, active_profile: str):
        with self.lock:
            was_running = self.running
            if self.running:
                self._unhook_all_unlocked()

            self.profiles = profiles or {}
            self.active_profile = active_profile

            if was_running:
                self.start(self.profiles, self.active_profile)

    def set_active_profile(self, profile_name: str):
        with self.lock:
            self.active_profile = profile_name
            if self.running:
                self.start(self.profiles, self.active_profile)

    def _run_steps_once(self, steps: list):
        for step in steps:
            t = step.get("type")
            # print(f"Executing step: {step}")
            if t == "delay":
                try:
                    secs = float(step.get("time", 0.0) or 0.0)
                except ValueError:
                    secs = 0.0
                if secs > 0:
                    time.sleep(secs)
                continue

            key = step.get("key")
            if not key and t not in ["click_left", "click_right"]:
                continue

            if t == "press":
                self._set_injecting(key, True)
                try:
                    pydirectinput.keyDown(key)
                    time.sleep(KEY_PRESS_DELAY_DEFAULT)
                    pydirectinput.keyUp(key)
                finally:
                    self._set_injecting(key, False)
            elif t == "down":
                self._set_injecting(key, True)
                try:
                    pydirectinput.keyDown(key)
                finally:
                    self._set_injecting(key, False)
            elif t == "up":
                self._set_injecting(key, True)
                try:
                    pydirectinput.keyUp(key)
                finally:
                    self._set_injecting(key, False)
            elif t == "click_left":
                pydirectinput.mouseDown(button='left')
                # print("click left down")
                time.sleep(CLICK_DELAY_DEFAULT)
                pydirectinput.mouseUp(button='left')
            elif t == "click_right":
                # print("click right down")
                pydirectinput.mouseDown(button='right')
                time.sleep(CLICK_DELAY_DEFAULT)
                pydirectinput.mouseUp(button='right')
            elif t == "mouse_down":
                pydirectinput.mouseDown(button=key)
            elif t == "mouse_up":
                pydirectinput.mouseUp(button=key)
            else:
                print(f"Unknown step type: {t}")

    def _run_steps_repeat(self, profile_name: str, trigger_key: str, steps: list, repeat: int, stop_event: threading.Event | None):
        try:
            while True:
                for _ in range(max(1, repeat)):
                    if stop_event and stop_event.is_set():
                        return
                    self._run_steps_once(steps)

                if not AUTO_REPLAY_WHILE_HELD:
                    break
                if not is_key_pressed_fast(trigger_key):
                    print(f"Trigger key '{trigger_key}' no longer physically pressed, stopping auto-replay.")
                    print(f"Trigger key '{trigger_key}' released, stopping auto-replay.")
                    break
        finally:
            end_time = time.monotonic()
            with self.lock:
                macro_id = (profile_name, trigger_key)
                self.macro_cooldowns[macro_id] = end_time + MACRO_COOLDOWN_EXTRA
                self.running_triggers.discard(trigger_key)

    def _run_steps_loop(self, trigger_key: str, steps: list, stop_event: threading.Event, stop_on_release: bool = False):
        try:
            while not stop_event.is_set():
                if stop_on_release and not is_key_pressed_fast(trigger_key):
                    if DEBUG_LOG_KEYS:
                        print(f"[key] release-detected name={trigger_key} time={time.monotonic():.3f}")
                    stop_event.set()
                    break
                if not is_target_window_focused():
                    time.sleep(0.01)
                    continue
                self._run_steps_once(steps)
        finally:
            with self.lock:
                self.running_triggers.discard(trigger_key)

    def _on_key(self, event, trigger_key: str):
        key_name = getattr(event, "name", None)
        if key_name and self._is_injecting_key(key_name):
            return
        if key_name and DEBUG_LOG_KEYS and event.event_type in {"down", "up"}:
            print(
                f"[key] {event.event_type} name={key_name} scan={getattr(event, 'scan_code', None)} "
                f"time={time.monotonic():.3f}"
            )

        if event.event_type == "down":
            if not is_target_window_focused():
                return

            if os.name == "nt":
                lang_id, _ = get_foreground_keyboard_layout()
                if lang_id != "0x409":
                    return

        with self.lock:
            prof = self.profiles.get(self.active_profile, {})
            raw = prof.get(trigger_key)
            if isinstance(raw, list):
                key_cfg = {"steps": raw, "repeat": 1, "loop_mode": "off"}
            elif isinstance(raw, dict):
                steps_val = raw.get("steps", [])
                if not isinstance(steps_val, list):
                    steps_val = []
                loop_mode = raw.get("loop_mode", "off")
                if loop_mode not in ["off", "toggle", "hold"]:
                    loop_mode = "off"
                try:
                    repeat = int(raw.get("repeat", 1))
                except (TypeError, ValueError):
                    repeat = 1
                key_cfg = {"steps": steps_val, "repeat": max(1, repeat), "loop_mode": loop_mode}
            else:
                key_cfg = {"steps": [], "repeat": 1, "loop_mode": "off"}

            steps = key_cfg.get("steps") or []
            loop_mode = key_cfg.get("loop_mode", "off")
            repeat = key_cfg.get("repeat", 1)

        if not steps:
            return

        if loop_mode == "toggle":
            if event.event_type != "down":
                return
            with self.lock:
                if trigger_key in self.loop_stop_events:
                    self.loop_stop_events[trigger_key].set()
                    self.loop_stop_events.pop(trigger_key, None)
                    if self.on_toggle_loop_stop:
                        try:
                            self.on_toggle_loop_stop()
                        except Exception:
                            pass
                    return

                stop_event = threading.Event()
                self.loop_stop_events[trigger_key] = stop_event
                self.running_triggers.add(trigger_key)

            if self.on_toggle_loop_start:
                try:
                    self.on_toggle_loop_start()
                except Exception:
                    pass

            threading.Thread(
                target=self._run_steps_loop,
                args=(trigger_key, steps, stop_event, False),
                daemon=True
            ).start()
            return

        if loop_mode == "hold":
            if event.event_type == "down":
                with self.lock:
                    if trigger_key in self.loop_stop_events:
                        return
                    stop_event = threading.Event()
                    self.loop_stop_events[trigger_key] = stop_event
                    self.running_triggers.add(trigger_key)

                threading.Thread(
                    target=self._run_steps_loop,
                    args=(trigger_key, steps, stop_event, True),
                    daemon=True
                ).start()
                return

            if event.event_type == "up":
                with self.lock:
                    stop_event = self.loop_stop_events.pop(trigger_key, None)
                if stop_event:
                    stop_event.set()
                return

        if event.event_type != "down":
            return

        with self.lock:
            macro_id = (self.active_profile, trigger_key)
            cooldown_until = self.macro_cooldowns.get(macro_id, 0.0)
            if time.monotonic() < cooldown_until:
                return
            if trigger_key in self.running_triggers:
                return
            self.running_triggers.add(trigger_key)

        threading.Thread(
            target=self._run_steps_repeat,
            args=(self.active_profile, trigger_key, steps, repeat, None),
            daemon=True
        ).start()


# ----------------- Step Dialog -----------------

class StepDialog(QDialog):
    """
    Create/Edit a step:
      - press/down/up -> key only
    - mouse_down/mouse_up -> key only (left/right)
    - delay -> time only
    """
    def __init__(self, parent=None, step=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Step" if step else "Add Step")
        self._result_step = None

        self.type_combo = QComboBox(self)
        self.type_combo.addItems(["press", "down", "up", "click_left", "click_right", "mouse_down", "mouse_up", "delay"])

        self.key_edit = QLineEdit(self)
        self.time_edit = QLineEdit(self)

        layout = QGridLayout()
        layout.addWidget(QLabel("Type:"), 0, 0)
        layout.addWidget(self.type_combo, 0, 1)

        layout.addWidget(QLabel("Key (press/down/up/mouse_down/mouse_up):"), 1, 0)
        layout.addWidget(self.key_edit, 1, 1)

        layout.addWidget(QLabel("Time seconds (delay):"), 2, 0)
        layout.addWidget(self.time_edit, 2, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self._on_ok)
        buttons.rejected.connect(self.reject)

        v = QVBoxLayout()
        v.addLayout(layout)
        v.addWidget(buttons)
        self.setLayout(v)

        # Prefill
        if step:
            t = step.get("type", "press")
            if t not in ["press", "down", "up", "click_left", "click_right", "mouse_down", "mouse_up", "delay"]:
                t = "press"
            self.type_combo.setCurrentText(t)
            if t == "delay":
                self.time_edit.setText(str(step.get("time", 0.0)))
                self.key_edit.setText("")
            elif t in ["click_left", "click_right"]:
                self.key_edit.setText("")
                self.time_edit.setText("0")
            else:
                self.key_edit.setText(str(step.get("key", "")))
                self.time_edit.setText("0")

    def _on_ok(self):
        t = self.type_combo.currentText().strip()

        if t == "delay":
            try:
                secs = float(self.time_edit.text().strip())
            except ValueError:
                QMessageBox.warning(self, "Invalid", "Delay time must be a number.")
                return
            self._result_step = {"type": "delay", "time": secs}
        elif t in ["click_left", "click_right"]:
            self._result_step = {"type": t}
        else:
            key = self.key_edit.text().strip()
            if not key:
                QMessageBox.warning(self, "Invalid", "Key is required for press/down/up/mouse_down/mouse_up.")
                return
            if t in ["mouse_down", "mouse_up"] and key not in ["left", "right"]:
                QMessageBox.warning(self, "Invalid", "Mouse button must be left or right.")
                return
            self._result_step = {"type": t, "key": key}

        self.accept()

    def get_step(self):
        return self._result_step


# ----------------- Main Window -----------------

class MainWindow(QMainWindow):
    autoSwitchProfile = pyqtSignal(str)

    def __init__(self):
        super().__init__()

        # Load Designer UI (must exist)
        uic.loadUi("main_window.ui", self)
        self.setWindowIcon(QIcon("icon.ico"))

        # your existing logic init
        self.engine = MacroEngine()
        self.engine.on_toggle_loop_start = self._play_toggle_on_sound
        self.engine.on_toggle_loop_stop = self._play_toggle_off_sound
        self.config = {"default_profile": None, "profiles": {}}
        self.selected_profile = None
        self.selected_key = None
        self.overlay = OverlayStatus(OverlayConfig(x=OVERLAY_X, y=OVERLAY_Y))
        self.hotkey_handle = keyboard.add_hotkey(TOGGLE_HOTKEY, self._on_toggle_hotkey)
        self.auto_switch_enabled = True
        self.auto_switch_interval = AUTO_SWITCH_INTERVAL_DEFAULT
        self._auto_switch_inflight = False
        self._quitting = False
        self.auto_switch_timer = QTimer(self)
        self.auto_switch_timer.timeout.connect(self._auto_switch_tick)

        # System tray
        self._init_tray()

        # Wire signals from UI -> your methods
        self._connect_signals()

        # Load data into UI
        self._bootstrap_storage()
        self.load_from_disk(select_default=True)
        self._update_status()
        self._update_overlay()

        self.editAutoSwitchInterval.setText(str(self.auto_switch_interval))
        self.chkAutoSwitch.blockSignals(True)
        self.chkAutoSwitch.setChecked(True)
        self.chkAutoSwitch.blockSignals(False)

        # status timer (window focus title)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick_status)
        self.timer.start(400)

        self.autoSwitchProfile.connect(self._apply_auto_profile)

        QTimer.singleShot(0, self._auto_start_if_possible)

    def _init_tray(self):
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(QIcon("icon.ico"))

        menu = QMenu(self)
        action_show = QAction("Show", self)
        action_hide = QAction("Hide", self)
        action_quit = QAction("Exit", self)

        action_show.triggered.connect(self._tray_show)
        action_hide.triggered.connect(self._tray_hide)
        action_quit.triggered.connect(self._tray_quit)

        menu.addAction(action_show)
        menu.addAction(action_hide)
        menu.addSeparator()
        menu.addAction(action_quit)

        self.tray_icon.setContextMenu(menu)
        self.tray_icon.activated.connect(self._on_tray_activated)
        self.tray_icon.show()

    def _on_tray_activated(self, reason):
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            if self.isVisible() and not self.isMinimized():
                self._tray_hide()
            else:
                self._tray_show()

    def _tray_show(self):
        self.showNormal()
        self.raise_()
        self.activateWindow()

    def _tray_hide(self):
        self.hide()

    def changeEvent(self, event):
        if event.type() == QEvent.Type.WindowStateChange:
            if self.isMinimized() and hasattr(self, "tray_icon"):
                QTimer.singleShot(0, self._tray_hide)
        super().changeEvent(event)

    def _tray_quit(self):
        self._quitting = True
        self.close()

    def _on_toggle_hotkey(self):
        QTimer.singleShot(0, self._toggle_running)

    def _toggle_running(self):
        if self.engine.running:
            self.on_stop()
            default_profile = self.config.get("default_profile")
            if default_profile and default_profile in self.config.get("profiles", {}):
                self.refresh_profiles(select=default_profile)
                self.comboProfile.setCurrentText(default_profile)
                self.on_profile_changed(default_profile)
        else:
            self.on_start()
            if self.auto_switch_enabled:
                self._auto_switch_tick()

    def _play_toggle_on_sound(self):
        if os.name != "nt":
            return
        base_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base_dir, "sounds", "on.wav")
        if not os.path.exists(path):
            print(f"Toggle on sound not found at {path}")
            return
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)

    def _play_toggle_off_sound(self):
        if os.name != "nt":
            return
        base_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base_dir, "sounds", "off.wav")
        if not os.path.exists(path):
            print(f"Toggle off sound not found at {path}")
            return
        winsound.PlaySound(path, winsound.SND_FILENAME | winsound.SND_ASYNC)

    def _auto_start_if_possible(self):
        if self.engine.running:
            return
        if not self.config.get("profiles"):
            return
        self.on_start()
        if self.auto_switch_enabled:
            self._auto_switch_tick()

    def _set_auto_switch_active(self, active: bool):
        if active and self.auto_switch_enabled:
            self.auto_switch_timer.start(int(self.auto_switch_interval * 1000))
        else:
            self.auto_switch_timer.stop()

    # ---------- UI ----------
    def _connect_signals(self):
        # Top buttons
        self.btnStart.clicked.connect(self.on_start)
        self.btnStop.clicked.connect(self.on_stop)
        self.btnReload.clicked.connect(self.on_reload)
        self.btnSave.clicked.connect(self.on_save)

        # Profile controls
        self.comboProfile.currentTextChanged.connect(self.on_profile_changed)
        self.btnAddProfile.clicked.connect(self.add_profile)
        self.btnDeleteProfile.clicked.connect(self.delete_profile)
        self.btnRenameProfile.clicked.connect(self.rename_profile)
        self.btnSetDefault.clicked.connect(self.set_default_profile)

        # Key controls
        self.comboKey.currentTextChanged.connect(self.on_key_changed)
        self.btnAddKey.clicked.connect(self.add_key)
        self.btnEditKey.clicked.connect(self.edit_key)
        self.btnDeleteKey.clicked.connect(self.delete_key)

        # Steps controls
        self.btnAddStep.clicked.connect(self.add_step)
        self.btnEditStep.clicked.connect(self.edit_step)
        self.btnDeleteStep.clicked.connect(self.delete_step)
        self.btnMoveUp.clicked.connect(lambda: self.move_step(-1))
        self.btnMoveDown.clicked.connect(lambda: self.move_step(1))

        # Loop controls
        self.comboLoopMode.currentTextChanged.connect(self.on_loop_mode_changed)
        self.spinRepeat.valueChanged.connect(self.on_repeat_changed)

        # Auto profile switch
        self.btnCaptureProfileImage.clicked.connect(self.capture_profile_image)
        self.chkAutoSwitch.toggled.connect(self.on_auto_switch_toggled)
        self.editAutoSwitchInterval.editingFinished.connect(self.on_auto_switch_interval_changed)
    # ---------- Config IO ----------

    def _bootstrap_storage(self):
        self._ensure_configs_dir()
        if not os.path.exists(CONFIG_PATH) or os.path.getsize(CONFIG_PATH) == 0:
            self._write_default_config()

    def _write_default_config(self):
        self.config = {"default_profile": None, "profiles": {}}
        try:
            with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2)
        except Exception as e:
            QMessageBox.warning(self, "Init warning", f"Could not initialize config file: {e}")

    def load_from_disk(self, select_default=False):
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            else:
                self._write_default_config()

            if "profiles" not in self.config or not isinstance(self.config["profiles"], dict):
                self.config["profiles"] = {}

            # Normalize per-trigger key configs for backward compatibility
            for prof_name, prof_data in self.config["profiles"].items():
                if not isinstance(prof_data, dict):
                    continue
                for key_name, val in list(prof_data.items()):
                    if isinstance(val, list):
                        prof_data[key_name] = {"steps": val, "repeat": 1, "loop_mode": "off"}
                    elif isinstance(val, dict):
                        steps_val = val.get("steps", [])
                        if not isinstance(steps_val, list):
                            steps_val = []
                        try:
                            repeat = int(val.get("repeat", 1))
                        except (TypeError, ValueError):
                            repeat = 1
                        loop_mode = val.get("loop_mode", "off")
                        if loop_mode not in ["off", "toggle", "hold"]:
                            loop_mode = "off"
                        prof_data[key_name] = {
                            "steps": steps_val,
                            "repeat": max(1, repeat),
                            "loop_mode": loop_mode,
                        }
                    else:
                        prof_data[key_name] = {"steps": [], "repeat": 1, "loop_mode": "off"}

        except Exception as e:
            QMessageBox.critical(self, "Load error", str(e))
            self._write_default_config()

        profiles = list(self.config["profiles"].keys())
        self.comboProfile.blockSignals(True)
        self.comboProfile.clear()
        self.comboProfile.addItems(profiles)
        self.comboProfile.blockSignals(False)

        # choose profile
        chosen = None
        if select_default:
            d = self.config.get("default_profile")
            if d in profiles:
                chosen = d
            elif profiles:
                chosen = profiles[0]
        else:
            if self.selected_profile in profiles:
                chosen = self.selected_profile
            elif self.config.get("default_profile") in profiles:
                chosen = self.config.get("default_profile")
            elif profiles:
                chosen = profiles[0]

        self.selected_profile = chosen
        if chosen:
            self.comboProfile.setCurrentText(chosen)
        else:
            self.comboKey.clear()
            self.listSteps.clear()

        self._update_default_label()
        self.refresh_keys()
        self.refresh_steps()

    def save_to_disk(self):
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(self.config, f, indent=2)

    # ---------- Status ----------

    def _update_default_label(self):
        d = self.config.get("default_profile")
        self.lblDefault.setText(f"Default: {d if d else '(none)'}")

    def _update_status(self):
        self.lblStatus.setText("Status: running" if self.engine.running else "Status: stopped")
        self._update_overlay()

    def _update_overlay(self):
        status = "ON" if self.engine.running else "OFF"
        profile = self.selected_profile or self.config.get("default_profile") or "(none)"
        self.overlay.update_text(f"Macro: {status}\nProfile: {profile}")

    def _tick_status(self):
        # show focused title + gate state
        if os.name == "nt":
            title = get_active_window_title()
            ok = is_target_window_focused()
            self.lblFocusGate.setText(
                f"Focus gate ({'ON' if ok else 'OFF'}): " + ", ".join(DNF_WINDOW_IDENTIFIERS) + f" | Active: {title}"
            )
        else:
            self.lblFocusGate.setText("Focus gate: (Windows only)")

    # ---------- Auto profile switch ----------

    def _ensure_configs_dir(self):
        os.makedirs(CONFIGS_DIR, exist_ok=True)

    def capture_profile_image(self):
        if not self.selected_profile:
            QMessageBox.warning(self, "No profile", "Select a profile first.")
            return

        img = window_capture(TITLE, crop=CAPTURE_CROP, bgr=True)
        if img is None:
            QMessageBox.warning(self, "Capture failed", "Could not capture target window.")
            return

        self._ensure_configs_dir()
        safe_name = self.selected_profile.replace(os.sep, "_")
        path = os.path.join(CONFIGS_DIR, f"{safe_name}.png")
        save_image(path, img)
        QMessageBox.information(self, "Captured", f"Saved profile image to {path}")

    def on_auto_switch_toggled(self, checked: bool):
        self.auto_switch_enabled = bool(checked)
        if self.auto_switch_enabled and self.engine.running:
            self._ensure_configs_dir()
            self.auto_switch_timer.start(int(self.auto_switch_interval * 1000))
        else:
            self.auto_switch_timer.stop()

    def on_auto_switch_interval_changed(self):
        text = self.editAutoSwitchInterval.text().strip()
        try:
            secs = float(text)
        except ValueError:
            QMessageBox.warning(self, "Invalid", "Interval must be a number (seconds).")
            self.editAutoSwitchInterval.setText(str(self.auto_switch_interval))
            return

        if secs <= 0:
            QMessageBox.warning(self, "Invalid", "Interval must be > 0.")
            self.editAutoSwitchInterval.setText(str(self.auto_switch_interval))
            return

        self.auto_switch_interval = secs
        if self.auto_switch_enabled:
            self.auto_switch_timer.start(int(self.auto_switch_interval * 1000))

    def _auto_switch_tick(self):
        if self._auto_switch_inflight:
            return
        if not self.auto_switch_enabled:
            return
        self._auto_switch_inflight = True
        threading.Thread(target=self._auto_switch_worker, daemon=True).start()

    def _auto_switch_worker(self):
        try:
            scene = window_capture(TITLE, crop=AUTO_SWITCH_CROP, bgr=True)
            if scene is None:
                return

            best_profile = None
            best_score = -1.0

            if not os.path.isdir(CONFIGS_DIR):
                return

            for name in os.listdir(CONFIGS_DIR):
                lower = name.lower()
                if not lower.endswith((".png", ".jpg", ".jpeg", ".bmp")):
                    continue

                profile_name = os.path.splitext(name)[0]
                if profile_name not in self.config.get("profiles", {}):
                    continue

                path = os.path.join(CONFIGS_DIR, name)
                found, score = template_match_any(
                    path,
                    scene,
                    threshold=AUTO_SWITCH_THRESHOLD,
                    return_score=True,
                )
                # print(f"Auto-switch: checking profile '{profile_name}' -> found={found}, score={score:.3f}")
                if found and score > best_score:
                    best_score = score
                    best_profile = profile_name
                    # print(f"Auto-switch: matched profile '{profile_name}' with score {score:.3f}")


            if best_profile and best_profile != self.selected_profile:
                self.autoSwitchProfile.emit(best_profile)
        finally:
            self._auto_switch_inflight = False

    def _apply_auto_profile(self, profile_name: str):
        if profile_name not in self.config.get("profiles", {}):
            return
        self.refresh_profiles(select=profile_name)
        self.comboProfile.setCurrentText(profile_name)
        self.on_profile_changed(profile_name)

    # ---------- Engine controls ----------

    def on_start(self):
        if not self.config["profiles"]:
            QMessageBox.warning(self, "No profiles", "No profiles found. Add one first.")
            return

        prof = self.selected_profile or self.config.get("default_profile") or next(iter(self.config["profiles"].keys()))
        if prof not in self.config["profiles"]:
            QMessageBox.warning(self, "Invalid", "Selected profile is invalid.")
            return

        self.selected_profile = prof
        self.engine.start(self.config["profiles"], prof)
        self._set_auto_switch_active(True)
        self._update_status()

    def on_stop(self):
        self.engine.stop()
        self._set_auto_switch_active(False)
        self._update_status()

    def on_reload(self):
        current_prof = self.selected_profile
        current_key = self.selected_key

        self.load_from_disk(select_default=False)

        # restore selection best-effort
        if current_prof in self.config["profiles"]:
            self.selected_profile = current_prof
            self.comboProfile.setCurrentText(current_prof)
            self.refresh_keys()
            if current_key and current_key in self.config["profiles"][current_prof]:
                self.selected_key = current_key
                self.comboKey.setCurrentText(current_key)
            self.refresh_steps()

        # apply to engine if running
        self.engine.reload(self.config["profiles"], self.selected_profile or self.config.get("default_profile"))
        self._update_status()

    def on_save(self):
        try:
            self.save_to_disk()
        except Exception as e:
            QMessageBox.critical(self, "Save error", str(e))
            return

        # apply to engine immediately if running
        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile or self.config.get("default_profile"))

        QMessageBox.information(self, "Saved", f"Saved to {CONFIG_PATH}")

    # ---------- Profile controls ----------

    def on_profile_changed(self, profile_name: str):
        profile_name = profile_name.strip()
        self.selected_profile = profile_name if profile_name else None

        # Clear key selection first
        self.selected_key = None

        # Rebuild key dropdown
        self.refresh_keys()

        # Auto-select first key if available
        if self.selected_profile and self.selected_profile in self.config["profiles"]:
            keys = list(self.config["profiles"][self.selected_profile].keys())
            if keys:
                self.comboKey.blockSignals(True)
                self.comboKey.setCurrentIndex(0)
                self.comboKey.blockSignals(False)
                self.selected_key = self.comboKey.currentText().strip() or None

        # Now refresh steps with the selected key
        self.refresh_steps()

        # Update loop controls for current key
        self._update_loop_controls()

        # If engine running, switch profile
        if self.engine.running:
            self.engine.set_active_profile(self.selected_profile)
        self._update_status()



    def add_profile(self):
        name, ok = QInputDialog.getText(self, "Add Profile", "Profile name (e.g. 1, 2, 3):")
        if not ok or not name.strip():
            return
        name = name.strip()

        if name in self.config["profiles"]:
            QMessageBox.warning(self, "Exists", f"Profile '{name}' already exists.")
            return

        self.config["profiles"][name] = {}

        if not self.config.get("default_profile"):
            self.config["default_profile"] = name

        self.refresh_profiles(select=name)

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)
        self._update_status()

    def delete_profile(self):
        if not self.selected_profile:
            QMessageBox.warning(self, "No profile", "Select a profile first.")
            return

        name = self.selected_profile
        resp = QMessageBox.question(self, "Confirm", f"Delete profile '{name}'?")
        if resp != QMessageBox.StandardButton.Yes:
            return

        self.config["profiles"].pop(name, None)

        # Remove stored image if present
        try:
            img_path = os.path.join(CONFIGS_DIR, f"{name}.png")
            if os.path.exists(img_path):
                os.remove(img_path)
        except Exception as e:
            QMessageBox.warning(self, "Image delete", f"Profile deleted, but image delete failed: {e}")

        # Fix default profile if needed
        if self.config.get("default_profile") == name:
            remaining = list(self.config["profiles"].keys())
            self.config["default_profile"] = remaining[0] if remaining else None

        # Refresh UI from memory (do NOT load_from_disk here)
        self.selected_profile = None
        self.selected_key = None
        self.refresh_profiles(select=self.config.get("default_profile"))

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile or self.config.get("default_profile"))
        self._update_status()

    def rename_profile(self):
        if not self.selected_profile:
            QMessageBox.warning(self, "No profile", "Select a profile first.")
            return

        old_name = self.selected_profile
        new_name, ok = QInputDialog.getText(self, "Rename Profile", "New profile name:")
        if not ok or not new_name.strip():
            return

        new_name = new_name.strip()
        if new_name == old_name:
            return

        if new_name in self.config["profiles"]:
            QMessageBox.warning(self, "Exists", f"Profile '{new_name}' already exists.")
            return

        # Move profile data
        self.config["profiles"][new_name] = self.config["profiles"].pop(old_name)

        # Update default profile if needed
        if self.config.get("default_profile") == old_name:
            self.config["default_profile"] = new_name

        # Rename stored image if present
        try:
            old_img = os.path.join(CONFIGS_DIR, f"{old_name}.png")
            new_img = os.path.join(CONFIGS_DIR, f"{new_name}.png")
            if os.path.exists(old_img):
                os.rename(old_img, new_img)
        except Exception as e:
            QMessageBox.warning(self, "Image rename", f"Profile renamed, but image rename failed: {e}")

        self.selected_profile = new_name
        self.refresh_profiles(select=new_name)

        try:
            self.save_to_disk()
        except Exception as e:
            QMessageBox.warning(self, "Save error", f"Profile renamed, but save failed: {e}")

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)
        self._update_status()


    def set_default_profile(self):
        if not self.selected_profile:
            QMessageBox.warning(self, "No profile", "Select a profile first.")
            return
        self.config["default_profile"] = self.selected_profile
        self._update_default_label()
        
    def refresh_profiles(self, select: str | None = None):
        profiles = list(self.config["profiles"].keys())

        self.comboProfile.blockSignals(True)
        self.comboProfile.clear()
        self.comboProfile.addItems(profiles)
        self.comboProfile.blockSignals(False)

        # Decide which profile to select
        chosen = None
        if select and select in profiles:
            chosen = select
        elif self.selected_profile in profiles:
            chosen = self.selected_profile
        elif self.config.get("default_profile") in profiles:
            chosen = self.config.get("default_profile")
        elif profiles:
            chosen = profiles[0]

        self.selected_profile = chosen
        if chosen:
            self.comboProfile.setCurrentText(chosen)
        else:
            self.comboKey.clear()
            self.listSteps.clear()
            self.selected_key = None

        self._update_default_label()
        self.refresh_keys()
        self.refresh_steps()

    # ---------- Key controls ----------

    def refresh_keys(self):
        self.comboKey.blockSignals(True)
        self.comboKey.clear()

        if not self.selected_profile or self.selected_profile not in self.config["profiles"]:
            self.comboKey.blockSignals(False)
            self.selected_key = None
            return

        keys = list(self.config["profiles"][self.selected_profile].keys())
        self.comboKey.addItems(keys)

        # If the previous selected_key is still valid, keep it; otherwise pick the first.
        if self.selected_key in keys:
            self.comboKey.setCurrentText(self.selected_key)
        elif keys:
            self.comboKey.setCurrentIndex(0)
            self.selected_key = self.comboKey.currentText().strip() or None
        else:
            self.selected_key = None

        self.comboKey.blockSignals(False)

        self._update_loop_controls()


    def on_key_changed(self, key_name: str):
        key_name = key_name.strip()
        self.selected_key = key_name if key_name else None
        self.refresh_steps()
        self._update_loop_controls()

    def add_key(self):
        if not self.selected_profile:
            QMessageBox.warning(self, "No profile", "Select a profile first.")
            return

        key_name, ok = QInputDialog.getText(self, "Add Key", "Trigger key (e.g. w, space, q):")
        if not ok or not key_name.strip():
            return
        key_name = key_name.strip()

        prof = self.config["profiles"].setdefault(self.selected_profile, {})
        if key_name in prof:
            QMessageBox.warning(self, "Exists", f"Key '{key_name}' already exists in this profile.")
            return

        prof[key_name] = {
            "steps": [],
            "repeat": self.spinRepeat.value(),
            "loop_mode": self.comboLoopMode.currentText().strip() or "off",
        }
        self.selected_key = key_name
        self.refresh_keys()
        self.comboKey.setCurrentText(key_name)
        self.refresh_steps()

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)
        self._update_status()

    def delete_key(self):
        if not self.selected_profile or not self.selected_key:
            QMessageBox.warning(self, "No selection", "Select a profile and key first.")
            return

        resp = QMessageBox.question(self, "Confirm", f"Delete key '{self.selected_key}'?")
        if resp != QMessageBox.StandardButton.Yes:
            return

        self.config["profiles"][self.selected_profile].pop(self.selected_key, None)
        self.selected_key = None
        self.refresh_keys()
        self.refresh_steps()

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)
        self._update_status()

    def edit_key(self):
        if not self.selected_profile or not self.selected_key:
            QMessageBox.warning(self, "No selection", "Select a profile and key first.")
            return

        new_key, ok = QInputDialog.getText(
            self,
            "Edit Key",
            f"New trigger key (current: '{self.selected_key}'):",
            text=self.selected_key
        )
        if not ok or not new_key.strip():
            return
        new_key = new_key.strip()

        prof = self.config["profiles"][self.selected_profile]
        current_cfg = prof.get(self.selected_key, {"steps": [], "repeat": 1, "loop_mode": "off"})
        if isinstance(current_cfg, list):
            current_cfg = {"steps": current_cfg, "repeat": 1, "loop_mode": "off"}
        elif not isinstance(current_cfg, dict):
            current_cfg = {"steps": [], "repeat": 1, "loop_mode": "off"}

        if new_key == self.selected_key:
            current_cfg["repeat"] = self.spinRepeat.value()
            current_cfg["loop_mode"] = self.comboLoopMode.currentText().strip() or "off"
            prof[self.selected_key] = current_cfg
            self.refresh_steps()
            if self.engine.running:
                self.engine.reload(self.config["profiles"], self.selected_profile)
            return

        if new_key in prof:
            QMessageBox.warning(self, "Exists", f"Key '{new_key}' already exists in this profile.")
            return

        # Rename the key by copying steps and deleting the old one
        current_cfg["repeat"] = self.spinRepeat.value()
        current_cfg["loop_mode"] = self.comboLoopMode.currentText().strip() or "off"
        prof[new_key] = current_cfg
        prof.pop(self.selected_key, None)
        self.selected_key = new_key
        self.refresh_keys()
        self.comboKey.setCurrentText(new_key)
        self.refresh_steps()

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)
        self._update_status()

    # ---------- Step controls ----------

    def current_steps(self):
        if not self.selected_profile or not self.selected_key:
            return None
        prof = self.config["profiles"].setdefault(self.selected_profile, {})
        val = prof.get(self.selected_key)
        if isinstance(val, list):
            val = {"steps": val, "repeat": 1, "loop_mode": "off"}
            prof[self.selected_key] = val
        elif not isinstance(val, dict):
            val = {"steps": [], "repeat": 1, "loop_mode": "off"}
            prof[self.selected_key] = val
        steps = val.get("steps")
        if not isinstance(steps, list):
            steps = []
            val["steps"] = steps
        return steps

    def _get_selected_key_cfg(self):
        if not self.selected_profile or not self.selected_key:
            return None
        prof = self.config["profiles"].setdefault(self.selected_profile, {})
        val = prof.get(self.selected_key)
        if isinstance(val, list):
            val = {"steps": val, "repeat": 1, "loop_mode": "off"}
            prof[self.selected_key] = val
        elif not isinstance(val, dict):
            val = {"steps": [], "repeat": 1, "loop_mode": "off"}
            prof[self.selected_key] = val
        if "steps" not in val or not isinstance(val["steps"], list):
            val["steps"] = []
        if "repeat" not in val:
            val["repeat"] = 1
        if val.get("loop_mode") not in ["off", "toggle", "hold"]:
            val["loop_mode"] = "off"
        return val

    def _update_loop_controls(self):
        cfg = self._get_selected_key_cfg()
        has_key = cfg is not None

        self.comboLoopMode.blockSignals(True)
        self.spinRepeat.blockSignals(True)

        if not has_key:
            self.comboLoopMode.setCurrentText("off")
            self.spinRepeat.setValue(1)
        else:
            self.comboLoopMode.setCurrentText(cfg.get("loop_mode", "off"))
            self.spinRepeat.setValue(int(cfg.get("repeat", 1) or 1))

        self.comboLoopMode.setEnabled(has_key)
        self.spinRepeat.setEnabled(has_key)

        self.comboLoopMode.blockSignals(False)
        self.spinRepeat.blockSignals(False)

    def on_loop_mode_changed(self, mode: str):
        cfg = self._get_selected_key_cfg()
        if not cfg:
            return
        mode = mode.strip() or "off"
        if mode not in ["off", "toggle", "hold"]:
            mode = "off"
        cfg["loop_mode"] = mode
        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    def on_repeat_changed(self, value: int):
        cfg = self._get_selected_key_cfg()
        if not cfg:
            return
        cfg["repeat"] = max(1, int(value))
        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    def refresh_steps(self):
        self.listSteps.clear()
        steps = self.current_steps()
        if not steps:
            return

        for i, s in enumerate(steps):
            t = s.get("type", "")
            if t == "delay":
                val = float(s.get("time", 0.0) or 0.0)
                self.listSteps.addItem(f"{i+1}. delay {val:.3f}s")
            elif t == "click_left":
                self.listSteps.addItem(f"{i+1}. click left")
            elif t == "click_right":
                self.listSteps.addItem(f"{i+1}. click right")
            else:
                k = s.get("key", "?")
                self.listSteps.addItem(f"{i+1}. {t} '{k}'")

    def add_step(self):
        if not self.selected_profile or not self.selected_key:
            QMessageBox.warning(self, "No selection", "Select a profile and key first.")
            return

        dlg = StepDialog(self)
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        step = dlg.get_step()
        if not step:
            return

        steps = self.current_steps()
        steps.append(step)
        self.refresh_steps()

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    def edit_step(self):
        steps = self.current_steps()
        if not steps:
            QMessageBox.warning(self, "No steps", "No steps to edit.")
            return

        row = self.listSteps.currentRow()
        if row < 0 or row >= len(steps):
            QMessageBox.warning(self, "No step", "Select a step first.")
            return

        dlg = StepDialog(self, step=steps[row])
        if dlg.exec() != QDialog.DialogCode.Accepted:
            return
        step = dlg.get_step()
        if not step:
            return

        steps[row] = step
        self.refresh_steps()
        self.listSteps.setCurrentRow(row)

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    def delete_step(self):
        steps = self.current_steps()
        if not steps:
            QMessageBox.warning(self, "No steps", "No steps to delete.")
            return

        row = self.listSteps.currentRow()
        if row < 0 or row >= len(steps):
            QMessageBox.warning(self, "No step", "Select a step first.")
            return

        resp = QMessageBox.question(self, "Confirm", "Delete selected step?")
        if resp != QMessageBox.StandardButton.Yes:
            return

        del steps[row]
        self.refresh_steps()

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    def move_step(self, direction: int):
        steps = self.current_steps()
        if not steps or len(steps) < 2:
            return

        row = self.listSteps.currentRow()
        if row < 0 or row >= len(steps):
            return

        new_row = row + direction
        if new_row < 0 or new_row >= len(steps):
            return

        steps[row], steps[new_row] = steps[new_row], steps[row]
        self.refresh_steps()
        self.listSteps.setCurrentRow(new_row)

        if self.engine.running:
            self.engine.reload(self.config["profiles"], self.selected_profile)

    # ---------- Close ----------

    def closeEvent(self, event):
        if hasattr(self, "tray_icon") and self.tray_icon.isVisible() and not self._quitting:
            self.hide()
            event.ignore()
            return
        try:
            self.engine.stop()
        except Exception:
            pass
        try:
            self.auto_switch_timer.stop()
        except Exception:
            pass
        try:
            keyboard.remove_hotkey(self.hotkey_handle)
        except Exception:
            pass
        try:
            self.overlay.close()
        except Exception:
            pass
        try:
            if hasattr(self, "tray_icon"):
                self.tray_icon.hide()
        except Exception:
            pass
        event.accept()


def main():
    if os.name == "nt":
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("dnf.macro.tool")
        except Exception:
            pass

    app = QApplication([])
    app.setWindowIcon(QIcon("icon.ico"))
    win = MainWindow()
    win.setFixedSize(500, 650)
    win.showMinimized()
    app.exec()


if __name__ == "__main__":
    main()
