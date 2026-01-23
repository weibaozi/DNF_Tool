import json
import os
import threading
import time

import keyboard
import pydirectinput
import cv2
from PyQt6 import uic

from overlay_status import OverlayConfig, OverlayStatus
from myUtils import window_capture, template_match_any, TITLE, SKILLA_CROP, save_image

from PyQt6.QtCore import Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QComboBox, QListWidget,
    QMessageBox, QInputDialog, QDialog, QDialogButtonBox,
    QLineEdit
)

# ----------------- Settings -----------------

CONFIG_PATH = "config.json"
DNF_WINDOW_IDENTIFIERS = ["地下城与勇士", "DNF"]  # only run macros when focused window title contains any of these
OVERLAY_X = 30
OVERLAY_Y = 30
TOGGLE_HOTKEY = "f1"
CONFIGS_DIR = "configs"
AUTO_SWITCH_THRESHOLD = 0.7
AUTO_SWITCH_INTERVAL_DEFAULT = 2.0
CAPTURE_CROP = SKILLA_CROP
AUTO_SWITCH_CROP = SKILLA_CROP
AUTO_SWITCH_LOW_SCORE_THRESHOLD = 0.5
AUTO_SWITCH_LOW_SCORE_DURATION = 35.0


# ----------------- Window detection (Windows) -----------------

if os.name == "nt":
    import ctypes
    user32 = ctypes.windll.user32

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

else:
    def get_active_window_title() -> str:
        return ""

    def is_target_window_focused() -> bool:
        return True


# ----------------- Macro Engine -----------------

class MacroEngine:
    """
    Hooks keys for the current profile, runs steps on key-down.

    Step format:
      - {"type": "press", "key": "..."}
      - {"type": "down",  "key": "..."}
      - {"type": "up",    "key": "..."}
      - {"type": "delay", "time": 0.3}
    """
    def __init__(self):
        self.profiles = {}
        self.active_profile = None
        self.profile_hooks = {}  # trigger_key -> hook handle
        self.running = False
        self.lock = threading.RLock()

    def _unhook_all_unlocked(self):
        for hook in self.profile_hooks.values():
            keyboard.unhook(hook)
        self.profile_hooks.clear()
        self.running = False

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

    def _run_steps(self, steps: list):
        for step in steps:
            t = step.get("type")

            if t == "delay":
                try:
                    secs = float(step.get("time", 0.0) or 0.0)
                except ValueError:
                    secs = 0.0
                if secs > 0:
                    time.sleep(secs)
                continue

            key = step.get("key")
            if not key:
                continue

            if t == "press":
                pydirectinput.press(key)
            elif t == "down":
                pydirectinput.keyDown(key)
            elif t == "up":
                pydirectinput.keyUp(key)

    def _on_key(self, event, trigger_key: str):
        if event.event_type != "down":
            return

        if not is_target_window_focused():
            return

        with self.lock:
            prof = self.profiles.get(self.active_profile, {})
            steps = prof.get(trigger_key)

        if not steps:
            return

        threading.Thread(target=self._run_steps, args=(steps,), daemon=True).start()


# ----------------- Step Dialog -----------------

class StepDialog(QDialog):
    """
    Create/Edit a step:
      - press/down/up -> key only
      - delay -> time only
    """
    def __init__(self, parent=None, step=None):
        super().__init__(parent)
        self.setWindowTitle("Edit Step" if step else "Add Step")
        self._result_step = None

        self.type_combo = QComboBox(self)
        self.type_combo.addItems(["press", "down", "up", "delay"])

        self.key_edit = QLineEdit(self)
        self.time_edit = QLineEdit(self)

        layout = QGridLayout()
        layout.addWidget(QLabel("Type:"), 0, 0)
        layout.addWidget(self.type_combo, 0, 1)

        layout.addWidget(QLabel("Key (press/down/up):"), 1, 0)
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
            if t not in ["press", "down", "up", "delay"]:
                t = "press"
            self.type_combo.setCurrentText(t)
            if t == "delay":
                self.time_edit.setText(str(step.get("time", 0.0)))
                self.key_edit.setText("")
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
        else:
            key = self.key_edit.text().strip()
            if not key:
                QMessageBox.warning(self, "Invalid", "Key is required for press/down/up.")
                return
            self._result_step = {"type": t, "key": key}

        self.accept()

    def get_step(self):
        return self._result_step


# ----------------- Main Window -----------------

class MainWindow(QMainWindow):
    autoSwitchProfile = pyqtSignal(str)
    autoSwitchFallback = pyqtSignal()

    def __init__(self):
        super().__init__()

        # Load Designer UI (must exist)
        uic.loadUi("main_window.ui", self)
        self.setWindowIcon(QIcon("icon.ico"))

        # your existing logic init
        self.engine = MacroEngine()
        self.config = {"default_profile": None, "profiles": {}}
        self.selected_profile = None
        self.selected_key = None
        self.overlay = OverlayStatus(OverlayConfig(x=OVERLAY_X, y=OVERLAY_Y))
        self.hotkey_handle = keyboard.add_hotkey(TOGGLE_HOTKEY, self._on_toggle_hotkey)
        self.auto_switch_enabled = True
        self.auto_switch_interval = AUTO_SWITCH_INTERVAL_DEFAULT
        self._auto_switch_inflight = False
        self._auto_switch_low_since = None
        self.auto_switch_timer = QTimer(self)
        self.auto_switch_timer.timeout.connect(self._auto_switch_tick)

        # Wire signals from UI -> your methods
        self._connect_signals()

        # Load data into UI
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
        self.autoSwitchFallback.connect(self._apply_auto_switch_fallback)

    def _on_toggle_hotkey(self):
        QTimer.singleShot(0, self._toggle_running)

    def _toggle_running(self):
        if self.engine.running:
            self.on_stop()
        else:
            self.on_start()

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
        self.btnDeleteKey.clicked.connect(self.delete_key)

        # Steps controls
        self.btnAddStep.clicked.connect(self.add_step)
        self.btnEditStep.clicked.connect(self.edit_step)
        self.btnDeleteStep.clicked.connect(self.delete_step)
        self.btnMoveUp.clicked.connect(lambda: self.move_step(-1))
        self.btnMoveDown.clicked.connect(lambda: self.move_step(1))

        # Auto profile switch
        self.btnCaptureProfileImage.clicked.connect(self.capture_profile_image)
        self.chkAutoSwitch.toggled.connect(self.on_auto_switch_toggled)
        self.editAutoSwitchInterval.editingFinished.connect(self.on_auto_switch_interval_changed)
    # ---------- Config IO ----------

    def load_from_disk(self, select_default=False):
        try:
            if os.path.exists(CONFIG_PATH):
                with open(CONFIG_PATH, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            else:
                self.config = {"default_profile": None, "profiles": {}}

            if "profiles" not in self.config or not isinstance(self.config["profiles"], dict):
                self.config["profiles"] = {}

        except Exception as e:
            QMessageBox.critical(self, "Load error", str(e))
            self.config = {"default_profile": None, "profiles": {}}

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
            self._auto_switch_low_since = None

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

            # Low-score fallback check for current profile
            current_profile = self.selected_profile or self.config.get("default_profile")
            if current_profile:
                current_img = os.path.join(CONFIGS_DIR, f"{current_profile}.png")
                if os.path.exists(current_img):
                    found_current, score_current = template_match_any(
                        current_img,
                        scene,
                        threshold=AUTO_SWITCH_LOW_SCORE_THRESHOLD,
                        return_score=True,
                    )
                    now = time.monotonic()
                    if found_current:
                        self._auto_switch_low_since = None
                    else:
                        if self._auto_switch_low_since is None:
                            self._auto_switch_low_since = now
                        elif now - self._auto_switch_low_since >= AUTO_SWITCH_LOW_SCORE_DURATION:
                            self.autoSwitchFallback.emit()

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

    def _apply_auto_switch_fallback(self):
        default_profile = self.config.get("default_profile")
        if default_profile and default_profile in self.config.get("profiles", {}):
            self.refresh_profiles(select=default_profile)
            self.comboProfile.setCurrentText(default_profile)
            self.on_profile_changed(default_profile)
        self._auto_switch_low_since = None

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


    def on_key_changed(self, key_name: str):
        key_name = key_name.strip()
        self.selected_key = key_name if key_name else None
        self.refresh_steps()

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

        prof[key_name] = []
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

    # ---------- Step controls ----------

    def current_steps(self):
        if not self.selected_profile or not self.selected_key:
            return None
        return self.config["profiles"].setdefault(self.selected_profile, {}).setdefault(self.selected_key, [])

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
    win.show()
    app.exec()


if __name__ == "__main__":
    main()
