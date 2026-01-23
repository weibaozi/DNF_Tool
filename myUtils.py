import cv2
import numpy as np
from PIL import Image
import time
import win32gui
import mss
import os
from typing import Union
# FULL_CROP = (0.3883, 0.8754, 0.2282, 0.1246)
SKILLA_CROP = (0.4075, 0.9445, 0.0156, 0.0406)
TITLE="地下城与勇士：创新世纪"
def window_capture(window_title, crop=None, bgr=True):
    """
    Capture a window region using ratio-based cropping.

    Args:
        window_title (str): Window title.
        crop (tuple): (x_ratio, y_ratio, w_ratio, h_ratio)
                      values are 0.0–1.0 relative to window size
        bgr (bool):
            True  -> return BGR numpy array (OpenCV default)
            False -> return RGB numpy array

    Returns:
        np.ndarray or None
    """
    hwnd = win32gui.FindWindow(None, window_title)
    if not hwnd:
        print(f'Window with title "{window_title}" not found.')
        return None

    win_left, win_top, win_right, win_bottom = win32gui.GetWindowRect(hwnd)
    win_w = win_right - win_left
    win_h = win_bottom - win_top

    if win_w <= 0 or win_h <= 0:
        return None

    # Convert ratio crop to absolute pixels
    if crop is not None:
        x_r, y_r, w_r, h_r = crop

        left   = win_left + int(x_r * win_w)
        top    = win_top  + int(y_r * win_h)
        width  = int(w_r * win_w)
        height = int(h_r * win_h)
    else:
        left   = win_left
        top    = win_top
        width  = win_w
        height = win_h

    if width <= 0 or height <= 0:
        return None

    with mss.mss() as sct:
        monitor = {"top": top, "left": left, "width": width, "height": height}
        shot = sct.grab(monitor)  # BGRA

    frame = np.asarray(shot, dtype=np.uint8)  # BGRA

    if bgr:
        return cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
    else:
        return cv2.cvtColor(frame, cv2.COLOR_BGRA2RGB)
    
def _load_image(input_item: Union[str, np.ndarray], color_flag=cv2.IMREAD_COLOR):
    """Return BGR image for either a filepath or already-loaded np.ndarray.

    Uses unicode-safe loading for Windows paths.
    """
    if isinstance(input_item, str):
        data = np.fromfile(input_item, dtype=np.uint8)
        img = cv2.imdecode(data, color_flag)
        if img is None:
            raise SystemExit(f"Failed to load image from path: {input_item}")
        return img
    elif isinstance(input_item, np.ndarray):
        return input_item.copy()
    else:
        raise ValueError("input must be a file path or a numpy.ndarray")


def save_image(path: str, img: np.ndarray):
    """Unicode-safe image save using cv2.imencode + tofile."""
    ext = os.path.splitext(path)[1].lower() or ".png"
    ok, buf = cv2.imencode(ext, img)
    if not ok:
        raise ValueError(f"Failed to encode image for: {path}")
    buf.tofile(path)
    
def template_match_any(
    templ_input: Union[str, np.ndarray],
    scene_input: Union[str, np.ndarray],
    write_result: bool = False,
    threshold: float = 0.70,
    scales: np.ndarray = None,
    return_score: bool = False,
) -> bool:
    """
    Template-matching ONLY (no SIFT).

    - Accepts template/scene as file paths or numpy BGR arrays.
    - If template has alpha, blends onto white.
    - Multi-scale matchTemplate with TM_CCOEFF_NORMED.
    - Returns True/False based on `threshold`.

    Args:
        threshold: required best match score to return True (typical 0.7~0.9).
        scales: optional array of scales to try (default: np.linspace(0.3, 1.5, 60)[::-1])
    """
    templ = _load_image(templ_input, color_flag=cv2.IMREAD_UNCHANGED)
    img = _load_image(scene_input, color_flag=cv2.IMREAD_COLOR)

    # Handle alpha template by blending onto white background
    if templ.ndim == 3 and templ.shape[2] == 4:
        alpha = templ[:, :, 3].astype(np.float32) / 255.0
        rgb = templ[:, :, :3].astype(np.float32)
        bg = np.ones_like(rgb) * 255.0
        templ = (rgb * alpha[:, :, None] + bg * (1.0 - alpha[:, :, None])).astype(np.uint8)

    gray_t = cv2.cvtColor(templ, cv2.COLOR_BGR2GRAY)
    gray_i = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Optional contrast boost (often helps template matching)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8, 8))
    gray_t = clahe.apply(gray_t)
    gray_i = clahe.apply(gray_i)

    t_h, t_w = gray_t.shape[:2]

    if scales is None:
        scales = np.linspace(0.3, 1.5, 60)[::-1]

    best_score = -1.0
    best_loc = None
    best_scale = None
    best_size = None

    for scale in scales:
        w_s = int(t_w * scale)
        h_s = int(t_h * scale)

        if w_s < 2 or h_s < 2 or w_s >= gray_i.shape[1] or h_s >= gray_i.shape[0]:
            continue

        resized = cv2.resize(gray_t, (w_s, h_s), interpolation=cv2.INTER_AREA if scale < 1 else cv2.INTER_CUBIC)
        res = cv2.matchTemplate(gray_i, resized, cv2.TM_CCOEFF_NORMED)
        _, maxv, _, maxloc = cv2.minMaxLoc(res)
        # print(maxv)
        if maxv > best_score:
            best_score = float(maxv)
            best_loc = maxloc
            best_scale = float(scale)
            best_size = (w_s, h_s)

    found = best_score >= float(threshold)

    if write_result:
        out = img.copy()
        if best_loc is not None and best_size is not None:
            x, y = best_loc
            w_s, h_s = best_size
            cv2.rectangle(out, (x, y), (x + w_s, y + h_s), (0, 255, 0) if found else (0, 0, 255), 2)
            cv2.putText(out, f"score={best_score:.3f} scale={best_scale:.3f}",
                        (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 255), 2)
        cv2.imwrite("result_template_match.png", out)

    if return_score:
        return found, best_score
    return found