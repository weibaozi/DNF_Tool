import queue
import threading
import tkinter as tk
from dataclasses import dataclass


@dataclass(frozen=True)
class OverlayConfig:
    x: int = 50
    y: int = 50
    font: str = "Segoe UI"
    font_size: int = 12
    fg: str = "#ffffff"
    bg: str = "#000000"
    alpha: float = 0.7


class OverlayStatus:
    """Small always-on-top overlay showing status text.

    Runs tkinter in a dedicated thread to avoid blocking the Qt event loop.
    """

    def __init__(self, config: OverlayConfig | None = None):
        self.config = config or OverlayConfig()
        self._queue: queue.Queue[str | None] = queue.Queue()
        self._thread = threading.Thread(target=self._run, daemon=True)
        self._ready = threading.Event()
        self._thread.start()
        self._ready.wait(timeout=2.0)

    def update_text(self, text: str):
        self._queue.put(text)

    def close(self):
        self._queue.put(None)

    def _run(self):
        root = tk.Tk()
        root.withdraw()

        win = tk.Toplevel(root)
        win.overrideredirect(True)
        win.attributes("-topmost", True)
        try:
            win.attributes("-alpha", float(self.config.alpha))
        except Exception:
            pass

        win.configure(bg=self.config.bg)
        win.geometry(f"+{int(self.config.x)}+{int(self.config.y)}")

        label = tk.Label(
            win,
            text="",
            fg=self.config.fg,
            bg=self.config.bg,
            font=(self.config.font, int(self.config.font_size)),
            padx=8,
            pady=4,
        )
        label.pack()

        def poll_queue():
            try:
                while True:
                    item = self._queue.get_nowait()
                    if item is None:
                        win.destroy()
                        root.destroy()
                        return
                    label.config(text=item)
            except queue.Empty:
                pass
            win.after(100, poll_queue)

        self._ready.set()
        win.after(100, poll_queue)
        root.mainloop()
