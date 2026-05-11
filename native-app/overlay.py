import json
import os
import sys
import time
import tkinter as tk
from pathlib import Path


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

STATE_FILE = BASE_DIR / "floating_controls" / "state.json"
BG = "#1d2226"
PANEL = "#242b30"
TEXT = "#f4f7f8"
MUTED = "#a8b3ba"
BORDER = "#40505a"
GREEN = "#0d6b3c"
GREEN_ACTIVE = "#119755"
RED = "#513237"
RED_ACTIVE = "#704545"
BLUE = "#33415f"
BLUE_ACTIVE = "#4d65a0"
HIDDEN_EXIT_SECONDS = 120


def read_state():
    if not STATE_FILE.exists():
        return {"seq": 0, "follow": False, "armed": False, "label": "Idle"}

    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"seq": 0, "follow": False, "armed": False, "label": "Idle"}

    return data if isinstance(data, dict) else {"seq": 0, "follow": False, "armed": False, "label": "Idle"}


def write_state(update):
    state = read_state()
    state.update(update)
    state["updatedAt"] = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    STATE_FILE.parent.mkdir(parents=True, exist_ok=True)
    temp = STATE_FILE.with_suffix(".tmp")
    temp.write_text(json.dumps(state, indent=2), encoding="utf-8")
    temp.replace(STATE_FILE)
    return state


def command(action):
    state = read_state()
    update = {
        "seq": int(state.get("seq") or 0) + 1,
        "command": action,
    }
    if action == "arm":
        update.update({"armed": True, "follow": True, "label": "Arming active tab..."})
    elif action == "disarm":
        update.update({"armed": False, "follow": False, "label": "Idle"})
    elif action == "close":
        update.update({
            "visible": False,
            "armed": False,
            "follow": False,
            "pid": os.getpid(),
            "label": "Idle",
        })
    return write_state(update)


class FloatingControls(tk.Tk):
    def __init__(self):
        super().__init__()
        self.geometry("358x114+120+120")
        self.minsize(300, 96)
        self.overrideredirect(True)
        self.attributes("-topmost", True)
        self.configure(bg=BG, highlightbackground=BORDER, highlightthickness=1)
        self._drag = None
        self._hidden_at = None
        self._last_seq = int(read_state().get("seq") or 0)
        self.protocol("WM_DELETE_WINDOW", self._close)

        self.status = tk.StringVar(value="Idle")
        self.follow = tk.BooleanVar(value=bool(read_state().get("follow")))
        self._build()
        self._refresh()

    def _build(self):
        header = tk.Frame(self, bg=BG)
        header.pack(fill="x", padx=10, pady=(8, 4))
        self._make_draggable(header)

        mark = tk.Label(
            header,
            text="LL",
            width=3,
            height=2,
            bg="#153d2b",
            fg="#75f0a0",
            font=("Segoe UI", 10, "bold"),
            relief="solid",
            bd=1,
        )
        mark.pack(side="left", padx=(0, 10))
        self._make_draggable(mark)

        title = tk.Frame(header, bg=BG)
        title.pack(side="left", fill="x", expand=True)
        self._make_draggable(title)
        title_label = tk.Label(title, text="Floating Capture", bg=BG, fg=TEXT, font=("Segoe UI", 10, "bold"), anchor="w")
        title_label.pack(fill="x")
        self._make_draggable(title_label)
        status_label = tk.Label(title, textvariable=self.status, bg=BG, fg=MUTED, font=("Segoe UI", 8), anchor="w")
        status_label.pack(fill="x")
        self._make_draggable(status_label)

        close = tk.Button(
            header,
            text="X",
            command=self._close,
            bg="#3a252b",
            fg=TEXT,
            activebackground=RED_ACTIVE,
            activeforeground=TEXT,
            relief="solid",
            bd=1,
            width=2,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
            highlightthickness=1,
            highlightbackground="#8a555f",
        )
        close.pack(side="right", padx=(8, 0))

        controls = tk.Frame(self, bg=BG)
        controls.pack(fill="x", padx=10, pady=(4, 10))
        self.arm_button = self._button(controls, "Arm", GREEN, GREEN_ACTIVE, lambda: command("arm"))
        self.arm_button.pack(side="left", fill="x", expand=True, padx=(0, 4))
        self.disarm_button = self._button(controls, "Disarm", RED, RED_ACTIVE, lambda: command("disarm"))
        self.disarm_button.pack(side="left", fill="x", expand=True, padx=4)
        self.follow_button = self._button(controls, "Follow Off", BLUE, BLUE_ACTIVE, self._toggle_follow)
        self.follow_button.pack(side="left", fill="x", expand=True, padx=(4, 0))

    def _button(self, parent, text, bg, active_bg, callback):
        return tk.Button(
            parent,
            text=text,
            command=callback,
            height=1,
            bg=bg,
            fg=TEXT,
            activebackground=active_bg,
            activeforeground=TEXT,
            disabledforeground="#8a969d",
            relief="flat",
            bd=0,
            font=("Segoe UI", 9, "bold"),
            cursor="hand2",
        )

    def _toggle_follow(self):
        state = read_state()
        follow = not bool(state.get("follow"))
        self.follow.set(follow)
        write_state({"follow": follow})
        command("follow_on" if follow else "follow_off")
        self._render(read_state())

    def _close(self):
        command("close")
        self._hidden_at = time.monotonic()
        self.withdraw()

    def _make_draggable(self, widget):
        widget.bind("<ButtonPress-1>", self._start_drag)
        widget.bind("<B1-Motion>", self._drag_window)

    def _start_drag(self, event):
        self._drag = (event.x_root - self.winfo_x(), event.y_root - self.winfo_y())

    def _drag_window(self, event):
        if not self._drag:
            return
        x_offset, y_offset = self._drag
        self.geometry(f"+{event.x_root - x_offset}+{event.y_root - y_offset}")

    def _refresh(self):
        state = read_state()
        seq = int(state.get("seq") or 0)
        if seq > self._last_seq:
            if state.get("command") == "close":
                self._hidden_at = time.monotonic()
                self.withdraw()
            elif state.get("command") == "show":
                self._hidden_at = None
                self.deiconify()
                self.lift()
                self.attributes("-topmost", True)
                state = write_state({
                    "visible": True,
                    "pid": os.getpid(),
                    "label": state.get("label") or "Idle",
                })
        self._last_seq = max(self._last_seq, seq)
        self._render(state)
        if self._hidden_at and time.monotonic() - self._hidden_at > HIDDEN_EXIT_SECONDS:
            write_state({"visible": False, "pid": "", "armed": False, "follow": False, "label": "Idle"})
            self.destroy()
            return
        self.after(500, self._refresh)

    def _render(self, state):
        follow = bool(state.get("follow"))
        armed = bool(state.get("armed"))
        label = str(state.get("label") or ("Armed" if armed else "Idle"))
        self.status.set(label)
        self.follow_button.configure(text="Follow On" if follow else "Follow Off")
        self.arm_button.configure(state="disabled" if armed else "normal")
        self.disarm_button.configure(state="normal" if armed else "disabled")


def main():
    write_state({
        "command": "",
        "visible": True,
        "pid": os.getpid(),
        "label": read_state().get("label") or "Idle",
    })
    app = FloatingControls()
    app.mainloop()


if __name__ == "__main__":
    main()
