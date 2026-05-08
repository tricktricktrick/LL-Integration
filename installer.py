import json
import os
import shutil
import sys
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import filedialog, messagebox


APP_NAME = "LLIntegration"
NATIVE_NAME = "ll_integration_native"
FIREFOX_EXTENSION_ID = "ll-integration@nm088.dev"
CHROMIUM_EXTENSION_ID = "ndnmgkboipaepgndebnikcnicechokln"
CHROMIUM_EXTENSION_ORIGIN = f"chrome-extension://{CHROMIUM_EXTENSION_ID}/"
FIREFOX_NATIVE_HOST_KEYS = [
    rf"Software\Mozilla\NativeMessagingHosts\{NATIVE_NAME}",
]
CHROMIUM_NATIVE_HOST_KEYS = [
    rf"Software\Google\Chrome\NativeMessagingHosts\{NATIVE_NAME}",
    rf"Software\Chromium\NativeMessagingHosts\{NATIVE_NAME}",
    rf"Software\Opera Software\NativeMessagingHosts\{NATIVE_NAME}",
    rf"Software\Microsoft\Edge\NativeMessagingHosts\{NATIVE_NAME}",
]
ROOT_DIR = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
INSTALL_ROOT = Path(os.environ.get("LOCALAPPDATA", Path.home() / "AppData" / "Local")) / APP_NAME
NATIVE_TARGET = INSTALL_ROOT / "native-app"
INSTALLER_ICON = ROOT_DIR / "installer_icon.png"
BG = "#15191d"
PANEL = "#20262b"
PANEL_ALT = "#2a3137"
TEXT = "#f4f7f8"
MUTED = "#a8b3ba"
BORDER = "#3a444b"
GREEN = "#26d07c"
GREEN_DARK = "#0d6b3c"
WARN = "#ffb454"
ERROR = "#ff6b6b"
DISABLED = "#637078"


def installer_with_toolbar() -> bool:
    args = {arg.lower() for arg in sys.argv[1:]}
    exe_name = Path(sys.argv[0]).stem.lower()
    if "--without-toolbar" in args or "--stable" in args:
        return False
    return (
        "--with-toolbar" in args
        or "--experimental" in args
        or "withtoolbar" in exe_name
        or "with-toolbar" in exe_name
        or "experimental" in exe_name
    )


class InstallerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.with_toolbar = installer_with_toolbar()
        self.mode_label = "With Toolbar" if self.with_toolbar else "Stable"
        self.title(f"LL Integration Installer - {self.mode_label}")
        self.geometry("680x360")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._icon_image = None
        self._set_window_icon()

        existing_state = self._load_existing_install_state()
        self.mo2_exe = tk.StringVar(value=existing_state.get("mo2_path", ""))
        self.downloads_path = tk.StringVar(value=existing_state.get("mo2_downloads_path", ""))
        self.status = tk.StringVar(value="Choose ModOrganizer.exe to begin.")
        self.delete_data = tk.BooleanVar(value=False)

        self._build_ui()
        self._validate()

    def _set_window_icon(self) -> None:
        if not INSTALLER_ICON.exists():
            return

        try:
            self._icon_image = tk.PhotoImage(file=str(INSTALLER_ICON))
            self.iconphoto(True, self._icon_image)
        except tk.TclError:
            pass

    def _build_ui(self) -> None:
        pad = {"padx": 14, "pady": 8}

        title = self._label(f"LL Integration Installer - {self.mode_label}", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", **pad)

        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="x", **pad)

        self._path_row(frame, "ModOrganizer.exe", self.mo2_exe, self._browse_mo2, 0)
        self._path_row(frame, "MO2 downloads folder", self.downloads_path, self._browse_downloads, 1)

        install_text = "Install With Toolbar" if self.with_toolbar else "Install Stable"
        self.install_button = self._button(self, install_text, self._install, accent=True, height=2)
        self.install_button.pack(fill="x", padx=14, pady=12)

        uninstall_frame = tk.Frame(self, bg=BG)
        uninstall_frame.pack(fill="x", padx=14, pady=2)
        self.uninstall_button = self._button(uninstall_frame, "Uninstall", self._uninstall)
        self.uninstall_button.pack(side="left")
        tk.Checkbutton(
            uninstall_frame,
            text="Also delete cookies, metadata, and config",
            variable=self.delete_data,
            bg=BG,
            fg=MUTED,
            activebackground=BG,
            activeforeground=TEXT,
            selectcolor=PANEL,
            highlightthickness=0,
        ).pack(side="left", padx=12)

        self.status_label = tk.Label(
            self,
            textvariable=self.status,
            anchor="w",
            justify="left",
            fg=WARN,
            bg=BG,
            font=("Segoe UI", 9),
        )
        self.status_label.pack(
            fill="x", padx=14, pady=8
        )

        note = (
            "This installs the native bridge in %LOCALAPPDATA%, copies the MO2 plugin, "
            "and registers browser Native Messaging for the current Windows user. "
            + (
                "This build enables the experimental MO2 toolbar button."
                if self.with_toolbar
                else "This build uses the stable Tools > LL Integration menu only."
            )
        )
        self._label(note, wraplength=640, fg=MUTED).pack(
            fill="x", padx=14, pady=6
        )

    def _path_row(self, parent, label, variable, command, row) -> None:
        self._label(label, master=parent).grid(row=row, column=0, sticky="w", pady=6)
        entry = tk.Entry(
            parent,
            textvariable=variable,
            width=70,
            bg=PANEL,
            fg=TEXT,
            insertbackground=TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=BORDER,
            highlightcolor=GREEN,
            disabledbackground=PANEL,
            disabledforeground=DISABLED,
        )
        entry.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        entry.bind("<KeyRelease>", lambda _event: self._validate())
        self._button(parent, "Browse", command).grid(row=row, column=2, pady=6)
        parent.grid_columnconfigure(1, weight=1)

    def _label(self, text: str, **kwargs) -> tk.Label:
        kwargs.setdefault("bg", BG)
        kwargs.setdefault("fg", TEXT)
        kwargs.setdefault("anchor", "w")
        kwargs.setdefault("justify", "left")
        kwargs.setdefault("font", ("Segoe UI", 9))
        return tk.Label(self if "master" not in kwargs else kwargs.pop("master"), text=text, **kwargs)

    def _button(self, parent, text: str, command, accent: bool = False, **kwargs) -> tk.Button:
        bg = GREEN_DARK if accent else PANEL_ALT
        active_bg = "#0f7f49" if accent else "#343d44"
        button = tk.Button(
            parent,
            text=text,
            command=command,
            bg=bg,
            fg=TEXT,
            activebackground=active_bg,
            activeforeground=TEXT,
            disabledforeground=DISABLED,
            relief="flat",
            bd=0,
            highlightthickness=1,
            highlightbackground=BORDER,
            font=("Segoe UI", 9, "bold" if accent else "normal"),
            cursor="hand2",
            **kwargs,
        )
        return button

    def _browse_mo2(self) -> None:
        path = filedialog.askopenfilename(
            title="Select ModOrganizer.exe",
            filetypes=[("ModOrganizer.exe", "ModOrganizer.exe"), ("Executable", "*.exe"), ("All files", "*.*")],
        )
        if path:
            self.mo2_exe.set(path)
            mo2_root = Path(path).parent
            guessed_downloads = self._guess_downloads_path(mo2_root)
            if guessed_downloads:
                self.downloads_path.set(str(guessed_downloads))
            self._validate()

    def _browse_downloads(self) -> None:
        path = filedialog.askdirectory(title="Select MO2 downloads folder")
        if path:
            self.downloads_path.set(path)
            self._validate()

    def _load_existing_install_state(self) -> dict[str, str]:
        config_path = NATIVE_TARGET / "config.json"
        if not config_path.exists():
            return {}

        try:
            data = json.loads(config_path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}

        if not isinstance(data, dict):
            return {}

        state = {}
        mo2_path = str(data.get("mo2_path") or "").strip()
        downloads_path = str(data.get("mo2_downloads_path") or "").strip()
        if mo2_path:
            state["mo2_path"] = mo2_path
        if downloads_path:
            state["mo2_downloads_path"] = downloads_path
        return state

    def _validate(self) -> bool:
        mo2_text = self.mo2_exe.get().strip()
        downloads_text = self.downloads_path.get().strip()
        ok = True
        errors = []

        if not mo2_text:
            ok = False
            errors.append("Choose ModOrganizer.exe.")
            mo2_path = None
        else:
            mo2_path = Path(mo2_text)

        if not downloads_text:
            ok = False
            errors.append("Choose the MO2 downloads folder.")
            downloads = None
        else:
            downloads = Path(downloads_text)

        if mo2_path is None:
            plugins_path = None
        else:
            plugins_path = mo2_path.parent / "plugins"

        if mo2_path is not None and (not mo2_path.exists() or mo2_path.name.lower() != "modorganizer.exe"):
            ok = False
            errors.append("ModOrganizer.exe was not found.")

        if plugins_path is not None and not plugins_path.exists():
            ok = False
            errors.append("MO2 plugins folder was not found next to ModOrganizer.exe.")

        if downloads is not None and not downloads.exists():
            ok = False
            errors.append("MO2 downloads folder does not exist.")

        self.install_button.config(state="normal" if ok else "disabled")
        self.uninstall_button.config(
            state="normal" if mo2_path is not None and mo2_path.parent.exists() else "disabled"
        )
        if hasattr(self, "status_label"):
            self.status_label.config(fg=GREEN if ok else WARN)
        self.status.set("Ready to install." if ok else "\n".join(errors))
        return ok

    def _guess_downloads_path(self, mo2_root: Path) -> Path | None:
        candidates = [
            mo2_root / "downloads",
            mo2_root.parent / "Mods Downloaded",
            mo2_root.parent / "downloads",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return None

    def _install(self) -> None:
        if not self._validate():
            return

        try:
            mo2_path = Path(self.mo2_exe.get().strip())
            downloads_path = Path(self.downloads_path.get().strip())
            mo2_root = mo2_path.parent

            self._install_native_app(mo2_path, downloads_path)
            self._install_mo2_plugin(mo2_root)
            self._register_native_messaging()

            messagebox.showinfo(
                "Installed",
                "LL Integration was installed.\n\n"
                + (
                    "Restart Firefox and MO2. The experimental toolbar button should appear in MO2."
                    if self.with_toolbar
                    else "Restart Firefox and MO2, then open Tools > LL Integration."
                ),
            )
            self.status.set(f"Installed to {INSTALL_ROOT}")
        except Exception as exc:
            messagebox.showerror("Install failed", str(exc))
            self.status.set(f"Install failed: {exc}")

    def _uninstall(self) -> None:
        mo2_path = Path(self.mo2_exe.get().strip())
        plugin_path = mo2_path.parent / "plugins" / "ll_integration"

        delete_data = self.delete_data.get()
        detail = (
            "This will remove the MO2 plugin and browser native messaging registrations.\n\n"
            f"MO2 plugin:\n{plugin_path}\n\n"
            "Native app registrations:\n"
            + "\n".join(f"HKCU\\{key}" for key in FIREFOX_NATIVE_HOST_KEYS + CHROMIUM_NATIVE_HOST_KEYS)
            + "\n\n"
        )
        if delete_data:
            detail += f"It will also delete:\n{INSTALL_ROOT}\n"
        else:
            detail += f"It will keep cookies, metadata, and config in:\n{INSTALL_ROOT}\n"

        if not messagebox.askyesno("Uninstall LL Integration", detail):
            return

        try:
            self._unregister_native_messaging()
            if plugin_path.exists():
                shutil.rmtree(plugin_path)
            if delete_data and INSTALL_ROOT.exists():
                shutil.rmtree(INSTALL_ROOT)

            messagebox.showinfo("Uninstalled", "LL Integration was uninstalled.")
            self.status.set("Uninstalled.")
        except Exception as exc:
            messagebox.showerror("Uninstall failed", str(exc))
            self.status.set(f"Uninstall failed: {exc}")

    def _install_native_app(self, mo2_path: Path, downloads_path: Path) -> None:
        NATIVE_TARGET.mkdir(parents=True, exist_ok=True)
        native_exe_source = ROOT_DIR / "native-app" / "ll_integration_native.exe"
        if native_exe_source.exists():
            native_exe_target = NATIVE_TARGET / "ll_integration_native.exe"
            shutil.copy2(native_exe_source, native_exe_target)
            native_launch_path = native_exe_target
            stale_run_bat = NATIVE_TARGET / "run.bat"
            if stale_run_bat.exists():
                stale_run_bat.unlink()
        else:
            for name in ["main.py"]:
                shutil.copy2(ROOT_DIR / "native-app" / name, NATIVE_TARGET / name)

            python_exe = Path(sys.executable)
            run_bat = f'@echo off\r\n"{python_exe}" "{NATIVE_TARGET / "main.py"}"\r\n'
            (NATIVE_TARGET / "run.bat").write_text(run_bat, encoding="utf-8")
            native_launch_path = NATIVE_TARGET / "run.bat"

        config = {
            "mo2_path": str(mo2_path),
            "mo2_downloads_path": str(downloads_path),
            "metadata_path": str(INSTALL_ROOT / "metadata"),
            "copy_archives_to_mo2_downloads": True,
            "overwrite_existing_downloads": True,
        }
        (NATIVE_TARGET / "config.json").write_text(json.dumps(config, indent=2), encoding="utf-8")

        firefox_manifest = {
            "name": NATIVE_NAME,
            "description": "LL Integration Native App",
            "path": str(native_launch_path),
            "type": "stdio",
            "allowed_extensions": [FIREFOX_EXTENSION_ID],
        }
        firefox_manifest_text = json.dumps(firefox_manifest, indent=2)
        (NATIVE_TARGET / "manifest.json").write_text(firefox_manifest_text, encoding="utf-8")
        (NATIVE_TARGET / "manifest.firefox.json").write_text(firefox_manifest_text, encoding="utf-8")

        chromium_manifest = {
            "name": NATIVE_NAME,
            "description": "LL Integration Native App",
            "path": str(native_launch_path),
            "type": "stdio",
            "allowed_origins": [CHROMIUM_EXTENSION_ORIGIN],
        }
        (NATIVE_TARGET / "manifest.chromium.json").write_text(
            json.dumps(chromium_manifest, indent=2),
            encoding="utf-8",
        )

    def _install_mo2_plugin(self, mo2_root: Path) -> None:
        source = ROOT_DIR / "mo2-plugin"
        target = mo2_root / "plugins" / "ll_integration"
        target.mkdir(parents=True, exist_ok=True)

        for name in ["__init__.py", "plugin.py", "utils.py", "check_update.py", "LL.sample.ini"]:
            shutil.copy2(source / name, target / name)

        icons_source = source / "icons"
        icons_target = target / "icons"
        if icons_source.exists():
            icons_target.mkdir(parents=True, exist_ok=True)
            for icon in icons_source.iterdir():
                if icon.is_file():
                    shutil.copy2(icon, icons_target / icon.name)

        experimental_source = source / "experimental"
        experimental_target = target / "experimental"
        if self.with_toolbar and experimental_source.exists():
            if experimental_target.exists():
                shutil.rmtree(experimental_target)
            shutil.copytree(experimental_source, experimental_target)
        elif experimental_target.exists():
            shutil.rmtree(experimental_target)

        plugin_paths = {
            "ll_ini_path": str(NATIVE_TARGET / "downloads_storage" / "latest_ll_download.ini"),
            "cookies_path": str(NATIVE_TARGET / "cookies_storage" / "cookies_ll.json"),
            "experimental_toolbar": self.with_toolbar,
        }
        (target / "plugin_paths.json").write_text(json.dumps(plugin_paths, indent=2), encoding="utf-8")

    def _register_native_messaging(self) -> None:
        for key_path in FIREFOX_NATIVE_HOST_KEYS:
            self._register_native_host(key_path, NATIVE_TARGET / "manifest.firefox.json")

        for key_path in CHROMIUM_NATIVE_HOST_KEYS:
            self._register_native_host(key_path, NATIVE_TARGET / "manifest.chromium.json")

    def _register_native_host(self, key_path: str, manifest_path: Path) -> None:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, str(manifest_path))

    def _unregister_native_messaging(self) -> None:
        for key_path in FIREFOX_NATIVE_HOST_KEYS + CHROMIUM_NATIVE_HOST_KEYS:
            try:
                winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
            except FileNotFoundError:
                pass


if __name__ == "__main__":
    app = InstallerApp()
    app.mainloop()
