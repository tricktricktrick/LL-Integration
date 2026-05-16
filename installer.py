import json
import os
import shutil
import subprocess
import sys
import time
import tkinter as tk
import winreg
from pathlib import Path
from tkinter import filedialog, messagebox

APP_NAME = "LLIntegration"
NATIVE_NAME = "ll_integration_native"
FIREFOX_EXTENSION_IDS = [
    "ll-integration-release@nm088.dev",
    "ll-integration-firefox@nm088.dev",
    "ll-integration@nm088.dev",
]
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
VORTEX_PLUGINS_DIR = Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "Vortex" / "plugins"
VORTEX_PLUGIN_TARGET = VORTEX_PLUGINS_DIR / "ll-integration"
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
    return False


class InstallerApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.mode_label = "Installer"
        self.title("LL Integration Installer")
        self.geometry("760x520")
        self.resizable(False, False)
        self.configure(bg=BG)
        self._icon_image = None
        self._set_window_icon()

        existing_state = self._load_existing_install_state()
        self.mo2_exe = tk.StringVar(value=existing_state.get("mo2_path", ""))
        self.downloads_path = tk.StringVar(value=existing_state.get("mo2_downloads_path", ""))
        self.vortex_downloads_path = tk.StringVar(value=existing_state.get("vortex_downloads_path", ""))
        self.status = tk.StringVar(value="Choose MO2, Vortex, or both to begin.")
        self.delete_data = tk.BooleanVar(value=False)
        self.floating_controls = tk.BooleanVar(value=existing_state.get("floating_controls_enabled", False))
        self.experimental_toolbar = tk.BooleanVar(value=existing_state.get("experimental_toolbar", False))

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

        title = self._label(f"LL Integration {self.mode_label}", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", **pad)

        frame = tk.Frame(self, bg=BG)
        frame.pack(fill="x", **pad)

        self._path_row(frame, "ModOrganizer.exe", self.mo2_exe, self._browse_mo2, 0)
        self._path_row(frame, "MO2 downloads folder", self.downloads_path, self._browse_downloads, 1)
        self._path_row(frame, "Vortex downloads folder", self.vortex_downloads_path, self._browse_vortex_downloads, 2)

        install_text = "Install"
        self.install_button = self._button(self, install_text, self._install, accent=True, height=2)
        self.install_button.pack(fill="x", padx=14, pady=12)

        tk.Checkbutton(
            self,
            text="Install optional floating capture controls",
            variable=self.floating_controls,
            bg=BG,
            fg=MUTED,
            activebackground=BG,
            activeforeground=TEXT,
            selectcolor=PANEL,
            highlightthickness=0,
        ).pack(anchor="w", padx=14, pady=(0, 4))

        tk.Checkbutton(
            self,
            text="Enable experimental MO2 toolbar button",
            variable=self.experimental_toolbar,
            bg=BG,
            fg=MUTED,
            activebackground=BG,
            activeforeground=TEXT,
            selectcolor=PANEL,
            highlightthickness=0,
        ).pack(anchor="w", padx=14, pady=(0, 4))

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
            "This installs the native bridge in %LOCALAPPDATA%, copies the MO2 plugin when MO2 is selected, "
            "installs a lightweight Vortex extension when a Vortex downloads folder is selected, "
            "and registers browser Native Messaging for the current Windows user. "
            "Floating capture controls are optional and only provide Arm / Disarm / Follow buttons. "
            "Enable the experimental MO2 toolbar checkbox only if you want the extra MO2 toolbar button."
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

    def _browse_vortex_downloads(self) -> None:
        path = filedialog.askdirectory(title="Select Vortex downloads folder")
        if path:
            self.vortex_downloads_path.set(path)
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
        vortex_downloads_path = str(data.get("vortex_downloads_path") or "").strip()
        if mo2_path:
            state["mo2_path"] = mo2_path
        if downloads_path:
            state["mo2_downloads_path"] = downloads_path
        if vortex_downloads_path:
            state["vortex_downloads_path"] = vortex_downloads_path
        state["active_downloads_target"] = str(data.get("active_downloads_target") or "mo2")
        state["floating_controls_enabled"] = bool(data.get("floating_controls_enabled", False))
        return state

    def _validate(self) -> bool:
        mo2_text = self.mo2_exe.get().strip()
        downloads_text = self.downloads_path.get().strip()
        vortex_downloads_text = self.vortex_downloads_path.get().strip()
        ok = True
        errors = []
        mo2_enabled = bool(mo2_text or downloads_text)
        vortex_enabled = bool(vortex_downloads_text)

        if not mo2_enabled and not vortex_enabled:
            ok = False
            errors.append("Choose an MO2 install or a Vortex downloads folder.")

        if mo2_enabled and not mo2_text:
            ok = False
            errors.append("Choose ModOrganizer.exe or clear the MO2 downloads folder.")
            mo2_path = None
        else:
            mo2_path = Path(mo2_text) if mo2_text else None

        if mo2_enabled and not downloads_text:
            ok = False
            errors.append("Choose the MO2 downloads folder or clear ModOrganizer.exe.")
            downloads = None
        else:
            downloads = Path(downloads_text) if downloads_text else None

        vortex_downloads = Path(vortex_downloads_text) if vortex_downloads_text else None

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

        if vortex_downloads is not None and not vortex_downloads.exists():
            ok = False
            errors.append("Vortex downloads folder does not exist.")

        self.install_button.config(state="normal" if ok else "disabled")
        self.uninstall_button.config(
            state="normal" if (mo2_path is not None and mo2_path.parent.exists()) or VORTEX_PLUGIN_TARGET.exists() else "disabled"
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
    def _powershell_literal(self, value: Path | str) -> str:
        return "'" + str(value).replace("'", "''") + "'"

    def _stop_running_native_processes(self) -> None:
        """
        Stop LL Integration native/overlay processes without killing the browser.

        This targets only processes whose ExecutablePath or CommandLine points inside
        %LOCALAPPDATA%\\LLIntegration, so it should not kill random python.exe processes.
        """
        exact_names = [
            "ll_integration_native",
            "ll_integration_overlay",
        ]
        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    "Stop-Process -Name ll_integration_native,ll_integration_overlay -Force -ErrorAction SilentlyContinue",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=5,
                check=False,
            )
        except Exception:
            pass

        roots = [
            INSTALL_ROOT,
            NATIVE_TARGET,
        ]

        escaped_names = ",".join(self._powershell_literal(name) for name in exact_names)
        escaped_roots = ",".join(self._powershell_literal(root) for root in roots)
        command = rf"""
$names = @({escaped_names})
$roots = @({escaped_roots})
$targets = Get-CimInstance Win32_Process | Where-Object {{
    $cmd = [string]$_.CommandLine
    $exe = [string]$_.ExecutablePath
    if ($names -contains $_.Name.Replace('.exe', '')) {{
        return $true
    }}
    foreach ($root in $roots) {{
        if ($cmd -like "*$root*" -or $exe -like "*$root*") {{
            return $true
        }}
    }}
    return $false
}}

foreach ($p in $targets) {{
    try {{
        Stop-Process -Id $p.ProcessId -Force -ErrorAction SilentlyContinue
    }} catch {{}}
}}
"""

        try:
            subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-ExecutionPolicy",
                    "Bypass",
                    "-Command",
                    command,
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                timeout=8,
                check=False,
            )
        except Exception:
            pass

        time.sleep(0.5)

    def _reset_floating_controls_state(self) -> None:
        state_dir = NATIVE_TARGET / "floating_controls"
        state_path = state_dir / "state.json"
        state = {
            "seq": 0,
            "command": "",
            "follow": False,
            "armed": False,
            "visible": False,
            "pid": "",
            "label": "Idle",
        }
        try:
            state_dir.mkdir(parents=True, exist_ok=True)
            state_path.write_text(json.dumps(state, indent=2), encoding="utf-8")
        except OSError:
            pass

    def _copy2_retry(self, source: Path, target: Path, attempts: int = 5) -> None:
        last_error = None

        for attempt in range(attempts):
            try:
                shutil.copy2(source, target)
                return
            except PermissionError as exc:
                last_error = exc
                self._stop_running_native_processes()
                time.sleep(0.4 + attempt * 0.2)

        if last_error:
            raise last_error

    def _rmtree_retry(self, target: Path, attempts: int = 5) -> None:
        last_error = None
        for attempt in range(attempts):
            try:
                shutil.rmtree(target)
                return
            except PermissionError as exc:
                last_error = exc
                self._stop_running_native_processes()
                time.sleep(0.4 + attempt * 0.2)
        if last_error:
            raise last_error

    def _copytree_retry(self, source: Path, target: Path, attempts: int = 5) -> None:
        last_error = None
        for attempt in range(attempts):
            try:
                shutil.copytree(source, target, dirs_exist_ok=True)
                return
            except PermissionError as exc:
                last_error = exc
                self._stop_running_native_processes()
                time.sleep(0.4 + attempt * 0.2)
        if last_error:
            raise last_error

    def _install(self) -> None:
        if not self._validate():
            return

        try:
            mo2_text = self.mo2_exe.get().strip()
            downloads_text = self.downloads_path.get().strip()
            vortex_downloads_text = self.vortex_downloads_path.get().strip()
            mo2_path = Path(mo2_text) if mo2_text else None
            downloads_path = Path(downloads_text) if downloads_text else None
            vortex_downloads_path = Path(vortex_downloads_text) if vortex_downloads_text else None
            mo2_root = mo2_path.parent if mo2_path else None

            # Prevent the browser extension from keeping the old native host locked
            # while files are being replaced.
            self._unregister_native_messaging()
            self._stop_running_native_processes()
            self._reset_floating_controls_state()

            self._install_native_app(mo2_path, downloads_path, vortex_downloads_path)
            if mo2_root is not None:
                self._install_mo2_plugin(mo2_root)
            if vortex_downloads_path is not None:
                self._install_vortex_extension(vortex_downloads_path)

            self._register_native_messaging()

            managers = []
            if mo2_root is not None:
                managers.append("MO2")
            if vortex_downloads_path is not None:
                managers.append("Vortex")
            restart_targets = ["Firefox"]
            if mo2_root is not None:
                restart_targets.append("MO2")
            if vortex_downloads_path is not None:
                restart_targets.append("Vortex")
            next_step = f"Restart {', '.join(restart_targets)}, then "
            if mo2_root is not None and vortex_downloads_path is not None:
                next_step += "open Tools > LL Integration or the Vortex LL Integration action."
            elif mo2_root is not None:
                next_step += "open Tools > LL Integration."
            else:
                next_step += "open the Vortex LL Integration action."
            if self.experimental_toolbar.get() and mo2_root is not None:
                next_step += " The experimental toolbar button should appear in MO2."
            messagebox.showinfo(
                "Installed",
                f"LL Integration was installed for {', '.join(managers)}.\n\n{next_step}",
            )
            self.status.set(f"Installed to {INSTALL_ROOT}")
        except Exception as exc:
            messagebox.showerror("Install failed", str(exc))
            self.status.set(f"Install failed: {exc}")

    def _uninstall(self) -> None:
        mo2_text = self.mo2_exe.get().strip()
        mo2_path = Path(mo2_text) if mo2_text else None
        plugin_path = mo2_path.parent / "plugins" / "ll_integration" if mo2_path else None

        delete_data = self.delete_data.get()
        detail = (
            "This will remove installed manager plugins and browser native messaging registrations.\n\n"
            f"MO2 plugin:\n{plugin_path if plugin_path is not None else '(not selected)'}\n\n"
            f"Vortex extension:\n{VORTEX_PLUGIN_TARGET}\n\n"
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
            self._stop_running_native_processes()

            if plugin_path is not None and plugin_path.exists():
                shutil.rmtree(plugin_path)
            if VORTEX_PLUGIN_TARGET.exists():
                shutil.rmtree(VORTEX_PLUGIN_TARGET)

            if delete_data and INSTALL_ROOT.exists():
                shutil.rmtree(INSTALL_ROOT)


            messagebox.showinfo("Uninstalled", "LL Integration was uninstalled.")
            self.status.set("Uninstalled.")
        except Exception as exc:
            messagebox.showerror("Uninstall failed", str(exc))
            self.status.set(f"Uninstall failed: {exc}")

    def _install_native_app(
        self,
        mo2_path: Path | None,
        downloads_path: Path | None,
        vortex_downloads_path: Path | None,
    ) -> None:
        NATIVE_TARGET.mkdir(parents=True, exist_ok=True)
        native_exe_source = ROOT_DIR / "native-app" / "ll_integration_native.exe"
        if native_exe_source.exists():
            native_exe_target = NATIVE_TARGET / "ll_integration_native.exe"
            self._copy2_retry(native_exe_source, native_exe_target)
            native_launch_path = native_exe_target
            stale_run_bat = NATIVE_TARGET / "run.bat"
            if stale_run_bat.exists():
                stale_run_bat.unlink()
        else:
            for name in ["main.py", "overlay.py", "manager_vortex.py"]:
                self._copy2_retry(ROOT_DIR / "native-app" / name, NATIVE_TARGET / name)

            python_exe = Path(sys.executable)
            run_bat = f'@echo off\r\n"{python_exe}" "{NATIVE_TARGET / "main.py"}"\r\n'
            (NATIVE_TARGET / "run.bat").write_text(run_bat, encoding="utf-8")
            native_launch_path = NATIVE_TARGET / "run.bat"
        
        config = {
            "mo2_path": str(mo2_path or ""),
            "mo2_downloads_path": str(downloads_path or ""),
            "vortex_downloads_path": str(vortex_downloads_path or ""),
            "active_downloads_target": "mo2",
            "download_routing_mode": "auto_open_manager",
            "when_both_managers_open": "both",
            "metadata_path": str(INSTALL_ROOT / "metadata"),
            "copy_archives_to_mo2_downloads": True,
            "copy_archives_to_vortex_downloads": True,
            "overwrite_existing_downloads": True,
            "floating_controls_enabled": bool(self.floating_controls.get()),
            "experimental_toolbar": bool(self.experimental_toolbar.get()),
        }
        (NATIVE_TARGET / "config.json").write_text(json.dumps(config, indent=2), encoding="utf-8")

        overlay_exe_source = ROOT_DIR / "native-app" / "ll_integration_overlay.exe"
        if overlay_exe_source.exists():
            self._copy2_retry(overlay_exe_source, NATIVE_TARGET / "ll_integration_overlay.exe")
        else:
            overlay_py_source = ROOT_DIR / "native-app" / "overlay.py"
            if overlay_py_source.exists():
                self._copy2_retry(overlay_py_source, NATIVE_TARGET / "overlay.py")

        manager_exe_source = ROOT_DIR / "native-app" / "ll_integration_vortex_manager.exe"
        if manager_exe_source.exists():
            self._copy2_retry(manager_exe_source, NATIVE_TARGET / "ll_integration_vortex_manager.exe")
        else:
            manager_py_source = ROOT_DIR / "native-app" / "manager_vortex.py"
            if manager_py_source.exists():
                self._copy2_retry(manager_py_source, NATIVE_TARGET / "manager_vortex.py")

        firefox_manifest = {
            "name": NATIVE_NAME,
            "description": "LL Integration Native App",
            "path": str(native_launch_path),
            "type": "stdio",
            "allowed_extensions": FIREFOX_EXTENSION_IDS,
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
            self._copy2_retry(source / name, target / name)

        icons_source = source / "icons"
        icons_target = target / "icons"
        if icons_source.exists():
            icons_target.mkdir(parents=True, exist_ok=True)
            for icon in icons_source.iterdir():
                if icon.is_file():
                    self._copy2_retry(icon, icons_target / icon.name)

        tutorial_source = ROOT_DIR / "Mo2_ImageTutorial"
        tutorial_target = target / "Mo2_ImageTutorial"
        if tutorial_source.exists():
            if tutorial_target.exists():
                self._rmtree_retry(tutorial_target)
            self._copytree_retry(tutorial_source, tutorial_target)

        experimental_source = source / "experimental"
        experimental_target = target / "experimental"

        if self.experimental_toolbar.get() and experimental_source.exists():
            if experimental_target.exists():
                self._rmtree_retry(experimental_target)
            self._copytree_retry(experimental_source, experimental_target)
        elif experimental_target.exists():
            self._rmtree_retry(experimental_target)

        plugin_paths = {
            "ll_ini_path": str(NATIVE_TARGET / "downloads_storage" / "latest_ll_download.ini"),
            "cookies_path": str(NATIVE_TARGET / "cookies_storage" / "cookies_ll.json"),
            "experimental_toolbar": bool(self.experimental_toolbar.get()),
        }
        (target / "plugin_paths.json").write_text(json.dumps(plugin_paths, indent=2), encoding="utf-8")

    def _install_vortex_extension(self, vortex_downloads_path: Path) -> None:
        source = ROOT_DIR / "vortex-extension"
        if not source.exists():
            raise RuntimeError(f"Vortex extension source was not found: {source}")

        VORTEX_PLUGINS_DIR.mkdir(parents=True, exist_ok=True)
        if VORTEX_PLUGIN_TARGET.exists():
            self._rmtree_retry(VORTEX_PLUGIN_TARGET)
        self._copytree_retry(source, VORTEX_PLUGIN_TARGET)
        extension_config = {
            "nativeAppPath": str(NATIVE_TARGET),
            "nativeConfigPath": str(NATIVE_TARGET / "config.json"),
            "vortexDownloadsPath": str(vortex_downloads_path),
        }
        (VORTEX_PLUGIN_TARGET / "ll-integration.config.json").write_text(
            json.dumps(extension_config, indent=2),
            encoding="utf-8",
        )

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
