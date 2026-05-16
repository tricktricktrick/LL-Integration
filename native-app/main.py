import json
import os
import re
import shutil
import subprocess
import struct
import sys
import hashlib
from pathlib import Path
from urllib.parse import parse_qs, urlparse

if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
CONFIG_FILE = BASE_DIR / "config.json"
OUT_FILE = BASE_DIR / "cookies_storage" / "cookies_ll.json"
DOWNLOAD_EVENTS_FILE = BASE_DIR / "downloads_storage" / "download_events.json"
DOWNLOAD_COMPLETIONS_FILE = BASE_DIR / "downloads_storage" / "download_completions.json"
LATEST_INI_FILE = BASE_DIR / "downloads_storage" / "latest_ll_download.ini"
FLOATING_CONTROLS_STATE_FILE = BASE_DIR / "floating_controls" / "state.json"
ARCHIVE_SUFFIXES = {".7z", ".zip", ".rar"}
QUICK_HASH_CHUNK_SIZE = 1024 * 1024
DEFAULT_CONFIG = {
    "mo2_path": "",
    "mo2_downloads_path": "",
    "vortex_downloads_path": "",
    "active_downloads_target": "mo2",
    "download_routing_mode": "auto_open_manager",
    "when_both_managers_open": "both",
    "metadata_path": str(BASE_DIR / "metadata"),
    "copy_archives_to_mo2_downloads": True,
    "copy_archives_to_vortex_downloads": True,
    "overwrite_existing_downloads": True,
    "floating_controls_enabled": False,
}

def read_message():
    raw_len = sys.stdin.buffer.read(4)
    if not raw_len:
        return None

    msg_len = struct.unpack("@I", raw_len)[0]
    return json.loads(sys.stdin.buffer.read(msg_len).decode("utf-8"))

def send_message(message):
    data = json.dumps(message).encode("utf-8")
    sys.stdout.buffer.write(struct.pack("@I", len(data)))
    sys.stdout.buffer.write(data)
    sys.stdout.buffer.flush()

def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")

def save_json_atomic(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_suffix(".tmp")
    temp.write_text(json.dumps(data, indent=2), encoding="utf-8")
    temp.replace(path)

def load_config():
    if not CONFIG_FILE.exists():
        save_json(CONFIG_FILE, DEFAULT_CONFIG)
        return DEFAULT_CONFIG.copy()

    try:
        data = json.loads(CONFIG_FILE.read_text(encoding="utf-8-sig"))
    except json.JSONDecodeError:
        return DEFAULT_CONFIG.copy()

    config = DEFAULT_CONFIG.copy()
    if isinstance(data, dict):
        config.update(data)
    return config

def load_json_list(path):
    if not path.exists():
        return []

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []

    return data if isinstance(data, list) else []

def load_json_object(path, default=None):
    if default is None:
        default = {}
    if not path.exists():
        return default.copy()

    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return default.copy()

    return data if isinstance(data, dict) else default.copy()

def ini_value(value):
    if value is None:
        return ""
    return str(value).replace("\r", " ").replace("\n", " ").strip()

def ll_file_id(page_url):
    if not page_url:
        return ""

    match = re.search(r"/files/file/(\d+)", page_url)
    return match.group(1) if match else ""

def ll_resource_id(download_url):
    if not download_url:
        return ""

    query = parse_qs(urlparse(download_url).query)
    values = query.get("r")
    return values[0] if values else ""

def source_type_for_event(event):
    source_type = str(event.get("sourceType") or "").strip().lower()
    if source_type:
        return source_type

    if event.get("action") == "capture_external_archive":
        return "external"

    page_url = str(event.get("pageUrl") or "")
    host = (urlparse(page_url).hostname or "").lower()

    if host == "www.loverslab.com" or host == "loverslab.com" or host.endswith(".loverslab.com"):
        return "loverslab"

    if host == "www.nexusmods.com" or host == "nexusmods.com" or host.endswith(".nexusmods.com"):
        return "nexus"

    if host == "dwemermods.com" or host.endswith(".dwemermods.com"):
        return "dwemermods"

    return "external"

def source_update_supported(source_type):
    return source_type == "loverslab"

def floating_controls_default_state():
    return {
        "seq": 0,
        "command": "",
        "follow": False,
        "armed": False,
        "visible": False,
        "label": "Idle",
    }

def floating_controls_state():
    state = floating_controls_default_state()
    state.update(load_json_object(FLOATING_CONTROLS_STATE_FILE, state))
    return state

def write_floating_controls_state(update):
    state = floating_controls_state()
    state.update(update)
    save_json_atomic(FLOATING_CONTROLS_STATE_FILE, state)
    return state

def process_name_running(names):
    wanted = {name.lower() for name in names}

    if sys.platform == "win32":
        try:
            import subprocess
            result = subprocess.run(
                ["tasklist", "/FO", "CSV", "/NH"],
                capture_output=True,
                text=True,
                creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            )
            output = result.stdout.lower()
            return any(name.lower() in output for name in wanted)
        except Exception:
            return False

    return False


def mo2_is_running(config):
    return process_name_running([
        "ModOrganizer.exe",
        "ModOrganizer2.exe",
    ])


def vortex_is_running(config):
    return process_name_running([
        "Vortex.exe",
    ])

def process_is_running(pid):
    try:
        pid = int(pid)
    except (TypeError, ValueError):
        return False
    if pid <= 0:
        return False

    if sys.platform == "win32":
        try:
            import ctypes
            from ctypes import wintypes

            kernel32 = ctypes.windll.kernel32
            process_query_limited_information = 0x1000
            still_active = 259
            handle = kernel32.OpenProcess(process_query_limited_information, False, pid)
            if not handle:
                return False
            try:
                exit_code = wintypes.DWORD()
                if not kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code)):
                    return False
                return exit_code.value == still_active
            finally:
                kernel32.CloseHandle(handle)
        except Exception:
            return False

    try:
        os.kill(pid, 0)
    except OSError:
        return False
    return True

def overlay_launcher():
    overlay_exe = BASE_DIR / "ll_integration_overlay.exe"
    if overlay_exe.exists():
        return [str(overlay_exe)]

    overlay_py = BASE_DIR / "overlay.py"
    if overlay_py.exists():
        pythonw = Path(sys.executable).with_name("pythonw.exe")
        python = pythonw if pythonw.exists() else Path(sys.executable)
        return [str(python), str(overlay_py)]

    return None

def open_floating_controls():
    config = load_config()
    if not config.get("floating_controls_enabled", False):
        return {
            "ok": False,
            "error": "Floating Capture Controls are not installed/enabled. Run the installer and enable the optional floating controls.",
            "disabled": True,
        }

    launcher = overlay_launcher()
    if not launcher:
        return {"ok": False, "error": "Floating Capture Controls were not found in the native app folder."}

    state = floating_controls_state()
    overlay_running = process_is_running(state.get("pid"))
    if state.get("visible") and not overlay_running:
        state = write_floating_controls_state({
            "visible": False,
            "pid": "",
            "armed": False,
            "label": "Idle",
        })
        overlay_running = False

    if state.get("visible"):
        update = {
            "seq": int(state.get("seq") or 0) + 1,
            "command": "close",
            "visible": False,
            "armed": False,
            "follow": False,
            "label": "Idle",
        }
        if overlay_running:
            update["pid"] = state.get("pid")
        else:
            update["pid"] = ""
        state = write_floating_controls_state(update)
        return {"ok": True, "closed": True, "disarmed": True, "state": state}

    if overlay_running:
        state = write_floating_controls_state({
            "seq": int(state.get("seq") or 0) + 1,
            "command": "show",
            "visible": True,
            "pid": state.get("pid"),
            "label": state.get("label") or "Idle",
        })
        return {"ok": True, "reused": True, "state": state}

    try:
        process = subprocess.Popen(
            launcher,
            cwd=str(BASE_DIR),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
            creationflags=getattr(subprocess, "DETACHED_PROCESS", 0) | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0),
        )
    except OSError as exc:
        return {"ok": False, "error": str(exc)}

    state = write_floating_controls_state({"command": "", "visible": True, "pid": process.pid, "label": "Idle"})
    return {"ok": True, "state": state}

def archive_quick_hash(path):
    size = path.stat().st_size
    digest = hashlib.sha256()
    digest.update(str(size).encode("ascii"))

    with path.open("rb") as file:
        digest.update(file.read(QUICK_HASH_CHUNK_SIZE))
        if size > QUICK_HASH_CHUNK_SIZE:
            file.seek(max(size - QUICK_HASH_CHUNK_SIZE, 0))
            digest.update(file.read(QUICK_HASH_CHUNK_SIZE))

    return digest.hexdigest()

def unique_path(path):
    if not path.exists():
        return path

    stem = path.stem
    suffix = path.suffix
    parent = path.parent
    counter = 2
    while True:
        candidate = parent / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1

def active_downloads_path(config):
    target = str(config.get("active_downloads_target") or "mo2").lower()
    if target == "vortex":
        vortex_text = str(config.get("vortex_downloads_path") or "").strip()
        if vortex_text:
            return Path(vortex_text)
    mo2_text = str(config.get("mo2_downloads_path") or "").strip()
    if mo2_text:
        return Path(mo2_text)
    vortex_text = str(config.get("vortex_downloads_path") or "").strip()
    return Path(vortex_text) if vortex_text else None


def ll_ini_lines(event, archive_path=None, browser_download_url=None, completed_at=None):
    download = event.get("download") or {}
    archive = Path(archive_path) if archive_path else None
    archive_size = archive.stat().st_size if archive and archive.exists() else ""
    quick_hash = archive_quick_hash(archive) if archive and archive.exists() else ""
    source_type = source_type_for_event(event)
    update_supported = source_update_supported(source_type)
    file_name = download.get("name") or (archive.name if archive else "")
    download_url = download.get("url") or browser_download_url or ""
    version = str(download.get("version") or "").strip()
    if not version and not update_supported:
        version = "0.0.0"

    return [
        "[LoversLab]",
        f"source={ini_value(source_type)}",
        f"ll_file_id={ini_value(ll_file_id(event.get('pageUrl')))}",
        f"ll_resource_id={ini_value(ll_resource_id(download_url))}",
        f"page_url={ini_value(event.get('pageUrl'))}",
        f"page_title={ini_value(event.get('pageTitle'))}",
        f"download_url={ini_value(download_url)}",
        f"file_name={ini_value(file_name)}",
        f"original_archive_name={ini_value(file_name)}",
        f"archive_name={ini_value(archive.name if archive else file_name)}",
        f"archive_size_bytes={ini_value(archive_size)}",
        f"archive_quick_hash={ini_value(quick_hash)}",
        f"version={ini_value(version)}",
        f"size={ini_value(download.get('size'))}",
        f"date_iso={ini_value(download.get('date_iso'))}",
        f"captured_at={ini_value(event.get('capturedAt'))}",
        f"archive_path={ini_value(archive_path)}",
        f"browser_download_url={ini_value(browser_download_url)}",
        f"completed_at={ini_value(completed_at)}",
        f"update_mode={'manual' if update_supported else 'skip'}",
        f"fixed_version={'false' if update_supported else 'true'}",
        f"manual_update={'false' if update_supported else 'true'}",
        f"skip_update_check={'false' if update_supported else 'true'}",
        "",
    ]

def write_latest_ini(path, event):
    lines = ll_ini_lines(event)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")

def is_archive(path):
    return path.suffix.lower() in ARCHIVE_SUFFIXES

def write_sidecar_files(archive_path, event, browser_download_url, completed_at):
    config = load_config()
    source_archive = Path(archive_path)
    archive, copy_error, copied_targets = copy_archive_to_manager_downloads(source_archive, config)
    if not is_archive(archive):
        raise ValueError(f"Not a supported archive: {archive}")

    ini_path = archive.with_name(f"{archive.name}.ll.ini")
    json_path = archive.with_name(f"{archive.name}.ll.json")
    metadata_downloads = Path(str(config.get("metadata_path"))) / "downloads"
    metadata_ini_path = metadata_downloads / f"{archive.name}.ll.ini"
    metadata_json_path = metadata_downloads / f"{archive.name}.ll.json"
    payload = {
        "sourceType": source_type_for_event(event),
        "llFileId": ll_file_id(event.get("pageUrl")),
        "llResourceId": ll_resource_id((event.get("download") or {}).get("url") or browser_download_url),
        "archiveName": archive.name,
        "originalArchiveName": (event.get("download") or {}).get("name"),
        "archiveSizeBytes": archive.stat().st_size,
        "archiveQuickHash": archive_quick_hash(archive),
        "archivePath": str(archive),
        "sourceArchivePath": str(source_archive),
        "copyError": copy_error,
        "browserDownloadUrl": browser_download_url,
        "completedAt": completed_at,
        "event": event,
        "copiedTargets": copied_targets,
    }

    ini_text = "\n".join(ll_ini_lines(event, archive, browser_download_url, completed_at))

    for copied in copied_targets:
        copied_archive = Path(copied.get("archivePath") or "")
        if copied_archive.exists():
            copied_ini = copied_archive.with_name(f"{copied_archive.name}.ll.ini")
            copied_json = copied_archive.with_name(f"{copied_archive.name}.ll.json")
            copied_ini.write_text(ini_text, encoding="utf-8")
            save_json(copied_json, payload)

    ini_path.write_text(ini_text, encoding="utf-8")
    save_json(json_path, payload)
    metadata_ini_path.parent.mkdir(parents=True, exist_ok=True)
    metadata_ini_path.write_text(ini_text, encoding="utf-8")
    save_json(metadata_json_path, payload)
    return ini_path, json_path, metadata_ini_path, metadata_json_path, archive, copy_error

def copy_archive_to_manager_downloads(archive, config):
    targets = manager_download_targets(config)

    if not targets:
        return archive, None, []

    copied = []
    errors = []

    for manager, downloads_path in targets:
        if manager == "mo2" and not config.get("copy_archives_to_mo2_downloads", True):
            continue
        if manager == "vortex" and not config.get("copy_archives_to_vortex_downloads", True):
            continue

        downloads_path.mkdir(parents=True, exist_ok=True)
        target = downloads_path / archive.name

        if archive.resolve() != target.resolve():
            if target.exists() and not config.get("overwrite_existing_downloads", True):
                target = unique_path(target)

            try:
                shutil.copy2(archive, target)
            except OSError as exc:
                errors.append(f"{manager}: {exc}")
                continue

        copied.append({
            "manager": manager,
            "archivePath": str(target),
        })

    primary = Path(copied[0]["archivePath"]) if copied else archive
    copy_error = "; ".join(errors) if errors else None
    return primary, copy_error, copied

def manager_download_targets(config):
    mode = str(config.get("download_routing_mode") or "auto_open_manager").lower()
    both_policy = str(config.get("when_both_managers_open") or "both").lower()

    mo2_open = mo2_is_running(config)
    vortex_open = vortex_is_running(config)

    mo2_text = str(config.get("mo2_downloads_path") or "").strip()
    vortex_text = str(config.get("vortex_downloads_path") or "").strip()

    mo2_path = Path(mo2_text) if mo2_text else None
    vortex_path = Path(vortex_text) if vortex_text else None

    targets = []

    if mode == "manual":
        active = str(config.get("active_downloads_target") or "mo2").lower()
        if active == "vortex" and vortex_path:
            targets.append(("vortex", vortex_path))
        elif mo2_path:
            targets.append(("mo2", mo2_path))
        elif vortex_path:
            targets.append(("vortex", vortex_path))
        return targets

    if mo2_open and vortex_open:
        if both_policy == "vortex":
            if vortex_path:
                targets.append(("vortex", vortex_path))
        elif both_policy == "mo2":
            if mo2_path:
                targets.append(("mo2", mo2_path))
        else:
            if mo2_path:
                targets.append(("mo2", mo2_path))
            if vortex_path:
                targets.append(("vortex", vortex_path))
        return targets

    if vortex_open and vortex_path:
        targets.append(("vortex", vortex_path))
        return targets

    if mo2_open and mo2_path:
        targets.append(("mo2", mo2_path))
        return targets

    active = str(config.get("active_downloads_target") or "mo2").lower()
    if active == "vortex" and vortex_path:
        targets.append(("vortex", vortex_path))
    elif mo2_path:
        targets.append(("mo2", mo2_path))
    elif vortex_path:
        targets.append(("vortex", vortex_path))

    return targets

def status_payload():
    config = load_config()

    mo2_path = Path(str(config.get("mo2_path") or ""))
    mo2_root = mo2_path.parent if mo2_path else Path("")
    plugins_path = mo2_root / "plugins" if mo2_root else Path("")
    plugin_path = plugins_path / "ll_integration" if plugins_path else Path("")

    downloads_text = str(config.get("mo2_downloads_path") or "").strip()
    downloads_path = Path(downloads_text) if downloads_text else Path("")

    metadata_path = Path(str(config.get("metadata_path") or ""))

    active_instance_path = Path(str(config.get("active_mo2_instance_path") or ""))
    active_plugin_path = Path(str(config.get("active_mo2_plugin_path") or ""))
    active_game = str(config.get("active_mo2_game") or "").strip()
    active_synced_at = str(config.get("active_mo2_synced_at") or "").strip()

    vortex_state_candidates = []

    configured_vortex_state = str(config.get("vortex_state_path") or "").strip()
    if configured_vortex_state:
        vortex_state_candidates.append(Path(configured_vortex_state))

    vortex_state_candidates.extend([
        BASE_DIR / "vortex_state.json",
        BASE_DIR / "vortex" / "vortex_state.json",
        BASE_DIR / "vortex_state" / "vortex_state.json",
        Path(os.environ.get("LOCALAPPDATA", "")) / "LLIntegration" / "native-app" / "vortex_state.json",
        Path(os.environ.get("APPDATA", "")) / "Vortex" / "ll_integration" / "vortex_state.json",
    ])

    vortex_state_path = next(
        (p for p in vortex_state_candidates if p and p.exists()),
        vortex_state_candidates[0]
    )
    vortex_state = load_json_object(vortex_state_path, {})

    vortex_downloads_text = str(
        config.get("vortex_downloads_path")
        or vortex_state.get("downloadsPath")
        or vortex_state.get("downloadPath")
        or ""
    ).strip()
    vortex_downloads_path = Path(vortex_downloads_text) if vortex_downloads_text else Path("")

    active_vortex_game = str(
        config.get("active_vortex_game")
        or vortex_state.get("activeGameId")
        or vortex_state.get("gameId")
        or ""
    ).strip()

    active_vortex_profile = str(
        config.get("active_vortex_profile")
        or vortex_state.get("activeProfileName")
        or vortex_state.get("profileName")
        or vortex_state.get("activeProfileId")
        or vortex_state.get("profileId")
        or ""
    ).strip()

    active_vortex_synced_at = str(
        config.get("active_vortex_synced_at")
        or vortex_state.get("capturedAt")
        or vortex_state.get("syncedAt")
        or ""
    ).strip()

    vortex_staging_text = str(
        config.get("vortex_staging_path")
        or config.get("vortex_mods_path")
        or vortex_state.get("stagingPath")
        or vortex_state.get("modsPath")
        or vortex_state.get("installPath")
        or ""
    ).strip()
    vortex_staging_path = Path(vortex_staging_text) if vortex_staging_text else Path("")

    return {
        "ok": True,
        "nativeApp": {
            "baseDir": str(BASE_DIR),
            "configPath": str(CONFIG_FILE),
            "configExists": CONFIG_FILE.exists(),
        },
        "mo2": {
            "path": str(mo2_path),
            "exists": mo2_path.exists(),
            "pluginsPath": str(plugins_path),
            "pluginsPathExists": plugins_path.exists(),
            "llPluginPath": str(plugin_path),
            "llPluginInstalled": plugin_path.exists(),
            "activeInstancePath": str(active_instance_path),
            "activeInstanceExists": active_instance_path.exists(),
            "activePluginPath": str(active_plugin_path),
            "activePluginInstalled": active_plugin_path.exists(),
            "activeGame": active_game,
            "activeSyncedAt": active_synced_at,
        },
        "vortex": {
            "statePath": str(vortex_state_path),
            "stateExists": vortex_state_path.exists(),
            "running": vortex_is_running(config),
            "downloadsPath": str(vortex_downloads_path),
            "downloadsExists": bool(vortex_downloads_text) and vortex_downloads_path.exists(),
            "stagingPath": str(vortex_staging_path),
            "stagingExists": bool(vortex_staging_text) and vortex_staging_path.exists(),
            "activeGame": active_vortex_game,
            "activeProfile": active_vortex_profile,
            "activeSyncedAt": active_vortex_synced_at,
            "modCount": len(vortex_state.get("mods") or []),
            "downloadCount": len(vortex_state.get("downloads") or []),
            "stateCandidates": [str(p) for p in vortex_state_candidates],
            "downloadsParent": str(vortex_downloads_path.parent) if vortex_downloads_text else "",
            "downloadsParentExists": bool(vortex_downloads_text) and vortex_downloads_path.parent.exists(),
            "stagingParent": str(vortex_staging_path.parent) if vortex_staging_text else "",
            "stagingParentExists": bool(vortex_staging_text) and vortex_staging_path.parent.exists(),
        },
        "downloads": {
            "path": str(downloads_path),
            "exists": bool(downloads_text) and downloads_path.exists(),
            "vortexPath": str(vortex_downloads_path),
            "vortexExists": bool(vortex_downloads_text) and vortex_downloads_path.exists(),
            "activeTarget": str(config.get("active_downloads_target") or "mo2"),
            "copyArchivesToMo2Downloads": bool(config.get("copy_archives_to_mo2_downloads", True)),
            "copyArchivesToVortexDownloads": bool(config.get("copy_archives_to_vortex_downloads", True)),
            "overwriteExistingDownloads": bool(config.get("overwrite_existing_downloads", True)),
        },
        "metadata": {
            "path": str(metadata_path),
            "exists": metadata_path.exists(),
        },
        "cookies": {
            "path": str(OUT_FILE),
            "exists": OUT_FILE.exists(),
        },
        "latestDownload": {
            "iniPath": str(LATEST_INI_FILE),
            "exists": LATEST_INI_FILE.exists(),
        },
        "floatingControls": {
            "enabled": bool(config.get("floating_controls_enabled", False)),
            "statePath": str(FLOATING_CONTROLS_STATE_FILE),
            "overlayExists": bool((BASE_DIR / "ll_integration_overlay.exe").exists() or (BASE_DIR / "overlay.py").exists()),
        },
    }

def handle_message(msg):
    if msg.get("action") == "status":
        return status_payload()

    if msg.get("action") == "save_ll_cookies":
        OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
        OUT_FILE.write_text(json.dumps(msg, indent=2), encoding="utf-8")
        return {"ok": True, "savedTo": str(OUT_FILE)}

    if msg.get("action") == "save_ll_download_event":
        events = load_json_list(DOWNLOAD_EVENTS_FILE)
        events.append(msg)
        save_json(DOWNLOAD_EVENTS_FILE, events[-100:])
        write_latest_ini(LATEST_INI_FILE, msg)
        return {
            "ok": True,
            "savedTo": str(DOWNLOAD_EVENTS_FILE),
            "latestIni": str(LATEST_INI_FILE)
        }

    if msg.get("action") == "save_ll_download_completed":
        event = msg.get("event") or {}
        archive_path = msg.get("archivePath")
        if not archive_path:
            return {"ok": False, "error": "Missing archivePath"}

        completions = load_json_list(DOWNLOAD_COMPLETIONS_FILE)
        completions.append(msg)
        save_json(DOWNLOAD_COMPLETIONS_FILE, completions[-100:])

        ini_path, json_path, metadata_ini_path, metadata_json_path, archive, copy_error = write_sidecar_files(
            archive_path,
            event,
            msg.get("browserDownloadUrl"),
            msg.get("completedAt"),
        )
        return {
            "ok": True,
            "savedTo": str(DOWNLOAD_COMPLETIONS_FILE),
            "sidecarIni": str(ini_path),
            "sidecarJson": str(json_path),
            "metadataIni": str(metadata_ini_path),
            "metadataJson": str(metadata_json_path),
            "archivePath": str(archive),
            "copyError": copy_error,
        }

    if msg.get("action") == "open_floating_controls":
        return open_floating_controls()

    if msg.get("action") == "floating_controls_state":
        return {"ok": True, "state": floating_controls_state()}

    if msg.get("action") == "floating_controls_status":
        update = {}
        for key in ("armed", "follow", "label", "visible"):
            if key in msg:
                update[key] = msg.get(key)
        return {"ok": True, "state": write_floating_controls_state(update)}

    return {"ok": False, "error": "Unknown action"}

def main():
    while True:
        msg = read_message()
        if msg is None:
            return

        response = handle_message(msg)
        request_id = msg.get("requestId")
        if request_id is not None and isinstance(response, dict):
            response["requestId"] = request_id
        send_message(response)


if __name__ == "__main__":
    main()
