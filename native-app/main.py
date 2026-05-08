import json
import re
import shutil
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
ARCHIVE_SUFFIXES = {".7z", ".zip", ".rar"}
QUICK_HASH_CHUNK_SIZE = 1024 * 1024
DEFAULT_CONFIG = {
    "mo2_path": "",
    "mo2_downloads_path": "",
    "metadata_path": str(BASE_DIR / "metadata"),
    "copy_archives_to_mo2_downloads": True,
    "overwrite_existing_downloads": True,
}

def read_message():
    raw_len = sys.stdin.buffer.read(4)
    if not raw_len:
        sys.exit(0)

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

def copy_archive_to_mo2_downloads(archive, config):
    if not config.get("copy_archives_to_mo2_downloads", True):
        return archive, None

    downloads_path = Path(str(config.get("mo2_downloads_path") or ""))
    if not downloads_path:
        return archive, None

    downloads_path.mkdir(parents=True, exist_ok=True)
    target = downloads_path / archive.name
    if archive.resolve() == target.resolve():
        return archive, None

    if not config.get("overwrite_existing_downloads", True):
        target = unique_path(target)

    try:
        shutil.copy2(archive, target)
    except OSError as exc:
        return archive, str(exc)

    return target, None

def ll_ini_lines(event, archive_path=None, browser_download_url=None, completed_at=None):
    download = event.get("download") or {}
    archive = Path(archive_path) if archive_path else None
    archive_size = archive.stat().st_size if archive and archive.exists() else ""
    quick_hash = archive_quick_hash(archive) if archive and archive.exists() else ""
    source_type = event.get("sourceType") or ("external" if event.get("action") == "capture_external_archive" else "loverslab")
    is_external = source_type == "external"
    file_name = download.get("name") or (archive.name if archive else "")
    download_url = download.get("url") or browser_download_url or ""
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
        f"version={ini_value(download.get('version'))}",
        f"size={ini_value(download.get('size'))}",
        f"date_iso={ini_value(download.get('date_iso'))}",
        f"captured_at={ini_value(event.get('capturedAt'))}",
        f"archive_path={ini_value(archive_path)}",
        f"browser_download_url={ini_value(browser_download_url)}",
        f"completed_at={ini_value(completed_at)}",
        f"fixed_version={'true' if is_external else 'false'}",
        f"manual_update={'true' if is_external else 'false'}",
        f"skip_update_check={'true' if is_external else 'false'}",
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
    archive, copy_error = copy_archive_to_mo2_downloads(source_archive, config)
    if not is_archive(archive):
        raise ValueError(f"Not a supported archive: {archive}")

    ini_path = archive.with_name(f"{archive.name}.ll.ini")
    json_path = archive.with_name(f"{archive.name}.ll.json")
    metadata_downloads = Path(str(config.get("metadata_path"))) / "downloads"
    metadata_ini_path = metadata_downloads / f"{archive.name}.ll.ini"
    metadata_json_path = metadata_downloads / f"{archive.name}.ll.json"
    payload = {
        "sourceType": event.get("sourceType") or ("external" if event.get("action") == "capture_external_archive" else "loverslab"),
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
    }

    ini_text = "\n".join(ll_ini_lines(event, archive, browser_download_url, completed_at))
    ini_path.write_text(ini_text, encoding="utf-8")
    save_json(json_path, payload)
    metadata_ini_path.parent.mkdir(parents=True, exist_ok=True)
    metadata_ini_path.write_text(ini_text, encoding="utf-8")
    save_json(metadata_json_path, payload)
    return ini_path, json_path, metadata_ini_path, metadata_json_path, archive, copy_error

def status_payload():
    config = load_config()
    mo2_path = Path(str(config.get("mo2_path") or ""))
    mo2_root = mo2_path.parent if mo2_path else Path("")
    plugins_path = mo2_root / "plugins" if mo2_root else Path("")
    plugin_path = plugins_path / "ll_integration" if plugins_path else Path("")
    downloads_path = Path(str(config.get("mo2_downloads_path") or ""))
    metadata_path = Path(str(config.get("metadata_path") or ""))

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
        },
        "downloads": {
            "path": str(downloads_path),
            "exists": downloads_path.exists(),
            "copyArchivesToMo2Downloads": bool(config.get("copy_archives_to_mo2_downloads", True)),
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

    return {"ok": False, "error": "Unknown action"}

def main():
    msg = read_message()
    send_message(handle_message(msg))

if __name__ == "__main__":
    main()
