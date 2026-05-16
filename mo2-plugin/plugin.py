import configparser
import html as html_lib
import json
import queue
import re
import shutil
import threading
import time
from pathlib import Path
from typing import Sequence
from urllib.parse import parse_qs, quote_plus, urljoin, urlparse
from urllib.request import Request, urlopen
import webbrowser

import mobase
from PyQt6.QtCore import QObject, Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QAction, QColor, QPixmap
from PyQt6.QtGui import QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QGridLayout,
    QHBoxLayout,
    QAbstractItemView,
    QInputDialog,
    QLabel,
    QListWidget,
    QMessageBox,
    QDoubleSpinBox,
    QProgressBar,
    QPushButton,
    QHeaderView,
    QLineEdit,
    QMenu,
    QSplitter,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QTextEdit,
    QVBoxLayout,
)

from .check_update import DEFAULT_COOKIES, check_ini_for_updates, choose_latest, with_query_value
from .utils import archive_quick_hash, compare_versions, cookie_header, extract_downloads, fetch_ll_html, load_ll_cookies


PLUGIN_NAME = "LL Integration"
ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_LATEST_INI = ROOT_DIR / "native-app" / "downloads_storage" / "latest_ll_download.ini"
PLUGIN_PATHS_FILE = Path(__file__).resolve().parent / "plugin_paths.json"
AUTO_BIND_WINDOW_SECONDS = 15 * 60
UPDATE_REQUEST_DELAY_SECONDS = 0.8
UPDATE_BATCH_SIZE = 25
UPDATE_BATCH_PAUSE_SECONDS = 0.0
UPDATE_REQUEST_TIMEOUT_SECONDS = 15.0
UPDATE_CACHE_VERSION = 1
LL_SECTION = "LoversLab"
MOD_META_FILE = "meta.ini"
LEGACY_MOD_LL_FILE = "LL.ini"

DOWNLOAD_CHUNK_SIZE = 1024 * 1024
UPDATE_MODE_MANUAL = "manual"
UPDATE_MODE_DOWNLOAD_ONLY = "download_only"
UPDATE_MODE_ASSISTED = "assisted"
UPDATE_MODE_AUTOMATIC = "automatic"
UPDATE_MODE_SKIP = "skip"
UPDATE_MODE_OPTIONS = [
    (UPDATE_MODE_MANUAL, "Manual install", True),
    (UPDATE_MODE_DOWNLOAD_ONLY, "Download only", True),
    (UPDATE_MODE_ASSISTED, "Assisted install", True),
    (UPDATE_MODE_AUTOMATIC, "Automatic install (experimental, coming later)", False),
    (UPDATE_MODE_SKIP, "Skip updates", True),
]
UPDATE_MODE_LABELS = {value: label for value, label, _enabled in UPDATE_MODE_OPTIONS}
VOICE_KEYWORDS = (
    "voice",
    "voices",
    "voiced",
    "dbvo",
    "dvo",
    "IVDT",
    "ivdt",
    "dvit",
    "dialogue voice",
    "dialogue voices",
    "voice pack",
    "voice files",
    "voicefiles",
    "voice addon",
    "voice add-on",
    "silent voice",
)

VOICE_NOISE_WORDS_RE = re.compile(
    r"""
    \b(
        dbvo
        | dvo
        | IVDT
        | ivdt
        | dvit
        | voice
        | voices
        | voiced
        | voicepack
        | voicefiles
        | voice[\s_-]*files
        | voice[\s_-]*pack
        | voice[\s_-]*addon
        | add[\s_-]*on
        | addon
        | dialogue[\s_-]*voice
        | dialogue[\s_-]*voices
        | silent[\s_-]*voice
        | patch
        | se
        | ae
        | le
    )\b
    """,
    re.IGNORECASE | re.VERBOSE,
)

VOICE_CATEGORIES = (
    ("npc", "Voicepack"),
    ("player", "DBVO"),
    ("scene", "IVDT / IVDT"),
)
VOICE_CATEGORY_LABELS = dict(VOICE_CATEGORIES)
VOICE_CATEGORY_IDS = {label: key for key, label in VOICE_CATEGORIES}
DEFAULT_VOICE_MATCH_THRESHOLD = 55

NEXUS_SOURCE_RE = re.compile(
    r"^https?://(?:www\.)?nexusmods\.com/(?P<game>[^/\s?#]+)/mods/(?P<mod_id>\d+)",
    re.IGNORECASE,
)

def normalized_voice_name(value: str) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bv?\d+(?:[._-]\d+){1,4}\b", " ", text, flags=re.IGNORECASE)
    text = VOICE_NOISE_WORDS_RE.sub(" ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(token for token in text.split() if len(token) > 1)


def voice_keyword_present(value: str) -> bool:
    text = re.sub(r"[^a-z0-9]+", " ", str(value or "").lower())
    padded = f" {text} "
    return any(f" {keyword} " in padded for keyword in VOICE_KEYWORDS)


def voice_category_guess(value: str) -> str:
    text = re.sub(r"[^a-z0-9]+", " ", str(value or "").lower())
    padded = f" {text} "
    if any(token in padded for token in (" dbvo ", " dvo ", " dragonborn voice over ")):
        return "player"
    if any(token in padded for token in (" IVDT ", " ivdt ", " dvit ", " dirty talk ", " scene ", " addon ")):
        return "scene"
    if any(token in padded for token in (" voicepack ", " voice pack ", " voicefiles ", " voice files ", " npc ", " dialogue ", " dialogues ")):
        return "npc"
    if voice_keyword_present(value):
        return "npc"
    return ""

def voice_match_score(base_name: str, voice_name: str) -> int:
    base = normalized_voice_name(base_name)
    voice = normalized_voice_name(voice_name)

    if not base or not voice:
        return 0

    score = 0

    if base == voice:
        score += 120
    elif base in voice or voice in base:
        score += 85

    compact_base = base.replace(" ", "")
    compact_voice = voice.replace(" ", "")
    if compact_base and compact_voice and compact_base != compact_voice:
        if compact_base in compact_voice or compact_voice in compact_base:
            score += 65

    base_tokens = set(base.split())
    voice_tokens = set(voice.split())
    common = base_tokens & voice_tokens

    if base_tokens:
        score += int((len(common) / len(base_tokens)) * 70)

    if voice_keyword_present(voice_name):
        score += 25

    return min(score, 160)


def normalize_voice_source_url(value: str) -> str:
    url = str(value or "").strip()
    if not url:
        return ""
    if "loverslab.com/files/file/" in url.lower():
        return with_query_value(url, "do", "download")
    match = NEXUS_SOURCE_RE.match(url)
    if match:
        return f"https://www.nexusmods.com/{match.group('game')}/mods/{match.group('mod_id')}?tab=files"
    return url


def is_nexus_source_url(value: str) -> bool:
    return bool(NEXUS_SOURCE_RE.match(str(value or "").strip()))


def nexus_source_parts(value: str) -> tuple[str, str] | None:
    match = NEXUS_SOURCE_RE.match(str(value or "").strip())
    if not match:
        return None
    return match.group("game"), match.group("mod_id")


def validate_nexus_api_key(api_key: str, timeout: float = 30.0) -> dict:
    request = Request(
        "https://api.nexusmods.com/v1/users/validate.json",
        headers={
            "apikey": api_key,
            "User-Agent": "LL Integration",
            "Accept": "application/json",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))
    return payload if isinstance(payload, dict) else {}


def fetch_nexus_files_api(game: str, mod_id: str, api_key: str, timeout: float = 30.0) -> list[dict]:
    url = f"https://api.nexusmods.com/v1/games/{game}/mods/{mod_id}/files.json"
    request = Request(
        url,
        headers={
            "apikey": api_key,
            "User-Agent": "LL Integration",
            "Accept": "application/json",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        payload = json.loads(response.read().decode("utf-8"))
    files = payload.get("files", payload if isinstance(payload, list) else [])
    if not isinstance(files, list):
        return []
    source_url = f"https://www.nexusmods.com/{game}/mods/{mod_id}?tab=files"
    downloads: list[dict] = []
    for item in files:
        if not isinstance(item, dict):
            continue
        file_id = str(item.get("file_id") or "").strip()
        name = str(item.get("name") or item.get("file_name") or "").strip()
        if not file_id or not name:
            continue
        size_kb = item.get("size_kb")
        size = f"{size_kb} KB" if size_kb not in (None, "") else ""
        downloads.append({
            "source_url": source_url,
            "source_title": "Nexus Mods API source",
            "download_name": name,
            "download_url": f"https://www.nexusmods.com/{game}/mods/{mod_id}?tab=files&file_id={file_id}&nmm=1",
            "voice_category": voice_category_guess(name),
            "size": size,
            "date_iso": str(item.get("uploaded_time") or ""),
            "version": str(item.get("version") or item.get("mod_version") or ""),
            "source_type": "nexus",
            "nexus_file_id": file_id,
            "nexus_category": str(item.get("category_name") or ""),
        })
    return downloads


def fetch_nexus_html(url: str, timeout: float = 30.0) -> str:
    request = Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) "
                "Gecko/20100101 Firefox/125.0"
            ),
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Encoding": "identity",
            "Accept-Language": "en-US,en;q=0.9",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        return response.read().decode(charset, errors="replace")


def strip_html_text(value: str) -> str:
    text = re.sub(r"<script\b.*?</script>", " ", value or "", flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<style\b.*?</style>", " ", text, flags=re.IGNORECASE | re.DOTALL)
    text = re.sub(r"<[^>]+>", " ", text)
    return " ".join(html_lib.unescape(text).split())


def nexus_file_name_from_text(text: str, file_id: str) -> str:
    cleaned = re.sub(r"\s+", " ", text or "").strip()
    patterns = [
        r"([^|]{4,180}?)(?:\s+Date uploaded|\s+File size|\s+Unique DLs|\s+Total DLs|\s+Version|\s+Mod manager download|\s+Manual download)",
        r"([A-Za-z0-9][^|]{4,180}?)\s+\(?\d+(?:\.\d+)?\s*(?:KB|MB|GB)\)?",
    ]
    for pattern in patterns:
        match = re.search(pattern, cleaned, re.IGNORECASE)
        if not match:
            continue
        name = match.group(1).strip(" -|")
        name = re.sub(r"^(Main files|Optional files|Miscellaneous files)\s+", "", name, flags=re.IGNORECASE)
        if name and file_id not in name:
            return name[:180]
    return ""


def extract_nexus_downloads(html: str, source_url: str) -> list[dict]:
    parsed = urlparse(source_url)
    base_url = f"{parsed.scheme}://{parsed.netloc}{parsed.path}"
    seen: set[str] = set()
    downloads: list[dict] = []

    for match in re.finditer(r"(?:file_id=|data-file-id=[\"']?)(\d+)", html or "", re.IGNORECASE):
        file_id = match.group(1)
        if file_id in seen:
            continue
        seen.add(file_id)
        chunk = html[max(0, match.start() - 2200): min(len(html), match.end() + 2200)]
        text = strip_html_text(chunk)
        name = nexus_file_name_from_text(text, file_id) or f"Nexus file {file_id}"
        size_match = re.search(r"\b\d+(?:\.\d+)?\s*(?:KB|MB|GB)\b", text, re.IGNORECASE)
        date_match = re.search(
            r"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}(?:,\s*\d{1,2}:\d{2}\s*(?:AM|PM)?)?",
            text,
        )
        downloads.append({
            "source_url": source_url,
            "source_title": "Nexus Mods source",
            "download_name": name,
            "download_url": f"{base_url}?tab=files&file_id={file_id}&nmm=1",
            "voice_category": voice_category_guess(name),
            "size": size_match.group(0) if size_match else "",
            "date_iso": date_match.group(0) if date_match else "",
            "version": "",
            "source_type": "nexus",
            "nexus_file_id": file_id,
        })

    return downloads


def voice_search_query(base_name: str) -> str:
    clean = normalized_voice_name(base_name)
    return f"{clean} voice OR voices OR DBVO OR IVDT".strip()

def normalized_update_mode(value: str | None, fixed: bool = False) -> str:
    mode = str(value or "").strip().lower()
    valid = {item[0] for item in UPDATE_MODE_OPTIONS}
    if mode in valid:
        return mode
    return UPDATE_MODE_SKIP if fixed else UPDATE_MODE_MANUAL


def update_mode_label(mode: str) -> str:
    return UPDATE_MODE_LABELS.get(normalized_update_mode(mode), UPDATE_MODE_LABELS[UPDATE_MODE_MANUAL])


def configure_update_mode_combo(combo: QComboBox, selected: str) -> None:
    selected_mode = normalized_update_mode(selected)
    selected_index = 0
    for index, (value, label, enabled) in enumerate(UPDATE_MODE_OPTIONS):
        combo.addItem(label, value)
        if value == selected_mode:
            selected_index = index
        item = combo.model().item(index)
        if item is not None and not enabled:
            item.setEnabled(False)
            item.setToolTip("Placeholder for a future release.")
    combo.setCurrentIndex(selected_index)


def ini_value(value) -> str:
    return "" if value is None else str(value).replace("\n", " ").replace("\r", " ").strip()


def ll_resource_id(download_url: str) -> str:
    if not download_url:
        return ""
    query = parse_qs(urlparse(download_url).query)
    return (query.get("r") or [""])[0]


def safe_archive_name(name: str) -> str:
    cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", str(name or "").strip())
    return cleaned.strip(" .") or "loverslab-download.archive"


def download_loverslab_archive(url: str, target: Path, cookies_path: Path, referer: str, timeout: float) -> Path:
    cookies = load_ll_cookies(cookies_path, required_only=False)
    if not cookies:
        raise RuntimeError(f"No usable LoversLab cookies found in {cookies_path}")

    target.parent.mkdir(parents=True, exist_ok=True)
    headers = {
        "Cookie": cookie_header(cookies),
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) "
            "Gecko/20100101 Firefox/125.0"
        ),
        "Accept": "application/octet-stream,text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "identity",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "DNT": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    if referer:
        headers["Referer"] = referer

    temp = target.with_name(f"{target.name}.part")
    request = Request(url, headers=headers)
    try:
        with urlopen(request, timeout=timeout) as response, temp.open("wb") as file:
            while True:
                chunk = response.read(DOWNLOAD_CHUNK_SIZE)
                if not chunk:
                    break
                file.write(chunk)
        temp.replace(target)
    except Exception:
        try:
            temp.unlink(missing_ok=True)
        except Exception:
            pass
        raise
    return target


def write_update_download_sidecar(
    ini_path: Path,
    archive_path: Path,
    latest: dict,
    download_url: str,
) -> Path:
    source = read_ll_section(ini_path)
    page_url = source.get("page_url", "").strip()
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    archive_hash = archive_quick_hash(archive_path) if archive_path.exists() else ""
    archive_size = str(archive_path.stat().st_size) if archive_path.exists() else ""
    ll_id_match = re.search(r"/files/file/(\d+)", page_url)
    ll_id = ll_id_match.group(1) if ll_id_match else source.get("ll_file_id", "").strip()
    update_mode = normalized_update_mode(source.get("update_mode"))
    lines = [
        "[LoversLab]",
        "source=loverslab",
        f"ll_file_id={ini_value(ll_id)}",
        f"ll_resource_id={ini_value(ll_resource_id(download_url))}",
        f"page_url={ini_value(page_url)}",
        f"page_title={ini_value(source.get('page_title', ''))}",
        f"download_url={ini_value(download_url)}",
        f"file_name={ini_value(latest.get('name'))}",
        f"original_archive_name={ini_value(latest.get('name'))}",
        f"archive_name={ini_value(archive_path.name)}",
        f"archive_size_bytes={archive_size}",
        f"archive_quick_hash={archive_hash}",
        f"version={ini_value(latest.get('version'))}",
        f"size={ini_value(latest.get('size'))}",
        f"date_iso={ini_value(latest.get('date_iso'))}",
        f"captured_at={now}",
        f"archive_path={ini_value(archive_path)}",
        f"browser_download_url={ini_value(download_url)}",
        f"completed_at={now}",
        f"update_mode={update_mode}",
        "fixed_version=false",
        "manual_update=false",
        "skip_update_check=false",
        f"multipart={ini_value(source.get('multipart', 'false'))}",
        f"file_pattern={ini_value(source.get('file_pattern', latest.get('name') or ''))}",
        "",
    ]
    sidecar = Path(f"{archive_path}.ll.ini")
    sidecar.write_text("\n".join(lines), encoding="utf-8")
    return sidecar


def write_voice_download_sidecar(
    archive_path: Path,
    candidate: dict,
    download_url: str,
) -> Path:
    now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
    archive_hash = archive_quick_hash(archive_path) if archive_path.exists() else ""
    archive_size = str(archive_path.stat().st_size) if archive_path.exists() else ""
    page_url = str(candidate.get("source_url") or candidate.get("page_url") or "").strip()
    download_name = str(candidate.get("download_name") or archive_path.name).strip()
    lines = [
        "[LoversLab]",
        "source=loverslab",
        f"ll_file_id={ini_value(ll_file_id_from_url(page_url))}",
        f"ll_resource_id={ini_value(ll_resource_id(download_url))}",
        f"page_url={ini_value(page_url)}",
        f"page_title={ini_value(candidate.get('source_title') or 'Voice pack source')}",
        f"download_url={ini_value(download_url)}",
        f"file_name={ini_value(download_name)}",
        f"original_archive_name={ini_value(download_name)}",
        f"archive_name={ini_value(archive_path.name)}",
        f"archive_size_bytes={archive_size}",
        f"archive_quick_hash={archive_hash}",
        f"version={ini_value(candidate.get('version') or '')}",
        f"size={ini_value(candidate.get('size') or '')}",
        f"date_iso={ini_value(candidate.get('date_iso') or '')}",
        f"captured_at={now}",
        f"archive_path={ini_value(archive_path)}",
        f"browser_download_url={ini_value(download_url)}",
        f"completed_at={now}",
        "update_mode=manual",
        "fixed_version=true",
        "manual_update=true",
        "skip_update_check=true",
        "manual_install=true",
        "multipart=false",
        f"file_pattern={ini_value(download_name)}",
        f"voice_base_mod={ini_value(candidate.get('base_mod') or '')}",
        f"voice_category={ini_value(candidate.get('voice_category') or '')}",
        f"voice_category_label={ini_value(candidate.get('voice_category_label') or '')}",
        f"voice_match_score={ini_value(candidate.get('online_score') or candidate.get('score') or '')}",
        "",
    ]
    sidecar = Path(f"{archive_path}.ll.ini")
    sidecar.write_text("\n".join(lines), encoding="utf-8")
    Path(f"{archive_path}.ll.json").write_text(
        json.dumps(
            {
                "sourceType": "loverslab",
                "voiceCandidate": candidate,
                "archiveName": archive_path.name,
                "archivePath": str(archive_path),
                "browserDownloadUrl": download_url,
                "completedAt": now,
            },
            indent=2,
        ),
        encoding="utf-8",
    )
    return sidecar


def read_mod_meta_general(mod) -> dict[str, str]:
    meta_path = mod_meta_path(mod)
    if not meta_path.exists():
        return {}

    config = configparser.ConfigParser(interpolation=None)
    config.read(meta_path, encoding="utf-8")
    if "General" not in config:
        return {}

    return {str(key).lower(): str(value).strip() for key, value in config["General"].items()}


def _positive_int(value: str | None) -> bool:
    if not value:
        return False

    try:
        return int(str(value).strip()) > 0
    except ValueError:
        return False


def mod_has_nexus_identity(mod) -> tuple[bool, str]:
    general = read_mod_meta_general(mod)
    url = general.get("url", "").lower()
    mod_id = general.get("modid") or general.get("mod_id") or general.get("nexusid")

    if "nexusmods.com" in url or url.startswith("nxm://"):
        return True, "meta.ini URL points to Nexus"
    if _positive_int(mod_id):
        return True, f"meta.ini has Nexus mod id {mod_id}"

    return False, ""


def mod_has_purgeable_nexus_identity(mod) -> tuple[bool, str]:
    general = read_mod_meta_general(mod)
    url = general.get("url", "").lower()
    repository = general.get("repository", "").lower()
    mod_id = general.get("modid") or general.get("mod_id") or general.get("nexusid")
    has_nexus_markers = _positive_int(mod_id) and (
        repository == "nexus"
        or any(
            general.get(key)
            for key in ("lastnexusquery", "lastnexusupdate", "nexuslastmodified", "nexusfilestatus")
        )
    )

    if "loverslab.com" in url and has_nexus_markers:
        return True, f"meta.ini has LoversLab URL but Nexus mod id {mod_id}"
    if has_nexus_markers:
        return True, f"meta.ini has Nexus mod id {mod_id}"

    return False, ""


def truthy(value: str | None) -> bool:
    return str(value or "").strip().lower() in {"1", "true", "yes", "on"}


def _section_key(section: configparser.SectionProxy, name: str) -> str:
    wanted = name.lower()
    for key in section.keys():
        if str(key).lower() == wanted:
            return str(key)
    return name


def cleanup_loverslab_meta(mod) -> bool:
    meta_path = mod_meta_path(mod)
    if not meta_path.exists():
        return False

    config = configparser.ConfigParser(interpolation=None)
    config.optionxform = str
    config.read(meta_path, encoding="utf-8")
    if "General" not in config:
        return False

    general = config["General"]
    changed = False
    if LL_SECTION in config:
        config.remove_section(LL_SECTION)
        changed = True

    url_key = _section_key(general, "url")
    url = general.get(url_key, "")
    if "loverslab.com" in url.lower():
        general[url_key] = ""
        changed = True

    if changed:
        general[_section_key(general, "hasCustomURL")] = "false"
        repo_key = _section_key(general, "repository")
        mod_id = general.get(_section_key(general, "modid"), "")
        if _positive_int(mod_id):
            general[repo_key] = "Nexus"

        with meta_path.open("w", encoding="utf-8") as file:
            config.write(file, space_around_delimiters=False)

    return changed


def unique_backup_path(path: Path, suffix: str) -> Path:
    backup = path.with_name(f"{path.name}{suffix}")
    if not backup.exists():
        return backup

    for index in range(2, 1000):
        candidate = path.with_name(f"{path.name}.purged-{index}.bak")
        if not candidate.exists():
            return candidate

    return path.with_name(f"{path.name}.purged-{int(time.time())}.bak")


def mod_root_path(mod) -> Path:
    return Path(str(mod.absolutePath()))


def mod_meta_path(mod) -> Path:
    return mod_root_path(mod) / MOD_META_FILE


def legacy_mod_ll_path(mod) -> Path:
    return mod_root_path(mod) / LEGACY_MOD_LL_FILE


def read_ini_file(path: Path, preserve_case: bool = False) -> configparser.ConfigParser:
    config = configparser.ConfigParser(interpolation=None)
    if preserve_case:
        config.optionxform = str
    config.read(path, encoding="utf-8")
    return config


def has_ll_section(path: Path) -> bool:
    if not path.exists():
        return False
    return LL_SECTION in read_ini_file(path)


def read_ll_section(path: Path) -> configparser.SectionProxy:
    config = read_ini_file(path)
    if LL_SECTION not in config:
        raise RuntimeError(f"{path} is missing a [{LL_SECTION}] section")
    return config[LL_SECTION]


def ll_file_id_from_url(url: str) -> str:
    match = re.search(r"/files/file/(\d+)", str(url or ""))
    return match.group(1) if match else ""


def ll_metadata_identity(section: configparser.SectionProxy) -> str:
    file_id = section.get("ll_file_id", "").strip()
    if file_id:
        return f"file:{file_id}"

    for key in ("page_url", "download_url", "browser_download_url"):
        file_id = ll_file_id_from_url(section.get(key, ""))
        if file_id:
            return f"file:{file_id}"

    page_url = section.get("page_url", "").strip().lower()
    if page_url:
        parsed = urlparse(page_url)
        normalized = parsed._replace(query="", fragment="").geturl().rstrip("/")
        return f"url:{normalized}"

    return ""


def ll_metadata_same_source(left_path: Path, right_path: Path) -> bool:
    try:
        left = read_ll_section(left_path)
        right = read_ll_section(right_path)
    except Exception:
        return False

    left_identity = ll_metadata_identity(left)
    right_identity = ll_metadata_identity(right)
    return bool(left_identity and right_identity and left_identity == right_identity)


def write_mod_ll_metadata_from_file(mod, source_path: Path) -> Path:
    source = read_ini_file(source_path)
    if LL_SECTION not in source:
        raise RuntimeError(f"{source_path} is missing a [{LL_SECTION}] section")
    return write_mod_ll_metadata_from_section(mod, source[LL_SECTION])


def write_mod_ll_metadata_from_text(mod, text: str) -> Path:
    source = configparser.ConfigParser(interpolation=None)
    source.read_string(text)
    if LL_SECTION not in source:
        raise RuntimeError(f"Generated metadata is missing a [{LL_SECTION}] section")
    return write_mod_ll_metadata_from_section(mod, source[LL_SECTION])


def write_mod_ll_metadata_from_section(mod, section: configparser.SectionProxy) -> Path:
    meta_path = mod_meta_path(mod)
    config = read_ini_file(meta_path, preserve_case=True)
    if "General" not in config:
        config["General"] = {}

    if LL_SECTION in config:
        config.remove_section(LL_SECTION)
    config.add_section(LL_SECTION)
    for key, value in section.items():
        config[LL_SECTION][str(key)] = str(value)

    with meta_path.open("w", encoding="utf-8") as file:
        config.write(file, space_around_delimiters=False)

    legacy = legacy_mod_ll_path(mod)
    if legacy.exists():
        legacy.unlink()

    return meta_path


def write_mod_general_source_metadata(mod, page_url: str, version: str) -> Path:
    meta_path = mod_meta_path(mod)
    config = read_ini_file(meta_path, preserve_case=True)
    if "General" not in config:
        config["General"] = {}

    general = config["General"]
    if page_url:
        general["url"] = page_url
        general["hasCustomURL"] = "true"
        general["repository"] = "LoversLab"
    if version:
        general["version"] = f"{version}.0" if version.count(".") == 2 else version

    with meta_path.open("w", encoding="utf-8") as file:
        config.write(file, space_around_delimiters=False)

    return meta_path


def mod_ll_metadata_path(mod, migrate_legacy: bool = True) -> Path | None:
    meta_path = mod_meta_path(mod)
    legacy_path = legacy_mod_ll_path(mod)

    if has_ll_section(meta_path):
        if migrate_legacy and legacy_path.exists():
            legacy_path.unlink()
        return meta_path

    if legacy_path.exists() and has_ll_section(legacy_path):
        if migrate_legacy:
            return write_mod_ll_metadata_from_file(mod, legacy_path)
        return legacy_path

    return None


def remove_mod_ll_metadata(mod) -> list[str]:
    actions = []
    meta_path = mod_meta_path(mod)
    if meta_path.exists():
        config = read_ini_file(meta_path, preserve_case=True)
        if LL_SECTION in config:
            config.remove_section(LL_SECTION)
            with meta_path.open("w", encoding="utf-8") as file:
                config.write(file, space_around_delimiters=False)
            actions.append(f"removed {MOD_META_FILE} [{LL_SECTION}]")

    legacy_path = legacy_mod_ll_path(mod)
    if legacy_path.exists():
        legacy_path.unlink()
        actions.append(f"removed legacy {LEGACY_MOD_LL_FILE}")

    return actions


class CheckAllWorker(QObject):
    rowReady = pyqtSignal(object)
    progressChanged = pyqtSignal(int, int)
    statusChanged = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(
        self,
        jobs: list[dict],
        cookies_path: Path,
        log_path: Path,
        request_delay: float = UPDATE_REQUEST_DELAY_SECONDS,
        batch_size: int = UPDATE_BATCH_SIZE,
        batch_pause: float = UPDATE_BATCH_PAUSE_SECONDS,
        request_timeout: float = UPDATE_REQUEST_TIMEOUT_SECONDS,
    ) -> None:
        super().__init__()
        self._jobs = jobs
        self._cookies_path = cookies_path
        self._log_path = log_path
        self._request_delay = request_delay
        self._batch_size = batch_size
        self._batch_pause = batch_pause
        self._request_timeout = request_timeout
        self._cancelled = False
        self._run_id = time.strftime("%Y%m%d-%H%M%S")

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        total = len(self._jobs)
        network_total = sum(1 for job in self._jobs if self._is_network_job(job))
        network_count = 0
        self._log_event({"event": "run_started", "run_id": self._run_id, "jobs": total, "network_jobs": network_total})

        for index, job in enumerate(self._jobs):
            if self._cancelled:
                break

            is_network_job = self._is_network_job(job)
            if is_network_job:
                if network_count > 0 and self._request_delay > 0:
                    self.statusChanged.emit(f"Waiting {self._request_delay:.1f}s before next request")
                    if not self._sleep_cancelable(self._request_delay):
                        break

                network_count += 1
                self.statusChanged.emit(f"Fetching {network_count} / {network_total}: {job.get('mod') or ''}")

            result = self._check_job(job)
            result["row_index"] = job.get("row_index", index)
            self.rowReady.emit(result)
            self.progressChanged.emit(index + 1, total)

            if (
                is_network_job
                and self._batch_size > 0
                and self._batch_pause > 0
                and network_count < network_total
                and network_count % self._batch_size == 0
            ):
                self.statusChanged.emit(
                    f"Cooldown {self._batch_pause:.0f}s after {network_count} LoversLab requests"
                )
                if not self._sleep_cancelable(self._batch_pause):
                    break

        self._log_event({"event": "run_finished", "run_id": self._run_id, "cancelled": self._cancelled})
        self.finished.emit(self._cancelled)

    def _is_network_job(self, job: dict) -> bool:
        return (
            Path(str(job.get("ini_path") or "")).exists()
            and not job.get("fixed")
            and bool(job.get("page_url"))
        )

    def _sleep_cancelable(self, seconds: float) -> bool:
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            if self._cancelled:
                return False
            time.sleep(min(0.25, max(deadline - time.monotonic(), 0)))
        return not self._cancelled

    def _log_event(self, payload: dict) -> None:
        try:
            self._log_path.parent.mkdir(parents=True, exist_ok=True)
            payload.setdefault("ts", time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()))
            with self._log_path.open("a", encoding="utf-8") as file:
                file.write(json.dumps(payload, ensure_ascii=False) + "\n")
        except Exception:
            pass

    def _check_job(self, job: dict) -> dict:
        if not Path(str(job.get("ini_path") or "")).exists():
            return {
                "ini_path": job.get("ini_path") or "",
                "internal_name": job.get("internal_name") or "",
                "mod": job["mod"],
                "status": "Purged",
                "current": "",
                "latest": "",
                "file": job.get("file") or "",
                "page_url": job.get("page_url") or "",
                "fixed": bool(job.get("fixed")),
                "update_mode": job.get("update_mode") or "",
                "info": "LL Integration metadata was removed",
            }

        if job.get("fixed"):
            return {
                "ini_path": job.get("ini_path") or "",
                "internal_name": job.get("internal_name") or "",
                "mod": job["mod"],
                "status": "Manual",
                "current": job.get("current") or "",
                "latest": "",
                "file": job.get("file") or "",
                "page_url": job.get("page_url") or "",
                "fixed": bool(job.get("fixed")),
                "update_mode": job.get("update_mode") or "",
                "info": "Skip updates; update fetch skipped",
            }

        return self._check_job_cancelable(job)

    def _check_job_cancelable(self, job: dict) -> dict:
        result_queue: queue.Queue[dict] = queue.Queue(maxsize=1)

        def run_fetch() -> None:
            result_queue.put(self._check_job_fetch(job))

        thread = threading.Thread(target=run_fetch, daemon=True)
        thread.start()
        while thread.is_alive():
            if self._cancelled:
                return self._cancelled_result(job)
            thread.join(0.1)

        try:
            return result_queue.get_nowait()
        except queue.Empty:
            return self._cancelled_result(job)

    def _check_job_fetch(self, job: dict) -> dict:
        try:
            started = time.monotonic()
            result = check_ini_for_updates(Path(job["ini_path"]), self._cookies_path, timeout=self._request_timeout)
            duration = time.monotonic() - started
            latest = result.get("latest") or {}
            current = result.get("currentVersion") or ""
            self._log_event({
                "event": "request",
                "run_id": self._run_id,
                "mod": job.get("mod") or "",
                "page_url": job.get("page_url") or "",
                "status": "ok",
                "duration_s": round(duration, 3),
                "timeout_s": round(self._request_timeout, 3),
                "downloads_seen": len(result.get("downloadsSeen") or []),
                "latest": latest.get("version") or "",
            })
            return {
                "ini_path": job.get("ini_path") or "",
                "internal_name": job.get("internal_name") or "",
                "mod": job["mod"],
                "status": "Unknown" if not current else ("Update" if result.get("updateAvailable") else "OK"),
                "current": current,
                "latest": latest.get("version") or "",
                "file": latest.get("name") or result.get("knownFile") or "",
                "latest_url": latest.get("url") or "",
                "latest_size": latest.get("size") or "",
                "latest_date_iso": latest.get("date_iso") or "",
                "page_url": job.get("page_url") or "",
                "fixed": bool(job.get("fixed")),
                "update_mode": job.get("update_mode") or "",
                "info": (
                    f"Fetched in {duration:.1f}s; current version missing"
                    if not current
                    else f"Fetched in {duration:.1f}s"
                ),
            }
        except Exception as exc:
            duration = time.monotonic() - started if "started" in locals() else 0.0
            timed_out = isinstance(exc, TimeoutError) or "timed out" in str(exc).lower()
            info = f"Timed out after {self._request_timeout:.1f}s; skipped" if timed_out else str(exc)
            self._log_event({
                "event": "request",
                "run_id": self._run_id,
                "mod": job.get("mod") or "",
                "page_url": job.get("page_url") or "",
                "status": "timeout" if timed_out else "error",
                "duration_s": round(duration, 3),
                "timeout_s": round(self._request_timeout, 3),
                "error": str(exc),
            })
            return {
                "ini_path": job.get("ini_path") or "",
                "internal_name": job.get("internal_name") or "",
                "mod": job["mod"],
                "status": "Skipped" if timed_out else "Error",
                "current": job.get("current") or "",
                "latest": "",
                "file": job.get("file") or "",
                "page_url": job.get("page_url") or "",
                "fixed": bool(job.get("fixed")),
                "update_mode": job.get("update_mode") or "",
                "info": info,
            }

    def _cancelled_result(self, job: dict) -> dict:
        return {
            "ini_path": job.get("ini_path") or "",
            "internal_name": job.get("internal_name") or "",
            "mod": job["mod"],
            "status": "Queued",
            "current": job.get("current") or "",
            "latest": "",
            "file": job.get("file") or "",
            "page_url": job.get("page_url") or "",
            "fixed": bool(job.get("fixed")),
            "update_mode": job.get("update_mode") or "",
            "info": "Cancelled; abandoned current HTTP request in background",
        }


class TryUpdateWorker(QObject):
    rowReady = pyqtSignal(object)
    statusChanged = pyqtSignal(str)
    finished = pyqtSignal(bool)

    def __init__(
        self,
        row: dict,
        cookies_path: Path,
        downloads_path: Path,
        request_timeout: float = UPDATE_REQUEST_TIMEOUT_SECONDS,
    ) -> None:
        super().__init__()
        self._row = dict(row)
        self._cookies_path = cookies_path
        self._downloads_path = downloads_path
        self._request_timeout = request_timeout

    def run(self) -> None:
        row = dict(self._row)
        try:
            ini_path = Path(str(row.get("ini_path") or ""))
            if not ini_path.exists():
                raise RuntimeError(f"LL Integration metadata was not found:\n{ini_path}")
            if not self._downloads_path:
                raise RuntimeError("MO2 downloads path is not available.")

            self.statusChanged.emit(f"Finding latest archive: {row.get('mod') or ''}")
            result = check_ini_for_updates(ini_path, self._cookies_path, timeout=self._request_timeout)
            latest = result.get("latest") or {}
            if not latest:
                raise RuntimeError(f"No matching download found for:\n{result.get('knownFile') or row.get('file') or ''}")
            if not result.get("updateAvailable"):
                row.update({
                    "status": "OK",
                    "current": result.get("currentVersion") or row.get("current") or "",
                    "latest": latest.get("version") or row.get("latest") or "",
                    "file": latest.get("name") or row.get("file") or "",
                    "info": "No update available after recheck",
                })
                self.rowReady.emit(row)
                self.finished.emit(False)
                return

            download_url = urljoin(result.get("sourceUrl") or row.get("page_url") or "", latest.get("url") or "")
            if not download_url:
                raise RuntimeError("Latest download URL is missing.")

            archive_name = safe_archive_name(latest.get("name") or row.get("file") or "")
            archive_path = self._downloads_path / archive_name
            already_exists = archive_path.exists()
            if not already_exists:
                self.statusChanged.emit(f"Downloading update: {archive_name}")
                download_loverslab_archive(
                    download_url,
                    archive_path,
                    self._cookies_path,
                    referer=row.get("page_url") or result.get("sourceUrl") or "",
                    timeout=max(self._request_timeout, 30.0),
                )

            sidecar = write_update_download_sidecar(ini_path, archive_path, latest, download_url)
            row.update({
                "status": "Downloaded",
                "current": result.get("currentVersion") or row.get("current") or "",
                "latest": latest.get("version") or row.get("latest") or "",
                "file": archive_name,
                "archive_path": str(archive_path),
                "sidecar_path": str(sidecar),
                "downloaded_now": True,
                "latest_url": download_url,
                "latest_size": latest.get("size") or "",
                "latest_date_iso": latest.get("date_iso") or "",
                "info": (
                    "Already in MO2 downloads; metadata refreshed"
                    if already_exists
                    else f"Downloaded to MO2 downloads; metadata: {sidecar.name}"
                ),
            })
            self.rowReady.emit(row)
            self.finished.emit(False)
        except Exception as exc:
            row.update({
                "status": "Error",
                "info": str(exc),
            })
            self.rowReady.emit(row)
            self.finished.emit(False)


class LoversLabBaseTool(mobase.IPluginTool):
    TOOL_NAME = "LL Integration Base"
    TOOL_DISPLAY = "LL Integration Base"
    TOOL_DESCRIPTION = "Base LoversLab integration tool."

    def __init__(self) -> None:
        super().__init__()
        self._organizer = None

    def init(self, organizer: mobase.IOrganizer) -> bool:
        self._organizer = organizer
        return True

    def name(self) -> str:
        return self.TOOL_NAME

    def localizedName(self) -> str:
        return self.TOOL_NAME

    def author(self) -> str:
        return "LL Integration"

    def description(self) -> str:
        return self.TOOL_DESCRIPTION

    def version(self) -> mobase.VersionInfo:
        return mobase.VersionInfo("0.3.0")

    def settings(self) -> Sequence[mobase.PluginSetting]:
        paths = self._configured_paths()
        return [
            mobase.PluginSetting(
                "ll_ini_path",
                "Path to the latest download metadata generated from a LoversLab download click.",
                str(paths.get("ll_ini_path") or DEFAULT_LATEST_INI),
            ),
            mobase.PluginSetting(
                "cookies_path",
                "Path to cookies_ll.json generated by the Firefox native messaging helper.",
                str(paths.get("cookies_path") or DEFAULT_COOKIES),
            ),
        ]

    def displayName(self) -> str:
        return self.TOOL_DISPLAY

    def tooltip(self) -> str:
        return self.TOOL_DESCRIPTION

    def icon(self) -> QIcon:
        return QIcon()

    def display(self) -> None:
        raise NotImplementedError

    def _setting_path(self, key: str, default: Path) -> Path:
        paths = self._configured_paths()
        default_value = paths.get(key) or default

        if not self._organizer:
            return Path(default_value)

        value = self._organizer.pluginSetting(self.name(), key)
        if value:
            configured = Path(str(default_value))
            selected = Path(str(value))
            if self._is_stale_default_path(selected) and configured.exists():
                return configured
            if configured.exists() and not selected.exists():
                return configured
            return selected
        return Path(default_value)

    def _is_stale_default_path(self, path: Path) -> bool:
        text = str(path).replace("\\", "/").lower()
        return "/plugins/native-app/" in text

    def _configured_paths(self) -> dict:
        if not PLUGIN_PATHS_FILE.exists():
            return {}

        try:
            data = json.loads(PLUGIN_PATHS_FILE.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError:
            return {}

        return data if isinstance(data, dict) else {}

    def _downloads_storage_path(self) -> Path:
        latest_ini = self._setting_path("ll_ini_path", DEFAULT_LATEST_INI)
        return latest_ini.parent

    def _native_config_path(self) -> Path:
        cookies = self._setting_path("cookies_path", DEFAULT_COOKIES)
        native_app = cookies.parents[1] if len(cookies.parents) > 1 else cookies.parent
        return native_app / "config.json"

    def _read_native_config(self, path: Path) -> dict:
        if not path.exists():
            return {}
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _parentWidget(self):
        if self._organizer and hasattr(self._organizer, "mainWindow"):
            return self._organizer.mainWindow()
        return None

    def _choose_mod(self):
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        names = [
            name
            for name in mod_list.allModsByProfilePriority()
            if mod_list.getMod(name) is not None
        ]
        if not names:
            raise RuntimeError("No installed mods found")

        display_names = [mod_list.displayName(name) for name in names]
        selected, ok = QInputDialog.getItem(
            self._parentWidget(),
            PLUGIN_NAME,
            "Choose the installed mod:",
            display_names,
            0,
            False,
        )
        if not ok or not selected:
            raise RuntimeError("No mod selected")

        index = display_names.index(selected)
        internal_name = names[index]
        mod = mod_list.getMod(internal_name)
        if mod is None:
            raise RuntimeError(f"Could not open mod: {selected}")

        return selected, mod

    def _choose_mod_with_hint(self, hint: str):
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        names = [
            name
            for name in mod_list.allModsByProfilePriority()
            if mod_list.getMod(name) is not None
        ]
        if not names:
            raise RuntimeError("No installed mods found")

        display_names = [mod_list.displayName(name) for name in names]
        current = self._best_mod_index(display_names, hint)
        selected, ok = QInputDialog.getItem(
            self._parentWidget(),
            PLUGIN_NAME,
            "Choose the installed mod:",
            display_names,
            current,
            False,
        )
        if not ok or not selected:
            raise RuntimeError("No mod selected")

        index = display_names.index(selected)
        mod = mod_list.getMod(names[index])
        if mod is None:
            raise RuntimeError(f"Could not open mod: {selected}")

        return selected, mod

    def _best_mod_index(self, mod_names: list[str], hint: str) -> int:
        hint_tokens = self._tokens(hint)
        if not hint_tokens:
            return 0

        best_index = 0
        best_score = -1
        for index, name in enumerate(mod_names):
            tokens = self._tokens(name)
            score = len(hint_tokens & tokens)
            if score > best_score:
                best_index = index
                best_score = score
        return best_index

    def _tokens(self, text: str) -> set[str]:
        clean = re.sub(r"\.(?:7z|zip|rar)$", "", text, flags=re.IGNORECASE)
        clean = re.sub(r"\bv?\d+(?:\.\d+){1,3}\b", " ", clean, flags=re.IGNORECASE)
        return {token for token in re.findall(r"[a-z0-9]+", clean.lower()) if len(token) > 2}

    def _write_mo2_meta_ini(self, mod, page_url: str, version: str) -> None:
        meta_path = mod_meta_path(mod)
        config = configparser.ConfigParser(interpolation=None)
        config.optionxform = str
        config.read(meta_path, encoding="utf-8")
        if "General" not in config:
            config["General"] = {}

        general = config["General"]
        if page_url:
            general["url"] = page_url
            general["hasCustomURL"] = "true"
            general["repository"] = "LoversLab"
        if version:
            general["version"] = f"{version}.0" if version.count(".") == 2 else version

        with meta_path.open("w", encoding="utf-8") as file:
            config.write(file, space_around_delimiters=False)


class LoversLabIntegrationTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration"
    TOOL_DISPLAY = "Check LoversLab Updates"
    TOOL_DESCRIPTION = "Checks LoversLab download pages for updates using exported Firefox cookies."

    def display(self) -> None:
        try:
            mod_name, mod = self._choose_mod()
            ini_path = self._mod_ll_ini_path(mod)
            cookies_path = self._setting_path("cookies_path", DEFAULT_COOKIES)
            out_path = self._output_path()
            result = check_ini_for_updates(ini_path, cookies_path, out_path=out_path)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                self._format_error(exc),
            )
            return

        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            self._format_result(result, mod_name),
        )

    def _mod_ll_ini_path(self, mod) -> Path:
        ini_path = mod_ll_metadata_path(mod)
        if ini_path is None:
            expected = mod_meta_path(mod)
            raise RuntimeError(
                f"This mod has no LL Integration metadata yet.\n\n"
                f"Expected metadata section in:\n{expected}\n\n"
                "Use 'Create Source Link' or install a captured download first."
            )
        return ini_path

    def _output_path(self) -> Path:
        if self._organizer:
            data_path = Path(str(self._organizer.pluginDataPath())) / "ll_integration"
            data_path.mkdir(parents=True, exist_ok=True)
            return data_path / "update_check.json"

        return Path(__file__).resolve().parent / "update_check.json"

    def _format_result(self, result: dict, mod_name: str = "") -> str:
        latest = result.get("latest")
        seen = len(result.get("downloadsSeen") or [])
        header = f"Mod: {mod_name}\n\n" if mod_name else ""

        if not latest:
            return (
                f"{header}"
                f"No matching download found for:\n{result.get('knownFile')}\n\n"
                f"Downloads seen: {seen}"
            )

        current_version = result.get("currentVersion")
        latest_version = latest.get("version")
        status = "Update available" if result.get("updateAvailable") else "Up to date"

        return (
            f"{header}"
            f"{status}\n\n"
            f"Current: {current_version}\n"
            f"Latest: {latest_version}\n"
            f"File: {latest.get('name')}\n"
            f"Size: {latest.get('size') or 'unknown'}\n"
            f"Downloads seen: {seen}"
        )

    def _format_error(self, exc: Exception) -> str:
        ini_path = self._setting_path("ll_ini_path", DEFAULT_LATEST_INI)
        cookies_path = self._setting_path("cookies_path", DEFAULT_COOKIES)
        cookie_names = []

        if cookies_path.exists():
            try:
                cookie_names = sorted(load_ll_cookies(cookies_path, required_only=False).keys())
            except Exception:
                cookie_names = ["<could not read cookies>"]

        return (
            f"LoversLab update check failed:\n\n{exc}\n\n"
            f"INI path:\n{ini_path}\n\n"
            f"Cookies path:\n{cookies_path}\n"
            f"Cookies file exists: {cookies_path.exists()}\n"
            f"Cookie names: {', '.join(cookie_names) if cookie_names else '<none>'}"
        )


class LoversLabBindLatestTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Bind Latest"
    TOOL_DISPLAY = "Bind Latest LoversLab Download to Mod"
    TOOL_DESCRIPTION = "Stores the latest LoversLab sidecar metadata in the installed mod meta.ini."

    def display(self) -> None:
        try:
            sidecar_ini = self._latest_sidecar_ini()
            ll_info = self._read_ll_info(sidecar_ini)
            mod_name, mod = self._choose_mod_with_hint(
                ll_info.get("archive_name") or ll_info.get("file_name") or ""
            )
            if not self._confirm_bind(mod_name, sidecar_ini, ll_info):
                return

            target = write_mod_ll_metadata_from_file(mod, sidecar_ini)

            try:
                self._apply_mod_metadata(mod, target)
                if self._organizer:
                    self._organizer.modDataChanged(mod)
            except Exception:
                pass
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Bind failed:\n\n{exc}",
            )
            return

        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            f"Bound LoversLab metadata to:\n{mod_name}\n\nStored in:\n{target}",
        )

    def _latest_sidecar_ini(self) -> Path:
        completions_path = self._downloads_storage_path() / "download_completions.json"
        if not completions_path.exists():
            raise RuntimeError(f"No download completions found:\n{completions_path}")

        completions = json.loads(completions_path.read_text(encoding="utf-8"))
        if not isinstance(completions, list) or not completions:
            raise RuntimeError(f"No download completions found in:\n{completions_path}")

        for completion in reversed(completions):
            archive = completion.get("archivePath")
            if not archive:
                continue
            sidecar = Path(archive).with_name(f"{Path(archive).name}.ll.ini")
            if sidecar.exists():
                return sidecar

        raise RuntimeError("No matching .ll.ini sidecar found for recent downloads")

    def _read_ll_info(self, ini_path: Path) -> dict:
        config = configparser.ConfigParser(interpolation=None)
        config.read(ini_path, encoding="utf-8")
        if LL_SECTION not in config:
            raise RuntimeError(f"Invalid LL metadata file:\n{ini_path}")

        ll = config[LL_SECTION]
        return {
            "file_name": ll.get("file_name", "").strip(),
            "archive_name": ll.get("archive_name", "").strip(),
            "version": ll.get("version", "").strip(),
            "page_url": ll.get("page_url", "").strip(),
        }

    def _confirm_bind(self, mod_name: str, sidecar_ini: Path, ll_info: dict) -> bool:
        archive_name = ll_info.get("archive_name") or ll_info.get("file_name") or "<unknown>"
        version = ll_info.get("version") or "<unknown>"
        page_url = ll_info.get("page_url") or "<unknown>"
        result = QMessageBox.question(
            self._parentWidget(),
            PLUGIN_NAME,
            "Bind this LoversLab metadata?\n\n"
            f"Mod:\n{mod_name}\n\n"
            f"Archive:\n{archive_name}\n"
            f"Version:\n{version}\n\n"
            f"Page:\n{page_url}\n\n"
            f"Source:\n{sidecar_ini}",
        )
        return result == QMessageBox.StandardButton.Yes

    def _apply_mod_metadata(self, mod, ini_path: Path) -> None:
        if mod_has_nexus_identity(mod)[0]:
            return

        config = configparser.ConfigParser(interpolation=None)
        config.read(ini_path, encoding="utf-8")
        ll = config[LL_SECTION]

        page_url = ll.get("page_url", "").strip()
        version = ll.get("version", "").strip()
        if page_url:
            mod.setUrl(page_url)
        if version:
            mod.setVersion(mobase.VersionInfo(version))
        self._write_mo2_meta_ini(mod, page_url, version)

    def _write_mo2_meta_ini(self, mod, page_url: str, version: str) -> None:
        meta_path = mod_meta_path(mod)
        config = configparser.ConfigParser(interpolation=None)
        config.optionxform = str
        config.read(meta_path, encoding="utf-8")
        if "General" not in config:
            config["General"] = {}

        general = config["General"]
        if page_url:
            general["url"] = page_url
            general["hasCustomURL"] = "true"
            general["repository"] = "LoversLab"
        if version:
            general["version"] = f"{version}.0" if version.count(".") == 2 else version

        with meta_path.open("w", encoding="utf-8") as file:
            config.write(file, space_around_delimiters=False)

class LoversLabCheckAllTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Check All"
    TOOL_DISPLAY = "Check All LoversLab Updates"
    TOOL_DESCRIPTION = "Checks every installed mod with LL Integration metadata for LoversLab updates."

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def display(self) -> None:
        try:
            jobs = self._collect_jobs()
            cookies_path = self._setting_path("cookies_path", DEFAULT_COOKIES)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Check all failed:\n\n{exc}",
            )
            return

        if not jobs:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                    "No installed mods with LL Integration metadata found.",
            )
            return

        self._show_results(jobs, cookies_path)

    def _collect_jobs(self) -> list[dict]:
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        jobs = []

        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is None:
                continue

            ini_path = mod_ll_metadata_path(mod)
            if ini_path is None:
                continue

            display_name = mod_list.displayName(internal_name)
            page_url = self._ll_page_url(ini_path)
            current_version = self._ll_current_version(ini_path)
            jobs.append({
                "mod": display_name,
                "internal_name": internal_name,
                "mod_path": str(mod.absolutePath()),
                "ini_path": str(ini_path),
                "page_url": page_url,
                "current": current_version,
                "file": self._ll_file_name(ini_path),
                "fixed": self._ll_fixed_update(ini_path),
                "update_mode": self._ll_update_mode(ini_path),
            })

        return jobs

    def _ll_page_url(self, ini_path: Path) -> str:
        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return ""
        return ll.get("page_url", "").strip()

    def _ll_current_version(self, ini_path: Path) -> str:
        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return ""
        return ll.get("version", "").strip()

    def _ll_file_name(self, ini_path: Path) -> str:
        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return ""
        return (
            ll.get("file_pattern", "").strip()
            or ll.get("file_name", "").strip()
            or ll.get("archive_name", "").strip()
        )

    def _ll_fixed_update(self, ini_path: Path) -> bool:
        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return False
        mode = normalized_update_mode(
            ll.get("update_mode"),
            fixed=(
                truthy(ll.get("fixed_version"))
                or truthy(ll.get("manual_update"))
                or truthy(ll.get("skip_update_check"))
            ),
        )
        return (
            mode == UPDATE_MODE_SKIP
            or truthy(ll.get("fixed_version"))
            or truthy(ll.get("manual_update"))
            or truthy(ll.get("skip_update_check"))
        )

    def _ll_update_mode(self, ini_path: Path) -> str:
        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return UPDATE_MODE_MANUAL
        return normalized_update_mode(
            ll.get("update_mode"),
            fixed=(
                truthy(ll.get("fixed_version"))
                or truthy(ll.get("manual_update"))
                or truthy(ll.get("skip_update_check"))
            ),
        )

    def _show_results(self, jobs: list[dict], cookies_path: Path) -> None:
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("LoversLab Update Check")
        dialog.resize(1370, 640)
        dialog.setMinimumSize(1180, 520)

        table = QTableWidget(dialog)
        progress = QProgressBar(dialog)
        progress.setRange(0, len(jobs))
        progress.setValue(0)
        update_cache = self._load_update_cache()
        cached_count = sum(1 for job in jobs if self._cache_entry_for_job(job, update_cache))
        progress_label = QLabel(
            f"Loaded {len(jobs)} LoversLab links"
            + (f" with {cached_count} cached results. " if cached_count else ". ")
            + "Click Fetch Updates to check versions."
        )
        pacing = self._load_fetch_pacing()

        fetch_updates = QPushButton("Fetch Updates")
        cancel = QPushButton("Cancel")
        close = QPushButton("Close")
        cancel.setEnabled(False)
        close.clicked.connect(dialog.accept)


        delay = QDoubleSpinBox(dialog)
        delay.setRange(0.0, 10.0)
        delay.setDecimals(1)
        delay.setSingleStep(0.1)
        delay.setSuffix(" s")
        delay.setValue(float(pacing["request_delay"]))
        delay.setToolTip("Delay before each LoversLab request. Increase this if LoversLab returns 403 or slows down.")

        batch_size = QSpinBox(dialog)
        batch_size.setRange(0, 500)
        batch_size.setValue(int(pacing["batch_size"]))
        batch_size.setToolTip("Number of requests before a cooldown. Set 0 to disable cooldown batches.")

        batch_pause = QDoubleSpinBox(dialog)
        batch_pause.setRange(0.0, 300.0)
        batch_pause.setDecimals(1)
        batch_pause.setSingleStep(5.0)
        batch_pause.setSuffix(" s")
        batch_pause.setValue(float(pacing["batch_pause"]))
        batch_pause.setToolTip("Cooldown duration after each batch. Set 0 to disable.")

        request_timeout = QDoubleSpinBox(dialog)
        request_timeout.setRange(3.0, 120.0)
        request_timeout.setDecimals(1)
        request_timeout.setSingleStep(1.0)
        request_timeout.setSuffix(" s")
        request_timeout.setValue(float(pacing["request_timeout"]))
        request_timeout.setToolTip(
            "Maximum time to wait for one LoversLab request before skipping that row. "
            "Use a high value like 120s to wait longer; timeout cannot be disabled safely."
        )

        filter_mode = QComboBox(dialog)
        filter_mode.addItems([
            "All links",
            "Updates",
            "OK",
            "Unknown / missing version",
            "Manual links",
            "Errors / skipped",
            "Not checked",
        ])
        filter_mode.setToolTip("Filter the visible LL Integration rows. This does not change what metadata is stored.")

        filter_text = QLineEdit(dialog)
        filter_text.setPlaceholderText("Search mod, file, or page")
        filter_text.setClearButtonEnabled(True)

        filter_count = QLabel("", dialog)

        pacing_widgets = [delay, batch_size, batch_pause, request_timeout]
        fetch_updates.clicked.connect(
            lambda _checked=False: self._start_check_worker(
                dialog,
                table,
                progress,
                progress_label,
                fetch_updates,
                cancel,
                close,
                jobs,
                cookies_path,
                delay.value(),
                batch_size.value(),
                batch_pause.value(),
                request_timeout.value(),
                pacing_widgets,
            )
        )
        cancel.clicked.connect(lambda _checked=False: self._cancel_check_worker(dialog, cancel))

        self._prepare_results_table(table)
        table._ll_jobs = jobs
        table._ll_dialog = dialog
        table._ll_progress = progress
        table._ll_progress_label = progress_label
        table._ll_fetch_updates = fetch_updates
        table._ll_cancel = cancel
        table._ll_close = close
        table._ll_cookies_path = cookies_path
        table._ll_pacing_widgets = pacing_widgets
        table._ll_update_cache = update_cache
        table._ll_filter_mode = filter_mode
        table._ll_filter_text = filter_text
        table._ll_filter_count = filter_count
        self._populate_results_table(table, jobs)
        filter_mode.currentTextChanged.connect(lambda _text: self._apply_results_filter(table))
        filter_text.textChanged.connect(lambda _text: self._apply_results_filter(table))
        self._apply_results_filter(table)

        controls = QHBoxLayout()
        controls.addWidget(fetch_updates)
        controls.addWidget(cancel)
        controls.addStretch(1)
        controls.addWidget(close)

        pacing_controls = QHBoxLayout()
        pacing_controls.addWidget(QLabel("Delay"))
        pacing_controls.addWidget(delay)
        pacing_controls.addWidget(QLabel("Cooldown every"))
        pacing_controls.addWidget(batch_size)
        pacing_controls.addWidget(QLabel("requests for"))
        pacing_controls.addWidget(batch_pause)
        timeout_label = QLabel("Timeout / request")
        timeout_label.setToolTip(
            "This stays enabled so one stuck LoversLab page cannot block the whole list forever."
        )
        pacing_controls.addWidget(timeout_label)
        pacing_controls.addWidget(request_timeout)
        pacing_controls.addStretch(1)

        filter_controls = QHBoxLayout()
        filter_controls.addWidget(QLabel("Filter"))
        filter_controls.addWidget(filter_mode)
        filter_controls.addWidget(filter_text, 1)
        filter_controls.addWidget(filter_count)

        layout = QVBoxLayout(dialog)
        layout.addLayout(filter_controls)
        layout.addWidget(table)
        layout.addWidget(progress)
        layout.addWidget(progress_label)
        layout.addLayout(pacing_controls)
        layout.addLayout(controls)
        dialog.setLayout(layout)
        dialog.exec()

    def _start_check_worker(
        self,
        dialog: QDialog,
        table: QTableWidget,
        progress: QProgressBar,
        progress_label: QLabel,
        fetch_updates: QPushButton,
        cancel: QPushButton,
        close: QPushButton,
        jobs: list[dict],
        cookies_path: Path,
        request_delay: float,
        batch_size: int,
        batch_pause: float,
        request_timeout: float,
        pacing_widgets: list,
    ) -> None:
        current_worker = getattr(dialog, "_ll_worker", None)
        if current_worker is not None:
            current_worker.cancel()

        self._save_fetch_pacing(request_delay, batch_size, batch_pause, request_timeout)
        self._reset_rows_for_fetch(table, jobs)
        progress.setRange(0, len(jobs))
        progress.setValue(0)
        progress_label.setText(
            f"Fetching updates 0 / {len(jobs)}. "
            f"Pacing: {request_delay:.1f}s/request"
            + (f", {batch_pause:.0f}s cooldown every {batch_size} requests." if batch_size > 0 and batch_pause > 0 else ".")
            + f" Timeout: {request_timeout:.1f}s."
        )
        fetch_updates.setEnabled(False)
        cancel.setEnabled(True)
        close.setEnabled(False)
        for widget in pacing_widgets:
            widget.setEnabled(False)

        thread = QThread(dialog)
        log_path = self._fetch_log_path()
        worker = CheckAllWorker(jobs, cookies_path, log_path, request_delay, batch_size, batch_pause, request_timeout)
        worker.moveToThread(thread)

        worker.rowReady.connect(lambda row: self._update_result_row(table, row))
        worker.progressChanged.connect(
            lambda done, total: self._set_progress(progress, progress_label, done, total)
        )
        worker.statusChanged.connect(lambda message: progress_label.setText(message))
        worker.finished.connect(
            lambda cancelled: self._finish_check_worker(
                dialog,
                thread,
                progress_label,
                fetch_updates,
                cancel,
                close,
                cancelled,
                pacing_widgets,
            )
        )
        thread.started.connect(worker.run)

        dialog._ll_thread = thread
        dialog._ll_worker = worker
        dialog._ll_fetch_log_path = log_path
        dialog._ll_progress_label = progress_label
        thread.start()

    def _cancel_check_worker(self, dialog: QDialog, cancel: QPushButton) -> None:
        worker = getattr(dialog, "_ll_worker", None)
        if worker is not None:
            worker.cancel()
            cancel.setEnabled(False)
            cancel.setText("Cancelling...")
            progress_label = getattr(dialog, "_ll_progress_label", None)
            if progress_label is not None:
                progress_label.setText("Cancelling after current UI handoff...")

    def _start_single_check_worker(
        self,
        table: QTableWidget,
        row_index: int,
        row: dict,
        fetch_button: QPushButton,
    ) -> None:
        dialog = getattr(table, "_ll_dialog", None)
        if dialog is None:
            return

        current_worker = getattr(dialog, "_ll_worker", None)
        if current_worker is not None:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "A fetch is already running. Cancel or wait for it to finish first.",
            )
            return

        jobs = getattr(table, "_ll_jobs", [])
        job = dict(jobs[row_index]) if 0 <= row_index < len(jobs) else dict(row)
        job["row_index"] = row_index
        pacing = self._current_fetch_pacing_from_table(table)
        self._save_fetch_pacing(
            pacing["request_delay"],
            pacing["batch_size"],
            pacing["batch_pause"],
            pacing["request_timeout"],
        )

        progress = getattr(table, "_ll_progress", None)
        progress_label = getattr(table, "_ll_progress_label", None)
        fetch_updates = getattr(table, "_ll_fetch_updates", None)
        cancel = getattr(table, "_ll_cancel", None)
        close = getattr(table, "_ll_close", None)
        pacing_widgets = getattr(table, "_ll_pacing_widgets", [])
        cookies_path = getattr(table, "_ll_cookies_path", DEFAULT_COOKIES)

        self._set_result_row_values(table, row_index, self._pending_row(job, "Queued", "Waiting"))
        if progress is not None:
            progress.setRange(0, 1)
            progress.setValue(0)
        if progress_label is not None:
            progress_label.setText(
                f"Fetching 1 mod: {job.get('mod') or ''}. Timeout: {pacing['request_timeout']:.1f}s."
            )
        if fetch_updates is not None:
            fetch_updates.setEnabled(False)
        if cancel is not None:
            cancel.setText("Cancel")
            cancel.setEnabled(True)
        if close is not None:
            close.setEnabled(False)
        fetch_button.setEnabled(False)
        for widget in pacing_widgets:
            widget.setEnabled(False)

        thread = QThread(dialog)
        log_path = self._fetch_log_path()
        worker = CheckAllWorker(
            [job],
            Path(str(cookies_path)),
            log_path,
            request_delay=0.0,
            batch_size=0,
            batch_pause=0.0,
            request_timeout=pacing["request_timeout"],
        )
        worker.moveToThread(thread)

        worker.rowReady.connect(lambda result: self._update_result_row(table, result))
        if progress is not None and progress_label is not None:
            worker.progressChanged.connect(
                lambda done, total: self._set_progress(progress, progress_label, done, total)
            )
        if progress_label is not None:
            worker.statusChanged.connect(lambda message: progress_label.setText(message))
        worker.finished.connect(
            lambda cancelled: self._finish_single_check_worker(
                dialog,
                thread,
                progress_label,
                fetch_updates,
                cancel,
                close,
                fetch_button,
                cancelled,
                pacing_widgets,
            )
        )
        thread.started.connect(worker.run)

        dialog._ll_thread = thread
        dialog._ll_worker = worker
        dialog._ll_fetch_log_path = log_path
        dialog._ll_progress_label = progress_label
        thread.start()

    def _start_try_update_worker(
        self,
        table: QTableWidget,
        row_index: int,
        action_button: QPushButton,
    ) -> None:
        dialog = getattr(table, "_ll_dialog", None)
        if dialog is None:
            return

        current_worker = getattr(dialog, "_ll_worker", None)
        if current_worker is not None:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "A fetch or update download is already running. Cancel or wait for it to finish first.",
            )
            return

        row = self._current_row_data(table, row_index)
        if not self._row_try_update_enabled(row):
            return

        pacing = self._current_fetch_pacing_from_table(table)
        cookies_path = getattr(table, "_ll_cookies_path", DEFAULT_COOKIES)
        downloads_path = self._active_downloads_path()
        if not downloads_path:
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, "MO2 downloads path is not available.")
            return

        progress = getattr(table, "_ll_progress", None)
        progress_label = getattr(table, "_ll_progress_label", None)
        fetch_updates = getattr(table, "_ll_fetch_updates", None)
        cancel = getattr(table, "_ll_cancel", None)
        close = getattr(table, "_ll_close", None)
        pacing_widgets = getattr(table, "_ll_pacing_widgets", [])

        row["info"] = "Downloading update..."
        self._set_result_row_values(table, row_index, row)
        if progress is not None:
            progress.setRange(0, 0)
        if progress_label is not None:
            progress_label.setText(f"Downloading update: {row.get('mod') or ''}")
        if fetch_updates is not None:
            fetch_updates.setEnabled(False)
        if cancel is not None:
            cancel.setEnabled(False)
        if close is not None:
            close.setEnabled(False)
        action_button.setEnabled(False)
        for widget in pacing_widgets:
            widget.setEnabled(False)

        thread = QThread(dialog)
        worker = TryUpdateWorker(
            row,
            Path(str(cookies_path)),
            downloads_path,
            request_timeout=pacing["request_timeout"],
        )
        worker.moveToThread(thread)
        worker.rowReady.connect(lambda result: self._update_result_row(table, result))
        if progress_label is not None:
            worker.statusChanged.connect(lambda message: progress_label.setText(message))
        worker.finished.connect(
            lambda cancelled: self._finish_try_update_worker(
                dialog,
                thread,
                progress,
                progress_label,
                fetch_updates,
                cancel,
                close,
                action_button,
                cancelled,
                pacing_widgets,
            )
        )
        thread.started.connect(worker.run)

        dialog._ll_thread = thread
        dialog._ll_worker = worker
        thread.start()

    def _active_downloads_path(self) -> Path | None:
        try:
            if self._organizer:
                downloads = Path(str(self._organizer.downloadsPath()))
                if str(downloads):
                    return downloads
        except Exception:
            pass

        try:
            config = self._read_native_config(self._native_config_path())
            downloads = Path(str(config.get("mo2_downloads_path") or ""))
            return downloads if str(downloads) else None
        except Exception:
            return None

    def _current_fetch_pacing_from_table(self, table: QTableWidget) -> dict:
        pacing = self._load_fetch_pacing()
        widgets = getattr(table, "_ll_pacing_widgets", [])
        if len(widgets) >= 4:
            pacing["request_delay"] = float(widgets[0].value())
            pacing["batch_size"] = int(widgets[1].value())
            pacing["batch_pause"] = float(widgets[2].value())
            pacing["request_timeout"] = float(widgets[3].value())
        return pacing

    def _finish_single_check_worker(
        self,
        dialog: QDialog,
        thread: QThread,
        progress_label: QLabel,
        fetch_updates: QPushButton,
        cancel: QPushButton,
        close: QPushButton,
        fetch_button: QPushButton,
        cancelled: bool,
        pacing_widgets: list,
    ) -> None:
        log_path = getattr(dialog, "_ll_fetch_log_path", None)
        suffix = f" Timing log: {log_path}" if log_path else ""
        if progress_label is not None:
            progress_label.setText(("Cancelled" if cancelled else "Done") + suffix)
        if fetch_updates is not None:
            fetch_updates.setEnabled(True)
        if cancel is not None:
            cancel.setText("Cancel")
            cancel.setEnabled(False)
        if close is not None:
            close.setEnabled(True)
        fetch_button.setEnabled(True)
        for widget in pacing_widgets:
            widget.setEnabled(True)
        dialog._ll_worker = None
        thread.quit()
        thread.wait(1000)

    def _finish_try_update_worker(
        self,
        dialog: QDialog,
        thread: QThread,
        progress: QProgressBar,
        progress_label: QLabel,
        fetch_updates: QPushButton,
        cancel: QPushButton,
        close: QPushButton,
        action_button: QPushButton,
        cancelled: bool,
        pacing_widgets: list,
    ) -> None:
        if progress is not None:
            progress.setRange(0, 1)
            progress.setValue(1)
        if progress_label is not None:
            progress_label.setText("Done" if not cancelled else "Cancelled")
        if fetch_updates is not None:
            fetch_updates.setEnabled(True)
        if cancel is not None:
            cancel.setText("Cancel")
            cancel.setEnabled(False)
        if close is not None:
            close.setEnabled(True)
        for widget in pacing_widgets:
            widget.setEnabled(True)
        dialog._ll_worker = None
        thread.quit()
        thread.wait(1000)

    def _finish_check_worker(
        self,
        dialog: QDialog,
        thread: QThread,
        progress_label: QLabel,
        fetch_updates: QPushButton,
        cancel: QPushButton,
        close: QPushButton,
        cancelled: bool,
        pacing_widgets: list,
    ) -> None:
        log_path = getattr(dialog, "_ll_fetch_log_path", None)
        suffix = f" Timing log: {log_path}" if log_path else ""
        progress_label.setText(("Cancelled" if cancelled else "Done") + suffix)
        fetch_updates.setEnabled(True)
        cancel.setText("Cancel")
        cancel.setEnabled(False)
        close.setEnabled(True)
        for widget in pacing_widgets:
            widget.setEnabled(True)
        dialog._ll_worker = None
        thread.quit()
        thread.wait(1000)

    def _set_progress(self, progress: QProgressBar, label: QLabel, done: int, total: int) -> None:
        progress.setValue(done)
        label.setText(f"Fetching updates {done} / {total}")

    def _fetch_log_path(self) -> Path:
        if self._organizer:
            return Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "fetch_update_timings.jsonl"
        return Path(__file__).resolve().parent / "fetch_update_timings.jsonl"

    def _fetch_pacing_config_path(self) -> Path:
        if self._organizer:
            return Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "fetch_pacing.json"
        return Path(__file__).resolve().parent / "fetch_pacing.json"

    def _update_cache_path(self) -> Path:
        if self._organizer:
            return Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "update_cache.json"
        return Path(__file__).resolve().parent / "update_cache.json"

    def _load_update_cache(self) -> dict:
        path = self._update_cache_path()
        if not path.exists():
            return {"version": UPDATE_CACHE_VERSION, "entries": {}}

        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {"version": UPDATE_CACHE_VERSION, "entries": {}}

        if not isinstance(data, dict):
            return {"version": UPDATE_CACHE_VERSION, "entries": {}}

        entries = data.get("entries")
        if not isinstance(entries, dict):
            entries = {}

        return {"version": UPDATE_CACHE_VERSION, "entries": entries}

    def _save_update_cache(self, cache: dict) -> None:
        path = self._update_cache_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(cache, indent=2, ensure_ascii=False), encoding="utf-8")
        except Exception:
            pass

    def _job_cache_key(self, job: dict) -> str:
        return str(Path(str(job.get("ini_path") or "")).resolve()).lower()

    def _job_signature(self, job: dict) -> dict:
        return {
            "page_url": job.get("page_url") or "",
            "current": job.get("current") or "",
            "file": job.get("file") or "",
            "fixed": bool(job.get("fixed")),
            "update_mode": job.get("update_mode") or "",
        }

    def _cache_entry_for_job(self, job: dict, cache: dict) -> dict | None:
        key = self._job_cache_key(job)
        entry = (cache.get("entries") or {}).get(key)
        if not isinstance(entry, dict):
            return None
        if entry.get("signature") != self._job_signature(job):
            return None
        return entry

    def _now_label(self) -> str:
        return time.strftime("%Y-%m-%d %H:%M")

    def _cache_row_payload(self, row: dict, checked_at: str) -> dict:
        return {
            "checked_at": checked_at,
            "ini_path": row.get("ini_path") or "",
            "internal_name": row.get("internal_name") or "",
            "mod": row.get("mod") or "",
            "status": row.get("status") or "",
            "current": row.get("current") or "",
            "latest": row.get("latest") or "",
            "file": row.get("file") or "",
            "archive_path": row.get("archive_path") or "",
            "sidecar_path": row.get("sidecar_path") or "",
            "latest_url": row.get("latest_url") or "",
            "latest_size": row.get("latest_size") or "",
            "latest_date_iso": row.get("latest_date_iso") or "",
            "page_url": row.get("page_url") or "",
            "fixed": bool(row.get("fixed")),
            "update_mode": row.get("update_mode") or (UPDATE_MODE_SKIP if row.get("fixed") else UPDATE_MODE_MANUAL),
            "info": row.get("info") or "",
        }

    def _cached_display_row(self, job: dict, row_index: int, entry: dict) -> dict | None:
        last_result = entry.get("last_result")
        if not isinstance(last_result, dict):
            return None

        display = dict(last_result)
        display.update({
            "row_index": row_index,
            "ini_path": job.get("ini_path") or "",
            "internal_name": job.get("internal_name") or "",
            "mod": job.get("mod") or display.get("mod") or "",
            "page_url": job.get("page_url") or display.get("page_url") or "",
            "current": job.get("current") or display.get("current") or "",
            "fixed": bool(job.get("fixed")),
            "update_mode": job.get("update_mode") or display.get("update_mode") or "",
            "archive_path": display.get("archive_path") or "",
            "sidecar_path": display.get("sidecar_path") or "",
            "latest_url": display.get("latest_url") or "",
            "latest_size": display.get("latest_size") or "",
            "latest_date_iso": display.get("latest_date_iso") or "",
        })

        last_success = entry.get("last_success")
        if display.get("status") == "Downloaded":
            if self._downloaded_row_was_installed(display):
                display["status"] = "Ready"
                display["latest"] = ""
                display["info"] = (
                    f"Cached {display.get('checked_at') or 'unknown'}; "
                    "downloaded archive appears installed"
                )
            elif not self._downloaded_archive_exists(display):
                display["status"] = "Update"
                display["info"] = (
                    f"Cached {display.get('checked_at') or 'unknown'}; "
                    "downloaded archive is missing"
                )
            else:
                display["info"] = (
                    f"Cached {display.get('checked_at') or 'unknown'}; "
                    "downloaded archive waiting for manual install"
                )
            return display


        if display.get("status") in {"Error", "Skipped"} and isinstance(last_success, dict):
            display["latest"] = display.get("latest") or last_success.get("latest") or ""
            display["file"] = display.get("file") or last_success.get("file") or job.get("file") or ""
            display["info"] = (
                f"Last fetch {display.get('checked_at') or 'unknown'}: {display.get('info') or display.get('status')}; "
                f"last OK {last_success.get('checked_at') or 'unknown'}"
            )
        else:
            display["info"] = f"Cached {display.get('checked_at') or 'unknown'}; {display.get('info') or ''}".strip()

        return display

    def _downloaded_row_was_installed(self, row: dict) -> bool:
        current = str(row.get("current") or "").strip()
        latest = str(row.get("latest") or "").strip()
        if not current or not latest:
            return False
        try:
            return compare_versions(current, latest) >= 0
        except Exception:
            return current == latest

    def _downloaded_archive_exists(self, row: dict) -> bool:
        archive_path = str(row.get("archive_path") or "").strip()
        if archive_path and Path(archive_path).exists():
            return True

        file_name = str(row.get("file") or "").strip()
        if not file_name:
            return False

        explicit_path = Path(file_name)
        if explicit_path.is_absolute():
            return explicit_path.exists()

        downloads = self._active_downloads_path()
        if downloads:
            candidate = downloads / file_name
            if candidate.exists():
                return True

        ini_path = Path(str(row.get("ini_path") or ""))
        if ini_path.exists():
            try:
                ll = read_ll_section(ini_path)
                archive_path = ll.get("archive_path", "").strip()
                if archive_path and Path(archive_path).exists():
                    return True
            except Exception:
                pass

        return False

    def _mark_update_as_downloaded_if_archive_exists(self, row: dict) -> dict:
        archive = self._existing_latest_archive(row)
        if not archive:
            return row

        updated = dict(row)
        sidecar = Path(f"{archive}.ll.ini")
        if not sidecar.exists():
            try:
                download_url = str(row.get("latest_url") or "")
                write_update_download_sidecar(
                    Path(str(row.get("ini_path") or "")),
                    archive,
                    {
                        "name": archive.name,
                        "version": row.get("latest") or "",
                        "size": row.get("latest_size") or "",
                        "date_iso": row.get("latest_date_iso") or "",
                    },
                    download_url,
                )
            except Exception:
                pass

        updated["status"] = "Downloaded"
        updated["archive_path"] = str(archive)
        updated["sidecar_path"] = str(sidecar)
        updated["info"] = "Latest archive already exists in MO2 downloads; waiting for manual install"
        return updated

    def _existing_latest_archive(self, row: dict) -> Path | None:
        file_name = safe_archive_name(str(row.get("file") or "").strip())
        if not file_name:
            return None

        candidates = []
        downloads = self._active_downloads_path()
        if downloads:
            candidates.append(downloads / file_name)

        ini_path = Path(str(row.get("ini_path") or ""))
        if ini_path.exists():
            try:
                ll = read_ll_section(ini_path)
                archive_path = ll.get("archive_path", "").strip()
                if archive_path:
                    candidates.append(Path(archive_path))
            except Exception:
                pass

        for candidate in candidates:
            if candidate.exists() and candidate.is_file():
                return candidate
        return None

    def _cache_result_row(self, table: QTableWidget, row: dict) -> None:
        status = row.get("status") or ""
        if status not in {"OK", "Update", "Unknown", "Downloaded", "Error", "Skipped"}:
            return

        jobs = getattr(table, "_ll_jobs", [])
        row_index = int(row.get("row_index", -1))
        job = jobs[row_index] if 0 <= row_index < len(jobs) else row
        cache = getattr(table, "_ll_update_cache", None)
        if cache is None:
            cache = self._load_update_cache()
            table._ll_update_cache = cache

        entry = self._cache_entry_for_job(job, cache) or {}
        checked_at = self._now_label()
        payload = self._cache_row_payload(row, checked_at)
        entry["signature"] = self._job_signature(job)
        entry["last_result"] = payload
        if status in {"OK", "Update"}:
            entry["last_success"] = payload

        cache.setdefault("entries", {})[self._job_cache_key(job)] = entry
        self._save_update_cache(cache)

    def _remove_cached_result(self, table: QTableWidget, row: dict) -> None:
        cache = getattr(table, "_ll_update_cache", None)
        if cache is None:
            cache = self._load_update_cache()
            table._ll_update_cache = cache

        entries = cache.get("entries") or {}
        key = self._job_cache_key(row)
        if key in entries:
            entries.pop(key, None)
            cache["entries"] = entries
            self._save_update_cache(cache)

    def _load_fetch_pacing(self) -> dict:
        pacing = {
            "request_delay": UPDATE_REQUEST_DELAY_SECONDS,
            "batch_size": UPDATE_BATCH_SIZE,
            "batch_pause": UPDATE_BATCH_PAUSE_SECONDS,
            "request_timeout": UPDATE_REQUEST_TIMEOUT_SECONDS,
        }
        path = self._fetch_pacing_config_path()
        if not path.exists():
            return pacing

        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
            if isinstance(data, dict):
                pacing["request_delay"] = max(0.0, min(10.0, float(data.get("request_delay", pacing["request_delay"]))))
                pacing["batch_size"] = max(0, min(500, int(data.get("batch_size", pacing["batch_size"]))))
                pacing["batch_pause"] = max(0.0, min(300.0, float(data.get("batch_pause", pacing["batch_pause"]))))
                pacing["request_timeout"] = max(3.0, min(120.0, float(data.get("request_timeout", pacing["request_timeout"]))))
        except Exception:
            return pacing

        return pacing

    def _save_fetch_pacing(
        self,
        request_delay: float,
        batch_size: int,
        batch_pause: float,
        request_timeout: float,
    ) -> None:
        path = self._fetch_pacing_config_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(
                json.dumps(
                    {
                        "request_delay": round(float(request_delay), 2),
                        "batch_size": int(batch_size),
                        "batch_pause": round(float(batch_pause), 2),
                        "request_timeout": round(float(request_timeout), 2),
                    },
                    indent=2,
                ),
                encoding="utf-8",
            )
        except Exception:
            pass

    def _prepare_results_table(self, table: QTableWidget) -> None:
        columns = ["Mod", "Status", "Current", "Latest", "File", "Info", "Action", "Page", "Folder", "Edit", "Purge"]
        table.clear()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels(columns)
        table.setRowCount(0)
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        table._ll_row_data = {}
        self._set_results_column_widths(table)

    def _apply_results_filter(self, table: QTableWidget) -> None:
        mode_widget = getattr(table, "_ll_filter_mode", None)
        text_widget = getattr(table, "_ll_filter_text", None)
        count_widget = getattr(table, "_ll_filter_count", None)
        mode = mode_widget.currentText() if mode_widget is not None else "All links"
        query = text_widget.text().strip().lower() if text_widget is not None else ""
        row_data = getattr(table, "_ll_row_data", {})

        visible = 0
        total = table.rowCount()
        for row_index in range(total):
            row = row_data.get(row_index) or self._row_from_table(table, row_index)
            matches = self._row_matches_filter(row, mode, query)
            table.setRowHidden(row_index, not matches)
            if matches:
                visible += 1

        if count_widget is not None:
            count_widget.setText(f"{visible} / {total}")

    def _row_from_table(self, table: QTableWidget, row_index: int) -> dict:
        def item_text(column: int) -> str:
            item = table.item(row_index, column)
            return item.text() if item is not None else ""

        return {
            "mod": item_text(0),
            "status": item_text(1),
            "current": item_text(2),
            "latest": item_text(3),
            "file": item_text(4),
            "info": item_text(5),
        }

    def _row_matches_filter(self, row: dict, mode: str, query: str) -> bool:
        status = str(row.get("status") or "")
        fixed = bool(row.get("fixed")) or status == "Manual"
        current = str(row.get("current") or "").strip()

        if mode == "Updates" and status != "Update":
            return False
        if mode == "OK" and status != "OK":
            return False
        if mode == "Unknown / missing version" and status != "Unknown" and (current or fixed):
            return False
        if mode == "Manual links" and not fixed:
            return False
        if mode == "Errors / skipped" and status not in {"Error", "Skipped"}:
            return False
        if mode == "Not checked" and status not in {"Ready", "Queued"}:
            return False

        if not query:
            return True

        haystack = " ".join(
            str(row.get(key) or "")
            for key in ("mod", "file", "page_url", "info", "status", "current", "latest", "update_mode")
        ).lower()
        return query in haystack

    def _populate_results_table(self, table: QTableWidget, jobs: list[dict]) -> None:
        cache = getattr(table, "_ll_update_cache", {"entries": {}})
        for index, job in enumerate(jobs):
            job["row_index"] = index
            row = self._pending_row(job, "Ready", "Not checked")
            entry = self._cache_entry_for_job(job, cache)
            if entry:
                cached = self._cached_display_row(job, index, entry)
                if cached:
                    row.update(cached)
            self._append_result_row(table, row)

    def _reset_rows_for_fetch(self, table: QTableWidget, jobs: list[dict]) -> None:
        for index, job in enumerate(jobs):
            job["row_index"] = index
            if index >= table.rowCount():
                self._append_result_row(table, self._pending_row(job, "Queued", "Waiting"))
            else:
                self._set_result_row_values(table, index, self._pending_row(job, "Queued", "Waiting"))
        self._apply_results_filter(table)

    def _pending_row(self, job: dict, status: str, info: str) -> dict:
        update_mode = normalized_update_mode(
            job.get("update_mode"),
            fixed=bool(job.get("fixed")),
        )
        if job.get("fixed"):
            status = "Manual"
            info = "Skip updates; update fetch skipped"
        elif info in {"Not checked", "Waiting"}:
            info = f"{update_mode_label(update_mode)}; {info.lower()}"

        return {
            "row_index": job.get("row_index", 0),
            "ini_path": job.get("ini_path") or "",
            "internal_name": job.get("internal_name") or "",
            "mod_path": job.get("mod_path") or "",
            "mod": job.get("mod") or "",
            "status": status,
            "current": job.get("current") or "",
            "latest": "",
            "file": job.get("file") or "",
            "page_url": job.get("page_url") or "",
            "fixed": bool(job.get("fixed")),
            "update_mode": update_mode,
            "info": info,
        }

    def _set_results_column_widths(self, table: QTableWidget) -> None:
        header = table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(6, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(7, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(8, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(9, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(10, QHeaderView.ResizeMode.Fixed)

        widths = {
            0: 260,
            1: 80,
            2: 80,
            3: 80,
            4: 240,
            6: 82,
            7: 82,
            8: 88,
            9: 82,
            10: 82,
        }
        for column, width in widths.items():
            table.setColumnWidth(column, width)

    def _append_result_row(self, table: QTableWidget, row: dict) -> None:
        row_index = table.rowCount()
        table.insertRow(row_index)
        self._set_result_row_values(table, row_index, row)
        self._configure_action_button(table, row_index, row)

        page_url = row.get("page_url") or ""
        open_button = QPushButton("Open")
        open_button.setEnabled(bool(page_url))
        open_button.clicked.connect(lambda _checked=False, url=page_url: webbrowser.open(url))
        table.setCellWidget(row_index, 7, open_button)

        folder_button = QPushButton("Folder")
        folder_button.setEnabled(bool(row.get("mod_path")))
        folder_button.clicked.connect(
            lambda _checked=False, data=row:
                self._open_mod_folder(data)
        )
        table.setCellWidget(row_index, 8, folder_button)

        edit_button = QPushButton("Edit")
        edit_button.setEnabled(bool(row.get("ini_path")))
        edit_button.clicked.connect(
            lambda _checked=False, index=row_index, data=row:
                self._edit_row_link(table, index, data)
        )
        table.setCellWidget(row_index, 9, edit_button)

        purge_button = QPushButton("Purge")
        purge_button.setEnabled(bool(row.get("ini_path")))
        purge_button.clicked.connect(
            lambda _checked=False, index=row_index, data=row, button=purge_button:
                self._purge_row_ll_ini(table, index, data, button)
        )
        table.setCellWidget(row_index, 10, purge_button)

    def _open_mod_folder(self, row: dict) -> None:
        path = Path(str(row.get("mod_path") or ""))
        if not path.exists():
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Mod folder was not found:\n{path}",
            )
            return

        try:
            webbrowser.open(path.as_uri())
        except ValueError:
            webbrowser.open(str(path))

    def _row_fetch_enabled(self, row: dict) -> bool:
        return (
            bool(row.get("ini_path"))
            and bool(row.get("page_url"))
            and not bool(row.get("fixed"))
            and Path(str(row.get("ini_path") or "")).exists()
        )

    def _row_try_update_enabled(self, row: dict) -> bool:
        return (
            str(row.get("status") or "") == "Update"
            and bool(row.get("ini_path"))
            and bool(row.get("page_url"))
            and not bool(row.get("fixed"))
            and normalized_update_mode(row.get("update_mode")) != UPDATE_MODE_SKIP
            and Path(str(row.get("ini_path") or "")).exists()
        )

    def _configure_action_button(self, table: QTableWidget, row_index: int, row: dict) -> QPushButton:
        button = table.cellWidget(row_index, 6)
        if not isinstance(button, QPushButton):
            button = QPushButton()
            table.setCellWidget(row_index, 6, button)
        try:
            button.clicked.disconnect()
        except Exception:
            pass

        status = str(row.get("status") or "")
        if status == "Downloaded":
            button.setText("Install")
            button.setEnabled(self._downloaded_archive_exists(row))
            button.clicked.connect(
                lambda _checked=False, index=row_index:
                    self._prompt_install_downloaded_update(table, index)
            )
            return button

        if self._row_try_update_enabled(row):
            button.setText("Try Update")
            button.setEnabled(True)
            button.clicked.connect(
                lambda _checked=False, index=row_index, action_button=button:
                    self._start_try_update_worker(table, index, action_button)
            )
            return button

        button.setText("Fetch")
        button.setEnabled(self._row_fetch_enabled(row))
        button.clicked.connect(
            lambda _checked=False, index=row_index, action_button=button:
                self._start_single_check_worker(
                    table,
                    index,
                    self._current_row_data(table, index),
                    action_button,
                )
        )
        return button

    def _current_row_data(self, table: QTableWidget, row_index: int) -> dict:
        row_data = getattr(table, "_ll_row_data", {})
        if row_index in row_data:
            return dict(row_data[row_index])
        return self._row_from_table(table, row_index)

    def _update_result_row(self, table: QTableWidget, row: dict) -> None:
        row_index = int(row.get("row_index", -1))
        if row_index < 0 or row_index >= table.rowCount():
            self._append_result_row(table, row)
            self._cache_result_row(table, row)
            self._apply_results_filter(table)
            return

        display_row = dict(row)
        jobs = getattr(table, "_ll_jobs", [])
        job = jobs[row_index] if 0 <= row_index < len(jobs) else row
        cache = getattr(table, "_ll_update_cache", {"entries": {}})
        entry = self._cache_entry_for_job(job, cache)
        if row.get("status") in {"Error", "Skipped"} and entry and isinstance(entry.get("last_success"), dict):
            success = entry["last_success"]
            display_row["latest"] = display_row.get("latest") or success.get("latest") or ""
            display_row["file"] = display_row.get("file") or success.get("file") or display_row.get("file") or ""
            display_row["info"] = (
                f"{display_row.get('info') or display_row.get('status')}; "
                f"last OK {success.get('checked_at') or 'unknown'}"
            )

        if display_row.get("status") == "Update":
            display_row = self._mark_update_as_downloaded_if_archive_exists(display_row)

        self._set_result_row_values(table, row_index, display_row)
        self._cache_result_row(table, display_row)
        self._maybe_prompt_assisted_install(table, row_index, display_row)
        self._apply_results_filter(table)

    def _set_result_row_values(self, table: QTableWidget, row_index: int, row: dict) -> None:
        row_data = getattr(table, "_ll_row_data", {})
        row_data[row_index] = dict(row)
        table._ll_row_data = row_data

        values = [
            row.get("mod") or "",
            row.get("status") or "",
            row.get("current") or "",
            row.get("latest") or "",
            row.get("file") or "",
            row.get("info") or row.get("error") or "",
        ]
        for column_index, value in enumerate(values):
            item = table.item(row_index, column_index)
            if item is None:
                item = QTableWidgetItem()
                table.setItem(row_index, column_index, item)

            item.setText(str(value))
            if column_index == 1:
                self._style_status_item(item, row.get("status") or "")
        if table.columnCount() > 6:
            self._configure_action_button(table, row_index, row)

    def _maybe_prompt_assisted_install(self, table: QTableWidget, row_index: int, row: dict) -> None:
        if not row.get("downloaded_now"):
            return
        if str(row.get("status") or "") != "Downloaded":
            return
        if normalized_update_mode(row.get("update_mode")) != UPDATE_MODE_ASSISTED:
            return
        self._prompt_install_downloaded_update(table, row_index, row)

    def _prompt_install_downloaded_update(
        self,
        table: QTableWidget,
        row_index: int,
        row: dict | None = None,
    ) -> None:
        row = dict(row or self._current_row_data(table, row_index))
        if not self._organizer or not hasattr(self._organizer, "installMod"):
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "The update archive was downloaded, but this MO2 version does not expose installMod to tools.",
            )
            return

        archive_path = Path(str(row.get("archive_path") or ""))
        if not archive_path.exists():
            archive_path = self._existing_latest_archive(row) or archive_path
        if not archive_path.exists():
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                f"The update archive was downloaded but could not be found:\n{archive_path}",
            )
            return

        mod_name = row.get("mod") or row.get("internal_name") or archive_path.stem
        message = QMessageBox(self._parentWidget())
        message.setWindowTitle("LL Integration Assisted Install")
        message.setIcon(QMessageBox.Icon.Question)
        message.setText(f"Install downloaded update for:\n{mod_name}")
        message.setInformativeText(
            f"Current: {row.get('current') or '?'}\n"
            f"Latest: {row.get('latest') or '?'}\n"
            f"Archive: {archive_path.name}\n"
            f"Target mod folder: {row.get('mod_path') or '?'}\n\n"
            "MO2 will handle the install/replace flow."
        )
        install_button = message.addButton("Install / Replace in MO2", QMessageBox.ButtonRole.AcceptRole)
        keep_button = message.addButton("Keep Downloaded", QMessageBox.ButtonRole.RejectRole)
        message.setDefaultButton(install_button)
        message.exec()

        if message.clickedButton() != install_button:
            self._mark_assisted_prompt_seen(table, row_index, row, "Downloaded; waiting for manual install")
            return

        try:
            installed_mod = self._organizer.installMod(str(archive_path), str(mod_name))
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"MO2 install failed:\n\n{exc}",
            )
            self._mark_assisted_prompt_seen(table, row_index, row, "Downloaded; MO2 install failed")
            return

        if installed_mod is None:
            self._mark_assisted_prompt_seen(table, row_index, row, "Downloaded; MO2 install cancelled")
            return

        sidecar = Path(str(row.get("sidecar_path") or f"{archive_path}.ll.ini"))
        try:
            if sidecar.exists():
                target = write_mod_ll_metadata_from_file(installed_mod, sidecar)
                write_mod_general_source_metadata(installed_mod, row.get("page_url") or "", row.get("latest") or "")
                self._organizer.modDataChanged(installed_mod)
        except Exception as exc:
            QMessageBox.warning(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Installed through MO2, but metadata refresh failed:\n\n{exc}",
            )

        updated = dict(row)
        updated["downloaded_now"] = False
        updated["status"] = "Ready"
        updated["current"] = row.get("latest") or row.get("current") or ""
        updated["latest"] = ""
        updated["info"] = "Assisted install completed; not checked"
        self._set_result_row_values(table, row_index, updated)
        self._update_backing_job(table, row_index, updated)
        self._remove_cached_result(table, updated)

    def _mark_assisted_prompt_seen(self, table: QTableWidget, row_index: int, row: dict, info: str) -> None:
        updated = dict(row)
        updated["downloaded_now"] = False
        updated["info"] = info
        self._set_result_row_values(table, row_index, updated)

    def _edit_row_link(self, table: QTableWidget, row_index: int, row: dict) -> None:
        ini_path = Path(str(row.get("ini_path") or ""))
        if not ini_path.exists():
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                f"LL Integration metadata is missing:\n{ini_path}",
            )
            return

        config = configparser.ConfigParser(interpolation=None)
        config.optionxform = str
        config.read(ini_path, encoding="utf-8")
        if LL_SECTION not in config:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Invalid LL Integration metadata:\n{ini_path}",
            )
            return

        ll = config[LL_SECTION]
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("Edit LoversLab Link")
        dialog.resize(700, 240)

        page_url = QLineEdit(ll.get("page_url", "").strip(), dialog)
        version = QLineEdit(ll.get("version", "").strip(), dialog)
        file_pattern = QLineEdit(
            ll.get("file_pattern", "").strip()
            or ll.get("file_name", "").strip()
            or ll.get("archive_name", "").strip(),
            dialog,
        )
        existing_fixed = (
            truthy(ll.get("fixed_version"))
            or truthy(ll.get("manual_update"))
            or truthy(ll.get("skip_update_check"))
        )
        update_mode = QComboBox(dialog)
        configure_update_mode_combo(
            update_mode,
            normalized_update_mode(ll.get("update_mode"), fixed=existing_fixed),
        )
        try_pattern = QPushButton("Try Pattern", dialog)
        pattern_help = QPushButton("Pattern Help", dialog)

        layout = QGridLayout(dialog)
        layout.addWidget(QLabel(f"Mod: {row.get('mod') or ''}"), 0, 0, 1, 2)
        layout.addWidget(QLabel("LoversLab page URL"), 1, 0)
        layout.addWidget(page_url, 1, 1)
        layout.addWidget(QLabel("Current version"), 2, 0)
        layout.addWidget(version, 2, 1)
        layout.addWidget(QLabel("File pattern"), 3, 0)
        layout.addWidget(file_pattern, 3, 1)
        layout.addWidget(QLabel("Update mode"), 4, 0)
        layout.addWidget(update_mode, 4, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        action_row = QHBoxLayout()
        action_row.addWidget(try_pattern)
        action_row.addWidget(pattern_help)
        action_row.addStretch(1)
        action_row.addWidget(buttons)
        layout.addLayout(action_row, 5, 0, 1, 2)

        try_pattern.clicked.connect(
            lambda _checked=False:
                self._try_link_pattern(
                    table,
                    row_index,
                    row,
                    config,
                    page_url,
                    version,
                    file_pattern,
                )
        )
        pattern_help.clicked.connect(lambda _checked=False: self._show_pattern_help())

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        new_url = page_url.text().strip()
        new_version = version.text().strip()
        new_pattern = file_pattern.text().strip()
        new_update_mode = normalized_update_mode(str(update_mode.currentData() or ""))
        new_fixed = new_update_mode == UPDATE_MODE_SKIP
        if not new_url:
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, "LoversLab page URL is required.")
            return

        ll["page_url"] = new_url
        ll["version"] = new_version
        ll["file_pattern"] = new_pattern
        ll["update_mode"] = new_update_mode
        ll["fixed_version"] = "true" if new_fixed else "false"
        ll["manual_update"] = "true" if new_fixed else "false"
        ll["skip_update_check"] = "true" if new_fixed else "false"
        if new_pattern and not ll.get("file_name", "").strip():
            ll["file_name"] = new_pattern

        with ini_path.open("w", encoding="utf-8") as file:
            config.write(file, space_around_delimiters=False)

        row["page_url"] = new_url
        row["current"] = new_version
        row["file"] = new_pattern
        row["fixed"] = new_fixed
        row["update_mode"] = new_update_mode
        row["status"] = "Manual" if new_fixed else "Ready"
        row["latest"] = ""
        row["info"] = "Updates skipped" if new_fixed else f"Mode: {update_mode_label(new_update_mode)}; not checked"
        self._set_result_row_values(table, row_index, row)
        self._update_backing_job(table, row_index, row)

        self._configure_action_button(table, row_index, row)

        open_button = QPushButton("Open")
        open_button.setEnabled(bool(new_url))
        open_button.clicked.connect(lambda _checked=False, url=new_url: webbrowser.open(url))
        table.setCellWidget(row_index, 7, open_button)
        self._apply_results_filter(table)

    def _update_backing_job(self, table: QTableWidget, row_index: int, row: dict) -> None:
        jobs = getattr(table, "_ll_jobs", [])
        if row_index < 0 or row_index >= len(jobs):
            return

        jobs[row_index].update({
            "page_url": row.get("page_url") or "",
            "current": row.get("current") or "",
            "file": row.get("file") or "",
            "fixed": bool(row.get("fixed")),
            "update_mode": row.get("update_mode") or "",
        })

    def _try_link_pattern(
        self,
        table: QTableWidget,
        row_index: int,
        row: dict,
        config: configparser.ConfigParser,
        page_url_widget: QLineEdit,
        version_widget: QLineEdit,
        pattern_widget: QLineEdit,
    ) -> None:
        page_url = page_url_widget.text().strip()
        pattern = pattern_widget.text().strip()
        if not page_url:
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, "LoversLab page URL is required.")
            return
        if not pattern:
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, "File pattern is required.")
            return

        try:
            cookies_path = self._setting_path("cookies_path", DEFAULT_COOKIES)
            timeout = self._load_fetch_pacing()["request_timeout"]
            fetch_url = with_query_value(page_url, "do", "download")
            html = fetch_ll_html(fetch_url, cookies_path, referer=page_url, timeout=timeout)
            downloads = extract_downloads(html)
            match = choose_latest(downloads, pattern)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Pattern test failed:\n\n{exc}",
            )
            return

        if not match:
            sample = "\n".join(download.name for download in downloads[:8])
            QMessageBox.warning(
                self._parentWidget(),
                PLUGIN_NAME,
                f"No matching download found for:\n{pattern}\n\n"
                f"Downloads seen: {len(downloads)}\n\n{sample}",
            )
            return

        message = QMessageBox(self._parentWidget())
        message.setWindowTitle(PLUGIN_NAME)
        message.setIcon(QMessageBox.Icon.Information)
        message.setText("Pattern matched.")
        message.setInformativeText(
            f"File: {match.name}\n"
            f"Detected version: {match.version or '<unknown>'}\n"
            f"Pattern: {pattern}\n"
            f"Size: {match.size or '<unknown>'}\n\n"
            f"Downloads seen: {len(downloads)}"
        )
        use_button = message.addButton("Use Match", QMessageBox.ButtonRole.AcceptRole)
        cancel_button = message.addButton(QMessageBox.StandardButton.Cancel)
        message.setDefaultButton(use_button)
        message.exec()
        if message.clickedButton() != use_button:
            return

        self._apply_pattern_match(
            table,
            row_index,
            row,
            config,
            page_url,
            pattern,
            match,
            version_widget,
            pattern_widget,
        )

    def _apply_pattern_match(
        self,
        table: QTableWidget,
        row_index: int,
        row: dict,
        config: configparser.ConfigParser,
        page_url: str,
        pattern: str,
        match,
        version_widget: QLineEdit,
        pattern_widget: QLineEdit,
    ) -> None:
        if LL_SECTION not in config:
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, "LoversLab metadata section is missing.")
            return

        detected_version = match.version or ""
        if not detected_version:
            QMessageBox.warning(self._parentWidget(), PLUGIN_NAME, "No version was detected from the matched file.")
            return

        ini_path = Path(str(row.get("ini_path") or ""))
        if not ini_path.exists():
            QMessageBox.critical(self._parentWidget(), PLUGIN_NAME, f"Metadata file is missing:\n{ini_path}")
            return

        ll = config[LL_SECTION]
        ll["page_url"] = page_url
        ll["version"] = detected_version
        ll["file_pattern"] = pattern
        ll["file_name"] = match.name
        ll["archive_name"] = match.name
        ll["original_archive_name"] = match.name
        if match.size:
            ll["size"] = match.size
        if match.date_iso:
            ll["date_iso"] = match.date_iso
        if match.url:
            ll["download_url"] = match.url

        with ini_path.open("w", encoding="utf-8") as file:
            config.write(file, space_around_delimiters=False)

        version_widget.setText(detected_version)
        pattern_widget.setText(pattern)

        updated = dict(row)
        updated["page_url"] = page_url
        updated["current"] = detected_version
        updated["file"] = match.name
        updated["latest"] = ""
        updated["status"] = "Ready"
        updated["info"] = "Pattern match applied; not checked"
        updated["fixed"] = bool(row.get("fixed"))
        updated["update_mode"] = row.get("update_mode") or UPDATE_MODE_MANUAL
        self._set_result_row_values(table, row_index, updated)
        self._update_backing_job(table, row_index, updated)
        self._remove_cached_result(table, updated)

    def _show_pattern_help(self) -> None:
        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            "File pattern examples:\n\n"
            "* matches any text.\n"
            "{version} marks where the version is in the file name.\n\n"
            "ModName_0.45Beta5.7z\n"
            "Pattern: ModName_{version}.7z\n"
            "Detected version: 0.45.5\n\n"
            "ModName0.13_SE.zip\n"
            "Pattern: ModName{version}_SE.zip\n"
            "Detected version: 0.13\n\n"
            "ModName0.80beta8 - SE.7z\n"
            "Pattern: ModName{version}beta{version} - SE.7z\n"
            "Detected version: 0.80.8",
        )

    def _purge_row_ll_ini(
        self,
        table: QTableWidget,
        row_index: int,
        row: dict,
        purge_button: QPushButton,
    ) -> None:
        ini_path = Path(str(row.get("ini_path") or ""))
        internal_name = row.get("internal_name")
        mod = None
        if self._organizer and internal_name:
            mod = self._organizer.modList().getMod(str(internal_name))

        if not ini_path.exists() and mod is None:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                f"LL Integration metadata is already gone:\n{ini_path}",
            )
            purge_button.setEnabled(False)
            return

        result = QMessageBox.question(
            self._parentWidget(),
            PLUGIN_NAME,
            f"Purge LoversLab link for this mod?\n\n{row.get('mod')}\n\n"
            "This removes the [LoversLab] metadata section from meta.ini "
            "and deletes any legacy LL.ini file.",
        )
        if result != QMessageBox.StandardButton.Yes:
            return

        actions = []
        if mod is not None:
            actions.extend(remove_mod_ll_metadata(mod))
        elif ini_path.exists() and ini_path.name == LEGACY_MOD_LL_FILE:
            ini_path.unlink()
            actions.append(f"removed legacy {LEGACY_MOD_LL_FILE}")

        purge_button.setEnabled(False)
        self._remove_cached_result(table, row)
        fetch_button = table.cellWidget(row_index, 6)
        if fetch_button is not None:
            fetch_button.setEnabled(False)

        meta_cleaned = False
        if mod is not None:
            meta_cleaned = cleanup_loverslab_meta(mod)
            if meta_cleaned:
                actions.append("cleaned meta.ini URL")

        status_item = table.item(row_index, 1)
        if status_item is not None:
            status_item.setText("Purged")
            self._style_status_item(status_item, "Purged")

        error_item = table.item(row_index, 5)
        if error_item is not None:
            error_item.setText("; ".join(actions) if actions else "nothing changed")

        row["status"] = "Purged"
        row["info"] = "; ".join(actions) if actions else "nothing changed"
        self._set_result_row_values(table, row_index, row)
        self._apply_results_filter(table)

    def _style_status_item(self, item: QTableWidgetItem, status: str) -> None:
        colors = {
            "OK": QColor(44, 140, 68),
            "Update": QColor(180, 130, 0),
            "Manual": QColor(105, 150, 190),
            "Downloaded": QColor(44, 140, 68),
            "Ready": QColor(145, 145, 145),
            "Queued": QColor(145, 145, 145),
            "Untracked": QColor(145, 145, 145),
            "Unknown": QColor(145, 145, 145),
            "Purged": QColor(145, 145, 145),
            "Skipped": QColor(145, 145, 145),
            "Error": QColor(190, 45, 45),
        }
        item.setForeground(colors.get(status, QColor(220, 220, 220)))
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)


class LoversLabMenuTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Menu"
    TOOL_DISPLAY = "LL Integration"
    TOOL_DESCRIPTION = "Opens LL Integration tools."

    def init(self, organizer: mobase.IOrganizer) -> bool:
        super().init(organizer)
        if self._experimental_toolbar_enabled():
            self._experimental_log("menu tool init; toolbar enabled")
            self._install_experimental_toolbar_button()
        return True

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def _install_experimental_toolbar_button(self) -> None:
        try:
            from .experimental.toolbar import install_toolbar_button

            install_toolbar_button(
                self._parentWidget,
                Path(__file__).resolve().parent / "icons" / "ll_check_all.svg",
                self.display,
                self._experimental_log,
            )
        except Exception as exc:
            self._experimental_log(f"toolbar experiment: install failed: {type(exc).__name__}: {exc}")

    def _experimental_toolbar_enabled(self) -> bool:
        return truthy(str(self._configured_paths().get("experimental_toolbar", "")))

    def _experimental_log(self, message: str) -> None:
        line = f"{time.strftime('%Y-%m-%dT%H:%M:%SZ', time.gmtime())} {message}\n"
        paths = [Path(__file__).resolve().parent / "experimental_toolbar.log"]
        if self._organizer:
            try:
                paths.insert(
                    0,
                    Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "experimental_toolbar.log",
                )
            except Exception:
                pass

        for log_path in paths:
            try:
                log_path.parent.mkdir(parents=True, exist_ok=True)
                with log_path.open("a", encoding="utf-8") as file:
                    file.write(line)
            except Exception:
                pass

    def display(self) -> None:
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("LL Integration")
        dialog.resize(430, 300)

        title = QLabel("LL Integration")
        title.setStyleSheet("font-weight: 700; font-size: 15px;")
        subtitle = QLabel("Manage LoversLab links, update checks, and integration paths.")
        subtitle.setWordWrap(True)

        actions = [
            (
                "Manage LoversLab Links",
                "Open the mod list, page links, purge buttons, and optional update fetch.",
                LoversLabCheckAllTool,
            ),
            (
                "Find Voice Packs",
                "Scan LoversLab mods and detect installed or missing voice packs.",
                LoversLabVoiceFinderTool,
            ),
            (
                "Create Manual Link",
                "Create source metadata for multipart archives or manually installed mods.",
                LoversLabCreateLinkTool,
            ),
            (
                "Purge Suspicious Links",
                "Clean accidental LoversLab links from Nexus-identified mods.",
                LoversLabPurgeSuspiciousLinksTool,
            ),
            (
                "Integration Paths",
                "Show native bridge, cookies, metadata, and MO2 plugin paths.",
                LoversLabPathsTool,
            ),
        ]

        layout = QVBoxLayout(dialog)
        layout.addWidget(title)
        layout.addWidget(subtitle)

        for label, tooltip, tool_class in actions:
            button = QPushButton(label)
            button.setToolTip(tooltip)
            button.clicked.connect(
                lambda _checked=False, current_tool=tool_class:
                    self._open_child_tool(dialog, current_tool)
            )
            layout.addWidget(button)

        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close)
        dialog.setLayout(layout)
        dialog.exec()

    def _open_child_tool(self, dialog: QDialog, tool_class: type[LoversLabBaseTool]) -> None:
        dialog.accept()
        tool = tool_class()
        tool.init(self._organizer)
        tool.display()


class VoiceSourceFetchWorker(QObject):
    candidatesReady = pyqtSignal(object)
    downloadsReady = pyqtSignal(object)
    statusChanged = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        rows: list[dict],
        source_urls: list[str],
        cookies_path: Path,
        false_matches: list[dict],
        nexus_api_key: str = "",
        timeout: float = UPDATE_REQUEST_TIMEOUT_SECONDS,
    ) -> None:
        super().__init__()
        self._rows = [dict(row) for row in rows]
        self._source_urls = list(source_urls)
        self._cookies_path = cookies_path
        self._false_matches = list(false_matches)
        self._nexus_api_key = str(nexus_api_key or "").strip()
        self._timeout = timeout
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            candidates = []
            all_downloads = []
            total = len(self._source_urls)
            for index, source_url in enumerate(self._source_urls, start=1):
                if self._cancelled:
                    self.finished.emit(False, "Cancelled")
                    return

                self.statusChanged.emit(f"Fetching voice source {index} / {total}: {source_url}")
                if is_nexus_source_url(source_url):
                    parts = nexus_source_parts(source_url)
                    if self._nexus_api_key and parts:
                        download_candidates = fetch_nexus_files_api(parts[0], parts[1], self._nexus_api_key, timeout=self._timeout)
                    else:
                        html = fetch_nexus_html(source_url, timeout=self._timeout)
                        download_candidates = extract_nexus_downloads(html, source_url)
                else:
                    html = fetch_ll_html(source_url, self._cookies_path, timeout=self._timeout)
                    download_candidates = []
                    for download in extract_downloads(html):
                        download_candidates.append({
                            "source_url": with_query_value(source_url, "do", "download"),
                            "source_title": "Voice pack source",
                            "download_name": download.name,
                            "download_url": download.url,
                            "voice_category": voice_category_guess(download.name),
                            "size": download.size or "",
                            "date_iso": download.date_iso or "",
                            "version": download.version or "",
                            "source_type": "loverslab",
                        })

                for download_candidate in download_candidates:
                    all_downloads.append(dict(download_candidate))
                    best = self._best_base_match(download_candidate)
                    if best:
                        candidates.append(best)

                self.candidatesReady.emit(candidates)
                self.downloadsReady.emit(all_downloads)
                time.sleep(UPDATE_REQUEST_DELAY_SECONDS)

            self.candidatesReady.emit(candidates)
            self.downloadsReady.emit(all_downloads)
            self.finished.emit(True, f"Fetched {total} voice source page(s).")
        except Exception as exc:
            self.finished.emit(False, str(exc))

    def _best_base_match(self, candidate: dict) -> dict | None:
        best = None
        best_score = 0
        for row in self._rows:
            if row.get("status") == "Ignored":
                continue
            if self._is_false_match(row, candidate):
                continue

            score = voice_match_score(row.get("base_mod") or "", candidate.get("download_name") or "")
            if candidate.get("voice_category") == row.get("voice_category"):
                score += 20
            elif row.get("voice_category") and candidate.get("voice_category"):
                score = 0
            else:
                score = max(0, score - 35)
            if score > best_score:
                best_score = score
                best = row

        if not best or best_score < 55:
            return None

        result = dict(candidate)
        result.update({
            "base_mod": best.get("base_mod") or "",
            "base_internal_name": best.get("base_internal_name") or "",
            "base_page_url": best.get("base_page_url") or "",
            "voice_category": best.get("voice_category") or "",
            "voice_category_label": best.get("voice_category_label") or "",
            "online_score": best_score,
        })
        return result

    def _is_false_match(self, row: dict, candidate: dict) -> bool:
        base_key = str(row.get("base_internal_name") or row.get("base_mod") or "").lower()
        candidate_key = str(candidate.get("download_name") or "").lower()
        source_url = str(candidate.get("source_url") or "").lower()
        category = str(row.get("voice_category") or "").lower()
        for item in self._false_matches:
            if str(item.get("base") or "").lower() != base_key:
                continue
            item_category = str(item.get("voice_category") or "").lower()
            if item_category and category and item_category != category:
                continue
            if str(item.get("candidate") or "").lower() != candidate_key:
                continue
            if str(item.get("source_url") or "").lower() in ("", source_url):
                return True
        return False


class VoiceCandidateDownloadWorker(QObject):
    statusChanged = pyqtSignal(str)
    finished = pyqtSignal(bool, str, object)

    def __init__(
        self,
        candidate: dict,
        downloads_path: Path,
        cookies_path: Path,
        timeout: float = 60.0,
    ) -> None:
        super().__init__()
        self._candidate = dict(candidate)
        self._downloads_path = downloads_path
        self._cookies_path = cookies_path
        self._timeout = timeout

    def run(self) -> None:
        try:
            download_url = str(self._candidate.get("download_url") or "")
            source_url = str(self._candidate.get("source_url") or "")
            if not download_url:
                raise RuntimeError("Candidate download URL is missing.")
            if not self._downloads_path:
                raise RuntimeError("MO2 downloads path is not available.")

            archive_name = safe_archive_name(self._candidate.get("download_name") or "voice-pack.7z")
            archive_path = self._downloads_path / archive_name
            self.statusChanged.emit(f"Downloading voice candidate: {archive_name}")
            already_exists = archive_path.exists()
            if not already_exists:
                download_loverslab_archive(
                    urljoin(source_url, download_url),
                    archive_path,
                    self._cookies_path,
                    referer=source_url,
                    timeout=self._timeout,
                )

            sidecar = write_voice_download_sidecar(
                archive_path,
                self._candidate,
                urljoin(source_url, download_url),
            )
            result = {
                "archive_path": str(archive_path),
                "sidecar_path": str(sidecar),
                "already_exists": already_exists,
            }
            self.finished.emit(True, "Voice candidate downloaded." if not already_exists else "Voice candidate already existed; metadata refreshed.", result)
        except Exception as exc:
            self.finished.emit(False, str(exc), {})


class LoversLabVoiceFinderTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Voice Finder"
    TOOL_DISPLAY = "Find LoversLab Voice Packs"
    TOOL_DESCRIPTION = "Finds installed or likely missing voice packs for LoversLab mods."

    COL_STATUS = 0
    COL_BASE_MOD = 1
    COL_VOICEPACK = 2
    COL_DBVO = 3
    COL_IVDT = 4
    COL_CATEGORY = 5
    COL_INSTALLED_VOICE = 6
    COL_SCORE = 7
    COL_ONLINE = 8
    COL_ONLINE_SCORE = 9
    COL_SOURCE = 10
    COL_PAGE = 11

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def display(self) -> None:
        try:
            config = self._load_voice_config()
            rows = self._collect_voice_rows(config)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Voice scan failed:\n\n{exc}",
            )
            return

        self._show_results(rows)

    def _load_all_fetch_table_widths(self) -> list[int]:
        config = self._load_voice_config()
        widths = config.get("allFetchedDownloadsColumnWidths") or []
        if not isinstance(widths, list):
            return []

        result = []
        for value in widths:
            try:
                result.append(int(value))
            except (TypeError, ValueError):
                pass
        return result


    def _save_all_fetch_table_widths(self, table: QTableWidget) -> None:
        config = self._load_voice_config()
        config["allFetchedDownloadsColumnWidths"] = [
            table.columnWidth(column)
            for column in range(table.columnCount())
        ]
        self._save_voice_config(config)


    def _restore_all_fetch_table_widths(self, table: QTableWidget) -> None:
        widths = self._load_all_fetch_table_widths()
        if not widths:
            return

        for column, width in enumerate(widths):
            if column >= table.columnCount():
                break
            if width > 20:
                table.setColumnWidth(column, width)

    def _load_all_fetch_splitter_sizes(self) -> list[int]:
        config = self._load_voice_config()
        sizes = config.get("allFetchedDownloadsSplitterSizes") or []
        if not isinstance(sizes, list):
            return []

        result = []
        for value in sizes:
            try:
                result.append(int(value))
            except (TypeError, ValueError):
                pass
        return result

    def _save_all_fetch_splitter_sizes(self, splitter: QSplitter) -> None:
        config = self._load_voice_config()
        config["allFetchedDownloadsSplitterSizes"] = [int(value) for value in splitter.sizes()]
        self._save_voice_config(config)


    def _restore_all_fetch_splitter_sizes(self, splitter: QSplitter) -> None:
        sizes = self._load_all_fetch_splitter_sizes()
        if len(sizes) >= 2:
            splitter.setSizes(sizes[:2])

    def _voice_config_path(self) -> Path:
        if self._organizer:
            path = Path(str(self._organizer.pluginDataPath())) / "ll_integration"
            path.mkdir(parents=True, exist_ok=True)
            return path / "voice_finder.json"
        return Path(__file__).resolve().parent / "voice_finder.json"

    def _load_voice_config(self) -> dict:
        path = self._voice_config_path()
        if not path.exists():
            return {
                "version": 1,
                "voiceSourceUrls": [],
                "nexusApiKey": "",
                "falseMatches": [],
                "ignoredBaseMods": [],
                "manualVoiceMatches": {},
                "forcedVoiceMods": [],
                "forcedBaseMods": [],
                "completeVoiceSlots": [],
                "noneVoiceSlots": [],
                "localMatchThreshold": DEFAULT_VOICE_MATCH_THRESHOLD,
                "window": {},
            }

        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            data = {}

        if not isinstance(data, dict):
            data = {}

        return {
            "version": 1,
            "voiceSourceUrls": list(data.get("voiceSourceUrls") or []),
            "nexusApiKey": str(data.get("nexusApiKey") or ""),
            "falseMatches": list(data.get("falseMatches") or []),
            "ignoredBaseMods": list(data.get("ignoredBaseMods") or []),
            "manualVoiceMatches": dict(data.get("manualVoiceMatches") or {}),
            "forcedVoiceMods": list(data.get("forcedVoiceMods") or []),
            "forcedBaseMods": list(data.get("forcedBaseMods") or []),
            "completeVoiceSlots": list(data.get("completeVoiceSlots") or []),
            "noneVoiceSlots": list(data.get("noneVoiceSlots") or []),
            "localMatchThreshold": self._local_match_threshold(data),
            "window": dict(data.get("window") or {}),
        }

    def _local_match_threshold(self, config: dict | None) -> int:
        try:
            return max(0, min(100, int((config or {}).get("localMatchThreshold") or DEFAULT_VOICE_MATCH_THRESHOLD)))
        except (TypeError, ValueError):
            return DEFAULT_VOICE_MATCH_THRESHOLD

    def _save_voice_config(self, config: dict) -> None:
        path = self._voice_config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(config, indent=2), encoding="utf-8")

    def _collect_voice_rows(self, config: dict) -> list[dict]:
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        all_mods = []
        forced_voice_mods = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base_mods = {str(value).lower() for value in config.get("forcedBaseMods", [])}

        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is None:
                continue

            display_name = mod_list.displayName(internal_name)
            key = str(internal_name).lower()
            auto_voice = voice_keyword_present(display_name)
            is_voice = (auto_voice or key in forced_voice_mods) and key not in forced_base_mods
            all_mods.append({
                "internal_name": internal_name,
                "display_name": display_name,
                "mod": mod,
                "is_voice": is_voice,
                "auto_is_voice": auto_voice,
                "classification_override": "voice" if key in forced_voice_mods else "base" if key in forced_base_mods else "auto",
                "ll_metadata": mod_ll_metadata_path(mod, migrate_legacy=False),
            })

        installed_voice_mods = [item for item in all_mods if item["is_voice"]]
        base_mods = [
            item
            for item in all_mods
            if item["ll_metadata"] is not None and not item["is_voice"]
        ]

        rows = []
        threshold = self._local_match_threshold(config)
        complete_slots = {str(value).lower() for value in config.get("completeVoiceSlots", [])}
        none_slots = {str(value).lower() for value in config.get("noneVoiceSlots", [])}
        for base in base_mods:
            page_url = self._ll_page_url(base["ll_metadata"])
            ignored = self._base_key(base) in {str(value).lower() for value in config.get("ignoredBaseMods", [])}

            for category, category_label in VOICE_CATEGORIES:
                slot_key = self._slot_key(base, category)
                voice_candidates = self._installed_voice_candidates(base, installed_voice_mods, config, category)
                manual_voice = self._manual_voice_candidate(config, base, all_mods, category)
                if manual_voice:
                    voice_candidates = [
                        manual_voice,
                        *[
                            candidate
                            for candidate in voice_candidates
                            if str(candidate.get("internal_name") or "").lower()
                            != str(manual_voice.get("internal_name") or "").lower()
                        ],
                    ]
                best_voice = next(
                    (
                        candidate
                        for candidate in voice_candidates
                        if candidate.get("manual") or int(candidate.get("score") or 0) >= threshold
                    ),
                    None,
                ) or {
                    "display_name": "",
                    "internal_name": "",
                    "score": 0,
                    "mod_path": "",
                    "manual": False,
                }

                slot_status = self._voice_row_status(
                        bool(ignored),
                        slot_key in complete_slots,
                        slot_key in none_slots,
                        bool(best_voice.get("manual")),
                        int(best_voice.get("score") or 0),
                        False,
                        threshold,
                    )

                rows.append({
                    "status": slot_status,
                    "slot_status": slot_status,
                    "base_mod": base["display_name"],
                    "base_internal_name": base["internal_name"],
                    "voice_category": category,
                    "voice_category_label": category_label,
                    "slot_key": slot_key,
                    "classification_override": base.get("classification_override") or "auto",
                    "base_page_url": page_url,
                    "installed_voice": best_voice["display_name"],
                    "installed_voice_internal_name": best_voice["internal_name"],
                    "score": best_voice["score"],
                    "manual_voice": bool(best_voice.get("manual")),
                    "complete_voice": slot_key in complete_slots,
                    "none_voice": slot_key in none_slots,
                    "local_match_threshold": threshold,
                    "installed_voice_candidates": voice_candidates,
                    "online_candidate": "",
                    "online_download_url": "",
                    "online_source_url": "",
                    "online_score": 0,
                    "online_size": "",
                    "online_date_iso": "",
                    "online_version": "",
                    "online_candidates": [],
                    "search_query": f"{voice_search_query(base['display_name'])} {category_label}",
                })

        self._attach_voice_overview(rows)

        category_order = {category: index for index, (category, _label) in enumerate(VOICE_CATEGORIES)}
        rows.sort(key=lambda row: (
            row["base_mod"].lower(),
            category_order.get(str(row.get("voice_category") or ""), 99),
            {"Missing": 0, "Online found": 1, "Possible": 2, "Installed": 3, "Complete": 4, "None": 5, "Ignored": 6}.get(row.get("slot_status") or row["status"], 9),
        ))
        return rows

    def _attach_voice_overview(self, rows: list[dict]) -> None:
        by_base: dict[str, dict[str, dict]] = {}
        for row in rows:
            by_base.setdefault(str(row.get("base_internal_name") or row.get("base_mod") or ""), {})[
                str(row.get("voice_category") or "")
            ] = row

        for row in rows:
            siblings = by_base.get(str(row.get("base_internal_name") or row.get("base_mod") or ""), {})
            primary = self._primary_voice_row(siblings)
            row["voice_overview"] = {
                category: self._voice_overview_text(siblings.get(category))
                for category, _label in VOICE_CATEGORIES
            }
            row["status"] = self._aggregate_voice_status(list(siblings.values()))
            row["_overview_duplicate"] = row is not primary

    def _primary_voice_row(self, siblings: dict[str, dict]) -> dict | None:
        for category, _label in VOICE_CATEGORIES:
            row = siblings.get(category)
            if row:
                return row
        return next(iter(siblings.values()), None)

    def _aggregate_voice_status(self, rows: list[dict]) -> str:
        statuses = [str(row.get("slot_status") or row.get("status") or "") for row in rows]
        if statuses and all(status == "Ignored" for status in statuses):
            return "Ignored"
        if any(
            status == "Installed"
            or (status == "Complete" and str(row.get("installed_voice") or "").strip())
            for status, row in zip(statuses, rows)
        ):
            return "Installed"
        if any(status in ("Possible", "Online found") for status in statuses):
            return "Possible"
        if statuses and all(status in ("Complete", "None", "Ignored") for status in statuses):
            return "Complete"
        if any(status == "Missing" for status in statuses):
            return "Missing"
        if any(status == "Complete" for status in statuses):
            return "Complete"
        if any(status == "None" for status in statuses):
            return "None"
        return "Missing"

    def _voice_overview_text(self, row: dict | None) -> str:
        if not row:
            return ""
        status = str(row.get("slot_status") or row.get("status") or "")
        candidate = str(row.get("installed_voice") or "").strip()
        if status == "Complete" and candidate:
            return f"OK: {candidate}"
        if status in ("Missing", "Ignored", "Complete", "None"):
            return status
        if status == "Online found":
            online_candidate = str(row.get("online_candidate") or "").strip()
            return f"Online: {online_candidate}" if online_candidate else "Online found"
        if candidate:
            prefix = "Manual" if row.get("manual_voice") else status
            return f"{prefix}: {candidate}"
        return status

    def _base_key(self, base: dict) -> str:
        return str(base.get("internal_name") or base.get("base_internal_name") or base.get("display_name") or base.get("base_mod") or "").lower()

    def _slot_key(self, base: dict, category: str) -> str:
        return f"{self._base_key(base)}|{category}"

    def _manual_match_key(self, base: dict, category: str) -> str:
        return self._slot_key(base, category)

    def _manual_voice_candidate(self, config: dict, base: dict, all_mods: list[dict], category: str) -> dict | None:
        matches = config.get("manualVoiceMatches")
        if not isinstance(matches, dict):
            return None

        entry = matches.get(self._manual_match_key(base, category))
        if not isinstance(entry, dict) and category == "player":
            entry = matches.get(self._base_key(base))
        if not isinstance(entry, dict):
            return None

        internal_name = str(entry.get("internal_name") or "").lower()
        if not internal_name:
            return None

        for mod in all_mods:
            if str(mod.get("internal_name") or "").lower() != internal_name:
                continue
            if str(mod.get("internal_name") or "") == str(base.get("internal_name") or ""):
                return None
            return {
                "display_name": mod.get("display_name") or entry.get("display_name") or "",
                "internal_name": mod.get("internal_name") or entry.get("internal_name") or "",
                "mod_path": str(mod_root_path(mod["mod"])),
                "score": 1000,
                "manual": True,
                "voice_category": category,
            }

        return {
            "display_name": entry.get("display_name") or entry.get("internal_name") or "",
            "internal_name": entry.get("internal_name") or "",
            "mod_path": entry.get("mod_path") or "",
            "score": 1000,
            "manual": True,
            "voice_category": category,
        }

    def _installed_voice_candidates(self, base: dict, voice_mods: list[dict], config: dict, category: str) -> list[dict]:
        candidates = []
        for voice in voice_mods:
            if voice["internal_name"] == base["internal_name"]:
                continue
            if self._is_false_match(config, base, voice["display_name"], "", category):
                continue

            score = voice_match_score(base["display_name"], voice["display_name"])
            guessed_category = voice_category_guess(voice["display_name"])
            if guessed_category == category:
                score += 20
            elif guessed_category:
                score = 0
            else:
                score = max(0, score - 35)
            candidates.append({
                "display_name": voice["display_name"],
                "internal_name": voice["internal_name"],
                "mod_path": str(mod_root_path(voice["mod"])),
                "score": score,
                "voice_category": guessed_category,
            })

        return sorted(
            candidates,
            key=lambda item: (-int(item.get("score") or 0), str(item.get("display_name") or "").lower()),
        )

    def _is_false_match(
        self,
        config: dict,
        base: dict,
        candidate: str,
        source_url: str,
        voice_category: str = "",
    ) -> bool:
        base_key = self._base_key(base)
        category_key = str(voice_category or base.get("voice_category") or "").lower()
        candidate_key = str(candidate or "").lower()
        source_key = str(source_url or "").lower()
        for item in config.get("falseMatches", []):
            if str(item.get("base") or "").lower() != base_key:
                continue
            item_category = str(item.get("voice_category") or "").lower()
            if item_category and category_key and item_category != category_key:
                continue
            if str(item.get("candidate") or "").lower() != candidate_key:
                continue
            if str(item.get("source_url") or "").lower() in ("", source_key):
                return True
        return False

    def _voice_row_status(
        self,
        ignored: bool,
        complete: bool,
        none: bool,
        manual: bool,
        score: int,
        has_online: bool,
        threshold: int | None = None,
    ) -> str:
        threshold_value = threshold if threshold is not None else DEFAULT_VOICE_MATCH_THRESHOLD
        if ignored:
            return "Ignored"
        if none:
            return "None"
        if complete:
            return "Complete"
        if manual or score >= 90:
            return "Installed"
        if score >= threshold_value:
            return "Possible"
        if has_online:
            return "Online found"
        return "Missing"

    def _ll_page_url(self, ini_path: Path | None) -> str:
        if ini_path is None:
            return ""

        try:
            ll = read_ll_section(ini_path)
        except Exception:
            return ""

        return ll.get("page_url", "").strip()

    def _show_results(self, rows: list[dict]) -> None:
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("LL Integration - Voice Finder")
        dialog.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        window_config = self._voice_window_config()
        dialog.resize(
            int(window_config.get("width") or 1180),
            int(window_config.get("height") or 620),
        )
        dialog.setMinimumSize(980, 480)

        table = QTableWidget(dialog)
        table.setColumnCount(12)
        table.setHorizontalHeaderLabels([
            "Status",
            "Base mod",
            "Voice Pack",
            "DBVO",
            "IVDT",
            "Voice type",
            "Installed voice candidate",
            "Score",
            "Online candidate",
            "Online score",
            "Source",
            "LL page",
        ])
        table.setRowCount(len(rows))
        display_rows = self._build_voice_display_rows(rows)

        table.setRowCount(len(display_rows))
        table._ll_voice_rows = rows
        table._ll_voice_display_rows = display_rows
        table._ll_voice_sort_column = -1
        table._ll_voice_sort_ascending = True
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

        for row_index, display_row in enumerate(display_rows):
            self._fill_display_table_row(table, row_index, display_row)

        header = table.horizontalHeader()
        header.setSectionResizeMode(self.COL_STATUS, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_BASE_MOD, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_VOICEPACK, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_DBVO, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_IVDT, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_CATEGORY, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_INSTALLED_VOICE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_SCORE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_ONLINE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_ONLINE_SCORE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(self.COL_SOURCE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(self.COL_PAGE, QHeaderView.ResizeMode.Stretch)

        filter_mode = QComboBox(dialog)
        filter_mode.addItems([
            "All",
            "Missing",
            "Online found",
            "Possible",
            "Installed",
            "Complete",
            "None",
            "Ignored",
        ])

        filter_text = QLineEdit(dialog)
        filter_text.setPlaceholderText("Search mod / voice / query")
        filter_text.setClearButtonEnabled(True)

        count_label = QLabel(dialog)
        selected_label = QLabel("Selected: none")
        selected_label.setWordWrap(True)
        threshold_spin = QSpinBox(dialog)
        threshold_spin.setRange(0, 100)
        threshold_spin.setValue(self._local_match_threshold(self._load_voice_config()))
        threshold_spin.setToolTip("Minimum local score required before an installed voice candidate is shown as a match.")

        open_page = QPushButton("Open LL page")
        source_urls = QPushButton("Voice source URLs")
        nexus_link = QPushButton("Handle Nexus Link")
        fetch_sources = QPushButton("Fetch sources")
        download_candidate = QPushButton("Choose / Download")
        all_downloads = QPushButton("Show All Source Downloads")
        false_match = QPushButton("False local match")
        manage_false_matches = QPushButton("False matches")
        complete_slot = QPushButton("Complete / Reopen")
        ignore_mod = QPushButton("Ignore / Unignore")
        classify_mod = QPushButton("Classify")
        voice_mods = QPushButton("Voice mods")
        search_web = QPushButton("Web search")
        close = QPushButton("Close")
        progress_label = QLabel("Ready")
        all_downloads.setMinimumHeight(32)
        all_downloads.setStyleSheet(
            "QPushButton { background-color: #214d7a; color: white; font-weight: 700; padding: 6px 14px; }"
            "QPushButton:hover { background-color: #2c669f; }"
            "QPushButton:disabled { background-color: #3a3a3a; color: #aaa; }"
        )
        download_candidate.setStyleSheet(
            "QPushButton { background-color: #1f6f46; color: white; font-weight: 700; }"
            "QPushButton:hover { background-color: #278a58; }"
        )
        fetch_sources.setStyleSheet(
            "QPushButton { background-color: #6a4d1c; color: white; font-weight: 700; }"
            "QPushButton:hover { background-color: #856124; }"
        )

        header.sectionClicked.connect(lambda column: self._sort_voice_table(table, column, filter_mode, filter_text, count_label))
        table.customContextMenuRequested.connect(
            lambda pos: self._show_voice_context_menu(table, pos, filter_mode, filter_text, count_label)
        )
        table.itemSelectionChanged.connect(lambda: self._update_selected_label(table, selected_label))
        table.itemDoubleClicked.connect(lambda item: self._inspect_double_clicked_cell(table, item))
        filter_mode.currentTextChanged.connect(
            lambda _text: self._apply_filter(table, filter_mode, filter_text, count_label)
        )
        filter_text.textChanged.connect(
            lambda _text: self._apply_filter(table, filter_mode, filter_text, count_label)
        )
        open_page.clicked.connect(lambda _checked=False: self._open_selected_page(table))
        source_urls.clicked.connect(lambda _checked=False: self._edit_source_urls(dialog))
        nexus_link.clicked.connect(lambda _checked=False: self._handle_nexus_link(dialog))
        fetch_sources.clicked.connect(
            lambda _checked=False: self._fetch_sources(
                dialog,
                table,
                progress_label,
                fetch_sources,
                source_urls,
                download_candidate,
            )
        )
        download_candidate.clicked.connect(
            lambda _checked=False: self._download_selected_candidate(
                dialog,
                table,
                progress_label,
                download_candidate,
                fetch_sources,
            )
        )
        all_downloads.clicked.connect(lambda _checked=False: self._show_all_fetched_downloads(dialog, table, progress_label, download_candidate, fetch_sources))
        false_match.clicked.connect(lambda _checked=False: self._mark_false_match(table, filter_mode, filter_text, count_label))
        manage_false_matches.clicked.connect(lambda _checked=False: self._manage_false_matches(table, filter_mode, filter_text, count_label))
        complete_slot.clicked.connect(lambda _checked=False: self._toggle_complete_slot(table, filter_mode, filter_text, count_label))
        ignore_mod.clicked.connect(lambda _checked=False: self._toggle_ignore(table, filter_mode, filter_text, count_label))
        classify_mod.clicked.connect(lambda _checked=False: self._classify_selected_mod(table, filter_mode, filter_text, count_label))
        voice_mods.clicked.connect(lambda _checked=False: self._show_voice_mods_inventory(dialog))
        search_web.clicked.connect(lambda _checked=False: self._search_selected_voice(table))
        close.clicked.connect(dialog.accept)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Filter"))
        controls.addWidget(filter_mode)
        controls.addWidget(filter_text, 1)
        controls.addWidget(QLabel("Local score"))
        controls.addWidget(threshold_spin)
        controls.addWidget(count_label)
        threshold_spin.valueChanged.connect(
            lambda value: self._set_local_match_threshold(
                table,
                value,
                filter_mode,
                filter_text,
                count_label,
                summary,
            )
        )

        false_match.hide()
        complete_slot.hide()
        ignore_mod.hide()
        classify_mod.hide()

        source_buttons = QHBoxLayout()
        source_buttons.addWidget(source_urls)
        source_buttons.addWidget(nexus_link)
        source_buttons.addWidget(fetch_sources)
        source_buttons.addWidget(all_downloads, 1)

        row_buttons = QHBoxLayout()
        row_buttons.addWidget(download_candidate)
        row_buttons.addWidget(open_page)
        row_buttons.addWidget(search_web)
        row_buttons.addWidget(false_match)
        row_buttons.addWidget(manage_false_matches)
        row_buttons.addWidget(complete_slot)
        row_buttons.addWidget(ignore_mod)
        row_buttons.addWidget(classify_mod)
        row_buttons.addWidget(voice_mods)
        row_buttons.addStretch(1)
        row_buttons.addWidget(close)

        summary = QLabel(self._summary_text(rows))
        summary.setWordWrap(True)

        layout = QVBoxLayout(dialog)
        layout.addWidget(summary)
        layout.addLayout(controls)
        layout.addWidget(selected_label)
        layout.addWidget(table)
        layout.addWidget(progress_label)
        layout.addLayout(source_buttons)
        layout.addLayout(row_buttons)
        dialog.setLayout(layout)

        self._apply_filter(table, filter_mode, filter_text, count_label)
        dialog.finished.connect(lambda _result: self._save_voice_window_config(dialog))
        if bool(window_config.get("maximized", True)):
            dialog.showMaximized()
        dialog.exec()

    def _build_voice_display_rows(self, rows: list[dict]) -> list[dict]:
        grouped: dict[str, dict] = {}

        for row in rows:
            base_key = str(row.get("base_internal_name") or row.get("base_mod") or "").lower()
            if not base_key:
                continue

            if base_key not in grouped:
                grouped[base_key] = {
                    "base_mod": row.get("base_mod") or "",
                    "base_internal_name": row.get("base_internal_name") or "",
                    "base_page_url": row.get("base_page_url") or "",
                    "classification_override": row.get("classification_override") or "auto",
                    "slots": {},
                }

            category = str(row.get("voice_category") or "")
            grouped[base_key]["slots"][category] = row

        display_rows = list(grouped.values())

        for display_row in display_rows:
            slot_rows = list((display_row.get("slots") or {}).values())
            display_row["status"] = self._aggregate_voice_status(slot_rows)
            display_row["installed_voice"] = self._best_display_candidate(slot_rows)
            display_row["score"] = self._best_display_score(slot_rows)
            display_row["online_candidate"] = self._best_display_online_candidate(slot_rows)
            display_row["online_score"] = self._best_display_online_score(slot_rows)

        return display_rows


    def _best_display_candidate(self, rows: list[dict]) -> str:
        candidates = [
            str(row.get("installed_voice") or "").strip()
            for row in rows
            if str(row.get("installed_voice") or "").strip()
        ]
        return " | ".join(candidates[:3])


    def _best_display_score(self, rows: list[dict]) -> int:
        scores = []
        for row in rows:
            try:
                scores.append(int(row.get("score") or 0))
            except (TypeError, ValueError):
                pass
        return max(scores) if scores else 0


    def _best_display_online_candidate(self, rows: list[dict]) -> str:
        candidates = [
            str(row.get("online_candidate") or "").strip()
            for row in rows
            if str(row.get("online_candidate") or "").strip()
        ]
        return " | ".join(candidates[:3])


    def _best_display_online_score(self, rows: list[dict]) -> int:
        scores = []
        for row in rows:
            try:
                scores.append(int(row.get("online_score") or 0))
            except (TypeError, ValueError):
                pass
        return max(scores) if scores else 0


    def _fill_display_table_row(self, table: QTableWidget, row_index: int, display_row: dict) -> None:
        slots = display_row.get("slots") or {}

        voicepack = slots.get("npc")
        dbvo = slots.get("player")
        ivdt = slots.get("scene")

        self._set_item(table, row_index, self.COL_STATUS, display_row.get("status", ""))
        self._set_item(table, row_index, self.COL_BASE_MOD, display_row.get("base_mod", ""))
        self._set_item(table, row_index, self.COL_VOICEPACK, self._voice_overview_text(voicepack) or "Missing")
        self._set_item(table, row_index, self.COL_DBVO, self._voice_overview_text(dbvo) or "Missing")
        self._set_item(table, row_index, self.COL_IVDT, self._voice_overview_text(ivdt) or "Missing")
        self._set_item(table, row_index, self.COL_CATEGORY, "")
        self._set_item(table, row_index, self.COL_INSTALLED_VOICE, display_row.get("installed_voice", ""))
        self._set_item(table, row_index, self.COL_SCORE, str(display_row.get("score") or ""))
        self._set_item(table, row_index, self.COL_ONLINE, display_row.get("online_candidate", ""))
        self._set_item(table, row_index, self.COL_ONLINE_SCORE, str(display_row.get("online_score") or ""))
        self._set_item(table, row_index, self.COL_SOURCE, "")
        self._set_item(table, row_index, self.COL_PAGE, display_row.get("base_page_url", ""))

        self._apply_display_row_background(table, row_index, display_row)


    def _apply_display_row_background(self, table: QTableWidget, row_index: int, display_row: dict) -> None:
        color = self._voice_row_background_color(display_row)

        for column in range(table.columnCount()):
            item = table.item(row_index, column)
            if item is None:
                continue

            if column in (self.COL_VOICEPACK, self.COL_DBVO, self.COL_IVDT):
                item.setBackground(self._overview_background_color(item.text()))
            else:
                item.setBackground(color)

            item.setForeground(QColor(242, 242, 242))


    def _refresh_display_rows(
        self,
        table: QTableWidget,
        filter_mode: QComboBox | None = None,
        filter_text: QLineEdit | None = None,
        count_label: QLabel | None = None,
    ) -> None:
        rows = getattr(table, "_ll_voice_rows", [])
        display_rows = self._build_voice_display_rows(rows)

        table._ll_voice_display_rows = display_rows
        table.setRowCount(len(display_rows))

        table.setUpdatesEnabled(False)
        try:
            for row_index, display_row in enumerate(display_rows):
                self._fill_display_table_row(table, row_index, display_row)
        finally:
            table.setUpdatesEnabled(True)

        if filter_mode is not None and filter_text is not None and count_label is not None:
            self._apply_filter(table, filter_mode, filter_text, count_label)

    def _voice_window_config(self) -> dict:
        config = self._load_voice_config()
        window = config.get("window")
        return window if isinstance(window, dict) else {}

    def _save_voice_window_config(self, dialog: QDialog) -> None:
        config = self._load_voice_config()
        size = dialog.normalGeometry().size() if dialog.isMaximized() else dialog.size()
        config["window"] = {
            "width": max(980, int(size.width())),
            "height": max(480, int(size.height())),
            "maximized": bool(dialog.isMaximized()),
        }
        self._save_voice_config(config)

    def _fill_table_row(self, table: QTableWidget, row_index: int, row: dict) -> None:
        self._set_item(table, row_index, self.COL_STATUS, row.get("status", ""))
        self._set_item(table, row_index, self.COL_BASE_MOD, row.get("base_mod", ""))
        overview = row.get("voice_overview") or {}
        self._set_item(table, row_index, self.COL_VOICEPACK, overview.get("npc", ""))
        self._set_item(table, row_index, self.COL_DBVO, overview.get("player", ""))
        self._set_item(table, row_index, self.COL_IVDT, overview.get("scene", ""))
        self._set_item(table, row_index, self.COL_CATEGORY, row.get("voice_category_label", ""))
        self._set_item(table, row_index, self.COL_INSTALLED_VOICE, row.get("installed_voice", ""))
        self._set_item(table, row_index, self.COL_SCORE, "Manual" if row.get("manual_voice") else str(row.get("score") or ""))
        self._set_item(table, row_index, self.COL_ONLINE, row.get("online_candidate", ""))
        self._set_item(table, row_index, self.COL_ONLINE_SCORE, str(row.get("online_score") or ""))
        self._set_item(table, row_index, self.COL_SOURCE, row.get("online_source_url", ""))
        self._set_item(table, row_index, self.COL_PAGE, row.get("base_page_url", ""))
        self._apply_voice_row_background(table, row_index, row)

    def _update_selected_label(self, table: QTableWidget, label: QLabel) -> None:
        row = self._selected_row(table)
        if not row:
            label.setText("Selected: none")
            return

        candidate = row.get("online_candidate") or row.get("installed_voice") or "no candidate"
        label.setText(
            f"Selected: {row.get('base_mod') or ''} | "
            f"{row.get('voice_category_label') or ''} | "
            f"{row.get('slot_status') or row.get('status') or ''} | {candidate}"
        )

    def _set_item(self, table: QTableWidget, row: int, column: int, value: str) -> None:
        item = QTableWidgetItem(str(value or ""))
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
        item.setData(Qt.ItemDataRole.UserRole, row)
        full_text = str(value or "")
        if full_text:
            item.setToolTip(full_text)

        if column == self.COL_STATUS:
            status = str(value or "")
            if status == "Missing":
                item.setToolTip("Missing: no installed voice-like mod matched this slot.")
            elif status == "Possible":
                item.setToolTip("Possible: a candidate was found, but the score is not high enough to fully trust.")
            elif status == "Installed":
                item.setToolTip("Installed: a likely installed voice pack was found.")
            elif status == "Complete":
                item.setToolTip("This voice slot is marked as done and will stay out of the missing queue.")
            elif status == "None":
                item.setToolTip("None: this voice slot is intentionally marked as not existing for this mod.")
        elif column in (self.COL_VOICEPACK, self.COL_DBVO, self.COL_IVDT):
            item.setToolTip(str(value or "Missing"))
        elif column == self.COL_CATEGORY:
            item.setToolTip("Voice slot: Voicepack for NPC/dialogue voices, DBVO/DVO for player voice, or IVDT/IVDT/DVIT for scene/addon voices.")
        elif column == self.COL_INSTALLED_VOICE:
            item.setToolTip(f"{full_text}\n\nDouble-click to inspect all installed voice candidates.".strip())
        elif column == self.COL_ONLINE:
            item.setToolTip(f"{full_text}\n\nUse Choose / Download to inspect all online candidates.".strip())

        table.setItem(row, column, item)

    def _apply_voice_row_background(self, table: QTableWidget, row_index: int, row: dict) -> None:
        color = self._voice_row_background_color(row)
        for column in range(table.columnCount()):
            item = table.item(row_index, column)
            if item is None:
                continue
            if column in (self.COL_VOICEPACK, self.COL_DBVO, self.COL_IVDT):
                item.setBackground(self._overview_background_color(item.text()))
            else:
                item.setBackground(color)
            item.setForeground(QColor(242, 242, 242))

    def _overview_background_color(self, value: str) -> QColor:
        text = str(value or "").lower()
        if text.startswith("manual") or text.startswith("installed") or text.startswith("ok"):
            return QColor(34, 78, 52)
        if text.startswith("possible"):
            return QColor(88, 70, 30)
        if text.startswith("online"):
            return QColor(34, 58, 86)
        if text == "complete":
            return QColor(34, 68, 88)
        if text == "none":
            return QColor(58, 58, 68)
        if text == "ignored":
            return QColor(54, 54, 54)
        if text == "missing":
            return QColor(78, 42, 42)
        return QColor(38, 42, 46)

    def _voice_row_background_color(self, row: dict) -> QColor:
        status = str(row.get("status") or "")
        score = int(row.get("score") or 0)
        online_score = int(row.get("online_score") or 0)

        if row.get("manual_voice"):
            return QColor(34, 88, 62)
        if status == "Complete":
            return QColor(34, 68, 88)
        if status == "None":
            return QColor(58, 58, 68)
        if status == "Installed" or score >= 90:
            return QColor(34, 78, 52)
        if status == "Possible" or score >= 55:
            return QColor(88, 70, 30)
        if status == "Online found":
            if online_score >= 90:
                return QColor(34, 78, 52)
            if online_score >= 55:
                return QColor(88, 70, 30)
            return QColor(34, 58, 86)
        if status == "Missing":
            return QColor(78, 42, 42)
        if status == "Ignored":
            return QColor(54, 54, 54)
        return QColor(38, 42, 46)

    def _candidate_background_color(self, score: int, manual: bool = False) -> QColor:
        if manual:
            return QColor(34, 88, 62)
        if score >= 90:
            return QColor(34, 78, 52)
        if score >= 70:
            return QColor(64, 82, 42)
        if score >= 55:
            return QColor(88, 70, 30)
        if score > 0:
            return QColor(76, 48, 40)
        return QColor(45, 45, 45)

    def _inspect_double_clicked_cell(self, table: QTableWidget, item: QTableWidgetItem) -> None:
        self._show_installed_voice_candidates(table)

    def _summary_text(self, rows: list[dict]) -> str:
        total = len(rows)
        missing = sum(1 for row in rows if row["status"] == "Missing")
        online = sum(1 for row in rows if row["status"] == "Online found")
        possible = sum(1 for row in rows if row["status"] == "Possible")
        installed = sum(1 for row in rows if row["status"] == "Installed")
        complete = sum(1 for row in rows if row["status"] == "Complete")
        none = sum(1 for row in rows if row["status"] == "None")
        ignored = sum(1 for row in rows if row["status"] == "Ignored")
        missing_by_type = []
        for category, label in VOICE_CATEGORIES:
            count = sum(
                1
                for row in rows
                if row.get("voice_category") == category and (row.get("slot_status") or row.get("status")) in ("Missing", "Online found")
            )
            missing_by_type.append(f"{label}: {count}")

        return (
            f"Scanned {total} LoversLab base mods. "
            f"Missing: {missing}. Online found: {online}. Possible: {possible}. "
            f"Installed: {installed}. Complete: {complete}. None: {none}. Ignored: {ignored}. "
            f"Needs check by type: {', '.join(missing_by_type)}. "
            "Add voice source URLs, fetch them, then download selected candidates into MO2 downloads."
        )

    def _apply_filter(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        mode = filter_mode.currentText()
        needle = filter_text.text().strip().lower()
        visible = 0

        display_rows = getattr(table, "_ll_voice_display_rows", [])

        for row_index, display_row in enumerate(display_rows):
            show = True
            status = str(display_row.get("status") or "")

            if mode != "All" and status != mode:
                slots = display_row.get("slots") or {}
                if not any(str(slot.get("slot_status") or slot.get("status") or "") == mode for slot in slots.values()):
                    show = False

            if needle:
                slots = display_row.get("slots") or {}
                haystack_parts = [
                    status,
                    display_row.get("base_mod", ""),
                    display_row.get("base_page_url", ""),
                    display_row.get("installed_voice", ""),
                    display_row.get("online_candidate", ""),
                ]

                for slot in slots.values():
                    haystack_parts.extend([
                        slot.get("slot_status", ""),
                        slot.get("status", ""),
                        slot.get("voice_category", ""),
                        slot.get("voice_category_label", ""),
                        slot.get("installed_voice", ""),
                        slot.get("online_candidate", ""),
                        slot.get("online_source_url", ""),
                        slot.get("search_query", ""),
                    ])

                haystack = " ".join(str(value or "") for value in haystack_parts).lower()
                if needle not in haystack:
                    show = False

            table.setRowHidden(row_index, not show)
            if show:
                visible += 1

        count_label.setText(f"{visible} / {table.rowCount()}")

    def _sort_voice_table(
        self,
        table: QTableWidget,
        column: int,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        rows = getattr(table, "_ll_voice_rows", [])
        selected = self._selected_row(table)
        selected_key = (
            str(selected.get("base_internal_name") or selected.get("base_mod") or ""),
            str(selected.get("voice_category") or ""),
        ) if selected else None

        previous_column = int(getattr(table, "_ll_voice_sort_column", -1))
        ascending = bool(getattr(table, "_ll_voice_sort_ascending", True))
        if previous_column == column:
            ascending = not ascending
        else:
            ascending = True
        table._ll_voice_sort_column = column
        table._ll_voice_sort_ascending = ascending

        rows.sort(key=lambda row: self._voice_sort_key(row, column), reverse=not ascending)
        self._refresh_voice_overviews(table)
        if selected_key:
            for index, row in enumerate(rows):
                key = (
                    str(row.get("base_internal_name") or row.get("base_mod") or ""),
                    str(row.get("voice_category") or ""),
                )
                if key == selected_key:
                    table.selectRow(index)
                    break
        self._apply_filter(table, filter_mode, filter_text, count_label)

    def _voice_sort_key(self, row: dict, column: int) -> tuple:
        if column == self.COL_STATUS:
            return (self._status_sort_bucket(row.get("status")), str(row.get("base_mod") or "").lower())
        if column == self.COL_BASE_MOD:
            return (str(row.get("base_mod") or "").lower(), self._category_sort_index(row))
        if column == self.COL_VOICEPACK:
            return self._overview_sort_key(row, "npc")
        if column == self.COL_DBVO:
            return self._overview_sort_key(row, "player")
        if column == self.COL_IVDT:
            return self._overview_sort_key(row, "scene")
        if column == self.COL_CATEGORY:
            return (self._category_sort_index(row), str(row.get("base_mod") or "").lower())
        if column == self.COL_INSTALLED_VOICE:
            return (str(row.get("installed_voice") or "").lower(), str(row.get("base_mod") or "").lower())
        if column == self.COL_SCORE:
            return (-int(row.get("score") or 0), str(row.get("base_mod") or "").lower())
        if column == self.COL_ONLINE:
            return (str(row.get("online_candidate") or "").lower(), str(row.get("base_mod") or "").lower())
        if column == self.COL_ONLINE_SCORE:
            return (-int(row.get("online_score") or 0), str(row.get("base_mod") or "").lower())
        if column == self.COL_SOURCE:
            return (str(row.get("online_source_url") or "").lower(), str(row.get("base_mod") or "").lower())
        if column == self.COL_PAGE:
            return (str(row.get("base_page_url") or "").lower(), str(row.get("base_mod") or "").lower())
        return (str(row.get("base_mod") or "").lower(), self._category_sort_index(row))

    def _overview_sort_key(self, row: dict, category: str) -> tuple:
        value = str((row.get("voice_overview") or {}).get(category) or "Missing")
        return (self._status_sort_bucket(value), value.lower(), str(row.get("base_mod") or "").lower())

    def _status_sort_bucket(self, value: object) -> int:
        text = str(value or "").lower()
        if text.startswith("complete"):
            return 0
        if text.startswith("none"):
            return 0
        if text.startswith("manual") or text.startswith("installed"):
            return 1
        if text.startswith("possible") or text.startswith("online"):
            return 2
        if text.startswith("missing"):
            return 3
        if text.startswith("ignored"):
            return 4
        return 5

    def _category_sort_index(self, row: dict) -> int:
        category_order = {category: index for index, (category, _label) in enumerate(VOICE_CATEGORIES)}
        return category_order.get(str(row.get("voice_category") or ""), 99)

    def _refresh_voice_overviews(self, table: QTableWidget) -> None:
        self._attach_voice_overview(getattr(table, "_ll_voice_rows", []))
        self._refresh_display_rows(table)

    def _refresh_voice_base(self, table: QTableWidget, target_row: dict) -> None:
        self._attach_voice_overview(getattr(table, "_ll_voice_rows", []))
        self._refresh_display_rows(table)

    def _set_local_match_threshold(
        self,
        table: QTableWidget,
        value: int,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
        summary: QLabel,
    ) -> None:
        config = self._load_voice_config()
        config["localMatchThreshold"] = int(value)
        self._save_voice_config(config)

        for row in getattr(table, "_ll_voice_rows", []):
            row["local_match_threshold"] = int(value)
            if (row.get("slot_status") or row.get("status")) in ("Ignored", "Complete", "None"):
                continue
            self._apply_local_threshold_to_row(row, int(value))

        self._attach_voice_overview(getattr(table, "_ll_voice_rows", []))
        self._refresh_display_rows(table, filter_mode, filter_text, count_label)
        summary.setText(self._summary_text(getattr(table, "_ll_voice_rows", [])))

    def _apply_local_threshold_to_row(self, row: dict, threshold: int) -> None:
        candidates = list(row.get("installed_voice_candidates") or [])
        best = next(
            (
                candidate
                for candidate in candidates
                if candidate.get("manual") or int(candidate.get("score") or 0) >= threshold
            ),
            None,
        )
        if best:
            row["installed_voice"] = best.get("display_name") or ""
            row["installed_voice_internal_name"] = best.get("internal_name") or ""
            row["score"] = int(best.get("score") or 0)
            row["manual_voice"] = bool(best.get("manual"))
        else:
            row["installed_voice"] = ""
            row["installed_voice_internal_name"] = ""
            row["online_candidate"] = ""
            row["online_download_url"] = ""
            row["online_source_url"] = ""
            row["online_score"] = 0
            row["score"] = 0
            row["manual_voice"] = False

        row["slot_status"] = self._voice_row_status(
            False,
            False,
            bool(row.get("none_voice")),
            bool(row.get("manual_voice")),
            int(row.get("score") or 0),
            bool(row.get("online_candidate")),
            threshold,
        )
        row["status"] = row["slot_status"]

    def _selected_row(self, table: QTableWidget) -> dict | None:
        context_row = getattr(table, "_ll_context_voice_row", None)
        if isinstance(context_row, dict):
            table._ll_context_voice_row = None
            return context_row

        selected = table.selectedItems()
        if not selected:
            return None

        row_index = selected[0].row()
        display_rows = getattr(table, "_ll_voice_display_rows", [])
        if row_index < 0 or row_index >= len(display_rows):
            return None

        display_row = display_rows[row_index]
        slots = display_row.get("slots") or {}

        # Fallback seulement pour boutons globaux non-slot.
        return slots.get("npc") or slots.get("player") or slots.get("scene")

    def _show_voice_context_menu(
        self,
        table: QTableWidget,
        pos,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        item = table.itemAt(pos)
        if item is None:
            return

        table.setCurrentCell(item.row(), item.column())
        slot_row = self._context_target_row(table, item.row(), item.column())

        # Pas une colonne slot => pas de menu de taggage.
        if not slot_row:
            return

        menu = QMenu(table)

        title = QAction(
            f"{slot_row.get('base_mod') or ''} | {slot_row.get('voice_category_label') or ''}",
            menu,
        )
        title.setEnabled(False)
        menu.addAction(title)
        menu.addSeparator()

        complete_action = QAction("Mark this slot complete", menu)
        none_action = QAction("Set this slot to None / does not exist", menu)
        reopen_action = QAction("Reopen this slot", menu)
        inspect_action = QAction("Inspect installed candidates", menu)

        complete_action.triggered.connect(
            lambda _checked=False: self._set_slot_flag(
                table, slot_row, "complete", filter_mode, filter_text, count_label
            )
        )
        none_action.triggered.connect(
            lambda _checked=False: self._set_slot_flag(
                table, slot_row, "none", filter_mode, filter_text, count_label
            )
        )
        reopen_action.triggered.connect(
            lambda _checked=False: self._set_slot_flag(
                table, slot_row, "reopen", filter_mode, filter_text, count_label
            )
        )
        inspect_action.triggered.connect(
            lambda _checked=False: self._select_row_and_inspect(table, slot_row)
        )

        menu.addAction(complete_action)
        menu.addAction(none_action)
        menu.addAction(reopen_action)
        menu.addSeparator()
        menu.addAction(inspect_action)

        if str(slot_row.get("installed_voice") or "").strip():
            false_match_action = QAction("Reject this installed candidate / false local match", menu)
            false_match_action.triggered.connect(
                lambda _checked=False: self._select_row_and_mark_false_match(
                    table, slot_row, filter_mode, filter_text, count_label
                )
            )
            menu.addAction(false_match_action)

        menu.exec(table.viewport().mapToGlobal(pos))

    def _context_target_row(self, table: QTableWidget, row_index: int, column: int) -> dict | None:
        display_rows = getattr(table, "_ll_voice_display_rows", [])
        if row_index < 0 or row_index >= len(display_rows):
            return None

        display_row = display_rows[row_index]
        slots = display_row.get("slots") or {}

        category = self._category_for_overview_column(column)
        if category:
            return slots.get(category)

        return None

    def _category_for_overview_column(self, column: int) -> str:
        if column == self.COL_VOICEPACK:
            return "npc"
        if column == self.COL_DBVO:
            return "player"
        if column == self.COL_IVDT:
            return "scene"
        return ""

    def _select_row_and_inspect(self, table: QTableWidget, row: dict) -> None:
        self._select_voice_row(table, row)
        self._show_installed_voice_candidates(table)

    def _select_row_and_mark_false_match(
        self,
        table: QTableWidget,
        row: dict,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        self._select_voice_row(table, row)
        self._mark_false_match(table, filter_mode, filter_text, count_label)

    def _select_voice_row(self, table: QTableWidget, row: dict) -> None:
        table._ll_context_voice_row = row

    def _set_slot_flag(
        self,
        table: QTableWidget,
        row: dict,
        mode: str,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        slot_key = str(row.get("slot_key") or self._slot_key(row, str(row.get("voice_category") or "player"))).lower()
        if not slot_key:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Selected voice slot cannot be updated.")
            return

        config = self._load_voice_config()
        complete_slots = {str(value).lower() for value in config.get("completeVoiceSlots", [])}
        none_slots = {str(value).lower() for value in config.get("noneVoiceSlots", [])}
        manual_matches = dict(config.get("manualVoiceMatches") or {})

        if mode == "complete":
            complete_slots.add(slot_key)
            none_slots.discard(slot_key)
            row["complete_voice"] = True
            row["none_voice"] = False
            row["slot_status"] = "Complete"
            row["status"] = "Complete"
        elif mode == "none":
            confirm = QMessageBox.question(
                self._parentWidget(),
                PLUGIN_NAME,
                "Mark this voice slot as None / does not exist?\n\n"
                f"Base: {row.get('base_mod') or ''}\n"
                f"Slot: {row.get('voice_category_label') or ''}\n\n"
                "This hides it from the missing list until you reopen it.",
            )
            if confirm != QMessageBox.StandardButton.Yes:
                return
            none_slots.add(slot_key)
            complete_slots.discard(slot_key)
            row["complete_voice"] = False
            row["none_voice"] = True
            row["installed_voice"] = ""
            row["installed_voice_internal_name"] = ""
            row["score"] = 0
            row["manual_voice"] = False
            manual_matches.pop(slot_key, None)
            row["installed_voice_candidates"] = [
                candidate for candidate in row.get("installed_voice_candidates", []) if not candidate.get("manual")
            ]
            row["slot_status"] = "None"
            row["status"] = "None"
        else:
            complete_slots.discard(slot_key)
            none_slots.discard(slot_key)
            row["complete_voice"] = False
            row["none_voice"] = False
            self._apply_local_threshold_to_row(row, self._local_match_threshold(config))

        config["completeVoiceSlots"] = sorted(complete_slots)
        config["noneVoiceSlots"] = sorted(none_slots)
        config["manualVoiceMatches"] = manual_matches
        self._save_voice_config(config)
        self._refresh_voice_base(table, row)
        self._apply_filter(table, filter_mode, filter_text, count_label)

    def _open_selected_page(self, table: QTableWidget) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        url = row.get("base_page_url") or ""
        if not url:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "This row has no LoversLab page URL.",
            )
            return

        if "loverslab.com/files/file/" in url.lower():
            url = with_query_value(url, "do", "download")
        webbrowser.open(url)

    def _show_installed_voice_candidates(self, table: QTableWidget) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        candidates = list(row.get("installed_voice_candidates") or [])
        candidates = sorted(
            candidates,
            key=lambda item: (-int(item.get("score") or 0), str(item.get("display_name") or "").lower()),
        )
        if not candidates:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "No installed voice candidates were found for this mod.",
            )
            return

        dialog = QDialog(table.window())
        dialog.setWindowTitle(f"Installed voice candidates - {row.get('base_mod') or ''}")
        dialog.resize(900, 500)

        candidate_table = QTableWidget(dialog)
        candidate_table.setColumnCount(4)
        candidate_table.setHorizontalHeaderLabels(["Score", "Type", "Installed voice candidate", "Folder"])
        candidate_table.setRowCount(len(candidates))
        candidate_table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        candidate_table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        candidate_table.setAlternatingRowColors(True)
        candidate_table._ll_installed_voice_candidates = candidates

        for index, candidate in enumerate(candidates):
            for column, value in enumerate([
                "Manual" if candidate.get("manual") else str(candidate.get("score") or ""),
                VOICE_CATEGORY_LABELS.get(candidate.get("voice_category") or "", candidate.get("voice_category") or ""),
                candidate.get("display_name") or "",
                candidate.get("mod_path") or "",
            ]):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(str(value or ""))
                item.setBackground(self._candidate_background_color(
                    int(candidate.get("score") or 0),
                    bool(candidate.get("manual")),
                ))
                item.setForeground(QColor(242, 242, 242))
                candidate_table.setItem(index, column, item)

        header = candidate_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        candidate_table.selectRow(0)

        title = QLabel(f"Base mod: {row.get('base_mod') or ''} | {row.get('voice_category_label') or ''}")
        title.setStyleSheet("font-weight: 700;")
        hint = QLabel("Installed candidates are sorted by score. Open a folder to inspect archive contents, or fix the selected candidate as the manual match for this base mod.")
        hint.setWordWrap(True)

        filter_text = QLineEdit(dialog)
        filter_text.setPlaceholderText("Filter candidate name or folder")
        filter_text.setClearButtonEnabled(True)
        filter_count = QLabel(dialog)
        filter_text.textChanged.connect(
            lambda _text: self._apply_installed_candidate_filter(candidate_table, filter_text, filter_count)
        )

        open_folder = QPushButton("Open selected folder")
        open_folder.clicked.connect(lambda _checked=False: self._open_selected_voice_folder(candidate_table))
        candidate_table.itemDoubleClicked.connect(lambda _item: self._open_selected_voice_folder(candidate_table))

        fix_manual = QPushButton("Fix manual selected")
        fix_manual.clicked.connect(
            lambda _checked=False: self._fix_manual_voice_candidate(
                dialog,
                table,
                candidate_table,
                row,
            )
        )

        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)

        buttons = QHBoxLayout()
        buttons.addWidget(open_folder)
        buttons.addWidget(fix_manual)
        buttons.addStretch(1)
        buttons.addWidget(close)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Filter"))
        filter_row.addWidget(filter_text, 1)
        filter_row.addWidget(filter_count)

        layout = QVBoxLayout(dialog)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addLayout(filter_row)
        layout.addWidget(candidate_table)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        self._apply_installed_candidate_filter(candidate_table, filter_text, filter_count)
        dialog.exec()

    def _apply_installed_candidate_filter(
        self,
        table: QTableWidget,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        needle = filter_text.text().strip().lower()
        candidates = getattr(table, "_ll_installed_voice_candidates", [])
        visible = 0
        first_visible = -1

        for row_index, candidate in enumerate(candidates):
            haystack = " ".join([
                str(candidate.get("score") or ""),
                "manual" if candidate.get("manual") else "",
                VOICE_CATEGORY_LABELS.get(candidate.get("voice_category") or "", candidate.get("voice_category") or ""),
                candidate.get("display_name") or "",
                candidate.get("internal_name") or "",
                candidate.get("mod_path") or "",
            ]).lower()
            show = not needle or needle in haystack
            table.setRowHidden(row_index, not show)
            if show:
                visible += 1
                if first_visible < 0:
                    first_visible = row_index

        count_label.setText(f"{visible} / {len(candidates)}")
        selected = table.selectedItems()
        if first_visible >= 0 and (not selected or table.isRowHidden(selected[0].row())):
            table.selectRow(first_visible)

    def _fix_manual_voice_candidate(
        self,
        dialog: QDialog,
        main_table: QTableWidget,
        candidate_table: QTableWidget,
        row: dict,
    ) -> None:
        selected = candidate_table.selectedItems()
        if not selected:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a voice candidate first.")
            return

        candidates = getattr(candidate_table, "_ll_installed_voice_candidates", [])
        candidate_index = selected[0].row()
        if candidate_index < 0 or candidate_index >= len(candidates):
            return

        candidate = dict(candidates[candidate_index])
        slot_key = str(row.get("slot_key") or self._slot_key(row, str(row.get("voice_category") or "player"))).lower()
        if not slot_key or not candidate.get("internal_name"):
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "This candidate cannot be bound manually.")
            return

        config = self._load_voice_config()
        manual_matches = dict(config.get("manualVoiceMatches") or {})
        manual_matches[slot_key] = {
            "internal_name": candidate.get("internal_name") or "",
            "display_name": candidate.get("display_name") or "",
            "mod_path": candidate.get("mod_path") or "",
            "voice_category": row.get("voice_category") or "",
        }
        config["manualVoiceMatches"] = manual_matches
        self._save_voice_config(config)

        candidate["score"] = 1000
        candidate["manual"] = True
        candidate["voice_category"] = row.get("voice_category") or candidate.get("voice_category") or ""
        row["installed_voice"] = candidate.get("display_name") or ""
        row["installed_voice_internal_name"] = candidate.get("internal_name") or ""
        row["score"] = 1000
        row["manual_voice"] = True
        row["slot_status"] = "Installed"
        row["status"] = "Installed"
        row["installed_voice_candidates"] = [
            candidate,
            *[
                item
                for item in row.get("installed_voice_candidates", [])
                if str(item.get("internal_name") or "").lower()
                != str(candidate.get("internal_name") or "").lower()
            ],
        ]

        self._refresh_voice_overviews(main_table)
        QMessageBox.information(
            dialog,
            PLUGIN_NAME,
            f"Manual voice match saved:\n\n{row.get('base_mod') or ''}\n"
            f"{row.get('voice_category_label') or ''} -> {candidate.get('display_name') or ''}",
        )

    def _open_selected_voice_folder(self, table: QTableWidget) -> None:
        selected = table.selectedItems()
        if not selected:
            return
        candidates = getattr(table, "_ll_installed_voice_candidates", [])
        row_index = selected[0].row()
        if row_index < 0 or row_index >= len(candidates):
            return
        path = Path(str(candidates[row_index].get("mod_path") or ""))
        if not path.exists():
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, f"Folder not found:\n{path}")
            return
        self._open_path(path)

    def _open_path(self, path: Path) -> None:
        target = path if path.is_dir() else path.parent
        try:
            webbrowser.open(target.as_uri())
        except ValueError:
            webbrowser.open(str(target))

    def _edit_source_urls(self, parent: QDialog) -> None:
        config = self._load_voice_config()
        dialog = QDialog(parent)
        dialog.setWindowTitle("Voice Source URLs")
        dialog.resize(780, 420)

        text = QTextEdit(dialog)
        text.setPlainText("\n".join(config.get("voiceSourceUrls", [])))
        text.setPlaceholderText("One LoversLab or Nexus Mods source URL per line")

        help_text = QLabel(
            "Add LoversLab download pages or Nexus Mods pages. Nexus URLs are normalized to their Files tab automatically."
        )
        help_text.setWordWrap(True)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        layout = QVBoxLayout(dialog)
        layout.addWidget(help_text)
        layout.addWidget(text)
        layout.addWidget(buttons)
        dialog.setLayout(layout)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        urls = []
        for line in text.toPlainText().splitlines():
            value = line.strip()
            if not value or value.startswith("#"):
                continue
            value = normalize_voice_source_url(value)
            if value not in urls:
                urls.append(value)
        config["voiceSourceUrls"] = urls
        self._save_voice_config(config)
        text.setPlainText("\n".join(urls))

    def _nexus_api_tutorial_image_path(self) -> Path | None:
        here = Path(__file__).resolve()
        candidates = [
            here.parent / "Mo2_ImageTutorial" / "API_Explain.png",
            here.parent.parent / "Mo2_ImageTutorial" / "API_Explain.png",
            ROOT_DIR / "Mo2_ImageTutorial" / "API_Explain.png",
            Path.cwd() / "Mo2_ImageTutorial" / "API_Explain.png",
        ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return None

    def _handle_nexus_link(self, parent: QDialog) -> None:
        config = self._load_voice_config()
        dialog = QDialog(parent)
        dialog.setWindowTitle("Handle Nexus Link")
        dialog.resize(760, 620)

        help_text = QLabel(
            "Paste a Nexus Mods personal API key here. It is stored only in LL Integration's local MO2 plugin config "
            "and is used to list Nexus file downloads for source scoring."
        )
        help_text.setWordWrap(True)

        image_label = QLabel(dialog)
        tutorial_path = self._nexus_api_tutorial_image_path()
        if tutorial_path is not None:
            pixmap = QPixmap(str(tutorial_path))
            if not pixmap.isNull():
                image_label.setPixmap(pixmap.scaledToWidth(620, Qt.TransformationMode.SmoothTransformation))
                image_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            else:
                image_label.setText("Nexus API key tutorial image could not be loaded.")
        else:
            image_label.setText("Nexus API key tutorial image not found.")

        key_input = QLineEdit(dialog)
        key_input.setEchoMode(QLineEdit.EchoMode.Password)
        key_input.setPlaceholderText("Nexus Mods API key")
        key_input.setText(str(config.get("nexusApiKey") or ""))

        status = QLabel("Ready")
        status.setWordWrap(True)

        def set_api_status(text: str, ready: bool) -> None:
            color = "#0f5f36" if ready else "#6f2323"
            border = "#35d07f" if ready else "#e05c5c"
            status.setText(text)
            status.setStyleSheet(f"padding: 6px; border: 1px solid {border}; background: {color};")

        set_api_status(
            "Nexus API key is configured." if key_input.text().strip() else "No Nexus API key configured.",
            bool(key_input.text().strip()),
        )

        open_api = QPushButton("Open Nexus API Access")
        save_key = QPushButton("Save API key")
        validate_key = QPushButton("Validate API key")
        clear_key = QPushButton("Clear")
        close = QPushButton("Close")

        def save() -> None:
            updated = self._load_voice_config()
            updated["nexusApiKey"] = key_input.text().strip()
            self._save_voice_config(updated)
            set_api_status(
                "Nexus API key saved." if key_input.text().strip() else "Nexus API key removed.",
                bool(key_input.text().strip()),
            )

        def validate() -> None:
            api_key = key_input.text().strip()
            if not api_key:
                QMessageBox.information(dialog, PLUGIN_NAME, "Paste a Nexus API key first.")
                return
            try:
                payload = validate_nexus_api_key(api_key, timeout=30.0)
            except Exception as exc:
                set_api_status("Nexus API key validation failed.", False)
                QMessageBox.warning(dialog, PLUGIN_NAME, f"Nexus API key validation failed:\n\n{exc}")
                return
            name = str(payload.get("name") or payload.get("user_id") or "").strip()
            set_api_status(f"Nexus API key validated{f' for {name}' if name else ''}.", True)

        open_api.clicked.connect(lambda _checked=False: webbrowser.open("https://www.nexusmods.com/users/myaccount?tab=api"))
        save_key.clicked.connect(lambda _checked=False: save())
        validate_key.clicked.connect(lambda _checked=False: validate())
        clear_key.clicked.connect(lambda _checked=False: (key_input.clear(), set_api_status("Key field cleared. Save to remove it from config.", False)))
        close.clicked.connect(dialog.accept)

        key_row = QHBoxLayout()
        key_row.addWidget(QLabel("API key"))
        key_row.addWidget(key_input, 1)

        buttons = QHBoxLayout()
        buttons.addWidget(open_api)
        buttons.addWidget(save_key)
        buttons.addWidget(validate_key)
        buttons.addWidget(clear_key)
        buttons.addStretch(1)
        buttons.addWidget(close)

        layout = QVBoxLayout(dialog)
        layout.addWidget(help_text)
        layout.addWidget(image_label)
        layout.addLayout(key_row)
        layout.addWidget(status)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec()

    def _fetch_sources(
        self,
        dialog: QDialog,
        table: QTableWidget,
        progress_label: QLabel,
        fetch_button: QPushButton,
        sources_button: QPushButton,
        download_button: QPushButton,
    ) -> None:
        config = self._load_voice_config()
        urls = [str(url).strip() for url in config.get("voiceSourceUrls", []) if str(url).strip()]
        if not urls:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "Add at least one voice source URL first.",
            )
            return

        cookies_path = self._setting_path("cookies_path", DEFAULT_COOKIES)
        rows = getattr(table, "_ll_voice_rows", [])
        thread = QThread(dialog)
        worker = VoiceSourceFetchWorker(
            rows,
            urls,
            cookies_path,
            config.get("falseMatches", []),
            config.get("nexusApiKey") or "",
            timeout=UPDATE_REQUEST_TIMEOUT_SECONDS,
        )
        worker.moveToThread(thread)

        fetch_button.setEnabled(False)
        sources_button.setEnabled(False)
        download_button.setEnabled(False)
        progress_label.setText("Fetching voice sources...")

        worker.statusChanged.connect(progress_label.setText)
        worker.candidatesReady.connect(lambda candidates: self._apply_online_candidates(table, list(candidates)))
        worker.downloadsReady.connect(lambda downloads: setattr(table, "_ll_voice_all_downloads", list(downloads)))
        worker.finished.connect(
            lambda ok, message: self._fetch_sources_finished(
                dialog,
                thread,
                worker,
                progress_label,
                fetch_button,
                sources_button,
                download_button,
                ok,
                message,
            )
        )
        thread.started.connect(worker.run)
        thread.start()

    def _fetch_sources_finished(
        self,
        dialog: QDialog,
        thread: QThread,
        worker: VoiceSourceFetchWorker,
        progress_label: QLabel,
        fetch_button: QPushButton,
        sources_button: QPushButton,
        download_button: QPushButton,
        ok: bool,
        message: str,
    ) -> None:
        progress_label.setText(message)
        fetch_button.setEnabled(True)
        sources_button.setEnabled(True)
        download_button.setEnabled(True)
        thread.quit()
        thread.wait()
        worker.deleteLater()
        thread.deleteLater()
        if not ok and message != "Cancelled":
            QMessageBox.warning(dialog, PLUGIN_NAME, f"Voice source fetch failed:\n\n{message}")

    def _apply_online_candidates(self, table: QTableWidget, candidates: list[dict]) -> None:
        rows = getattr(table, "_ll_voice_rows", [])
        candidates_by_base = {}
        for candidate in candidates:
            base = candidate.get("base_internal_name") or candidate.get("base_mod") or ""
            category = candidate.get("voice_category") or ""
            candidates_by_base.setdefault((base, category), []).append(candidate)

        for row_index, row in enumerate(rows):
            row_candidates = candidates_by_base.get((
                row.get("base_internal_name") or row.get("base_mod") or "",
                row.get("voice_category") or "",
            ))
            if not row_candidates:
                continue

            row_candidates = sorted(
                row_candidates,
                key=lambda item: int(item.get("online_score") or 0),
                reverse=True,
            )
            candidate = row_candidates[0]
            row.update({
                "online_candidate": candidate.get("download_name") or "",
                "online_download_url": candidate.get("download_url") or "",
                "online_source_url": candidate.get("source_url") or "",
                "online_score": candidate.get("online_score") or 0,
                "online_size": candidate.get("size") or "",
                "online_date_iso": candidate.get("date_iso") or "",
                "online_version": candidate.get("version") or "",
                "online_candidates": row_candidates,
            })
            if (row.get("slot_status") or row.get("status")) == "Missing":
                row["slot_status"] = "Online found"
                row["status"] = "Online found"
            self._fill_table_row(table, row_index, row)
        self._refresh_voice_overviews(table)

    def _download_selected_candidate(
        self,
        dialog: QDialog,
        table: QTableWidget,
        progress_label: QLabel,
        download_button: QPushButton,
        fetch_button: QPushButton,
    ) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        downloads_path = self._mo2_downloads_path()
        if not downloads_path:
            QMessageBox.warning(self._parentWidget(), PLUGIN_NAME, "MO2 downloads path is not available.")
            return

        self._show_online_candidate_download_dialog(
            dialog,
            row,
            downloads_path,
            progress_label,
            download_button,
            fetch_button,
        )

    def _candidate_from_online_table(self, table: QTableWidget) -> dict | None:
        selected = table.selectedItems()
        if not selected:
            return None

        candidates = getattr(table, "_ll_voice_candidates", [])
        index = selected[0].row()
        if index < 0 or index >= len(candidates):
            return None
        return dict(candidates[index])

    def _show_online_candidate_download_dialog(
        self,
        parent: QDialog,
        row: dict,
        downloads_path: Path,
        main_progress_label: QLabel,
        main_download_button: QPushButton,
        main_fetch_button: QPushButton,
    ) -> None:
        candidates = list(row.get("online_candidates") or [])
        if not candidates and row.get("online_download_url"):
            candidates = [{
                "download_name": row.get("online_candidate") or "",
                "download_url": row.get("online_download_url") or "",
                "source_url": row.get("online_source_url") or "",
                "online_score": row.get("online_score") or 0,
                "size": row.get("online_size") or "",
                "date_iso": row.get("online_date_iso") or "",
                "version": row.get("online_version") or "",
                "voice_category": row.get("voice_category") or "",
                "voice_category_label": row.get("voice_category_label") or "",
            }]

        candidates = sorted(
            candidates,
            key=lambda item: int(item.get("online_score") or 0),
            reverse=True,
        )
        if not candidates:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Selected row has no online candidate yet.")
            return

        dialog = QDialog(parent)
        dialog.setWindowTitle(f"Download voice candidate - {row.get('base_mod') or ''}")
        dialog.resize(920, 500)

        table = QTableWidget(dialog)
        table.setColumnCount(6)
        table.setHorizontalHeaderLabels(["Score", "Type", "Candidate", "Size", "Date", "Source"])
        table.setRowCount(len(candidates))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table._ll_voice_candidates = candidates

        for index, candidate in enumerate(candidates):
            score = int(candidate.get("online_score") or 0)
            for column, value in enumerate([
                str(score or ""),
                candidate.get("voice_category_label") or VOICE_CATEGORY_LABELS.get(candidate.get("voice_category") or "", candidate.get("voice_category") or ""),
                candidate.get("download_name") or "",
                candidate.get("size") or "",
                candidate.get("date_iso") or "",
                candidate.get("source_url") or "",
            ]):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(str(value or ""))
                item.setBackground(self._candidate_background_color(score))
                item.setForeground(QColor(242, 242, 242))
                table.setItem(index, column, item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        table.selectRow(0)

        title = QLabel(f"Base mod: {row.get('base_mod') or ''} | {row.get('voice_category_label') or ''}")
        title.setStyleSheet("font-weight: 700;")
        hint = QLabel("Select a candidate, then download. This window stays open and shows the download status.")
        hint.setWordWrap(True)
        status_label = QLabel("Ready")
        status_label.setWordWrap(True)

        download_selected = QPushButton("Download selected")
        hide_online = QPushButton("Hide online match")
        close = QPushButton("Close")
        close.clicked.connect(dialog.reject)

        def start_download() -> None:
            candidate = self._candidate_from_online_table(table)
            if not candidate:
                QMessageBox.information(dialog, PLUGIN_NAME, "Select a candidate first.")
                return
            self._start_online_candidate_download(
                dialog,
                row,
                candidate,
                downloads_path,
                status_label,
                main_progress_label,
                download_selected,
                close,
                main_download_button,
                main_fetch_button,
            )

        download_selected.clicked.connect(lambda _checked=False: start_download())
        hide_online.clicked.connect(lambda _checked=False: self._hide_selected_online_match(dialog, table, row))
        table.itemDoubleClicked.connect(lambda _item: start_download())

        buttons = QHBoxLayout()
        buttons.addWidget(download_selected)
        buttons.addWidget(hide_online)
        buttons.addStretch(1)
        buttons.addWidget(close)

        layout = QVBoxLayout(dialog)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addWidget(table)
        layout.addWidget(status_label)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec()

    def _show_all_fetched_downloads(
        self,
        parent: QDialog,
        main_table: QTableWidget,
        main_progress_label: QLabel,
        main_download_button: QPushButton,
        main_fetch_button: QPushButton,
    ) -> None:
        downloads = list(getattr(main_table, "_ll_voice_all_downloads", []) or [])
        if not downloads:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "No fetched downloads are available yet.\n\nUse Fetch sources first.",
            )
            return

        cache_key = self._all_source_downloads_cache_key(downloads)
        cached = getattr(main_table, "_ll_voice_all_downloads_score_cache", None)
        if isinstance(cached, dict) and cached.get("key") == cache_key:
            scored_downloads = [dict(item) for item in cached.get("rows", [])]
            dialog = QDialog(parent)
            dialog.setWindowTitle("All fetched source downloads")
            dialog.resize(1280, 760)
            dialog.setMinimumSize(1040, 560)
            title = QLabel("All fetched source downloads")
            title.setStyleSheet("font-weight: 700;")
            hint = QLabel("Each fetched archive is scored against all installed base mods. Filter or sort, then download using the best target shown in the row.")
            hint.setWordWrap(True)
            title.setMaximumHeight(22)
            hint.setMaximumHeight(24)
            hint.setStyleSheet("color: #c8c8c8;")
            layout = QVBoxLayout(dialog)
            dialog.setLayout(layout)
        else:
            dialog = QDialog(parent)
            dialog.setWindowTitle("All fetched source downloads")
            dialog.resize(1280, 760)
            dialog.setMinimumSize(1040, 560)

            title = QLabel("All fetched source downloads")
            title.setStyleSheet("font-weight: 700;")
            hint = QLabel("Each fetched archive is scored against all installed base mods. Filter or sort, then download using the best target shown in the row.")
            hint.setWordWrap(True)
            title.setMaximumHeight(22)
            hint.setMaximumHeight(24)
            hint.setStyleSheet("color: #c8c8c8;")
            loading_label = QLabel("Preparing scored download list...")
            loading_progress = QProgressBar(dialog)
            loading_progress.setRange(0, max(1, len(downloads)))

            layout = QVBoxLayout(dialog)
            layout.setContentsMargins(6, 6, 6, 6)
            layout.setSpacing(4)
            layout.addWidget(title)
            layout.addWidget(hint)
            layout.addWidget(loading_label)
            layout.addWidget(loading_progress)
            dialog.setLayout(layout)
            dialog.show()
            QApplication.processEvents()

            voice_rows = self._download_target_rows(list(getattr(main_table, "_ll_voice_rows", []) or []))
            voice_inventory = self._collect_voice_mods_inventory(show_message=False)
            downloaded_archives = self._downloaded_archive_names()
            score_cache: dict[tuple[str, str], int] = {}
            scored_downloads = []
            for index, download in enumerate(downloads, start=1):
                candidate = dict(download)
                target, score = self._best_download_target(voice_rows, candidate, score_cache)
                candidate["online_score"] = score
                candidate["already_status"] = self._download_already_status(candidate, voice_inventory, downloaded_archives)
                if target:
                    candidate["target_base_mod"] = target.get("base_mod") or ""
                    candidate["target_base_internal_name"] = target.get("base_internal_name") or ""
                    candidate["target_voice_category"] = target.get("voice_category") or ""
                    candidate["target_voice_category_label"] = target.get("voice_category_label") or ""
                scored_downloads.append(candidate)
                if index == len(downloads) or index % 25 == 0:
                    loading_progress.setValue(index)
                    loading_label.setText(f"Scoring source downloads... {index} / {len(downloads)}")
                    QApplication.processEvents()
            scored_downloads.sort(
                key=lambda item: (
                    -int(item.get("online_score") or 0),
                    str(item.get("target_base_mod") or "").lower(),
                    str(item.get("download_name") or "").lower(),
                )
            )
            main_table._ll_voice_all_downloads_score_cache = {
                "key": cache_key,
                "rows": [dict(item) for item in scored_downloads],
            }

        if "voice_rows" not in locals():
            voice_rows = self._download_target_rows(list(getattr(main_table, "_ll_voice_rows", []) or []))

        if "voice_inventory" not in locals():
            voice_inventory = self._collect_voice_mods_inventory(show_message=False)

        table = QTableWidget(dialog)
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Base mods", "Score", "Target online to download", "State"])
        table.setRowCount(len(scored_downloads))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table._ll_voice_candidates = scored_downloads

        table.verticalHeader().setDefaultSectionSize(22)
        table.verticalHeader().setMinimumSectionSize(18)
        table.horizontalHeader().setMinimumSectionSize(70)
        table.setAlternatingRowColors(True)

        inventory_panel = self._build_voice_inventory_panel(dialog, voice_inventory)
        inventory_panel.setMinimumWidth(280)
        inventory_panel.setMaximumWidth(380)
        
        for index, download in enumerate(scored_downloads):
            score = int(download.get("online_score") or 0)
            for column, value in enumerate([
                download.get("target_base_mod") or "",
                str(score or ""),
                download.get("download_name") or "",
                download.get("already_status") or "Not found",
            ]):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(self._source_download_tooltip(download))
                item.setBackground(self._candidate_background_color(score))
                item.setForeground(QColor(242, 242, 242))
                table.setItem(index, column, item)

        header = table.horizontalHeader()
        header.setSectionsMovable(False)
        header.setStretchLastSection(False)

        for column in range(table.columnCount()):
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Interactive)

        default_widths = [380, 70, 360, 170]
        for column, width in enumerate(default_widths):
            table.setColumnWidth(column, width)

        self._restore_all_fetch_table_widths(table)
        
        table.selectRow(0)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table.customContextMenuRequested.connect(lambda pos: self._show_source_download_context_menu(table, pos))

        filter_text = QLineEdit(dialog)
        filter_text.setPlaceholderText("Filter download name / source / date")
        filter_text.setClearButtonEnabled(True)
        count_label = QLabel(dialog)
        status_label = QLabel("Ready")
        status_label.setWordWrap(True)

        def apply_filter() -> None:
            needle = filter_text.text().strip().lower()
            visible = 0
            first_visible = -1
            for row_index, download in enumerate(scored_downloads):
                haystack = " ".join([
                    str(download.get("online_score") or ""),
                    download.get("already_status") or "",
                    download.get("target_base_mod") or "",
                    download.get("target_voice_category_label") or "",
                    VOICE_CATEGORY_LABELS.get(download.get("voice_category") or "", download.get("voice_category") or ""),
                    download.get("download_name") or "",
                    download.get("size") or "",
                    download.get("date_iso") or "",
                    download.get("source_url") or "",
                ]).lower()
                show = not needle or needle in haystack
                table.setRowHidden(row_index, not show)
                if show:
                    visible += 1
                    if first_visible < 0:
                        first_visible = row_index
            count_label.setText(f"{visible} / {len(scored_downloads)}")
            selected = table.selectedItems()
            if first_visible >= 0 and (not selected or table.isRowHidden(selected[0].row())):
                table.selectRow(first_visible)

        filter_text.textChanged.connect(lambda _text: apply_filter())

        download_selected = QPushButton("Download selected for best target")
        close = QPushButton("Close")
        close.clicked.connect(dialog.reject)

        def start_download() -> None:
            candidate = self._candidate_from_online_table(table)
            if not candidate:
                QMessageBox.information(dialog, PLUGIN_NAME, "Select a download first.")
                return
            downloads_path = self._mo2_downloads_path()
            if not downloads_path:
                QMessageBox.warning(dialog, PLUGIN_NAME, "MO2 downloads path is not available.")
                return
            target = self._target_row_for_download(voice_rows, candidate)
            if not target:
                QMessageBox.information(
                    dialog,
                    PLUGIN_NAME,
                    "No target mod was found for this download.\n\nUse the normal Choose / Download flow for a manually selected target.",
                )
                return
            self._start_online_candidate_download(
                dialog,
                target,
                candidate,
                downloads_path,
                status_label,
                main_progress_label,
                download_selected,
                close,
                main_download_button,
                main_fetch_button,
            )

        download_selected.clicked.connect(lambda _checked=False: start_download())
        table.itemDoubleClicked.connect(lambda _item: start_download())

        filter_row = QHBoxLayout()
        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(4)
        filter_row.addWidget(QLabel("Filter"))
        filter_row.addWidget(filter_text, 1)
        filter_row.addWidget(count_label)

        while layout.count():
            child = layout.takeAt(0)
            widget = child.widget()
            if widget is not None:
                widget.setParent(None)

        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        splitter = QSplitter(Qt.Orientation.Horizontal, dialog)
        splitter.addWidget(table)
        splitter.addWidget(inventory_panel)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)

        self._restore_all_fetch_splitter_sizes(splitter)

        status_label.setMaximumHeight(22)

        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(0, 0, 0, 0)
        bottom_row.setSpacing(6)
        bottom_row.addWidget(download_selected)
        bottom_row.addWidget(status_label, 1)
        bottom_row.addWidget(close)

        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(8)
        header_row.addWidget(title)
        header_row.addWidget(hint, 1)

        layout.addLayout(header_row)
        layout.addLayout(filter_row)
        layout.addWidget(splitter, 1)
        layout.addLayout(bottom_row)

        apply_filter()
        try:
            dialog.exec()
        finally:
            self._save_all_fetch_table_widths(table)
            self._save_all_fetch_splitter_sizes(splitter)

    def _source_download_tooltip(self, item: dict) -> str:
        return "\n".join([
            f"Matched base mod: {item.get('target_base_mod') or ''}",
            f"Score: {item.get('online_score') or ''}",
            f"Target type: {item.get('target_voice_category_label') or ''}",
            f"File type: {VOICE_CATEGORY_LABELS.get(item.get('voice_category') or '', item.get('voice_category') or '')}",
            f"State: {item.get('already_status') or 'Not found'}",
            f"Download: {item.get('download_name') or ''}",
            f"Size: {item.get('size') or ''}",
            f"Date: {item.get('date_iso') or ''}",
            f"Source: {item.get('source_url') or ''}",
        ]).strip()

    def _show_source_download_context_menu(self, table: QTableWidget, pos) -> None:
        item = table.itemAt(pos)
        if item is None:
            return
        table.selectRow(item.row())
        rows = getattr(table, "_ll_voice_candidates", [])
        index = item.row()
        if index < 0 or index >= len(rows):
            return
        download = rows[index]

        menu = QMenu(table)
        title = QAction(download.get("download_name") or "", menu)
        title.setEnabled(False)
        menu.addAction(title)
        menu.addSeparator()
        for label, value in [
            ("Matched base mod", download.get("target_base_mod") or ""),
            ("Score", download.get("online_score") or ""),
            ("Target type", download.get("target_voice_category_label") or ""),
            ("File type", VOICE_CATEGORY_LABELS.get(download.get("voice_category") or "", download.get("voice_category") or "")),
            ("State", download.get("already_status") or "Not found"),
            ("Size", download.get("size") or ""),
            ("Date", download.get("date_iso") or ""),
            ("Source", download.get("source_url") or ""),
        ]:
            action = QAction(f"{label}: {value}", menu)
            action.setEnabled(False)
            menu.addAction(action)
        source_url = str(download.get("source_url") or "").strip()
        if source_url:
            menu.addSeparator()
            open_source = QAction("Open source page", menu)
            open_source.triggered.connect(lambda _checked=False, url=source_url: webbrowser.open(url))
            menu.addAction(open_source)
        menu.exec(table.viewport().mapToGlobal(pos))

    def _download_target_rows(self, voice_rows: list[dict]) -> list[dict]:
        rows = [dict(row) for row in voice_rows if not row.get("_overview_duplicate")]
        seen = {
            (
                str(row.get("base_internal_name") or row.get("base_mod") or "").lower(),
                str(row.get("voice_category") or ""),
            )
            for row in rows
        }
        config = self._load_voice_config()
        forced_voice_mods = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base_mods = {str(value).lower() for value in config.get("forcedBaseMods", [])}
        try:
            mod_list = self._organizer.modList() if self._organizer else None
            internal_names = mod_list.allModsByProfilePriority() if mod_list else []
        except Exception:
            return rows
        for internal_name in internal_names:
            try:
                display_name = mod_list.displayName(internal_name)
            except Exception:
                display_name = str(internal_name)
            key = str(internal_name).lower()
            auto_voice = voice_keyword_present(display_name)
            is_voice = (auto_voice or key in forced_voice_mods) and key not in forced_base_mods
            if is_voice:
                continue
            for category, category_label in VOICE_CATEGORIES:
                row_key = (key, category)
                if row_key in seen:
                    continue
                seen.add(row_key)
                rows.append({
                    "base_mod": display_name,
                    "base_internal_name": internal_name,
                    "voice_category": category,
                    "voice_category_label": category_label,
                    "slot_status": "",
                    "status": "",
                    "_synthetic_download_target": True,
                })
        return rows

    def _all_source_downloads_cache_key(self, downloads: list[dict]) -> tuple:
        config = self._load_voice_config()
        forced = tuple(sorted(str(value).lower() for value in config.get("forcedVoiceMods", [])))
        forced_base = tuple(sorted(str(value).lower() for value in config.get("forcedBaseMods", [])))
        none_slots = tuple(sorted(str(value).lower() for value in config.get("noneVoiceSlots", [])))
        complete_slots = tuple(sorted(str(value).lower() for value in config.get("completeVoiceSlots", [])))
        archives = tuple(sorted(self._downloaded_archive_names()))
        mods = []
        try:
            mod_list = self._organizer.modList() if self._organizer else None
            internal_names = mod_list.allModsByProfilePriority() if mod_list else []
            for internal_name in internal_names:
                mods.append((str(internal_name), str(mod_list.displayName(internal_name))))
        except Exception:
            mods = []
        source_rows = tuple(
            (
                str(item.get("download_name") or ""),
                str(item.get("source_url") or ""),
                str(item.get("size") or ""),
                str(item.get("date_iso") or ""),
                str(item.get("voice_category") or ""),
            )
            for item in downloads
        )
        return (source_rows, tuple(mods), forced, forced_base, none_slots, complete_slots, archives)

    def _best_download_target(
        self,
        rows: list[dict],
        download: dict,
        score_cache: dict[tuple[str, str], int] | None = None,
    ) -> tuple[dict | None, int]:
        best_row = None
        best_score = 0
        for row in rows:
            if row.get("_overview_duplicate"):
                continue
            slot_status = str(row.get("slot_status") or row.get("status") or "")
            if slot_status in ("Ignored", "None"):
                continue
            score = self._download_target_score(row, download, score_cache)
            if score > best_score:
                best_score = score
                best_row = row
        return best_row, best_score

    def _download_target_score(
        self,
        row: dict,
        download: dict,
        score_cache: dict[tuple[str, str], int] | None = None,
    ) -> int:
        base_name = str(row.get("base_mod") or "")
        download_name = str(download.get("download_name") or "")
        cache_key = (base_name, download_name)
        if score_cache is not None and cache_key in score_cache:
            score = score_cache[cache_key]
        else:
            score = voice_match_score(base_name, download_name)
            if score_cache is not None:
                score_cache[cache_key] = score
        download_category = str(download.get("voice_category") or "")
        row_category = str(row.get("voice_category") or "")
        if download_category and row_category:
            if download_category == row_category:
                score += 20
            else:
                return 0
        elif row_category:
            score = max(0, score - 35)
        return score

    def _target_row_for_download(self, rows: list[dict], candidate: dict) -> dict | None:
        base = str(candidate.get("target_base_internal_name") or candidate.get("target_base_mod") or "")
        category = str(candidate.get("target_voice_category") or "")
        if not base or not category:
            return None
        for row in rows:
            row_base = str(row.get("base_internal_name") or row.get("base_mod") or "")
            if row_base == base and str(row.get("voice_category") or "") == category:
                return row
        return None

    def _downloaded_archive_names(self) -> set[str]:
        downloads_path = self._mo2_downloads_path()
        if not downloads_path or not downloads_path.exists():
            return set()
        try:
            return {path.name.lower() for path in downloads_path.iterdir() if path.is_file()}
        except Exception:
            return set()

    def _download_already_status(self, candidate: dict, voice_mods: list[dict], archive_names: set[str]) -> str:
        download_name = str(candidate.get("download_name") or "").strip()
        if not download_name:
            return ""
        archive_name = safe_archive_name(download_name).lower()
        if archive_name in archive_names:
            return "Downloaded"

        best_installed = None
        best_installed_score = 0
        for voice in voice_mods:
            score = voice_match_score(voice.get("display_name") or "", download_name)
            if score > best_installed_score:
                best_installed_score = score
                best_installed = voice
        if best_installed and best_installed_score >= 90:
            return f"Installed: {best_installed.get('display_name') or ''}"

        best_archive = ""
        best_archive_score = 0
        for archive in archive_names:
            if not archive.endswith((".7z", ".zip", ".rar")):
                continue
            score = voice_match_score(archive, download_name)
            if score > best_archive_score:
                best_archive_score = score
                best_archive = archive
        if best_archive and best_archive_score >= 95:
            return f"Downloaded: {best_archive}"
        return ""

    def _hide_selected_online_match(self, dialog: QDialog, table: QTableWidget, row: dict) -> None:
        candidate = self._candidate_from_online_table(table)
        if not candidate:
            QMessageBox.information(dialog, PLUGIN_NAME, "Select an online candidate first.")
            return

        item = {
            "base": str(row.get("base_internal_name") or row.get("base_mod") or "").lower(),
            "voice_category": str(row.get("voice_category") or "").lower(),
            "candidate": str(candidate.get("download_name") or "").lower(),
            "source_url": str(candidate.get("source_url") or "").lower(),
        }
        if not item["base"] or not item["candidate"]:
            QMessageBox.information(dialog, PLUGIN_NAME, "This online candidate cannot be hidden.")
            return

        confirm = QMessageBox.question(
            dialog,
            PLUGIN_NAME,
            "Hide this online candidate from future source fetches?\n\n"
            f"Base: {row.get('base_mod') or ''}\n"
            f"Voice type: {row.get('voice_category_label') or ''}\n"
            f"Candidate: {candidate.get('download_name') or ''}\n\n"
            "You can restore it later with the False matches button.",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        config = self._load_voice_config()
        if item not in config["falseMatches"]:
            config["falseMatches"].append(item)
        self._save_voice_config(config)

        candidates = getattr(table, "_ll_voice_candidates", [])
        index = table.selectedItems()[0].row() if table.selectedItems() else -1
        if 0 <= index < len(candidates):
            candidates.pop(index)
            table.removeRow(index)
            if candidates:
                table.selectRow(min(index, len(candidates) - 1))
        row["online_candidates"] = candidates
        QMessageBox.information(dialog, PLUGIN_NAME, "Online false match saved.")

    def _start_online_candidate_download(
        self,
        dialog: QDialog,
        row: dict,
        candidate: dict,
        downloads_path: Path,
        status_label: QLabel,
        main_progress_label: QLabel,
        download_button: QPushButton,
        close_button: QPushButton,
        main_download_button: QPushButton,
        main_fetch_button: QPushButton,
    ) -> None:
        candidate = {
            **candidate,
            "base_mod": row.get("base_mod") or "",
            "base_internal_name": row.get("base_internal_name") or "",
            "base_page_url": row.get("base_page_url") or "",
            "voice_category": row.get("voice_category") or candidate.get("voice_category") or "",
            "voice_category_label": row.get("voice_category_label") or candidate.get("voice_category_label") or "",
        }

        thread = QThread(dialog)
        worker = VoiceCandidateDownloadWorker(
            candidate,
            downloads_path,
            self._setting_path("cookies_path", DEFAULT_COOKIES),
        )
        worker.moveToThread(thread)
        download_button.setEnabled(False)
        close_button.setEnabled(False)
        main_download_button.setEnabled(False)
        main_fetch_button.setEnabled(False)
        status_label.setText("Starting download...")
        main_progress_label.setText("Starting voice candidate download...")
        worker.statusChanged.connect(status_label.setText)
        worker.statusChanged.connect(main_progress_label.setText)
        worker.finished.connect(
            lambda ok, message, result: self._online_candidate_download_finished(
                dialog,
                thread,
                worker,
                status_label,
                main_progress_label,
                download_button,
                close_button,
                main_download_button,
                main_fetch_button,
                ok,
                message,
                dict(result or {}),
            )
        )
        thread.started.connect(worker.run)
        thread.start()

    def _online_candidate_download_finished(
        self,
        dialog: QDialog,
        thread: QThread,
        worker: VoiceCandidateDownloadWorker,
        status_label: QLabel,
        main_progress_label: QLabel,
        download_button: QPushButton,
        close_button: QPushButton,
        main_download_button: QPushButton,
        main_fetch_button: QPushButton,
        ok: bool,
        message: str,
        result: dict,
    ) -> None:
        download_button.setEnabled(True)
        close_button.setEnabled(True)
        main_download_button.setEnabled(True)
        main_fetch_button.setEnabled(True)
        thread.quit()
        thread.wait()
        worker.deleteLater()
        thread.deleteLater()
        if ok:
            text = f"{message} {result.get('archive_path') or ''}".strip()
            status_label.setText(text)
            main_progress_label.setText(text)
        else:
            status_label.setText(f"Download failed: {message}")
            main_progress_label.setText("Download failed")
            QMessageBox.warning(dialog, PLUGIN_NAME, f"Voice candidate download failed:\n\n{message}")

    def _choose_online_candidate(self, parent: QDialog, row: dict) -> dict | None:
        candidates = list(row.get("online_candidates") or [])
        if not candidates and row.get("online_download_url"):
            candidates = [{
                "download_name": row.get("online_candidate") or "",
                "download_url": row.get("online_download_url") or "",
                "source_url": row.get("online_source_url") or "",
                "online_score": row.get("online_score") or 0,
                "size": row.get("online_size") or "",
                "date_iso": row.get("online_date_iso") or "",
                "version": row.get("online_version") or "",
            }]

        candidates = sorted(
            candidates,
            key=lambda item: int(item.get("online_score") or 0),
            reverse=True,
        )
        if not candidates:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Selected row has no online candidate yet.")
            return None

        dialog = QDialog(parent)
        dialog.setWindowTitle(f"Download voice candidate - {row.get('base_mod') or ''}")
        dialog.resize(920, 460)

        table = QTableWidget(dialog)
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["Score", "Candidate", "Size", "Date", "Source"])
        table.setRowCount(len(candidates))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table._ll_voice_candidates = candidates

        for index, candidate in enumerate(candidates):
            score = int(candidate.get("online_score") or 0)
            for column, value in enumerate([
                str(score or ""),
                candidate.get("download_name") or "",
                candidate.get("size") or "",
                candidate.get("date_iso") or "",
                candidate.get("source_url") or "",
            ]):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(str(value or ""))
                item.setBackground(self._candidate_background_color(score))
                item.setForeground(QColor(242, 242, 242))
                table.setItem(index, column, item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Stretch)
        if candidates:
            table.selectRow(0)

        title = QLabel(f"Base mod: {row.get('base_mod') or ''}")
        title.setStyleSheet("font-weight: 700;")
        hint = QLabel("Candidates are sorted by score, highest first. Choose the archive to download.")
        hint.setWordWrap(True)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("Download selected")
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        table.itemDoubleClicked.connect(lambda _item: dialog.accept())

        layout = QVBoxLayout(dialog)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addWidget(table)
        layout.addWidget(buttons)
        dialog.setLayout(layout)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        selected = table.selectedItems()
        if not selected:
            return None
        index = selected[0].row()
        if index < 0 or index >= len(candidates):
            return None
        return dict(candidates[index])

    def _download_candidate_finished(
        self,
        dialog: QDialog,
        thread: QThread,
        worker: VoiceCandidateDownloadWorker,
        progress_label: QLabel,
        download_button: QPushButton,
        fetch_button: QPushButton,
        ok: bool,
        message: str,
        result: dict,
    ) -> None:
        download_button.setEnabled(True)
        fetch_button.setEnabled(True)
        thread.quit()
        thread.wait()
        worker.deleteLater()
        thread.deleteLater()
        if ok:
            progress_label.setText(f"{message} {result.get('archive_path') or ''}")
        else:
            progress_label.setText("Download failed")
            QMessageBox.warning(dialog, PLUGIN_NAME, f"Voice candidate download failed:\n\n{message}")

    def _mo2_downloads_path(self) -> Path | None:
        if self._organizer:
            try:
                path = Path(str(self._organizer.downloadsPath()))
                if str(path):
                    return path
            except Exception:
                pass

        config = self._read_native_config(self._native_config_path())
        downloads = config.get("mo2_downloads_path")
        return Path(str(downloads)) if downloads else None

    def _mark_false_match(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        candidate = row.get("installed_voice") or ""
        source_url = ""
        if not candidate:
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "Selected row has no installed/local voice candidate to mark.\n\n"
                "Use Choose / Download > Hide online match for online candidates.",
            )
            return

        confirm = QMessageBox.question(
            self._parentWidget(),
            PLUGIN_NAME,
            "Hide this candidate from future voice scans?\n\n"
            f"Base: {row.get('base_mod') or ''}\n"
            f"Voice type: {row.get('voice_category_label') or ''}\n"
            f"Candidate: {candidate}\n\n"
            "This only affects installed/local voice matching. You can restore it later with the False matches button.",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        config = self._load_voice_config()
        item = {
            "base": str(row.get("base_internal_name") or row.get("base_mod") or "").lower(),
            "voice_category": str(row.get("voice_category") or "").lower(),
            "candidate": str(candidate).lower(),
            "source_url": str(source_url).lower(),
        }
        if item not in config["falseMatches"]:
            config["falseMatches"].append(item)
        manual_matches = dict(config.get("manualVoiceMatches") or {})
        manual_matches.pop(str(row.get("slot_key") or "").lower(), None)
        config["manualVoiceMatches"] = manual_matches
        self._save_voice_config(config)

        row["online_candidate"] = ""
        row["online_download_url"] = ""
        row["online_source_url"] = ""
        row["online_score"] = 0
        if candidate == row.get("installed_voice"):
            row["installed_voice"] = ""
            row["installed_voice_internal_name"] = ""
            row["score"] = 0
            row["manual_voice"] = False
            row["slot_status"] = "Missing"
            row["status"] = "Missing"
        self._refresh_voice_overviews(table)
        self._apply_filter(table, filter_mode, filter_text, count_label)

    def _manage_false_matches(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        config = self._load_voice_config()
        false_matches = list(config.get("falseMatches") or [])
        if not false_matches:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "No false matches are saved.")
            return

        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("False voice matches")
        dialog.resize(860, 420)

        items = QListWidget(dialog)
        for item in false_matches:
            base = item.get("base") or ""
            candidate = item.get("candidate") or ""
            source_url = item.get("source_url") or ""
            category = item.get("voice_category") or ""
            label = f"{base}"
            if category:
                label = f"{label} [{VOICE_CATEGORY_LABELS.get(category, category)}]"
            label = f"{label} -> {candidate}"
            if source_url:
                label = f"{label} | {source_url}"
            items.addItem(label)

        hint = QLabel("Remove selected entries to let them appear again after the next Fetch sources / scan.")
        hint.setWordWrap(True)
        remove_selected = QPushButton("Remove selected")
        close = QPushButton("Close")

        def remove_items() -> None:
            selected = sorted((item.row() for item in items.selectedIndexes()), reverse=True)
            if not selected:
                QMessageBox.information(dialog, PLUGIN_NAME, "Select a false match first.")
                return
            for index in selected:
                if 0 <= index < len(false_matches):
                    false_matches.pop(index)
                    items.takeItem(index)
            config["falseMatches"] = false_matches
            self._save_voice_config(config)
            self._apply_filter(table, filter_mode, filter_text, count_label)

        remove_selected.clicked.connect(lambda _checked=False: remove_items())
        close.clicked.connect(dialog.accept)

        buttons = QHBoxLayout()
        buttons.addWidget(remove_selected)
        buttons.addStretch(1)
        buttons.addWidget(close)

        layout = QVBoxLayout(dialog)
        layout.addWidget(hint)
        layout.addWidget(items)
        layout.addLayout(buttons)
        dialog.setLayout(layout)
        dialog.exec()

    def _toggle_complete_slot(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return
        mode = "reopen" if (row.get("slot_status") or row.get("status")) in ("Complete", "None") else "complete"
        self._set_slot_flag(table, row, mode, filter_mode, filter_text, count_label)

    def _toggle_ignore(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        key = str(row.get("base_internal_name") or row.get("base_mod") or "").lower()
        config = self._load_voice_config()
        ignored = {str(value).lower() for value in config.get("ignoredBaseMods", [])}
        if key in ignored:
            ignored.remove(key)
            row["slot_status"] = self._voice_row_status(
                False,
                bool(row.get("complete_voice")),
                bool(row.get("none_voice")),
                bool(row.get("manual_voice")),
                int(row.get("score") or 0),
                bool(row.get("online_candidate")),
                self._local_match_threshold(config),
            )
            row["status"] = row["slot_status"]
        else:
            ignored.add(key)
            row["slot_status"] = "Ignored"
            row["status"] = "Ignored"
        config["ignoredBaseMods"] = sorted(ignored)
        self._save_voice_config(config)
        self._refresh_voice_overviews(table)
        self._apply_filter(table, filter_mode, filter_text, count_label)

    def _classify_selected_mod(
        self,
        table: QTableWidget,
        filter_mode: QComboBox,
        filter_text: QLineEdit,
        count_label: QLabel,
    ) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        key = str(row.get("base_internal_name") or row.get("base_mod") or "").lower()
        if not key:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Selected row cannot be classified.")
            return

        current = str(row.get("classification_override") or "auto")
        labels = [
            "Auto",
            "Base mod",
            "Voice pack",
        ]
        current_index = {"auto": 0, "base": 1, "voice": 2}.get(current, 0)
        selected, ok = QInputDialog.getItem(
            self._parentWidget(),
            "Classify LL mod",
            f"{row.get('base_mod') or ''}\n\nTreat this mod as:",
            labels,
            current_index,
            False,
        )
        if not ok:
            return

        config = self._load_voice_config()
        forced_voice = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base = {str(value).lower() for value in config.get("forcedBaseMods", [])}

        forced_voice.discard(key)
        forced_base.discard(key)

        if selected == "Voice pack":
            forced_voice.add(key)
        elif selected == "Base mod":
            forced_base.add(key)

        config["forcedVoiceMods"] = sorted(forced_voice)
        config["forcedBaseMods"] = sorted(forced_base)
        self._save_voice_config(config)

        row_index = self._selected_table_index(table)
        if selected == "Voice pack":
            rows = getattr(table, "_ll_voice_rows", [])
            if 0 <= row_index < len(rows):
                rows.pop(row_index)
                table.removeRow(row_index)
            self._apply_filter(table, filter_mode, filter_text, count_label)
            QMessageBox.information(
                self._parentWidget(),
                PLUGIN_NAME,
                "Saved as Voice pack.\n\nIt was removed from the base-mod list and will be used as a voice candidate on the next scan.",
            )
            return

        row["classification_override"] = "base" if selected == "Base mod" else "auto"
        self._fill_table_row(table, row_index, row)
        self._apply_filter(table, filter_mode, filter_text, count_label)

    def _show_voice_mods_inventory(self, parent: QDialog) -> None:
        voice_mods = self._collect_voice_mods_inventory(show_message=True)

        if not voice_mods:
            return

        dialog = QDialog(parent)
        dialog.setWindowTitle("Installed voice mods")
        dialog.resize(940, 520)

        panel = self._build_voice_inventory_panel(dialog, voice_mods)
        close = QPushButton("Close")
        close.clicked.connect(dialog.accept)

        layout = QVBoxLayout(dialog)
        layout.addWidget(panel)
        layout.addWidget(close)
        dialog.setLayout(layout)
        dialog.exec()

    def _collect_voice_mods_inventory(self, show_message: bool = True) -> list[dict]:
        if not self._organizer:
            if show_message:
                QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "MO2 organizer is not available.")
            return []

        config = self._load_voice_config()
        forced_voice_mods = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base_mods = {str(value).lower() for value in config.get("forcedBaseMods", [])}
        mod_list = self._organizer.modList()
        voice_mods = []
        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is None:
                continue
            key = str(internal_name).lower()
            display_name = mod_list.displayName(internal_name)
            is_voice = (voice_keyword_present(display_name) or key in forced_voice_mods) and key not in forced_base_mods
            if not is_voice:
                continue
            path = mod_root_path(mod)
            meta_path = path / MOD_META_FILE
            install_time = meta_path.stat().st_mtime if meta_path.exists() else path.stat().st_mtime if path.exists() else 0
            voice_mods.append({
                "display_name": display_name,
                "internal_name": internal_name,
                "mod_path": str(path),
                "installed_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(install_time)) if install_time else "",
                "install_time": install_time,
                "classification": "Forced voice" if key in forced_voice_mods else "Auto voice",
            })

        if show_message and not voice_mods:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "No installed voice-like mods were found.")
        return voice_mods

    def _build_voice_inventory_panel(self, parent: QDialog, voice_mods: list[dict]):
        panel = QDialog(parent)
        panel.setWindowFlags(Qt.WindowType.Widget)
        table = QTableWidget(panel)
        table.setColumnCount(2)
        table.setHorizontalHeaderLabels(["Voice mod", "Installed"])
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table._ll_voice_inventory = voice_mods

        filter_text = QLineEdit(panel)
        filter_text.setPlaceholderText("Filter voice mod / folder")
        filter_text.setClearButtonEnabled(True)
        sort_mode = QComboBox(panel)
        sort_mode.addItems(["Name A-Z", "Name Z-A", "Newest first", "Oldest first"])
        count_label = QLabel(panel)

        def populate() -> None:
            needle = filter_text.text().strip().lower()
            sort_text = sort_mode.currentText()
            rows = [
                item for item in voice_mods
                if not needle or needle in " ".join([
                    item.get("display_name") or "",
                    item.get("internal_name") or "",
                    item.get("mod_path") or "",
                    item.get("classification") or "",
                ]).lower()
            ]
            if sort_text == "Name Z-A":
                rows.sort(key=lambda item: str(item.get("display_name") or "").lower(), reverse=True)
            elif sort_text == "Newest first":
                rows.sort(key=lambda item: float(item.get("install_time") or 0), reverse=True)
            elif sort_text == "Oldest first":
                rows.sort(key=lambda item: float(item.get("install_time") or 0))
            else:
                rows.sort(key=lambda item: str(item.get("display_name") or "").lower())

            table.setRowCount(len(rows))
            table._ll_voice_inventory_visible = rows
            for row_index, item in enumerate(rows):
                for column, value in enumerate([
                    item.get("display_name") or "",
                    item.get("installed_at") or "",
                ]):
                    cell = QTableWidgetItem(str(value or ""))
                    cell.setFlags(cell.flags() & ~Qt.ItemFlag.ItemIsEditable)
                    cell.setToolTip(self._voice_inventory_tooltip(item))
                    table.setItem(row_index, column, cell)
            count_label.setText(f"{len(rows)} / {len(voice_mods)}")
            if rows:
                table.selectRow(0)

        def open_selected() -> None:
            selected = table.selectedItems()
            rows = getattr(table, "_ll_voice_inventory_visible", [])
            if not selected:
                return
            index = selected[0].row()
            if index < 0 or index >= len(rows):
                return
            path = Path(str(rows[index].get("mod_path") or ""))
            if path.exists():
                self._open_path(path)

        filter_text.textChanged.connect(lambda _text: populate())
        sort_mode.currentTextChanged.connect(lambda _text: populate())
        table.itemDoubleClicked.connect(lambda _item: open_selected())
        table.customContextMenuRequested.connect(lambda pos: self._show_voice_inventory_context_menu(table, pos))

        open_folder = QPushButton("Open selected folder")
        open_folder.clicked.connect(lambda _checked=False: open_selected())

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Filter"))
        controls.addWidget(filter_text, 1)
        controls.addWidget(QLabel("Sort"))
        controls.addWidget(sort_mode)
        controls.addWidget(count_label)

        buttons = QHBoxLayout()
        buttons.addWidget(open_folder)
        buttons.addStretch(1)

        layout = QVBoxLayout(panel)
        title = QLabel("Installed voice mods")
        title.setStyleSheet("font-weight: 700;")
        layout.addWidget(title)
        layout.addLayout(controls)
        layout.addWidget(table)
        layout.addLayout(buttons)
        panel.setLayout(layout)
        populate()
        return panel

    def _voice_inventory_tooltip(self, item: dict) -> str:
        return "\n".join([
            item.get("display_name") or "",
            f"Installed: {item.get('installed_at') or ''}",
            f"Class: {item.get('classification') or ''}",
            f"Internal: {item.get('internal_name') or ''}",
            f"Folder: {item.get('mod_path') or ''}",
        ]).strip()

    def _show_voice_inventory_context_menu(self, table: QTableWidget, pos) -> None:
        selected = table.selectedItems()
        rows = getattr(table, "_ll_voice_inventory_visible", [])
        item = table.itemAt(pos)
        if item is None:
            return
        table.selectRow(item.row())
        index = item.row()
        if index < 0 or index >= len(rows):
            return
        voice = rows[index]

        menu = QMenu(table)
        title = QAction(voice.get("display_name") or "", menu)
        title.setEnabled(False)
        menu.addAction(title)
        menu.addSeparator()
        for label, value in [
            ("Installed", voice.get("installed_at") or ""),
            ("Class", voice.get("classification") or ""),
            ("Internal", voice.get("internal_name") or ""),
            ("Folder", voice.get("mod_path") or ""),
        ]:
            action = QAction(f"{label}: {value}", menu)
            action.setEnabled(False)
            menu.addAction(action)
        menu.addSeparator()
        open_folder = QAction("Open folder", menu)
        open_folder.triggered.connect(lambda _checked=False, path=voice.get("mod_path") or "": self._open_path(Path(path)) if path and Path(path).exists() else None)
        menu.addAction(open_folder)
        menu.exec(table.viewport().mapToGlobal(pos))

    def _selected_table_index(self, table: QTableWidget) -> int:
        selected = table.selectedItems()
        return selected[0].row() if selected else -1

    def _search_selected_voice(self, table: QTableWidget) -> None:
        row = self._selected_row(table)
        if not row:
            QMessageBox.information(self._parentWidget(), PLUGIN_NAME, "Select a row first.")
            return

        query = row.get("search_query") or row.get("base_mod") or ""
        if not query:
            return

        url = "https://www.google.com/search?q=" + quote_plus(
            f"site:loverslab.com/files/file/ {query}"
        )
        webbrowser.open(url)

class LoversLabPathsTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Paths"
    TOOL_DISPLAY = "LL Integration Paths"
    TOOL_DESCRIPTION = "Shows LL Integration install, cookies, metadata, and MO2 paths."

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def display(self) -> None:
        paths = self._paths()
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("LL Integration Paths")
        dialog.resize(760, 280)

        layout = QGridLayout(dialog)
        for row, (label, path) in enumerate(paths.items()):
            layout.addWidget(QLabel(label), row, 0)
            value = QLabel(str(path))
            value.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)
            layout.addWidget(value, row, 1)
            button = QPushButton("Open")
            button.setEnabled(path.exists())
            button.clicked.connect(lambda _checked=False, target=path: self._open_path(target))
            layout.addWidget(button, row, 2)

        close = QDialogButtonBox(QDialogButtonBox.StandardButton.Close)
        close.rejected.connect(dialog.reject)
        layout.addWidget(close, len(paths), 0, 1, 3)
        dialog.setLayout(layout)
        dialog.exec()

    def _paths(self) -> dict[str, Path]:
        paths = self._configured_paths()
        cookies = Path(str(paths.get("cookies_path") or DEFAULT_COOKIES))
        native_app = cookies.parents[1] if len(cookies.parents) > 1 else cookies.parent
        config = native_app / "config.json"
        native_config = self._read_native_config(config)
        metadata = Path(str(native_config.get("metadata_path") or native_app / "metadata"))
        downloads = Path(str(native_config.get("mo2_downloads_path") or ""))
        plugin = Path(__file__).resolve().parent
        fetch_log = (
            Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "fetch_update_timings.jsonl"
            if self._organizer
            else plugin / "fetch_update_timings.jsonl"
        )

        return {
            "Integration folder": native_app.parent,
            "Native app": native_app,
            "Config": config,
            "Cookies": cookies.parent,
            "Metadata": metadata,
            "MO2 downloads": downloads,
            "MO2 plugin": plugin,
            "Fetch timing log": fetch_log,
        }

    def _read_native_config(self, path: Path) -> dict:
        if not path.exists():
            return {}
        try:
            data = json.loads(path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _open_path(self, path: Path) -> None:
        target = path if path.is_dir() else path.parent
        webbrowser.open(str(target))


class LoversLabPurgeSuspiciousLinksTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Purge Suspicious Links"
    TOOL_DISPLAY = "Purge Suspicious LoversLab Links"
    TOOL_DESCRIPTION = "Removes LL metadata and clears LL meta URLs from mods that MO2 identifies as Nexus mods."

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def display(self) -> None:
        try:
            candidates = self._find_candidates()
            if not candidates:
                QMessageBox.information(
                    self._parentWidget(),
                    PLUGIN_NAME,
                    "No suspicious LoversLab links found.",
                )
                return

            if not self._confirm_purge(candidates):
                return

            purged = self._purge(candidates)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Purge failed:\n\n{exc}",
            )
            return

        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            "Purged suspicious LoversLab links:\n\n" + "\n".join(purged),
        )

    def _find_candidates(self) -> list[dict]:
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        candidates = []
        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is None:
                continue

            ll_metadata = mod_ll_metadata_path(mod, migrate_legacy=False)
            general = read_mod_meta_general(mod)
            meta_url = general.get("url", "").lower()
            has_stale_ll_meta = "loverslab.com" in meta_url
            if ll_metadata is None and not has_stale_ll_meta:
                continue

            is_nexus, reason = mod_has_purgeable_nexus_identity(mod)
            if not is_nexus:
                continue

            candidates.append({
                "mod_name": mod_list.displayName(internal_name),
                "mod": mod,
                "ll_metadata": ll_metadata,
                "reason": reason,
            })

        return candidates

    def _confirm_purge(self, candidates: list[dict]) -> bool:
        shown = candidates[:40]
        lines = [f"{item['mod_name']} ({item['reason']})" for item in shown]
        if len(candidates) > len(shown):
            lines.append(f"...and {len(candidates) - len(shown)} more")

        result = QMessageBox.question(
            self._parentWidget(),
            PLUGIN_NAME,
            "Clean suspicious LoversLab links from these Nexus-linked mods?\n\n"
            "LL Integration metadata and stale LoversLab URLs will be cleared from meta.ini. "
            "Legacy LL.ini files will be deleted.\n\n"
            + "\n".join(lines),
        )
        return result == QMessageBox.StandardButton.Yes

    def _purge(self, candidates: list[dict]) -> list[str]:
        purged = []
        for item in candidates:
            actions = remove_mod_ll_metadata(item["mod"])
            if cleanup_loverslab_meta(item["mod"]):
                actions.append("cleaned meta.ini URL")
            purged.append(f"{item['mod_name']} -> {', '.join(actions) if actions else 'nothing changed'}")

        return purged


class LoversLabCreateLinkTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Create Link"
    TOOL_DISPLAY = "Create Source Link"
    TOOL_DESCRIPTION = "Creates source metadata manually for LoversLab, Patreon, or other external mod pages."

    def icon(self) -> QIcon:
        return QIcon(str(Path(__file__).resolve().parent / "icons" / "ll_check_all.svg"))

    def display(self) -> None:
        try:
            mod_name, mod = self._choose_mod_filtered()
            values = self._prompt_link_values(mod_name)
            if not values:
                return

            target = write_mod_ll_metadata_from_text(mod, self._ll_ini_text(values))
            self._write_mo2_meta_ini(mod, values["page_url"], values["version"])
            if self._organizer:
                self._organizer.modDataChanged(mod)
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Create link failed:\n\n{exc}",
            )
            return

        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            f"Created source link for:\n{mod_name}\n\nStored in:\n{target}",
        )

    def _choose_mod_filtered(self):
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        mods = []
        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is not None:
                mods.append((mod_list.displayName(internal_name), mod))

        if not mods:
            raise RuntimeError("No installed mods found")

        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("Choose Mod")
        dialog.resize(520, 520)

        filter_box = QLineEdit(dialog)
        filter_box.setPlaceholderText("Filter mods...")
        list_widget = QListWidget(dialog)

        def refill() -> None:
            text = filter_box.text().strip().lower()
            list_widget.clear()
            for name, _mod in mods:
                if not text or text in name.lower():
                    list_widget.addItem(name)
            if list_widget.count() > 0:
                list_widget.setCurrentRow(0)

        filter_box.textChanged.connect(refill)
        refill()

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        list_widget.itemDoubleClicked.connect(lambda _item: dialog.accept())

        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("Choose the installed mod:"))
        layout.addWidget(filter_box)
        layout.addWidget(list_widget)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            raise RuntimeError("No mod selected")

        item = list_widget.currentItem()
        if item is None:
            raise RuntimeError("No mod selected")

        selected = item.text()
        for name, mod in mods:
            if name == selected:
                return name, mod

        raise RuntimeError(f"Could not open mod: {selected}")

    def _prompt_link_values(self, mod_name: str) -> dict | None:
        dialog = QDialog(self._parentWidget())
        dialog.setWindowTitle("Create Source Link")
        dialog.resize(700, 260)

        url = QLineEdit(dialog)
        version = QLineEdit(dialog)
        file_pattern = QLineEdit(dialog)
        url.setPlaceholderText("https://www.loverslab.com/files/file/... or https://www.patreon.com/posts/...")
        version.setPlaceholderText("Installed version, for example 1.2.3 or 5-60")
        file_pattern.setPlaceholderText("Optional, for example Example Mod 5-* - FULL*")
        multipart = QCheckBox("Multipart or manual install", dialog)
        manual = QCheckBox("Manual source link", dialog)
        update_mode = QComboBox(dialog)
        configure_update_mode_combo(update_mode, UPDATE_MODE_MANUAL)
        multipart.setChecked(True)
        manual.setChecked(True)

        layout = QGridLayout(dialog)
        layout.addWidget(QLabel(f"Mod: {mod_name}"), 0, 0, 1, 2)
        layout.addWidget(QLabel("Source page URL"), 1, 0)
        layout.addWidget(url, 1, 1)
        layout.addWidget(QLabel("Current version"), 2, 0)
        layout.addWidget(version, 2, 1)
        layout.addWidget(QLabel("Download file pattern (optional)"), 3, 0)
        layout.addWidget(file_pattern, 3, 1)
        layout.addWidget(multipart, 4, 1)
        layout.addWidget(manual, 5, 1)
        layout.addWidget(QLabel("Update mode"), 6, 0)
        layout.addWidget(update_mode, 6, 1)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        layout.addWidget(buttons, 7, 0, 1, 2)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        page_url = url.text().strip()
        current_version = version.text().strip()
        if not page_url:
            raise RuntimeError("Source page URL is required")
        if not current_version:
            raise RuntimeError("Current version is required")

        return {
            "page_url": page_url,
            "version": current_version,
            "file_pattern": file_pattern.text().strip(),
            "multipart": multipart.isChecked(),
            "manual_install": manual.isChecked(),
            "update_mode": normalized_update_mode(str(update_mode.currentData() or "")),
        }

    def _ll_ini_text(self, values: dict) -> str:
        ll_id_match = re.search(r"/files/file/(\d+)", values["page_url"])
        ll_id = ll_id_match.group(1) if ll_id_match else ""
        is_loverslab = "loverslab.com" in values["page_url"].lower()
        update_mode = normalized_update_mode(values.get("update_mode"))
        if not is_loverslab:
            update_mode = UPDATE_MODE_SKIP
        skip_updates = update_mode == UPDATE_MODE_SKIP
        lines = [
            "[LoversLab]",
            f"source={'loverslab' if is_loverslab else 'external'}",
            f"ll_file_id={ll_id}",
            "ll_resource_id=",
            f"page_url={values['page_url']}",
            "page_title=",
            "download_url=",
            f"file_name={values.get('file_pattern') or ''}",
            f"original_archive_name={values.get('file_pattern') or ''}",
            f"archive_name={values.get('file_pattern') or ''}",
            "archive_size_bytes=",
            "archive_quick_hash=",
            f"version={values['version']}",
            "size=",
            "date_iso=",
            "captured_at=",
            "archive_path=",
            "browser_download_url=",
            "completed_at=",
            f"update_mode={update_mode}",
            f"fixed_version={str(skip_updates).lower()}",
            f"manual_install={str(values['manual_install']).lower()}",
            f"manual_update={str(skip_updates).lower()}",
            f"skip_update_check={str(skip_updates).lower()}",
            f"multipart={str(values['multipart']).lower()}",
            f"file_pattern={values.get('file_pattern') or ''}",
            "",
        ]
        return "\n".join(lines)


class LoversLabAutoBindTool(LoversLabBaseTool):
    TOOL_NAME = "LL Integration Auto Bind"
    TOOL_DISPLAY = "Auto Bind LoversLab Metadata"
    TOOL_DESCRIPTION = "Finds likely LoversLab metadata matches for installed mods and binds them with confirmation."

    MIN_SCORE = 5

    def display(self) -> None:
        try:
            candidates = self._find_candidates()
            if not candidates:
                QMessageBox.information(
                    self._parentWidget(),
                    PLUGIN_NAME,
                    "No safe LoversLab bind candidates found.",
                )
                return

            if not self._confirm_candidates(candidates):
                return

            bound = []
            for candidate in candidates:
                if mod_ll_metadata_path(candidate["mod"]):
                    continue
                target = write_mod_ll_metadata_from_file(candidate["mod"], candidate["sidecar"])
                try:
                    LoversLabBindLatestTool._apply_mod_metadata(self, candidate["mod"], target)
                    if self._organizer:
                        self._organizer.modDataChanged(candidate["mod"])
                except Exception:
                    pass
                bound.append(f"{candidate['mod_name']} <- {candidate['archive_name']}")
        except Exception as exc:
            QMessageBox.critical(
                self._parentWidget(),
                PLUGIN_NAME,
                f"Auto bind failed:\n\n{exc}",
            )
            return

        QMessageBox.information(
            self._parentWidget(),
            PLUGIN_NAME,
            "Bound LoversLab metadata:\n\n" + "\n".join(bound),
        )

    def _find_candidates(self) -> list[dict]:
        sidecars = self._metadata_sidecars()
        mods = self._mods_without_ll_metadata()
        candidates = []
        used_mods = set()

        for sidecar in sidecars:
            info = self._read_ll_info(sidecar)
            archive_name = info.get("archive_name") or info.get("file_name") or sidecar.name
            match = self._best_match(archive_name, mods, used_mods)
            if not match:
                continue

            score, mod_name, mod = match
            if score < self.MIN_SCORE:
                continue

            candidates.append({
                "sidecar": sidecar,
                "archive_name": archive_name,
                "version": info.get("version") or "",
                "mod_name": mod_name,
                "mod": mod,
                "score": score,
            })
            used_mods.add(mod_name)

        return candidates

    def _metadata_sidecars(self) -> list[Path]:
        metadata_path = self._metadata_downloads_path()
        sidecars = list(metadata_path.glob("*.ll.ini")) if metadata_path.exists() else []

        completions_path = self._downloads_storage_path() / "download_completions.json"
        if completions_path.exists():
            completions = json.loads(completions_path.read_text(encoding="utf-8"))
            if isinstance(completions, list):
                for completion in completions:
                    archive = completion.get("archivePath")
                    if archive:
                        sidecar = Path(archive).with_name(f"{Path(archive).name}.ll.ini")
                        if sidecar.exists():
                            sidecars.append(sidecar)

        unique = {str(path).lower(): path for path in sidecars}
        return sorted(unique.values(), key=lambda path: path.stat().st_mtime, reverse=True)

    def _metadata_downloads_path(self) -> Path:
        config_path = self._native_config_path()
        if config_path.exists():
            try:
                config = json.loads(config_path.read_text(encoding="utf-8-sig"))
                metadata = config.get("metadata_path")
                if metadata:
                    return Path(str(metadata)) / "downloads"
            except Exception:
                pass

        return self._downloads_storage_path().parent / "metadata" / "downloads"

    def _native_config_path(self) -> Path:
        return self._downloads_storage_path().parent / "config.json"

    def _mods_without_ll_metadata(self) -> list[tuple[str, object]]:
        if not self._organizer:
            raise RuntimeError("MO2 organizer is not available")

        mod_list = self._organizer.modList()
        mods = []
        for internal_name in mod_list.allModsByProfilePriority():
            mod = mod_list.getMod(internal_name)
            if mod is None:
                continue
            display_name = mod_list.displayName(internal_name)
            if mod_ll_metadata_path(mod):
                continue
            if mod_has_nexus_identity(mod)[0]:
                continue
            mods.append((display_name, mod))
        return mods

    def _best_match(self, archive_name: str, mods: list[tuple[str, object]], used_mods: set[str]):
        archive_tokens = self._tokens(archive_name)
        archive_acronym = self._acronym(archive_name)
        best = None

        for mod_name, mod in mods:
            if mod_name in used_mods:
                continue

            mod_tokens = self._tokens(mod_name)
            shared_tokens = archive_tokens & mod_tokens
            score = len(shared_tokens) * 3
            mod_acronym = self._acronym(mod_name)
            compact_archive = self._compact(archive_name)
            compact_mod = self._compact(mod_name)

            if archive_acronym and mod_acronym and archive_acronym == mod_acronym:
                score += 5
            if archive_acronym and archive_acronym in compact_mod:
                score += 4
            if mod_acronym and mod_acronym in compact_archive:
                score += 4
            if compact_archive and compact_archive in compact_mod:
                score += 3

            if best is None or score > best[0]:
                best = (score, mod_name, mod)

        return best

    def _acronym(self, text: str) -> str:
        stop_words = {"the", "and", "with", "for"}
        tokens = [token for token in self._tokens(text) if token not in stop_words]
        if len(tokens) == 1:
            token = tokens[0]
            return token if 3 <= len(token) <= 12 else ""
        if len(tokens) < 2:
            return ""

        acronym = "".join(token[0] for token in tokens)
        if len(acronym) < 3:
            return ""
        return acronym

    def _compact(self, text: str) -> str:
        return "".join(re.findall(r"[a-z0-9]+", text.lower()))

    def _read_ll_info(self, ini_path: Path) -> dict:
        config = configparser.ConfigParser(interpolation=None)
        config.read(ini_path, encoding="utf-8")
        if LL_SECTION not in config:
            return {}

        ll = config[LL_SECTION]
        return {
            "archive_name": ll.get("archive_name", "").strip(),
            "file_name": ll.get("file_name", "").strip(),
            "version": ll.get("version", "").strip(),
        }

    def _confirm_candidates(self, candidates: list[dict]) -> bool:
        lines = [
            f"{candidate['mod_name']} <- {candidate['archive_name']} "
            f"(v{candidate['version'] or '?'}, score {candidate['score']})"
            for candidate in candidates
        ]
        result = QMessageBox.question(
            self._parentWidget(),
            PLUGIN_NAME,
            "Bind these LoversLab metadata matches?\n\n" + "\n".join(lines),
        )
        return result == QMessageBox.StandardButton.Yes


class LoversLabInstallBinder(mobase.IPluginInstallerSimple):
    ARCHIVE_EXTENSIONS = {"7z", "zip", "rar"}

    def __init__(self) -> None:
        super().__init__()
        self._organizer = None
        self._archive_path = None

    def init(self, organizer: mobase.IOrganizer) -> bool:
        self._organizer = organizer
        self._log("init")
        return True

    def name(self) -> str:
        return "LL Integration Install Binder"

    def localizedName(self) -> str:
        return "LL Integration Install Binder"

    def author(self) -> str:
        return "LL Integration"

    def description(self) -> str:
        return "Binds LoversLab sidecar metadata to a mod after installation."

    def version(self) -> mobase.VersionInfo:
        return mobase.VersionInfo("0.3.0")

    def settings(self) -> Sequence[mobase.PluginSetting]:
        return []

    def priority(self) -> int:
        return -1000

    def isManualInstaller(self) -> bool:
        return False

    def isArchiveSupported(self, tree) -> bool:
        return True

    def supportedExtensions(self) -> set[str]:
        return self.ARCHIVE_EXTENSIONS

    def install(self, name, tree, version: str, nexus_id: int):
        self._log(f"install passthrough name={name} version={version} nexus_id={nexus_id}")
        return tree

    def onInstallationStart(self, archive: str, reinstallation: bool, current_mod) -> None:
        self._archive_path = Path(str(archive))
        self._log(f"start archive={self._archive_path} reinstallation={reinstallation}")

    def onInstallationEnd(self, result, new_mod) -> None:
        self._log(f"end result={result} archive={self._archive_path}")
        if not self._is_success(result) or not new_mod or not self._archive_path:
            return

        try:
            sidecar = self._find_sidecar(self._archive_path)
            if not sidecar:
                self._log("no sidecar found")
                return

            target = mod_meta_path(new_mod)
            existing = mod_ll_metadata_path(new_mod)
            if existing and not ll_metadata_same_source(existing, sidecar):
                self._log(f"target exists with different LL source, skip target={target} sidecar={sidecar}")
                return
            is_nexus, reason = mod_has_nexus_identity(new_mod)
            if is_nexus:
                self._log(f"skip Nexus-identified mod: {reason} target={target}")
                return

            target = write_mod_ll_metadata_from_file(new_mod, sidecar)
            self._apply_mod_metadata(new_mod, target)
            if self._organizer:
                self._organizer.modDataChanged(new_mod)
            self._log(f"bound sidecar={sidecar} target={target} replaced_existing={bool(existing)}")
        except Exception as exc:
            self._log(f"bind error: {exc}")

    def _is_success(self, result) -> bool:
        name = getattr(result, "name", "")
        return str(name).upper() == "SUCCESS" or "SUCCESS" in str(result).upper()

    def _find_sidecar(self, archive_path: Path) -> Path | None:
        direct = Path(f"{archive_path}.ll.ini")
        if direct.exists():
            return direct

        try:
            quick_hash = archive_quick_hash(archive_path)
        except OSError as exc:
            self._log(f"quick hash failed: {exc}")
            return None

        for sidecar in self._candidate_sidecars(archive_path):
            if self._sidecar_value(sidecar, "archive_quick_hash") == quick_hash:
                return sidecar

        return None

    def _candidate_sidecars(self, archive_path: Path) -> list[Path]:
        candidates = list(archive_path.parent.glob("*.ll.ini"))
        metadata = self._metadata_downloads_path()
        if metadata.exists():
            candidates.extend(metadata.glob("*.ll.ini"))

        unique = {str(path).lower(): path for path in candidates}
        return list(unique.values())

    def _metadata_downloads_path(self) -> Path:
        config_path = self._native_config_path()
        if config_path.exists():
            try:
                config = json.loads(config_path.read_text(encoding="utf-8-sig"))
                metadata = config.get("metadata_path")
                if metadata:
                    return Path(str(metadata)) / "downloads"
            except Exception:
                pass
        return Path(__file__).resolve().parents[1] / "native-app" / "metadata" / "downloads"

    def _native_config_path(self) -> Path:
        paths = self._configured_paths()
        ll_ini = paths.get("ll_ini_path")
        if ll_ini:
            return Path(str(ll_ini)).parent.parent / "config.json"
        return Path(__file__).resolve().parents[1] / "native-app" / "config.json"

    def _configured_paths(self) -> dict:
        if not PLUGIN_PATHS_FILE.exists():
            return {}

        try:
            data = json.loads(PLUGIN_PATHS_FILE.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError:
            return {}
        return data if isinstance(data, dict) else {}

    def _sidecar_value(self, sidecar: Path, key: str) -> str:
        config = configparser.ConfigParser(interpolation=None)
        config.read(sidecar, encoding="utf-8")
        if LL_SECTION not in config:
            return ""
        return config[LL_SECTION].get(key, "").strip()

    def _apply_mod_metadata(self, mod, ini_path: Path) -> None:
        is_nexus, reason = mod_has_nexus_identity(mod)
        if is_nexus:
            self._log(f"skip metadata for Nexus-identified mod: {reason}")
            return

        config = configparser.ConfigParser(interpolation=None)
        config.read(ini_path, encoding="utf-8")
        ll = config[LL_SECTION]

        page_url = ll.get("page_url", "").strip()
        version = ll.get("version", "").strip()
        write_mod_general_source_metadata(mod, page_url, version)

    def _log(self, message: str) -> None:
        try:
            path = self._log_path()
            path.parent.mkdir(parents=True, exist_ok=True)
            with path.open("a", encoding="utf-8") as file:
                file.write(message + "\n")
        except Exception:
            pass

    def _log_path(self) -> Path:
        if self._organizer:
            return Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "install_binder.log"
        return Path(__file__).resolve().parent / "install_binder.log"


class LoversLabInstallObserver(mobase.IPlugin):
    def __init__(self) -> None:
        super().__init__()
        self._organizer = None
        self._sync_timer = None
        self._last_synced_downloads_path = ""

    def init(self, organizer: mobase.IOrganizer) -> bool:
        self._organizer = organizer
        self._sync_active_instance_config()
        self._start_active_instance_sync_timer()
        ok = organizer.modList().onModInstalled(self._on_mod_installed)
        profile_ok = False
        try:
            profile_ok = organizer.onProfileChanged(self._on_profile_changed)
        except Exception as exc:
            self._log(f"profile change hook failed: {exc}")
        download_ok = False
        try:
            download_ok = organizer.downloadManager().onDownloadRemoved(self._on_download_removed)
        except Exception as exc:
            self._log(f"download remove hook failed: {exc}")

        self._log(f"init onModInstalled={ok} onProfileChanged={profile_ok} onDownloadRemoved={download_ok}")
        return True

    def name(self) -> str:
        return "LL Integration Install Observer"

    def localizedName(self) -> str:
        return "LL Integration Install Observer"

    def author(self) -> str:
        return "LL Integration"

    def description(self) -> str:
        return "Binds LoversLab metadata when MO2 reports a new mod installation."

    def version(self) -> mobase.VersionInfo:
        return mobase.VersionInfo("0.3.0")

    def settings(self) -> Sequence[mobase.PluginSetting]:
        return []

    def _start_active_instance_sync_timer(self) -> None:
        if self._sync_timer is not None:
            return

        self._sync_timer = QTimer()
        self._sync_timer.setInterval(5000)
        self._sync_timer.timeout.connect(self._sync_active_instance_config)
        self._sync_timer.start()

    def _sync_active_instance_config(self) -> None:
        if not self._organizer:
            return

        config_path = self._native_config_path()
        try:
            if config_path.exists():
                data = json.loads(config_path.read_text(encoding="utf-8-sig"))
                config = data if isinstance(data, dict) else {}
            else:
                config = {}

            mo2_root = Path(str(self._organizer.basePath()))
            downloads_path = Path(str(self._organizer.downloadsPath()))
            downloads_text = str(downloads_path)
            if downloads_text == self._last_synced_downloads_path and config_path.exists():
                return

            mo2_exe = mo2_root / "ModOrganizer.exe"
            config.update(
                {
                    "mo2_downloads_path": downloads_text,
                    "active_mo2_instance_path": str(mo2_root),
                    "active_mo2_plugin_path": str(Path(__file__).resolve().parent),
                    "active_mo2_game": self._managed_game_name(),
                    "active_mo2_synced_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime()),
                }
            )
            if mo2_exe.exists():
                config["mo2_path"] = str(mo2_exe)
            config_path.parent.mkdir(parents=True, exist_ok=True)
            config_path.write_text(json.dumps(config, indent=2), encoding="utf-8")
            self._last_synced_downloads_path = downloads_text
            self._log(f"active instance synced downloads={downloads_path}")
        except Exception as exc:
            self._log(f"active instance sync failed: {exc}")

    def _managed_game_name(self) -> str:
        try:
            game = self._organizer.managedGame() if self._organizer else None
            if game and hasattr(game, "gameName"):
                return str(game.gameName())
        except Exception:
            pass
        return ""

    def _on_profile_changed(self, _old_profile, _new_profile) -> None:
        self._sync_active_instance_config()

    def _on_mod_installed(self, mod_or_name) -> None:
        try:
            mod = self._resolve_mod(mod_or_name)
            if mod is None:
                self._log(f"mod installed but could not resolve: {mod_or_name}")
                return

            target = mod_meta_path(mod)
            is_nexus, reason = mod_has_nexus_identity(mod)
            if is_nexus:
                self._log(f"skip Nexus-identified mod: {reason} target={target}")
                return

            sidecar = self._latest_installed_sidecar(mod)
            if not sidecar:
                self._log("no recent installed LL sidecar found")
                return
            existing = mod_ll_metadata_path(mod)
            if existing and not ll_metadata_same_source(existing, sidecar):
                self._log(f"skip existing different LL source target={target} sidecar={sidecar}")
                return

            target = write_mod_ll_metadata_from_file(mod, sidecar)
            self._apply_mod_metadata(mod, target)
            if self._organizer:
                self._organizer.modDataChanged(mod)
            self._log(f"auto-bound sidecar={sidecar} target={target} replaced_existing={bool(existing)}")
        except Exception as exc:
            self._log(f"auto-bind error: {exc}")

    def _on_download_removed(self, download_id: int) -> None:
        try:
            if not self._organizer:
                return

            download_path = Path(str(self._organizer.downloadManager().downloadPath(download_id)))
            self._log(f"download removed id={download_id} path={download_path}")
            if not str(download_path):
                return

            if download_path.exists():
                self._log(f"skip metadata cleanup because archive still exists: {download_path}")
                return

            removed = self._remove_metadata_for_archive(download_path)
            self._log(f"removed LL metadata count={removed} archive={download_path}")
        except Exception as exc:
            self._log(f"download removed cleanup error: {exc}")

    def _remove_metadata_for_archive(self, archive_path: Path) -> int:
        removed = 0
        archive_name = archive_path.name

        direct_paths = [
            Path(f"{archive_path}.ll.ini"),
            Path(f"{archive_path}.ll.json"),
        ]

        for path in direct_paths:
            if path.exists():
                path.unlink()
                removed += 1

        metadata_path = self._metadata_downloads_path()
        for sidecar in metadata_path.glob("*.ll.ini") if metadata_path.exists() else []:
            if self._sidecar_matches_archive(sidecar, archive_path, archive_name):
                json_sidecar = sidecar.with_suffix(".json")
                sidecar.unlink(missing_ok=True)
                removed += 1
                if json_sidecar.exists():
                    json_sidecar.unlink()
                    removed += 1

        return removed

    def _sidecar_matches_archive(self, sidecar: Path, archive_path: Path, archive_name: str) -> bool:
        config = configparser.ConfigParser(interpolation=None)
        config.read(sidecar, encoding="utf-8")
        if LL_SECTION not in config:
            return False

        ll = config[LL_SECTION]
        return (
            ll.get("archive_path", "").strip().lower() == str(archive_path).lower()
            or ll.get("archive_name", "").strip().lower() == archive_name.lower()
        )

    def _resolve_mod(self, mod_or_name):
        if hasattr(mod_or_name, "absolutePath"):
            return mod_or_name

        if self._organizer:
            try:
                return self._organizer.modList().getMod(str(mod_or_name))
            except Exception:
                return None
        return None

    def _latest_installed_sidecar(self, mod) -> Path | None:
        install_time = self._mod_install_time(mod)
        candidates = []
        for sidecar in self._candidate_sidecars():
            archive_path = self._archive_path_from_sidecar(sidecar)
            if not archive_path:
                continue

            meta_path = Path(f"{archive_path}.meta")
            if not meta_path.exists() or not self._meta_says_installed(meta_path):
                continue

            delta = abs(meta_path.stat().st_mtime - install_time)
            if delta > AUTO_BIND_WINDOW_SECONDS:
                continue

            candidates.append((delta, meta_path.stat().st_mtime, sidecar, archive_path))

        if not candidates:
            return None

        candidates.sort(key=lambda item: (item[0], -item[1]))
        self._log(
            "candidate sidecars: "
            + "; ".join(f"{sidecar.name} archive={archive.name} delta={delta:.1f}s" for delta, _, sidecar, archive in candidates[:5])
        )
        return candidates[0][2]

    def _mod_install_time(self, mod) -> float:
        mod_path = Path(str(mod.absolutePath()))
        meta_ini = mod_path / "meta.ini"
        if meta_ini.exists():
            return meta_ini.stat().st_mtime
        return time.time()

    def _candidate_sidecars(self) -> list[Path]:
        candidates = []
        for path in self._metadata_downloads_path().glob("*.ll.ini"):
            candidates.append(path)

        downloads_path = self._mo2_downloads_path()
        if downloads_path and downloads_path.exists():
            candidates.extend(downloads_path.glob("*.ll.ini"))

        unique = {str(path).lower(): path for path in candidates}
        return list(unique.values())

    def _archive_path_from_sidecar(self, sidecar: Path) -> Path | None:
        config = configparser.ConfigParser(interpolation=None)
        config.read(sidecar, encoding="utf-8")
        if LL_SECTION not in config:
            return None

        archive_path = config[LL_SECTION].get("archive_path", "").strip()
        if archive_path:
            path = Path(archive_path)
            if path.exists():
                return path

        name = config[LL_SECTION].get("archive_name", "").strip()
        downloads_path = self._mo2_downloads_path()
        if name and downloads_path:
            path = downloads_path / name
            if path.exists():
                return path

        return None

    def _meta_says_installed(self, meta_path: Path) -> bool:
        config = configparser.ConfigParser(interpolation=None)
        config.read(meta_path, encoding="utf-8")
        if "General" not in config:
            return False

        general = config["General"]
        return (
            general.get("installed", "").lower() == "true"
            and general.get("uninstalled", "").lower() == "false"
        )

    def _metadata_downloads_path(self) -> Path:
        config = self._native_config()
        metadata = config.get("metadata_path")
        if metadata:
            return Path(str(metadata)) / "downloads"
        return Path(__file__).resolve().parents[1] / "native-app" / "metadata" / "downloads"

    def _mo2_downloads_path(self) -> Path | None:
        config = self._native_config()
        downloads = config.get("mo2_downloads_path")
        return Path(str(downloads)) if downloads else None

    def _native_config(self) -> dict:
        config_path = self._native_config_path()
        if not config_path.exists():
            return {}

        try:
            data = json.loads(config_path.read_text(encoding="utf-8-sig"))
        except Exception:
            return {}
        return data if isinstance(data, dict) else {}

    def _native_config_path(self) -> Path:
        paths = self._configured_paths()
        ll_ini = paths.get("ll_ini_path")
        if ll_ini:
            return Path(str(ll_ini)).parent.parent / "config.json"
        return Path(__file__).resolve().parents[1] / "native-app" / "config.json"

    def _configured_paths(self) -> dict:
        if not PLUGIN_PATHS_FILE.exists():
            return {}

        try:
            data = json.loads(PLUGIN_PATHS_FILE.read_text(encoding="utf-8-sig"))
        except json.JSONDecodeError:
            return {}
        return data if isinstance(data, dict) else {}

    def _apply_mod_metadata(self, mod, ini_path: Path) -> None:
        is_nexus, reason = mod_has_nexus_identity(mod)
        if is_nexus:
            self._log(f"skip metadata for Nexus-identified mod: {reason}")
            return

        config = configparser.ConfigParser(interpolation=None)
        config.read(ini_path, encoding="utf-8")
        ll = config[LL_SECTION]

        page_url = ll.get("page_url", "").strip()
        version = ll.get("version", "").strip()
        write_mod_general_source_metadata(mod, page_url, version)

    def _log(self, message: str) -> None:
        for path in self._log_paths():
            try:
                path.parent.mkdir(parents=True, exist_ok=True)
                with path.open("a", encoding="utf-8") as file:
                    file.write(message + "\n")
            except Exception:
                pass

    def _log_paths(self) -> list[Path]:
        paths = [Path(__file__).resolve().parent / "install_observer.log"]
        if self._organizer:
            paths.append(Path(str(self._organizer.pluginDataPath())) / "ll_integration" / "install_observer.log")
        return paths
