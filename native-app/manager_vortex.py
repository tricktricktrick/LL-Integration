import configparser
import fnmatch
import gzip
import json
import re
import shutil
import sys
import time
from tkinter import dialog
import webbrowser
import zlib
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import parse_qs, quote_plus, unquote, urlencode, urljoin, urlparse, urlunparse
from urllib.request import Request, urlopen

from PyQt6.QtCore import QObject, QThread, QTimer, Qt, pyqtSignal
from PyQt6.QtGui import QAction, QColor, QIcon
from PyQt6.QtWidgets import (
    QApplication,
    QAbstractItemView,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QDoubleSpinBox,
    QGridLayout,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QTextEdit,
    QVBoxLayout,
    QMenu,
    QMessageBox,
    QProgressBar,
    QSpinBox,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QProgressBar,
    QPushButton,
    QSplitter
)


if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

APP_ID_BASE = "LLIntegration.Tools"


def resource_base_dir() -> Path:
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(sys._MEIPASS)
    return BASE_DIR


def set_windows_app_id(tool_id: str = "manager") -> None:
    if sys.platform != "win32":
        return

    try:
        import ctypes
        app_id = f"{APP_ID_BASE}.{tool_id}"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(app_id)
    except Exception:
        pass


def app_icon_path(mode: str = "manager") -> Path:
    candidates = []

    for root in (BASE_DIR, resource_base_dir()):
        icons_dir = root / "icons"
        candidates.extend([
            icons_dir / f"ll_integration_{mode}.ico",
            icons_dir / "ll_integration.ico",
            icons_dir / "ll_integration_128.png",
            icons_dir / "ll_integration.png",
        ])

    for path in candidates:
        if path.exists():
            return path

    return Path()


def apply_app_icon(app: QApplication, mode: str = "manager") -> QIcon:
    path = app_icon_path(mode)
    icon = QIcon(str(path)) if path else QIcon()

    if not icon.isNull():
        app.setWindowIcon(icon)

    return icon

CONFIG_FILE = BASE_DIR / "config.json"
VORTEX_STATE_FILE = BASE_DIR / "vortex_state.json"
VORTEX_COMMANDS_FILE = BASE_DIR / "vortex_commands.json"
VORTEX_FETCH_PACING_FILE = BASE_DIR / "vortex_fetch_pacing.json"
ARCHIVE_SUFFIXES = {".7z", ".zip", ".rar"}
COOKIE_NAMES = {"ips4_IPSSessionFront", "ips4_member_id", "ips4_login_key"}
VERSION_RE = r"(?<!\d)v?(\d+(?:[.-]\d+){1,3})(?:\b|(?=\D))"
UPDATE_REQUEST_DELAY_SECONDS = 0.3
UPDATE_BATCH_SIZE = 25
UPDATE_BATCH_PAUSE_SECONDS = 0.0
UPDATE_REQUEST_TIMEOUT_SECONDS = 5.0
VOICE_DOWNLOAD_TIMEOUT_SECONDS = 120.0

DARK_STYLE = """
QDialog {
    background: #181b1f;
    color: #f4f7f8;
}
QLabel {
    color: #f4f7f8;
}
QLineEdit {
    background: #242a2f;
    color: #f4f7f8;
    border: 1px solid #3a444b;
    padding: 5px;
}
QTableWidget {
    background: #202428;
    color: #f4f7f8;
    gridline-color: #111417;
    alternate-background-color: #282d32;
    selection-background-color: #353d45;
    selection-color: #ffffff;
}
QHeaderView::section {
    background: #3b414b;
    color: #ffffff;
    border: 1px solid #242a2f;
    padding: 4px;
}
QPushButton {
    background: #3f3f3f;
    color: #ffffff;
    border: 1px solid #686868;
    border-radius: 2px;
    padding: 4px 10px;
    min-height: 20px;
}
QPushButton:hover {
    background: #505050;
}
QPushButton:pressed {
    background: #0d6b3c;
}
QPushButton:disabled {
    background: #333333;
    color: #8b8b8b;
    border-color: #4a4a4a;
}
QComboBox, QDoubleSpinBox, QSpinBox {
    background: #3f3f3f;
    color: #ffffff;
    border: 1px solid #686868;
    padding: 3px;
}
QMenu {
    background: #242a2f;
    color: #f4f7f8;
    border: 1px solid #3a444b;
}
QMenu::item:selected {
    background: #0d6b3c;
}
"""

ACTION_COLUMNS = {6, 7, 8, 9, 10}
VOICE_COL_STATUS = 0
VOICE_COL_BASE_MOD = 1
VOICE_COL_VOICEPACK = 2
VOICE_COL_DBVO = 3
VOICE_COL_IVDT = 4
VOICE_COL_CATEGORY = 5
VOICE_COL_INSTALLED_VOICE = 6
VOICE_COL_SCORE = 7
VOICE_COL_ONLINE = 8
VOICE_COL_ONLINE_SCORE = 9
VOICE_COL_SOURCE = 10
VOICE_COL_PAGE = 11
UPDATE_MODE_MANUAL = "manual"
UPDATE_MODE_DOWNLOAD_ONLY = "download_only"
UPDATE_MODE_ASSISTED = "assisted"
UPDATE_MODE_AUTOMATIC = "automatic"
UPDATE_MODE_SKIP = "skip"
UPDATE_MODE_OPTIONS = [
    (UPDATE_MODE_MANUAL, "Manual install"),
    (UPDATE_MODE_DOWNLOAD_ONLY, "Download only"),
    (UPDATE_MODE_ASSISTED, "Assisted install"),
    (UPDATE_MODE_AUTOMATIC, "Automatic install"),
    (UPDATE_MODE_SKIP, "Skip updates"),
]
UPDATE_MODE_LABELS = {value: label for value, label in UPDATE_MODE_OPTIONS}
VERSION_MARKERS = ("{version}", "{v}", "<version>", "<v>")

DEFAULT_VOICE_MATCH_THRESHOLD = 55

NEXUS_SOURCE_RE = re.compile(
    r"^https?://(?:www\.)?nexusmods\.com/(?P<game>[^/\s?#]+)/mods/(?P<mod_id>\d+)",
    re.IGNORECASE,
)

VOICE_KEYWORDS = {
    "voice",
    "voices",
    "voicepack",
    "voicepacks",
    "dbvo",
    "IVDT",
    "ivdt",
    "dialogue",
    "dialogues",
    "addon",
}

VOICE_NOISE_WORDS_RE = re.compile(
    r"""
    \b(
        se|ae|skyrim|special|edition|anniversary|edition|
        fomod|main|file|files|optional|miscellaneous|
        voice|voices|voicepack|voicepacks|voice|pack|
        dbvo|IVDT|ivdt|dialogue|dialogues|addon|patch|
        fixed|fix|update|hotfix|version
    )\b
    """,
    re.IGNORECASE | re.VERBOSE,
)
VOICE_CATEGORIES = [
    ("npc", "Voicepack"),
    ("player", "DBVO"),
    ("scene", "IVDT"),
]

VOICE_CATEGORY_LABELS = {
    "npc": "Voicepack",
    "player": "DBVO",
    "scene": "IVDT",
}
NEXUS_SOURCE_RE = re.compile(
    r"^https?://(?:www\.)?nexusmods\.com/(?P<game>[^/\s?#]+)/mods/(?P<mod_id>\d+)",
    re.IGNORECASE,
)

BAD_NEXUS_TITLE_PREFIX_RE = re.compile(
    r"^\s*(?:download\s+your\s+files?|download\s+files?|files?)\s*[-:–—]\s*",
    re.IGNORECASE,
)

DEFAULT_VOICE_MATCH_THRESHOLD = 55

def clean_source_download_name(value: str) -> str:
    name = str(value or "").strip()
    if not name:
        return ""

    # Nexus peut parfois retourner un titre UI genre:
    # "Download your files - Ciri DBVO Voice..."
    previous = None
    while previous != name:
        previous = name
        name = BAD_NEXUS_TITLE_PREFIX_RE.sub("", name).strip()

    name = re.sub(r"\s+", " ", name).strip()
    return name


def display_name_from_archive_name(value: str) -> str:
    name = safe_archive_name(value)
    if not name:
        return ""

    name = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", name, flags=re.IGNORECASE)
    name = clean_source_download_name(name)
    return name


def safe_archive_name(value: str) -> str:
    name = str(value or "").strip()
    if not name:
        return ""

    # Si Nexus/API retourne juste un nom sans extension, on garde le nom,
    # mais on évite les caractères invalides Windows.
    name = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    name = name.rstrip(". ")

    return name

def normalized_voice_name(value: str) -> str:
    text = str(value or "").lower()
    text = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bv?\d+(?:[._-]\d+){1,4}\b", " ", text, flags=re.IGNORECASE)
    text = VOICE_NOISE_WORDS_RE.sub(" ", text)
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return " ".join(token for token in text.split() if len(token) > 1)

def voice_search_query(value: str) -> str:
    text = normalized_voice_name(value)

    if not text:
        text = str(value or "").strip().lower()
        text = re.sub(r"[^a-z0-9]+", " ", text)
        text = " ".join(text.split())

    return text

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

def is_loverslab_source_url(value: str) -> bool:
    url = str(value or "").strip().lower()
    return "loverslab.com/files/file/" in url

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

def fetch_loverslab_source_downloads(
    source_url: str,
    cookies_path: Path,
    timeout: float = 30.0,
) -> list[dict]:
    downloads = fetch_ll_downloads(source_url, cookies_path, timeout)

    results: list[dict] = []
    normalized_source = with_query_value(source_url, "do", "download")

    for item in downloads:
        name = str(getattr(item, "name", "") or "").strip()
        url = str(getattr(item, "url", "") or "").strip()

        if not name or not url:
            continue

        results.append({
            "source_url": normalized_source,
            "source_title": "LoversLab source",
            "download_name": name,
            "file_name": name,
            "archive_file_name": name,
            "download_url": url,
            "voice_category": voice_category_guess(name),
            "size": str(getattr(item, "size", "") or ""),
            "date_iso": str(getattr(item, "date_iso", "") or ""),
            "version": str(getattr(item, "version", "") or extract_version(name) or ""),
            "source_type": "loverslab",
            "nexus_file_id": "",
            "nexus_category": "",
        })

    return results

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
        file_name = safe_archive_name(str(item.get("file_name") or "").strip())
        raw_name = str(item.get("name") or "").strip()

        archive_name = file_name or safe_archive_name(raw_name)
        display_name = clean_source_download_name(raw_name) or display_name_from_archive_name(archive_name)

        if not file_id or not archive_name:
            continue

        size_kb = item.get("size_kb")
        size = f"{size_kb} KB" if size_kb not in (None, "") else ""

        downloads.append({
            "source_url": source_url,
            "source_title": "Nexus Mods API source",
            "download_name": display_name,
            "file_name": archive_name,
            "archive_file_name": archive_name,
            "download_url": f"https://www.nexusmods.com/{game}/mods/{mod_id}?tab=files&file_id={file_id}&nmm=1",
            "voice_category": voice_category_guess(display_name or archive_name),
            "size": size,
            "date_iso": str(item.get("uploaded_time") or ""),
            "version": str(item.get("version") or item.get("mod_version") or ""),
            "source_type": "nexus",
            "nexus_file_id": file_id,
            "nexus_category": str(item.get("category_name") or ""),
        })

    return downloads

def normalize_voice_source_url(value: str) -> str:
    url = str(value or "").strip()
    if not url:
        return ""

    # LoversLab file page:
    # https://www.loverslab.com/files/file/1234-name/
    # becomes:
    # https://www.loverslab.com/files/file/1234-name/?do=download
    if "loverslab.com/files/file/" in url.lower():
        return with_query_value(url, "do", "download")

    # Nexus page:
    # https://www.nexusmods.com/skyrimspecialedition/mods/12345
    # becomes:
    # https://www.nexusmods.com/skyrimspecialedition/mods/12345?tab=files
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

@dataclass(frozen=True)
class LLDownload:
    name: str
    url: str
    size: str = ""
    date_iso: str = ""
    version: str = ""


class LLDownloadParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.downloads: list[LLDownload] = []
        self._in_title = False
        self._in_meta = False
        self._current_name = ""
        self._current_size_parts: list[str] = []
        self._current_date_iso = ""

    def handle_starttag(self, tag: str, attrs: list[tuple]) -> None:
        attrs_dict = dict(attrs)
        classes = attrs_dict.get("class", "")
        if tag == "span" and "ipsType_break" in classes:
            self._in_title = True
            self._current_name = ""
            self._current_size_parts = []
            self._current_date_iso = ""
            return
        if tag == "p" and "ipsDataItem_meta" in classes:
            self._in_meta = True
            return
        if tag == "time" and self._in_meta:
            self._current_date_iso = attrs_dict.get("datetime", "")
            return
        if tag == "a" and attrs_dict.get("data-action") == "download":
            href = attrs_dict.get("href", "")
            if href and self._current_name:
                size = " ".join(" ".join(self._current_size_parts).split())
                self.downloads.append(LLDownload(
                    name=self._current_name.strip(),
                    url=unquote(href).replace("&amp;", "&"),
                    size=(size.split(" / ")[0].strip() if size else ""),
                    date_iso=self._current_date_iso,
                    version=extract_version(self._current_name) or "",
                ))

    def handle_endtag(self, tag: str) -> None:
        if tag == "span" and self._in_title:
            self._in_title = False
        elif tag == "p" and self._in_meta:
            self._in_meta = False

    def handle_data(self, data: str) -> None:
        if self._in_title:
            self._current_name += data
        elif self._in_meta:
            self._current_size_parts.append(data)

def configured_cookies_path(config: dict) -> Path:
    return Path(str(
        config.get("cookies_path")
        or BASE_DIR / "cookies_storage" / "cookies_ll.json"
    ))

def read_json(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def save_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2), encoding="utf-8")


def with_query_value(url: str, key: str, value: str) -> str:
    parsed = urlparse(url)
    query = parse_qs(parsed.query)
    query[key] = [value]
    return urlunparse(parsed._replace(query=urlencode(query, doseq=True)))


def open_path(path: Path) -> None:
    target = path if path.is_dir() else path.parent
    if not target.exists():
        QMessageBox.information(None, "LL Integration", f"Path not found:\n{path}")
        return
    try:
        webbrowser.open(target.as_uri())
    except ValueError:
        webbrowser.open(str(target))


def load_ini(path: Path) -> dict:
    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(path, encoding="utf-8-sig")
    except Exception:
        return {}
    if not parser.has_section("LoversLab"):
        return {}
    try:
        return {key: value for key, value in parser.items("LoversLab", raw=True)}
    except Exception:
        return {}


def write_ini(path: Path, data: dict) -> None:
    parser = configparser.ConfigParser(interpolation=None)
    parser["LoversLab"] = {key: str(value or "") for key, value in data.items()}
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as handle:
        parser.write(handle)


def extract_version(file_name: str) -> str:
    stem = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", file_name, flags=re.IGNORECASE)
    match = re.search(VERSION_RE, stem, re.IGNORECASE)
    return match.group(1) if match else ""


def version_key(version: str) -> tuple:
    return tuple(int(part) for part in re.findall(r"\d+", version or ""))


def compare_versions(left: str, right: str) -> int:
    left_parts = version_key(left)
    right_parts = version_key(right)
    max_len = max(len(left_parts), len(right_parts), 1)
    left_parts += (0,) * (max_len - len(left_parts))
    right_parts += (0,) * (max_len - len(right_parts))
    return (left_parts > right_parts) - (left_parts < right_parts)


def normalized_update_mode(value: str | None, fixed: bool = False) -> str:
    mode = str(value or "").strip().lower()
    if mode == "download":
        mode = UPDATE_MODE_DOWNLOAD_ONLY
    if mode in UPDATE_MODE_LABELS:
        return mode
    return UPDATE_MODE_SKIP if fixed else UPDATE_MODE_MANUAL


def update_mode_label(mode: str) -> str:
    return UPDATE_MODE_LABELS.get(normalized_update_mode(mode), UPDATE_MODE_LABELS[UPDATE_MODE_MANUAL])


def configure_update_mode_combo(combo: QComboBox, selected: str) -> None:
    selected_mode = normalized_update_mode(selected)
    selected_index = 0
    for index, (value, label) in enumerate(UPDATE_MODE_OPTIONS):
        combo.addItem(label, value)
        if value == selected_mode:
            selected_index = index
    combo.setCurrentIndex(selected_index)


def filename_prefix(file_name: str) -> str:
    stem = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", file_name, flags=re.IGNORECASE)
    marker_positions = [stem.find(marker) for marker in VERSION_MARKERS if marker in stem]
    if marker_positions:
        prefix = stem[:min(marker_positions)]
    else:
        prefix = re.split(VERSION_RE, stem, maxsplit=1, flags=re.IGNORECASE)[0]
    return re.sub(r"[^a-z0-9]+", " ", prefix.lower()).strip()


def compact_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value or "").lower())


def version_marker_match_pattern(pattern: str) -> str:
    match_pattern = str(pattern or "")
    for marker in VERSION_MARKERS:
        match_pattern = match_pattern.replace(marker, "*")
    return match_pattern


def link_pattern_value(data: dict, row: dict) -> str:
    return (
        str(data.get("file_pattern") or "").strip()
        or str(data.get("file_name") or "").strip()
        or str(data.get("archive_name") or "").strip()
        or str(row.get("archive") or "").strip()
    )


def strip_archive_extension(value: str) -> str:
    return re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", str(value or ""), flags=re.IGNORECASE)


def has_archive_extension(value: str) -> bool:
    return bool(re.search(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", str(value or ""), flags=re.IGNORECASE))


def version_marker_version(file_name: str, pattern: str) -> str:
    pattern = pattern.strip()
    if not any(item in pattern for item in VERSION_MARKERS):
        return ""

    def build_regex(source_pattern: str) -> tuple[str, bool]:
        regex = ""
        index = 0
        group_added = False
        while index < len(source_pattern):
            marker = next((item for item in VERSION_MARKERS if source_pattern.startswith(item, index)), "")
            if marker:
                regex += "(.+?)"
                index += len(marker)
                group_added = True
                continue
            char = source_pattern[index]
            regex += ".*?" if char == "*" else re.escape(char)
            index += 1
        return regex, group_added

    regex, group_added = build_regex(pattern)
    if not group_added:
        return ""

    match = re.fullmatch(regex, file_name, flags=re.IGNORECASE) if has_archive_extension(pattern) else None
    if not match:
        stem_regex, _group_added = build_regex(strip_archive_extension(pattern))
        match = re.fullmatch(stem_regex, strip_archive_extension(file_name), flags=re.IGNORECASE)
    if not match:
        return ""

    value = ".".join(part for group in match.groups() for part in re.findall(r"\d+", group))
    return value


def wildcard_version(file_name: str, pattern: str) -> str:
    marked = version_marker_version(file_name, pattern)
    if marked:
        return marked

    pattern = pattern.strip()
    if "*" not in pattern:
        return ""

    regex = ""
    for char in strip_archive_extension(pattern):
        regex += "(.*?)" if char == "*" else re.escape(char)
    match = re.fullmatch(regex, strip_archive_extension(file_name), flags=re.IGNORECASE)
    if not match:
        return ""

    parts = []
    for value in match.groups():
        parts.extend(re.findall(r"\d+", value))
    return ".".join(parts)


def candidate_version(download: LLDownload, known_file: str) -> str:
    version = download.version or wildcard_version(download.name, known_file)
    if version:
        return version
    if any(marker in known_file for marker in VERSION_MARKERS):
        marker_index = min(
            (known_file.find(marker) for marker in VERSION_MARKERS if marker in known_file),
            default=-1,
        )
        marker_prefix = known_file[:marker_index] if marker_index >= 0 else ""
        if compact_name(marker_prefix) and compact_name(marker_prefix) == compact_name(filename_prefix(download.name)):
            return extract_version(download.name) or ""
    return ""


def score_download(download: LLDownload, known_file: str) -> int:
    pattern = known_file.strip()
    match_pattern = version_marker_match_pattern(pattern)
    known_prefix = filename_prefix(known_file)
    download_prefix = filename_prefix(download.name)
    score = 0
    if any(char in match_pattern for char in "*?[]"):
        lower_name = download.name.lower()
        lower_pattern = match_pattern.lower()
        if fnmatch.fnmatch(lower_name, lower_pattern):
            score += 130
        elif fnmatch.fnmatch(lower_name, f"*{lower_pattern}*"):
            score += 95
    if known_prefix and download_prefix == known_prefix:
        score += 100
    elif known_prefix and compact_name(known_prefix) == compact_name(download_prefix):
        score += 100
    elif known_prefix and (known_prefix in download_prefix or download_prefix in known_prefix):
        score += 60
    if not any(char in pattern for char in "*?[]") and Path(download.name).suffix.lower() == Path(known_file).suffix.lower():
        score += 20
    if candidate_version(download, known_file):
        score += 10
    return score

def google_voice_search_terms(value: str) -> str:
    text = str(value or "").strip()

    # Retire extension/archive/version/bruit
    text = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", " ", text, flags=re.IGNORECASE)
    text = re.sub(r"\bv?\d+(?:[._-]\d+){1,4}\b", " ", text, flags=re.IGNORECASE)

    # Sépare les CamelCase: BarefootRealismNG -> Barefoot Realism NG
    text = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", text)

    # Retire les mots qui rendent Google trop étroit
    text = re.sub(
        r"\b(voicepack|voicepacks|voice\s*pack|voices?|dbvo|ivdt|dvit|addon|patch|se|ae|le)\b",
        " ",
        text,
        flags=re.IGNORECASE,
    )

    text = re.sub(r"[^a-zA-Z0-9]+", " ", text)
    text = " ".join(part for part in text.split() if len(part) > 1)

    return text

def choose_latest(downloads: list[LLDownload], known_file: str) -> LLDownload | None:
    scored = [
        (score_download(download, known_file), download, candidate_version(download, known_file))
        for download in downloads
    ]
    scored = [item for item in scored if item[0] >= 80 and item[2]]
    if not scored:
        return None

    _score, download, version = max(scored, key=lambda item: (item[0], version_key(item[2] or "0")))
    if download.version == version:
        return download
    return LLDownload(download.name, download.url, download.size, download.date_iso, version)


def fixed_update(data: dict) -> bool:
    fixed = any(str(data.get(key) or "").strip().lower() == "true" for key in ("fixed_version", "manual_update", "skip_update_check"))
    return normalized_update_mode(data.get("update_mode"), fixed=fixed) == UPDATE_MODE_SKIP or fixed


def load_fetch_pacing() -> dict:
    pacing = {
        "request_delay": UPDATE_REQUEST_DELAY_SECONDS,
        "batch_size": UPDATE_BATCH_SIZE,
        "batch_pause": UPDATE_BATCH_PAUSE_SECONDS,
        "request_timeout": UPDATE_REQUEST_TIMEOUT_SECONDS,
    }
    if not VORTEX_FETCH_PACING_FILE.exists():
        return pacing

    try:
        data = json.loads(VORTEX_FETCH_PACING_FILE.read_text(encoding="utf-8-sig"))
    except Exception:
        return pacing
    if not isinstance(data, dict):
        return pacing

    try:
        pacing["request_delay"] = max(0.0, min(30.0, float(data.get("request_delay", pacing["request_delay"]))))
        pacing["batch_size"] = max(1, min(999, int(data.get("batch_size", pacing["batch_size"]))))
        pacing["batch_pause"] = max(0.0, min(120.0, float(data.get("batch_pause", pacing["batch_pause"]))))
        pacing["request_timeout"] = max(1.0, min(120.0, float(data.get("request_timeout", pacing["request_timeout"]))))
    except Exception:
        return pacing
    return pacing


def load_vortex_state(config: dict) -> dict:
    state_path = Path(str(config.get("vortex_state_path") or VORTEX_STATE_FILE))
    if not state_path.exists():
        return {}
    try:
        data = json.loads(state_path.read_text(encoding="utf-8-sig"))
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def append_vortex_command(command: dict) -> str:
    queue = load_json_object(VORTEX_COMMANDS_FILE, {"commands": []})
    commands = queue.get("commands") if isinstance(queue.get("commands"), list) else []
    command_id = f"ll-{int(__import__('time').time() * 1000)}"
    command = dict(command)
    command["id"] = command_id
    command["createdAt"] = __import__("time").strftime("%Y-%m-%dT%H:%M:%SZ", __import__("time").gmtime())
    commands.append(command)
    save_json(VORTEX_COMMANDS_FILE, {"commands": commands[-100:]})
    return command_id


def load_json_object(path: Path, default: dict | None = None) -> dict:
    if default is None:
        default = {}
    if not path.exists():
        return dict(default)
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception:
        return dict(default)
    return data if isinstance(data, dict) else dict(default)


def save_fetch_pacing(request_delay: float, batch_size: int, batch_pause: float, request_timeout: float) -> None:
    save_json(
        VORTEX_FETCH_PACING_FILE,
        {
            "request_delay": round(float(request_delay), 2),
            "batch_size": int(batch_size),
            "batch_pause": round(float(batch_pause), 2),
            "request_timeout": round(float(request_timeout), 2),
        },
    )


def load_cookies(path: Path) -> dict:
    if not path.exists():
        raise RuntimeError(f"LoversLab cookies are not exported yet:\n{path}")
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    cookies = data.get("cookies", [])
    return {
        cookie["name"]: cookie["value"]
        for cookie in cookies
        if cookie.get("name") and cookie.get("value")
    }


def fetch_ll_downloads(url: str, cookies_path: Path, timeout: float) -> list[LLDownload]:
    cookies = load_cookies(cookies_path)
    download_url = with_query_value(url, "do", "download")
    headers = {
        "Cookie": "; ".join(f"{name}={value}" for name, value in cookies.items()),
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "identity",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "DNT": "1",
        "Referer": url,
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    request = Request(download_url, headers=headers)
    with urlopen(request, timeout=timeout) as response:
        data = response.read()
        encoding = (response.headers.get("Content-Encoding") or "").lower()
        if encoding == "gzip":
            data = gzip.decompress(data)
        elif encoding == "deflate":
            data = zlib.decompress(data)
        html = data.decode(response.headers.get_content_charset() or "utf-8", errors="replace")
    parser = LLDownloadParser()
    try:
        parser.feed(html)
    except AssertionError:
        parser = LLDownloadParser()
        parser.feed(html.replace("<![", "&lt;!["))
    return parser.downloads


def ll_request_headers(cookies_path: Path, referer: str = "") -> dict:
    cookies = load_cookies(cookies_path)
    headers = {
        "Cookie": "; ".join(f"{name}={value}" for name, value in cookies.items()),
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Encoding": "identity",
        "Accept-Language": "en-US,en;q=0.9",
        "Connection": "keep-alive",
        "DNT": "1",
        "Sec-Fetch-Dest": "document",
        "Sec-Fetch-Mode": "navigate",
        "Sec-Fetch-Site": "same-origin" if referer else "none",
        "Sec-Fetch-User": "?1",
        "Upgrade-Insecure-Requests": "1",
    }
    if referer:
        headers["Referer"] = referer
    return headers
def archive_file_name_from_candidate(candidate: dict) -> str:
    for key in ("archive_file_name", "file_name", "download_file_name", "download_name", "name"):
        value = safe_archive_name(str(candidate.get(key) or "").strip())
        if value:
            name = value
            break
    else:
        name = "voice-pack.7z"

    suffix = Path(name).suffix.lower()
    if suffix not in ARCHIVE_SUFFIXES:
        # Fallback. Nexus API usually gives file_name with extension.
        # If not, .7z is the safest common fallback for Skyrim mod archives.
        name = f"{name}.7z"

    return name


def download_raw_url(url: str, target: Path, timeout: float, headers: dict | None = None) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    request = Request(url, headers=headers or {
        "User-Agent": "LL Integration",
        "Accept": "*/*",
    })
    tmp_target = target.with_name(f"{target.name}.lltmp")

    with urlopen(request, timeout=timeout) as response, tmp_target.open("wb") as output:
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            output.write(chunk)

    tmp_target.replace(target)
    return target

def debug_icon_status(mode: str, icon: QIcon) -> None:
    try:
        log_path = BASE_DIR / "icon_debug.txt"
        log_path.write_text(
            f"mode={mode}\n"
            f"BASE_DIR={BASE_DIR}\n"
            f"resource_base_dir={resource_base_dir()}\n"
            f"icon_path={app_icon_path(mode)}\n"
            f"icon_null={icon.isNull()}\n",
            encoding="utf-8",
        )
    except Exception:
        pass

def fetch_nexus_download_links_api(
    game: str,
    mod_id: str,
    file_id: str,
    api_key: str,
    timeout: float = 30.0,
) -> list[str]:
    url = f"https://api.nexusmods.com/v1/games/{game}/mods/{mod_id}/files/{file_id}/download_link.json"
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

    links: list[str] = []

    if isinstance(payload, list):
        for item in payload:
            if isinstance(item, dict):
                uri = str(item.get("URI") or item.get("uri") or "").strip()
                if uri:
                    links.append(uri)
            elif isinstance(item, str) and item.strip():
                links.append(item.strip())

    elif isinstance(payload, dict):
        for key in ("URI", "uri", "download_link", "url"):
            uri = str(payload.get(key) or "").strip()
            if uri:
                links.append(uri)

        data = payload.get("data")
        if isinstance(data, list):
            for item in data:
                if isinstance(item, dict):
                    uri = str(item.get("URI") or item.get("uri") or "").strip()
                    if uri:
                        links.append(uri)

    return links


def download_nexus_candidate(
    candidate: dict,
    target: Path,
    api_key: str,
    timeout: float = VOICE_DOWNLOAD_TIMEOUT_SECONDS,
) -> Path:
    file_id = str(candidate.get("nexus_file_id") or "").strip()
    source_url = str(candidate.get("source_url") or "").strip()
    parts = nexus_source_parts(source_url)

    if not parts or not file_id:
        raise RuntimeError("Nexus candidate is missing game/mod/file information.")

    game, mod_id = parts
    links = fetch_nexus_download_links_api(game, mod_id, file_id, api_key, timeout=30.0)

    if not links:
        raise RuntimeError(
            "Nexus API returned no download link for this file.\n\n"
            "This can happen if the file requires manual download, permission, adult confirmation, "
            "or if Nexus blocks direct API downloads for this account/file."
        )

    return download_raw_url(
        links[0],
        target,
        timeout=timeout,
        headers={
            "User-Agent": "LL Integration",
            "Accept": "*/*",
        },
    )

def download_ll_file(url: str, target: Path, cookies_path: Path, referer: str, timeout: float) -> Path:
    target.parent.mkdir(parents=True, exist_ok=True)
    request = Request(url, headers=ll_request_headers(cookies_path, referer))
    tmp_target = target.with_name(f"{target.name}.lltmp")
    with urlopen(request, timeout=timeout) as response, tmp_target.open("wb") as output:
        while True:
            chunk = response.read(1024 * 256)
            if not chunk:
                break
            output.write(chunk)
    tmp_target.replace(target)
    return target

def vortex_mod_folder(config: dict, mod: dict) -> Path:
    staging = Path(str(config.get("vortex_staging_path") or config.get("vortex_mods_path") or ""))
    install_path = str(mod.get("installationPath") or mod.get("id") or "").strip()
    return staging / install_path if staging and install_path else Path()


def vortex_mod_meta_path(config: dict, mod: dict) -> Path:
    return vortex_mod_folder(config, mod) / "meta.ini"


def read_meta_general_path(meta_path: Path) -> dict[str, str]:
    if not meta_path.exists():
        return {}

    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(meta_path, encoding="utf-8-sig")
    except Exception:
        return {}

    if "General" not in parser:
        return {}

    return {str(k).lower(): str(v).strip() for k, v in parser["General"].items()}


def vortex_has_ll_metadata(meta_path: Path) -> bool:
    if not meta_path.exists():
        return False

    parser = configparser.ConfigParser(interpolation=None)
    try:
        parser.read(meta_path, encoding="utf-8-sig")
    except Exception:
        return False

    general = {str(k).lower(): str(v).strip().lower() for k, v in parser["General"].items()} if "General" in parser else {}
    return (
        "LoversLab" in parser
        or "loverslab.com" in general.get("url", "")
        or "loverslab.com" in general.get("website", "")
        or "loverslab.com" in general.get("homepage", "")
    )


def vortex_has_purgeable_nexus_identity(mod: dict, meta_path: Path) -> tuple[bool, str]:
    general = read_meta_general_path(meta_path)
    url = str(general.get("url") or mod.get("source") or "").lower()
    repository = str(general.get("repository") or "").lower()

    mod_id = (
        general.get("modid")
        or general.get("mod_id")
        or general.get("nexusid")
        or str(mod.get("modId") or "")
    )

    file_id = (
        general.get("fileid")
        or general.get("file_id")
        or str(mod.get("fileId") or "")
    )

    def positive_int(value: str) -> bool:
        try:
            return int(str(value or "").strip()) > 0
        except ValueError:
            return False

    has_nexus_id = positive_int(mod_id) or positive_int(file_id)
    has_nexus_url = "nexusmods.com" in url or url.startswith("nxm://")
    has_nexus_repo = repository == "nexus"

    if has_nexus_url:
        return True, "URL points to Nexus"
    if has_nexus_id and has_nexus_repo:
        return True, f"Nexus repository/id marker {mod_id or file_id}"
    if has_nexus_id:
        return True, f"Nexus id marker {mod_id or file_id}"

    return False, ""

def archive_rows(
    downloads_path: Path,
    metadata_path: Path,
    mods_path: Path | None = None,
    vortex_state: dict | None = None,
) -> list[dict]:
    rows: list[dict] = []
    state = vortex_state or {}
    state_downloads = {
        Path(str(download.get("localPath") or download.get("fileName") or "")).name.lower(): download
        for download in state.get("downloads", [])
        if isinstance(download, dict)
        and Path(str(download.get("localPath") or download.get("fileName") or "")).name
    }
    state_mods = {
        archive_id: mod
        for mod in state.get("mods", [])
        if isinstance(mod, dict)
        for archive_id in [str(mod.get("archiveId") or "").strip()]
        if archive_id
    }
    staging_path = Path(str(state.get("stagingPath") or mods_path or ""))
    if downloads_path.exists():
        for archive in sorted(
            [item for item in downloads_path.iterdir() if item.is_file() and item.suffix.lower() in ARCHIVE_SUFFIXES],
            key=lambda item: item.name.lower(),
        ):
            sidecar = archive.with_name(f"{archive.name}.ll.ini")
            metadata_sidecar = metadata_path / "downloads" / f"{archive.name}.ll.ini"
            data = load_ini(sidecar) if sidecar.exists() else load_ini(metadata_sidecar)
            replaced_by = str(data.get("replaced_by") or "").strip()
            if replaced_by and (downloads_path / replaced_by).exists():
                continue
            vortex_download = state_downloads.get(archive.name.lower(), {})
            download_id = str(vortex_download.get("id") or "").strip()
            vortex_mod = state_mods.get(download_id, {}) if download_id else {}
            installed_folder = None
            if vortex_mod and str(vortex_mod.get("state") or "").lower() == "installed":
                install_rel = str(vortex_mod.get("installationPath") or vortex_mod.get("id") or "").strip()
                candidate = staging_path / install_rel if install_rel else Path("")
                installed_folder = candidate if candidate.exists() else candidate
            if installed_folder is None and not state.get("mods"):
                installed_folder = find_installed_folder(staging_path, data, archive) if str(staging_path) else None
            mod_name = (
                str(vortex_mod.get("name") or "")
                if vortex_mod
                else (
                Path(str(installed_folder)).name
                if installed_folder
                else data.get("page_title") or archive.stem
                )
            )
            vortex_status = ""
            if vortex_mod:
                if str(vortex_mod.get("state") or "").lower() == "installed":
                    vortex_status = "Enabled" if vortex_mod.get("enabled") else "Installed"
                else:
                    vortex_status = str(vortex_mod.get("state") or "")
            elif vortex_download:
                vortex_status = "Downloaded"
            rows.append({
                "mod": mod_name,
                "archive": archive.name,
                "path": str(archive),
                "sidecar": str(sidecar if sidecar.exists() else metadata_sidecar if metadata_sidecar.exists() else ""),
                "installed_folder": str(installed_folder or ""),
                "vortex_download_id": str(vortex_download.get("id") or ""),
                "vortex_mod_id": str(vortex_mod.get("id") or "") if vortex_mod else "",
                "vortex_status": vortex_status,
                "source": data.get("source", ""),
                "page_url": data.get("page_url", ""),
                "page_title": data.get("page_title", ""),
                "file_name": data.get("file_pattern") or data.get("file_name", ""),
                "version": data.get("version", ""),
                "fixed": fixed_update(data),
                "update_mode": normalized_update_mode(data.get("update_mode"), fixed=fixed_update(data)),
                "latest": "",
                "has_metadata": "Yes" if data else "No",
            })
    return rows


def find_installed_folder(mods_path: Path | None, data: dict, archive: Path) -> Path | None:
    if not mods_path or not mods_path.exists():
        return None
    candidates = [
        data.get("page_title", ""),
        data.get("file_name", ""),
        archive.stem,
    ]
    prefixes = [filename_prefix(candidate) for candidate in candidates if candidate]
    best = None
    best_score = 0
    try:
        folders = [item for item in mods_path.iterdir() if item.is_dir()]
    except OSError:
        return None
    for folder in folders:
        folder_key = filename_prefix(folder.name)
        for prefix in prefixes:
            score = 0
            if prefix and folder_key == prefix:
                score = 120
            elif prefix and (prefix in folder_key or folder_key in prefix):
                score = 80
            if score > best_score:
                best = folder
                best_score = score
    return best if best_score >= 80 else None

class VortexVoiceSourceFetchWorker(QObject):
    candidatesReady = pyqtSignal(object)
    downloadsReady = pyqtSignal(object)
    statusChanged = pyqtSignal(str)
    finished = pyqtSignal(bool, str)

    def __init__(
        self,
        rows: list[dict],
        source_urls: list[str],
        nexus_api_key: str,
        cookies_path: Path,
        timeout: float = 30.0,
    ) -> None:
        super().__init__()
        self._rows = [dict(row) for row in rows]
        self._source_urls = list(source_urls)
        self._nexus_api_key = str(nexus_api_key or "").strip()
        self._cookies_path = cookies_path
        self._timeout = timeout
        self._cancelled = False
        

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            candidates = []
            all_downloads = []

            nexus_urls = [url for url in self._source_urls if is_nexus_source_url(url)]
            loverslab_urls = [url for url in self._source_urls if is_loverslab_source_url(url)]
            unsupported = len(self._source_urls) - len(nexus_urls) - len(loverslab_urls)

            if not nexus_urls and not loverslab_urls:
                self.finished.emit(False, "No supported source URL found. Add LoversLab file pages or Nexus mod pages.")
                return

            if nexus_urls and not self._nexus_api_key:
                self.statusChanged.emit(
                    f"Nexus API key is missing; skipping {len(nexus_urls)} Nexus source(s)."
                )

            total = len(nexus_urls) + len(loverslab_urls)
            done = 0

            for source_url in loverslab_urls:
                if self._cancelled:
                    self.finished.emit(False, "Cancelled.")
                    return

                done += 1
                self.statusChanged.emit(f"Fetching LoversLab source {done} / {total}")

                try:
                    download_candidates = fetch_loverslab_source_downloads(
                        source_url,
                        self._cookies_path,
                        timeout=self._timeout,
                    )
                except Exception as exc:
                    self.statusChanged.emit(f"LoversLab fetch failed: {exc}")
                    continue

                for download_candidate in download_candidates:
                    all_downloads.append(dict(download_candidate))
                    candidates.extend(self._best_base_matches(download_candidate))

                self.candidatesReady.emit(candidates)
                self.downloadsReady.emit(all_downloads)
                time.sleep(0.2)

            for index, source_url in enumerate(nexus_urls, start=1):
                if self._cancelled:
                    self.finished.emit(False, "Cancelled.")
                    return

                if not self._nexus_api_key:
                    continue

                done += 1
                parts = nexus_source_parts(source_url)
                if not parts:
                    continue

                game, mod_id = parts
                self.statusChanged.emit(f"Fetching Nexus source {done} / {total}: {game} #{mod_id}")

                try:
                    download_candidates = fetch_nexus_files_api(
                        game,
                        mod_id,
                        self._nexus_api_key,
                        timeout=self._timeout,
                    )
                except Exception as exc:
                    self.statusChanged.emit(f"Nexus fetch failed: {exc}")
                    continue

                for download_candidate in download_candidates:
                    all_downloads.append(dict(download_candidate))
                    candidates.extend(self._best_base_matches(download_candidate))                  

                self.candidatesReady.emit(candidates)
                self.downloadsReady.emit(all_downloads)
                time.sleep(0.2)

            self.candidatesReady.emit(candidates)
            self.downloadsReady.emit(all_downloads)

            message = (
                f"Fetched {len(loverslab_urls)} LoversLab source(s), "
                f"{len(nexus_urls) if self._nexus_api_key else 0} Nexus source(s), "
                f"{len(all_downloads)} file(s), {len(candidates)} candidate match(es)."
            )

            if nexus_urls and not self._nexus_api_key:
                message += f" Skipped {len(nexus_urls)} Nexus source(s): missing API key."

            if unsupported:
                message += f" Skipped {unsupported} unsupported source(s)."

            self.finished.emit(True, message)

        except Exception as exc:
            self.finished.emit(False, str(exc))

    def _best_base_matches(self, candidate: dict) -> list[dict]:
        matches: list[dict] = []

        candidate_category = str(candidate.get("voice_category") or "").lower()

        for row in self._rows:
            if row.get("status") == "Ignored":
                continue

            row_category = str(row.get("voice_category") or "").lower()
            if candidate_category and row_category and candidate_category != row_category:
                continue

            score = voice_match_score(
                row.get("base_mod") or "",
                candidate.get("download_name") or candidate.get("file_name") or "",
            )

            if score < DEFAULT_VOICE_MATCH_THRESHOLD:
                continue

            result = dict(candidate)
            result.update({
                "base_mod": row.get("base_mod") or "",
                "base_internal_name": row.get("base_internal_name") or "",
                "base_page_url": row.get("base_page_url") or "",
                "voice_category": row.get("voice_category") or candidate.get("voice_category") or "",
                "voice_category_label": row.get("voice_category_label") or "",
                "online_score": score,
            })
            matches.append(result)

        matches.sort(
            key=lambda item: (
                str(item.get("base_mod") or "").lower(),
                str(item.get("voice_category") or "").lower(),
                -int(item.get("online_score") or 0),
            )
        )
        return matches
    
class VortexVoiceFinder(QDialog):

    COL_STATUS = VOICE_COL_STATUS
    COL_BASE_MOD = VOICE_COL_BASE_MOD
    COL_VOICEPACK = VOICE_COL_VOICEPACK
    COL_DBVO = VOICE_COL_DBVO
    COL_IVDT = VOICE_COL_IVDT
    COL_CATEGORY = VOICE_COL_CATEGORY
    COL_INSTALLED_VOICE = VOICE_COL_INSTALLED_VOICE
    COL_SCORE = VOICE_COL_SCORE
    COL_ONLINE = VOICE_COL_ONLINE
    COL_ONLINE_SCORE = VOICE_COL_ONLINE_SCORE
    COL_SOURCE = VOICE_COL_SOURCE
    COL_PAGE = VOICE_COL_PAGE

    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("LL Integration - Voice Finder")
        self.setWindowFlag(Qt.WindowType.WindowMinMaxButtonsHint, True)
        self.resize(1400, 760)
        self.setMinimumSize(980, 480)
        self.setStyleSheet(DARK_STYLE)

        self.config = read_json(CONFIG_FILE)
        self.vortex_state = load_vortex_state(self.config)

        self.downloads_path = Path(str(self.config.get("vortex_downloads_path") or ""))
        if not str(self.downloads_path) and self.vortex_state.get("downloadsPath"):
            self.downloads_path = Path(str(self.vortex_state.get("downloadsPath")))

        mods_text = str(self.config.get("vortex_mods_path") or self.config.get("vortex_staging_path") or "").strip()
        if not mods_text:
            mods_text = str(self.vortex_state.get("stagingPath") or "").strip()

        self.mods_path = Path(mods_text) if mods_text else None
        self.metadata_path = Path(str(self.config.get("metadata_path") or BASE_DIR / "metadata"))

        self.cookies_path = configured_cookies_path(self.config)

        self.rows = self._build_initial_rows()
        self.online_candidates: list[dict] = []
        self.source_downloads: list[dict] = []
        self.fetch_thread: QThread | None = None
        self.fetch_worker: VortexVoiceSourceFetchWorker | None = None
        self._build_ui()

    def _voice_config_path(self) -> Path:
        return BASE_DIR / "voice_finder.json"

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

    def _save_voice_config(self, config: dict) -> None:
        path = self._voice_config_path()
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(config, indent=2), encoding="utf-8")

    def _local_match_threshold(self, config: dict | None) -> int:
        try:
            return max(0, min(100, int((config or {}).get("localMatchThreshold") or DEFAULT_VOICE_MATCH_THRESHOLD)))
        except (TypeError, ValueError):
            return DEFAULT_VOICE_MATCH_THRESHOLD
            
    def _handle_nexus_link(self) -> None:
        config = self._load_voice_config()

        dialog = QDialog(self)
        dialog.setWindowTitle("Handle Nexus Link")
        dialog.resize(720, 260)
        dialog.setStyleSheet(DARK_STYLE)

        help_text = QLabel(
            "Paste a Nexus Mods personal API key here.\n\n"
            "It is stored locally in LL Integration's voice_finder.json and is used only to list Nexus file downloads for source scoring."
        )
        help_text.setWordWrap(True)

        key_input = QLineEdit(dialog)
        key_input.setEchoMode(QLineEdit.EchoMode.Password)
        key_input.setPlaceholderText("Nexus Mods API key")
        key_input.setText(str(config.get("nexusApiKey") or ""))

        status = QLabel("")
        status.setWordWrap(True)

        def set_api_status(message: str, ready: bool) -> None:
            color = "#0f5f36" if ready else "#6f2323"
            border = "#35d07f" if ready else "#e05c5c"
            status.setText(message)
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
            self.progress_label.setText("Nexus API key saved.")

        def validate() -> None:
            api_key = key_input.text().strip()
            if not api_key:
                QMessageBox.information(dialog, "LL Integration", "Paste a Nexus API key first.")
                return

            try:
                payload = validate_nexus_api_key(api_key, timeout=30.0)
            except Exception as exc:
                set_api_status("Nexus API key validation failed.", False)
                QMessageBox.warning(dialog, "LL Integration", f"Nexus API key validation failed:\n\n{exc}")
                return

            name = str(payload.get("name") or payload.get("user_id") or "").strip()
            set_api_status(f"Nexus API key validated{f' for {name}' if name else ''}.", True)
            self.progress_label.setText("Nexus API key validated.")

        open_api.clicked.connect(
            lambda _checked=False: webbrowser.open("https://www.nexusmods.com/users/myaccount?tab=api")
        )
        save_key.clicked.connect(lambda _checked=False: save())
        validate_key.clicked.connect(lambda _checked=False: validate())
        clear_key.clicked.connect(
            lambda _checked=False: (
                key_input.clear(),
                set_api_status("Key field cleared. Save to remove it from config.", False),
            )
        )
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
        layout.addLayout(key_row)
        layout.addWidget(status)
        layout.addLayout(buttons)

        dialog.exec()

    def _edit_source_urls(self) -> None:
        config = self._load_voice_config()

        dialog = QDialog(self)
        dialog.setWindowTitle("Voice Source URLs")
        dialog.resize(780, 420)
        dialog.setStyleSheet(DARK_STYLE)

        text = QTextEdit(dialog)
        text.setPlainText("\n".join(config.get("voiceSourceUrls", [])))
        text.setPlaceholderText("One LoversLab or Nexus Mods source URL per line")

        help_text = QLabel(
            "Add LoversLab download pages or Nexus Mods pages.\n\n"
            "LoversLab URLs are normalized to ?do=download.\n"
            "Nexus URLs are normalized to their Files tab automatically."
        )
        help_text.setWordWrap(True)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Save | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        layout = QVBoxLayout(dialog)
        layout.addWidget(help_text)
        layout.addWidget(text)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return

        urls = []
        rejected = []

        for line in text.toPlainText().splitlines():
            value = line.strip()
            if not value or value.startswith("#"):
                continue

            normalized = normalize_voice_source_url(value)

            if not normalized.startswith(("http://", "https://")):
                rejected.append(value)
                continue

            if normalized not in urls:
                urls.append(normalized)

        config["voiceSourceUrls"] = urls
        self._save_voice_config(config)

        message = f"Saved {len(urls)} voice source URL(s)."
        if rejected:
            message += "\n\nIgnored invalid line(s):\n" + "\n".join(rejected[:8])
            if len(rejected) > 8:
                message += f"\n...and {len(rejected) - 8} more."

        QMessageBox.information(self, "LL Integration", message)
        self.progress_label.setText(message)

    def _base_key(self, base: dict) -> str:
        return str(
            base.get("internal_name")
            or base.get("base_internal_name")
            or base.get("display_name")
            or base.get("base_mod")
            or ""
        ).lower()

    def _slot_key(self, base: dict, category: str) -> str:
        return f"{self._base_key(base)}|{category}"

    def _manual_match_key(self, base: dict, category: str) -> str:
        return self._slot_key(base, category)

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
            if str(mod.get("internal_name") or "").lower() == self._base_key(base):
                return None
            return {
                "display_name": mod.get("display_name") or entry.get("display_name") or "",
                "internal_name": mod.get("internal_name") or entry.get("internal_name") or "",
                "mod_path": mod.get("mod_path") or entry.get("mod_path") or "",
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
            if str(voice.get("internal_name") or "").lower() == self._base_key(base):
                continue
            if self._is_false_match(config, base, voice.get("display_name") or "", "", category):
                continue

            score = voice_match_score(base.get("display_name") or base.get("base_mod") or "", voice.get("display_name") or "")
            guessed_category = voice_category_guess(voice.get("display_name") or "")
            if guessed_category == category:
                score += 20
            elif guessed_category:
                score = 0
            else:
                score = max(0, score - 35)

            candidates.append({
                "display_name": voice.get("display_name") or "",
                "internal_name": voice.get("internal_name") or "",
                "mod_path": voice.get("mod_path") or "",
                "score": score,
                "voice_category": guessed_category,
            })

        return sorted(
            candidates,
            key=lambda item: (-int(item.get("score") or 0), str(item.get("display_name") or "").lower()),
        )

    def _build_initial_rows(self) -> list[dict]:
        config = self._load_voice_config()
        threshold = self._local_match_threshold(config)
        forced_voice_mods = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base_mods = {str(value).lower() for value in config.get("forcedBaseMods", [])}
        ignored_base_mods = {str(value).lower() for value in config.get("ignoredBaseMods", [])}
        complete_slots = {str(value).lower() for value in config.get("completeVoiceSlots", [])}
        none_slots = {str(value).lower() for value in config.get("noneVoiceSlots", [])}

        metadata_rows = []
        for row in archive_rows(self.downloads_path, self.metadata_path, self.mods_path, self.vortex_state):
            if row.get("has_metadata") == "Yes":
                metadata_rows.append(row)

        installed_by_key: dict[str, dict] = {}
        for mod in self._vortex_installed_mods():
            internal_name = str(mod.get("internal_name") or mod.get("id") or "").strip()
            if internal_name:
                installed_by_key[internal_name.lower()] = mod

        all_mods: list[dict] = []
        seen_mods: set[str] = set()

        for row in metadata_rows:
            internal_name = str(row.get("mod_id") or row.get("mod") or row.get("archive") or row.get("page_title") or "").strip()
            display_name = str(row.get("mod") or row.get("page_title") or row.get("archive") or internal_name).strip()
            if not internal_name:
                internal_name = display_name
            key = internal_name.lower()
            installed = installed_by_key.get(key) or installed_by_key.get(display_name.lower()) or {}
            mod_path = installed.get("mod_path") or row.get("installed_folder") or ""
            auto_voice = voice_keyword_present(display_name)
            is_voice = (auto_voice or key in forced_voice_mods) and key not in forced_base_mods
            all_mods.append({
                "internal_name": internal_name,
                "display_name": display_name,
                "mod_path": mod_path,
                "is_voice": is_voice,
                "auto_is_voice": auto_voice,
                "classification_override": "voice" if key in forced_voice_mods else "base" if key in forced_base_mods else "auto",
                "base_page_url": row.get("page_url") or "",
                "raw": row,
            })
            seen_mods.add(key)

        for mod in self._vortex_installed_mods():
            internal_name = str(mod.get("internal_name") or mod.get("id") or "").strip()
            display_name = str(mod.get("display_name") or mod.get("name") or internal_name).strip()
            key = internal_name.lower()
            if not key or key in seen_mods:
                continue
            auto_voice = voice_keyword_present(display_name)
            is_voice = (auto_voice or key in forced_voice_mods) and key not in forced_base_mods
            if not is_voice:
                continue
            all_mods.append({
                "internal_name": internal_name,
                "display_name": display_name,
                "mod_path": mod.get("mod_path") or "",
                "is_voice": True,
                "auto_is_voice": auto_voice,
                "classification_override": "voice" if key in forced_voice_mods else "auto",
                "base_page_url": "",
                "raw": mod,
            })

        installed_voice_mods = [item for item in all_mods if item.get("is_voice")]
        base_mods = [item for item in all_mods if not item.get("is_voice") and item.get("base_page_url")]

        rows = []
        for base in base_mods:
            ignored = self._base_key(base) in ignored_base_mods
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
                    ignored,
                    slot_key.lower() in complete_slots,
                    slot_key.lower() in none_slots,
                    bool(best_voice.get("manual")),
                    int(best_voice.get("score") or 0),
                    False,
                    threshold,
                )

                rows.append({
                    "status": slot_status,
                    "slot_status": slot_status,
                    "base_mod": base.get("display_name") or "",
                    "base_internal_name": base.get("internal_name") or "",
                    "voicepack": "",
                    "dbvo": "",
                    "IVDT": "",
                    "voice_category": category,
                    "voice_category_label": category_label,
                    "slot_key": slot_key,
                    "classification_override": base.get("classification_override") or "auto",
                    "base_page_url": base.get("base_page_url") or "",
                    "installed_voice": best_voice.get("display_name") or "",
                    "installed_voice_internal_name": best_voice.get("internal_name") or "",
                    "score": best_voice.get("score") or 0,
                    "manual_voice": bool(best_voice.get("manual")),
                    "complete_voice": slot_key.lower() in complete_slots,
                    "none_voice": slot_key.lower() in none_slots,
                    "local_match_threshold": threshold,
                    "installed_voice_candidates": voice_candidates,
                    "online_candidate": "",
                    "online_download_url": "",
                    "online_source_url": "",
                    "online_score": 0,
                    "source": "",
                    "download_url": "",
                    "source_type": "",
                    "raw": base.get("raw") or {},
                    "search_query": f"{voice_search_query(base.get('display_name') or '')} {category_label}",
                })

        self._attach_voice_overview(rows)
        category_order = {category: index for index, (category, _label) in enumerate(VOICE_CATEGORIES)}
        rows.sort(key=lambda row: (
            str(row.get("base_mod") or "").lower(),
            category_order.get(str(row.get("voice_category") or ""), 99),
        ))
        return rows

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

        display_rows.sort(key=lambda item: str(item.get("base_mod") or "").lower())
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

    def _voice_row_background_color(self, row: dict) -> QColor:
        status = str(row.get("slot_status") or row.get("status") or "")

        if status == "Installed":
            return QColor(38, 74, 48)
        if status == "Possible":
            return QColor(76, 63, 31)
        if status == "Online found":
            return QColor(44, 67, 88)
        if status == "Complete":
            return QColor(44, 67, 88)
        if status == "Ignored":
            return QColor(55, 55, 55)
        if status == "None":
            return QColor(45, 45, 55)

        return QColor(82, 38, 38)

    def _overview_background_color(self, text: str) -> QColor:
        value = str(text or "")

        if value.startswith(("Installed:", "Possible:", "Manual:", "OK:")):
            return QColor(38, 74, 48)
        if value.startswith("Online:"):
            return QColor(44, 67, 88)
        if value == "Complete":
            return QColor(44, 67, 88)
        if value == "None":
            return QColor(45, 45, 55)
        if value == "Ignored":
            return QColor(55, 55, 55)

        return QColor(82, 38, 38)

    def _fill_display_table_row(self, row_index: int, display_row: dict) -> None:
        slots = display_row.get("slots") or {}

        voicepack = slots.get("npc")
        dbvo = slots.get("player")
        IVDT = slots.get("scene")

        self._set_item(row_index, self.COL_STATUS, display_row.get("status", ""))
        self._set_item(row_index, self.COL_BASE_MOD, display_row.get("base_mod", ""))
        self._set_item(row_index, self.COL_VOICEPACK, self._voice_overview_text(voicepack) or "Missing")
        self._set_item(row_index, self.COL_DBVO, self._voice_overview_text(dbvo) or "Missing")
        self._set_item(row_index, self.COL_IVDT, self._voice_overview_text(IVDT) or "Missing")
        self._set_item(row_index, self.COL_CATEGORY, "")
        self._set_item(row_index, self.COL_INSTALLED_VOICE, display_row.get("installed_voice", ""))
        self._set_item(row_index, self.COL_SCORE, str(display_row.get("score") or ""))
        self._set_item(row_index, self.COL_ONLINE, display_row.get("online_candidate", ""))
        self._set_item(row_index, self.COL_ONLINE_SCORE, str(display_row.get("online_score") or ""))
        self._set_item(row_index, self.COL_SOURCE, "")
        self._set_item(row_index, self.COL_PAGE, display_row.get("base_page_url", ""))

        self._apply_display_row_background(row_index, display_row)

    def _apply_display_row_background(self, row_index: int, display_row: dict) -> None:
        color = self._voice_row_background_color(display_row)

        for column in range(self.table.columnCount()):
            item = self.table.item(row_index, column)
            if item is None:
                continue

            if column in (self.COL_VOICEPACK, self.COL_DBVO, self.COL_IVDT):
                item.setBackground(self._overview_background_color(item.text()))
            else:
                item.setBackground(color)

            item.setForeground(QColor(242, 242, 242))

    def _refresh_display_rows(self) -> None:
        self._attach_voice_overview(self.rows)
        display_rows = self._build_voice_display_rows(self.rows)

        self.table._ll_voice_rows = self.rows
        self.table._ll_voice_display_rows = display_rows
        self.table.setRowCount(len(display_rows))

        self.table.setUpdatesEnabled(False)
        try:
            for row_index, display_row in enumerate(display_rows):
                self._fill_display_table_row(row_index, display_row)
        finally:
            self.table.setUpdatesEnabled(True)

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
            status == "Installed" or (status == "Complete" and str(row.get("installed_voice") or "").strip())
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
            row["voicepack"] = row["voice_overview"].get("npc", "")
            row["dbvo"] = row["voice_overview"].get("player", "")
            row["IVDT"] = row["voice_overview"].get("scene", "")

    def _refresh_voice_base(self, target_row: dict) -> None:
        self._refresh_display_rows()

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

    def _fetch_sources(
        self,
        fetch_button: QPushButton,
        sources_button: QPushButton,
        download_button: QPushButton,
    ) -> None:
        config = self._load_voice_config()
        urls = [str(url).strip() for url in config.get("voiceSourceUrls", []) if str(url).strip()]

        if not urls:
            QMessageBox.information(
                self,
                "LL Integration",
                "Add at least one voice source URL first.",
            )
            return

        nexus_urls = [url for url in urls if is_nexus_source_url(url)]
        loverslab_urls = [url for url in urls if is_loverslab_source_url(url)]

        if not nexus_urls and not loverslab_urls:
            QMessageBox.information(
                self,
                "LL Integration",
                "No supported source URL found.\n\nAdd LoversLab file pages or Nexus Mods pages.",
            )
            return

        api_key = str(config.get("nexusApiKey") or "").strip()

        if nexus_urls and not api_key and not loverslab_urls:
            QMessageBox.information(
                self,
                "LL Integration",
                "Nexus API key is missing.\n\nUse Handle Nexus Link first, or add a LoversLab source URL.",
            )
            return

        self.fetch_thread = QThread(self)
        self.fetch_worker = VortexVoiceSourceFetchWorker(
            self.rows,
            urls,
            api_key,
            self.cookies_path,
            timeout=30.0,
        )
        self.fetch_worker.moveToThread(self.fetch_thread)

        fetch_button.setEnabled(False)
        sources_button.setEnabled(False)
        download_button.setEnabled(False)
        self.progress_label.setText("Fetching voice sources...")

        self.fetch_worker.statusChanged.connect(self.progress_label.setText)
        self.fetch_worker.candidatesReady.connect(lambda candidates: self._apply_online_candidates(list(candidates)))
        self.fetch_worker.downloadsReady.connect(lambda downloads: self._store_source_downloads(list(downloads)))
        self.fetch_worker.finished.connect(
            lambda ok, message: self._fetch_sources_finished(
                self.fetch_thread,
                self.fetch_worker,
                fetch_button,
                sources_button,
                download_button,
                ok,
                message,
            )
        )

        self.fetch_thread.started.connect(self.fetch_worker.run)
        self.fetch_thread.start()

    def _store_source_downloads(self, downloads: list[dict]) -> None:
        self.source_downloads = list(downloads)

    def _apply_online_candidates(self, candidates: list[dict]) -> None:
        merged: dict[tuple[str, str, str, str], dict] = {
            self._candidate_identity(candidate): dict(candidate)
            for candidate in self.online_candidates
        }

        for candidate in candidates or []:
            candidate = dict(candidate)
            key = self._candidate_identity(candidate)
            if not key[0] or not key[1] or not key[2]:
                continue
            merged[key] = candidate

        self.online_candidates = list(merged.values())

        self._apply_online_candidates_to_rows()
        self._attach_voice_overview(self.rows)
        self._populate_table()
        self._apply_filter()
        self._update_summary()
        self._update_selected_label()

    def _fetch_sources_finished(
        self,
        thread: QThread,
        worker: VortexVoiceSourceFetchWorker,
        fetch_button: QPushButton,
        sources_button: QPushButton,
        download_button: QPushButton,
        ok: bool,
        message: str,
    ) -> None:
        self.progress_label.setText(message)

        fetch_button.setEnabled(True)
        sources_button.setEnabled(True)
        download_button.setEnabled(True)

        try:
            thread.quit()
            thread.wait(3000)
        except Exception:
            pass

        try:
            worker.deleteLater()
            thread.deleteLater()
        except Exception:
            pass

        self.fetch_thread = None
        self.fetch_worker = None

        if not ok:
            QMessageBox.warning(self, "LL Integration", message)

    def _download_target_rows(self) -> list[dict]:
        rows = []

        for row in list(getattr(self, "rows", []) or []):
            if row.get("_overview_duplicate"):
                continue
            rows.append(dict(row))

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

        for mod in self._vortex_installed_mods():
            internal_name = str(mod.get("internal_name") or mod.get("id") or "").strip()
            display_name = str(mod.get("display_name") or mod.get("name") or internal_name).strip()
            key = internal_name.lower()

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

        mods = tuple(
            sorted(
                (
                    str(item.get("internal_name") or item.get("id") or ""),
                    str(item.get("display_name") or item.get("name") or ""),
                )
                for item in self._vortex_installed_mods()
            )
        )

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

        return (source_rows, mods, forced, forced_base, none_slots, complete_slots, archives)

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
        if not self.downloads_path or not self.downloads_path.exists():
            return set()

        try:
            return {path.name.lower() for path in self.downloads_path.iterdir() if path.is_file()}
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
        download_url = str(download.get("download_url") or "").strip()

        if source_url or download_url:
            menu.addSeparator()

        if source_url:
            open_source = QAction("Open source page", menu)
            open_source.triggered.connect(lambda _checked=False, url=source_url: webbrowser.open(url))
            menu.addAction(open_source)

        if download_url:
            open_download = QAction("Open download URL", menu)
            open_download.triggered.connect(
                lambda _checked=False, url=urljoin(source_url, download_url): webbrowser.open(url)
            )
            menu.addAction(open_download)

        menu.exec(table.viewport().mapToGlobal(pos))

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

    def _candidate_from_online_table(self, table: QTableWidget) -> dict | None:
        selected = table.selectedItems()
        if not selected:
            return None

        candidates = getattr(table, "_ll_voice_candidates", [])
        index = selected[0].row()

        if index < 0 or index >= len(candidates):
            return None

        return dict(candidates[index])
    def _vortex_installed_mods(self) -> list[dict]:
        state = self.vortex_state or {}
        staging_path = Path(str(state.get("stagingPath") or self.mods_path or ""))

        mods = []
        raw_mods = state.get("mods", [])

        if not isinstance(raw_mods, list):
            raw_mods = []

        for item in raw_mods:
            if not isinstance(item, dict):
                continue

            mod_id = str(item.get("id") or "").strip()
            display_name = str(item.get("name") or item.get("displayName") or mod_id).strip()
            install_rel = str(item.get("installationPath") or mod_id).strip()
            mod_path = staging_path / install_rel if str(staging_path) and install_rel else Path("")

            state_text = str(item.get("state") or "").lower()
            if state_text and state_text != "installed":
                continue

            mods.append({
                "id": mod_id,
                "internal_name": mod_id,
                "display_name": display_name,
                "name": display_name,
                "mod_path": str(mod_path),
                "enabled": bool(item.get("enabled")),
                "state": item.get("state") or "",
                "archive_id": item.get("archiveId") or "",
            })

        if mods:
            return mods

        # Fallback si le snapshot Vortex ne contient pas encore les mods.
        if self.mods_path and self.mods_path.exists():
            try:
                for folder in self.mods_path.iterdir():
                    if not folder.is_dir():
                        continue
                    mods.append({
                        "id": folder.name,
                        "internal_name": folder.name,
                        "display_name": folder.name,
                        "name": folder.name,
                        "mod_path": str(folder),
                        "enabled": False,
                        "state": "",
                        "archive_id": "",
                    })
            except OSError:
                pass

        return mods

    def _collect_voice_mods_inventory(self, show_message: bool = True) -> list[dict]:
        config = self._load_voice_config()
        forced_voice_mods = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base_mods = {str(value).lower() for value in config.get("forcedBaseMods", [])}

        voice_mods = []

        for mod in self._vortex_installed_mods():
            internal_name = str(mod.get("internal_name") or mod.get("id") or "").strip()
            display_name = str(mod.get("display_name") or mod.get("name") or internal_name).strip()
            key = internal_name.lower()

            is_voice = (voice_keyword_present(display_name) or key in forced_voice_mods) and key not in forced_base_mods

            if not is_voice:
                continue

            path = Path(str(mod.get("mod_path") or ""))
            install_time = path.stat().st_mtime if path.exists() else 0

            voice_mods.append({
                "display_name": display_name,
                "internal_name": internal_name,
                "mod_path": str(path),
                "installed_at": time.strftime("%Y-%m-%d %H:%M", time.localtime(install_time)) if install_time else "",
                "install_time": install_time,
                "classification": "Forced voice" if key in forced_voice_mods else "Auto voice",
            })

        if show_message and not voice_mods:
            QMessageBox.information(self, "LL Integration", "No installed voice-like mods were found.")

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
        table.customContextMenuRequested.connect(
            lambda pos: self._show_voice_inventory_context_menu(table, pos)
        )

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
        item = table.itemAt(pos)
        if item is None:
            return

        table.selectRow(item.row())

        rows = getattr(table, "_ll_voice_inventory_visible", [])
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

        mod_path = str(voice.get("mod_path") or "")
        if mod_path:
            open_folder = QAction("Open folder", menu)
            open_folder.triggered.connect(
                lambda _checked=False, path=mod_path: self._open_path(Path(path)) if Path(path).exists() else None
            )
            menu.addAction(open_folder)

        menu.exec(table.viewport().mapToGlobal(pos))

    def _open_path(self, path: Path) -> None:
        target = path if path.is_dir() else path.parent

        try:
            webbrowser.open(target.as_uri())
        except ValueError:
            webbrowser.open(str(target))

    def _download_selected_voice_candidate(self) -> None:
        row = self._selected_context_row() or self._selected_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a base mod row first.")
            return

        candidate = self._candidate_from_main_row(row)
        if not candidate:
            QMessageBox.information(
                self,
                "LL Integration",
                "Selected row has no online candidate yet.\n\n"
                "Use Fetch sources first, or use Show All Source Downloads to manually pick a candidate.",
            )
            return

        target = {
            "base_mod": row.get("base_mod") or "",
            "base_internal_name": row.get("base_internal_name") or row.get("base_mod") or "",
            "voice_category": row.get("voice_category") or candidate.get("voice_category") or "npc",
            "voice_category_label": VOICE_CATEGORY_LABELS.get(
                row.get("voice_category") or candidate.get("voice_category") or "npc",
                row.get("voice_category") or candidate.get("voice_category") or "npc",
            ),
            "base_page_url": row.get("base_page_url") or "",
        }

        self._download_source_candidate(candidate, target, parent=self)

    def _candidate_from_main_row(self, row: dict) -> dict | None:
        candidates = list(row.get("online_candidates") or [])
        if candidates:
            candidates = sorted(
                [dict(candidate) for candidate in candidates],
                key=lambda item: int(item.get("online_score") or 0),
                reverse=True,
            )
            candidate = candidates[0]
        else:
            download_name = str(row.get("online_candidate") or "").strip()
            download_url = str(row.get("online_download_url") or "").strip()
            source_url = str(row.get("online_source_url") or "").strip()

            if not download_name and not download_url:
                return None

            candidate = {
                "download_name": download_name,
                "file_name": row.get("file_name") or row.get("archive_file_name") or download_name,
                "archive_file_name": row.get("archive_file_name") or row.get("file_name") or download_name,
                "download_url": download_url,
                "source_url": source_url,
                "source_type": row.get("source_type") or "",
                "online_score": row.get("online_score") or 0,
                "size": row.get("online_size") or "",
                "date_iso": row.get("online_date_iso") or "",
                "version": row.get("online_version") or "",
            }

        candidate.update({
            "base_mod": row.get("base_mod") or "",
            "base_internal_name": row.get("base_internal_name") or row.get("base_mod") or "",
            "base_page_url": row.get("base_page_url") or "",
            "voice_category": row.get("voice_category") or candidate.get("voice_category") or "npc",
            "voice_category_label": row.get("voice_category_label")
                or candidate.get("voice_category_label")
                or VOICE_CATEGORY_LABELS.get(row.get("voice_category") or candidate.get("voice_category") or "npc", "Voicepack"),
        })

        return candidate

    def _download_source_candidate(self, candidate: dict, target_row: dict, parent: QDialog | None = None) -> None:
        parent = parent or self

        archive_name = archive_file_name_from_candidate(candidate)
        target_path = self.downloads_path / archive_name

        if not self.downloads_path:
            QMessageBox.warning(parent, "LL Integration", "Vortex downloads path is not configured.")
            return

        if not self.downloads_path.exists():
            try:
                self.downloads_path.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                QMessageBox.warning(parent, "LL Integration", f"Could not create downloads folder:\n{exc}")
                return

        if target_path.exists():
            choice = QMessageBox.question(
                parent,
                "LL Integration",
                "This archive already exists in the Vortex downloads folder.\n\n"
                f"{target_path.name}\n\n"
                "Yes = overwrite and download again\n"
                "No = queue the existing archive for Vortex install\n"
                "Cancel = do nothing",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No | QMessageBox.StandardButton.Cancel,
                QMessageBox.StandardButton.No,
            )

            if choice == QMessageBox.StandardButton.Cancel:
                return

            if choice == QMessageBox.StandardButton.No:
                self._write_voice_sidecars(target_path, candidate, target_row)
                self._queue_voice_install(target_path, candidate, target_row)
                return

        if not self._confirm_voice_download(candidate, target_row, target_path, parent):
            return

        self.progress_label.setText(f"Downloading {archive_name}...")
        QApplication.processEvents()

        try:
            downloaded = self._download_candidate_file(candidate, target_path)
            self._write_voice_sidecars(downloaded, candidate, target_row)
            self._queue_voice_install(downloaded, candidate, target_row)
        except Exception as exc:
            self.progress_label.setText("Voice download failed.")
            QMessageBox.warning(parent, "LL Integration", f"Voice download failed:\n\n{exc}")
            return

        self.progress_label.setText(f"Downloaded and queued for Vortex: {downloaded.name}")

        QMessageBox.information(
            parent,
            "LL Integration",
            "Voice archive downloaded and queued for Vortex install.\n\n"
            f"Archive: {downloaded.name}\n"
            f"Target: {target_row.get('base_mod') or ''}\n"
            f"Type: {target_row.get('voice_category_label') or target_row.get('voice_category') or ''}\n\n"
            "Keep Vortex open; the LL Integration extension will start the archive install.",
        )

    def _confirm_voice_download(self, candidate: dict, target_row: dict, target_path: Path, parent: QDialog) -> bool:
        source_type = str(candidate.get("source_type") or "").strip() or "unknown"
        return QMessageBox.question(
            parent,
            "LL Integration",
            "Download and install this voice candidate?\n\n"
            f"Target base mod: {target_row.get('base_mod') or ''}\n"
            f"Voice type: {target_row.get('voice_category_label') or target_row.get('voice_category') or ''}\n"
            f"Download: {candidate.get('download_name') or ''}\n"
            f"Source: {source_type}\n"
            f"Archive: {target_path.name}\n\n"
            "The archive will be saved to the Vortex downloads folder, sidecar metadata will be written, "
            "then an install request will be queued for Vortex.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes,
        ) == QMessageBox.StandardButton.Yes

    def _download_candidate_file(self, candidate: dict, target_path: Path) -> Path:
        source_type = str(candidate.get("source_type") or "").strip().lower()
        download_url = str(candidate.get("download_url") or "").strip()
        source_url = str(candidate.get("source_url") or "").strip()

        if source_type == "nexus" or is_nexus_source_url(source_url):
            config = self._load_voice_config()
            api_key = str(config.get("nexusApiKey") or "").strip()
            if not api_key:
                raise RuntimeError("Nexus API key is missing. Use Handle Nexus Link first.")
            return download_nexus_candidate(candidate, target_path, api_key, timeout=VOICE_DOWNLOAD_TIMEOUT_SECONDS)

        if "loverslab.com" in download_url.lower() or "loverslab.com" in source_url.lower():
            if not download_url:
                raise RuntimeError("LoversLab candidate is missing its download URL.")
            return download_ll_file(
                download_url,
                target_path,
                self.cookies_path,
                source_url or download_url,
                VOICE_DOWNLOAD_TIMEOUT_SECONDS,
            )

        if download_url.startswith(("http://", "https://")):
            return download_raw_url(download_url, target_path, timeout=VOICE_DOWNLOAD_TIMEOUT_SECONDS)

        raise RuntimeError("Candidate has no usable download URL.")

    def _write_voice_sidecars(self, archive_path: Path, candidate: dict, target_row: dict) -> None:
        source_type = str(candidate.get("source_type") or "").strip()
        source_url = str(candidate.get("source_url") or "").strip()
        download_url = str(candidate.get("download_url") or "").strip()
        base_mod = str(target_row.get("base_mod") or "").strip()
        base_internal = str(target_row.get("base_internal_name") or base_mod).strip()
        category = str(target_row.get("voice_category") or candidate.get("voice_category") or "npc").strip()
        category_label = VOICE_CATEGORY_LABELS.get(category, category)

        display_source = "LoversLab" if "loverslab.com" in source_url.lower() else "Nexus Mods"

        data = {
            "source": source_type or ("nexus" if is_nexus_source_url(source_url) else "loverslab"),
            "display_source": display_source,
            "mod_homepage": source_url,
            "page_url": source_url,
            "download_url": download_url,
            "file_name": archive_path.name,
            "archive_name": archive_path.name,
            "original_archive_name": archive_path.name,
            "version": str(candidate.get("version") or extract_version(archive_path.name) or ""),
            "update_mode": UPDATE_MODE_MANUAL,
            "fixed_version": "false",
            "manual_update": "false",
            "skip_update_check": "false",

            # Voice Finder metadata
            "ll_integration_kind": "voice_pack",
            "voice_for_base_mod": base_mod,
            "voice_for_base_internal_name": base_internal,
            "voice_category": category,
            "voice_category_label": category_label,
            "voice_source_download_name": clean_source_download_name(candidate.get("download_name") or ""),
            "voice_source_raw_download_name": str(candidate.get("download_name") or ""),
            "voice_source_score": str(candidate.get("online_score") or ""),
            "voice_source_size": str(candidate.get("size") or ""),
            "voice_source_date": str(candidate.get("date_iso") or ""),
            "nexus_file_id": str(candidate.get("nexus_file_id") or ""),
            
        }

        sidecar = archive_path.with_name(f"{archive_path.name}.ll.ini")
        metadata_sidecar = self.metadata_path / "downloads" / f"{archive_path.name}.ll.ini"

        write_ini(sidecar, data)
        write_ini(metadata_sidecar, data)

    def _queue_voice_install(self, archive_path: Path, candidate: dict, target_row: dict) -> None:
        game_id = self.vortex_state.get("activeGameId") or self.config.get("active_vortex_game") or "skyrimse"
        profile_id = self.vortex_state.get("activeProfileId") or self.config.get("active_vortex_profile_id") or ""

        source_type = str(candidate.get("source_type") or "").strip()
        source_url = str(candidate.get("source_url") or "").strip()
        download_url = str(candidate.get("download_url") or "").strip()
        is_ll = "loverslab.com" in source_url.lower() or "loverslab.com" in download_url.lower()
        display_source = "LoversLab" if is_ll else "Nexus Mods"

        command_id = append_vortex_command({
            "action": "install_archive",
            "operation": "install_voice",
            "archivePath": str(archive_path),
            "gameId": game_id,
            "profileId": profile_id,
            "allowAutoEnable": True,
            "enableAfterInstall": True,

            "voicePack": True,
            "voiceForBaseMod": target_row.get("base_mod") or "",
            "voiceForBaseInternalName": target_row.get("base_internal_name") or target_row.get("base_mod") or "",
            "voiceCategory": target_row.get("voice_category") or candidate.get("voice_category") or "npc",

            "sourceType": source_type or ("nexus" if is_nexus_source_url(source_url) else "loverslab"),
            "sourceName": "Website",
            "displaySource": display_source,
            "sourceUrl": source_url,
            "pageUrl": source_url,
            "modHomepage": source_url,
            "downloadUrl": download_url,
            "downloadName": clean_source_download_name(candidate.get("download_name") or "") or display_name_from_archive_name(archive_path.name) or archive_path.name,
            "archiveName": archive_path.name,
            "version": str(candidate.get("version") or extract_version(archive_path.name) or ""),
            "nexusFileId": str(candidate.get("nexus_file_id") or ""),
        })

        self.progress_label.setText(f"Voice install queued for Vortex ({command_id}).")
        self._schedule_vortex_refresh()

    def _show_all_source_downloads(self) -> None:
        downloads = list(getattr(self, "source_downloads", []) or [])

        if not downloads:
            QMessageBox.information(
                self,
                "LL Integration",
                "No fetched downloads are available yet.\n\nUse Fetch sources first.",
            )
            return

        cache_key = self._all_source_downloads_cache_key(downloads)
        cached = getattr(self, "_all_source_downloads_score_cache", None)

        if isinstance(cached, dict) and cached.get("key") == cache_key:
            scored_downloads = [dict(item) for item in cached.get("rows", [])]
            dialog = QDialog(self)
            dialog.setWindowTitle("All fetched source downloads")
            dialog.resize(1160, 620)
            dialog.setStyleSheet(DARK_STYLE)

            title = QLabel("All fetched source downloads")
            title.setStyleSheet("font-weight: 700;")
            hint = QLabel(
                "Each fetched archive is scored against all installed base mods. "
                "Filter or sort, then download using the best target shown in the row."
            )
            hint.setWordWrap(True)

            layout = QVBoxLayout(dialog)
            dialog.setLayout(layout)
        else:
            dialog = QDialog(self)
            dialog.setWindowTitle("All fetched source downloads")
            dialog.resize(1160, 620)
            dialog.setStyleSheet(DARK_STYLE)

            title = QLabel("All fetched source downloads")
            title.setStyleSheet("font-weight: 700;")
            hint = QLabel(
                "Each fetched archive is scored against all installed base mods. "
                "Filter or sort, then download using the best target shown in the row."
            )
            hint.setWordWrap(True)

            loading_label = QLabel("Preparing scored download list...")
            loading_progress = QProgressBar(dialog)
            loading_progress.setRange(0, max(1, len(downloads)))

            layout = QVBoxLayout(dialog)
            layout.addWidget(title)
            layout.addWidget(hint)
            layout.addWidget(loading_label)
            layout.addWidget(loading_progress)
            dialog.setLayout(layout)
            dialog.show()
            QApplication.processEvents()

            voice_rows = self._download_target_rows()
            voice_inventory = self._collect_voice_mods_inventory(show_message=False)
            downloaded_archives = self._downloaded_archive_names()
            score_cache: dict[tuple[str, str], int] = {}
            scored_downloads = []

            for index, download in enumerate(downloads, start=1):
                candidate = dict(download)
                target, score = self._best_download_target(voice_rows, candidate, score_cache)
                candidate["online_score"] = score
                candidate["already_status"] = self._download_already_status(
                    candidate,
                    voice_inventory,
                    downloaded_archives,
                )

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

            self._all_source_downloads_score_cache = {
                "key": cache_key,
                "rows": [dict(item) for item in scored_downloads],
            }

        voice_rows = self._download_target_rows()
        voice_inventory = self._collect_voice_mods_inventory(show_message=False)

        table = QTableWidget(dialog)
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Base Mods", "Score", "Target online to download", "State"])
        table.setRowCount(len(scored_downloads))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        table._ll_voice_candidates = scored_downloads

        for row_index, download in enumerate(scored_downloads):
            score = int(download.get("online_score") or 0)

            values = [
                download.get("target_base_mod") or "",
                str(score or ""),
                download.get("download_name") or "",
                download.get("already_status") or "Not found",
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(self._source_download_tooltip(download))
                item.setBackground(self._candidate_background_color(score))
                item.setForeground(QColor(242, 242, 242))
                table.setItem(row_index, column, item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)

        if scored_downloads:
            table.selectRow(0)

        table.customContextMenuRequested.connect(
            lambda pos: self._show_source_download_context_menu(table, pos)
        )

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
                    VOICE_CATEGORY_LABELS.get(
                        download.get("voice_category") or "",
                        download.get("voice_category") or "",
                    ),
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

        inventory_panel = self._build_voice_inventory_panel(dialog, voice_inventory)

        download_selected = QPushButton("Download selected for best target")
        close = QPushButton("Close")
        close.clicked.connect(dialog.reject)

        def start_download() -> None:
            candidate = self._candidate_from_online_table(table)
            if not candidate:
                QMessageBox.information(dialog, "LL Integration", "Select a download first.")
                return

            target = self._target_row_for_download(voice_rows, candidate)
            if not target:
                QMessageBox.information(
                    dialog,
                    "LL Integration",
                    "No target mod was found for this download.\n\n"
                    "The selected row has no best target. Use a higher scored candidate or improve the source/name match.",
                )
                return

            self._download_source_candidate(candidate, target, parent=dialog)

        download_selected.clicked.connect(lambda _checked=False: start_download())
        table.itemDoubleClicked.connect(lambda _item: start_download())

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Filter"))
        filter_row.addWidget(filter_text, 1)
        filter_row.addWidget(count_label)

        buttons = QHBoxLayout()
        buttons.addWidget(download_selected)
        buttons.addStretch(1)
        buttons.addWidget(close)

        while layout.count():
            child = layout.takeAt(0)
            widget = child.widget()
            if widget is not None:
                widget.setParent(None)

        title.setMaximumHeight(22)
        hint.setMaximumHeight(24)
        hint.setStyleSheet("color: #c8c8c8;")
        status_label.setMaximumHeight(22)

        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        header_row = QHBoxLayout()
        header_row.setContentsMargins(0, 0, 0, 0)
        header_row.setSpacing(8)
        header_row.addWidget(title)
        header_row.addWidget(hint, 1)

        filter_row.setContentsMargins(0, 0, 0, 0)
        filter_row.setSpacing(4)

        splitter = QSplitter(Qt.Orientation.Horizontal, dialog)
        splitter.addWidget(table)
        splitter.addWidget(inventory_panel)
        splitter.setStretchFactor(0, 4)
        splitter.setStretchFactor(1, 1)

        bottom_row = QHBoxLayout()
        bottom_row.setContentsMargins(0, 0, 0, 0)
        bottom_row.setSpacing(6)
        bottom_row.addWidget(download_selected)
        bottom_row.addWidget(status_label, 1)
        bottom_row.addWidget(close)

        layout.addLayout(header_row)
        layout.addLayout(filter_row)
        layout.addWidget(splitter, 1)
        layout.addLayout(bottom_row)

        apply_filter()
        dialog.exec()

    def _build_ui(self) -> None:
        self.summary = QLabel(self._summary_text())
        self.summary.setWordWrap(True)

        self.filter_mode = QComboBox(self)
        self.filter_mode.addItems([
            "All",
            "Missing",
            "Online found",
            "Possible",
            "Installed",
            "Complete",
            "None",
            "Ignored",
        ])
        self.filter_mode.currentTextChanged.connect(lambda _text: self._apply_filter())

        self.filter_text = QLineEdit(self)
        self.filter_text.setPlaceholderText("Search mod / voice / query")
        self.filter_text.setClearButtonEnabled(True)
        self.filter_text.textChanged.connect(lambda _text: self._apply_filter())

        self.threshold_spin = QSpinBox(self)
        self.threshold_spin.setRange(0, 100)
        self.threshold_spin.setValue(70)
        self.threshold_spin.setToolTip("Minimum local score required before an installed voice candidate is shown as a match.")

        self.filter_count = QLabel("0 / 0", self)
        self.selected_label = QLabel("Selected: none")
        self.selected_label.setWordWrap(True)

        self.table = QTableWidget(self)
        self.table.setColumnCount(12)
        self.table.setHorizontalHeaderLabels([
            "Status",
            "Base LoversLab mod",
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
        self.table.setRowCount(0)
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.itemSelectionChanged.connect(self._update_selected_label)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(lambda pos: self._show_voice_context_menu(pos))
        self.table.itemDoubleClicked.connect(lambda _item: self._show_installed_voice_candidates())
        self.table.horizontalHeader().sectionClicked.connect(lambda column: self._sort_voice_table(column))

        self._set_column_widths()
        self._populate_table()

        self.progress_label = QLabel("Ready")

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
        refresh = QPushButton("Refresh")
        close = QPushButton("Close")

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

        open_page.clicked.connect(lambda _checked=False: self._open_selected_page())
        refresh.clicked.connect(lambda _checked=False: self._refresh())
        close.clicked.connect(self.accept)

        source_urls.clicked.connect(lambda _checked=False: self._edit_source_urls())
        nexus_link.clicked.connect(lambda _checked=False: self._handle_nexus_link())
        fetch_sources.clicked.connect(lambda _checked=False: self._fetch_sources(fetch_sources, source_urls, download_candidate))
        all_downloads.clicked.connect(lambda _checked=False: self._show_all_source_downloads())
        download_candidate.clicked.connect(lambda _checked=False: self._download_selected_voice_candidate())

        false_match.clicked.connect(lambda _checked=False: self._mark_false_match())
        complete_slot.clicked.connect(lambda _checked=False: self._toggle_complete_selected())
        ignore_mod.clicked.connect(lambda _checked=False: self._toggle_ignore_selected())
        classify_mod.clicked.connect(lambda _checked=False: self._show_classify_dialog())
        voice_mods.clicked.connect(lambda _checked=False: self._show_installed_voice_candidates())
        search_web.clicked.connect(lambda _checked=False: self._search_context_voice())
        manage_false_matches.clicked.connect(lambda _checked=False: self._show_false_matches_dialog())

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Filter"))
        controls.addWidget(self.filter_mode)
        controls.addWidget(self.filter_text, 1)
        controls.addWidget(QLabel("Local score"))
        controls.addWidget(self.threshold_spin)
        controls.addWidget(self.filter_count)

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
        row_buttons.addWidget(refresh)
        row_buttons.addStretch(1)
        row_buttons.addWidget(close)

        # Boutons ambigus: taggage seulement par right-click sur Voicepack / DBVO / IDTV.
        false_match.hide()
        complete_slot.hide()
        ignore_mod.hide()
        classify_mod.hide()

        self.summary.setMaximumHeight(42)
        self.progress_label.setMaximumHeight(22)
        self.selected_label.setMaximumHeight(26)

        controls.setContentsMargins(0, 0, 0, 0)
        controls.setSpacing(4)

        source_buttons.setContentsMargins(0, 0, 0, 0)
        source_buttons.setSpacing(4)

        row_buttons.setContentsMargins(0, 0, 0, 0)
        row_buttons.setSpacing(4)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)

        layout.addWidget(self.summary)
        layout.addLayout(controls)
        layout.addWidget(self.selected_label)
        layout.addWidget(self.table, 1)
        layout.addWidget(self.progress_label)
        layout.addLayout(source_buttons)
        layout.addLayout(row_buttons)

        self._apply_filter()

    def _populate_table(self) -> None:
        self._refresh_display_rows()

        if self.table.rowCount() > 0 and not self.table.selectedItems():
            self.table.selectRow(0)

        if hasattr(self, "status_label"):
            self._update_summary()

        self._update_selected_label()

    def _update_summary(self) -> None:
        if hasattr(self, "summary"):
            self.summary.setText(self._summary_text())

    def _set_item(self, row: int, column: int, value: str) -> None:
        item = QTableWidgetItem(str(value or ""))
        item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)

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
            elif status == "Online found":
                item.setToolTip("Online found: a source download matched this base mod.")
            elif status == "Complete":
                item.setToolTip("This voice slot is marked as done.")
            elif status == "None":
                item.setToolTip("None: this voice slot is intentionally marked as not existing for this mod.")
            elif status == "Ignored":
                item.setToolTip("Ignored: this base mod is hidden from the missing queue.")

        elif column in (self.COL_VOICEPACK, self.COL_DBVO, self.COL_IVDT):
            item.setToolTip(str(value or "Missing"))

        elif column == self.COL_CATEGORY:
            item.setToolTip("Voice slot: Voicepack, DBVO/DVO, or IVDT/DVIT.")

        elif column == self.COL_INSTALLED_VOICE:
            item.setToolTip(
                f"{full_text}\n\nDouble-click or use Voice mods to inspect installed candidates.".strip()
            )

        elif column == self.COL_ONLINE:
            item.setToolTip(
                f"{full_text}\n\nUse Choose / Download or Show All Source Downloads to inspect online candidates.".strip()
            )

        self.table.setItem(row, column, item)

    def _summary_text(self) -> str:
        total = len(self.rows)
        missing = sum(1 for row in self.rows if row.get("status") == "Missing")
        return (
            f"Vortex voice finder preview. Base LoversLab mods: {total}. "
            f"Missing voice slots: {missing}. "
            "This first Vortex pass is UI-only; matching and Nexus/API fetch will be plugged next."
        )

    def _fill_table_row(self, row_index: int, row: dict) -> None:
        overview = row.get("voice_overview") or {}
        values = [
            row.get("status", ""),
            row.get("base_mod", ""),
            overview.get("npc", row.get("voicepack", "")),
            overview.get("player", row.get("dbvo", "")),
            overview.get("scene", row.get("IVDT", "")),
            row.get("voice_category_label") or row.get("voice_category", ""),
            row.get("installed_voice", ""),
            "Manual" if row.get("manual_voice") else str(row.get("score") or ""),
            row.get("online_candidate", ""),
            str(row.get("online_score") or ""),
            row.get("online_source_url") or row.get("source", ""),
            row.get("base_page_url", ""),
        ]

        for column, value in enumerate(values):
            item = QTableWidgetItem(str(value or ""))
            item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            item.setToolTip(str(value or ""))

            if column == VOICE_COL_INSTALLED_VOICE:
                item.setToolTip(f"{value}\n\nDouble-click or use Voice mods to inspect installed candidates.".strip())
            elif column in (VOICE_COL_VOICEPACK, VOICE_COL_DBVO, VOICE_COL_IVDT):
                item.setToolTip(str(value or "Missing"))
            elif column == VOICE_COL_CATEGORY:
                item.setToolTip("Voice slot: Voicepack, DBVO, or IVDT.")

            item.setForeground(QColor(242, 242, 242))
            self.table.setItem(row_index, column, item)

        self._apply_row_background(row_index, row)

    def _set_column_widths(self) -> None:
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(VOICE_COL_STATUS, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(VOICE_COL_BASE_MOD, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_VOICEPACK, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_DBVO, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_IVDT, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_CATEGORY, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(VOICE_COL_INSTALLED_VOICE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_SCORE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(VOICE_COL_ONLINE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_ONLINE_SCORE, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(VOICE_COL_SOURCE, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(VOICE_COL_PAGE, QHeaderView.ResizeMode.Stretch)

    def _apply_row_background(self, row_index: int, row: dict) -> None:
        status = str(row.get("status") or "")
        if status == "Installed":
            color = QColor(34, 78, 52)
        elif status == "Possible":
            color = QColor(88, 70, 30)
        elif status == "Online found":
            color = QColor(34, 58, 86)
        elif status == "Complete":
            color = QColor(34, 68, 88)
        elif status == "None":
            color = QColor(58, 58, 68)
        elif status == "Ignored":
            color = QColor(54, 54, 54)
        elif status == "Missing":
            color = QColor(78, 42, 42)
        else:
            color = QColor(38, 42, 46)

        for column in range(self.table.columnCount()):
            item = self.table.item(row_index, column)
            if item is not None:
                item.setBackground(color)
                item.setForeground(QColor(242, 242, 242))

    def _selected_row_index(self) -> int:
        selected = self.table.selectedItems()
        return selected[0].row() if selected else -1

    def _selected_row(self) -> dict | None:
        row_index = self._selected_row_index()
        if row_index < 0:
            return None

        display_rows = getattr(self.table, "_ll_voice_display_rows", [])
        if row_index < 0 or row_index >= len(display_rows):
            return None

        display_row = display_rows[row_index]
        slots = display_row.get("slots") or {}

        return slots.get("npc") or slots.get("player") or slots.get("scene")

    def _update_selected_label(self) -> None:
        row = self._selected_row()
        if not row:
            self.selected_label.setText("Selected: none")
            return

        candidate = row.get("online_candidate") or row.get("installed_voice") or "no candidate"
        self.selected_label.setText(
            f"Selected: {row.get('base_mod') or ''} | "
            f"{row.get('voice_category') or ''} | "
            f"{row.get('status') or ''} | {candidate}"
        )

    def _apply_filter(self) -> None:
        mode = self.filter_mode.currentText()
        needle = self.filter_text.text().strip().lower()
        visible = 0
        first_visible = -1

        display_rows = getattr(self.table, "_ll_voice_display_rows", [])

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

            self.table.setRowHidden(row_index, not show)

            if show:
                visible += 1
                if first_visible < 0:
                    first_visible = row_index

        self.filter_count.setText(f"{visible} / {len(display_rows)}")

        selected = self.table.selectedItems()
        if first_visible >= 0 and (not selected or self.table.isRowHidden(selected[0].row())):
            self.table.selectRow(first_visible)

    def _refresh(self) -> None:
        self._refresh_voice_rows_preserve_fetch()

    def _schedule_vortex_refresh(self) -> None:
        QTimer.singleShot(1200, self._refresh)

    def _sort_rows(self, column: int, ascending: bool, apply_filter: bool = True) -> None:
        status_order = {
            "Missing": 0,
            "Online found": 1,
            "Possible": 2,
            "Installed": 3,
            "Complete": 4,
            "None": 5,
            "Ignored": 6,
        }
        category_order = {category: index for index, (category, _label) in enumerate(VOICE_CATEGORIES)}

        def column_value(row: dict):
            if column == VOICE_COL_STATUS:
                return status_order.get(str(row.get("status") or ""), 99)
            if column == VOICE_COL_BASE_MOD:
                return str(row.get("base_mod") or "").lower()
            if column == VOICE_COL_VOICEPACK:
                return str((row.get("voice_overview") or {}).get("npc") or row.get("voicepack") or "").lower()
            if column == VOICE_COL_DBVO:
                return str((row.get("voice_overview") or {}).get("player") or row.get("dbvo") or "").lower()
            if column == VOICE_COL_IVDT:
                return str((row.get("voice_overview") or {}).get("scene") or row.get("IVDT") or "").lower()
            if column == VOICE_COL_CATEGORY:
                return category_order.get(str(row.get("voice_category") or ""), 99)
            if column == VOICE_COL_INSTALLED_VOICE:
                return str(row.get("installed_voice") or "").lower()
            if column == VOICE_COL_SCORE:
                return int(row.get("score") or 0)
            if column == VOICE_COL_ONLINE:
                return str(row.get("online_candidate") or "").lower()
            if column == VOICE_COL_ONLINE_SCORE:
                return int(row.get("online_score") or 0)
            if column == VOICE_COL_SOURCE:
                return str(row.get("online_source_url") or row.get("source") or "").lower()
            if column == VOICE_COL_PAGE:
                return str(row.get("base_page_url") or "").lower()
            return ""

        self.rows.sort(
            key=lambda row: (
                column_value(row),
                str(row.get("base_mod") or "").lower(),
                category_order.get(str(row.get("voice_category") or ""), 99),
            ),
            reverse=not ascending,
        )

        if hasattr(self, "table"):
            self.table.setRowCount(len(self.rows))
            for row_index, row in enumerate(self.rows):
                self._fill_table_row(row_index, row)
            if apply_filter:
                self._apply_filter()

    def _sort_voice_table(self, column: int) -> None:
        old_column = getattr(self, "_ll_voice_sort_column", -1)
        old_ascending = bool(getattr(self, "_ll_voice_sort_ascending", True))
        ascending = not old_ascending if old_column == column else True
        self._ll_voice_sort_column = column
        self._ll_voice_sort_ascending = ascending
        self._sort_rows(column, ascending, apply_filter=True)

    def _category_for_overview_column(self, column: int) -> str:
        if column == VOICE_COL_VOICEPACK:
            return "npc"
        if column == VOICE_COL_DBVO:
            return "player"
        if column == VOICE_COL_IVDT:
            return "scene"
        return ""

    def _context_target_row(self, row_index: int, column: int) -> dict | None:
        display_rows = getattr(self.table, "_ll_voice_display_rows", [])
        if row_index < 0 or row_index >= len(display_rows):
            return None

        display_row = display_rows[row_index]
        slots = display_row.get("slots") or {}

        category = self._category_for_overview_column(column)
        if category:
            return slots.get(category)

        return None

    def _selected_context_row(self) -> dict | None:
        context_row = getattr(self.table, "_ll_context_voice_row", None)
        if isinstance(context_row, dict):
            self.table._ll_context_voice_row = None
            return context_row
        return self._selected_row()

    def _set_slot_flag(self, row: dict, mode: str) -> None:
        slot_key = str(row.get("slot_key") or self._slot_key(row, str(row.get("voice_category") or "player"))).lower()
        if not slot_key:
            QMessageBox.information(self, "LL Integration", "Selected voice slot cannot be updated.")
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
                self,
                "LL Integration",
                "Mark this voice slot as None / does not exist?\n\n"
                f"Base: {row.get('base_mod') or ''}\n"
                f"Slot: {row.get('voice_category_label') or row.get('voice_category') or ''}\n\n"
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

        self._refresh_voice_base(row)
        self._apply_filter()
        self._update_selected_label()

    def _row_identity(self, row: dict) -> tuple[str, str]:
        return (
            str(row.get("base_internal_name") or row.get("base_mod") or "").lower(),
            str(row.get("voice_category") or "").lower(),
        )

    def _candidate_identity(self, candidate: dict) -> tuple[str, str, str, str]:
        return (
            str(candidate.get("base_internal_name") or candidate.get("base_mod") or "").lower(),
            str(candidate.get("voice_category") or "").lower(),
            str(candidate.get("download_name") or candidate.get("file_name") or "").lower(),
            str(candidate.get("source_url") or "").lower(),
        )

    def _candidate_is_false_match(self, config: dict, row: dict, candidate: dict) -> bool:
        base_key = self._base_key(row)
        category_key = str(row.get("voice_category") or "").lower()
        candidate_key = str(candidate.get("download_name") or candidate.get("file_name") or "").lower()
        source_key = str(candidate.get("source_url") or "").lower()

        for item in config.get("falseMatches", []):
            if str(item.get("base") or "").lower() != base_key:
                continue

            item_category = str(item.get("voice_category") or "").lower()
            if item_category and item_category != category_key:
                continue

            if str(item.get("candidate") or "").lower() != candidate_key:
                continue

            item_source = str(item.get("source_url") or "").lower()
            if item_source and item_source != source_key:
                continue

            return True

        return False

    def _apply_online_candidates_to_rows(self) -> None:
        config = self._load_voice_config()
        candidates_by_slot: dict[tuple[str, str], list[dict]] = {}

        for candidate in list(self.online_candidates or []):
            candidate = dict(candidate)
            base = str(candidate.get("base_internal_name") or candidate.get("base_mod") or "").lower()
            category = str(candidate.get("voice_category") or "").lower()
            if not base or not category:
                continue
            candidates_by_slot.setdefault((base, category), []).append(candidate)

        for row in self.rows:
            key = self._row_identity(row)
            row_candidates = candidates_by_slot.get(key, [])
            row_candidates = [
                candidate
                for candidate in row_candidates
                if not self._candidate_is_false_match(config, row, candidate)
            ]

            if not row_candidates:
                row["online_candidate"] = ""
                row["online_download_url"] = ""
                row["online_source_url"] = ""
                row["online_score"] = 0
                row["online_size"] = ""
                row["online_date_iso"] = ""
                row["online_version"] = ""
                row["online_candidates"] = []
                continue

            row_candidates = sorted(
                row_candidates,
                key=lambda item: int(item.get("online_score") or 0),
                reverse=True,
            )
            candidate = row_candidates[0]

            row.update({
                "online_candidate": candidate.get("download_name") or candidate.get("file_name") or "",
                "online_download_url": candidate.get("download_url") or "",
                "online_source_url": candidate.get("source_url") or "",
                "online_score": candidate.get("online_score") or 0,
                "online_size": candidate.get("size") or "",
                "online_date_iso": candidate.get("date_iso") or "",
                "online_version": candidate.get("version") or "",
                "online_candidates": row_candidates,
                "source_type": candidate.get("source_type") or "",
            })

            if (row.get("slot_status") or row.get("status")) == "Missing":
                row["slot_status"] = "Online found"
                row["status"] = "Online found"

    def _refresh_voice_rows_preserve_fetch(self, select_key: tuple[str, str] | None = None) -> None:
        old_context = getattr(self.table, "_ll_context_voice_row", None)
        if select_key is None and isinstance(old_context, dict):
            select_key = self._row_identity(old_context)

        self.config = read_json(CONFIG_FILE)
        self.vortex_state = load_vortex_state(self.config)

        state_downloads = str(self.vortex_state.get("downloadsPath") or "").strip()
        state_staging = str(self.vortex_state.get("stagingPath") or "").strip()
        if state_downloads:
            self.downloads_path = Path(state_downloads)
        if state_staging:
            self.mods_path = Path(state_staging)

        self.rows = self._build_initial_rows()
        self._apply_online_candidates_to_rows()
        self._attach_voice_overview(self.rows)
        self._populate_table()
        self._apply_filter()
        self._update_summary()
        self._update_selected_label()

        if select_key:
            display_rows = getattr(self.table, "_ll_voice_display_rows", [])
            for display_index, display_row in enumerate(display_rows):
                slots = display_row.get("slots") or {}
                if any(self._row_identity(slot) == select_key for slot in slots.values()):
                    self.table.selectRow(display_index)
                    break

        self.table._ll_context_voice_row = None

    def _show_voice_context_menu(self, pos) -> None:
        item = self.table.itemAt(pos)
        if item is None:
            return

        self.table.setCurrentCell(item.row(), item.column())

        display_rows = getattr(self.table, "_ll_voice_display_rows", [])
        if item.row() < 0 or item.row() >= len(display_rows):
            return

        display_row = display_rows[item.row()]
        slots = display_row.get("slots") or {}

        slot_row = self._context_target_row(item.row(), item.column())
        base_row = slot_row or slots.get("npc") or slots.get("player") or slots.get("scene")

        if not base_row:
            return

        # Important: les vieux boutons / helpers utilisent _selected_context_row().
        self.table._ll_context_voice_row = base_row

        menu = QMenu(self.table)

        title = QAction(
            f"{display_row.get('base_mod') or base_row.get('base_mod') or ''}",
            menu,
        )
        title.setEnabled(False)
        menu.addAction(title)
        menu.addSeparator()

        open_page_action = QAction("Open LL page", menu)
        open_page_action.triggered.connect(lambda _checked=False: self._open_selected_page())
        menu.addAction(open_page_action)

        web_search_action = QAction("Web search voice sources", menu)
        web_search_action.triggered.connect(lambda _checked=False: self._search_context_voice(base_row))
        menu.addAction(web_search_action)

        inspect_action = QAction("Inspect installed candidates", menu)
        inspect_action.triggered.connect(lambda _checked=False: self._show_installed_voice_candidates(base_row))
        menu.addAction(inspect_action)

        menu.addSeparator()

        # Actions seulement valides quand on right-click directement sur Voicepack / DBVO / IVDT.
        if slot_row:
            slot_title = QAction(
                f"Slot: {slot_row.get('voice_category_label') or slot_row.get('voice_category') or ''}",
                menu,
            )
            slot_title.setEnabled(False)
            menu.addAction(slot_title)

            complete_action = QAction("Mark this slot complete", menu)
            none_action = QAction("Set this slot to None / does not exist", menu)
            reopen_action = QAction("Reopen this slot", menu)

            complete_action.triggered.connect(lambda _checked=False: self._set_slot_flag(slot_row, "complete"))
            none_action.triggered.connect(lambda _checked=False: self._set_slot_flag(slot_row, "none"))
            reopen_action.triggered.connect(lambda _checked=False: self._set_slot_flag(slot_row, "reopen"))

            menu.addAction(complete_action)
            menu.addAction(none_action)
            menu.addAction(reopen_action)

            if str(slot_row.get("installed_voice") or "").strip():
                false_match_action = QAction("Reject this installed candidate / false local match", menu)
                false_match_action.triggered.connect(lambda _checked=False: self._mark_false_match(slot_row))
                menu.addAction(false_match_action)

            menu.addSeparator()

        ignore_action = QAction("Ignore / Unignore this base mod", menu)
        ignore_action.triggered.connect(lambda _checked=False: self._toggle_ignore_selected())
        menu.addAction(ignore_action)

        classify_action = QAction("Classify this mod", menu)
        classify_action.triggered.connect(lambda _checked=False: self._show_classify_dialog())
        menu.addAction(classify_action)

        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _show_installed_voice_candidates(self, row: dict | None = None) -> None:
        row = row or self._selected_context_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return

        candidates = sorted(
            list(row.get("installed_voice_candidates") or []),
            key=lambda item: (-int(item.get("score") or 0), str(item.get("display_name") or "").lower()),
        )
        if not candidates:
            QMessageBox.information(self, "LL Integration", "No installed voice candidates were found for this slot.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle(f"Installed voice candidates - {row.get('base_mod') or ''}")
        dialog.resize(900, 500)
        dialog.setStyleSheet(DARK_STYLE)

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
                item.setBackground(self._candidate_background_color(int(candidate.get("score") or 0), bool(candidate.get("manual"))))
                item.setForeground(QColor(242, 242, 242))
                candidate_table.setItem(index, column, item)

        header = candidate_table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        if candidates:
            candidate_table.selectRow(0)

        title = QLabel(f"Base mod: {row.get('base_mod') or ''} | {row.get('voice_category_label') or ''}")
        title.setStyleSheet("font-weight: 700;")
        hint = QLabel("Installed candidates are sorted by score. Fix selected binds that candidate to this exact slot.")
        hint.setWordWrap(True)

        open_folder = QPushButton("Open selected folder")
        fix_manual = QPushButton("Fix manual selected")
        close = QPushButton("Close")
        close.clicked.connect(dialog.reject)

        def selected_candidate() -> dict | None:
            selected = candidate_table.selectedItems()
            if not selected:
                return None
            index = selected[0].row()
            if index < 0 or index >= len(candidates):
                return None
            return candidates[index]

        def open_selected_folder() -> None:
            candidate = selected_candidate()
            if not candidate:
                return
            path = Path(str(candidate.get("mod_path") or ""))
            if not path.exists():
                QMessageBox.information(dialog, "LL Integration", f"Folder not found:\n{path}")
                return
            webbrowser.open(path.as_uri() if path.is_absolute() else str(path))

        def fix_selected() -> None:
            candidate = selected_candidate()
            if not candidate:
                QMessageBox.information(dialog, "LL Integration", "Select a voice candidate first.")
                return

            slot_key = str(row.get("slot_key") or self._slot_key(row, str(row.get("voice_category") or "player"))).lower()
            if not slot_key or not candidate.get("internal_name"):
                QMessageBox.information(dialog, "LL Integration", "This candidate cannot be bound manually.")
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

            candidate = dict(candidate)
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
                    if str(item.get("internal_name") or "").lower() != str(candidate.get("internal_name") or "").lower()
                ],
            ]
            self._refresh_voice_base(row)
            self._apply_filter()
            self._update_selected_label()
            QMessageBox.information(dialog, "LL Integration", "Manual voice match saved.")

        open_folder.clicked.connect(lambda _checked=False: open_selected_folder())
        candidate_table.itemDoubleClicked.connect(lambda _item: open_selected_folder())
        fix_manual.clicked.connect(lambda _checked=False: fix_selected())

        buttons = QHBoxLayout()
        buttons.addWidget(open_folder)
        buttons.addWidget(fix_manual)
        buttons.addStretch(1)
        buttons.addWidget(close)

        layout = QVBoxLayout(dialog)
        layout.addWidget(title)
        layout.addWidget(hint)
        layout.addWidget(candidate_table)
        layout.addLayout(buttons)
        dialog.exec()

    def _mark_false_match(self, row: dict | None = None) -> None:
        row = row or self._selected_context_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return
        candidate = str(row.get("installed_voice") or "").strip()
        if not candidate:
            QMessageBox.information(self, "LL Integration", "This slot has no local candidate to mark false.")
            return

        confirm = QMessageBox.question(
            self,
            "LL Integration",
            "Mark this local candidate as a false match?\n\n"
            f"Base: {row.get('base_mod') or ''}\n"
            f"Slot: {row.get('voice_category_label') or ''}\n"
            f"Candidate: {candidate}",
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        config = self._load_voice_config()
        false_matches = list(config.get("falseMatches") or [])
        entry = {
            "base": self._base_key(row),
            "candidate": candidate,
            "source_url": "",
            "voice_category": row.get("voice_category") or "",
        }
        if entry not in false_matches:
            false_matches.append(entry)
        config["falseMatches"] = false_matches

        slot_key = str(row.get("slot_key") or "").lower()
        manual_matches = dict(config.get("manualVoiceMatches") or {})
        manual_matches.pop(slot_key, None)
        config["manualVoiceMatches"] = manual_matches
        self._save_voice_config(config)
        self._refresh_voice_rows_preserve_fetch(self._row_identity(row))
        self.progress_label.setText("False local match saved.")

    def _toggle_ignore_selected(self) -> None:
        row = self._selected_context_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return

        base_key = self._base_key(row)
        config = self._load_voice_config()
        ignored = {str(value).lower() for value in config.get("ignoredBaseMods", [])}
        if base_key in ignored:
            ignored.remove(base_key)
            action = "unignored"
        else:
            ignored.add(base_key)
            action = "ignored"

        config["ignoredBaseMods"] = sorted(ignored)
        self._save_voice_config(config)
        self._refresh_voice_rows_preserve_fetch(self._row_identity(row))
        self.progress_label.setText(f"Base mod {action}.")

    def _toggle_complete_selected(self) -> None:
        row = self._selected_context_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return
        current = str(row.get("slot_status") or row.get("status") or "")
        self._set_slot_flag(row, "reopen" if current in ("Complete", "None") else "complete")

    def _show_classify_dialog(self) -> None:
        row = self._selected_context_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return

        base_key = self._base_key(row)
        config = self._load_voice_config()
        forced_voice = {str(value).lower() for value in config.get("forcedVoiceMods", [])}
        forced_base = {str(value).lower() for value in config.get("forcedBaseMods", [])}

        menu = QMenu(self)
        voice_action = QAction("Force selected mod as voice mod", menu)
        base_action = QAction("Force selected mod as base mod", menu)
        auto_action = QAction("Clear classification override", menu)
        voice_action.triggered.connect(lambda _checked=False: self._set_classification_override(base_key, "voice", forced_voice, forced_base, config))
        base_action.triggered.connect(lambda _checked=False: self._set_classification_override(base_key, "base", forced_voice, forced_base, config))
        auto_action.triggered.connect(lambda _checked=False: self._set_classification_override(base_key, "auto", forced_voice, forced_base, config))
        menu.addAction(voice_action)
        menu.addAction(base_action)
        menu.addAction(auto_action)
        menu.exec(self.mapToGlobal(self.rect().center()))

    def _set_classification_override(self, base_key: str, mode: str, forced_voice: set[str], forced_base: set[str], config: dict) -> None:
        forced_voice.discard(base_key)
        forced_base.discard(base_key)
        if mode == "voice":
            forced_voice.add(base_key)
        elif mode == "base":
            forced_base.add(base_key)
        config["forcedVoiceMods"] = sorted(forced_voice)
        config["forcedBaseMods"] = sorted(forced_base)
        self._save_voice_config(config)
        self._refresh_voice_rows_preserve_fetch()
        self.progress_label.setText(f"Classification set to {mode}.")

    def _search_context_voice(self, row: dict | None = None) -> None:
        row = row or self._selected_context_row() or self._selected_row()
        if not row:
            QMessageBox.information(self, "LL Integration", "Select a row first.")
            return

        search_text = (
            str(row.get("online_candidate") or "").strip()
            or str(row.get("file_name") or "").strip()
            or str(row.get("archive") or "").strip()
            or str(row.get("base_mod") or "").strip()
            or str(row.get("search_query") or "").strip()
        )

        if not search_text:
            QMessageBox.information(self, "LL Integration", "This row has no filename/search value.")
            return

        search_text = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", search_text, flags=re.IGNORECASE)
        search_text = re.sub(r"[_\-\.]+", " ", search_text)
        search_text = re.sub(r"\s+", " ", search_text).strip()

        source_url = (
            str(row.get("online_source_url") or "").strip()
            or str(row.get("source_url") or "").strip()
            or str(row.get("base_page_url") or "").strip()
            or str(row.get("page_url") or "").strip()
        )

        domain = ""
        if source_url:
            parsed = urlparse(source_url if re.match(r"^[a-z][a-z0-9+.-]*://", source_url, re.I) else "https://" + source_url)
            domain = (parsed.netloc or "").lower()
            if domain.startswith("www."):
                domain = domain[4:]

        google_query = f"site:{domain} {search_text}" if domain else search_text
        url = "https://www.google.com/search?q=" + quote_plus(google_query)
        webbrowser.open(url)

    def _show_false_matches_dialog(self) -> None:
        config = self._load_voice_config()
        false_matches = list(config.get("falseMatches") or [])

        dialog = QDialog(self)
        dialog.setWindowTitle("LL Integration - False Matches")
        dialog.resize(900, 460)
        dialog.setStyleSheet(DARK_STYLE)

        info = QLabel(
            "These entries block local voice candidates from being auto-matched again. "
            "Remove an entry if you want LL Integration to consider that candidate again."
        )
        info.setWordWrap(True)

        table = QTableWidget(dialog)
        table.setColumnCount(4)
        table.setHorizontalHeaderLabels(["Base", "Voice type", "Candidate", "Source URL"])
        table.setRowCount(len(false_matches))
        table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        table.setAlternatingRowColors(True)
        table.setWordWrap(False)
        table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        for row_index, entry in enumerate(false_matches):
            values = [
                entry.get("base") or "",
                VOICE_CATEGORY_LABELS.get(entry.get("voice_category") or "", entry.get("voice_category") or ""),
                entry.get("candidate") or "",
                entry.get("source_url") or "",
            ]

            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(str(value or ""))
                table.setItem(row_index, column, item)

        header = table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.ResizeToContents)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        if false_matches:
            table.selectRow(0)

        count_label = QLabel(f"{len(false_matches)} false match(es)")

        remove_selected = QPushButton("Remove selected")
        clear_all = QPushButton("Clear all")
        close = QPushButton("Close")

        def selected_index() -> int:
            selected = table.selectedItems()
            return selected[0].row() if selected else -1

        def save_and_refresh(updated: list[dict], message: str) -> None:
            new_config = self._load_voice_config()
            new_config["falseMatches"] = updated
            self._save_voice_config(new_config)
            dialog.accept()
            self._refresh_voice_rows_preserve_fetch()
            self.progress_label.setText(message)

        def remove() -> None:
            index = selected_index()
            if index < 0 or index >= len(false_matches):
                QMessageBox.information(dialog, "LL Integration", "Select a false match first.")
                return

            removed = false_matches[index]
            updated = [
                entry
                for entry_index, entry in enumerate(false_matches)
                if entry_index != index
            ]

            save_and_refresh(
                updated,
                f"Removed false match: {removed.get('candidate') or ''}",
            )

        def clear() -> None:
            if not false_matches:
                return

            confirm = QMessageBox.question(
                dialog,
                "LL Integration",
                f"Clear all {len(false_matches)} false match(es)?",
            )
            if confirm != QMessageBox.StandardButton.Yes:
                return

            save_and_refresh([], "All false matches cleared.")

        remove_selected.clicked.connect(lambda _checked=False: remove())
        clear_all.clicked.connect(lambda _checked=False: clear())
        close.clicked.connect(dialog.reject)

        buttons = QHBoxLayout()
        buttons.addWidget(remove_selected)
        buttons.addWidget(clear_all)
        buttons.addStretch(1)
        buttons.addWidget(close)

        layout = QVBoxLayout(dialog)
        layout.addWidget(info)
        layout.addWidget(count_label)
        layout.addWidget(table, 1)
        layout.addLayout(buttons)

        dialog.exec()

    def _open_selected_page(self) -> None:
        row = self._selected_row()
        if not row:
            return
        url = row.get("base_page_url") or ""
        if not url:
            QMessageBox.information(self, "LL Integration", "Selected row has no LoversLab page URL.")
            return
        webbrowser.open(url)

    def _coming_soon(self, tool_name: str) -> None:
        QMessageBox.information(
            self,
            "LL Integration",
            f"{tool_name} will be wired in the next pass.\n\n"
            "This patch only adds the Vortex Voice Finder window and MO2-like layout.",
        )

class VortexManager(QDialog):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle("LL Integration - Vortex Manager")
        self.resize(1400, 760)
        self.setMinimumSize(1180, 560)
        self.setStyleSheet(DARK_STYLE)
        self.config = read_json(CONFIG_FILE)
        self.vortex_state = load_vortex_state(self.config)
        self.downloads_path = Path(str(self.config.get("vortex_downloads_path") or ""))
        mods_text = str(self.config.get("vortex_mods_path") or self.config.get("vortex_staging_path") or "").strip()
        if not mods_text:
            mods_text = str(self.vortex_state.get("stagingPath") or "").strip()
        if not str(self.downloads_path) and self.vortex_state.get("downloadsPath"):
            self.downloads_path = Path(str(self.vortex_state.get("downloadsPath")))
        self.mods_path = Path(mods_text) if mods_text else None
        self.cookies_path = configured_cookies_path(self.config)
        self.metadata_path = Path(str(self.config.get("metadata_path") or BASE_DIR / "metadata"))
        self.rows = archive_rows(self.downloads_path, self.metadata_path, self.mods_path, self.vortex_state)
        self._build_ui()

    def _build_ui(self) -> None:
        self.filter_mode = QComboBox(self)
        self.filter_mode.addItems([
            "All links",
            "Updates",
            "OK",
            "Unknown / missing version",
            "Manual links",
            "Errors / skipped",
            "Not checked",
        ])
        self.filter_mode.currentTextChanged.connect(lambda _text: self._apply_filter())
        self.filter_text = QLineEdit(self)
        self.filter_text.setPlaceholderText("Search mod, file, or page")
        self.filter_text.setClearButtonEnabled(True)
        self.filter_text.textChanged.connect(lambda _text: self._apply_filter())
        self.filter_count = QLabel("", self)

        self.table = QTableWidget(self)
        self.table.setColumnCount(11)
        self.table.setHorizontalHeaderLabels([
            "Mod",
            "Status",
            "Current",
            "Latest",
            "File",
            "Info",
            "Fetch",
            "Page",
            "Folder",
            "Edit",
            "Purge",
        ])
        self.table.setRowCount(len(self.rows))
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setAlternatingRowColors(True)
        self.table.setWordWrap(False)
        self.table.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self._show_context_menu)

        self.progress = QProgressBar(self)
        self.progress.setRange(0, 100)
        self.progress.setValue(0)
        self.progress.setFormat("0%")
        self.status_label = QLabel()
        self.status_label.setWordWrap(True)

        self._populate_table()

        self._set_column_widths()
        if self.rows:
            self.table.selectRow(0)

        pacing = load_fetch_pacing()
        self.delay = QDoubleSpinBox(self)
        self.delay.setRange(0.0, 30.0)
        self.delay.setSingleStep(0.1)
        self.delay.setDecimals(1)
        self.delay.setSuffix(" s")
        self.delay.setValue(float(pacing["request_delay"]))
        self.cooldown_every = QSpinBox(self)
        self.cooldown_every.setRange(1, 999)
        self.cooldown_every.setValue(int(pacing["batch_size"]))
        self.cooldown_for = QDoubleSpinBox(self)
        self.cooldown_for.setRange(0.0, 120.0)
        self.cooldown_for.setSingleStep(0.5)
        self.cooldown_for.setDecimals(1)
        self.cooldown_for.setSuffix(" s")
        self.cooldown_for.setValue(float(pacing["batch_pause"]))
        self.timeout = QDoubleSpinBox(self)
        self.timeout.setRange(1.0, 120.0)
        self.timeout.setSingleStep(1.0)
        self.timeout.setDecimals(1)
        self.timeout.setSuffix(" s")
        self.timeout.setValue(float(pacing["request_timeout"]))


        refresh = QPushButton("Refresh")
        fetch_updates = QPushButton("Fetch Updates")
        create_manual = QPushButton("Create Manual Link")
        cancel = QPushButton("Cancel")
        close = QPushButton("Close")


        refresh.clicked.connect(lambda _checked=False: self._refresh())
        fetch_updates.clicked.connect(lambda _checked=False: self._fetch_updates())
        create_manual.clicked.connect(lambda _checked=False: self.create_manual_link())
        cancel.clicked.connect(lambda _checked=False: self._cancel_fetch())

        close.clicked.connect(self.accept)

        controls = QHBoxLayout()
        controls.addWidget(QLabel("Delay"))
        controls.addWidget(self.delay)
        controls.addWidget(QLabel("Cooldown every"))
        controls.addWidget(self.cooldown_every)
        controls.addWidget(QLabel("requests for"))
        controls.addWidget(self.cooldown_for)
        controls.addWidget(QLabel("Timeout / request"))
        controls.addWidget(self.timeout)
        controls.addStretch(1)

        buttons = QHBoxLayout()
        buttons.addWidget(refresh)
        buttons.addWidget(fetch_updates)
        buttons.addWidget(create_manual)
        buttons.addWidget(cancel)
        buttons.addStretch(1)
        buttons.addWidget(close)

        filters = QHBoxLayout()
        filters.addWidget(QLabel("Filter"))
        filters.addWidget(self.filter_mode)
        filters.addWidget(self.filter_text, 1)
        filters.addWidget(self.filter_count)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(4)
        layout.addLayout(filters)
        layout.addWidget(self.table, 1)
        layout.addWidget(self.progress)
        layout.addWidget(self.status_label)
        layout.addLayout(controls)
        layout.addLayout(buttons)
        self._update_summary()
        self._apply_filter()

    def _refresh_display_rows(self) -> None:
        self.table.setRowCount(len(self.rows))

    def _populate_table(self) -> None:
        self.table.setRowCount(len(self.rows))

        for row_index, row in enumerate(self.rows):
            status = self._row_status(row)
            values = [
                row.get("mod", ""),
                status,
                row.get("version", ""),
                row.get("latest", ""),
                row.get("file_name") or row.get("archive", ""),
                self._row_info(row),
            ]

            tooltip = self._row_tooltip(row)
            for column, value in enumerate(values):
                item = QTableWidgetItem(str(value or ""))
                item.setFlags(item.flags() & ~Qt.ItemFlag.ItemIsEditable)
                item.setToolTip(tooltip)
                if column == 1:
                    self._style_status_item(item, status)
                self.table.setItem(row_index, column, item)

            self._set_action_buttons(row_index, row)

        if self.table.rowCount() > 0 and not self.table.selectedItems():
            self.table.selectRow(0)

        if hasattr(self, "status_label"):
            self._update_summary()

    def _set_column_widths(self) -> None:
        header = self.table.horizontalHeader()
        header.setStretchLastSection(False)
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Fixed)
        header.setSectionResizeMode(4, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(5, QHeaderView.ResizeMode.Stretch)
        for column in ACTION_COLUMNS:
            header.setSectionResizeMode(column, QHeaderView.ResizeMode.Fixed)
        widths = {
            0: 280,
            1: 112,
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
            self.table.setColumnWidth(column, width)

    def _set_action_buttons(self, row_index: int, row: dict) -> None:
        actions = [
            (6, "Fetch", lambda _checked=False, index=row_index: self._fetch_row_index(index)),
            (7, "Open", lambda _checked=False, index=row_index: self._open_row_page(index)),
            (8, "Folder", lambda _checked=False, index=row_index: self._open_row_archive(index)),
            (9, "Edit", lambda _checked=False, index=row_index: self._edit_row_link(index)),
            (10, "Purge", lambda _checked=False, index=row_index: self._purge_row_link(index)),
        ]
        for column, label, callback in actions:
            button = QPushButton(label, self.table)
            button.clicked.connect(callback)
            self.table.setCellWidget(row_index, column, button)
        if self._row_status(row) == "Downloaded":
            self.table.cellWidget(row_index, 6).setText("Install")
        elif self._row_status(row) == "Installed":
            self.table.cellWidget(row_index, 6).setText("Enable")
        elif self._row_status(row) == "Update":
            self.table.cellWidget(row_index, 6).setText("Download")
        elif self._row_status(row) == "Update Ready":
            self.table.cellWidget(row_index, 6).setText("Install NV")
        self.table.cellWidget(row_index, 6).setEnabled(self._row_fetch_enabled(row))
        if self._row_status(row) in {"Downloaded", "Installed", "Update", "Update Ready"}:
            self.table.cellWidget(row_index, 6).setEnabled(True)
        self.table.cellWidget(row_index, 7).setEnabled(bool(row.get("page_url")))
        self.table.cellWidget(row_index, 8).setEnabled(bool(row.get("installed_folder")) and Path(str(row.get("installed_folder"))).exists())
        self.table.cellWidget(row_index, 9).setEnabled(bool(row.get("sidecar")))
        self.table.cellWidget(row_index, 10).setEnabled(bool(row.get("sidecar")))

    def purge_suspicious_links(self) -> None:
        candidates = []
        mods = list((self.vortex_state or {}).get("mods") or [])

        for mod in mods:
            meta_path = vortex_mod_meta_path(self.config, mod)
            if not vortex_has_ll_metadata(meta_path):
                continue

            purgeable, reason = vortex_has_purgeable_nexus_identity(mod, meta_path)
            if not purgeable:
                continue

            candidates.append({
                "mod": mod,
                "meta_path": meta_path,
                "reason": reason,
            })

        if not candidates:
            QMessageBox.information(self, "LL Integration", "No suspicious LoversLab links found.")
            return

        shown = candidates[:40]
        lines = [
            f"{item['mod'].get('name') or item['mod'].get('id')} ({item['reason']})"
            for item in shown
        ]
        if len(candidates) > len(shown):
            lines.append(f"...and {len(candidates) - len(shown)} more")

        confirm = QMessageBox.question(
            self,
            "LL Integration",
            "Clean suspicious LoversLab links from these Nexus-linked Vortex mods?\n\n"
            "This removes [LoversLab] metadata and stale LoversLab URL fields from meta.ini.\n\n"
            + "\n".join(lines),
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        purged = []
        for item in candidates:
            meta_path = item["meta_path"]
            if self._purge_vortex_meta_ini(meta_path):
                purged.append(str(item["mod"].get("name") or item["mod"].get("id") or meta_path.parent.name))

        QMessageBox.information(
            self,
            "LL Integration",
            "Purged suspicious LoversLab links:\n\n" + "\n".join(purged[:80]),
        )

    def _purge_vortex_meta_ini(self, meta_path: Path) -> bool:
        if not meta_path.exists():
            return False

        backup = meta_path.with_name(f"{meta_path.name}.purged-{int(time.time())}.bak")
        try:
            shutil.copy2(meta_path, backup)
        except Exception:
            pass

        parser = configparser.ConfigParser(interpolation=None)
        parser.optionxform = str
        parser.read(meta_path, encoding="utf-8-sig")

        changed = False

        if "LoversLab" in parser:
            parser.remove_section("LoversLab")
            changed = True

        if "General" in parser:
            general = parser["General"]

            for key in list(general.keys()):
                key_l = str(key).lower()
                value_l = str(general.get(key) or "").lower()

                if key_l in {"llintegration", "llsource", "llsourcename", "llvoicepack"}:
                    general.pop(key, None)
                    changed = True
                    continue

                if "loverslab.com" in value_l:
                    general[key] = ""
                    changed = True

            if changed:
                if "hasCustomURL" in general:
                    general["hasCustomURL"] = "false"

                if "repository" in general:
                    modid = (
                        general.get("modid")
                        or general.get("modId")
                        or general.get("nexusid")
                        or general.get("nexusId")
                        or ""
                    )
                    try:
                        if int(str(modid or "0")) > 0:
                            general["repository"] = "Nexus"
                    except ValueError:
                        pass

        if not changed:
            return False

        with meta_path.open("w", encoding="utf-8") as handle:
            parser.write(handle, space_around_delimiters=False)

        legacy = meta_path.parent / "LL.ini"
        try:
            if legacy.exists():
                legacy.unlink()
        except Exception:
            pass

        return True

    def _row_status(self, row: dict) -> str:
        if row.get("status_override"):
            return str(row.get("status_override"))
        if row.get("vortex_status") in {"Enabled", "Installed"}:
            return str(row.get("vortex_status"))
        if row.get("vortex_status") == "Downloaded":
            return "Downloaded"
        if row.get("has_metadata") != "Yes":
            return "Untracked"
        if row.get("fixed"):
            return "Manual"
        return "Ready"

    def _row_fetch_enabled(self, row: dict) -> bool:
        if self._row_status(row) in {"Downloaded", "Installed", "Update", "Update Ready"}:
            return True
        return (
            row.get("has_metadata") == "Yes"
            and bool(row.get("page_url"))
            and not bool(row.get("fixed"))
            and row.get("source") in ("", "loverslab")
        )

    def _style_status_item(self, item: QTableWidgetItem, status: str) -> None:
        colors = {
            "OK": QColor(44, 140, 68),
            "Update": QColor(180, 130, 0),
            "Enabled": QColor(44, 140, 68),
            "Installed": QColor(44, 140, 68),
            "Downloaded": QColor(105, 150, 190),
            "Update Ready": QColor(180, 130, 0),
            "Queued": QColor(145, 145, 145),
            "Manual": QColor(105, 150, 190),
            "Ready": QColor(145, 145, 145),
            "Untracked": QColor(145, 145, 145),
            "Unknown": QColor(145, 145, 145),
            "Error": QColor(190, 45, 45),
        }
        item.setForeground(colors.get(status, QColor(220, 220, 220)))
        item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

    def _row_info(self, row: dict) -> str:
        if row["has_metadata"] != "Yes":
            return "No LL metadata sidecar"
        bits = []
        if row.get("fixed"):
            bits.append("Updates skipped")
        elif row.get("fetch_info"):
            bits.append(row["fetch_info"])
        elif row.get("update_mode"):
            bits.append(f"{update_mode_label(row.get('update_mode') or '')}; not checked")
        if row.get("installed_folder"):
            bits.append("Installed folder found")
        elif self.mods_path:
            bits.append("Installed folder unknown")
        return "; ".join(bits)

    def _row_tooltip(self, row: dict) -> str:
        return "\n".join([
            f"Archive: {row.get('archive') or ''}",
            f"File: {row.get('file_name') or ''}",
            f"Page: {row.get('page_title') or ''}",
            f"URL: {row.get('page_url') or ''}",
            f"Archive path: {row.get('path') or ''}",
            f"Installed folder: {row.get('installed_folder') or ''}",
        ]).strip()

    def _update_summary(self) -> None:
        linked = sum(1 for row in self.rows if row["has_metadata"] == "Yes")
        missing = len(self.rows) - linked
        self.status_label.setText(
            f"Loaded {linked} LoversLab links"
            + (f" and {missing} untracked archives. " if missing else ". ")
            + "Click Fetch Updates to check versions."
        )

    def _selected_row(self) -> dict | None:
        selected = self.table.selectedItems()
        if not selected:
            return None

        row_index = selected[0].row()
        if row_index < 0 or row_index >= len(self.rows):
            return None

        return self.rows[row_index]

    def _row_for_button(self, button) -> dict | None:
        index = self.table.indexAt(button.pos())
        if not index.isValid():
            return self._selected_row()
        self.table.selectRow(index.row())
        return self.rows[index.row()]

    def _row_at(self, row_index: int) -> dict | None:
        if row_index < 0 or row_index >= len(self.rows):
            return None
        self.table.selectRow(row_index)
        return self.rows[row_index]

    def _open_row_archive(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None:
            self._open_archive(row)

    def _open_row_page(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None:
            self._open_page(row)

    def _edit_row_link(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None and self._edit_link_dialog(row):
            self._refresh()

    def _purge_row_link(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None:
            self._purge_link(row)

    def _open_selected_archive(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        self._open_archive(row)

    def _open_archive(self, row: dict) -> None:
        installed = Path(row.get("installed_folder") or "")
        if installed.exists():
            open_path(installed)
        else:
            open_path(Path(row["path"]))

    def _open_selected_page(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        self._open_page(row)

    def _open_page(self, row: dict) -> None:
        url = row.get("page_url") or ""
        if not url:
            QMessageBox.information(self, "LL Integration", "Selected archive has no source page URL.")
            return
        webbrowser.open(url)

    def _inspect_selected_metadata(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        sidecar = Path(row.get("sidecar") or "")
        if not sidecar.exists():
            QMessageBox.information(self, "LL Integration", "Selected archive has no LL metadata sidecar.")
            return
        data = load_ini(sidecar)
        text = "\n".join(f"{key}: {value}" for key, value in data.items()) or sidecar.read_text(encoding="utf-8-sig")
        QMessageBox.information(self, "LL metadata", text)

    def _edit_selected_link(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        if self._edit_link_dialog(row):
            self._refresh()

    def _purge_selected_link(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        self._purge_link(row)

    def _remove_downloaded_archive_from_manager(self) -> None:
        row = self._selected_row()
        if row is None:
            return

        archive_path = Path(row.get("path") or "")
        archive_name = row.get("archive") or archive_path.name or row.get("mod") or "(unknown)"
        status = self._row_status(row)

        if status != "Downloaded":
            QMessageBox.information(
                self,
                "LL Integration",
                "This hard-remove action is only intended for rows stuck as Downloaded.",
            )
            return

        targets: list[Path] = []

        if archive_path.exists() and archive_path.is_file():
            targets.append(archive_path)

        # Sidecar direct: archive.7z.ll.ini
        sidecar = Path(row.get("sidecar") or "")
        if sidecar.exists() and sidecar.is_file():
            targets.append(sidecar)

        # Sidecar dans metadata/downloads
        if archive_name and archive_name != "(unknown)":
            metadata_sidecar = self.metadata_path / "downloads" / f"{archive_name}.ll.ini"
            if metadata_sidecar.exists() and metadata_sidecar.is_file():
                targets.append(metadata_sidecar)

        # Nettoyage doublons éventuels autour de l'archive
        if archive_path.name:
            possible_sidecars = [
                archive_path.with_name(f"{archive_path.name}.ll.ini"),
                archive_path.with_suffix(archive_path.suffix + ".ll.ini"),
            ]
            for candidate in possible_sidecars:
                if candidate.exists() and candidate.is_file() and candidate not in targets:
                    targets.append(candidate)

        if not targets:
            QMessageBox.information(
                self,
                "LL Integration",
                "No local archive or LL sidecar was found to remove for this row.",
            )
            return

        lines = "\n".join(str(path) for path in targets[:12])
        if len(targets) > 12:
            lines += f"\n...and {len(targets) - 12} more"

        confirm = QMessageBox.warning(
            self,
            "Hard remove downloaded archive",
            "This will delete the local downloaded archive/sidecar from disk so it disappears from this manager view.\n\n"
            "It does not uninstall an installed mod.\n\n"
            f"Target row:\n{archive_name}\n\n"
            f"Files to delete:\n{lines}\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if confirm != QMessageBox.StandardButton.Yes:
            return

        deleted = []
        failed = []

        for target in targets:
            try:
                if target.exists() and target.is_file():
                    target.unlink()
                    deleted.append(str(target))
            except OSError as exc:
                failed.append(f"{target}: {exc}")

        self._refresh()

        if failed:
            QMessageBox.warning(
                self,
                "LL Integration",
                "Some files could not be deleted:\n\n" + "\n".join(failed[:20]),
            )
            return

        QMessageBox.information(
            self,
            "LL Integration",
            f"Removed downloaded archive from manager:\n\n{archive_name}",
        )

    def _purge_link(self, row: dict) -> None:
        sidecar = Path(row.get("sidecar") or "")
        if not sidecar.exists():
            QMessageBox.information(self, "LL Integration", "Selected archive has no sidecar to purge.")
            return
        if not QMessageBox.question(
            self,
            "Purge LL metadata",
            f"Delete metadata sidecar?\n\n{sidecar}",
        ) == QMessageBox.StandardButton.Yes:
            return
        try:
            sidecar.unlink()
        except OSError as exc:
            QMessageBox.warning(self, "LL Integration", f"Could not delete sidecar:\n{exc}")
            return
        self._refresh()

    def _fetch_selected(self) -> None:
        row = self._selected_row()
        if row is None:
            return
        self._fetch_row(row)

    def _fetch_row_index(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None:
            if self._row_status(row) == "Downloaded":
                self._queue_vortex_install(row)
                return
            if self._row_status(row) == "Installed":
                self._queue_vortex_enable(row)
                return
            if self._row_status(row) == "Update":
                self._download_update(row)
                return
            if self._row_status(row) == "Update Ready":
                self._queue_vortex_install(row, update_archive=True)
                return
            self._fetch_row(row)

    def _queue_vortex_install(self, row: dict, update_archive: bool = False) -> None:
        archive_path = row.get("update_archive_path") if update_archive else row.get("path")
        archive_path = archive_path or ""
        if not archive_path or not Path(archive_path).exists():
            QMessageBox.warning(self, "LL Integration", f"Archive not found:\n{archive_path}")
            return
        if update_archive and not self._confirm_update_replace(row, Path(archive_path)):
            return
        enable_after_install = normalized_update_mode(row.get("update_mode")) == UPDATE_MODE_AUTOMATIC
        if update_archive and self._row_status(row) == "Enabled":
            enable_after_install = True

        old_installed_folder = row.get("installed_folder") if update_archive else ""
        old_mod_enabled = self._row_status(row) == "Enabled"

        sidecar_data = {}
        sidecar_path = Path(row.get("sidecar") or "")

        if sidecar_path.exists():
            sidecar_data = load_ini(sidecar_path)

        source_type = str(sidecar_data.get("source") or row.get("source") or "").strip()
        page_url = str(row.get("page_url") or sidecar_data.get("page_url") or sidecar_data.get("mod_homepage") or "").strip()
        download_url = str(
            row.get("latest_url")
            if update_archive and row.get("latest_url")
            else sidecar_data.get("download_url") or ""
        ).strip()

        is_ll = "loverslab.com" in page_url.lower() or "loverslab.com" in download_url.lower()
        display_source = "LoversLab" if is_ll else (sidecar_data.get("display_source") or "Nexus Mods")

        command_id = append_vortex_command({
            "action": "install_archive",
            "operation": "replace" if update_archive else "install",
            "archivePath": archive_path,
            "gameId": self.vortex_state.get("activeGameId") or self.config.get("active_vortex_game") or "skyrimse",
            "allowAutoEnable": enable_after_install,
            "enableAfterInstall": enable_after_install,

            # Vortex replacement metadata
            "replaceModId": row.get("vortex_mod_id") if update_archive else "",
            "replaceModName": row.get("mod") if update_archive else "",
            "removeOldBeforeInstall": bool(update_archive and row.get("vortex_mod_id")),
            "deleteOldArchive": bool(update_archive),

            # Old install data
            "oldArchivePath": row.get("path") if update_archive else "",
            "oldArchiveName": row.get("archive") if update_archive else "",
            "oldDownloadId": row.get("vortex_download_id") if update_archive else "",
            "oldInstalledFolder": old_installed_folder or "",
            "oldWasEnabled": bool(old_mod_enabled),

            # Metadata for direct Vortex write.
            "sourceType": source_type or ("loverslab" if is_ll else "nexus"),
            "sourceName": "Website" if is_ll else display_source,
            "displaySource": display_source,
            "sourceUrl": page_url,
            "pageUrl": page_url,
            "modHomepage": page_url,
            "downloadUrl": download_url,
            "downloadName": row.get("latest_file_name") if update_archive else row.get("file_name") or row.get("archive") or Path(archive_path).name,
            "archiveName": Path(archive_path).name,
            "version": row.get("latest") if update_archive else row.get("version") or sidecar_data.get("version") or "",

            # Sidecars
            "oldSidecarPath": row.get("sidecar") if update_archive else "",
            "oldMetadataSidecarPath": str(self.metadata_path / "downloads" / f"{row.get('archive') or ''}.ll.ini") if update_archive else "",

            # Backup/cleanup behavior
            "backupOldModFolder": bool(update_archive),
            "deleteOldModFolder": bool(update_archive),
            "cleanupDuplicateAfterInstall": bool(update_archive),

            "profileId": self.vortex_state.get("activeProfileId") or self.config.get("active_vortex_profile_id") or "",
        })
        row["status_override"] = "Queued"
        row["fetch_info"] = f"{'New version install' if update_archive else 'Install'} queued for Vortex ({command_id}). Keep Vortex open."
        self._populate_table()
        self._apply_filter()
        self._schedule_vortex_refresh()
        QMessageBox.information(
            self,
            "LL Integration",
            "Install request queued for Vortex.\n\n"
            "Keep Vortex open; the LL Integration extension will import/start the archive install.",
        )

    def _confirm_update_replace(self, row: dict, new_archive: Path) -> bool:
        old_archive = Path(row.get("path") or "")
        old_mod = row.get("mod") or row.get("archive") or ""
        return QMessageBox.warning(
            self,
            "LL Integration",
            "LL Integration will replace this mod manually because Vortex does not merge LoversLab updates reliably.\n\n"
            "A backup of the old installed mod folder will be created first. Then the old Vortex mod entry, old staging folder, old archive, and old sidecars will be removed before installing the new archive.\n\n"
            f"Old mod: {old_mod}\n"
            f"Old archive: {old_archive.name if old_archive.name else '(unknown)'}\n"
            f"New archive: {new_archive.name}\n\n"
            "Continue?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        ) == QMessageBox.StandardButton.Yes

    def _download_update(self, row: dict) -> None:
        url = row.get("latest_url") or ""
        file_name = row.get("latest_file_name") or row.get("file_name") or ""
        if not url or not file_name:
            QMessageBox.warning(self, "LL Integration", "No update download URL is available. Fetch updates first.")
            return
        target = self.downloads_path / Path(file_name).name
        try:
            downloaded = download_ll_file(
                url,
                target,
                self.cookies_path,
                row.get("page_url") or "",
                float(self.timeout.value()),
            )
            self._write_update_sidecars(row, downloaded)
        except Exception as exc:
            row["status_override"] = "Error"
            row["fetch_info"] = f"Update download failed: {exc}"
            self._populate_table()
            self._apply_filter()
            QMessageBox.warning(self, "LL Integration", f"Update download failed:\n\n{exc}")
            return
        row["status_override"] = "Update Ready"
        row["update_archive_path"] = str(downloaded)
        row["fetch_info"] = f"Downloaded new version: {downloaded.name}"
        self._populate_table()
        self._apply_filter()

    def _write_update_sidecars(self, row: dict, downloaded: Path) -> None:
        old_sidecar = Path(row.get("sidecar") or "")
        data = load_ini(old_sidecar) if old_sidecar.exists() else {}
        updated = dict(data)
        updated.update({
            "source": "loverslab",
            "page_url": row.get("page_url") or data.get("page_url") or "",
            "download_url": row.get("latest_url") or data.get("download_url") or "",
            "file_pattern": data.get("file_pattern") or row.get("file_name") or row.get("archive") or downloaded.name,
            "file_name": row.get("latest_file_name") or downloaded.name,
            "archive_name": downloaded.name,
            "original_archive_name": data.get("original_archive_name") or row.get("archive") or downloaded.name,
            "version": row.get("latest") or data.get("version") or extract_version(downloaded.name),
            "update_mode": normalized_update_mode(row.get("update_mode") or data.get("update_mode")),
            "fixed_version": "false",
            "manual_update": "false",
            "skip_update_check": "false",
            "replaces_archive": row.get("archive") or "",
        })
        write_ini(downloaded.with_name(f"{downloaded.name}.ll.ini"), updated)
        write_ini(self.metadata_path / "downloads" / f"{downloaded.name}.ll.ini", updated)

    def _mark_update_replacement(self, row: dict, new_archive_name: str) -> None:
        for sidecar in [
            Path(row.get("sidecar") or ""),
            self.metadata_path / "downloads" / f"{row.get('archive') or ''}.ll.ini",
        ]:
            if not str(sidecar) or not sidecar.exists():
                continue
            data = load_ini(sidecar)
            if not data:
                continue
            data["replaced_by"] = new_archive_name
            write_ini(sidecar, data)

    def _cancel_update_index(self, row_index: int) -> None:
        row = self._row_at(row_index)
        if row is not None:
            self._cancel_update(row)

    def _cancel_update(self, row: dict) -> None:
        archive_path = Path(row.get("update_archive_path") or "")
        if archive_path.exists():
            try:
                archive_path.unlink()
            except OSError as exc:
                QMessageBox.warning(self, "LL Integration", f"Could not delete update archive:\n{exc}")
                return
        for sidecar in [
            archive_path.with_name(f"{archive_path.name}.ll.ini") if archive_path.name else Path(""),
            self.metadata_path / "downloads" / f"{archive_path.name}.ll.ini" if archive_path.name else Path(""),
        ]:
            if str(sidecar) and sidecar.exists():
                try:
                    sidecar.unlink()
                except OSError:
                    pass
        row.pop("update_archive_path", None)
        row["status_override"] = "Update"
        row["fetch_info"] = "Update download cancelled"
        self._populate_table()
        self._apply_filter()

    def _queue_vortex_enable(self, row: dict) -> None:
        mod_id = row.get("vortex_mod_id") or ""
        profile_id = self.vortex_state.get("activeProfileId") or self.config.get("active_vortex_profile_id") or ""
        if not mod_id:
            QMessageBox.warning(self, "LL Integration", "This row has no Vortex mod id yet.")
            return
        if not profile_id:
            QMessageBox.warning(self, "LL Integration", "Vortex active profile id is not available yet.")
            return
        command_id = append_vortex_command({
            "action": "enable_mod",
            "profileId": profile_id,
            "modId": mod_id,
        })
        row["status_override"] = "Queued"
        row["fetch_info"] = f"Enable queued for Vortex ({command_id}). Keep Vortex open."
        self._populate_table()
        self._apply_filter()
        self._schedule_vortex_refresh()

    def _schedule_vortex_refresh(self) -> None:
        QTimer.singleShot(2500, self._refresh)
        QTimer.singleShot(7000, self._refresh)

    def _fetch_updates(self) -> None:
        save_fetch_pacing(
            self.delay.value(),
            self.cooldown_every.value(),
            self.cooldown_for.value(),
            self.timeout.value(),
        )
        total = max(1, len(self.rows))
        checked = 0
        network_checked = 0
        for row in self.rows:
            if self._row_fetch_enabled(row):
                if network_checked > 0 and self.delay.value() > 0:
                    self.status_label.setText(f"Waiting {self.delay.value():.1f}s before next request")
                    self._sleep_with_events(float(self.delay.value()))
                network_checked += 1
                self._fetch_row(row, quiet=True)
                if (
                    self.cooldown_every.value() > 0
                    and self.cooldown_for.value() > 0
                    and network_checked % self.cooldown_every.value() == 0
                ):
                    self.status_label.setText(
                        f"Cooldown {self.cooldown_for.value():.1f}s after {network_checked} LoversLab requests"
                    )
                    self._sleep_with_events(float(self.cooldown_for.value()))
            checked += 1
            self.progress.setValue(int((checked / total) * 100))
            self.progress.setFormat(f"{int((checked / total) * 100)}%")
            QApplication.processEvents()
        self._populate_table()
        self._apply_filter()
        self.status_label.setText(f"Fetch complete. Checked {checked} archives.")

    def _sleep_with_events(self, seconds: float) -> None:
        import time

        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            QApplication.processEvents()
            time.sleep(min(0.1, max(deadline - time.monotonic(), 0)))

    def _cancel_fetch(self) -> None:
        self.status_label.setText("No fetch is currently running.")

    def _refresh(self) -> None:
        self.config = read_json(CONFIG_FILE)
        self.vortex_state = load_vortex_state(self.config)

        state_downloads = str(self.vortex_state.get("downloadsPath") or "").strip()
        state_staging = str(self.vortex_state.get("stagingPath") or "").strip()

        if state_downloads:
            self.downloads_path = Path(state_downloads)
        if state_staging:
            self.mods_path = Path(state_staging)

        self.metadata_path = Path(str(self.config.get("metadata_path") or BASE_DIR / "metadata"))
        self.rows = archive_rows(
            self.downloads_path,
            self.metadata_path,
            self.mods_path,
            self.vortex_state,
        )

        self._populate_table()
        self._apply_filter()

        if self.rows and not self.table.selectedItems():
            self.table.selectRow(0)

        self.status_label.setText(f"Refreshed. Loaded {len(self.rows)} archive link(s).")

    def _show_context_menu(self, pos) -> None:
        item = self.table.itemAt(pos)
        row_index = item.row() if item is not None else self.table.indexAt(pos).row()
        if row_index < 0:
            return
        self.table.selectRow(row_index)
        menu = QMenu(self)
        row = self._selected_row()

        open_archive = menu.addAction("Open archive folder")
        open_page = menu.addAction("Open source page")
        inspect = menu.addAction("Inspect metadata")
        edit = menu.addAction("Edit link")
        purge = menu.addAction("Purge metadata")

        remove_from_manager = None
        if row and self._row_status(row) == "Downloaded":
            menu.addSeparator()
            remove_from_manager = menu.addAction("Remove this downloaded archive from manager")
            remove_from_manager.setToolTip("Hard-removes the downloaded archive and LL sidecars from LL Integration's manager view.")

        menu.addSeparator()
        open_downloads = menu.addAction("Open Vortex downloads")
        open_metadata = menu.addAction("Open metadata folder")
        action = menu.exec(self.table.viewport().mapToGlobal(pos))

        if action == open_archive:
            self._open_selected_archive()
        elif action == open_page:
            self._open_selected_page()
        elif action == inspect:
            self._inspect_selected_metadata()
        elif action == edit:
            self._edit_selected_link()
        elif action == purge:
            self._purge_selected_link()
        elif remove_from_manager is not None and action == remove_from_manager:
            self._remove_downloaded_archive_from_manager()
        elif action == open_downloads:
            open_path(self.downloads_path)
        elif action == open_metadata:
            open_path(self.metadata_path)

    def _fetch_row(self, row: dict, quiet: bool = False) -> None:
        if row.get("has_metadata") != "Yes":
            if not quiet:
                QMessageBox.information(self, "LL Integration", "Selected archive has no metadata to fetch from.")
            return
        if row.get("fixed"):
            row["status_override"] = "Manual"
            row["latest"] = ""
            row["fetch_info"] = "Updates skipped"
            if not quiet:
                QMessageBox.information(self, "LL Integration", "Selected archive is marked to skip updates.")
            return
        if row.get("source") not in ("", "loverslab"):
            if not quiet:
                QMessageBox.information(self, "LL Integration", "Only LoversLab update fetch is implemented in this first Vortex pass.")
            return
        url = row.get("page_url") or ""
        if not url:
            if not quiet:
                QMessageBox.information(self, "LL Integration", "Selected archive has no page URL.")
            return
        try:
            downloads = fetch_ll_downloads(url, self.cookies_path, float(self.timeout.value()))
            latest = choose_latest(downloads, row.get("file_name") or row.get("archive") or "")
        except Exception as exc:
            row["latest"] = ""
            row["status_override"] = "Error"
            row["fetch_info"] = str(exc)
            if not quiet:
                QMessageBox.warning(self, "LL Integration", f"Fetch failed:\n\n{exc}")
            return
        if latest is None:
            row["latest"] = ""
            row["status_override"] = "Unknown"
            self.status_label.setText("No matching download found.")
            return
        row["latest"] = latest.version or ""
        current = row.get("version") or ""
        if current and latest.version and compare_versions(latest.version, current) > 0:
            row["status_override"] = "Update"
            row["latest_url"] = latest.url
            row["latest_file_name"] = latest.name
            row["file_name"] = latest.name
            row["fetch_info"] = f"Latest file: {latest.name}"
            self.status_label.setText(f"Update found: {row.get('archive')} -> {latest.name}")
        else:
            row["status_override"] = "OK"
            row["latest_url"] = latest.url
            row["latest_file_name"] = latest.name
            row["file_name"] = latest.name
            row["fetch_info"] = f"Latest file: {latest.name}"
            self.status_label.setText(f"OK: {row.get('archive')} -> {latest.name}")
        self._populate_table()
        self._apply_filter()

    def create_manual_link(self) -> None:
        try:
            mod_name, mod = self._choose_vortex_mod_filtered()
            values = self._prompt_manual_link_values(mod_name)
            if not values:
                return

            target = self._write_vortex_manual_link(mod, values)
            self._refresh()

        except Exception as exc:
            QMessageBox.critical(
                self,
                "LL Integration",
                f"Create manual link failed:\n\n{exc}",
            )
            return

        QMessageBox.information(
            self,
            "LL Integration",
            f"Created manual source link for:\n{mod_name}\n\nStored in:\n{target}",
        )

    def _choose_vortex_mod_filtered(self) -> tuple[str, dict]:
        mods = []
        staging_path = Path(str((self.vortex_state or {}).get("stagingPath") or self.mods_path or ""))

        for item in (self.vortex_state or {}).get("mods", []):
            if not isinstance(item, dict):
                continue

            mod_id = str(item.get("id") or "").strip()
            display_name = str(item.get("name") or item.get("displayName") or mod_id).strip()
            install_rel = str(item.get("installationPath") or mod_id).strip()
            mod_path = staging_path / install_rel if str(staging_path) and install_rel else Path("")

            state_text = str(item.get("state") or "").lower()
            if state_text and state_text != "installed":
                continue

            if not mod_path.exists():
                continue

            mods.append((
                display_name,
                {
                    "id": mod_id,
                    "name": display_name,
                    "installationPath": install_rel,
                    "mod_path": str(mod_path),
                    "version": str(item.get("version") or ""),
                    "archiveId": str(item.get("archiveId") or ""),
                },
            ))

        if not mods and self.mods_path and self.mods_path.exists():
            for folder in self.mods_path.iterdir():
                if folder.is_dir():
                    mods.append((
                        folder.name,
                        {
                            "id": folder.name,
                            "name": folder.name,
                            "installationPath": folder.name,
                            "mod_path": str(folder),
                            "version": "",
                            "archiveId": "",
                        },
                    ))

        if not mods:
            raise RuntimeError("No installed Vortex mods found. Sync Vortex metadata first.")

        mods.sort(key=lambda item: item[0].lower())

        dialog = QDialog(self)
        dialog.setWindowTitle("Choose Vortex Mod")
        dialog.resize(620, 560)
        dialog.setStyleSheet(DARK_STYLE)

        filter_box = QLineEdit(dialog)
        filter_box.setPlaceholderText("Filter installed mods")

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

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        list_widget.itemDoubleClicked.connect(lambda _item: dialog.accept())

        layout = QVBoxLayout(dialog)
        layout.addWidget(QLabel("Choose the installed Vortex mod:"))
        layout.addWidget(filter_box)
        layout.addWidget(list_widget, 1)
        layout.addWidget(buttons)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            raise RuntimeError("No mod selected.")

        item = list_widget.currentItem()
        if item is None:
            raise RuntimeError("No mod selected.")

        selected = item.text()
        for name, mod in mods:
            if name == selected:
                return name, mod

        raise RuntimeError(f"Could not resolve selected mod: {selected}")


    def _prompt_manual_link_values(self, mod_name: str) -> dict | None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Create Manual Source Link")
        dialog.resize(760, 340)
        dialog.setStyleSheet(DARK_STYLE)

        source = QComboBox(dialog)
        source.addItems(["loverslab", "patreon", "website", "nexus", "other"])

        page_url = QLineEdit(dialog)
        page_url.setPlaceholderText("https://...")

        page_title = QLineEdit(dialog)
        page_title.setText(mod_name)

        version = QLineEdit(dialog)
        version.setPlaceholderText("Example: 1.0.0")

        file_pattern = QLineEdit(dialog)
        file_pattern.setPlaceholderText("Optional. Example: ModName_{version}.7z")

        fixed_version = QComboBox(dialog)
        fixed_version.addItem("Manual / skip update checks", True)
        fixed_version.addItem("Allow update checks when supported", False)

        update_mode = QComboBox(dialog)
        configure_update_mode_combo(update_mode, UPDATE_MODE_SKIP)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)

        layout = QGridLayout(dialog)
        layout.addWidget(QLabel(f"Mod: {mod_name}"), 0, 0, 1, 2)

        layout.addWidget(QLabel("Source type"), 1, 0)
        layout.addWidget(source, 1, 1)

        layout.addWidget(QLabel("Page URL"), 2, 0)
        layout.addWidget(page_url, 2, 1)

        layout.addWidget(QLabel("Page title"), 3, 0)
        layout.addWidget(page_title, 3, 1)

        layout.addWidget(QLabel("Version"), 4, 0)
        layout.addWidget(version, 4, 1)

        layout.addWidget(QLabel("File pattern"), 5, 0)
        layout.addWidget(file_pattern, 5, 1)

        layout.addWidget(QLabel("Mode"), 6, 0)
        layout.addWidget(fixed_version, 6, 1)

        layout.addWidget(QLabel("Update mode"), 7, 0)
        layout.addWidget(update_mode, 7, 1)

        layout.addWidget(buttons, 8, 0, 1, 2)

        if dialog.exec() != QDialog.DialogCode.Accepted:
            return None

        url = page_url.text().strip()
        if not url:
            QMessageBox.information(self, "LL Integration", "Page URL is required.")
            return None

        selected_source = source.currentText().strip().lower()
        fixed = bool(fixed_version.currentData())

        return {
            "source": selected_source,
            "page_url": url,
            "page_title": page_title.text().strip() or mod_name,
            "version": version.text().strip(),
            "file_pattern": file_pattern.text().strip(),
            "update_mode": update_mode.currentData() or UPDATE_MODE_SKIP,
            "fixed_version": "true" if fixed else "false",
            "manual_update": "true" if fixed else "false",
            "skip_update_check": "true" if fixed else "false",
        }


    def _write_vortex_manual_link(self, mod: dict, values: dict) -> Path:
        mod_path = Path(str(mod.get("mod_path") or ""))
        if not mod_path.exists():
            raise RuntimeError(f"Vortex mod folder not found:\n{mod_path}")

        meta_path = mod_path / "meta.ini"

        parser = configparser.ConfigParser(interpolation=None)
        parser.optionxform = str

        if meta_path.exists():
            parser.read(meta_path, encoding="utf-8-sig")

        if "General" not in parser:
            parser["General"] = {}

        general = parser["General"]
        page_url = values.get("page_url", "").strip()
        version = values.get("version", "").strip()
        source = values.get("source", "website").strip().lower()

        general["url"] = page_url
        general["website"] = page_url
        general["homepage"] = page_url
        general["hasCustomURL"] = "true"
        general["repository"] = "LoversLab" if source == "loverslab" else source

        if version:
            general["version"] = version

        if "LoversLab" in parser:
            parser.remove_section("LoversLab")

        parser.add_section("LoversLab")

        now = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime())
        ll = parser["LoversLab"]

        archive_name = str(mod.get("archiveId") or mod.get("id") or mod.get("name") or "").strip()

        ll["source"] = source
        ll["page_url"] = page_url
        ll["page_title"] = values.get("page_title", "").strip()
        ll["version"] = version
        ll["file_name"] = values.get("file_pattern", "").strip()
        ll["file_pattern"] = values.get("file_pattern", "").strip()
        ll["archive_name"] = archive_name
        ll["captured_at"] = now
        ll["completed_at"] = now
        ll["update_mode"] = values.get("update_mode", UPDATE_MODE_SKIP)
        ll["fixed_version"] = values.get("fixed_version", "true")
        ll["manual_update"] = values.get("manual_update", "true")
        ll["skip_update_check"] = values.get("skip_update_check", "true")
        ll["manual_install"] = "true"
        ll["multipart"] = "false"

        meta_path.parent.mkdir(parents=True, exist_ok=True)
        with meta_path.open("w", encoding="utf-8") as handle:
            parser.write(handle, space_around_delimiters=False)

        return meta_path
    def _edit_link_dialog(self, row: dict) -> bool:
        sidecar = Path(row.get("sidecar") or "")
        data = load_ini(sidecar) if sidecar.exists() else {}
        dialog = QDialog(self)
        dialog.setWindowTitle("Edit LoversLab Link")
        dialog.setStyleSheet(DARK_STYLE)
        dialog.resize(720, 260)

        page_url = QLineEdit(data.get("page_url") or row.get("page_url") or "", dialog)
        version = QLineEdit(data.get("version") or row.get("version") or "", dialog)
        file_pattern = QLineEdit(link_pattern_value(data, row), dialog)
        update_mode = QComboBox(dialog)
        configure_update_mode_combo(
            update_mode,
            normalized_update_mode(data.get("update_mode"), fixed=fixed_update(data)),
        )

        try_pattern = QPushButton("Try Pattern")
        help_button = QPushButton("Pattern Help")
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(dialog.accept)
        buttons.rejected.connect(dialog.reject)
        try_pattern.clicked.connect(
            lambda _checked=False: self._try_link_pattern(
                dialog,
                page_url,
                version,
                file_pattern,
                data,
            )
        )
        help_button.clicked.connect(lambda _checked=False: self._show_pattern_help(dialog))

        layout = QGridLayout(dialog)
        layout.addWidget(QLabel(f"Archive: {row.get('archive') or ''}"), 0, 0, 1, 3)
        layout.addWidget(QLabel("LoversLab page URL"), 1, 0)
        layout.addWidget(page_url, 1, 1, 1, 2)
        layout.addWidget(QLabel("Current version"), 2, 0)
        layout.addWidget(version, 2, 1, 1, 2)
        layout.addWidget(QLabel("File pattern"), 3, 0)
        layout.addWidget(file_pattern, 3, 1, 1, 2)
        layout.addWidget(QLabel("Update mode"), 4, 0)
        layout.addWidget(update_mode, 4, 1, 1, 2)
        layout.addWidget(try_pattern, 5, 0)
        layout.addWidget(help_button, 5, 1)
        layout.addWidget(buttons, 5, 2)
        if dialog.exec() != QDialog.DialogCode.Accepted:
            return False

        target = sidecar if str(sidecar) else Path(row["path"]).with_name(f"{row['archive']}.ll.ini")
        new_update_mode = normalized_update_mode(str(update_mode.currentData() or ""))
        new_fixed = new_update_mode == UPDATE_MODE_SKIP
        updated = dict(data)
        updated.update({
            "source": "loverslab",
            "page_url": page_url.text().strip(),
            "download_url": data.get("download_url") or (with_query_value(page_url.text().strip(), "do", "download") if page_url.text().strip() else ""),
            "file_pattern": file_pattern.text().strip(),
            "file_name": data.get("file_name") or data.get("archive_name") or row.get("archive") or file_pattern.text().strip(),
            "archive_name": row.get("archive") or "",
            "version": version.text().strip(),
            "update_mode": new_update_mode,
            "fixed_version": "true" if new_fixed else "false",
            "manual_update": "true" if new_fixed else "false",
            "skip_update_check": "true" if new_fixed else "false",
        })
        write_ini(target, updated)
        metadata_target = self.metadata_path / "downloads" / target.name
        write_ini(metadata_target, updated)
        return True

    def _try_link_pattern(
        self,
        parent: QDialog,
        page_url_widget: QLineEdit,
        version_widget: QLineEdit,
        pattern_widget: QLineEdit,
        data: dict,
    ) -> None:
        page_url = page_url_widget.text().strip()
        pattern = pattern_widget.text().strip()
        if not page_url:
            QMessageBox.critical(parent, "LL Integration", "LoversLab page URL is required.")
            return
        if not pattern:
            QMessageBox.critical(parent, "LL Integration", "File pattern is required.")
            return

        try:
            downloads = fetch_ll_downloads(page_url, self.cookies_path, float(self.timeout.value()))
            match = choose_latest(downloads, pattern)
        except Exception as exc:
            QMessageBox.critical(parent, "LL Integration", f"Pattern test failed:\n\n{exc}")
            return

        if not match:
            sample = "\n".join(download.name for download in downloads[:8])
            QMessageBox.warning(
                parent,
                "LL Integration",
                f"No matching download found for:\n{pattern}\n\n"
                f"Downloads seen: {len(downloads)}\n\n{sample}",
            )
            return

        message = QMessageBox(parent)
        message.setWindowTitle("LL Integration")
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
        message.addButton(QMessageBox.StandardButton.Cancel)
        message.setDefaultButton(use_button)
        message.exec()
        if message.clickedButton() != use_button:
            return

        version_widget.setText(match.version or "")
        pattern_widget.setText(pattern)
        data["file_name"] = match.name
        data["archive_name"] = match.name
        data["original_archive_name"] = match.name
        data["download_url"] = match.url
        data["size"] = match.size
        data["date_iso"] = match.date_iso

    def _show_pattern_help(self, parent: QDialog) -> None:
        QMessageBox.information(
            parent,
            "Pattern Help",
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

    def _apply_filter(self) -> None:
        mode = self.filter_mode.currentText() if hasattr(self, "filter_mode") else "All links"
        needle = self.filter_text.text().strip().lower()
        first_visible = -1
        visible = 0
        for row_index, row in enumerate(self.rows):
            status = self._row_status(row)
            current = str(row.get("version") or "").strip()
            fixed = bool(row.get("fixed"))
            mode_match = True
            if mode == "Updates":
                mode_match = status == "Update"
            elif mode == "OK":
                mode_match = status == "OK"
            elif mode == "Unknown / missing version":
                mode_match = status == "Unknown" or (not current and not fixed)
            elif mode == "Manual links":
                mode_match = fixed or status == "Manual"
            elif mode == "Errors / skipped":
                mode_match = status in {"Error", "Manual", "Untracked"}
            elif mode == "Not checked":
                mode_match = status == "Ready"

            haystack = " ".join(str(value or "") for value in row.values()).lower()
            show = mode_match and (not needle or needle in haystack)
            self.table.setRowHidden(row_index, not show)
            if show:
                visible += 1
            if show and first_visible < 0:
                first_visible = row_index
        if hasattr(self, "filter_count"):
            self.filter_count.setText(f"{visible} / {len(self.rows)}")
        selected = self.table.selectedItems()
        if first_visible >= 0 and (not selected or self.table.isRowHidden(selected[0].row())):
            self.table.selectRow(first_visible)


def main() -> int:
    mode = "links"
    for arg in sys.argv[1:]:
        if arg.startswith("--mode="):
            mode = arg.split("=", 1)[1].strip().lower()

    tool_id = mode.replace("_", "-")
    set_windows_app_id(tool_id)

    app = QApplication(sys.argv)
    icon = apply_app_icon(app, tool_id)

    if mode == "voice":
        dialog = VortexVoiceFinder()
        if not icon.isNull():
            dialog.setWindowIcon(icon)
        dialog.exec()
        return 0

    if mode in {"purge", "purge-suspicious", "purge_suspicious"}:
        dialog = VortexManager()
        if not icon.isNull():
            dialog.setWindowIcon(icon)
        QTimer.singleShot(0, dialog.purge_suspicious_links)
        QTimer.singleShot(250, dialog.close)
        dialog.exec()
        return 0

    if mode in {"create-link", "create_link", "manual-link", "manual_link"}:
        manager = VortexManager()
        if not icon.isNull():
            manager.setWindowIcon(icon)
        manager.hide()

        def run_manual_link_only() -> None:
            try:
                manager.create_manual_link()
            finally:
                QApplication.instance().quit()

        QTimer.singleShot(0, run_manual_link_only)
        return app.exec()

    debug_icon_status(tool_id, icon)

    dialog = VortexManager()
    if not icon.isNull():
        dialog.setWindowIcon(icon)
    dialog.exec()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
