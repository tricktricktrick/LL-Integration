import json
import re
import hashlib
import gzip
import zlib
from dataclasses import dataclass
from html.parser import HTMLParser
from pathlib import Path
from typing import Dict, Iterable, List, Optional
from urllib.parse import unquote
from urllib.request import Request, urlopen


COOKIE_NAMES = {
    "ips4_IPSSessionFront",
    "ips4_member_id",
    "ips4_login_key",
}
QUICK_HASH_CHUNK_SIZE = 1024 * 1024
VERSION_RE = r"\bv?(\d+(?:[.-]\d+){1,3})(?:\b|(?=\D))"


@dataclass(frozen=True)
class LLDownload:
    name: str
    url: str
    size: Optional[str] = None
    date_iso: Optional[str] = None
    version: Optional[str] = None


class _LLDownloadParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.downloads: List[LLDownload] = []
        self._in_title = False
        self._in_meta = False
        self._current_name: Optional[str] = None
        self._current_size_parts: List[str] = []
        self._current_date_iso: Optional[str] = None

    def handle_starttag(self, tag: str, attrs: List[tuple]) -> None:
        attrs_dict = dict(attrs)
        classes = attrs_dict.get("class", "")

        if tag == "span" and "ipsType_break" in classes:
            self._in_title = True
            self._current_name = ""
            self._current_size_parts = []
            self._current_date_iso = None
            return

        if tag == "p" and "ipsDataItem_meta" in classes:
            self._in_meta = True
            return

        if tag == "time" and self._in_meta:
            self._current_date_iso = attrs_dict.get("datetime")
            return

        if tag == "a" and attrs_dict.get("data-action") == "download":
            href = attrs_dict.get("href")
            if href and self._current_name:
                size = " ".join(" ".join(self._current_size_parts).split()) or None
                self.downloads.append(
                    LLDownload(
                        name=self._current_name.strip(),
                        url=unquote(href).replace("&amp;", "&"),
                        size=size.split(" / ")[0].strip() if size else None,
                        date_iso=self._current_date_iso,
                        version=extract_version(self._current_name),
                    )
                )

    def handle_endtag(self, tag: str) -> None:
        if tag == "span" and self._in_title:
            self._in_title = False
        elif tag == "p" and self._in_meta:
            self._in_meta = False

    def handle_data(self, data: str) -> None:
        if self._in_title and self._current_name is not None:
            self._current_name += data
        elif self._in_meta:
            self._current_size_parts.append(data)


def load_ll_cookies(path: Path, required_only: bool = True) -> Dict[str, str]:
    if not path.exists():
        raise RuntimeError(
            f"LoversLab cookies are not exported yet. Open the Firefox LL Integration popup and click Export Cookies.\n"
            f"Expected file: {path}"
        )

    data = json.loads(path.read_text(encoding="utf-8"))
    cookies = data.get("cookies", [])
    return {
        cookie["name"]: cookie["value"]
        for cookie in cookies
        if cookie.get("value") and (not required_only or cookie.get("name") in COOKIE_NAMES)
    }


def cookie_header(cookies: Dict[str, str]) -> str:
    return "; ".join(f"{name}={value}" for name, value in cookies.items())


def fetch_ll_html(url: str, cookies_path: Path, referer: Optional[str] = None, timeout: float = 30.0) -> str:
    cookies = load_ll_cookies(cookies_path, required_only=False)
    if not cookies:
        raise RuntimeError(f"No usable LoversLab cookies found in {cookies_path}")

    headers = {
        "Cookie": cookie_header(cookies),
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) "
            "Gecko/20100101 Firefox/125.0"
        ),
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

    request = Request(url, headers=headers)
    with urlopen(request, timeout=timeout) as response:
        charset = response.headers.get_content_charset() or "utf-8"
        data = response.read()
        encoding = (response.headers.get("Content-Encoding") or "").lower()
        if encoding == "gzip":
            data = gzip.decompress(data)
        elif encoding == "deflate":
            data = zlib.decompress(data)
        return data.decode(charset, errors="replace")


def extract_downloads(html: str) -> List[LLDownload]:
    parser = _LLDownloadParser()
    try:
        parser.feed(html)
    except AssertionError as exc:
        if "expected name token" not in str(exc):
            raise

        # LoversLab pages can contain malformed <![...]> script/comment fragments
        # from embedded post content. They are irrelevant for the download list,
        # but Python's HTMLParser is strict enough to abort on them.
        parser = _LLDownloadParser()
        parser.feed(html.replace("<![", "&lt;!["))
    return parser.downloads


def extract_version(file_name: str) -> Optional[str]:
    stem = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", file_name, flags=re.IGNORECASE)
    match = re.search(VERSION_RE, stem, re.IGNORECASE)
    return match.group(1) if match else None


def downloads_to_json(downloads: Iterable[LLDownload]) -> str:
    return json.dumps([download.__dict__ for download in downloads], indent=2)


def version_key(version: str) -> tuple:
    return tuple(int(part) for part in re.findall(r"\d+", version))


def compare_versions(left: str, right: str) -> int:
    left_parts = version_key(left)
    right_parts = version_key(right)
    max_len = max(len(left_parts), len(right_parts))
    left_parts += (0,) * (max_len - len(left_parts))
    right_parts += (0,) * (max_len - len(right_parts))
    return (left_parts > right_parts) - (left_parts < right_parts)


def filename_prefix(file_name: str) -> str:
    stem = re.sub(r"\.(?:7z|zip|rar|tar|gz|bz2|xz)$", "", file_name, flags=re.IGNORECASE)
    prefix = re.split(VERSION_RE, stem, maxsplit=1, flags=re.IGNORECASE)[0]
    return re.sub(r"[^a-z0-9]+", " ", prefix.lower()).strip()


def archive_quick_hash(path: Path) -> str:
    size = path.stat().st_size
    digest = hashlib.sha256()
    digest.update(str(size).encode("ascii"))

    with path.open("rb") as file:
        digest.update(file.read(QUICK_HASH_CHUNK_SIZE))
        if size > QUICK_HASH_CHUNK_SIZE:
            file.seek(max(size - QUICK_HASH_CHUNK_SIZE, 0))
            digest.update(file.read(QUICK_HASH_CHUNK_SIZE))

    return digest.hexdigest()
