import argparse
import configparser
import fnmatch
import json
from dataclasses import asdict
from pathlib import Path
from typing import List, Optional
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode, urlsplit, urlunsplit

try:
    from .utils import (
        LLDownload,
        compare_versions,
        extract_downloads,
        fetch_ll_html,
        filename_prefix,
        version_key,
    )
except ImportError:
    from utils import (
        LLDownload,
        compare_versions,
        extract_downloads,
        fetch_ll_html,
        filename_prefix,
        version_key,
    )


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_COOKIES = ROOT_DIR / "native-app" / "cookies_storage" / "cookies_ll.json"
DEFAULT_INI = Path(__file__).resolve().parent / "LL.sample.ini"
DEFAULT_OUT = Path(__file__).resolve().parent / "update_check.json"


def read_ll_ini(path: Path) -> configparser.SectionProxy:
    config = configparser.ConfigParser(interpolation=None)
    loaded = config.read(path, encoding="utf-8")
    if not loaded:
        raise FileNotFoundError(path)
    if "LoversLab" not in config:
        raise RuntimeError(f"{path} is missing a [LoversLab] section")
    return config["LoversLab"]


def with_query_value(url: str, key: str, value: str) -> str:
    parts = urlsplit(url)
    path = parts.path
    if path.startswith("/files/file/") and not path.endswith("/"):
        path = f"{path}/"
    query = dict(parse_qsl(parts.query, keep_blank_values=True))
    query[key] = value
    return urlunsplit((parts.scheme, parts.netloc, path, urlencode(query), parts.fragment))


def download_page_url(config: configparser.SectionProxy) -> str:
    page_url = config.get("page_url", "").strip()
    download_url = config.get("download_url", "").strip()
    base = page_url or download_url
    if not base:
        raise RuntimeError("LL metadata needs page_url or download_url")
    return with_query_value(base, "do", "download")


def score_candidate(download: LLDownload, known_file: str) -> int:
    score = 0
    pattern = known_file.strip()
    if any(char in pattern for char in "*?[]"):
        lower_name = download.name.lower()
        lower_pattern = pattern.lower()
        if fnmatch.fnmatch(lower_name, lower_pattern):
            score += 130
        elif fnmatch.fnmatch(lower_name, f"*{lower_pattern}*"):
            score += 95

    known_prefix = filename_prefix(known_file)
    download_prefix = filename_prefix(download.name)

    if known_prefix and download_prefix == known_prefix:
        score += 100
    elif known_prefix and known_prefix in download_prefix:
        score += 60

    if not any(char in pattern for char in "*?[]") and Path(download.name).suffix.lower() == Path(known_file).suffix.lower():
        score += 20

    if download.version:
        score += 10

    return score


def choose_latest(downloads: List[LLDownload], known_file: str) -> Optional[LLDownload]:
    scored = [
        (score_candidate(download, known_file), download)
        for download in downloads
        if download.version
    ]
    scored = [item for item in scored if item[0] >= 80]
    if not scored:
        return None

    return max(scored, key=lambda item: (item[0], version_key(item[1].version or "0")))[1]


def load_html(args: argparse.Namespace, config: configparser.SectionProxy) -> tuple[str, str]:
    if args.html:
        return args.html.read_text(encoding="utf-8"), str(args.html)

    url = download_page_url(config)
    html = fetch_ll_html(
        url,
        args.cookies,
        referer=config.get("page_url", fallback=None),
        timeout=getattr(args, "timeout", 30.0),
    )
    return html, url


def write_result(path: Path, payload: dict) -> None:
    path.write_text(json.dumps(payload, indent=2), encoding="utf-8")


def check_ini_for_updates(
    ini_path: Path,
    cookies_path: Path = DEFAULT_COOKIES,
    html_path: Optional[Path] = None,
    out_path: Optional[Path] = None,
    timeout: float = 30.0,
) -> dict:
    config = read_ll_ini(ini_path)
    known_file = (
        config.get("file_pattern", "").strip()
        or config.get("file_name", "").strip()
        or config.get("archive_name", "").strip()
    )
    current_version = config.get("version", "").strip()

    if not known_file:
        raise RuntimeError("LL metadata needs file_name")
    if not current_version:
        raise RuntimeError("LL metadata needs version")

    class Args:
        pass

    Args.cookies = cookies_path
    Args.html = html_path
    Args.timeout = timeout

    html, source_url = load_html(Args, config)
    downloads = extract_downloads(html)
    latest = choose_latest(downloads, known_file)
    latest_version = latest.version if latest else None
    comparison = compare_versions(latest_version, current_version) if latest_version else None

    payload = {
        "sourceUrl": source_url,
        "knownFile": known_file,
        "currentVersion": current_version,
        "latest": asdict(latest) if latest else None,
        "updateAvailable": comparison is not None and comparison > 0,
        "downloadsSeen": [asdict(download) for download in downloads],
    }

    if out_path:
        write_result(out_path, payload)

    return payload


def main() -> int:
    parser = argparse.ArgumentParser(description="Check a LoversLab-backed mod for updates.")
    parser.add_argument("--ini", type=Path, default=DEFAULT_INI, help=f"LL.ini path, default: {DEFAULT_INI}")
    parser.add_argument("--cookies", type=Path, default=DEFAULT_COOKIES, help=f"Cookie export path, default: {DEFAULT_COOKIES}")
    parser.add_argument("--html", type=Path, help="Parse local HTML instead of fetching LoversLab")
    parser.add_argument("--out", type=Path, default=DEFAULT_OUT, help=f"Output JSON path, default: {DEFAULT_OUT}")
    parser.add_argument("--timeout", type=float, default=30.0, help="HTTP timeout in seconds, default: 30")
    args = parser.parse_args()

    try:
        payload = check_ini_for_updates(args.ini, args.cookies, args.html, args.out)
    except HTTPError as exc:
        print(f"HTTP error {exc.code}: {exc.reason}")
        return 1
    except URLError as exc:
        print(f"Network error: {exc.reason}")
        return 1

    latest = payload["latest"]
    if latest:
        status = "update available" if payload["updateAvailable"] else "up to date"
        print(f"{status}: current {payload['currentVersion']}, latest {latest['version']}")
    else:
        print(f"No matching download found for {payload['knownFile']}")
    print(f"Wrote result to {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
