import argparse
import json
from dataclasses import asdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional
from urllib.error import HTTPError, URLError

try:
    from .utils import extract_downloads, fetch_ll_html
except ImportError:
    from utils import extract_downloads, fetch_ll_html


ROOT_DIR = Path(__file__).resolve().parents[1]
DEFAULT_COOKIES = ROOT_DIR / "native-app" / "cookies_storage" / "cookies_ll.json"
DEFAULT_OUT = Path(__file__).resolve().parent / "downloads_ll.json"


def fetch_html(url: str, cookies_path: Path) -> str:
    return fetch_ll_html(url, cookies_path)


def write_downloads(
    html: str,
    out_path: Path,
    source_url: Optional[str] = None,
) -> int:
    downloads = extract_downloads(html)
    payload = {
        "sourceUrl": source_url or "",
        "fetchedAt": datetime.now(timezone.utc).isoformat(),
        "downloads": [asdict(download) for download in downloads],
    }

    out_path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
    return len(downloads)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Fetch a LoversLab download page and write detected files to JSON."
    )
    parser.add_argument("url", nargs="?", help="LoversLab URL to fetch")
    parser.add_argument(
        "--html",
        type=Path,
        help="Parse a local HTML file instead of fetching a URL",
    )
    parser.add_argument(
        "--cookies",
        type=Path,
        default=DEFAULT_COOKIES,
        help=f"Cookie export path, default: {DEFAULT_COOKIES}",
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=DEFAULT_OUT,
        help=f"Output JSON path, default: {DEFAULT_OUT}",
    )
    args = parser.parse_args()

    if args.html:
        html = args.html.read_text(encoding="utf-8")
        source_url = str(args.html)
    elif args.url:
        try:
            html = fetch_html(args.url, args.cookies)
        except HTTPError as exc:
            print(f"HTTP error {exc.code}: {exc.reason}")
            return 1
        except URLError as exc:
            print(f"Network error: {exc.reason}")
            return 1
        source_url = args.url
    else:
        parser.error("provide a URL or --html")

    count = write_downloads(html, args.out, source_url)
    print(f"Wrote {count} downloads to {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
