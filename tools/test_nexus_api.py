import json
import os
import sys
from urllib.request import Request, urlopen


def main() -> int:
    api_key = os.environ.get("NEXUS_API_KEY", "").strip()
    if not api_key:
        print("Missing NEXUS_API_KEY environment variable.", file=sys.stderr)
        return 2

    game = sys.argv[1] if len(sys.argv) > 1 else "skyrimspecialedition"
    mod_id = sys.argv[2] if len(sys.argv) > 2 else "88908"
    url = f"https://api.nexusmods.com/v1/games/{game}/mods/{mod_id}/files.json"
    request = Request(
        url,
        headers={
            "apikey": api_key,
            "User-Agent": "LL Integration API probe",
            "Accept": "application/json",
        },
    )
    with urlopen(request, timeout=30) as response:
        payload = json.loads(response.read().decode("utf-8"))

    files = payload.get("files", payload if isinstance(payload, list) else [])
    print(f"Fetched {len(files)} files from {game}/{mod_id}")
    for item in files[:20]:
        print(json.dumps({
            "file_id": item.get("file_id"),
            "name": item.get("name"),
            "category_name": item.get("category_name"),
            "version": item.get("version"),
            "size_kb": item.get("size_kb"),
            "uploaded_time": item.get("uploaded_time"),
            "mod_version": item.get("mod_version"),
        }, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
