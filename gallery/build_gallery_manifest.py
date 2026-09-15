#!/usr/bin/env python3
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parent
MANIFEST_PATH = ROOT / 'gallery.json'
IMAGE_EXTENSIONS = {'.png', '.jpg', '.jpeg', '.webp', '.gif', '.avif'}


def is_image(path: Path) -> bool:
    return path.is_file() and path.suffix.lower() in IMAGE_EXTENSIONS


def build_manifest() -> dict:
    images = []
    seen = set()
    for path in sorted(ROOT.iterdir(), key=lambda p: p.name.lower()):
        if path.name == MANIFEST_PATH.name:
            continue
        if not is_image(path):
            continue
        src = f"gallery/{path.name}"
        if src in seen:
            continue
        seen.add(src)
        images.append({"src": src})
    return {"images": images}


def main() -> None:
    payload = build_manifest()
    with MANIFEST_PATH.open('w', encoding='utf-8') as f:
        json.dump(payload, f, indent=2)
        f.write('\n')
    print(f"Wrote {len(payload['images'])} image(s) to {MANIFEST_PATH.name}")


if __name__ == '__main__':
    main()
