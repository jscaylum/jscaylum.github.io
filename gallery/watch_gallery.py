#!/usr/bin/env python3
import json
import time
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


def write_manifest() -> None:
    payload = build_manifest()
    data = json.dumps(payload, indent=2) + '\n'
    current = None
    if MANIFEST_PATH.exists():
        current = MANIFEST_PATH.read_text(encoding='utf-8')
    if current != data:
        MANIFEST_PATH.write_text(data, encoding='utf-8')
        print(f"Updated gallery manifest: {len(payload['images'])} image(s)")


if __name__ == '__main__':
    print('Watching gallery folder for new images...')
    write_manifest()
    while True:
        write_manifest()
        time.sleep(2)
