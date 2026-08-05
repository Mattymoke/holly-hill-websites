"""Publish a single lot from pending_lot.json to the Holly Hill Surplus shop.

Reads config.ini and pending_lot.json in this same folder, POSTs the lot
(plus any photos) to /api/admin/sync-lot, and writes the outcome to
result.json.
"""

import configparser
import json
import mimetypes
import os
import sys
from datetime import datetime, timezone

import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.ini")
PENDING_LOT_PATH = os.path.join(SCRIPT_DIR, "pending_lot.json")
RESULT_PATH = os.path.join(SCRIPT_DIR, "result.json")

VALID_STATUSES = ("available", "reserved", "sold", "shipped")

# mimetypes.guess_type() depends on the OS's MIME database (e.g. Windows
# registry entries), which isn't always reliable -- fall back to this
# explicit map for the extensions we actually expect before trusting it.
EXTENSION_CONTENT_TYPES = {
    ".jpg": "image/jpeg",
    ".jpeg": "image/jpeg",
    ".png": "image/png",
    ".webp": "image/webp",
    ".gif": "image/gif",
}


def guess_content_type(filename):
    ext = os.path.splitext(filename)[1].lower()
    if ext in EXTENSION_CONTENT_TYPES:
        return EXTENSION_CONTENT_TYPES[ext]
    guessed, _ = mimetypes.guess_type(filename)
    return guessed or "image/jpeg"


class SyncError(Exception):
    pass


def write_result(success, **fields):
    result = {"success": success, "timestamp": datetime.now(timezone.utc).isoformat()}
    result.update(fields)
    with open(RESULT_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)
    return result


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise SyncError(f"Config file not found: {CONFIG_PATH}")

    config = configparser.ConfigParser()
    try:
        config.read(CONFIG_PATH)
        api_base_url = config.get("sync", "api_base_url").strip()
        sync_api_key = config.get("sync", "sync_api_key").strip()
    except configparser.Error as err:
        raise SyncError(f"Could not read config.ini: {err}")

    if not sync_api_key or sync_api_key == "PASTE_YOUR_KEY_HERE":
        raise SyncError("sync_api_key in config.ini has not been set yet")

    return api_base_url, sync_api_key


def load_pending_lot():
    if not os.path.exists(PENDING_LOT_PATH):
        raise SyncError(f"pending_lot.json not found: {PENDING_LOT_PATH}")

    with open(PENDING_LOT_PATH, "r", encoding="utf-8") as f:
        try:
            return json.load(f)
        except json.JSONDecodeError as err:
            raise SyncError(f"pending_lot.json is not valid JSON: {err}")


def main():
    api_base_url, sync_api_key = load_config()
    lot = load_pending_lot()

    lot_id = lot.get("id")
    name = lot.get("name")
    category = lot.get("category")
    description = lot.get("description")
    price_dollars = lot.get("price_dollars")
    status = lot.get("status")
    photo_paths = lot.get("photo_paths") or []
    clear_images = bool(lot.get("clear_images", False))
    # Featured comes straight from pending_lot.json, which the VBA side
    # already populated from the sheet's Featured column (I). Original
    # Price / Profit (G/H) are PRIVATE and never appear in pending_lot.json
    # in the first place, so there's nothing to filter out here.
    featured = bool(lot.get("featured", False))

    if not lot_id or not name:
        raise SyncError("pending_lot.json must include at least 'id' and 'name'")
    if price_dollars is None:
        raise SyncError("pending_lot.json must include 'price_dollars'")
    if status not in VALID_STATUSES:
        raise SyncError(f"status must be one of {', '.join(VALID_STATUSES)}, got: {status!r}")

    website_price_cents = round(float(price_dollars) * 100)

    missing_photos = [p for p in photo_paths if not os.path.isfile(p)]
    if missing_photos:
        raise SyncError("Missing photo file(s): " + ", ".join(missing_photos))

    data = {
        "id": lot_id,
        "name": name,
        "website_price_cents": str(website_price_cents),
        "status": status,
        "clear_images": "true" if clear_images else "false",
        "featured": "true" if featured else "false",
    }
    if category:
        data["category"] = category
    if description:
        data["description"] = description

    opened_files = [open(path, "rb") for path in photo_paths]
    try:
        files = [
            ("images", (os.path.basename(path), fh, guess_content_type(path)))
            for path, fh in zip(photo_paths, opened_files)
        ]
        url = api_base_url.rstrip("/") + "/api/admin/sync-lot"
        headers = {"Authorization": f"Bearer {sync_api_key}"}

        try:
            response = requests.post(url, headers=headers, data=data, files=files or None, timeout=60)
        except requests.exceptions.RequestException as err:
            raise SyncError(f"Network error while contacting {url}: {err}")
    finally:
        for fh in opened_files:
            fh.close()

    if response.status_code != 200:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise SyncError(f"Server returned {response.status_code}: {detail}")

    try:
        lot_response = response.json()
    except ValueError:
        raise SyncError(f"Server returned 200 but the response wasn't valid JSON: {response.text}")

    if "lot" not in lot_response:
        raise SyncError(f"Server returned 200 but the response had no 'lot' field: {lot_response}")

    write_result(True, lot=lot_response["lot"])
    print(f"OK: synced lot '{lot_id}' -- {len(photo_paths)} photo(s) uploaded.")


if __name__ == "__main__":
    try:
        main()
    except SyncError as err:
        write_result(False, error=str(err))
        print(f"FAILED: {err}")
        sys.exit(1)
    except Exception as err:
        write_result(False, error=f"Unexpected error: {err}")
        print(f"FAILED: Unexpected error: {err}")
        sys.exit(1)
