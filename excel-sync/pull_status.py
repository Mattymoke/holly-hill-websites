"""Pull current lot statuses from the Holly Hill Surplus shop.

Reads config.ini, makes one GET to /api/sync/lots, and writes
remote_status.json shaped as:

  {"success": true, "lots": {"<id>": {"status": "...", "updated_at": "..."}, ...}, "timestamp": "..."}

or on failure:

  {"success": false, "error": "...", "timestamp": "..."}

One fast call per invocation, not per-row -- meant to be run once at the
start of a save, not once per lot.
"""

import configparser
import json
import os
import sys
from datetime import datetime, timezone

import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.ini")
REMOTE_STATUS_PATH = os.path.join(SCRIPT_DIR, "remote_status.json")


class PullError(Exception):
    pass


def write_result(success, **fields):
    result = {"success": success, "timestamp": datetime.now(timezone.utc).isoformat()}
    result.update(fields)
    with open(REMOTE_STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(result, f, indent=2)
    return result


def load_config():
    if not os.path.exists(CONFIG_PATH):
        raise PullError(f"Config file not found: {CONFIG_PATH}")

    config = configparser.ConfigParser()
    try:
        config.read(CONFIG_PATH)
        api_base_url = config.get("sync", "api_base_url").strip()
        sync_api_key = config.get("sync", "sync_api_key").strip()
    except configparser.Error as err:
        raise PullError(f"Could not read config.ini: {err}")

    if not sync_api_key or sync_api_key == "PASTE_YOUR_KEY_HERE":
        raise PullError("sync_api_key in config.ini has not been set yet")

    return api_base_url, sync_api_key


def main():
    api_base_url, sync_api_key = load_config()

    url = api_base_url.rstrip("/") + "/api/sync/lots"
    headers = {"Authorization": f"Bearer {sync_api_key}"}

    try:
        response = requests.get(url, headers=headers, timeout=30)
    except requests.exceptions.RequestException as err:
        raise PullError(f"Network error while contacting {url}: {err}")

    if response.status_code != 200:
        try:
            detail = response.json()
        except ValueError:
            detail = response.text
        raise PullError(f"Server returned {response.status_code}: {detail}")

    try:
        data = response.json()
    except ValueError:
        raise PullError(f"Server returned 200 but the response wasn't valid JSON: {response.text}")

    if "lots" not in data or not isinstance(data["lots"], list):
        raise PullError(f"Server returned 200 but the response had no 'lots' array: {data}")

    lots_by_id = {}
    for lot in data["lots"]:
        lot_id = lot.get("id")
        if not lot_id:
            continue
        lots_by_id[lot_id] = {"status": lot.get("status"), "updated_at": lot.get("updated_at")}

    write_result(True, lots=lots_by_id)
    print(f"OK: pulled status for {len(lots_by_id)} lot(s).")


if __name__ == "__main__":
    try:
        main()
    except PullError as err:
        write_result(False, error=str(err))
        print(f"FAILED: {err}")
        sys.exit(1)
    except Exception as err:
        write_result(False, error=f"Unexpected error: {err}")
        print(f"FAILED: Unexpected error: {err}")
        sys.exit(1)
