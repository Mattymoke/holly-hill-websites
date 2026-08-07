"""Pull sold/all order records from the Holly Hill Surplus shop for the
private Sold Lots tracking spreadsheet.

Reads config.ini, makes one GET to /api/sync/orders, and writes
sold_orders.json shaped as:

  {"success": true, "orders": [{"id": "...", "lot_id": "...", "lot_name": "...",
   "category": "...", "amount_cents": 7000, "buyer_email": "...",
   "status": "paid", "created_at": "...", "paid_at": "..."}, ...], "timestamp": "..."}

or on failure:

  {"success": false, "error": "...", "timestamp": "..."}

This is a pull-only script -- nothing here ever writes back to the website.
"""

import configparser
import json
import os
import sys
from datetime import datetime, timezone

import requests

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(SCRIPT_DIR, "config.ini")
SOLD_ORDERS_PATH = os.path.join(SCRIPT_DIR, "sold_orders.json")


class PullError(Exception):
    pass


def write_result(success, **fields):
    result = {"success": success, "timestamp": datetime.now(timezone.utc).isoformat()}
    result.update(fields)
    with open(SOLD_ORDERS_PATH, "w", encoding="utf-8") as f:
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

    url = api_base_url.rstrip("/") + "/api/sync/orders"
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

    if "orders" not in data or not isinstance(data["orders"], list):
        raise PullError(f"Server returned 200 but the response had no 'orders' array: {data}")

    write_result(True, orders=data["orders"])
    print(f"OK: pulled {len(data['orders'])} order(s).")


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
