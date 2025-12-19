import requests
import json
from time import sleep

# --- CONFIGURATION ---
BASE_URL = "http://127.0.0.1:5000/api/assets"
TEST_SERIAL = "test serial"
# --- END CONFIGURATION ---

# 1. Simulated Intune device data (Use your internal keys)
# This is the data you would normally fetch from Microsoft Graph.
INTUNE_DEVICE_DATA = {
    "serial_number": TEST_SERIAL,
    "asset_type": "Laptop",
    "product": "Surface Laptop Studio",
    "used_by_name": "Test User Sync",
    "used_by_email": "test.user.sync@example.com",
    "department": "IT",
    "location": "Tel Aviv Office - Desk 301",
    "name": "INTUNE-XYZ-TLS",  # The device name/tag from Intune
    # Note: Keys here must match the fields in your inventory!
}


def create_or_update_asset(device_data):
    """
    Attempts to update the asset, and if it fails (404), creates it.
    """
    serial = device_data["serial_number"]

    print(f"\nAttempting to UPDATE asset with serial: {serial}...")

    # Try to UPDATE (PUT) the asset first
    update_url = f"{BASE_URL}/{serial}"
    try:
        response = requests.put(update_url, json=device_data)

        if response.status_code == 200:
            print(f"✅ Success: Asset {serial} updated.")
            return

        elif response.status_code == 404:
            print(f"Asset {serial} not found. Proceeding to CREATE...")
            # If not found, fall through to the POST logic

        else:
            print(
                f"❌ Failed to update. Status: {response.status_code}. Error: {response.json().get('error', 'Unknown Error')}")
            return

    except requests.exceptions.ConnectionError:
        print("❌ Connection Error: Is your Flask app running?")
        return

    # If the update failed with 404, or we skipped straight to create:
    print(f"Attempting to CREATE new asset with serial: {serial}...")
    try:
        response = requests.post(BASE_URL, json=device_data)

        if response.status_code == 201:
            print(f"✅ Success: Asset {serial} created. New ID: {response.json().get('id')}")
        elif response.status_code == 409:
            # Asset exists but PUT failed earlier (shouldn't happen, but good check)
            print(f"❌ Conflict: Asset {serial} already exists (POST failed).")
        else:
            print(
                f"❌ Failed to create. Status: {response.status_code}. Error: {response.json().get('error', 'Unknown Error')}")

    except requests.exceptions.RequestException as e:
        print(f"❌ An error occurred during POST: {e}")


if __name__ == '__main__':
    print("--- Starting Local Intune Sync Test ---")

    # 1. Simulate initial run (will CREATE the device)
    create_or_update_asset(INTUNE_DEVICE_DATA)

    # Modify the data to simulate an update (e.g., user changed departments)
    INTUNE_DEVICE_DATA['department'] = 'Finance'
    INTUNE_DEVICE_DATA['location'] = 'London Office - Cubicle 1'
    print("\n[Waiting 2 seconds... simulating updated Intune data]")
    sleep(2)

    # 2. Simulate subsequent run (will UPDATE the existing device)
    create_or_update_asset(INTUNE_DEVICE_DATA)

    print("\n--- Sync Test Complete ---")