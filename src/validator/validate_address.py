# AIzaSyDmVcjPphH6eG_Bx35tzWhZDoSyF8EH67o
import pandas as pd
import requests
import time
import os
import datetime
import tkinter as tk
from tkinter import filedialog

# === CONFIGURATION ===
GOOGLE_API_KEY = "AIzaSyDmVcjPphH6eG_Bx35tzWhZDoSyF8EH67o"  # Replace with your actual API key

# === Address helpers ===
def build_full_address(row):
    parts = [
        str(row.get("Address 1") or "").strip(),
        str(row.get("Address 2") or "").strip(),
        str(row.get("City") or "").strip(),
        str(row.get("State") or "").strip(),
        str(row.get("Zip") or "").strip()
    ]
    return ', '.join([part for part in parts if part])

def get_component(components, type_name, use_short_name=False):
    for comp in components:
        if type_name in comp["types"]:
            return comp["short_name"] if use_short_name else comp["long_name"]
    return ""

def validate_and_parse_address(address, api_key):
    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {"address": address, "key": api_key}
        response = requests.get(url, params=params)
        result = response.json()

        if result["status"] == "OK":
            components = result["results"][0]["address_components"]
            street_number = get_component(components, "street_number")
            route = get_component(components, "route")
            address1 = f"{street_number} {route}".strip()
            city = get_component(components, "locality")
            state = get_component(components, "administrative_area_level_1", use_short_name=True)
            zip_code = get_component(components, "postal_code")
            return address1, city, state, zip_code
        else:
            return ("ERROR",) * 4
    except Exception as e:
        return (f"ERROR: {str(e)}",) * 4

# === Main process ===
def process_addresses(input_file, api_key):
    df = pd.read_excel(input_file)

    # ✅ Ensure expected headers in columns C–K
    expected_headers = [
        "Name", "Address 1", "City", "State", "Zip",
        "Sex", "Race", "Public Defender", "Capture Date"
    ]
    while len(df.columns) < 11:
        df[f"Unnamed {len(df.columns)+1}"] = ""
    for i, header in enumerate(expected_headers):
        df.columns.values[2 + i] = header

    # ✅ Remove rows with "General Delivery" or "Homeless" in Address 1
    df = df[~df["Address 1"].astype(str).str.contains("General Delivery|Homeless", case=False, na=False)]

    df["Zip"] = df["Zip"].astype(str)
    original_row_count = len(df)

    for index, row in df.iterrows():
        full_address = build_full_address(row)
        print(f"[{index+1}/{original_row_count}] Validating: {full_address}")
        address1, city, state, zip_code = validate_and_parse_address(full_address, api_key)

        df.at[index, "Address 1"] = address1
        df.at[index, "City"] = city
        df.at[index, "State"] = state
        df.at[index, "Zip"] = str(zip_code)

        time.sleep(0.2)

    # ✅ Remove duplicates
    dedup_df = df.drop_duplicates(subset=["Name", "Address 1", "City", "State", "Zip"])
    dedup_count = original_row_count - len(dedup_df)
    print(f"\n🧹 Removed {dedup_count} duplicate row(s).")

    # ✅ Save
    base, ext = os.path.splitext(input_file)
    output_file = f"{base}_validated{ext}"
    dedup_df.to_excel(output_file, index=False)
    print(f"✅ Validation complete. Output saved to: {output_file}")

# === File picker and run ===
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()
    input_file = filedialog.askopenfilename(
        title="Select Excel file to validate",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not input_file:
        print("❌ No file selected. Exiting.")
    else:
        process_addresses(input_file, GOOGLE_API_KEY)
