# src/validator/validate_address.py
import os
import sys
import time
import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
import tkinter.scrolledtext as scrolledtext

# === CONFIGURATION ===
GOOGLE_API_KEY = "AIzaSyDmVcjPphH6eG_Bx35tzWhZDoSyF8EH67o"  # <-- put your real key here

# === Safe print for Windows console (avoids UnicodeEncodeError) ===
def safe_print(*args, **kwargs):
    try:
        print(*args, **kwargs)
    except UnicodeEncodeError:
        txt = " ".join(map(str, args))
        enc = sys.stdout.encoding or "utf-8"
        sys.stdout.write(txt.encode(enc, errors="replace").decode(enc, "replace") + "\n")

# === Address helpers ===
def build_full_address(row):
    parts = [
        str(row.get("Address 1") or "").strip(),
        str(row.get("Address 2") or "").strip(),
        str(row.get("City") or "").strip(),
        str(row.get("State") or "").strip(),
        str(row.get("Zip") or "").strip()
    ]
    return ", ".join([p for p in parts if p])

def get_component(components, type_name, use_short_name=False):
    for comp in components:
        if type_name in comp["types"]:
            return comp["short_name"] if use_short_name else comp["long_name"]
    return ""

def validate_and_parse_address(address, api_key):
    try:
        url = "https://maps.googleapis.com/maps/api/geocode/json"
        params = {"address": address, "key": api_key}
        r = requests.get(url, params=params, timeout=20)
        r.raise_for_status()
        result = r.json()
        if result.get("status") == "OK" and result.get("results"):
            components = result["results"][0]["address_components"]
            street_number = get_component(components, "street_number")
            route = get_component(components, "route")
            address1 = f"{street_number} {route}".strip()
            city = get_component(components, "locality")
            state = get_component(components, "administrative_area_level_1", use_short_name=True)
            zip_code = get_component(components, "postal_code")
            return address1, city, state, zip_code, None
        else:
            return ("ERROR", "ERROR", "ERROR", "ERROR", result.get("status") or "Unknown status")
    except Exception as e:
        return ("ERROR", "ERROR", "ERROR", "ERROR", str(e))

# === Core processing ===
def process_file(input_file, api_key):
    """
    Returns dict:
      {
        "file": <input_file>,
        "ok": True/False,
        "validated_rows": int,
        "duplicates_removed": int,
        "errors": int,
        "error_samples": [str,...],
        "output": <output_file or None>,
        "message": <error message or None>
      }
    """
    info = {
        "file": input_file,
        "ok": False,
        "validated_rows": 0,
        "duplicates_removed": 0,
        "errors": 0,
        "error_samples": [],
        "output": None,
        "message": None,
    }

    try:
        df = pd.read_excel(input_file)

        # Ensure expected headers (C–K) exist / are named
        expected_headers = [
            "Name", "Address 1", "City", "State", "Zip",
            "Sex", "Race", "Public Defender", "Capture Date"
        ]
        while len(df.columns) < 11:
            df[f"Unnamed {len(df.columns)+1}"] = ""
        for i, header in enumerate(expected_headers):
            df.columns.values[2 + i] = header

        # Filter out General Delivery / Homeless
        df = df[~df["Address 1"].astype(str).str.contains("General Delivery|Homeless", case=False, na=False)]

        df["Zip"] = df["Zip"].astype(str)
        original_rows = len(df)

        error_count = 0
        error_examples = []

        for idx, row in df.iterrows():
            full_address = build_full_address(row)
            safe_print(f"[{idx+1}/{original_rows}] Validating: {full_address}")
            address1, city, state, zip_code, err = validate_and_parse_address(full_address, api_key)

            df.at[idx, "Address 1"] = address1
            df.at[idx, "City"] = city
            df.at[idx, "State"] = state
            df.at[idx, "Zip"] = str(zip_code)

            if err:
                error_count += 1
                if len(error_examples) < 5:
                    error_examples.append(f"{full_address} -> {err}")

            # light pacing
            time.sleep(0.2)

        # Deduplicate
        dedup_df = df.drop_duplicates(subset=["Name", "Address 1", "City", "State", "Zip"])
        dedup_removed = original_rows - len(dedup_df)

        base, ext = os.path.splitext(input_file)
        output_file = f"{base}_validated{ext}"
        dedup_df.to_excel(output_file, index=False)

        info.update({
            "ok": True,
            "validated_rows": len(dedup_df.index),
            "duplicates_removed": dedup_removed,
            "errors": error_count,
            "error_samples": error_examples,
            "output": output_file,
        })
        return info

    except Exception as e:
        info["message"] = str(e)
        return info

# === Report UI ===
def show_report_window(root, title, body):
    win = tk.Toplevel(root)
    win.title(title)
    win.geometry("780x560+120+120")
    win.minsize(560, 360)
    win.transient(root)
    win.grab_set()

    frm = tk.Frame(win, padx=12, pady=12)
    frm.pack(fill="both", expand=True)

    lbl = tk.Label(frm, text=title, font=("Segoe UI", 12, "bold"), anchor="w", justify="left")
    lbl.pack(fill="x", pady=(0, 8))

    txt = scrolledtext.ScrolledText(frm, wrap="none", font=("Menlo", 11) if sys.platform == "darwin" else ("Consolas", 10))
    txt.pack(fill="both", expand=True)
    txt.insert("1.0", body)
    txt.mark_set("insert", "1.0")
    txt.focus()

    btns = tk.Frame(frm)
    btns.pack(fill="x", pady=(8, 0))

    def do_copy():
        try:
            root.clipboard_clear()
            root.clipboard_append(txt.get("1.0", "end-1c"))
        except Exception:
            pass

    def do_save_as():
        path = filedialog.asksaveasfilename(
            parent=win,
            title="Save Report As…",
            defaultextension=".txt",
            initialfile="validation_report.txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            with open(path, "w", encoding="utf-8") as f:
                f.write(txt.get("1.0", "end-1c"))

    tk.Button(btns, text="Copy to Clipboard", command=do_copy).pack(side="left")
    tk.Button(btns, text="Save as…", command=do_save_as).pack(side="left", padx=(8, 0))
    tk.Button(btns, text="Close", command=win.destroy).pack(side="right")

    win.bind("<Escape>", lambda e: win.destroy())

# === Main ===
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()

    files = filedialog.askopenfilenames(
        title="Select Excel file(s) to validate",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not files:
        safe_print("No files selected. Exiting.")
        sys.exit(0)

    results = []
    for path in files:
        safe_print("\n" + "=" * 72)
        safe_print(f"Processing: {path}")
        info = process_file(path, GOOGLE_API_KEY)
        results.append(info)

    # Build one combined report
    lines = []
    lines.append("=" * 72)
    lines.append("📊 Validation Summary")
    lines.append("=" * 72)
    ok_count = sum(1 for r in results if r["ok"])
    fail_count = len(results) - ok_count
    lines.append(f"Files selected: {len(results)}  |  OK: {ok_count}  |  Failed: {fail_count}")
    lines.append("")

    for r in results:
        lines.append("-" * 72)
        lines.append(os.path.basename(r["file"]))
        if r["ok"]:
            lines.append(f"  ✔ Validated rows: {r['validated_rows']}")
            lines.append(f"  🧹 Duplicates removed: {r['duplicates_removed']}")
            lines.append(f"  💾 Output: {r['output']}")
            if r["errors"]:
                lines.append(f"  ⚠️ API/address errors: {r['errors']}")
                if r["error_samples"]:
                    lines.append("  Examples:")
                    for ex in r["error_samples"]:
                        lines.append(f"    - {ex}")
        else:
            lines.append("  ❌ FAILED")
            if r.get("message"):
                lines.append(f"  Reason: {r['message']}")
        lines.append("")

    report_text = "\n".join(lines)
    # Also print to stdout (the launcher can capture if needed)
    safe_print("\n" + report_text + "\n")

    # Show one scrollable window with the whole summary
    show_report_window(root, "Validation Summary", report_text)
    root.mainloop()
