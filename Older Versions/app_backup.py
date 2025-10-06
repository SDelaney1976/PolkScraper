# src/app.py
import os
import sys
import json
import subprocess
import shutil
import threading
from pathlib import Path
import socket
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.filedialog   # ensure PyInstaller bundles filedialog
import tkinter.scrolledtext as scrolledtext
from datetime import datetime

APP_TITLE = "POLK County Scraper"
APP_VERSION = "v1.2"
DEFAULT_URL = "https://showcase.polkcountyclerk.net/showcaseweb/"
DEFAULT_DEBUG_PORT = 9222

CONFIG_DIR = Path.home() / ".polk_scraper"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"

REPORT_DIR = CONFIG_DIR / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

FROZEN = bool(getattr(sys, "frozen", False))

# Path to an Output folder in the repo (works in dev & frozen builds)
if FROZEN:
    # Saved to: /Users/sdelaney/Documents/PolkScraper/Output
    repo_root = Path(sys.executable).resolve().parents[4]
else:
    # In dev, go two levels up from src/app.py to reach repo root
    repo_root = Path(__file__).resolve().parent.parent

EXPORT_DIR = repo_root / "Output"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# --- Force-bundle runtime deps used by validator/final_scrub ---
# PyInstaller only bundles modules it "sees" from the entrypoint.
# Your validator modules are run at runtime, so we import their deps here
# to force PyInstaller to include them in the app bundle.
try:
    import pandas, numpy, openpyxl, et_xmlfile  # pandas/openpyxl stack
    import requests, certifi, urllib3, chardet, idna  # requests stack
    import dateutil, dateutil.parser, dateutil.tz     # dateutil parsing
    import selenium, selenium.webdriver
    from selenium.webdriver import chrome as _chrome_mod  # keeps chrome driver pieces
    # Force-bundle capture modules
    import capture.gather_cases as _force_gather_cases     # noqa: F401
    import capture.gather_case_details as _force_details   # noqa: F401
except Exception:
    # If a dev env doesn't have all deps installed, don't crash the GUI.
    pass


# -------------------------------
# Config helpers
# -------------------------------
def load_config():
    if CONFIG_FILE.exists():
        try:
            return json.loads(CONFIG_FILE.read_text())
        except Exception:
            return {}
    return {}

def save_config(cfg):
    try:
        CONFIG_FILE.write_text(json.dumps(cfg, indent=2))
    except Exception as e:
        print("Failed to save config:", e, file=sys.stderr)

def resource_path(relpath: str) -> str:
    """Return an absolute path to resource, working in dev and in PyInstaller bundle."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        base = sys._MEIPASS
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.normpath(os.path.join(base, relpath))

# -------------------------------
# Chrome helpers
# -------------------------------
def is_windows():
    return sys.platform.startswith("win")

def is_macos():
    return sys.platform == "darwin"

def candidate_paths_windows():
    paths = []
    pf = os.environ.get("ProgramFiles", r"C:\Program Files")
    pf86 = os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")
    paths.extend([
        rf"{pf}\Google\Chrome\Application\chrome.exe",
        rf"{pf86}\Google\Chrome\Application\chrome.exe",
    ])
    found = shutil.which("chrome") or shutil.which("chrome.exe") or shutil.which("google-chrome")
    if found:
        paths.append(found)
    try:
        import winreg
        for root in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
            try:
                with winreg.OpenKey(root, r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe") as k:
                    val, _ = winreg.QueryValueEx(k, None)
                    if val:
                        paths.append(val)
            except OSError:
                pass
    except Exception:
        pass
    seen, uniq = set(), []
    for p in paths:
        if p and p not in seen and os.path.isfile(p):
            seen.add(p); uniq.append(p)
    return uniq

def detect_chrome():
    if is_macos():
        return ("mac_open", None)
    if is_windows():
        paths = candidate_paths_windows()
        if paths:
            return ("windows_path", paths[0])
    exe = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if exe:
        return ("path_exe", exe)
    return (None, None)

def launch_chrome(url, incognito=True, debug_port=None, status_label=None, parent=None):
    url = (url or "").strip() or DEFAULT_URL
    port = int(debug_port) if debug_port else DEFAULT_DEBUG_PORT
    mode, value = detect_chrome()

    try:
        if is_macos():
            # Use direct Chrome binary (more reliable than 'open -na')
            chrome_bin = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
            # Dedicated profile dir so remote debugging works even if normal Chrome is running
            user_dir   = str(CONFIG_DIR / "chrome_debug_profile")
            os.makedirs(user_dir, exist_ok=True)

            cmd = [
                chrome_bin,
                f"--remote-debugging-port={port}",
                "--remote-debugging-address=127.0.0.1",
                f"--user-data-dir={user_dir}",
            ]
            if incognito:
                cmd.append("--incognito")
            cmd.append(url)

            subprocess.Popen(cmd)
            if status_label:
                status_label.config(text=f"Launched Chrome on macOS → {url}")

        elif mode in ("windows_path", "path_exe"):
            chrome_exe = value
            if not chrome_exe:
                try:
                    messagebox.showerror("Chrome Not Found", "Could not find Google Chrome automatically.", parent=parent)
                except Exception:
                    pass
                return
            cmd = [chrome_exe]
            if incognito:
                cmd.append("--incognito")
            cmd.append(f"--remote-debugging-port={port}")
            cmd.append(url)
            creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if is_windows() else 0
            subprocess.Popen(cmd, creationflags=creationflags)
            if status_label:
                status_lbl_text = f"Launched Chrome → {url}"
                status_label.config(text=status_lbl_text)

        else:
            try:
                messagebox.showwarning("Chrome Not Found", "I couldn't find Google Chrome automatically.", parent=parent)
            except Exception:
                pass
            if status_label:
                status_label.config(text="Chrome not found.")

    except Exception as e:
        try:
            messagebox.showerror("Launch Failed", f"Could not launch Chrome:\n{e}", parent=parent)
        except Exception:
            pass
        if status_label:
            status_label.config(text="Launch failed.")


def devtools_port_alive(port: int = 9222) -> bool:
    """Quick check that something is listening on localhost:port (Chrome DevTools)."""
    try:
        with socket.create_connection(("127.0.0.1", int(port)), timeout=0.5):
            return True
    except Exception:
        return False

# -------------------------------
# Report window
# -------------------------------
def show_report_window(parent: tk.Tk, title: str, body: str, autosave: bool = True):
    saved_path = None
    if autosave:
        stamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        saved_path = REPORT_DIR / f"report_{stamp}.txt"
        try:
            saved_path.write_text(body, encoding="utf-8")
        except Exception:
            saved_path = None

    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.grab_set()
    win.geometry("820x560+120+120")
    win.minsize(560, 360)

    frm = ttk.Frame(win, padding=12); frm.pack(fill="both", expand=True)
    hdr_text = title if not saved_path else f"{title}\nSaved to: {saved_path}"
    ttk.Label(frm, text=hdr_text, font=("", 11, "bold"), justify="left").pack(anchor="w", pady=(0, 8))

    font_mono = ("Menlo", 11) if sys.platform == "darwin" else ("Consolas", 10)
    txt = scrolledtext.ScrolledText(frm, wrap="none", font=font_mono)
    txt.pack(fill="both", expand=True)
    txt.insert("1.0", body)
    txt.focus()

    btns = ttk.Frame(frm); btns.pack(fill="x", pady=(8, 0))
    def do_copy():
        try:
            parent.clipboard_clear()
            parent.clipboard_append(txt.get("1.0", "end-1c"))
        except Exception:
            pass
    def do_save_as():
        from tkinter import filedialog
        path = filedialog.asksaveasfilename(
            parent=win, title="Save Report As…",
            defaultextension=".txt",
            initialfile=(saved_path.name if saved_path else "report.txt"),
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if path:
            Path(path).write_text(txt.get("1.0", "end-1c"), encoding="utf-8")

    ttk.Button(btns, text="Copy", command=do_copy).pack(side="left")
    ttk.Button(btns, text="Save…", command=do_save_as).pack(side="left", padx=(8,0))
    ttk.Button(btns, text="Close ✖️", command=win.destroy).pack(side="right")
    win.bind("<Escape>", lambda e: win.destroy())

# -------------------------------
# Generic module runner
#   (lets the module drive its own picker; captures Unicode safely)
# -------------------------------
def run_module_blocking(module_name: str):
    """
    Run a module so its own __main__ executes (it shows its own picker).
      Dev:    python -m <module>
      Frozen: re-invoke this executable with --run-module <module>
    Returns (exit_code, stdout, stderr) with UTF-8 decode.
    """
    if FROZEN:
        cmd = [sys.executable, sys.argv[0], "--run-module", module_name]
    else:
        cmd = [sys.executable, "-m", module_name]

    env = os.environ.copy()
    # Force UTF-8 so emoji/logs never break decoding
    env.setdefault("PYTHONIOENCODING", "utf-8")
    # IMPORTANT: tell validator to suppress its own popup and exit cleanly
    if module_name == "validator.validate_address":
        env["VALIDATOR_NO_POPUPS"] = "1"

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if is_windows() else 0

    proc = subprocess.run(
        cmd,
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
        creationflags=creationflags,
        env=env,                     # <— pass our env
    )
    return proc.returncode, (proc.stdout or ""), (proc.stderr or "")

def run_capture_cases_with_stdin(start_str: str, end_str: str, case_abbr: str):
    """
    Run capture.gather_cases as a module (works in dev and frozen) and feed the three inputs via stdin.
    Returns (rc, stdout, stderr).
    """
    if FROZEN:
        cmd = [sys.executable, sys.argv[0], "--run-module", "capture.gather_cases"]
    else:
        cmd = [sys.executable, "-m", "capture.gather_cases"]

    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0) if is_windows() else 0
    proc = subprocess.Popen(
        cmd,
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        encoding="utf-8",
        errors="replace",
        creationflags=creationflags,
        env=dict(os.environ, PYTHONIOENCODING="utf-8"),
        cwd=str(EXPORT_DIR),  # <— ensure outputs are saved in Output folder
    )

    payload = f"{start_str}\n{end_str}\n{case_abbr}\n"
    out, err = proc.communicate(payload)

    # Optional: append save location to report output
    if out:
        out += f"\n\n📂 Saved to: {EXPORT_DIR}"

    return proc.returncode, (out or ""), (err or "")

# -------------------------------
# UI
# -------------------------------
def center_window(root, width=680, height=460):
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = int((sw / 2) - (width / 2))
    y = int((sh / 2) - (height / 2))
    root.geometry(f"{width}x{height}+{x}+{y}")

def _validate_mmddyyyy(s: str) -> str:
    s = s.strip()
    try:
        dt = datetime.strptime(s, "%m/%d/%Y")
    except Exception:
        raise ValueError("Use MM/DD/YYYY (e.g., 07/01/2025)")
    return dt.strftime("%m/%d/%Y")  # normalized, zero-padded

def build_ui():
    root = tk.Tk()
    root.title(f"{APP_TITLE} — {APP_VERSION}")
    center_window(root, 680, 620)
    root.resizable(True, True)

    main = ttk.Frame(root, padding=16); main.pack(fill="both", expand=True)

    ttk.Label(main, text=APP_TITLE, font=("", 16, "bold")).pack(anchor="w")
    ttk.Label(main, text="Launch Chrome or run tools below (each tool opens its own picker).",
              justify="left").pack(anchor="w", pady=(8, 12))

    cfg = load_config()

    # URL
    url_frame = ttk.Frame(main); url_frame.pack(fill="x", pady=(0, 6))
    ttk.Label(url_frame, text="URL:").pack(side="left")
    url_var = tk.StringVar(value=cfg.get("url", DEFAULT_URL))
    ttk.Entry(url_frame, textvariable=url_var, width=68).pack(side="left", padx=(8, 0), fill="x", expand=True)

    # Debug port
    port_frame = ttk.Frame(main); port_frame.pack(fill="x", pady=(0, 10))
    ttk.Label(port_frame, text="Remote Debug Port:").pack(side="left")
    port_var = tk.StringVar(value=str(cfg.get("debug_port", DEFAULT_DEBUG_PORT)))
    ttk.Entry(port_frame, textvariable=port_var, width=10).pack(side="left", padx=(8, 0))

    status_lbl = ttk.Label(main, text="", foreground="#0a6"); status_lbl.pack(anchor="w", pady=(8, 0))
    btns = ttk.Frame(main); btns.pack(pady=(10, 8), fill="x")

    def save_prefs():
        c = load_config()
        c["url"] = url_var.get().strip() or DEFAULT_URL
        try:
            c["debug_port"] = int(port_var.get())
        except Exception:
            c["debug_port"] = DEFAULT_DEBUG_PORT
        save_config(c)

    ttk.Button(
        btns,
        text="🚀 Launch Chrome (Incognito + Debug)",
        command=lambda: (save_prefs(), launch_chrome(url_var.get(), True, port_var.get(), status_lbl, root))
    ).pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    # --- Capture Cases button ---
    def open_capture_dialog():
        dlg = tk.Toplevel(root)
        dlg.title("Capture Cases")
        dlg.transient(root)
        dlg.grab_set()
        dlg.resizable(False, False)

        tk.Label(dlg, text="Start Date (MM/DD/YYYY):").grid(row=0, column=0, padx=8, pady=(10,4), sticky="w")
        start_var = tk.StringVar()
        tk.Entry(dlg, textvariable=start_var, width=16).grid(row=0, column=1, padx=8, pady=(10,4), sticky="w")

        tk.Label(dlg, text="End Date (MM/DD/YYYY):").grid(row=1, column=0, padx=8, pady=4, sticky="w")
        end_var = tk.StringVar()
        tk.Entry(dlg, textvariable=end_var, width=16).grid(row=1, column=1, padx=8, pady=4, sticky="w")

        tk.Label(dlg, text="Case Type:").grid(row=2, column=0, padx=8, pady=(8,4), sticky="w")
        case_var = tk.StringVar(value="MM")
        rb = ttk.Frame(dlg); rb.grid(row=2, column=1, padx=8, pady=(8,4), sticky="w")
        for opt in ("MM", "CT", "CF"):
            ttk.Radiobutton(rb, text=opt, value=opt, variable=case_var).pack(side="left", padx=(0,10))

        def on_run():
            # Verify Chrome DevTools is running on the expected port (default 9222)
            try:
                expected_port = int(port_var.get())
            except Exception:
                expected_port = DEFAULT_DEBUG_PORT

            if not devtools_port_alive(expected_port):
                messagebox.showwarning(
                    "Chrome Not Ready",
                    f"Chrome with remote debugging on port {expected_port} is not running.\n\n"
                    "Click 'Launch Chrome (Incognito + Debug)' first, then try again.",
                    parent=dlg
                )
                return

            try:
                s = _validate_mmddyyyy(start_var.get())
                e = _validate_mmddyyyy(end_var.get())
            except Exception as ex:
                messagebox.showerror("Invalid Date", str(ex), parent=dlg)
                return

            dlg.destroy()

            # show spinner & run on background thread (same UX as other tools)
            set_running(True)
            status_lbl.config(text="Running capture…")

            def worker():
                try:
                    rc, out, err = run_capture_cases_with_stdin(s, e, case_var.get())
                    report_title = "Capture Cases Report"
                    report_sections = [
                        "========================================================================",
                        "capture/gather_cases.py",
                        "========================================================================",
                    ]
                    if rc == 0:
                        body = (out.strip() or "Completed.")
                        report_sections.append(body)
                    else:
                        body = ((err or "") + ("\n" + (out or "") if out else "")).strip() or f"Exited with code {rc}"
                        report_sections.append(f"❌ {body}")
                    report = "\n".join(report_sections)

                    def done():
                        set_running(False)
                        title = f"{report_title} ✅" if rc == 0 else f"{report_title} ⚠️"
                        show_report_window(root, title, report, autosave=True)
                        status_lbl.config(text=("Success." if rc == 0 else f"Failed (code {rc})."))
                    root.after(0, done)
                except Exception as e2:
                    def done_err():
                        set_running(False)
                        messagebox.showerror("Error", f"Capture crashed:\n{e2}", parent=root)
                        status_lbl.config(text="Failed.")
                    root.after(0, done_err)

            threading.Thread(target=worker, daemon=True).start()

        btns2 = ttk.Frame(dlg); btns2.grid(row=3, column=0, columnspan=2, pady=10, sticky="e")
        ttk.Button(btns2, text="Cancel", command=dlg.destroy, width=10).pack(side="right", padx=6)
        ttk.Button(btns2, text="Run", command=on_run, width=10).pack(side="right", padx=6)

        dlg.bind("<Escape>", lambda e: dlg.destroy())

    # Place the new button alongside your existing ones:
    ttk.Button(btns, text="📥 Capture Cases", command=open_capture_dialog)\
        .pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    # --- Get Case Details button ---
    def run_details():
        try:
            expected_port = int(port_var.get())
        except Exception:
            expected_port = DEFAULT_DEBUG_PORT
        if not devtools_port_alive(expected_port):
            messagebox.showwarning(
                "Chrome Not Ready",
                f"Chrome with remote debugging on port {expected_port} is not running.\n\n"
                "Click 'Launch Chrome (Incognito + Debug)' first, then try again.",
                parent=root
            )
            return

        # Run as a module (same pattern as other tools)
        run_tool("capture.gather_case_details", "Get Case Details Report")

    ttk.Button(btns, text="📄 Get Case Details", command=run_details)\
        .pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    # Shared spinner
    prog = ttk.Progressbar(main, mode="indeterminate")

    def set_running(running: bool):
        for child in btns.winfo_children():
            try: child.configure(state="disabled" if running else "normal")
            except Exception: pass
        if running:
            if not prog.winfo_ismapped():
                prog.pack(fill="x", pady=(0, 8))
            prog.start(50)
        else:
            try: prog.stop()
            except Exception: pass
            if prog.winfo_ismapped():
                prog.pack_forget()

    def run_tool(module_name: str, report_title: str):
        """Run a module once (it shows its own picker), capture all output, show a report, and ALWAYS stop spinner."""
        set_running(True)
        status_lbl.config(text=f"Running {module_name}…")

        def worker():
            try:
                rc, out, err = run_module_blocking(module_name)
                report_sections = []
                report_sections.append("========================================================================")
                report_sections.append(module_name)
                report_sections.append("========================================================================")
                if rc == 0:
                    body = (out.strip() or "Completed.")
                    report_sections.append(body)
                else:
                    body = ((err or "") + ("\n" + (out or "") if out else "")).strip()
                    if not body:
                        body = f"Exited with code {rc}"
                    report_sections.append(f"❌ {body}")

                report = "\n".join(report_sections)

                def done():
                    set_running(False)
                    title = f"{report_title} ✅" if rc == 0 else f"{report_title} ⚠️"
                    show_report_window(root, title, report, autosave=True)
                    status_lbl.config(text=("Success." if rc == 0 else f"Failed (code {rc})."))

                root.after(0, done)
            except Exception as e:
                # Make SURE we stop spinner even on unexpected errors
                def done_err():
                    set_running(False)
                    messagebox.showerror("Error", f"{module_name} crashed:\n{e}", parent=root)
                    status_lbl.config(text="Failed.")
                root.after(0, done_err)

        threading.Thread(target=worker, daemon=True).start()

    # Buttons: each module drives its own multi-file picker inside itself
    ttk.Button(btns, text="✅ Validate Addresses", command=lambda: run_tool("validator.validate_address", "Validate Report"))\
        .pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    ttk.Button(btns, text="🧹 Final Cleanup", command=lambda: run_tool("validator.final_scrub", "Final Cleanup Report"))\
        .pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    ttk.Button(btns, text="🕵️ Launch Chrome (Incognito)",
               command=lambda: (save_prefs(), launch_chrome(url_var.get(), True, None, status_lbl, root)))\
        .pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    ttk.Button(btns, text="❌ Exit", command=root.destroy).pack(fill="x", ipadx=6, ipady=3)

    root.protocol("WM_DELETE_WINDOW", lambda: root.after_idle(root.destroy))
    return root

# -------------------------------
# Frozen child entry: run any module
# -------------------------------
def _run_module_entry(modname: str):
    import runpy
    # give child a clean argv so your scripts only see their own name
    sys.argv = [modname]
    try:
        runpy.run_module(modname, run_name="__main__")
    except SystemExit:
        pass

# -------------------------------
# Entrypoint
# -------------------------------
if __name__ == "__main__":
    # Run a module (used by validate/final_scrub and capture.*)
    if "--run-module" in sys.argv:
        try:
            idx = sys.argv.index("--run-module")
            module = sys.argv[idx + 1]
        except Exception:
            print("Missing module name after --run-module", file=sys.stderr)
            sys.exit(2)
        _run_module_entry(module)
        sys.exit(0)

    # Normal GUI startup
    try:
        app = build_ui()
        app.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            tk.Tk().withdraw()
            messagebox.showerror("Startup Error", f"The app failed to start:\n{e}")
        except Exception:
            pass
        sys.exit(1)
