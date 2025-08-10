import os
import sys
import json
import subprocess
import shutil
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

APP_TITLE = "POLK County Scraper"
DEFAULT_URL = "https://showcase.polkcountyclerk.net/showcaseweb/"
DEFAULT_DEBUG_PORT = 9222

CONFIG_DIR = Path.home() / ".polk_scraper"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"

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
        messagebox.showerror("Error", f"Failed to save config:\n{e}")

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

    seen = set()
    uniq = []
    for p in paths:
        if p and p not in seen and os.path.isfile(p):
            seen.add(p)
            uniq.append(p)
    return uniq

def detect_chrome():
    # macOS: use 'open -na' so we don't care where Chrome.app lives
    if is_macos():
        return ("mac_open", None)

    # Windows: look in registry/Program Files/PATH
    if is_windows():
        paths = candidate_paths_windows()
        if paths:
            return ("windows_path", paths[0])

    # Fallback: PATH on any OS
    exe = shutil.which("google-chrome") or shutil.which("chrome") or shutil.which("chromium")
    if exe:
        return ("path_exe", exe)

    return (None, None)

def launch_chrome(url, incognito=True, debug_port=None, status_label=None):
    url = (url or "").strip() or DEFAULT_URL
    mode, value = detect_chrome()

    try:
        if mode == "mac_open":
            args = ["open", "-na", "Google Chrome", "--args"]
            if incognito:
                args.append("--incognito")
            if debug_port:
                args.append(f"--remote-debugging-port={int(debug_port)}")
            args.append(url)
            subprocess.Popen(args)
            if status_label:
                status_label.config(text=f"Launched Chrome on macOS → {url}")
        elif mode in ("windows_path", "path_exe"):
            chrome_exe = value
            if not chrome_exe:
                messagebox.showerror("Chrome Not Found", "Could not find Google Chrome automatically.")
                return
            cmd = [chrome_exe]
            if incognito:
                cmd.append("--incognito")
            if debug_port:
                cmd.append(f"--remote-debugging-port={int(debug_port)}")
            cmd.append(url)
            subprocess.Popen(cmd)
            if status_label:
                status_label.config(text=f"Launched Chrome → {url}")
        else:
            messagebox.showwarning("Chrome Not Found", "I couldn't find Google Chrome automatically.")
            if status_label:
                status_label.config(text="Chrome not found.")
    except Exception as e:
        messagebox.showerror("Launch Failed", f"Could not launch Chrome:\n{e}")
        if status_label:
            status_label.config(text="Launch failed.")

# ---------- Address Validation integration ----------

def default_validate_candidates():
    """Return common places we might find validate_address.py"""
    candidates = []
    # Same folder as the app
    try:
        app_dir = Path(__file__).resolve().parent
        candidates.append(app_dir / "validate_address.py")
    except Exception:
        pass
    # User Documents/PolkScraper
    candidates.append(Path.home() / "Documents" / "PolkScraper" / "validate_address.py")
    return [str(p) for p in candidates]

def locate_validate_script(cfg):
    """Find the validate script using (1) saved path, (2) common defaults, else prompt."""
    # 1) saved path
    saved = cfg.get("validate_script_path")
    if saved and Path(saved).exists():
        return saved

    # 2) defaults
    for cand in default_validate_candidates():
        if Path(cand).exists():
            cfg["validate_script_path"] = cand
            save_config(cfg)
            return cand

    # 3) prompt user
    messagebox.showinfo(
        "Locate Address Validator",
        "Select your address validation script (e.g., validate_address.py, a .exe, or a .app). "
        "This will be saved for next time."
    )
    if is_macos():
        filetypes = [("Python or App", "*.py;*.app"), ("All files", "*")]
    else:
        filetypes = [("Python/Executable", "*.py;*.exe"), ("All files", "*.*")]
    path = filedialog.askopenfilename(title="Select validator", filetypes=filetypes)
    if path and Path(path).exists():
        cfg["validate_script_path"] = path
        save_config(cfg)
        return path
    return None

def run_validate_script(status_label):
    cfg = load_config()
    script_path = locate_validate_script(cfg)
    if not script_path:
        status_label.config(text="Validator not set.")
        return

    p = Path(script_path)
    try:
        if p.suffix.lower() == ".py":
            # Run with this Python environment
            subprocess.Popen([sys.executable, str(p)])
        elif p.suffix.lower() == ".exe":
            # Windows executable
            subprocess.Popen([str(p)])
        elif p.suffix.lower() == ".app" and is_macos():
            # macOS app bundle
            subprocess.Popen(["open", str(p)])
        else:
            # Try to execute directly
            subprocess.Popen([str(p)])
        status_label.config(text=f"Validator started: {p.name}")
    except Exception as e:
        messagebox.showerror("Validator Error", f"Could not run validator:\n{e}")
        status_label.config(text="Validator failed to start.")

# ---------- UI ----------

def center_window(root, width=620, height=300):
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = int((sw / 2) - (width / 2))
    y = int((sh / 2) - (height / 2))
    root.geometry(f"{width}x{height}+{x}+{y}")

def build_ui():
    root = tk.Tk()
    root.title(APP_TITLE)
    center_window(root, 620, 300)
    root.resizable(False, False)

    main = ttk.Frame(root, padding=16)
    main.pack(fill="both", expand=True)

    title_lbl = ttk.Label(main, text=APP_TITLE, font=("Segoe UI", 16, "bold"))
    title_lbl.pack(anchor="w")

    desc = ttk.Label(
        main,
        text="Launch Google Chrome to the Polk Showcase and run your Address Validator.",
        justify="left"
    )
    desc.pack(anchor="w", pady=(8, 12))

    # URL Row
    cfg = load_config()
    url_frame = ttk.Frame(main)
    url_frame.pack(fill="x", pady=(0, 6))
    ttk.Label(url_frame, text="URL:").pack(side="left")
    url_var = tk.StringVar(value=cfg.get("url", DEFAULT_URL))
    url_entry = ttk.Entry(url_frame, textvariable=url_var, width=68)
    url_entry.pack(side="left", padx=(8, 0), fill="x", expand=True)

    # Debug Port Row
    port_frame = ttk.Frame(main)
    port_frame.pack(fill="x", pady=(0, 10))
    ttk.Label(port_frame, text="Remote Debug Port:").pack(side="left")
    port_var = tk.StringVar(value=str(cfg.get("debug_port", DEFAULT_DEBUG_PORT)))
    port_entry = ttk.Entry(port_frame, textvariable=port_var, width=10)
    port_entry.pack(side="left", padx=(8, 0))

    status_lbl = ttk.Label(main, text="", foreground="#0a6")

    def save_prefs():
        c = load_config()
        c["url"] = url_var.get().strip() or DEFAULT_URL
        try:
            c["debug_port"] = int(port_var.get())
        except Exception:
            c["debug_port"] = DEFAULT_DEBUG_PORT
        save_config(c)
        status_lbl.config(text="Preferences saved.")

    # Buttons Row
    btns = ttk.Frame(main)
    btns.pack(pady=(10, 6))

    incog_btn = ttk.Button(
        btns,
        text="Launch Chrome (Incognito)",
        command=lambda: (save_prefs(),
                         launch_chrome(url_var.get(), incognito=True, debug_port=None, status_label=status_lbl))
    )
    incog_btn.grid(row=0, column=0, padx=6, ipadx=6, ipady=3)

    debug_btn = ttk.Button(
        btns,
        text="Launch Chrome (Incognito + Debug)",
        command=lambda: (save_prefs(),
                         launch_chrome(url_var.get(), incognito=True, debug_port=port_var.get(), status_label=status_lbl))
    )
    debug_btn.grid(row=0, column=1, padx=6, ipadx=6, ipady=3)

    validate_btn = ttk.Button(
        btns,
        text="Validate Addresses",
        command=lambda: run_validate_script(status_lbl)
    )
    validate_btn.grid(row=0, column=2, padx=6, ipadx=6, ipady=3)

    status_lbl.pack(anchor="w", pady=(8, 0))

    return root

if __name__ == "__main__":
    app = build_ui()
    app.mainloop()
