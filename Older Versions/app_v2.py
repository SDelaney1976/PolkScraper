# src/app.py
import os
import sys
import json
import subprocess
import shutil
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.filedialog  # ensure PyInstaller bundles filedialog
import tkinter.scrolledtext as scrolledtext
from datetime import datetime

APP_TITLE = "POLK County Scraper"
DEFAULT_URL = "https://showcase.polkcountyclerk.net/showcaseweb/"
DEFAULT_DEBUG_PORT = 9222
APP_VERSION = "v1.2"

CONFIG_DIR = Path.home() / ".polk_scraper"
CONFIG_DIR.mkdir(parents=True, exist_ok=True)
CONFIG_FILE = CONFIG_DIR / "config.json"

FROZEN = bool(getattr(sys, "frozen", False))

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
            seen.add(p)
            uniq.append(p)
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
                try:
                    messagebox.showerror("Chrome Not Found", "Could not find Google Chrome automatically.", parent=parent)
                except Exception:
                    pass
                return
            cmd = [chrome_exe]
            if incognito:
                cmd.append("--incognito")
            if debug_port:
                cmd.append(f"--remote-debugging-port={int(debug_port)}")
            cmd.append(url)
            creationflags = 0
            if is_windows():
                creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
            subprocess.Popen(cmd, creationflags=creationflags)
            if status_label:
                status_label.config(text=f"Launched Chrome → {url}")
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

# -------------------------------
# Module runners (with/without env)
# -------------------------------
def run_module_blocking(module_name: str):
    """
    Run a module so its own __main__ executes.
      Dev:    python -m <module>
      Frozen: re-invoke this executable with --run-module <module>
    Returns (exit_code, stdout, stderr)
    """
    if FROZEN:
        cmd = [sys.executable, sys.argv[0], "--run-module", module_name]
    else:
        cmd = [sys.executable, "-m", module_name]

    creationflags = 0
    if is_windows():
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    proc = subprocess.run(
        cmd,
        text=True,
        capture_output=True,
        creationflags=creationflags,
    )
    return proc.returncode, (proc.stdout or ""), (proc.stderr or "")

def run_module_blocking_with_env(module_name: str, extra_env: dict | None = None):
    """
    Same as run_module_blocking, but lets us pass environment variables
    (used to feed one file path per run).
    """
    if FROZEN:
        cmd = [sys.executable, sys.argv[0], "--run-module", module_name]
    else:
        cmd = [sys.executable, "-m", module_name]

    env = os.environ.copy()
    if extra_env:
        env.update(extra_env)
    # Make Windows happy with any non-ASCII output
    env.setdefault("PYTHONIOENCODING", "utf-8")

    creationflags = 0
    if is_windows():
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    proc = subprocess.run(
        cmd,
        text=True,
        capture_output=True,
        creationflags=creationflags,
        env=env,
    )
    return proc.returncode, (proc.stdout or ""), (proc.stderr or "")

# -------------------------------
# Report window (scrollable + autosave)
# -------------------------------
REPORT_DIR = CONFIG_DIR / "reports"
REPORT_DIR.mkdir(parents=True, exist_ok=True)

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

    frm = ttk.Frame(win, padding=12)
    frm.pack(fill="both", expand=True)

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
# UI
# -------------------------------
def center_window(root, width=640, height=420):
    root.update_idletasks()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    x = int((sw / 2) - (width / 2))
    y = int((sh / 2) - (height / 2))
    root.geometry(f"{width}x{height}+{x}+{y}")

def build_ui():
    from tkinter import filedialog

    root = tk.Tk()
    root.title(f"{APP_TITLE} — {APP_VERSION}")
    center_window(root, 680, 460)
    root.resizable(False, False)

    main = ttk.Frame(root, padding=16)
    main.pack(fill="both", expand=True)

    ttk.Label(main, text=APP_TITLE, font=("", 16, "bold")).pack(anchor="w")
    ttk.Label(
        main,
        text="Launch Chrome to Polk Showcase and run tools below.",
        justify="left"
    ).pack(anchor="w", pady=(8, 12))

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

    status_lbl = ttk.Label(main, text="", foreground="#0a6")
    status_lbl.pack(anchor="w", pady=(8, 0))

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

    # Progress bar shared by actions
    prog = ttk.Progressbar(main, mode="indeterminate")

    def set_running(running: bool):
        for child in btns.winfo_children():
            try:
                child.configure(state="disabled" if running else "normal")
            except Exception:
                pass
        if running:
            if not prog.winfo_ismapped():
                prog.pack(fill="x", pady=(0, 8))
            prog.start(50)
        else:
            try:
                prog.stop()
            except Exception:
                pass
            if prog.winfo_ismapped():
                prog.pack_forget()

    # ---------- Validate (multi-file) ----------
    def run_validate_multi():
        paths = filedialog.askopenfilenames(
            parent=root,
            title="Select file(s) to Validate",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")]
        )
        if not paths:
            status_lbl.config(text="Canceled.")
            return

        set_running(True)
        status_lbl.config(text="Validating files… this can take a while ⏳")

        def worker():
            sections = []
            ok, fail = 0, 0
            for p in paths:
                rc, out, err = run_module_blocking_with_env(
                    "validator.validate_address",
                    {"VALIDATOR_INPUT": p, "PYTHONIOENCODING": "utf-8"}
                )
                header = f"📄 {p}"
                if rc == 0:
                    ok += 1
                    body = (out or "").strip() or "Validated."
                    sections.append(f"{'='*72}\n{header}\n{'='*72}\n{body}\n")
                else:
                    fail += 1
                    details = ((err or "") + ("\n" + (out or "") if out else "")).strip()
                    if not details:
                        details = f"Exited with code {rc}"
                    sections.append(f"{'='*72}\n{header}\n{'='*72}\n❌ {details}\n")

            summary = [
                "========================================================================",
                "📊 Overall Summary — Validate",
                "========================================================================",
                f"Files selected: {len(paths)}",
                f"Succeeded: {ok}",
                f"Failed: {fail}",
                ""
            ]
            report = "\n".join(sections + ["\n"] + summary)

            def done():
                set_running(False)
                show_report_window(root,
                    "validator.validate_address — Report ✅" if fail == 0 else "validator.validate_address — Report ⚠️",
                    report, autosave=True)
                status_lbl.config(text=("Success." if fail == 0 else f"Completed with {fail} error(s)."))

            root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    ttk.Button(
        btns, text="✅ Validate Addresses (multi)",
        command=run_validate_multi
    ).pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    # ---------- Final Cleanup (multi-file) ----------
    def run_final_cleanup_multi():
        paths = filedialog.askopenfilenames(
            parent=root,
            title="Select file(s) for Final Cleanup",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if not paths:
            status_lbl.config(text="Canceled.")
            return

        set_running(True)
        status_lbl.config(text="Running Final Cleanup… ⏳")

        def worker():
            sections = []
            ok, fail = 0, 0
            for p in paths:
                rc, out, err = run_module_blocking_with_env(
                    "validator.final_scrub",
                    {"FINAL_SCRUB_INPUT": p, "PYTHONIOENCODING": "utf-8"}
                )
                header = f"📄 {p}"
                if rc == 0:
                    ok += 1
                    body = (out or "").strip() or "Cleaned."
                    sections.append(f"{'='*72}\n{header}\n{'='*72}\n{body}\n")
                else:
                    fail += 1
                    details = ((err or "") + ("\n" + (out or "") if out else "")).strip()
                    if not details:
                        details = f"Exited with code {rc}"
                    sections.append(f"{'='*72}\n{header}\n{'='*72}\n❌ {details}\n")

            summary = [
                "========================================================================",
                "📊 Overall Summary — Final Cleanup",
                "========================================================================",
                f"Files selected: {len(paths)}",
                f"Succeeded: {ok}",
                f"Failed: {fail}",
                ""
            ]
            report = "\n".join(sections + ["\n"] + summary)

            def done():
                set_running(False)
                show_report_window(root,
                    "validator.final_scrub — Report ✅" if fail == 0 else "validator.final_scrub — Report ⚠️",
                    report, autosave=True)
                status_lbl.config(text=("Success." if fail == 0 else f"Completed with {fail} error(s)."))

            root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    ttk.Button(
        btns, text="🧹 Final Cleanup (multi)",
        command=run_final_cleanup_multi
    ).pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    ttk.Button(
        btns,
        text="🕵️ Launch Chrome (Incognito)",
        command=lambda: (save_prefs(), launch_chrome(url_var.get(), True, None, status_lbl, root))
    ).pack(fill="x", pady=(0, 6), ipadx=6, ipady=3)

    ttk.Button(
        btns,
        text="❌ Exit",
        command=root.destroy
    ).pack(fill="x", ipadx=6, ipady=3)

    root.protocol("WM_DELETE_WINDOW", lambda: root.after_idle(root.destroy))
    return root

# -------------------------------
# Frozen child entry: run any module
# -------------------------------
def _run_module_entry(modname: str):
    import runpy, sys
    sys.argv = [modname]  # clean argv for child
    try:
        runpy.run_module(modname, run_name="__main__")
    except SystemExit:
        pass

# -------------------------------
# Entrypoint
# -------------------------------
if __name__ == "__main__":
    if "--run-module" in sys.argv:
        try:
            idx = sys.argv.index("--run-module")
            module = sys.argv[idx + 1]
        except Exception:
            print("Missing module name after --run-module", file=sys.stderr)
            sys.exit(2)
        _run_module_entry(module)
        sys.exit(0)

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
