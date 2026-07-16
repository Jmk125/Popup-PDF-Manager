"""
PDF Merge Popup — hotkey-summoned PDF merge utility
====================================================
Press Ctrl+Alt+M anywhere in Windows to pop up a window showing all
currently-open PDFs. Drag (or double-click) them into the merge queue,
set "All" or a page range per file, reorder, then Merge & Save.

Dependencies:
    pip install pypdf keyboard psutil

Build to exe:
    pyinstaller --noconsole --onefile --name PDFMergePopup pdf_merge_popup.py

Detection strategy (three tiers):
  1. Enumerate top-level window titles containing ".pdf" -> filenames
  2. Resolve full paths via open file handles on known PDF viewer
     processes (Acrobat, Bluebeam Revu, Foxit, Edge, Chrome, SumatraPDF)
  3. Fall back to Windows Recent Items (.lnk) parsing
  Unresolved files appear grayed out with a right-click "Locate..." option.
"""

import os
import re
import sys
import time
import queue
import ctypes
import threading
import traceback

try:
    import tkinter as tk
    from tkinter import filedialog, messagebox
except ImportError:  # headless environment; logic functions still importable
    tk = None

IS_WINDOWS = sys.platform.startswith("win")

# ---------------------------------------------------------------------------
# Theme (dark, matches the drawing overlay tool)
# ---------------------------------------------------------------------------
BG        = "#1e1e1e"   # window background
PANEL     = "#252526"   # panel background
CARD      = "#2d2d30"   # list item / card background
CARD_HOT  = "#3e3e42"   # hover
ACCENT    = "#0e639c"   # primary blue
ACCENT_HI = "#1177bb"
GREEN     = "#4ec9b0"
TEXT      = "#d4d4d4"
TEXT_DIM  = "#808080"
BORDER    = "#3e3e42"
DANGER    = "#c0504d"

FONT       = ("Segoe UI", 10)
FONT_BOLD  = ("Segoe UI", 10, "bold")
FONT_SMALL = ("Segoe UI", 8)
FONT_TITLE = ("Segoe UI", 12, "bold")

HOTKEY = "ctrl+alt+m"

# Known PDF viewer process names (lowercase) for handle scanning.
# NOTE: browsers (msedge/chrome/firefox) are deliberately excluded — they hold
# tens of thousands of handles and psutil.open_files() takes minutes on them.
# Browser PDFs still show up via window-title detection (usually unresolved).
PDF_VIEWER_PROCS = {
    "acrobat.exe", "acrord32.exe", "acrord64.exe",
    "revu.exe", "revux64.exe", "bluebeam revu.exe",
    "foxitpdfreader.exe", "foxitreader.exe", "foxitphantompdf.exe",
    "sumatrapdf.exe", "nitropdf.exe", "pdfxedit.exe",
}

# Soft time budget for the whole handle scan (seconds)
HANDLE_SCAN_BUDGET = 4.0

# ---------------------------------------------------------------------------
# Debug logging — writes pdfmerge_debug.log next to the script/exe
# ---------------------------------------------------------------------------
def _app_dir():
    if getattr(sys, "frozen", False):          # PyInstaller exe
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


LOG_PATH = os.path.join(_app_dir(), "pdfmerge_debug.log")
_log_lock = threading.Lock()


def log(msg):
    """Write diagnostics without relying on a console being attached.

    PyInstaller's ``--noconsole`` builds set ``sys.stdout`` to ``None``.  A
    regular ``print`` therefore raises during startup and terminates the app
    before tkinter is even created.
    """
    line = f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}"
    try:
        with _log_lock, open(LOG_PATH, "a", encoding="utf-8") as fh:
            fh.write(line + "\n")
    except OSError:
        pass

    stream = getattr(sys, "stdout", None)
    if stream is not None:
        try:
            stream.write(line + "\n")
            stream.flush()
        except (AttributeError, OSError):
            pass


def _log_unhandled_exception(kind, value, tb, thread_name=None):
    """Record otherwise-invisible crashes from a windowed executable."""
    where = f" in thread {thread_name}" if thread_name else ""
    log(f"UNHANDLED EXCEPTION{where}:\n" +
        "".join(traceback.format_exception(kind, value, tb)))


def install_exception_logging():
    """Log main- and worker-thread crashes before Python reports them."""
    def main_hook(kind, value, tb):
        _log_unhandled_exception(kind, value, tb)

    def thread_hook(args):
        _log_unhandled_exception(args.exc_type, args.exc_value, args.exc_traceback,
                                 args.thread.name)

    sys.excepthook = main_hook
    if hasattr(threading, "excepthook"):
        threading.excepthook = thread_hook


# ---------------------------------------------------------------------------
# Win32 prototypes — explicit argtypes/restype (64-bit safe)
# ---------------------------------------------------------------------------
if IS_WINDOWS:
    from ctypes import wintypes

    _user32 = ctypes.WinDLL("user32", use_last_error=True)
    WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND,
                                     wintypes.LPARAM)
    _user32.EnumWindows.argtypes = [WNDENUMPROC, wintypes.LPARAM]
    _user32.EnumWindows.restype = wintypes.BOOL
    _user32.IsWindowVisible.argtypes = [wintypes.HWND]
    _user32.IsWindowVisible.restype = wintypes.BOOL
    _user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    _user32.GetWindowTextLengthW.restype = ctypes.c_int
    _user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR,
                                       ctypes.c_int]
    _user32.GetWindowTextW.restype = ctypes.c_int
    _user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND,
                                                 ctypes.POINTER(wintypes.DWORD)]
    _user32.GetWindowThreadProcessId.restype = wintypes.DWORD


# ---------------------------------------------------------------------------
# Page range parsing  (e.g. "1-3, 5, 9-" for a 12-page doc -> [0,1,2,4,8..11])
# ---------------------------------------------------------------------------
def parse_page_range(spec: str, num_pages: int):
    """Parse a 1-based page range spec into a sorted list of 0-based indices.
    Supports: "3", "1-4", "6-" (to end), "-3" (from start), comma-separated.
    Raises ValueError with a friendly message on bad input."""
    spec = spec.strip()
    if not spec or spec.lower() == "all":
        return list(range(num_pages))
    indices = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d*)\s*-\s*(\d*)", part)
        if m and (m.group(1) or m.group(2)):
            start = int(m.group(1)) if m.group(1) else 1
            end = int(m.group(2)) if m.group(2) else num_pages
        elif re.fullmatch(r"\d+", part):
            start = end = int(part)
        else:
            raise ValueError(f'Bad range "{part}" — use forms like 1-3, 5, 9-')
        if start < 1 or end > num_pages or start > end:
            raise ValueError(
                f'Range "{part}" out of bounds (document has {num_pages} pages)')
        indices.extend(range(start - 1, end))
    # preserve order given, drop duplicates
    seen, out = set(), []
    for i in indices:
        if i not in seen:
            seen.add(i)
            out.append(i)
    return out


# ---------------------------------------------------------------------------
# Open-PDF detection (Windows only; harmless no-ops elsewhere)
# ---------------------------------------------------------------------------
def _enum_window_pdf_titles():
    """Return PDF filenames and owning process IDs from visible window titles."""
    names = set()
    raw_titles = []
    pids = set()
    if not IS_WINDOWS:
        return names, pids

    def cb(hwnd, _):
        try:
            if not _user32.IsWindowVisible(hwnd):
                return True
            length = _user32.GetWindowTextLengthW(hwnd)
            if length == 0:
                return True
            buf = ctypes.create_unicode_buffer(length + 1)
            _user32.GetWindowTextW(hwnd, buf, length + 1)
            title = buf.value
            if ".pdf" not in title.lower():
                return True
            raw_titles.append(title)
            pid = wintypes.DWORD()
            _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value:
                pids.add(pid.value)
            for m in re.finditer(r"([^\\/:*?\"<>|\r\n]+?\.pdf)", title, re.I):
                name = m.group(1).strip().strip("-–— ").strip()
                name = re.sub(r"^.*?[-–—]\s+(?=[^-]*\.pdf$)", "", name,
                              flags=re.I).strip()
                if name.lower().endswith(".pdf"):
                    names.add(name)
        except Exception as e:
            log(f"title cb error on hwnd {hwnd}: {e!r}")
        return True  # never stop enumeration

    ok = _user32.EnumWindows(WNDENUMPROC(cb), 0)
    if not ok:
        log(f"EnumWindows returned FALSE, GetLastError="
            f"{ctypes.get_last_error()}")
    log(f"tier1 titles: raw={raw_titles!r} -> parsed={sorted(names)!r}, "
        f"pids={sorted(pids)!r}")
    return names, pids


# Once a path is resolved during this session, remember it — psutil's
# open_files() is transiently flaky (random AccessDenied) so later scans
# can fall back on what earlier scans learned.
_SESSION_PATH_CACHE = {}


def _is_pdf_viewer_proc(name: str) -> bool:
    name = (name or "").lower()
    if name in PDF_VIEWER_PROCS:
        return True
    # Bluebeam ships as Revu.exe / Revu21.exe / RevuX64.exe depending on version
    return "revu" in name or "bluebeam" in name


def _pdfs_from_proc(proc):
    """Open .pdf handles for one process. Raises psutil errors upward."""
    found = {}
    for f in proc.open_files():
        if f.path.lower().endswith(".pdf"):
            found[os.path.basename(f.path).lower()] = f.path
    return found


def _paths_from_process_handles(extra_pids=None):
    """Return dict {filename_lower: full_path} from PDF viewers' open handles.
    Bounded by HANDLE_SCAN_BUDGET so it can never hang the app."""
    paths = {}
    if not IS_WINDOWS:
        return paths
    try:
        import psutil
    except ImportError:
        log("tier2: psutil not installed")
        return paths
    deadline = time.monotonic() + HANDLE_SCAN_BUDGET
    extra_pids = set(extra_pids or ())
    scanned = []
    for proc in psutil.process_iter(["name", "pid"]):
        if time.monotonic() > deadline:
            log("tier2: hit time budget, stopping early")
            break
        if (proc.info["pid"] not in extra_pids and
                not _is_pdf_viewer_proc(proc.info["name"])):
            continue
        scanned.append(f"{proc.info['name']}[{proc.info['pid']}]")
        try:
            paths.update(_pdfs_from_proc(proc))
        except psutil.AccessDenied:
            log(f"tier2: {proc.info['name']} AccessDenied — retrying "
                "with fresh handle")
        except (psutil.NoSuchProcess, OSError) as e:
            log(f"tier2: {proc.info['name']} -> {type(e).__name__}")
            continue
        # retry once with a brand-new Process object (fresh OpenProcess)
        try:
            time.sleep(0.25)
            paths.update(_pdfs_from_proc(psutil.Process(proc.info["pid"])))
            log("tier2: retry succeeded")
        except Exception as e:
            log(f"tier2: retry failed -> {type(e).__name__}")
        # Last resort: PDFs passed on the command line (double-click opens).
        # Do this even after a successful open_files() call because some PDF
        # viewers stop exposing file handles after save/merge operations while
        # still keeping the source document path in their command line.
        try:
            for arg in proc.cmdline():
                if arg.lower().endswith(".pdf") and os.path.isfile(arg):
                    paths[os.path.basename(arg).lower()] = arg
        except Exception as e:
            log(f"tier2: cmdline fallback -> {type(e).__name__}")
    if paths:
        _SESSION_PATH_CACHE.update(paths)
    log(f"tier2 handles: scanned={scanned!r} -> {list(paths.values())!r}")
    return paths


def _paths_from_recent_items():
    """Parse %APPDATA%\\Microsoft\\Windows\\Recent .lnk files for pdf targets.
    Lightweight shell-link parse: scan the binary for drive-letter paths."""
    paths = {}
    if not IS_WINDOWS:
        return paths
    recent = os.path.join(os.environ.get("APPDATA", ""),
                          "Microsoft", "Windows", "Recent")
    if not os.path.isdir(recent):
        return paths
    lnks = [os.path.join(recent, f) for f in os.listdir(recent)
            if f.lower().endswith(".pdf.lnk")]
    # newest first so the most recent duplicate wins
    lnks.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    for lnk in lnks:
        try:
            with open(lnk, "rb") as fh:
                data = fh.read()
        except OSError:
            continue
        target = None
        # try UTF-16 encoded path first (LinkInfo unicode / string data)
        for encoding in ("utf-16-le", "latin-1"):
            try:
                text = data.decode(encoding, errors="ignore")
            except Exception:
                continue
            m = re.search(r"[A-Za-z]:\\[^\x00:*?\"<>|]+?\.pdf", text, re.I)
            if m:
                target = m.group(0)
                break
        if target and os.path.isfile(target):
            key = os.path.basename(target).lower()
            paths.setdefault(key, target)
    return paths


def detect_open_pdfs():
    """Return list of dicts: {name, path (or None), source}.
    Each tier is isolated — one failing can't blank the others."""
    log("=== scan start ===")
    try:
        titles, title_pids = _enum_window_pdf_titles()
    except Exception:
        import traceback
        log("tier1 CRASH:\n" + traceback.format_exc())
        titles = set()
        title_pids = set()
    try:
        handle_paths = _paths_from_process_handles(title_pids)
    except Exception:
        import traceback
        log("tier2 CRASH:\n" + traceback.format_exc())
        handle_paths = {}
    try:
        recent_paths = _paths_from_recent_items()
        log(f"tier3 recent: {len(recent_paths)} pdf lnk target(s)")
    except Exception:
        import traceback
        log("tier3 CRASH:\n" + traceback.format_exc())
        recent_paths = {}

    results = []
    for name in sorted(titles, key=str.lower):
        key = name.lower()
        path = (handle_paths.get(key) or _SESSION_PATH_CACHE.get(key) or
                recent_paths.get(key))
        results.append({"name": name, "path": path, "section": "open",
                        "source": "handles" if key in handle_paths
                                  else "session" if key in _SESSION_PATH_CACHE
                                  else "recent" if key in recent_paths
                                  else None})
    # Also surface handle-detected PDFs whose window title didn't parse
    listed = {r["name"].lower() for r in results}
    for key, path in handle_paths.items():
        if key not in listed:
            results.append({"name": os.path.basename(path), "path": path,
                            "section": "open", "source": "handles"})
            listed.add(key)

    # Recent section: session cache first (was open earlier this run),
    # then Recent Items shortcuts — anything not already listed above.
    open_paths = {r["path"] for r in results if r["path"]}
    recent_count = 0
    for key, path in list(_SESSION_PATH_CACHE.items()) + \
                     list(recent_paths.items()):
        if recent_count >= 8:
            break
        if key in listed or path in open_paths or not os.path.isfile(path):
            continue
        results.append({"name": os.path.basename(path), "path": path,
                        "section": "recent", "source": "recent"})
        listed.add(key)
        open_paths.add(path)
        recent_count += 1

    n_open = sum(1 for r in results if r["section"] == "open")
    log(f"=== scan done: {n_open} open, {recent_count} recent ===")
    return results


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------
_TkFrame = tk.Frame if tk else object


class QueueRow(_TkFrame):
    """One PDF in the merge queue: name, All/pages toggle, remove, drag-reorder."""

    def __init__(self, master, app, path):
        super().__init__(master, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1)
        self.app = app
        self.path = path
        self.num_pages = self._count_pages(path)
        self.mode = tk.StringVar(value="all")
        self.range_var = tk.StringVar()

        grip = tk.Label(self, text="⠿", bg=CARD, fg=TEXT_DIM, font=FONT,
                        cursor="fleur", padx=6)
        grip.pack(side="left", fill="y")

        body = tk.Frame(self, bg=CARD)
        body.pack(side="left", fill="both", expand=True, padx=(0, 4), pady=4)

        name = os.path.basename(path)
        pages_txt = f"{self.num_pages} pgs" if self.num_pages else "?"
        tk.Label(body, text=name, bg=CARD, fg=TEXT, font=FONT_BOLD,
                 anchor="w").pack(fill="x")
        sub = tk.Frame(body, bg=CARD)
        sub.pack(fill="x", pady=(2, 0))
        tk.Label(sub, text=pages_txt, bg=CARD, fg=TEXT_DIM,
                 font=FONT_SMALL).pack(side="left", padx=(0, 8))

        self.rb_all = tk.Radiobutton(
            sub, text="All", variable=self.mode, value="all",
            bg=CARD, fg=TEXT, selectcolor=PANEL, activebackground=CARD,
            activeforeground=TEXT, font=FONT_SMALL, command=self._mode_changed)
        self.rb_all.pack(side="left")
        self.rb_rng = tk.Radiobutton(
            sub, text="Pages:", variable=self.mode, value="range",
            bg=CARD, fg=TEXT, selectcolor=PANEL, activebackground=CARD,
            activeforeground=TEXT, font=FONT_SMALL, command=self._mode_changed)
        self.rb_rng.pack(side="left", padx=(6, 2))
        self.entry = tk.Entry(sub, textvariable=self.range_var, width=12,
                              bg=PANEL, fg=TEXT, insertbackground=TEXT,
                              relief="flat", font=FONT_SMALL,
                              disabledbackground=CARD, state="disabled")
        self.entry.pack(side="left")
        self.entry.bind("<FocusIn>", lambda e: None)

        rm = tk.Label(self, text="✕", bg=CARD, fg=TEXT_DIM, font=FONT,
                      cursor="hand2", padx=8)
        rm.pack(side="right", fill="y")
        rm.bind("<Button-1>", lambda e: self.app.remove_row(self))
        rm.bind("<Enter>", lambda e: rm.config(fg=DANGER))
        rm.bind("<Leave>", lambda e: rm.config(fg=TEXT_DIM))

        # drag-to-reorder bindings on the grip and title area
        for w in (grip, body):
            w.bind("<ButtonPress-1>", self._drag_start)
            w.bind("<B1-Motion>", self._drag_motion)
            w.bind("<ButtonRelease-1>", self._drag_end)

    def _count_pages(self, path):
        try:
            from pypdf import PdfReader
            return len(PdfReader(path).pages)
        except Exception:
            return 0

    def _mode_changed(self):
        if self.mode.get() == "range":
            self.entry.config(state="normal")
            self.entry.focus_set()
        else:
            self.entry.config(state="disabled")

    # -- reorder drag ------------------------------------------------------
    def _drag_start(self, event):
        self._drag_y = event.y_root
        self.config(highlightbackground=ACCENT)

    def _drag_motion(self, event):
        self.app.reorder_drag(self, event.y_root)

    def _drag_end(self, event):
        self.config(highlightbackground=BORDER)
        self.app.reorder_commit()

    def get_page_indices(self):
        if self.mode.get() == "all":
            return list(range(self.num_pages))
        return parse_page_range(self.range_var.get(), self.num_pages)


class MergeApp:
    def __init__(self, root):
        self.root = root
        self.rows = []          # ordered QueueRow list
        self.detected = []      # detect_open_pdfs() results
        self._drag_source = None  # left-list drag state
        self._scan_queue = queue.Queue()  # worker -> tk thread results
        self._scan_gen = 0
        self._build_ui()
        self._poll_scan()       # runs forever on the tk thread

    def _poll_scan(self):
        """Drain scan results posted by worker threads. This is the only
        thread-safe way to get data back into tkinter — workers must never
        touch widgets or call root.after themselves."""
        try:
            while True:
                gen, results, err = self._scan_queue.get_nowait()
                self._refresh_done(gen, results, err)
        except queue.Empty:
            pass
        self.root.after(150, self._poll_scan)

    # -- UI construction ---------------------------------------------------
    def _build_ui(self):
        r = self.root
        r.title(f"PDF Merge  —  {HOTKEY.upper()}")
        r.configure(bg=BG)
        r.geometry("760x480")
        r.minsize(620, 380)
        r.attributes("-topmost", True)
        r.protocol("WM_DELETE_WINDOW", self.hide)
        r.bind("<Escape>", lambda e: self.hide())

        header = tk.Frame(r, bg=BG)
        header.pack(fill="x", padx=12, pady=(10, 6))
        tk.Label(header, text="PDF MERGE", bg=BG, fg=GREEN,
                 font=FONT_TITLE).pack(side="left")
        tk.Label(header, text=f"   Hotkey: {HOTKEY.upper()}  ·  Esc hides",
                 bg=BG, fg=TEXT, font=FONT).pack(
                     side="left", pady=(3, 0))

        main = tk.Frame(r, bg=BG)
        main.pack(fill="both", expand=True, padx=12, pady=(0, 8))
        main.columnconfigure(0, weight=1, uniform="col")
        main.columnconfigure(1, weight=1, uniform="col")
        main.rowconfigure(1, weight=1)

        # ---- left: open PDFs ----
        lt = tk.Frame(main, bg=BG)
        lt.grid(row=0, column=0, sticky="ew")
        tk.Label(lt, text="OPEN PDFs", bg=BG, fg=TEXT_DIM,
                 font=FONT_SMALL).pack(side="left")
        self._mini_btn(lt, "⟳ Refresh", self.refresh).pack(side="right")
        self._mini_btn(lt, "+ Add PDF…", self.add_manual).pack(
            side="right", padx=(0, 6))
        self._mini_btn(lt, "Log", self.open_log).pack(
            side="right", padx=(0, 6))

        left_wrap = tk.Frame(main, bg=PANEL, highlightbackground=BORDER,
                             highlightthickness=1)
        left_wrap.grid(row=1, column=0, sticky="nsew", padx=(0, 6), pady=(4, 0))
        self.left_canvas, self.left_inner = self._scroll_panel(left_wrap)

        # ---- right: merge queue ----
        rt = tk.Frame(main, bg=BG)
        rt.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        tk.Label(rt, text="MERGE QUEUE", bg=BG, fg=TEXT_DIM,
                 font=FONT_SMALL).pack(side="left")
        self._mini_btn(rt, "Clear", self.clear_queue).pack(side="right")

        right_wrap = tk.Frame(main, bg=PANEL, highlightbackground=BORDER,
                              highlightthickness=1)
        right_wrap.grid(row=1, column=1, sticky="nsew", padx=(6, 0), pady=(4, 0))
        self.right_canvas, self.right_inner = self._scroll_panel(right_wrap)
        self.right_wrap = right_wrap

        self.drop_hint = tk.Label(
            self.right_inner, text="Drag PDFs here\n(or double-click on the left)",
            bg=PANEL, fg=TEXT_DIM, font=FONT, pady=30)
        self.drop_hint.pack(fill="x")

        # ---- footer ----
        footer = tk.Frame(r, bg=BG)
        footer.pack(fill="x", padx=12, pady=(0, 12))
        self.status = tk.Label(footer, text="", bg=BG, fg=TEXT_DIM,
                               font=FONT_SMALL, anchor="w")
        self.status.pack(side="left", fill="x", expand=True)
        self.merge_btn = tk.Button(
            footer, text="Merge & Save", command=self.merge_and_save,
            bg=ACCENT, fg="white", activebackground=ACCENT_HI,
            activeforeground="white", relief="flat", font=FONT_BOLD,
            padx=18, pady=6, cursor="hand2")
        self.merge_btn.pack(side="right")

    def _mini_btn(self, parent, text, cmd):
        b = tk.Label(parent, text=text, bg=BG, fg=ACCENT_HI, font=FONT_SMALL,
                     cursor="hand2")
        b.bind("<Button-1>", lambda e: cmd())
        b.bind("<Enter>", lambda e: b.config(fg=GREEN))
        b.bind("<Leave>", lambda e: b.config(fg=ACCENT_HI))
        return b

    def _scroll_panel(self, wrap):
        canvas = tk.Canvas(wrap, bg=PANEL, highlightthickness=0)
        sb = tk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        inner = tk.Frame(canvas, bg=PANEL)
        inner.bind("<Configure>",
                   lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        win = canvas.create_window((0, 0), window=inner, anchor="nw")
        canvas.bind("<Configure>",
                    lambda e: canvas.itemconfigure(win, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        for w in (canvas, inner):
            w.bind("<MouseWheel>", lambda e, c=canvas:
                   c.yview_scroll(int(-e.delta / 120), "units"))
        return canvas, inner

    # -- left list ---------------------------------------------------------
    def refresh(self):
        """Kick off detection in a background thread; UI stays responsive."""
        for w in self.left_inner.winfo_children():
            w.destroy()
        tk.Label(self.left_inner, text="Scanning open PDFs…",
                 bg=PANEL, fg=TEXT_DIM, font=FONT, pady=30).pack(fill="x")
        self.set_status("Scanning…")
        self._scan_gen += 1
        gen = self._scan_gen

        def worker():
            try:
                results = detect_open_pdfs()
                err = None
            except Exception as e:
                results = []
                err = str(e)
            self._scan_queue.put((gen, results, err))

        threading.Thread(target=worker, daemon=True).start()

    def _refresh_done(self, gen, results, err):
        if gen != self._scan_gen:
            return  # a newer scan superseded this one
        for w in self.left_inner.winfo_children():
            w.destroy()
        self.detected = results
        if err:
            self.set_status(f"Scan error: {err} — see Log")
        open_items = [r for r in results if r["section"] == "open"]
        recent_items = [r for r in results if r["section"] == "recent"]
        if not open_items:
            tk.Label(self.left_inner,
                     text="No open PDFs detected right now.",
                     bg=PANEL, fg=TEXT_DIM, font=FONT, pady=12).pack(fill="x")
        for item in open_items:
            self._make_left_card(item)
        if recent_items:
            tk.Label(self.left_inner, text="RECENT / PREVIOUSLY SEEN",
                     bg=PANEL, fg=TEXT_DIM, font=FONT_SMALL, anchor="w",
                     padx=8, pady=4).pack(fill="x", pady=(8, 0))
            for item in recent_items:
                self._make_left_card(item)
        if not err:
            n = sum(1 for d in open_items if d["path"])
            self.set_status(f"{len(open_items)} open ({n} resolved), "
                            f"{len(recent_items)} recent")

    def _make_left_card(self, item):
        resolved = bool(item["path"])
        card = tk.Frame(self.left_inner, bg=CARD,
                        highlightbackground=BORDER, highlightthickness=1)
        card.pack(fill="x", padx=6, pady=3)
        fg = TEXT if resolved else TEXT_DIM
        lbl = tk.Label(card, text=item["name"], bg=CARD, fg=fg,
                       font=FONT, anchor="w", padx=8, pady=6)
        lbl.pack(side="left", fill="x", expand=True)
        if resolved:
            tag = "recent · → drag" if item.get("section") == "recent" \
                  else "→ drag"
        else:
            tag = "unresolved · dbl-click to locate"
        tk.Label(card, text=tag, bg=CARD, fg=TEXT_DIM, font=FONT_SMALL,
                 padx=8).pack(side="right")

        widgets = (card, lbl)
        if resolved:
            for w in widgets:
                w.bind("<ButtonPress-1>",
                       lambda e, it=item, c=card: self._left_drag_start(e, it, c))
                w.bind("<B1-Motion>", self._left_drag_motion)
                w.bind("<ButtonRelease-1>", self._left_drag_end)
                w.bind("<Double-Button-1>",
                       lambda e, it=item: self.add_to_queue(it["path"]))
                w.bind("<Enter>", lambda e, c=card: self._hover(c, True))
                w.bind("<Leave>", lambda e, c=card: self._hover(c, False))
        else:
            for w in widgets:
                w.bind("<Double-Button-1>",
                       lambda e, it=item: self._locate(it))

    def _hover(self, card, on):
        color = CARD_HOT if on else CARD
        card.config(bg=color)
        for child in card.winfo_children():
            child.config(bg=color)

    def _locate(self, item):
        path = filedialog.askopenfilename(
            parent=self.root, title=f"Locate {item['name']}",
            filetypes=[("PDF files", "*.pdf")])
        if path:
            item["path"] = path
            self.refresh_left_only()
            self.add_to_queue(path)

    def refresh_left_only(self):
        self._refresh_done(self._scan_gen, self.detected, None)

    def add_manual(self):
        paths = filedialog.askopenfilenames(
            parent=self.root, title="Add PDFs",
            filetypes=[("PDF files", "*.pdf")])
        for p in paths:
            self.add_to_queue(p)

    # -- drag from left to right --------------------------------------------
    def _left_drag_start(self, event, item, card):
        self._drag_source = item
        self._drag_started = False
        self._drag_origin = (event.x_root, event.y_root)

    def _left_drag_motion(self, event):
        if self._drag_source is None:
            return
        dx = abs(event.x_root - self._drag_origin[0])
        dy = abs(event.y_root - self._drag_origin[1])
        if not self._drag_started and (dx > 5 or dy > 5):
            self._drag_started = True
            self._ghost = tk.Toplevel(self.root)
            self._ghost.overrideredirect(True)
            self._ghost.attributes("-topmost", True)
            try:
                self._ghost.attributes("-alpha", 0.85)
            except tk.TclError:
                pass
            tk.Label(self._ghost, text=self._drag_source["name"],
                     bg=ACCENT, fg="white", font=FONT_SMALL,
                     padx=10, pady=4).pack()
        if self._drag_started:
            self._ghost.geometry(f"+{event.x_root + 12}+{event.y_root + 8}")
            over = self._over_queue(event.x_root, event.y_root)
            self.right_wrap.config(
                highlightbackground=GREEN if over else BORDER)

    def _left_drag_end(self, event):
        item = self._drag_source
        self._drag_source = None
        self.right_wrap.config(highlightbackground=BORDER)
        if getattr(self, "_ghost", None):
            self._ghost.destroy()
            self._ghost = None
        if item and getattr(self, "_drag_started", False):
            if self._over_queue(event.x_root, event.y_root):
                self.add_to_queue(item["path"])
        self._drag_started = False

    def _over_queue(self, x_root, y_root):
        w = self.right_wrap
        wx, wy = w.winfo_rootx(), w.winfo_rooty()
        return (wx <= x_root <= wx + w.winfo_width() and
                wy <= y_root <= wy + w.winfo_height())

    # -- merge queue ---------------------------------------------------------
    def add_to_queue(self, path):
        if not path:
            return
        self.drop_hint.pack_forget()
        row = QueueRow(self.right_inner, self, path)
        row.pack(fill="x", padx=6, pady=3)
        self.rows.append(row)
        self.set_status(f"{len(self.rows)} file(s) in queue")

    def remove_row(self, row):
        self.rows.remove(row)
        row.destroy()
        if not self.rows:
            self.drop_hint.pack(fill="x")
        self.set_status(f"{len(self.rows)} file(s) in queue")

    def clear_queue(self):
        for r in self.rows:
            r.destroy()
        self.rows.clear()
        self.drop_hint.pack(fill="x")
        self.set_status("Queue cleared")

    # reorder: swap when dragged row's center passes a neighbor's center
    def reorder_drag(self, row, y_root):
        idx = self.rows.index(row)
        ry = row.winfo_rooty() + row.winfo_height() / 2
        if idx > 0:
            above = self.rows[idx - 1]
            if y_root < above.winfo_rooty() + above.winfo_height() / 2:
                self.rows[idx - 1], self.rows[idx] = row, above
                self._repack_rows()
                return
        if idx < len(self.rows) - 1:
            below = self.rows[idx + 1]
            if y_root > below.winfo_rooty() + below.winfo_height() / 2:
                self.rows[idx + 1], self.rows[idx] = row, below
                self._repack_rows()

    def reorder_commit(self):
        pass  # order already committed live during drag

    def _repack_rows(self):
        for r in self.rows:
            r.pack_forget()
        for r in self.rows:
            r.pack(fill="x", padx=6, pady=3)

    # -- merge ----------------------------------------------------------------
    def merge_and_save(self):
        if not self.rows:
            messagebox.showinfo("PDF Merge", "Queue is empty — drag some "
                                "PDFs in first.", parent=self.root)
            return
        # validate all ranges before asking for a save path
        plan = []
        for row in self.rows:
            try:
                plan.append((row.path, row.get_page_indices()))
            except ValueError as e:
                messagebox.showerror(
                    "Bad page range",
                    f"{os.path.basename(row.path)}:\n{e}", parent=self.root)
                return
        out = filedialog.asksaveasfilename(
            parent=self.root, title="Save merged PDF",
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
            initialfile="merged.pdf")
        if not out:
            return
        try:
            from pypdf import PdfReader, PdfWriter
            writer = PdfWriter()
            total = 0
            for path, indices in plan:
                with open(path, "rb") as source:
                    reader = PdfReader(source)
                    for i in indices:
                        writer.add_page(reader.pages[i])
                        total += 1
            with open(out, "wb") as fh:
                writer.write(fh)
        except Exception as e:
            messagebox.showerror("Merge failed", str(e), parent=self.root)
            return
        self.set_status(f"Saved {total} pages → {os.path.basename(out)}")
        if messagebox.askyesno("PDF Merge",
                               f"Merged {total} pages.\n\nOpen it now?",
                               parent=self.root):
            os.startfile(out) if IS_WINDOWS else None

    # -- show / hide ------------------------------------------------------------
    def show(self):
        self.clear_queue()
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        self.refresh()

    def hide(self):
        self.root.withdraw()

    def open_log(self):
        if os.path.isfile(LOG_PATH) and IS_WINDOWS:
            os.startfile(LOG_PATH)
        else:
            self.set_status(f"Log: {LOG_PATH}")

    def set_status(self, msg):
        self.status.config(text=msg)


# ---------------------------------------------------------------------------
# Hotkey wiring (keyboard lib runs its own thread; marshal to tk via queue)
# ---------------------------------------------------------------------------
def main():
    try:
        open(LOG_PATH, "w").close()  # fresh log each session
    except OSError:
        pass
    install_exception_logging()
    log(f"app start · python {sys.version.split()[0]} · frozen="
        f"{getattr(sys, 'frozen', False)} · log={LOG_PATH}")
    if tk is None:
        raise RuntimeError("tkinter is unavailable; cannot start the PDF Merge UI")
    root = tk.Tk()
    app = MergeApp(root)
    root.withdraw()  # start hidden in the background

    events = queue.Queue()

    def hotkey_thread():
        try:
            import keyboard
        except ImportError:
            log("hotkey disabled: keyboard package is not installed")
            events.put(("no_keyboard", None))
            return
        try:
            keyboard.add_hotkey(HOTKEY, lambda: events.put(("show", None)))
            log(f"hotkey registered: {HOTKEY}")
            keyboard.wait()  # block forever; hotkeys stay registered
        except Exception as e:
            log("hotkey thread failed:\n" + traceback.format_exc())
            events.put(("hotkey_error", str(e)))

    threading.Thread(target=hotkey_thread, name="pdf-merge-hotkey",
                     daemon=True).start()

    def poll():
        try:
            while True:
                ev, _detail = events.get_nowait()
                if ev == "show":
                    app.show()
                elif ev == "no_keyboard":
                    app.show()
                    app.set_status("keyboard lib missing — hotkey disabled "
                                   "(pip install keyboard)")
                elif ev == "hotkey_error":
                    app.show()
                    app.set_status("Hotkey stopped — keep this window open or "
                                   "see Log for details")
        except queue.Empty:
            pass
        root.after(120, poll)

    poll()
    # first launch: show once so you know it's alive
    app.show()
    root.mainloop()


if __name__ == "__main__":
    main()
