import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import DoubleVar, BooleanVar
import json
import os
import shutil, tempfile
from tkinter import font as tkfont
import importlib
import sys
import re
import csv
import io
import time
import webbrowser
import subprocess, platform
import threading
from datetime import datetime, timezone
from urllib.parse import quote
from urllib.request import Request as UrlRequest, urlopen


#Global Variables
STOP_PROCESSING = False
APP_VERSION = "1.0.0"
GITHUB_OWNER = "kevinpcassidy"
GITHUB_REPOSITORY = "quiz_processing_system"
GITHUB_RELEASES_URL = f"https://github.com/{GITHUB_OWNER}/{GITHUB_REPOSITORY}/releases"
GITHUB_LATEST_RELEASE_API_URL = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPOSITORY}/releases/latest"
APP_DIR_NAME = "quiz_processing_system"
SCOPES = ["https://www.googleapis.com/auth/drive.file"]
OAUTH_CLIENT_FILE = "google_oauth_client.json"
TOKEN_FILE = "google_token.json"
SAMPLE_WORKBOOK = os.path.join("reference", "SAMPLE.xlsx")
SAMPLE_PDF = os.path.join("reference", "SAMPLE.pdf")
SAMPLE_GRADING_SCALE = [5, 6, 7, 8, 9, 10]

# These globals are populated by the staged dependency worker after Tk has
# displayed the main window. Keeping the names stable avoids spreading import
# plumbing throughout the existing processing code.
gspread = Request = RefreshError = Credentials = InstalledAppFlow = None
pd = convert_from_path = pdfinfo_from_path = None
Image = ImageEnhance = ImageTk = cv2 = np = None
pytesseract = process = None

DEPENDENCY_ORDER = ("google", "excel", "pdf", "ocr")
DEPENDENCY_LABELS = {
    "google": "Google Sheets tools",
    "excel": "Excel export tools",
    "pdf": "PDF and calibration tools",
    "ocr": "OCR and name-matching tools",
}

PACKAGED_RESOURCE_PATHS = (
    OAUTH_CLIENT_FILE,
    "reference",
    "LICENSE",
    "THIRD_PARTY_LICENSES.txt",
    os.path.join("vendor", "tesseract", "tesseract.exe"),
    os.path.join("vendor", "tesseract", "tessdata", "eng.traineddata"),
    os.path.join("vendor", "poppler", "Library", "bin", "pdfinfo.exe"),
    os.path.join("vendor", "poppler", "Library", "bin", "pdftoppm.exe"),
)


def resource_root():
    """Return the source tree or PyInstaller's unpacked resource directory."""
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.abspath(sys._MEIPASS)
    return os.path.dirname(os.path.abspath(__file__))


def bundled_tool_paths(root=None):
    """Return bundled tool locations when a complete local vendor tree exists."""
    root = root or resource_root()
    tesseract = os.path.join(root, "vendor", "tesseract", "tesseract.exe")
    poppler_bin = os.path.join(root, "vendor", "poppler", "Library", "bin")
    required = (
        tesseract,
        os.path.join(root, "vendor", "tesseract", "tessdata", "eng.traineddata"),
        os.path.join(poppler_bin, "pdfinfo.exe"),
        os.path.join(poppler_bin, "pdftoppm.exe"),
    )
    if all(os.path.isfile(path) for path in required):
        return tesseract, poppler_bin
    return None, None


def missing_packaged_resources(root=None):
    """List required release resources missing from a frozen application."""
    root = root or resource_root()
    return [
        relative_path
        for relative_path in PACKAGED_RESOURCE_PATHS
        if not os.path.exists(os.path.join(root, relative_path))
    ]


def configure_tesseract(tesseract_module, executable):
    """Point pytesseract at the bundled executable and hide its Windows process."""
    if executable:
        tesseract_module.pytesseract.tesseract_cmd = executable
    original_subprocess_args = tesseract_module.pytesseract.subprocess_args

    def hidden_subprocess_args(include_stdout=True):
        kwargs = original_subprocess_args(include_stdout)
        if platform.system() == "Windows" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW
        return kwargs

    tesseract_module.pytesseract.subprocess_args = hidden_subprocess_args


def configure_pdf2image(pdf2image_module):
    """Prevent bundled Poppler utilities from opening Windows console windows."""
    implementation = pdf2image_module.pdf2image
    original_popen = implementation.Popen

    def hidden_popen(*args, **kwargs):
        if platform.system() == "Windows" and hasattr(subprocess, "CREATE_NO_WINDOW"):
            kwargs.setdefault("creationflags", subprocess.CREATE_NO_WINDOW)
        return original_popen(*args, **kwargs)

    implementation.Popen = hidden_popen


def atomic_write_json(path, data):
    """Write JSON without risking a partially-written settings file."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temporary_path = f"{path}.tmp"
    with open(temporary_path, "w", encoding="utf-8") as handle:
        json.dump(data, handle, indent=4)
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary_path, path)


def install_sample_grading_scale(path, grading_scales):
    """Install the canonical SAMPLE scale, replacing any scale with that name."""
    grading_scales["SAMPLE"] = list(SAMPLE_GRADING_SCALE)
    atomic_write_json(path, grading_scales)
    return grading_scales


def unique_gradebook_title(existing_titles, year=None):
    """Return the annual default title with a numeric suffix when necessary."""
    year = year or datetime.now().year
    base = f"{year}-{year + 1} Quiz Processing System"
    existing_titles = set(existing_titles)
    if base not in existing_titles:
        return base
    suffix = 1
    while f"{base}_{suffix}" in existing_titles:
        suffix += 1
    return f"{base}_{suffix}"


def excel_sheet_title(class_name, used_titles):
    """Create a unique, Excel-compatible worksheet title."""
    cleaned = re.sub(r"[:\\/?*\[\]]", " ", str(class_name)).strip() or "Roster"
    cleaned = re.sub(r"\s+", " ", cleaned)[:31]
    candidate = cleaned
    suffix = 1
    used_lower = {title.casefold() for title in used_titles}
    while candidate.casefold() in used_lower:
        ending = f"_{suffix}"
        candidate = f"{cleaned[:31 - len(ending)]}{ending}"
        suffix += 1
    return candidate


def read_roster_names(path):
    """Read a standardized one-column roster whose A1 value is Name."""
    with open(path, newline="", encoding="utf-8-sig") as handle:
        rows = list(csv.reader(handle))
    if not rows or not rows[0] or rows[0][0].strip().casefold() != "name":
        raise ValueError("Roster cell A1 must contain 'Name'.")
    return [row[0].strip() for row in rows[1:] if row and row[0].strip()]


def write_roster_names(path, names):
    """Atomically write the canonical local roster format."""
    os.makedirs(os.path.dirname(path), exist_ok=True)
    temporary_path = f"{path}.tmp"
    with open(temporary_path, "w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.writer(handle)
        writer.writerow(["Name"])
        writer.writerows([[str(name).strip()] for name in names if str(name).strip()])
        handle.flush()
        os.fsync(handle.fileno())
    os.replace(temporary_path, path)


def normalize_score_value(value):
    """Return numeric score text as a number, collapsing whole floats to ints."""
    if value is None:
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return int(value) if isinstance(value, float) and value.is_integer() else value

    text = str(value).strip()
    if not re.fullmatch(r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)", text):
        return text
    number = float(text)
    return int(number) if number.is_integer() else number


def normalize_score_row(row):
    """Normalize score columns while preserving the name in the first column."""
    values = list(row)
    if not values:
        return values
    return [values[0], *(normalize_score_value(value) for value in values[1:])]


def format_local_timestamp(iso_timestamp):
    if not iso_timestamp:
        return "Rosters have not been updated yet."
    value = datetime.fromisoformat(iso_timestamp).astimezone()
    return f"Rosters last updated on {value.strftime('%B %d, %Y')} at {value.strftime('%I:%M:%S %p')}."


def format_google_progress_status(stage, elapsed_seconds, dot_count):
    """Return a concise animated status for Google startup work."""
    if stage == "connecting" and elapsed_seconds >= 15:
        label = "Still connecting"
    elif stage == "connecting":
        label = "Connecting"
    else:
        label = "Preparing"
    return f"Google Sheets: {label}{'.' * dot_count}"


def version_tuple(version):
    """Return a comparable three-part version tuple, accepting a leading v."""
    match = re.fullmatch(r"v?(\d+)\.(\d+)\.(\d+)", str(version).strip(), re.IGNORECASE)
    if not match:
        raise ValueError(f"Invalid application version: {version!r}")
    return tuple(int(part) for part in match.groups())


def release_download_url(release):
    """Prefer a versioned Windows ZIP asset and fall back to the release page."""
    assets = release.get("assets") or []
    zip_assets = [asset for asset in assets if str(asset.get("name", "")).lower().endswith(".zip")]
    preferred = next(
        (asset for asset in zip_assets if "windows" in str(asset.get("name", "")).lower()),
        zip_assets[0] if zip_assets else None,
    )
    if preferred and preferred.get("browser_download_url"):
        return preferred["browser_download_url"]
    return release.get("html_url") or GITHUB_RELEASES_URL


def fetch_latest_release(timeout=5):
    """Fetch the latest stable public GitHub release without requiring a token."""
    request = UrlRequest(
        GITHUB_LATEST_RELEASE_API_URL,
        headers={
            "Accept": "application/vnd.github+json",
            "User-Agent": f"Quiz-Processing-System/{APP_VERSION}",
            "X-GitHub-Api-Version": "2022-11-28",
        },
    )
    with urlopen(request, timeout=timeout) as response:
        release = json.load(response)
    tag = release.get("tag_name", "")
    version_tuple(tag)
    return release


class AutoScrollableFrame:
    """A vertical canvas/frame whose scrollbar appears only when required."""

    def __init__(self, parent):
        self.frame = ttk.Frame(parent)
        self.canvas = tk.Canvas(self.frame, highlightthickness=0, borderwidth=0)
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        self.scrollbar.grid_remove()
        self.frame.rowconfigure(0, weight=1)
        self.frame.columnconfigure(0, weight=1)
        self.content = ttk.Frame(self.canvas)
        self.window_id = self.canvas.create_window((0, 0), window=self.content, anchor="nw")
        self.content.bind("<Configure>", self._schedule_refresh)
        self.canvas.bind("<Configure>", self._schedule_refresh)
        self.toplevel = parent.winfo_toplevel()
        self._wheel_binding = self.toplevel.bind("<MouseWheel>", self._on_mousewheel, add="+")
        self.frame.bind("<Destroy>", self._on_destroy, add="+")
        self._refresh_job = None
        self._scroll_needed = False

    def pack(self, **kwargs):
        self.frame.pack(**kwargs)

    def _schedule_refresh(self, _event=None):
        try:
            if self._refresh_job:
                self.canvas.after_cancel(self._refresh_job)
            self._refresh_job = self.canvas.after_idle(self._refresh)
        except tk.TclError:
            pass

    def _refresh(self):
        try:
            self._refresh_job = None
            width = max(self.canvas.winfo_width(), 1)
            self.canvas.itemconfigure(self.window_id, width=width)
            self.content.update_idletasks()
            content_height = self.content.winfo_reqheight()
            viewport_height = max(self.canvas.winfo_height(), 1)
            self.canvas.configure(scrollregion=(0, 0, width, content_height))
            self._scroll_needed = content_height > viewport_height
            if self._scroll_needed:
                self.scrollbar.grid()
            else:
                self.scrollbar.grid_remove()
                self.canvas.yview_moveto(0)
        except tk.TclError:
            pass

    def _contains_widget(self, widget):
        while widget is not None:
            if widget in (self.canvas, self.content):
                return True
            widget = getattr(widget, "master", None)
        return False

    def _on_mousewheel(self, event):
        try:
            pointed = self.canvas.winfo_containing(event.x_root, event.y_root)
            if not self._scroll_needed or not self._contains_widget(pointed):
                return None
            top, bottom = self.canvas.yview()
            direction = -1 if event.delta > 0 else 1
            if (direction < 0 and top <= 0) or (direction > 0 and bottom >= 1):
                return "break"
            self.canvas.yview_scroll(direction, "units")
            return "break"
        except tk.TclError:
            return None

    def _on_destroy(self, event):
        if event.widget is self.frame and self._wheel_binding:
            try:
                self.toplevel.unbind("<MouseWheel>", self._wheel_binding)
            except tk.TclError:
                pass
            self._wheel_binding = None

class QuizAppGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Quiz Processing System")

        # Use full height of screen
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        #self.root.geometry(f"{int(screen_w * 0.8)}x{int(screen_h * 0.95)}")
        root.state('zoomed')
        root.minsize(width=1300, height=750)  # prevent too-small window
        
        # Main horizontal layout
        self.paned = ttk.PanedWindow(self.root, orient="horizontal")
        self.paned.pack(fill="both", expand=True)

        # Left, center, right frames (add 20px side padding)
        self.left_outer_frame = ttk.Frame(self.paned)
        self.left_scroll = AutoScrollableFrame(self.left_outer_frame)
        self.left_scroll.pack(fill="both", expand=True)
        self.left_frame = self.left_scroll.content
        self.left_frame.configure(padding=(20, 12, 20, 12))
        self.center_frame = ttk.Frame(self.paned, padding=12)
        self.right_frame = ttk.Frame(self.paned, padding=12)
        self.paned.add(self.left_outer_frame, weight=2)
        self.paned.add(self.center_frame, weight=1)
        self.paned.add(self.right_frame, weight=3)

        # Create custom styles
        style = ttk.Style()
        style.configure("Header.TLabel", font=("Segoe UI", 11, "bold"))
        style.configure("Bold.TLabel", font=("Segoe UI", 10, "bold"))
        style.configure("ProgressCheck.TLabel", foreground="#22aa22", font=("Segoe UI", 11, "bold"))  # green ✔️

        #Get background color
        self.bg_color = self.root.cget("bg")


        
        # Start maximized window
        self.root.state('zoomed')

        # Handle window close
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        

    
        # Mutable user files belong in per-user app data on Windows. During
        # source development on other platforms, keep the historical repo path.
        self.project_root = resource_root()
        self.tesseract_executable, self.poppler_bin = bundled_tool_paths(self.project_root)
        local_app_data = os.environ.get("LOCALAPPDATA")
        self.app_data_dir = (
            os.path.join(local_app_data, APP_DIR_NAME)
            if local_app_data else self.project_root
        )
        os.makedirs(self.app_data_dir, exist_ok=True)
        self.settings_file = os.path.join(self.app_data_dir, "quiz_settings.json")
        self.token_file = os.path.join(self.app_data_dir, TOKEN_FILE)
        self.oauth_client_file = os.path.join(self.project_root, OAUTH_CLIENT_FILE)

        # Dictionary to hold local and Google-backed class definitions.
        self.classes = {}
        self.rosters_dir = os.path.join(self.app_data_dir, "rosters")
        os.makedirs(self.rosters_dir, exist_ok=True)
        self.classes_file = os.path.join(self.app_data_dir, "saved_classes.json")
        self._migrate_legacy_user_files()
        self._load_classes()
        
        # Grading scales storage
        self.grading_file = os.path.join(self.app_data_dir, "saved_grading_scales.json")
        self.grading_scales = {}
        self._load_grading_scales()

        #Set up variables to hold coordinates for name box and score boxes
        self.name_box = None
        self.score_boxes = {}  # {'Topic 1': (x0, y0, x1, y1), ...}
        self.skipped_pages = []
        self.current_page_index = None
        self.page_side_box = None



        #Advanced Variables
        self.enable_side_detection = tk.BooleanVar(value=False)
        self.score_threshold = tk.DoubleVar(value=3.2)
        self.enable_gradebook_var = tk.BooleanVar(value=False)
        self.google_sheets_enabled_var = tk.BooleanVar(value=False)
        self.google_connection_status_var = tk.StringVar(value="Google Sheets: Not Enabled")
        self.google_sheet_title_var = tk.StringVar(value="None created")
        self.google_roster_status_var = tk.StringVar(value="Rosters have not been updated yet.")
        self.show_google_extraction_warning_var = tk.BooleanVar(value=True)
        self.show_sample_walkthrough_var = tk.BooleanVar(value=True)
        self.google_spreadsheet_id = ""
        self.google_rosters_last_updated = ""
        self.google_roster_snapshot = None
        self.google_session_connected = False
        self.google_session_rosters_refreshed = False
        self.google_status_animation_job = None
        self.google_status_animation_stage = None
        self.google_status_animation_started = 0
        self.google_status_animation_dots = 0
        self.update_status_var = tk.StringVar(value=f"Version {APP_VERSION}")
        self.update_status_animation_job = None
        self.update_status_animation_dots = 0
        self.latest_release = None
        self.ignored_update_version = ""

        # Heavy optional dependencies are prepared only after the window is
        # built. A requested feature can move its group to the front of the
        # remaining queue and will resume automatically when loading finishes.
        self.dependency_states = {name: "not_started" for name in DEPENDENCY_ORDER}
        self.dependency_errors = {}
        self.dependency_callbacks = {name: [] for name in DEPENDENCY_ORDER}
        self.dependency_wait_dialogs = {name: [] for name in DEPENDENCY_ORDER}
        self.dependency_priority = []
        self.dependency_lock = threading.Lock()
        
        # Load user preferences
        self.load_settings()

        # central storage for all calibration data
        self.calibration_data = {
            "name_box": None,
            "score_boxes": {},
            "score_calibrations": {}
        }
        
        #Data Processing Variables:
 
        # --- Initialize state flags for manual name prompt ---
        self.waiting_for_manual_name = False
        self.manual_name_selection = None

        self.waiting_for_manual_score = False
        self.manual_score_selection = None

        
        self.stop_processing = False
        self.pdf_thread = None
        
        # Build panels
        self._build_left_panel()
        self._build_center_panel()
        self._build_right_panel()
        if self.google_sheets_enabled_var.get():
            self._start_google_status_animation("preparing")
        self.root.after(100, self._start_dependency_loader)
        self.root.after(250, self._start_update_check)
        self.root.after_idle(self._show_sample_walkthrough_if_enabled)
 
  
    # ---------------- UTILITY ----------------
    def _start_update_check(self):
        """Check for a stable release without delaying or blocking the interface."""
        self.update_status_animation_dots = 0
        self._animate_update_status()
        threading.Thread(target=self._update_check_worker, daemon=True).start()

    def _animate_update_status(self):
        self.update_status_animation_dots = (self.update_status_animation_dots % 3) + 1
        self.update_status_var.set(f"Checking for updates{'.' * self.update_status_animation_dots}")
        self.update_status_animation_job = self.root.after(500, self._animate_update_status)

    def _stop_update_status_animation(self):
        if self.update_status_animation_job is not None:
            try:
                self.root.after_cancel(self.update_status_animation_job)
            except tk.TclError:
                pass
        self.update_status_animation_job = None

    def _update_check_worker(self):
        try:
            release = fetch_latest_release()
        except Exception:
            self.root.after(0, self._finish_update_check_unavailable)
            return
        self.root.after(0, lambda: self._finish_update_check(release))

    def _finish_update_check_unavailable(self):
        self._stop_update_status_animation()
        self.update_status_var.set(f"Version {APP_VERSION} · Update check unavailable")

    def _finish_update_check(self, release):
        self._stop_update_status_animation()
        remote_version = str(release["tag_name"]).lstrip("vV")
        if version_tuple(remote_version) <= version_tuple(APP_VERSION):
            self.update_status_var.set(f"Version {APP_VERSION} is up to date")
            return
        self.latest_release = release
        if remote_version == self.ignored_update_version:
            self.update_status_var.set(f"Version {APP_VERSION} · Update {remote_version} ignored")
        else:
            self.update_status_var.set(f"Version {remote_version} is available")
        self._refresh_update_button()

    def _refresh_update_button(self):
        if hasattr(self, "update_download_button"):
            if self.latest_release:
                self.update_download_button.pack(fill="x", padx=18, pady=(6, 10))
            else:
                self.update_download_button.pack_forget()

    def _show_update_dialog(self):
        if not self.latest_release:
            return
        remote_version = str(self.latest_release["tag_name"]).lstrip("vV")
        popup = tk.Toplevel(self.root)
        popup.title("Quiz Processing System Update")
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)
        frame = ttk.Frame(popup, padding=20)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame,
            text=f"Version {remote_version} is available",
            font=("TkDefaultFont", 14, "bold"),
        ).pack(pady=(0, 12))
        ttk.Label(
            frame,
            text=(
                "Your classes, settings, grading scales, roster caches, and Google authorization "
                "are saved locally in your Windows application-data folder.\n\n"
                "Close Quiz Processing System, delete the entire current Quiz Processing System "
                "program directory, and unzip the new version wherever you want to keep it.\n\n"
                f"Do not delete: %LOCALAPPDATA%\\{APP_DIR_NAME}"
            ),
            wraplength=520,
            justify="left",
        ).pack()
        ignore_var = tk.BooleanVar(value=remote_version == self.ignored_update_version)
        ttk.Checkbutton(frame, text="Ignore this update", variable=ignore_var).pack(anchor="w", pady=(16, 12))

        def close_dialog(open_download=False):
            self.ignored_update_version = remote_version if ignore_var.get() else ""
            self.save_settings()
            popup.destroy()
            if open_download:
                webbrowser.open(release_download_url(self.latest_release))
            self._finish_update_check(self.latest_release)

        buttons = ttk.Frame(frame)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Download Update", command=lambda: close_dialog(True)).pack(side="left")
        ttk.Button(buttons, text="Cancel", command=close_dialog).pack(side="right")
        popup.protocol("WM_DELETE_WINDOW", close_dialog)

    def _start_google_status_animation(self, stage):
        """Animate the startup status until Google reaches a final state."""
        self._stop_google_status_animation()
        self.google_status_animation_stage = stage
        self.google_status_animation_started = time.monotonic()
        self.google_status_animation_dots = 0
        self._animate_google_status()

    def _animate_google_status(self):
        if not self.google_status_animation_stage:
            return
        self.google_status_animation_dots = (self.google_status_animation_dots % 3) + 1
        elapsed = time.monotonic() - self.google_status_animation_started
        self.google_connection_status_var.set(format_google_progress_status(
            self.google_status_animation_stage,
            elapsed,
            self.google_status_animation_dots,
        ))
        self.google_status_animation_job = self.root.after(500, self._animate_google_status)

    def _stop_google_status_animation(self):
        if self.google_status_animation_job is not None:
            try:
                self.root.after_cancel(self.google_status_animation_job)
            except tk.TclError:
                pass
        self.google_status_animation_job = None
        self.google_status_animation_stage = None

    def _start_dependency_loader(self):
        """Prepare optional dependency groups without delaying the first paint."""
        if getattr(self, "dependency_thread", None) and self.dependency_thread.is_alive():
            return
        self.dependency_thread = threading.Thread(target=self._dependency_worker, daemon=True)
        self.dependency_thread.start()

    def _dependency_worker(self):
        while True:
            with self.dependency_lock:
                remaining = [name for name in DEPENDENCY_ORDER
                             if self.dependency_states[name] == "not_started"]
                if not remaining:
                    return
                priority = next((name for name in self.dependency_priority if name in remaining), None)
                group = priority or remaining[0]
                if group in self.dependency_priority:
                    self.dependency_priority.remove(group)
                self.dependency_states[group] = "loading"
            try:
                self._load_dependency_group(group)
            except Exception as error:
                with self.dependency_lock:
                    self.dependency_states[group] = "failed"
                    self.dependency_errors[group] = error
            else:
                with self.dependency_lock:
                    self.dependency_states[group] = "ready"
            self.root.after(0, lambda name=group: self._finish_dependency_group(name))

    def _load_dependency_group(self, group):
        global gspread, Request, RefreshError, Credentials, InstalledAppFlow
        global pd, convert_from_path, pdfinfo_from_path
        global Image, ImageEnhance, ImageTk, cv2, np, pytesseract, process
        if group == "google":
            gspread = importlib.import_module("gspread")
            Request = importlib.import_module("google.auth.transport.requests").Request
            RefreshError = importlib.import_module("google.auth.exceptions").RefreshError
            Credentials = importlib.import_module("google.oauth2.credentials").Credentials
            InstalledAppFlow = importlib.import_module("google_auth_oauthlib.flow").InstalledAppFlow
        elif group == "excel":
            pd = importlib.import_module("pandas")
            importlib.import_module("openpyxl")
        elif group == "pdf":
            pdf2image = importlib.import_module("pdf2image")
            configure_pdf2image(pdf2image)
            convert_from_path = pdf2image.convert_from_path
            pdfinfo_from_path = pdf2image.pdfinfo_from_path
            pil_image = importlib.import_module("PIL.Image")
            Image, ImageEnhance, ImageTk = (
                pil_image,
                importlib.import_module("PIL.ImageEnhance"),
                importlib.import_module("PIL.ImageTk"),
            )
            np = importlib.import_module("numpy")
            cv2 = importlib.import_module("cv2")
        elif group == "ocr":
            pytesseract = importlib.import_module("pytesseract")
            configure_tesseract(pytesseract, self.tesseract_executable)
            process = importlib.import_module("rapidfuzz.process")

    def _finish_dependency_group(self, group):
        dialogs = self.dependency_wait_dialogs[group]
        self.dependency_wait_dialogs[group] = []
        for dialog in dialogs:
            if dialog.winfo_exists():
                dialog.grab_release()
                dialog.destroy()

        state = self.dependency_states[group]
        callbacks = self.dependency_callbacks[group]
        self.dependency_callbacks[group] = []
        if state == "failed":
            error = self.dependency_errors[group]
            if group == "google":
                self._stop_google_status_animation()
                self.google_connection_status_var.set("Google Sheets: Not Connected")
            messagebox.showerror(
                f"{DEPENDENCY_LABELS[group]} unavailable",
                f"The required {DEPENDENCY_LABELS[group].lower()} could not be prepared:\n\n{error}\n\n"
                "Close and reopen the application to try to resolve this error.",
            )
            return
        if group == "google":
            self._start_google_startup_check()
        for callback in callbacks:
            self.root.after_idle(callback)

    def _run_when_dependency_ready(self, group, callback, message):
        """Run now when ready, otherwise show a modal and resume exactly once."""
        with self.dependency_lock:
            state = self.dependency_states[group]
            if state == "ready":
                return True
            if state == "failed":
                error = self.dependency_errors[group]
            else:
                error = None
                self.dependency_callbacks[group].append(callback)
                if state == "not_started" and group not in self.dependency_priority:
                    self.dependency_priority.insert(0, group)
        if error is not None:
            messagebox.showerror(
                f"{DEPENDENCY_LABELS[group]} unavailable",
                f"This feature is unavailable because its dependencies failed to load:\n\n{error}\n\n"
                "Close and reopen the application to try to resolve this error.",
            )
            return False

        popup = tk.Toplevel(self.root)
        popup.title("Preparing Feature")
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)
        popup.protocol("WM_DELETE_WINDOW", lambda: None)
        ttk.Label(popup, text=message, padding=24).pack()
        progress = ttk.Progressbar(popup, mode="indeterminate", length=280)
        progress.pack(padx=24, pady=(0, 24))
        progress.start(12)
        self.dependency_wait_dialogs[group].append(popup)
        self._start_dependency_loader()
        return False

    def _on_close(self):
        self._stop_google_status_animation()
        self._stop_update_status_animation()
        #Delete temp copies of gradebook
        self._cleanup_temp_gradebook_copies()
        
        """Delete or clear topics file when window closes."""
        if os.path.exists(self.topics_file):
            try:
                os.remove(self.topics_file)
            except Exception as e:
                print("Error deleting topics file:", e)
        self.root.destroy()

    def load_settings(self):
        """Load saved settings from file, or use defaults if missing."""
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.enable_side_detection.set(data.get("enable_side_detection", False))
                self.score_threshold.set(data.get("score_threshold", 3.2))
                self.enable_gradebook_var.set(data.get("enable_gradebook_var", False))
                self.google_sheets_enabled_var.set(data.get("google_sheets_enabled", False))
                self.google_connection_status_var.set(
                    "Google Sheets: Preparing…"
                    if self.google_sheets_enabled_var.get()
                    else "Google Sheets: Not Enabled"
                )
                self.show_google_extraction_warning_var.set(data.get("show_google_extraction_warning", True))
                self.show_sample_walkthrough_var.set(data.get("show_sample_walkthrough", True))
                self.google_rosters_last_updated = data.get("google_rosters_last_updated", "")
                self.google_roster_status_var.set(format_local_timestamp(self.google_rosters_last_updated))
                selected = data.get("google_spreadsheet", {})
                self.google_spreadsheet_id = selected.get("id", "")
                self.google_sheet_title_var.set(selected.get("title", "None created"))
                self.ignored_update_version = data.get("ignored_update_version", "")
            except Exception as e:
                print(f"[DEBUG] Failed to load settings: {e}")
        else:
            # Defaults if file doesn't exist
            self.enable_side_detection.set(False)
            self.score_threshold.set(3.2)
            self.enable_gradebook_var.set(False)
            self.google_sheets_enabled_var.set(False)
            self.google_connection_status_var.set("Google Sheets: Not Enabled")
            self.show_google_extraction_warning_var.set(True)
            self.show_sample_walkthrough_var.set(True)

    def save_settings(self):
        """Save current settings to file."""
        try:
            data = {
                "enable_side_detection": self.enable_side_detection.get(),
                "score_threshold": self.score_threshold.get(),
                "enable_gradebook_var": self.enable_gradebook_var.get(),
                "google_sheets_enabled": self.google_sheets_enabled_var.get(),
                "show_google_extraction_warning": self.show_google_extraction_warning_var.get(),
                "show_sample_walkthrough": self.show_sample_walkthrough_var.get(),
                "google_rosters_last_updated": getattr(self, "google_rosters_last_updated", ""),
                "ignored_update_version": self.ignored_update_version,
                "google_spreadsheet": {
                    "id": self.google_spreadsheet_id,
                    "title": self.google_sheet_title_var.get()
                }
            }
            atomic_write_json(self.settings_file, data)
        except Exception as e:
            print(f"[DEBUG] Failed to save settings: {e}")

    def _migrate_legacy_user_files(self):
        """Copy legacy mutable files into LOCALAPPDATA without deleting originals."""
        if self.app_data_dir == self.project_root:
            return
        for filename in ("quiz_settings.json", "saved_classes.json", "saved_grading_scales.json"):
            source = os.path.join(self.project_root, filename)
            destination = os.path.join(self.app_data_dir, filename)
            if os.path.exists(source) and not os.path.exists(destination):
                shutil.copy2(source, destination)
        legacy_rosters = os.path.join(self.project_root, "rosters")
        if os.path.isdir(legacy_rosters):
            for filename in os.listdir(legacy_rosters):
                if not filename.lower().endswith(".csv"):
                    continue
                source = os.path.join(legacy_rosters, filename)
                destination = os.path.join(self.rosters_dir, filename)
                if not os.path.exists(destination):
                    shutil.copy2(source, destination)



    # ---------------- LEFT PANEL ----------------
    def _build_left_panel(self):

        # Progress Section
        ttk.Label(self.left_frame, text="Workflow:", style="Bold.TLabel").pack(anchor="center", pady=(10, 2))

        progress_frame = ttk.Frame(self.left_frame)
        progress_frame.pack(fill="x", pady=(0, 4))

        self.progress_labels = {}

        steps = [
            ("pdf", "Select PDF"),
            ("class", "Select Class"),
            ("grading", "Select Grading Scale"),
            ("topics", "Save Topics"),
            ("calibration", "Calibration"),
            ("download", "Sync to Sheets or Export Grades"),
        ]

        for key, label_text in steps:
            row = ttk.Frame(progress_frame)
            row.pack(fill="x", pady=2)
            icon = ttk.Label(row, text="○", width=2)
            icon.pack(side="left")
            label = ttk.Label(row, text=label_text)
            label.pack(side="left")
            self.progress_labels[key] = icon


        # Select PDF
        ttk.Separator(self.left_frame, orient="horizontal").pack(fill="x", pady=(6,4))
        ttk.Label(self.left_frame, text="Select PDF:", style="Bold.TLabel").pack(anchor="w", pady=(8,2))
        self.pdf_path_var = tk.StringVar(value="No file selected.")
        ttk.Entry(self.left_frame, textvariable=self.pdf_path_var, width=40).pack(fill="x", pady=4)
        ttk.Button(self.left_frame, text="Browse PDF", command=self._select_pdf).pack(pady=4)


        # Select Class
        ttk.Label(self.left_frame, text="Select Class:", style="Bold.TLabel").pack(anchor="w", pady=(6,2))
        # Build list dynamically from saved classes
        class_names = ["-- Select Class --"] + list(self.classes.keys())
        self.class_combo = ttk.Combobox(self.left_frame, values=class_names, state="readonly")
        self.class_combo.current(0)  # start with placeholder
        self.class_combo.pack(fill="x", pady=(0,6))
        self.class_combo.bind("<<ComboboxSelected>>", lambda e: self._class_selected())

        # Select Grading Scale
        ttk.Label(self.left_frame, text="Select Grading Scale:", style="Bold.TLabel").pack(anchor="w", pady=(6,2))
        self.scale_combo = ttk.Combobox(self.left_frame, values=["-- Select Grading Scale --"], state="readonly")
        self.scale_combo.current(0)  # start with placeholder
        self.scale_combo.pack(fill="x", pady=(0,6))
        self.scale_combo.bind("<<ComboboxSelected>>", lambda e: self._grading_selected())

        # Populate with any existing grading scales
        if hasattr(self, "grading_scales") and self.grading_scales:
            self._refresh_grading_combobox()



        # Topics
        ttk.Label(self.left_frame, text="Enter Topic Names:", style="Bold.TLabel").pack(anchor="w", pady=(6,2))
        self.topic_frame = ttk.Frame(self.left_frame)
        self.topic_frame.pack(fill="x", pady=(0,6))

        self.topic_vars = []
        self.topic_entries = []

        # File where topics will be stored
        self.topics_file = os.path.join(os.getcwd(), "saved_topics.json")

        # Add at least one entry field
        self._add_topic_entry()

        # Add button to add new topics
        self.add_topic_btn = ttk.Button(self.left_frame, text="➕ Add Topic", command=self._add_topic_entry)
        self.add_topic_btn.pack(anchor="center", pady=(0,4))

        # Add "Save/Update Topics" button
        self.save_topics_btn = ttk.Button(self.left_frame, text="💾 Save These Topics", command=self._save_topics)
        self.save_topics_btn.pack(anchor="center", pady=(0,6))

        # Load any previously saved topics
        self._load_saved_topics()

        self.calibrate_button = ttk.Button(self.left_frame, text="Run Calibration and Extract Data", command=self._on_run_calibration)
        self.calibrate_button.state(["disabled"])   # disable it here
        self.calibrate_button.pack(fill="x", pady=4)


        # Preferences section
        ttk.Separator(self.left_frame, orient="horizontal").pack(fill="x", pady=(10,8))
        ttk.Label(self.left_frame, text="Preferences", style="Header.TLabel").pack(anchor="w", pady=(0,6))
        ttk.Label(
            self.left_frame,
            textvariable=self.google_connection_status_var,
            wraplength=300,
        ).pack(anchor="w", pady=(0, 2))
        ttk.Button(
            self.left_frame,
            text="Set-up Google Sheets",
            command=self._setup_google_sheets_panel,
        ).pack(fill="x", pady=4)
        ttk.Label(
            self.left_frame,
            textvariable=self.google_roster_status_var,
            wraplength=300,
        ).pack(anchor="w", pady=(3, 2))
        self.google_refresh_button = ttk.Button(
            self.left_frame,
            text="Refresh Rosters Now",
            command=lambda: self._refresh_google_rosters(background=True, notify=True),
        )
        self.google_refresh_button.pack(fill="x", pady=4)
        self.google_refresh_button.state(["disabled"])
        ttk.Separator(self.left_frame, orient="horizontal").pack(fill="x", pady=(6, 4))
        ttk.Button(self.left_frame, text="Set-up Classes", command=self._setup_classes_panel).pack(fill="x", pady=4)
        ttk.Button(self.left_frame, text="Set-up Grading Scale", command=self._open_grading_scale_setup).pack(fill="x", pady=4)
        if self.enable_gradebook_var.get():
            ttk.Button(self.left_frame, text="View Local Gradebook", command=self._on_view_gradebook).pack(fill="x", pady=4)
        ttk.Button(self.left_frame, text="Advanced", command=self.setup_advanced_pop_up).pack(fill="x", pady=4)





    # ---------------- SELECTION HANDLERS ----------------
    def _select_pdf(self):
        """Select PDF file and display its path."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self._set_selected_pdf(file_path)

    def _set_selected_pdf(self, file_path):
        """Select a PDF path and reset state from any previous extraction."""
        self.pdf_path_var.set(file_path)
        self.completed_steps = {
            "pdf_selected": False,
            "roster_loaded": False,
            "names_verified": False,
            "scores_verified": False,
            "topics_saved": False,
            "calibrated": False,
            "csv_downloaded": False,
        }
        if hasattr(self, "class_combo"):
            self.class_combo.current(0)
        if hasattr(self, "scale_combo"):
            self.scale_combo.current(0)
        self._mark_topics_modified()
        for step in self.progress_labels:
            self._set_check(self.progress_labels[step], is_done=False)
        self.mark_step_done("pdf")

    def _show_sample_walkthrough_if_enabled(self):
        if self.show_sample_walkthrough_var.get():
            self._show_sample_walkthrough()

    def _show_sample_walkthrough(self):
        popup = tk.Toplevel(self.root)
        popup.title("Welcome to Quiz Processing System")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("680x570")
        ttk.Label(
            popup,
            text="Welcome to Quiz Processing System!",
            font=("Segoe UI", 17, "bold"),
        ).pack(pady=(20, 8))
        ttk.Label(
            popup,
            text=(
                "We're happy you're here. A sample roster, quiz PDF, and grading scale are included "
                "so you can practice the complete workflow before using your own classes.\n\n"
                "Try the sample walkthrough:\n\n"
                "1. Enable Google Sheets and authorize your Google account.\n"
                "2. Create a Google Sheets gradebook. It will include a worksheet named SAMPLE.\n"
                "3. Under Set-up Classes, add a Google Sheets class mapped to the SAMPLE worksheet.\n"
                "4. Use the SAMPLE.pdf quiz file and choose the SAMPLE grading scale.\n"
                "5. Name the quiz topics, calibrate the sample, process it, and sync the results.\n\n"
                "Prefer not to use Google Sheets? Import SAMPLE.csv from the reference directory as "
                "a local roster, then follow the same PDF and grading-scale steps.\n\n"
                "After the practice run, you'll be ready to use your own rosters, grading scales, and quizzes."
            ),
            wraplength=620,
            justify="left",
        ).pack(fill="x", padx=25, pady=8)
        suppress_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            popup,
            text="Don't show this walkthrough again",
            variable=suppress_var,
        ).pack(anchor="w", padx=25, pady=(8, 12))
        buttons = ttk.Frame(popup)
        buttons.pack(pady=10)

        def close_walkthrough():
            if suppress_var.get():
                self.show_sample_walkthrough_var.set(False)
                self.save_settings()
            popup.destroy()

        def use_sample_pdf():
            sample_path = os.path.join(self.project_root, SAMPLE_PDF)
            if not os.path.isfile(sample_path):
                messagebox.showwarning(
                    "Sample PDF Not Available",
                    "SAMPLE.pdf is not available yet. You can select another PDF with Browse PDF.",
                    parent=popup,
                )
                return
            self._set_selected_pdf(sample_path)
            close_walkthrough()

        ttk.Button(buttons, text="Use SAMPLE.pdf", command=use_sample_pdf).pack(side="left", padx=7)
        ttk.Button(buttons, text="Get Started", command=close_walkthrough).pack(side="left", padx=7)
        popup.protocol("WM_DELETE_WINDOW", close_walkthrough)

    def _class_selected(self):
        if self.class_combo.get() != "-- Select Class --":
            self.mark_step_done("class")
        else:
            self._set_check(self.progress_labels["class"], False)

    # Function to refresh class combobox with current classes
    def _refresh_class_combobox(self):
        values = ["-- Select Class --"] + list(self.classes.keys())
        self.class_combo['values'] = values
        self.class_combo.current(0)  # reset to placeholder

    def _grading_selected(self):
        if self.scale_combo.get() != "-- Select Grading Scale --":
            self.mark_step_done("grading")
        else:
            self._set_check(self.progress_labels["grading"], False)

    def _refresh_grading_combobox(self):
        if os.path.exists(self.grading_file):
            with open(self.grading_file, "r", encoding="utf-8") as f:
                grading_scales = json.load(f)
            values = ["-- Select Grading Scale --"] + list(grading_scales.keys())
            self.scale_combo['values'] = values
            # Keep current selection if valid
            if self.scale_combo.get() not in values:
                self.scale_combo.current(0)



    def _add_topic_entry(self):
        """Add a new topic entry field."""
        frame = ttk.Frame(self.topic_frame)
        frame.pack(fill="x", pady=2)

        var = tk.StringVar()
        entry = ttk.Entry(frame, textvariable=var, width=30)
        entry.pack(side="left", fill="x", expand=True)

        # Only show delete (❌) if more than one entry exists
        if len(self.topic_vars) >= 1:
            del_btn = ttk.Button(frame, text="❌", width=3, command=lambda f=frame, v=var: self._remove_topic_entry(f, v))
            del_btn.pack(side="right", padx=3)
        else:
            del_btn = None

        self.topic_entries.append((frame, del_btn))
        self.topic_vars.append(var)
        var.trace_add("write", lambda *args: self._mark_topics_modified())

    def _remove_topic_entry(self, frame, var):
        """Remove a topic entry."""
        frame.destroy()
        if var in self.topic_vars:
            self.topic_vars.remove(var)
        self.topic_entries = [e for e in self.topic_entries if e[0] != frame]

    def _save_topics(self):
        """Save topics to JSON and update progress tracker."""
        topics = [v.get().strip() for v in self.topic_vars if v.get().strip()]
        if not topics:
            messagebox.showwarning("No Topics", "Please enter at least one topic name before saving.")
            return

        with open(self.topics_file, "w", encoding="utf-8") as f:
            json.dump(topics, f, indent=2)

        # Update button text to "Saved" and mark progress complete
        self.save_topics_btn.config(text="💾 Saved")
        self._set_check(self.progress_labels["topics"], is_done=True)
        
        self.mark_step_done("topics")


    def _mark_topics_modified(self):
        """If topics change after saving, revert button text to 'Update' and progress check to incomplete."""
        if "Saved" in self.save_topics_btn.cget("text"):
            self.save_topics_btn.config(text="💾 Update Topic Names")
            self._set_check(self.progress_labels["topics"], is_done=False)

    def _load_saved_topics(self):
        """Load topics from JSON (if available)."""
        if os.path.exists(self.topics_file):
            try:
                with open(self.topics_file, "r", encoding="utf-8") as f:
                    saved_topics = json.load(f)
                if saved_topics:
                    for frame, _ in self.topic_entries:
                        frame.destroy()
                    self.topic_vars.clear()
                    self.topic_entries.clear()
                    for topic in saved_topics:
                        self._add_topic_entry()
                        self.topic_vars[-1].set(topic)
                    self.save_topics_btn.config(text="✅ Update Topic Names")
            except Exception as e:
                print("Error loading topics:", e)

    def _open_grading_scale_setup(self):
        self._setup_grading_scale_panel()  # Update center panel
        self._show_grading_scales_panel()  # Update right panel



    # ---------------- PROGRESS TRACKING ----------------
    def _set_check(self, label_widget, is_done):
        """Display green check when complete."""
        if is_done:
            label_widget.config(text="✔", style="ProgressCheck.TLabel")
        else:
            label_widget.config(text="○", style="TLabel")

    def _update_progress(self):
        has_pdf = bool(self.pdf_path_var.get() and "No file selected" not in self.pdf_path_var.get())
        has_class = bool(self.class_combo.get())
        has_grading = bool(self.scale_combo.get())
        has_topics = (os.path.exists(self.topics_file) and os.path.getsize(self.topics_file) > 0) or any(v.get().strip() for v in self.topic_vars)

        self._set_check(self.progress_labels["pdf"], has_pdf)
        self._set_check(self.progress_labels["class"], has_class)
        self._set_check(self.progress_labels["grading"], has_grading)
        self._set_check(self.progress_labels["topics"], has_topics)
        
        # Update calibrate button availability
        self._update_calibrate_button_state()

    def mark_step_done(self, step_key):
        """Mark a progress step as done with a green checkmark."""
        if step_key in self.progress_labels:
            self.progress_labels[step_key].config(text="✔", style="ProgressCheck.TLabel")
            # Update calibrate button availability
            self._update_calibrate_button_state()

    def _update_calibrate_button_state(self):
        """
        Enables the Calibrate button only when all required steps are complete.
        """
        required_steps = ["pdf", "class", "grading", "topics"]

        # Check if each required step has a completed checkmark (✔)
        all_done = all(
            self.progress_labels[step].cget("text") == "✔"
            for step in required_steps
            if step in self.progress_labels
        )

        if all_done:
            self.calibrate_button.state(["!disabled"])   # enable
        else:
            self.calibrate_button.state(["disabled"])    # disable


    #Setup Advanced Pop-up
    def setup_advanced_pop_up(self):
        if not self._run_when_dependency_ready(
            "pdf", self.setup_advanced_pop_up, "Preparing PDF and calibration tools…"
        ):
            return
        # --- Create pop-up window ---
        advanced_window = tk.Toplevel(self.root)
        advanced_window.title("Advanced Settings")

        # --- Set geometry ---
        win_width, win_height = 700, 900
        x = (advanced_window.winfo_screenwidth() - win_width) // 2
        y = (advanced_window.winfo_screenheight() - win_height) // 4
        advanced_window.geometry(f"{win_width}x{win_height}+{x}+{y}")

        # --- Canvas with scrollbar ---
        canvas = tk.Canvas(advanced_window, width=win_width, height=win_height)
        scrollbar = ttk.Scrollbar(advanced_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # --- Mousewheel scrolling ---
        def _on_mousewheel(event):
            if event.num == 4: canvas.yview_scroll(-1, "units")
            elif event.num == 5: canvas.yview_scroll(1, "units")
            else: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        # --- Keep references for images ---
        content_frame.image_refs = []

        #Heading
        ttk.Label(content_frame, text="Advanced Settings", font=("Segoe UI", 18, "bold")).pack(pady=(5, 20))

        # --- Enable Local Gradebook Update ---
        ttk.Label(content_frame, text="Track and Update a Local Gradebook", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))
        # Text
        text = (
            "This setting will keep a master gradebook every time you run the program, per roster.\n"
            "It will add each new topic to a new column, and keep all older columns from previous runs.\n"
            "Note that if you keep a topic the same name, it will only update new scores being added to the gradebook.\n"
            "This would be helpful if you have two different pdfs with the same side of the quiz (it will add to the same column).\n"
            "If you enable this setting, it will add the option after processing a pdf to save a new copy of the entire gradebook.\n"
        )
        ttk.Label(content_frame, text=text, wraplength=win_width - 20, justify="left").pack(padx=10, pady=5)

        

        ttk.Checkbutton(
            content_frame,
            text="Enable Local Gradebook Updates",
            variable=self.enable_gradebook_var,
            onvalue=True,
            offvalue=False            
        ).pack(anchor="center", padx=10, pady=5)

        # Separator
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=15)
        # ---------------------------
        # Section 2: Multi-Side Detection
        # ---------------------------
        ttk.Label(content_frame, text="Multi-Side Detection", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))

        # Image
        try:
            img_path = os.path.join(os.getcwd(), "reference", "page_side_visual.png")
            side_img = Image.open(img_path)
            side_img.thumbnail((win_width - 40, 300))
            side_photo = ImageTk.PhotoImage(side_img)
            ttk.Label(content_frame, image=side_photo).pack(pady=(0, 10))
            content_frame.image_refs.append(side_photo)
        except Exception as e:
            ttk.Label(content_frame, text=f"[Image missing: page_side_visual.png]\n{e}", foreground="red").pack()

        # Text
        text = (
            "If your quizzes have two sides, you can enable this setting.\n\n"
            "On your quiz, you can include an 'A side' and 'B side' label, with a circle around the appropriate side.\n"
            "If you select the checkbox below, the program will also look at the side circled.\n"
            "If one quiz is upside down or mislabeled, it will be flagged for manual review.\n"
        )
        ttk.Label(content_frame, text=text, wraplength=win_width - 20, justify="left").pack(padx=10, pady=5)

        # Checkbox
        ttk.Checkbutton(
            content_frame,
            text="Enable Side Detection",
            variable=self.enable_side_detection,
            onvalue=True,
            offvalue=False
        ).pack(pady=(0, 10))

        # Separator
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=15)

        # ---------------------------
        # Section 3: Warning Dialogue Boxes
        # ---------------------------
        ttk.Label(
            content_frame,
            text="Warning Dialogue Boxes",
            font=("Segoe UI", 14, "bold"),
        ).pack(pady=(10, 5))
        ttk.Checkbutton(
            content_frame,
            text="Show Google Sheets safety confirmation before extraction",
            variable=self.show_google_extraction_warning_var,
        ).pack(anchor="center", padx=10, pady=(5, 10))
        ttk.Checkbutton(
            content_frame,
            text="Show sample walkthrough when the application starts",
            variable=self.show_sample_walkthrough_var,
        ).pack(anchor="center", padx=10, pady=(0, 10))

        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=15)

        # ---------------------------
        # Section 4: Detecting Multiple Scores
        # ---------------------------
        ttk.Label(content_frame, text="Detecting Multiple Scores", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))

        # Image
        try:
            img_path = os.path.join(os.getcwd(), "reference", "multiple_score_visual.png")
            score_img = Image.open(img_path)
            score_img.thumbnail((win_width - 40, 300))
            score_photo = ImageTk.PhotoImage(score_img)
            ttk.Label(content_frame, image=score_photo).pack(pady=(0, 10))
            content_frame.image_refs.append(score_photo)
        except Exception as e:
            ttk.Label(content_frame, text=f"[Image missing: multiple_score_visual.png]\n{e}", foreground="red").pack()

        # Text
        text = (
            "This number sets the threshold for detecting multiple circles drawn around scores.\n"
            "The default value of 3.2 worked well in testing. Lower numbers increase sensitivity, "
            "sometimes detecting slight ovals as multiple circles. Higher numbers reduce sensitivity, "
            "sometimes merging two overlapping circles into one.\n\n"
            "Adjust only if needed; most users should leave this at 3.2."
        )
        ttk.Label(content_frame, text=text, wraplength=win_width - 20, justify="left").pack(padx=10, pady=(5, 8))

        # Spinbox for threshold
        ttk.Label(content_frame, text="Circle Detection Threshold:").pack(pady=(5, 2))
        threshold_spin = ttk.Spinbox(
            content_frame,
            from_=1.0,
            to=6.0,
            increment=0.1,
            textvariable=self.score_threshold,
            width=5
        )
        threshold_spin.pack(pady=(0, 10))
        # --- Buttons: Restore Defaults / Cancel / Save & Close ---
        button_frame = ttk.Frame(content_frame)
        button_frame.pack(pady=15)

        # Restore Defaults
        def restore_defaults():
            self.enable_side_detection.set(False)
            self.score_threshold.set(3.2)
            self.enable_gradebook_var.set(False)
            self.show_google_extraction_warning_var.set(True)
            self.show_sample_walkthrough_var.set(True)

        ttk.Button(button_frame, text="Restore Defaults", command=restore_defaults).pack(side="left", padx=10)

        # Cancel button — just close the window without saving
        def cancel_changes():
            advanced_window.destroy()

        ttk.Button(button_frame, text="Cancel", command=cancel_changes).pack(side="left", padx=10)

        # Save & Close button — apply settings and close
        def save_and_close():
            print("Side detection:", self.enable_side_detection.get())
            print("Score threshold:", self.score_threshold.get())
            print("Enable gradebook:", self.enable_gradebook_var.get())
            self.save_settings()
            advanced_window.destroy()


        ttk.Button(button_frame, text="Save and Close", command=save_and_close).pack(side="left", padx=10)


        # --- Scroll region ---
        content_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))




    # ---------------- CENTER PANEL ----------------
    def _build_center_panel(self):
        """Build the scrollable Home content shown between workflow screens."""
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        self.center_home_scroll = AutoScrollableFrame(self.center_frame)
        self.center_home_scroll.pack(fill="both", expand=True)
        home = self.center_home_scroll.content

        ttk.Label(home, text="Quiz Processing System", font=("TkDefaultFont", 20, "bold"), anchor="center").pack(pady=(30, 10))
        ttk.Label(
            home,
            text="Welcome to the Quiz Processing System.\nPlease begin by following the steps on the left frame.\n",
            wraplength=420,
            justify="center",
        ).pack(pady=(0, 10))
        ttk.Label(home, textvariable=self.update_status_var, justify="center").pack(pady=(0, 4))
        self.update_download_button = ttk.Button(
            home,
            text="Download Update",
            command=self._show_update_dialog,
        )
        self._refresh_update_button()
        ttk.Separator(home, orient="horizontal").pack(fill="x", pady=(0, 20))
        ttk.Label(home, text="YouTube Tutorials:", font=("TkDefaultFont", 12, "bold")).pack(pady=(0, 10))
        link_label = tk.Label(
            home,
            text="Come visit me at:\nhttps://www.youtube.com/@KevinsTeacherTech",
            fg="blue",
            cursor="hand2",
            justify="center",
        )
        link_label.pack(pady=(5, 20))
        link_label.bind("<Button-1>", lambda _event: webbrowser.open("https://www.youtube.com/@KevinsTeacherTech"))

        ttk.Separator(home, orient="horizontal").pack(fill="x", pady=(0, 20))
        ttk.Label(home, text="Microsoft Word Templates:", font=("TkDefaultFont", 12, "bold")).pack(pady=(0, 10))
        templates = [
            ("One-Page Quiz Template", "https://docs.google.com/document/d/1EF0sel2g1I94xmV5VCxS2j-vzveeEm6P/export?format=docx"),
            ("Two-Page Quiz Template", "https://docs.google.com/document/d/1xk3f2LEKAum9tqkyix8UkZryegbYsVlA/export?format=docx"),
        ]
        for name, url in templates:
            link = tk.Label(home, text=name, fg="blue", cursor="hand2", justify="center")
            link.pack(pady=2)
            link.bind("<Button-1>", lambda _event, url=url: webbrowser.open(url))

        ttk.Separator(home, orient="horizontal").pack(fill="x", pady=15)
        ttk.Button(
            home,
            text="Export Current Rosters for Mail Merge",
            command=self._export_rosters_for_mail_merge,
        ).pack(fill="x", padx=18, pady=(0, 10))
        ttk.Frame(home, height=15).pack()

    def _setup_google_sheets_panel(self):
        """Show Google Sheets configuration without changing the right panel."""
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        scroll = AutoScrollableFrame(self.center_frame)
        scroll.pack(fill="both", expand=True)
        content = scroll.content
        ttk.Label(
            content,
            text="Set-up Google Sheets",
            font=("Segoe UI", 16, "bold"),
        ).pack(pady=(24, 14))
        self._build_google_controls(content)
        ttk.Separator(content, orient="horizontal").pack(fill="x", padx=18, pady=15)
        ttk.Button(content, text="Home", command=self._build_center_panel).pack(pady=(0, 18))

    def _build_google_controls(self, parent):
        ttk.Label(parent, text="Google Sheets Gradebook", style="Header.TLabel").pack(anchor="w", padx=18)
        ttk.Checkbutton(
            parent,
            text="Manage rosters and grades with Google Sheets",
            variable=self.google_sheets_enabled_var,
            command=self._on_google_opt_in_changed,
        ).pack(anchor="w", padx=18, pady=(5, 3))
        self.google_controls_frame = ttk.Frame(parent)
        self.google_controls_frame.pack(fill="x", padx=18)
        ttk.Label(self.google_controls_frame, textvariable=self.google_connection_status_var, wraplength=400).pack(anchor="w")
        ttk.Label(self.google_controls_frame, textvariable=self.google_sheet_title_var, wraplength=400).pack(anchor="w", pady=(0, 4))
        ttk.Label(self.google_controls_frame, textvariable=self.google_roster_status_var, wraplength=400).pack(anchor="w", pady=(0, 4))
        self.google_authorization_help_var = tk.StringVar(
            value="If the correct browser profile does not open, use the authorization dialog to copy the link."
        )
        ttk.Label(
            self.google_controls_frame,
            textvariable=self.google_authorization_help_var,
            wraplength=400,
            justify="left",
        ).pack(anchor="w", pady=(0, 4))
        self.google_authorize_button = ttk.Button(
            self.google_controls_frame,
            text="Reconnect / Change Google Account" if os.path.exists(self.token_file) else "Authorize Google Sheets",
            command=self._reauthorize_google,
        )
        self.google_authorize_button.pack(fill="x", pady=2)
        self.google_create_button = ttk.Button(
            self.google_controls_frame,
            text="Create Google Sheets Gradebook",
            command=self._confirm_create_google_gradebook,
        )
        self.google_create_button.pack(fill="x", pady=2)
        self.google_open_button = ttk.Button(
            self.google_controls_frame,
            text="Open Google Sheets Gradebook",
            command=self._open_google_gradebook,
        )
        self.google_open_button.pack(fill="x", pady=2)
        self.google_setup_refresh_button = ttk.Button(
            self.google_controls_frame,
            text="Refresh Rosters Now",
            command=lambda: self._refresh_google_rosters(background=True, notify=True),
        )
        self.google_setup_refresh_button.pack(fill="x", pady=2)
        ttk.Button(
            self.google_controls_frame,
            text="About Google Sheets",
            command=self._show_google_sheets_help,
        ).pack(fill="x", pady=2)
        self._update_google_controls_visibility()

    def _export_rosters_for_mail_merge(self):
        if not self._run_when_dependency_ready(
            "excel", self._export_rosters_for_mail_merge, "Preparing Excel export…"
        ):
            return
        if not self.classes:
            messagebox.showinfo("No Classes", "There are no class rosters to export.")
            return
        if not self._confirm_cached_roster_export():
            return
        output_path = filedialog.asksaveasfilename(
            title="Export Current Rosters for Mail Merge",
            defaultextension=".xlsx",
            initialfile="Quiz Processing System Rosters.xlsx",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not output_path:
            return
        used_titles = []
        exported_students = 0
        failures = []
        try:
            with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
                for class_name in self.classes:
                    try:
                        names = read_roster_names(self._resolve_class_path(class_name))
                    except (OSError, ValueError) as error:
                        names = []
                        failures.append(f"{class_name}: {error}")
                    title = excel_sheet_title(class_name, used_titles)
                    used_titles.append(title)
                    pd.DataFrame({"Name": names}).to_excel(writer, sheet_name=title, index=False)
                    exported_students += len(names)
            summary = f"Exported {len(self.classes)} class roster(s) and {exported_students} student(s)."
            if failures:
                summary += "\n\nSome rosters could not be read and were exported empty:\n" + "\n".join(failures)
                messagebox.showwarning("Rosters Exported", summary)
            else:
                messagebox.showinfo("Rosters Exported", summary)
        except Exception as error:
            messagebox.showerror("Roster Export Error", f"Could not export the workbook:\n{error}")

    def _confirm_cached_roster_export(self):
        """Warn when this session could not verify Google-backed roster caches."""
        has_google_rosters = any(
            isinstance(info, dict) and info.get("source") == "google_sheet"
            for info in self.classes.values()
        )
        if (not self.google_sheets_enabled_var.get()
                or not has_google_rosters
                or (self.google_session_connected and self.google_session_rosters_refreshed)):
            return True
        timestamp = format_local_timestamp(self.google_rosters_last_updated)
        if not self.google_rosters_last_updated:
            timestamp = "The date of the last successful roster sync is unknown."
        detail = (
            "Google Sheets has not connected and refreshed the rosters during this session. "
            "The downloaded file will be based on the last locally cached roster.\n\n"
            f"{timestamp}\n\n"
            "Check your internet connection or re-authorize your Google account if necessary."
        )
        result = {"proceed": False}
        popup = tk.Toplevel(self.root)
        popup.title("Google Sheets Rosters Not Refreshed")
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)
        ttk.Label(
            popup,
            text=detail,
            wraplength=540,
            justify="left",
            padding=(24, 24, 24, 12),
        ).pack()
        buttons = ttk.Frame(popup, padding=(24, 8, 24, 24))
        buttons.pack(fill="x")

        def proceed():
            result["proceed"] = True
            popup.destroy()

        ttk.Button(buttons, text="Cancel", command=popup.destroy).pack(side="right", padx=(8, 0))
        ttk.Button(
            buttons, text="Proceed with Cached Roster", command=proceed
        ).pack(side="right")
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)
        self.root.wait_window(popup)
        return result["proceed"]


    #Handle classes saving, loading, deleting, etc.
    def _setup_classes_panel(self):
        # Clear existing widgets
        for widget in self.center_frame.winfo_children():
            widget.destroy()

        # Outer container to place 1/4 of height from the top of the frame
        container = ttk.Frame(self.center_frame)
        container.pack(pady=int(self.root.winfo_height() * 0.10))

        # Heading
        ttk.Label(container, text="Setup Class Roster", font=("Segoe UI", 14, "bold")).pack(pady=(0, 20))

        directions = (
            "Local classes use CSV or Excel roster files. Select a local class to add or remove students."
            if not self.google_sheets_enabled_var.get()
            else "Local classes remain editable here. For Google classes, edit student names in column A of the mapped tab, then refresh rosters."
        )
        ttk.Label(container, text=directions, wraplength=350, justify="left").pack(fill="x", pady=(0, 8))

        #Help Button
        ttk.Button(container, text="How to Setup Classes", command=self._show_setup_classes_help).pack(fill="x", pady=15)
        
        # Add/Delete buttons
        ttk.Button(container, text="Add Local Class", command=self._add_class).pack(fill="x", pady=5)
        if self.google_sheets_enabled_var.get():
            ttk.Button(container, text="Add Google Sheets Class", command=self._add_google_class).pack(fill="x", pady=5)
        ttk.Button(container, text="Add or Remove Students", command=self._edit_class_students).pack(fill="x", pady=5)
        ttk.Button(container, text="Delete Class", command=self._delete_class).pack(fill="x", pady=5)

        # ---- NEW: show class list in right panel ----
        self._show_class_list_panel()
        

      
        # Horizontal line
        ttk.Separator(container, orient="horizontal").pack(fill="x", pady=(15, 5))
        
        # Button to return to original center panel
        ttk.Button(container, text="Home", command=lambda: self.reset_panels()).pack(fill="x", pady=8)


    def _add_class(self):
        """Open a custom window to add a class with name and CSV roster."""

        # Create pop-up window
        popup = tk.Toplevel(self.root)
        popup.title("Add Class to Program")
        popup.transient(self.root)  # Keep above main window
        popup.grab_set()  # Make modal

        # Calculate position: centered horizontally, 1/4 down vertically
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width, height = 500, 200
        x = (screen_w - width) // 2
        y = screen_h // 4
        popup.geometry(f"{width}x{height}+{x}+{y}")

        # Heading
        ttk.Label(popup, text="Add Class to Program", font=("Segoe UI", 14, "bold")).pack(pady=(10, 10))

        # Frame for inputs
        input_frame = ttk.Frame(popup, padding=10)
        input_frame.pack(fill="both", expand=True)

        # Class Name
        ttk.Label(input_frame, text="Name your class:").grid(row=0, column=0, sticky="w")
        class_name_var = tk.StringVar()
        ttk.Entry(input_frame, textvariable=class_name_var, width=30).grid(row=0, column=1, pady=5)

        # CSV File
        ttk.Label(input_frame, text="Select .CSV or Excel file of Class Roster:").grid(row=1, column=0, sticky="w")
        file_path_var = tk.StringVar()
        file_entry = ttk.Entry(input_frame, textvariable=file_path_var, width=30)
        file_entry.grid(row=1, column=1, pady=5)

        def browse_csv():
            path = filedialog.askopenfilename(
                filetypes=[
                    ("CSV Files", "*.csv"),
                    ("Excel Files", "*.xlsx *.xls")
                ]
            )
            if path:
                file_path_var.set(path)

        ttk.Button(input_frame, text="Browse", command=browse_csv).grid(row=1, column=2, padx=5)

        # Save and Close
        def save_and_close():
            class_name = class_name_var.get().strip()
            file_path = file_path_var.get().strip()
            if not class_name:
                messagebox.showwarning("Missing Class Name", "Please enter a class name.")
                return
            if not file_path or not os.path.exists(file_path):
                messagebox.showwarning("Missing CSV or Excel File", "Please select a valid CSV file.")
                return

            # Save roster to subdirectory (always as CSV)
            os.makedirs(self.rosters_dir, exist_ok=True)
            dest_path = os.path.join(self.rosters_dir, f"{class_name}.csv")

            try:
                if file_path.lower().endswith((".xlsx", ".xls")):
                    dataframe = pd.read_excel(file_path)
                    if not len(dataframe.columns) or str(dataframe.columns[0]).strip().casefold() != "name":
                        raise ValueError("Roster cell A1 must contain 'Name'.")
                    names = [str(value).strip() for value in dataframe.iloc[:, 0].dropna() if str(value).strip()]
                else:
                    names = read_roster_names(file_path)
                write_roster_names(dest_path, names)
            except Exception as error:
                messagebox.showerror(
                    "Roster Format Error",
                    f"Could not import the roster. Cell A1 must contain 'Name', with one student per row in column A.\n\n{error}",
                )
                return

            # Update classes dictionary and UI
            self.classes[class_name] = {
                "source": "local_csv",
                "roster_path": self._to_relative_roster_path(dest_path),
            }
            self._save_classes()
            self._refresh_classes_tree()
            self._refresh_class_combobox()

            popup.destroy()  # Close pop-up

        ttk.Button(popup, text="Save and Close", command=save_and_close).pack(pady=10)

        # Handle user closing the window without saving
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _add_google_class(self):
        """Map a named application class to a tab in the active Google gradebook."""
        if not self.google_spreadsheet_id:
            messagebox.showwarning("No Google Gradebook", "Create a Google Sheets gradebook first.")
            return
        try:
            spreadsheet = self.get_gsheet_client(interactive=False).open_by_key(self.google_spreadsheet_id)
            worksheets = spreadsheet.worksheets()
        except Exception as error:
            messagebox.showerror("Google Sheets Error", f"Could not load worksheet tabs:\n{error}")
            return

        popup = tk.Toplevel(self.root)
        popup.title("Add Google Sheets Class")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("520x220")
        frame = ttk.Frame(popup, padding=18)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Course name:").grid(row=0, column=0, sticky="w", pady=7)
        class_name_var = tk.StringVar()
        ttk.Entry(frame, textvariable=class_name_var, width=34).grid(row=0, column=1, sticky="ew")
        ttk.Label(frame, text="Roster tab:").grid(row=1, column=0, sticky="w", pady=7)
        tab_var = tk.StringVar()
        tab_combo = ttk.Combobox(frame, textvariable=tab_var, values=[w.title for w in worksheets], state="readonly")
        tab_combo.grid(row=1, column=1, sticky="ew")
        if worksheets:
            tab_combo.current(0)
        frame.columnconfigure(1, weight=1)
        ttk.Label(
            frame,
            text="The tab must use 'Name' in A1 and one student per row in column A.",
            wraplength=460,
        ).grid(row=2, column=0, columnspan=2, pady=10)

        def save_mapping():
            class_name = class_name_var.get().strip()
            existing = self.classes.get(class_name)
            can_remap = (isinstance(existing, dict)
                         and existing.get("source") == "google_sheet"
                         and existing.get("needs_remapping"))
            if not class_name or (class_name in self.classes and not can_remap):
                messagebox.showwarning("Invalid Course Name", "Enter a unique course name, or a Google course awaiting remapping.")
                return
            worksheet = next((item for item in worksheets if item.title == tab_var.get()), None)
            if worksheet is None:
                messagebox.showwarning("Missing Tab", "Select a roster tab.")
                return
            roster_path = os.path.join(self.rosters_dir, f"{class_name}.csv")
            previous_record = self.classes.get(class_name)
            self.classes[class_name] = {
                "source": "google_sheet",
                "spreadsheet_id": self.google_spreadsheet_id,
                "worksheet_id": worksheet.id,
                "worksheet_title": worksheet.title,
                "roster_path": self._to_relative_roster_path(roster_path),
                "last_refreshed_at": "",
            }
            if not self._refresh_one_google_roster(class_name, self.classes[class_name]):
                if previous_record is None:
                    del self.classes[class_name]
                else:
                    self.classes[class_name] = previous_record
                messagebox.showerror(
                    "Google Sheets Roster",
                    "The roster could not be refreshed. Confirm that the tab has 'Name' in A1.",
                )
                return
            self._save_classes()
            self._refresh_class_combobox()
            popup.destroy()

        ttk.Button(frame, text="Save Class", command=save_mapping).grid(row=3, column=0, columnspan=2, pady=8)

    def _to_relative_roster_path(self, path):
        """Normalize roster file paths to project-relative format for portability."""
        if not path:
            return path

        normalized = os.path.normpath(path)

        try:
            if os.path.isabs(normalized):
                normalized = os.path.relpath(normalized, self.app_data_dir)
        except Exception:
            pass

        return normalized.replace("\\", "/")

    def _resolve_class_path(self, class_name):
        """Return absolute path to class roster file from stored class mapping."""
        class_info = self.classes.get(class_name)
        raw_path = (
            class_info.get("roster_path")
            if isinstance(class_info, dict) else class_info
        )
        if not raw_path:
            return None

        normalized = os.path.normpath(raw_path)
        if os.path.isabs(normalized):
            return normalized

        return os.path.join(self.app_data_dir, normalized)

    def _load_classes(self):
        """Load classes info from JSON if available"""
        if os.path.exists(self.classes_file):
            try:
                with open(self.classes_file, "r", encoding="utf-8") as f:
                    loaded = json.load(f)

                records = loaded.get("classes", loaded) if isinstance(loaded, dict) else {}
                if isinstance(records, dict):
                    self.classes = {}
                    for class_name, record in records.items():
                        if isinstance(record, str):
                            record = {
                                "source": "local_csv",
                                "roster_path": self._to_relative_roster_path(record),
                            }
                        elif isinstance(record, dict):
                            record = dict(record)
                            if record.get("roster_path"):
                                record["roster_path"] = self._to_relative_roster_path(record["roster_path"])
                        self.classes[class_name] = record
                else:
                    self.classes = {}

                # Persist normalized relative paths for cross-machine consistency.
                atomic_write_json(self.classes_file, {"version": 2, "classes": self.classes})
            except Exception as e:
                print("Error loading classes:", e)
                
    def _save_classes(self, redraw=True):
        """Save current classes dictionary to JSON"""
        try:
            atomic_write_json(self.classes_file, {"version": 2, "classes": self.classes})
        except Exception as e:
            print("Error saving classes:", e)
        # After deleting and saving
        if redraw:
            self._setup_classes_panel()  # redraw the panel with updated class list
            self._show_class_list_panel()

    def _edit_class_students(self):
        # Ensure a class is selected in the right-hand Treeview
        if not hasattr(self, "class_tree") or not self.class_tree.get_children():
            messagebox.showinfo("No Classes", "There are no classes to edit.")
            return

        selected = self.class_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select a class from the right pane to edit students.")
            return

        class_name = self.class_tree.item(selected[0], "values")[0]
        class_info = self.classes.get(class_name, {})
        if isinstance(class_info, dict) and class_info.get("source") == "google_sheet":
            messagebox.showinfo(
                "Google Sheets Roster",
                "This roster is managed through Google Sheets. Add, remove, or rename "
                "students in column A of the mapped tab, then click 'Refresh Rosters Now' "
                "or restart Quiz Processing System."
            )
            return
        csv_path = self._resolve_class_path(class_name)
        if not csv_path or not os.path.exists(csv_path):
            messagebox.showerror("Error", f"CSV file for class '{class_name}' not found.")
            return

        # Pop-up window
        popup = tk.Toplevel(self.root)
        popup.title(f"Edit Students: {class_name}")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("400x250+{}+{}".format(self.root.winfo_screenwidth()//2 - 200, self.root.winfo_screenheight()//4))

        # Directions
        ttk.Label(popup, text="Your class roster will open now.\n"
                              "Keep 'Name' in cell A1 and one student per row in column A.\n"
                              "Please make any changes and save before closing.\n"
                              "If prompted, you may need to use a program like Microsoft Excel to make the changes.",
                  wraplength=380, justify="left").pack(pady=20, padx=10)

        # Button frame
        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)

        def proceed():
            # Open CSV in system default program
            try:
                os.startfile(csv_path)  # Windows
            except AttributeError:
                import subprocess, platform
                if platform.system() == "Darwin":  # macOS
                    subprocess.call(("open", csv_path))
                else:  # Linux
                    subprocess.call(("xdg-open", csv_path))

        def done():
            popup.destroy()
            self._refresh_classes_tree()
            self._refresh_class_combobox()

        ttk.Button(btn_frame, text="Proceed", command=proceed).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Done Editing", command=done).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=popup.destroy).pack(side="left", padx=5)

        # ---------------- Auto-refresh logic ----------------
        last_mod_time = os.path.getmtime(csv_path)

        def check_file_change():
            nonlocal last_mod_time
            try:
                current_mod_time = os.path.getmtime(csv_path)
                if current_mod_time != last_mod_time:
                    last_mod_time = current_mod_time
                    self._refresh_classes_tree()
                    self._refresh_class_combobox()
            except FileNotFoundError:
                pass  # CSV might have been deleted

            # Continue checking every 1 second if pop-up still exists
            if popup.winfo_exists():
                popup.after(1000, check_file_change)

        check_file_change()

    def _delete_class(self):
        """Delete the class currently selected in the right-hand Treeview."""

        if not hasattr(self, "class_tree") or not self.class_tree.get_children():
            messagebox.showinfo("No Classes", "There are no classes to delete.")
            return

        selected = self.class_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select a class from the list of classes in the right pane to delete it.")
            return

        class_name = self.class_tree.item(selected[0], "values")[0]
        class_info = self.classes.get(class_name, {})
        roster_file = self.class_tree.item(selected[0], "values")[2]

        # Create confirmation pop-up
        popup = tk.Toplevel(self.root)
        popup.title("Confirm Delete")
        popup.transient(self.root)
        popup.grab_set()

        # Center window: horizontally centered, 1/4 down vertically
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        width, height = 400, 150
        x = (screen_w - width) // 2
        y = screen_h // 4
        popup.geometry(f"{width}x{height}+{x}+{y}")

        # Confirmation message
        if isinstance(class_info, dict) and class_info.get("source") == "google_sheet":
            msg = (f'Please confirm that you want to remove "{class_name}" from the program. '
                   "Nothing in Google Sheets will be deleted or changed.")
        else:
            msg = f'Please confirm that you want to delete the local class "{class_name}" with roster "{roster_file}".'
        ttk.Label(popup, text=msg, wraplength=380, justify="center").pack(pady=(20, 10), padx=10)

        # Button frame
        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)

        def confirm_delete():
            # Remove CSV file
            csv_path = self._resolve_class_path(class_name)
            is_google = isinstance(class_info, dict) and class_info.get("source") == "google_sheet"
            if not is_google and csv_path and os.path.exists(csv_path):
                try:
                    os.remove(csv_path)
                except Exception as e:
                    print(f"Error deleting CSV for class {class_name}: {e}")

            # Remove class from dictionary
            if class_name in self.classes:
                del self.classes[class_name]

            self._save_classes()
            self._refresh_classes_tree()
            self._refresh_class_combobox()
            popup.destroy()

        ttk.Button(btn_frame, text="Delete and Close", command=confirm_delete).pack(side="left", padx=10)
        ttk.Button(btn_frame, text="Cancel", command=popup.destroy).pack(side="right", padx=10)

        # Handle window close as cancel
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _show_setup_classes_help(self):

        # Create pop-up window
        help_window = tk.Toplevel(self.root)
        help_window.title("How to Setup Classes")

        # Load image
        img_path = os.path.join(os.getcwd(), "reference", "class_setup_picture.jpg")
        img = Image.open(img_path)
        photo = ImageTk.PhotoImage(img)

        # Set window width to image width and reasonable height
        win_width = min(img.width+40, 950)
        win_height = min(img.height + 400, 1000)
        help_window.geometry(
            f"{win_width}x{win_height}+"
            f"{int((help_window.winfo_screenwidth()-win_width)/2)}+"
            f"{int(help_window.winfo_screenheight()/30)}"
        )

        # Create canvas + scrollbar
        canvas = tk.Canvas(help_window, width=win_width, height=win_height)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        # Frame inside canvas to hold content
        content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # ---------------- HEADER ----------------
        header = ttk.Label(
            content_frame,
            text="Class Setup Recommendations",
            font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=(10, 5))

        # ---------------- CLICKABLE LINK ----------------
        link_label = tk.Label(
            content_frame,
            text="For YouTube tutorials, visit:\nhttps://www.youtube.com/@KevinsTeacherTech",
            fg="blue",
            cursor="hand2",
            justify="center"
        )
        link_label.pack(pady=(0, 10))

        def open_link(event):
            webbrowser.open("https://www.youtube.com/@KevinsTeacherTech")

        link_label.bind("<Button-1>", open_link)

        # ---------------- HORIZONTAL LINE ----------------
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=10)

        # Mousewheel scrolling function
        def _on_mousewheel(event):
            if event.num == 4:   # Linux scroll up
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:  # Linux scroll down
                canvas.yview_scroll(1, "units")
            else:  # Windows/macOS
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Bind scrolling only when cursor is over canvas
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        # Text instructions
        text = ("Use a gradebook program to export a CSV roster. Put 'Name' in cell A1, "
                "with one student per row in column A.\n\n"
                "We recommend creating a roster of all the classes of one subject, "
                "so if you teach two different subjects or grade level classes, each would be its own class.\n"
                "Here is a picture of a properly formatted CSV file.\n"
                "If you choose to use Google Sheets, you can import rosters from there.\n"
                "If you'd prefer to use a .csv file locally, we recommend that you select the Manage Gradebook checkbox in the Advanced Menu so you can track all grade scans throughout the year.\n\n")
        ttk.Label(content_frame, text=text, wraplength=win_width-20, justify="left").pack(padx=10, pady=5)

        # Image
        ttk.Label(content_frame, image=photo).pack(pady=5)
        content_frame.image = photo  # keep reference to prevent garbage collection

        # Update scroll region
        content_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))


    #Handle Grade Scales Saving/Loading/Deleting/Etc.
    def _setup_grading_scale_panel(self):
        # Clear center panel
        for widget in self.center_frame.winfo_children():
            widget.destroy()

        # Outer container for layout
        container = ttk.Frame(self.center_frame)
        container.pack(pady=int(self.root.winfo_height() * 0.1))  # ~1/10 from top

        # Heading
        ttk.Label(container, text="Set-up Grading Scale", font=("Segoe UI", 14, "bold")).pack(pady=(0, 20))

        # Help Button
        ttk.Button(container, text="How to Set-up Grade Scales", command=self._show_grading_help).pack(fill="x", pady=5)
        ttk.Button(container, text="New Grade Scale", command=self._new_grade_scale_window).pack(fill="x", pady=5)
        ttk.Button(container, text="Edit Grading Scale", command=self._edit_grading_scale).pack(fill="x", pady=5)
        ttk.Button(container, text="Delete Grading Scale", command=self._delete_grading_scale).pack(fill="x", pady=5)
        
        # Horizontal line
        ttk.Separator(container, orient="horizontal").pack(fill="x", pady=(15, 5))
        
        # Button to return to original center panel
        ttk.Button(container, text="Home", command=lambda: self.reset_panels()).pack(fill="x", pady=8)


    def _show_grading_help(self):
        help_window = tk.Toplevel(self.root)
        help_window.title("How to Set-up Grade Scales")
        
        # Load image
        img_path = os.path.join(os.getcwd(), "reference", "grade_setup_picture.png")
        img = Image.open(img_path)
        photo = ImageTk.PhotoImage(img)

        # Window size
        win_width = img.width
        win_height = min(img.height + 400, 800)
        help_window.geometry(f"{win_width}x{win_height}+{int((help_window.winfo_screenwidth()-win_width)/2)}+{int(help_window.winfo_screenheight()/8)}")

        # Canvas and scrollbar
        canvas = tk.Canvas(help_window, width=win_width, height=win_height)
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)

        content_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=content_frame, anchor="nw")

        # ---------------- HEADER ----------------
        header = ttk.Label(
            content_frame,
            text="Grade Scale Setup",
            font=("Segoe UI", 16, "bold")
        )
        header.pack(pady=(10, 5))

        # ---------------- CLICKABLE LINK ----------------
        link_label = tk.Label(
            content_frame,
            text="For YouTube tutorials, visit:\nhttps://www.youtube.com/@KevinsTeacherTech",
            fg="blue",
            cursor="hand2",
            justify="center"
        )
        link_label.pack(pady=(0, 10))

        def open_link(event):
            webbrowser.open("https://www.youtube.com/@KevinsTeacherTech")

        link_label.bind("<Button-1>", open_link)

        # ---------------- HORIZONTAL LINE ----------------
        ttk.Separator(content_frame, orient="horizontal").pack(fill="x", pady=10)

        # Mousewheel scrolling function
        def _on_mousewheel(event):
            if event.num == 4:   # Linux scroll up
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:  # Linux scroll down
                canvas.yview_scroll(1, "units")
            else:  # Windows/macOS
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Bind scrolling only when cursor is over canvas
        def _bind_mousewheel(event):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            canvas.bind_all("<Button-4>", _on_mousewheel)
            canvas.bind_all("<Button-5>", _on_mousewheel)

        def _unbind_mousewheel(event):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        canvas.bind("<Enter>", _bind_mousewheel)
        canvas.bind("<Leave>", _unbind_mousewheel)

        # Text instructions
        text = ("Enter each score that you want to use on your topic quiz. \n\n"
                "I grade these on a four point scale, sometimes giving half points.\n"
                "So, I enterred the following as potential grades, and called it a 4-Point Scale:\n"
                "0, 1, 1.5, 2, 2.5, 3, 3.5, 4\n\n"
                "On each section of the quiz, I put that scale. I used the space bar about 6 times to space out each score.\n"
                "Then, I just have to circle the score for that section to grade the topic quiz.\n"
                "Feel free to download and copy the templates provided online.\n"
                "Below is a picture of the top portion of a quiz, with the score box on the right side.\n\n")
        ttk.Label(content_frame, text=text, wraplength=win_width-20, justify="left").pack(padx=10, pady=5)

        # Image
        ttk.Label(content_frame, image=photo).pack(pady=5)
        content_frame.image = photo  # keep reference to prevent garbage collection

        # Update scroll region
        content_frame.update_idletasks()
        canvas.configure(scrollregion=canvas.bbox("all"))

    def _new_grade_scale_window(self, edit_name=None, edit_scores=None):


        # Load existing grading scales if file exists
        grading_scales = {}
        if os.path.exists(self.grading_file):
            with open(self.grading_file, "r", encoding="utf-8") as f:
                grading_scales = json.load(f)

        
        popup = tk.Toplevel(self.root)
        popup.title("New Grade Scale")
        popup.geometry("400x600+{}+{}".format(self.root.winfo_screenwidth()//2 - 200, self.root.winfo_screenheight()//4))
        popup.transient(self.root)
        popup.grab_set()

        # Grading Scale Name
        ttk.Label(popup, text="Name your grading scale:").pack(anchor="w", padx=10, pady=(10,2))
        name_var = tk.StringVar()
        ttk.Entry(popup, textvariable=name_var, width=30).pack(fill="x", padx=10, pady=(0,10))
        
        # Pre-fill name if editing
        if edit_name:
            name_var.set(edit_name)

        ttk.Separator(popup, orient="horizontal").pack(fill="x", padx=10, pady=5)

        # Valid Scores Section
        ttk.Label(popup, text="Enter Valid Scores:").pack(anchor="w", padx=10, pady=(5,2))
        scores_frame = ttk.Frame(popup)
        scores_frame.pack(fill="x", padx=10)

        score_vars = []

        def add_score_entry(value=""):
            frame = ttk.Frame(scores_frame)
            frame.pack(fill="x", pady=2)

            var = tk.StringVar(value=value)
            entry = ttk.Entry(frame, textvariable=var, width=10)
            entry.pack(side="left", fill="x", expand=True)

            def remove():
                frame.destroy()
                score_vars.remove(var)

            if score_vars:
                ttk.Button(frame, text="❌", width=3, command=remove).pack(side="right", padx=3)
            score_vars.append(var)
            
            # Automatically focus the new entry
            entry.focus_set()
        # --- Pre-fill scores if editing ---
        if edit_scores:
            for s in edit_scores:
                add_score_entry(s)
        else:
            # Start with one empty entry
            add_score_entry()

        # "+" button to add more score entries dynamically
        ttk.Button(popup, text="➕", command=lambda: add_score_entry()).pack(pady=5)


        # Save & Close / Cancel
        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)

        def save_and_close():
            name = name_var.get().strip()
            if not name:
                messagebox.showwarning("Missing Name", "Please enter a grading scale name.")
                return

            # Collect valid scores
            scores = []
            has_skip = False
            non_empty_entries = []
            for v in score_vars:
                val = v.get().strip()
                if val:
                    non_empty_entries.append(val)
                    if val.lower() == "skip":
                        has_skip = True
                        continue
                    try:
                        scores.append(float(val))
                    except ValueError:
                        messagebox.showwarning("Invalid Score", f"'{val}' is not a number or 'Skip'.")
                        return
            scores = sorted(list(set(scores)))  # remove duplicates & sort

            # Convert floats that are whole numbers to int for display
            display_scores = [int(s) if s.is_integer() else s for s in scores]
            if has_skip:
                first_numeric_index = next(
                    (idx for idx, val in enumerate(non_empty_entries) if val.lower() != "skip"),
                    None
                )
                first_skip_index = next(
                    (idx for idx, val in enumerate(non_empty_entries) if val.lower() == "skip"),
                    None
                )

                # Keep existing numeric ordering, but place Skip first or last
                # based on whether the user's first Skip entry came before numbers.
                if first_numeric_index is None or (first_skip_index is not None and first_skip_index < first_numeric_index):
                    display_scores = ["Skip"] + display_scores
                else:
                    display_scores.append("Skip")

            # Update JSON dictionary
            grading_scales[name] = display_scores
            self.grading_scales[name] = display_scores 


            # Save to JSON
            with open(self.grading_file, "w", encoding="utf-8") as f:
                json.dump(grading_scales, f, indent=2)

            # Refresh right panel treeview
            self._refresh_grading_tree()

            # Refresh left panel combobox
            self._refresh_grading_combobox()

            popup.destroy()

        ttk.Button(btn_frame, text="Save and Close", command=save_and_close).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Cancel", command=popup.destroy).pack(side="right", padx=5)


    def _load_grading_scales(self):
        if os.path.exists(self.grading_file):
            try:
                with open(self.grading_file, "r", encoding="utf-8") as f:
                    self.grading_scales = json.load(f)
                # Convert loaded scores to normalized types: float values + optional "Skip"
                for key, scores in self.grading_scales.items():
                    normalized = []
                    for s in scores:
                        if isinstance(s, str) and s.strip().lower() == "skip":
                            normalized.append("Skip")
                        else:
                            normalized.append(float(s))
                    self.grading_scales[key] = normalized
            except Exception as e:
                print("Error loading grading scales:", e)
                self.grading_scales = {}
        else:
            self.grading_scales = {}
        install_sample_grading_scale(self.grading_file, self.grading_scales)


    def _edit_grading_scale(self):
        # Ensure a scale is selected in the right panel treeview
        if not hasattr(self, "grading_tree") or not self.grading_tree.get_children():
            messagebox.showinfo("No Grading Scales", "There are no grading scales to edit.")
            return

        selected = self.grading_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select a grading scale from the right pane to edit.")
            return

        scale_name, _ = self.grading_tree.item(selected[0], "values")
        scores_list = self.grading_scales.get(scale_name, [])

        # Open the New Grade Scale window pre-filled with selection
        self._new_grade_scale_window(edit_name=scale_name, edit_scores=scores_list)

    def _delete_grading_scale(self):
        if not hasattr(self, "grading_tree") or not self.grading_tree.get_children():
            messagebox.showinfo("No Scales", "There are no grading scales to delete.")
            return

        selected = self.grading_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select a grading scale from the list to delete it.")
            return

        scale_name = self.grading_tree.item(selected[0], "values")[0]

        confirm = messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete the grading scale '{scale_name}'?")
        if not confirm:
            return

        # Load JSON
        if os.path.exists(self.grading_file):
            with open(self.grading_file, "r", encoding="utf-8") as f:
                grading_scales = json.load(f)
        else:
            grading_scales = {}

        # Remove the selected scale
        if scale_name in grading_scales:
            del grading_scales[scale_name]

        # Update the in-memory dictionary
        self.grading_scales = grading_scales

        # Save updated JSON
        with open(self.grading_file, "w", encoding="utf-8") as f:
            json.dump(grading_scales, f, indent=2)

        # Refresh UI
        self._refresh_grading_tree()
        self._refresh_grading_combobox()


    #Calibrate Extraction functions

    def _selected_class_is_google(self):
        class_info = self.classes.get(self.class_combo.get(), {}) if hasattr(self, "class_combo") else {}
        return (isinstance(class_info, dict)
                and class_info.get("source") == "google_sheet"
                and not class_info.get("needs_remapping"))

    def _confirm_google_extraction_safety(self):
        result = {"continue": False}
        popup = tk.Toplevel(self.root)
        popup.title("Google Sheets Safety Warning")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("560x230")
        ttk.Label(
            popup,
            text=("Do not make any changes to the Google file while data is being extracted and updated. "
                  "Quiz Processing System will verify that the roster has not changed before synchronizing scores."),
            wraplength=510,
            justify="left",
        ).pack(padx=20, pady=(20, 12))
        suppress_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(popup, text="Don't show this warning again", variable=suppress_var).pack(anchor="w", padx=20)
        buttons = ttk.Frame(popup)
        buttons.pack(pady=18)

        def proceed():
            result["continue"] = True
            if suppress_var.get():
                self.show_google_extraction_warning_var.set(False)
                self.save_settings()
            popup.destroy()

        ttk.Button(buttons, text="Cancel", command=popup.destroy).pack(side="left", padx=8)
        ttk.Button(buttons, text="Continue", command=proceed).pack(side="left", padx=8)
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)
        self.root.wait_window(popup)
        return result["continue"]

    def _on_run_calibration(self):
        if not self._run_when_dependency_ready(
            "pdf", self._on_run_calibration, "Preparing PDF and calibration tools…"
        ):
            return
        if self._selected_class_is_google():
            if self.show_google_extraction_warning_var.get() and not self._confirm_google_extraction_safety():
                return
            try:
                self.google_roster_snapshot = read_roster_names(self._resolve_class_path(self.class_combo.get()))
            except (OSError, ValueError) as error:
                messagebox.showerror("Google Roster Error", f"Refresh this roster before extraction.\n\n{error}")
                return
        else:
            self.google_roster_snapshot = None
        # --- Clear center and right frames ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # --- Clear any previous calibration data to force fresh setup ---
        self.name_box = None
        self.score_boxes = {}
        self.calibration_data = {
            "name_box": None,
            "score_boxes": {},
            "score_calibrations": {}
        }
        self.page_side_box = None
        if hasattr(self, "page_side_clicks"):
            self.page_side_clicks.clear()
        if hasattr(self, "score_click_positions"):
            self.score_click_positions.clear()
        self.skipped_pages = []


        # Check for topics
        self.topics = [v.get().strip() for v in self.topic_vars if v.get().strip()]


        # --- Display PDF first page ---
        self._display_pdf_first_page()  # stores self.right_canvas internally
        
        selected_class = self.class_combo.get()
        selected_scale = self.scale_combo.get()

        # --- Pre-flight checks ---

        missing = []

        if not hasattr(self, "right_full_image") or self.right_full_image is None:
            missing.append("Choose a PDF")

        if not selected_class or selected_class == "-- Select Class --":
            missing.append("Select Class")

        if not selected_scale or selected_scale == "-- Select Grading Scale --":
            missing.append("Select Grading Scale")

        if not self.topics:
            missing.append("Add Topics")



        if missing:
            tk.Label(self.center_frame,
                     text="Please complete the following before starting calibration: \n" + "\n  ".join(missing),
                     wraplength=self.center_frame.winfo_width() - 20,  # wrap near the frame width
                     justify="left",
                     foreground="red").pack(pady=50)
            return



        # --- Add buttons for Name and Topic boxes ---
        ttk.Label(self.center_frame, text="Calibrate Data Extraction", style="Header.TLabel").pack(pady=(0, 10))
        if self._selected_class_is_google():
            ttk.Label(
                self.center_frame,
                text="Do not make any changes to the Google file while data is being extracted and updated.",
                foreground="#b00020",
                wraplength=max(self.center_frame.winfo_width() - 10, 250),
                justify="left",
            ).pack(pady=(0, 10))

        ttk.Button(self.center_frame, text="Set Name Area", command=lambda: self._enable_box_drawing(self.right_canvas)).pack(fill="x", pady=4)

        for topic in self.topics:
            ttk.Button(self.center_frame, text=f"Set {topic} Score Area",
                       command=lambda t=topic: self._enable_box_drawing(self.right_canvas, t)).pack(fill="x", pady=4)

        # --- Optional Page Side button (only if detection enabled) ---
        if self.enable_side_detection.get():  # check the BooleanVar's value
            ttk.Button(
                self.center_frame,
                text="Set Page Side",
                command=lambda: self._enable_box_drawing(
                    self.right_canvas,
                    topic_name="Page Side",
                    color="orange"  # use orange for page side
                )
            ).pack(fill="x", pady=4)



        # --- Next Button (after setting Name/Score boxes) ---
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=10)
        ttk.Label(self.center_frame, text="Press next to move to score box calibration.",
                  font=("Helvetica", 10, "bold")).pack(pady=(5, 2))

        self.next_button = ttk.Button(
            self.center_frame,
            text="Next",
            width=15,
            command=lambda: self._on_next_score_calibration(topic_index=0)
        )
        self.next_button.pack(pady=(5, 10))
        self.next_button.state(["disabled"])   # start disabled

 
    def _update_next_button_state(self):
        """Enable the Next button only when all required calibration boxes are set."""
        ready = True

        # Name box required
        if not self.calibration_data.get("name_box"):
            ready = False

        # Score boxes required for all topics
        for topic in self.topics:
            if topic not in self.calibration_data.get("score_boxes", {}):
                ready = False

        # Page side box required ONLY if enabled
        if self.enable_side_detection.get():
            if "Page Side" not in self.calibration_data.get("score_boxes", {}):
                ready = False

        # Update button
        if hasattr(self, "next_button"):
            if ready:
                self.next_button.state(["!disabled"])
            else:
                self.next_button.state(["disabled"])

 
    def _on_next_score_calibration(self, topic_index=0, bypass_page_side=False):
        """
        Handles middle/right panel updates for calibrating score positions for one topic at a time.
        Displays floating labels at each click and stores calibration data for extraction.
        """

        if not hasattr(self, "score_click_positions"):
            self.score_click_positions = {}

        if not hasattr(self, "calibration_data"):
            self.calibration_data = {"name_box": None, "score_boxes": {}, "score_calibrations": {}}

        # --- Clear center frame ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()
                
        # --- PAGE SIDE CALIBRATION CHECK ---
        if topic_index == 0 and self.enable_side_detection.get() and getattr(self, "name_box", None) and not bypass_page_side:
            self._prompt_page_side_calibration(topic_index=topic_index)
            return  # pause here until user confirms
            
        # --- Check if all topics are done ---
        if topic_index >= len(self.topics):
            ttk.Label(self.center_frame, text="All topics calibrated!", style="Header.TLabel").pack(pady=10)
            ttk.Label(self.center_frame, text="Press the Extract button to begin processing your pdf file.").pack(pady=15)

            ttk.Button(
                self.center_frame,
                text="Extract Data from PDF",
                command=self.run_data_extraction
            ).pack(pady=5)
            self.mark_step_done("calibration")
            return

        topic_name = self.topics[topic_index]

        # --- Center panel heading ---
        ttk.Label(
            self.center_frame,
            text=f"{topic_name} Score Box - Calibrate Each Value",
            style="Header.TLabel"
        ).pack(pady=(0, 6))

        # --- Grading scale check ---
        scale_name = self.scale_combo.get()
        if not scale_name:
            ttk.Label(
                self.center_frame,
                text="Please choose a grade scale in the dropdown menu on the left.",
                foreground="red"
            ).pack(pady=6)
            return

        score_labels = self.grading_scales.get(scale_name, [])
        ttk.Label(
            self.center_frame,
            text=f"Click at each score from least to greatest.\nCurrent grading scale: {score_labels}"
        ).pack(pady=4)

        clicked_count_var = tk.StringVar(value=f"Clicked: 0 / {len(score_labels)}")
        ttk.Label(self.center_frame, textvariable=clicked_count_var).pack(pady=(2, 4))

        # Reset clicks for this topic
        self.score_click_positions[topic_name] = []

        # --- Right panel ---
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        if topic_name not in self.score_boxes:
            ttk.Label(
                self.right_frame,
                text=f"No score box defined for {topic_name}",
                foreground="red"
            ).pack(pady=20)
            return

        # Crop the score box area from the full PDF page
        x0, y0, x1, y1 = self.calibration_data["score_boxes"][topic_name]
        pil_crop = self.right_full_image.crop((x0, y0, x1, y1))

        # --- Canvas setup ---
        self.right_canvas = tk.Canvas(self.right_frame, bg=self.bg_color)
        self.right_canvas.pack(fill="both", expand=True)
        self.right_photo = None
        self.score_coords = []       # original cropped-image x positions
        self.score_lines = []        # canvas line IDs
        self.score_labels_drawn = [] # canvas label IDs

        # --- Function to redraw image + overlays ---
        def redraw_canvas(widget, scale):
            widget.delete("all")
            resized, _ = self._resize_image_to_fit(pil_crop, self.right_frame)
            self.right_photo = ImageTk.PhotoImage(resized)
            widget.create_image(0, 0, anchor="nw", image=self.right_photo)
            self.image_scale = scale

            # redraw clicked red lines and labels proportionally
            for i, orig_x in enumerate(self.score_coords):
                x_display = int(orig_x * scale)
                line = widget.create_line(x_display, 0, x_display, resized.height, fill='red', width=1)
                label_text = str(score_labels[i])
                label = widget.create_text(
                    x_display + 4,
                    resized.height / 2 + 10,
                    text=label_text,
                    fill="red",
                    anchor="nw",
                    font=("TkDefaultFont", 9, "bold")
                )
                self.score_lines[i] = line
                self.score_labels_drawn[i] = label

        # --- Initial display ---
        initial_resized, initial_scale = self._resize_image_to_fit(pil_crop, self.right_frame)
        redraw_canvas(self.right_canvas, initial_scale)

        # --- Click logic ---
        def on_click(event):
            if len(self.score_coords) >= len(score_labels):
                return
            x_display = self.right_canvas.canvasx(event.x)
            x_original = int(x_display / self.image_scale)
            self.score_coords.append(x_original)

            # append placeholders for line/label IDs
            self.score_lines.append(None)
            self.score_labels_drawn.append(None)

            redraw_canvas(self.right_canvas, self.image_scale)
            clicked_count_var.set(f"Clicked: {len(self.score_coords)} / {len(score_labels)}")
            if len(self.score_coords) >= len(score_labels):
                next_btn.config(state='normal')

        self.right_canvas.bind("<Button-1>", on_click)

        # --- Buttons ---
        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=6)

        def reset_clicks():
            self.score_coords.clear()
            self.score_lines.clear()
            self.score_labels_drawn.clear()
            clicked_count_var.set(f"Clicked: 0 / {len(score_labels)}")
            next_btn.config(state='disabled')
            redraw_canvas(self.right_canvas, self.image_scale)

        reset_btn = ttk.Button(btn_frame, text="Reset", command=reset_clicks)
        reset_btn.pack(side="left", padx=4)

        def next_topic():
            # store calibration (real image coordinates)
            scale_dict = {score_labels[i]: self.score_coords[i] for i in range(len(self.score_coords))}
            self.score_click_positions[topic_name] = scale_dict
            self.calibration_data["score_calibrations"][topic_name] = scale_dict
            #print(f"[DEBUG] Saved calibration for '{topic_name}': {scale_dict}")
            self._on_next_score_calibration(topic_index=topic_index + 1)

        next_btn = ttk.Button(btn_frame, text="Next", state='disabled', command=next_topic)
        next_btn.pack(side="left", padx=4)

        # --- Bind resize event with overlays ---
        self._bind_resize_event(
            pil_crop,
            self.right_frame,
            self.right_canvas,
            is_center=True,
            redraw_overlays=lambda widget, scale: redraw_canvas(widget, scale)
        )

    def _prompt_page_side_calibration(self, topic_index=0):
        """
        Display the Page Side box for calibration.
        Allows the user to click positions (Side A, Side B)
        and select the expected side (A or B) via radio buttons.
        """

        if not hasattr(self, "calibration_data"):
            self.calibration_data = {"name_box": None, "score_boxes": {}, "score_calibrations": {}}

        if "Page Side" not in self.calibration_data.get("score_boxes", {}):
            tk.messagebox.showerror("Page Side Not Set", "Please draw the Page Side box first in calibration.")
            return

        # --- Clear frames ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # --- Right frame: cropped Page Side box ---
        x0, y0, x1, y1 = self.calibration_data["score_boxes"]["Page Side"]
        # Start with cropped image
        cropped_page_side = self.right_full_image.crop((x0, y0, x1, y1))

        # --- Fit into right_frame if needed ---
        resized, scale = self._resize_image_to_fit(cropped_page_side, self.right_frame)

        canvas = tk.Canvas(self.right_frame, width=resized.width, height=resized.height, bg=self.bg_color)
        canvas.pack(fill="both", expand=True)
        photo = ImageTk.PhotoImage(resized)
        canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas.photo = photo
        self.right_canvas = canvas
        self.right_img = photo
        self.image_scale = scale


        # --- Center frame: instruction and click guidance ---
        ttk.Label(self.center_frame, text="Calibrate Page Sides", style="Header.TLabel").pack(pady=(0, 10))
        ttk.Label(
            self.center_frame,
            text="Click at Side A, then Side B to calibrate the circle detection.",
            font=("Helvetica", 12),
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=(20, 6))

        # Click positions storage
        self.page_side_clicks = []
        self.page_side_lines = []
        self.page_side_labels_drawn = []

        # --- Click logic ---
        def on_click(event):
            if len(self.page_side_clicks) >= 2:
                return
            x_click = int(event.x / self.image_scale)
            self.page_side_clicks.append(x_click)

            # Draw vertical red line
            line = canvas.create_line(
                event.x, 0, event.x, resized.height,
                fill='red', width=1
            )
            label_text = "Side A" if len(self.page_side_clicks) == 1 else "Side B"
            label = canvas.create_text(
                event.x + 4,
                resized.height / 2 + 10,
                text=label_text,
                fill='red',
                anchor="nw",
                font=("TkDefaultFont", 9, "bold")
            )
            self.page_side_lines.append(line)
            self.page_side_labels_drawn.append(label)
            
            # --- Enable Next button once both clicks are made ---
            if len(self.page_side_clicks) == 2:
                next_btn.config(state='normal')

        canvas.bind("<Button-1>", on_click)

        # --- Reset Button ---
        def reset_page_side():
            self.page_side_clicks.clear()          # clear recorded clicks
            canvas.delete("all")                    # remove all overlays
            # redraw the cropped image
            canvas.create_image(0, 0, anchor="nw", image=photo)
            # reset the scale in case it's needed
            self.image_scale = scale
            # disable Next button until two clicks again
            next_btn.config(state='disabled')

        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=6)

        ttk.Button(btn_frame, text="Reset", command=reset_page_side).pack(side="left", padx=4)


        # --- Radio buttons for expected side ---
        ttk.Separator(self.center_frame, orient='horizontal').pack(fill='x', pady=10)
        ttk.Label(
            self.center_frame,
            text="Choose Expected Side:",
            font=("Helvetica", 11)
        ).pack(pady=(4, 2))

        side_var = tk.StringVar(value="A")
        ttk.Radiobutton(self.center_frame, text="Side A", variable=side_var, value="A").pack(pady=2)
        ttk.Radiobutton(self.center_frame, text="Side B", variable=side_var, value="B").pack(pady=2)



        # --- Next button ---
        def confirm():
                self.calibration_data["page_side_positions"] = self.page_side_clicks
                self.expected_page_side = side_var.get()
                self.calibration_data['expected_side'] = side_var.get()
                #print(f"[DEBUG] Page Side calibration saved: {self.calibration_data['page_side_positions']}, expected: {self.calibration_data['expected_side']}")
                # Resume score calibration, bypass page side check
                # Before calling _on_next_score_calibration
                self.right_frame.unbind("<Configure>")  # remove the debounced resize
                self._on_next_score_calibration(topic_index=topic_index, bypass_page_side=True)

        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=6)

        next_btn = ttk.Button(btn_frame, text="Next", state='disabled', command=confirm)
        next_btn.pack(side="left", padx=4)

        # --- Optional: handle frame resize to redraw canvas if needed ---
        def redraw_canvas(event=None):
            resized_new, scale_new = self._resize_image_to_fit(cropped_page_side, self.right_frame)
            canvas.config(width=resized_new.width, height=resized_new.height)
            photo_new = ImageTk.PhotoImage(resized_new)
            canvas.delete("all")
            canvas.create_image(0, 0, anchor="nw", image=photo_new)
            canvas.photo = photo_new
            self.image_scale = scale_new
            # redraw Side lines
            for i, x_orig in enumerate(self.page_side_clicks):
                x_display = int(x_orig * scale_new)
                line = canvas.create_line(x_display, 0, x_display, resized_new.height, fill='red', width=1)
                label_text = "Side A" if i == 0 else "Side B"
                label = canvas.create_text(x_display + 4, resized_new.height / 2 + 10, text=label_text,
                                           fill='red', anchor="nw", font=("TkDefaultFont", 9, "bold"))
                self.page_side_lines[i] = line
                self.page_side_labels_drawn[i] = label

        self.right_frame.bind("<Configure>", lambda e: self._debounce(redraw_canvas, 150, e))


    def _extract_pages_thread(self, pdf_path):
        start = time.time()

        try:
            # Poppler reading is heavy; runs in this worker thread
            pages = convert_from_path(pdf_path, dpi=200, poppler_path=self.poppler_bin)
            if self.stop_processing:
                return

            # Pass result back to main thread safely
            self.root.after(0, lambda: self._on_pages_loaded(pages, start))

        except Exception as e:
            self.root.after(
                0,
                lambda error=e: messagebox.showerror("PDF Error", str(error))
            )




    def _show_extraction_progress(self, message="Extracting..."):
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()
            
        ttk.Label(self.center_frame, text="Data Extraction", style="Header.TLabel").pack(pady=(30, 50))

        
        # --- Stop Processing Label and Button ---
        ttk.Label(
            self.center_frame,
            text="This button will interrupt all data processing and take you back to the calibration stage.",
            wraplength=self.center_frame.winfo_width() - 10,  # wrap near the frame width
            justify="left"
        ).pack(pady=(20, 4))
        
        stop_btn = ttk.Button(
            self.center_frame,
            text="STOP ALL PROCESSING",
            style="Danger.TButton",  # optional custom style if you define one
            command=self._stop_processing
        )
        stop_btn.pack(pady=(4, 20))

        if self._selected_class_is_google():
            ttk.Label(
                self.center_frame,
                text="Do not make any changes to the Google file while data is being extracted and updated.",
                foreground="#b00020",
                wraplength=max(self.center_frame.winfo_width() - 10, 250),
                justify="left",
            ).pack(pady=(0, 15))


        ttk.Label(
            self.right_frame,
            text=message,
            style="Header.TLabel",
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=20)

        pb = ttk.Progressbar(
            self.right_frame,
            mode="indeterminate",
            length=300
        )
        pb.pack(pady=10)
        pb.start(15)

        self._active_progressbar = pb

    
    def run_data_extraction(self):
        """
        Main data extraction loop.
        Uses the calibrated name box, score boxes, and score calibrations.
        Calls _show_extraction_progress() to update UI during processing.
        Manual review functions are placeholders for later implementation.
        """
        if not self._run_when_dependency_ready(
            "ocr", self.run_data_extraction, "Preparing OCR and name-matching tools…"
        ):
            return
        self.stop_processing = False

        # Pre-flight: get selected class and roster
        class_name = self.class_combo.get()
        scale_name = self.scale_combo.get()
        topics = [v.get().strip() for v in self.topic_vars if v.get().strip()]
        self.topics = topics
        pdf_path = self.pdf_path_var.get()

        if not class_name or class_name.startswith("Select"):
            tk.messagebox.showerror("No Class Selected", "Please select a class before running extraction.")
            return
        if not scale_name or scale_name.startswith("Select"):
            tk.messagebox.showerror("No Grading Scale Selected", "Please select a grading scale before running extraction.")
            return
        if not topics:
            tk.messagebox.showerror("No Topics Found", "Please ensure topics are loaded before extraction.")
            return
        if not pdf_path or not os.path.exists(pdf_path):
            tk.messagebox.showerror("No PDF Selected", "Please select a PDF to extract data from.")
            return

        # Load roster if exists
        roster_file = os.path.join(self.rosters_dir, f"{class_name}.csv")
        # In run_data_extraction, after loading CSV
        self.roster_names = []
        if os.path.exists(roster_file):
            try:
                self.roster_names = read_roster_names(roster_file)
                print(f"Loaded {len(self.roster_names)} students from {roster_file}")
            except ValueError as error:
                messagebox.showerror("Roster Format Error", str(error))
                return
        else:
            print(f"Roster file not found for {class_name}; proceeding without roster matching.")

        # Show initial extraction progress
        self._show_extraction_progress("Converting PDF to image files.\nThis process can take up to a minute for 100 page pdfs.")
        self.root.update_idletasks()
        time.sleep(0.5)
        
        # Show initial UI
        self._show_extraction_progress("Converting PDF to image files.\nThis process can take up to a minute for 100 page pdfs.")

        # Start worker thread
        self.pdf_thread = threading.Thread(
            target=self._extract_pages_thread,
            args=(pdf_path,),
            daemon=True
        )
        self.pdf_thread.start()

    def _on_pages_loaded(self, pages, start_time):
        if self.stop_processing:
            return

        load_time = time.time() - start_time
        est_per_page = load_time / max(1, len(pages))

        # Show progress for each page
        for i in range(len(pages)):
            if self.stop_processing:
                return

            eta = int((len(pages) - i) * est_per_page)
            msg = f"Loaded {i+1}/{len(pages)} pages\nEstimated time remaining: {eta}s"
            self._show_extraction_progress(msg)
            self.root.update_idletasks()

        # Now proceed to OCR + extraction
        self._proceed_with_extraction(pages)

        
    def _proceed_with_extraction(self, all_pages):
        self.extracted_data=[]
        
        for page_index, img in enumerate(all_pages):
            
            if self.stop_processing:
                print("Processing stopped.")
                return                
            
            self.current_page_index = page_index
            
            # --- Page Side Detection ---
            if getattr(self, "enable_side_detection", None) and self.enable_side_detection.get() and hasattr(self, "expected_page_side"):
                if "Page Side" in self.calibration_data.get("score_boxes", {}):
                    # Crop to the calibrated page side box
                    x0, y0, x1, y1 = self.calibration_data["score_boxes"]["Page Side"]
                    cropped_page_side = img.crop((x0, y0, x1, y1))

                    # Detect circled page side in the cropped box
                    detected_side = self.detect_circled_page_side(cropped_page_side, debug=False)

                    # Compare to expected side and prompt if different
                    if detected_side is not None and detected_side != self.expected_page_side:
                        # Page is flagged → prompt manual confirmation
                        self.prompt_manual_page_side(img, detected_side, self.expected_page_side)

                        # Wait for user decision
                        while getattr(self, "waiting_for_manual_page_side", False):
                            self.root.update()
                            time.sleep(0.05)


           

            # Extract student data using calibrated boxes
            student_data = {}
            try:
                # Name box
                cropped_name = img.crop(self.calibration_data["name_box"])
                name_text = pytesseract.image_to_string(cropped_name, config='--psm 6 --oem 3').strip()
                resolved_name = name_text

                # Resolve against roster if provided
                if self.roster_names:
                    if name_text in self.roster_names:
                        resolved_name = name_text
                    else:
                        best = process.extractOne(name_text, self.roster_names)
                        if best and best[1] >= 80:
                            resolved_name = best[0]
                            print(f"Fuzzy matched OCR name '{name_text}' -> '{resolved_name}'")
                        else:
                            resolved_name = self.prompt_manual_name(
                                pil_page_image=img,
                                cropped_name_image=cropped_name,
                                ocr_name_guess=name_text,
                                roster_names=self.roster_names,
                                page_index=page_index
                            )


                student_data['name'] = resolved_name


                # Scores
                topics_dict = {}
                # --- In your main extraction loop ---
                for topic in self.topics:
                    if topic in self.calibration_data["score_boxes"]:
                        cropped_score = img.crop(self.calibration_data["score_boxes"][topic])

                        # For a given topic
                        score_val, circle_count = self.detect_circled_score(
                            cropped_score,
                            self.calibration_data["score_calibrations"].get(topic, {})
                        )


                        # Multiple/no circle detected → manual prompt
                        if circle_count >= 2 or score_val is None:
                            if circle_count>=2:
                                text_of_error = "Multiple Scores Detected"
                            else:
                                text_of_error = "No Score Detected"
                            
                            self.prompt_manual_score(full_page_image=img,
                                                     cropped_image=cropped_score,
                                                     student_name=resolved_name,
                                                     topic_label=topic,
                                                     text_to_display = text_of_error)

                            # Wait for user input
                            while self.waiting_for_manual_score:
                                self.root.update()  # Keep GUI responsive
                                time.sleep(0.05)

                            score_val = self.manual_score_selection  # Get value from the prompt

                        # Store score
                        topics_dict[topic] = score_val

                student_data['topics'] = topics_dict
                self.extracted_data.append(student_data)

            except Exception as e:
                print(f"Error processing page {page_index+1}: {e}")
                continue

            # Update progress for UI
            self._show_extraction_progress(f"Extracted name and score(s) from page {page_index+1} of {len(all_pages)}")
            # Then force the UI to refresh so the user sees it immediately
            self.root.update_idletasks()
        
        # Final update - 
        self._show_extraction_progress("Data extraction complete. Total students processed: "+ str(len(self.extracted_data)))
        print("Extraction finished. Total students processed:", len(self.extracted_data))


        # Short pause to let UI refresh
        self.root.update()
        time.sleep(0.5)  # half-second pause

        # Show data preview with export buttons
        self.show_data_preview_with_export()

    def _stop_processing(self):
        """Triggered by the STOP ALL PROCESSING button to interrupt data extraction."""
        self.stop_processing = True
        # Optionally update UI immediately
        for widget in self.center_frame.winfo_children():
            widget.destroy()

        ttk.Label(
            self.center_frame,
            text="Processing stopped by user.\nReturning to calibration stage...",
            foreground="red",
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=20)

        # ⏳ Wait 1500 ms (1.5 seconds), then return to calibration
        self.center_frame.after(1500, self._on_run_calibration)

        self.reset_panels()
        self._on_run_calibration()

    def skip_current_page(self):
        """
        User chose to skip the current page during manual name/score selection.
        Marks the page as skipped and resumes extraction loop.
        """
        print(f"[INFO] Page {self.current_page_index + 1} skipped by user.")
        
        # Save skipped index for logging or later reference
        if not hasattr(self, "skipped_pages"):
            self.skipped_pages = []
        self.skipped_pages.append(self.current_page_index)

        # Signal to resume extraction (same as confirming name)
        self.manual_name_selection = None
        self.waiting_for_manual_name = False  # release the wait loop

        # Update progress view to resume normal flow
        self._show_extraction_progress(f"Skipping page {self.current_page_index + 1} and continuing...")
        self.root.update()

    def skip_current_score(self, student_name, topic_label):
        """Skip the current score prompt and continue extraction."""
        print(f"[INFO] Skipped score for {student_name} - {topic_label}")
        self.manual_score_selection = ""   # leave blank in the dictionary
        self.waiting_for_manual_score = False
        self._show_extraction_progress(f"Skipped score for {student_name} - {topic_label}")


    def prompt_manual_name(self, pil_page_image, cropped_name_image, ocr_name_guess, roster_names, page_index):
        """
        Display the full page (right frame) and a name correction dropdown (center frame)
        when OCR name doesn't match a roster entry.
        Allows manual entry of a new name.
        """
        print("Detected no name for page " + str(page_index))

        # --- Clear frames ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()


        # --- Right frame: show full page scaled ---
        for widget in self.right_frame.winfo_children():
            widget.destroy()  # remove old preview

        self.right_frame.update_idletasks()
        resized, scale = self._resize_image_to_fit(pil_page_image, self.right_frame)



        canvas = tk.Canvas(self.right_frame, width=resized.width, height=resized.height, bg=self.bg_color)
        canvas.pack(fill="both", expand=True)
        photo = ImageTk.PhotoImage(resized)
        canvas.create_image(0, 0, anchor="nw", image=photo)

        # Store references
        self.right_canvas = canvas
        self.right_img = photo
        self.image_scale = scale
        
        self._bind_resize_event(pil_page_image, self.right_frame, canvas)


        # --- Center frame: heading and dropdown ---
        ttk.Label(
            self.center_frame,
            text=f"Page {page_index + 1}: OCR guessed '{ocr_name_guess}'.\nPlease select or enter the correct student name:",
            style="Header.TLabel",
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=(10, 8))

        # Add placeholder at the top of the list
        roster_list = ["---Choose Student Name from Roster---"] + roster_names
        name_var = tk.StringVar(value=roster_list[0])
        dropdown = ttk.Combobox(self.center_frame, textvariable=name_var, values=roster_list, width=50, state="readonly")
        dropdown.pack(pady=8)

        # --- Manual entry toggle ---
        manual_entry_frame = ttk.Frame(self.center_frame)
        manual_entry_frame.pack(pady=4)

        manual_name_var = tk.StringVar()
        entry_visible = tk.BooleanVar(value=False)

        def show_manual_entry():
            if not entry_visible.get():
                ttk.Label(
                    manual_entry_frame,
                    text="Type the student's name. This student's scores will be put at the end of your file.\n"
                         "To get the name in the correct order for your gradebook, please update the class roster under Preferences.",
                    wraplength=self.center_frame.winfo_width() - 20
                ).pack(pady=(5, 3))
                ttk.Entry(manual_entry_frame, textvariable=manual_name_var, width=50).pack(pady=(2, 5))
                entry_visible.set(True)

        ttk.Button(self.center_frame, text="New Name", command=show_manual_entry).pack(pady=(8, 4))

        # --- Confirm button ---
        def confirm():
            selected = name_var.get().strip()
            if entry_visible.get() and manual_name_var.get().strip():
                self.manual_name_selection = manual_name_var.get().strip()
            elif selected != "---Choose Student Name from Roster---":
                self.manual_name_selection = selected
            else:
                self.manual_name_selection = None  # User didn’t make a valid choice
            self.waiting_for_manual_name = False

        ttk.Button(self.center_frame, text="Confirm Name", command=confirm).pack(pady=(10, 8))

        ttk.Label(
            self.center_frame,
            text="If you cannot find the correct name, you can type it manually or press the skip this page button.",
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=5)
        
        ttk.Button(self.center_frame, text="Skip This Page", command=self.skip_current_page).pack(pady=(10, 0))

        # --- Stop Processing Label and Button ---
        ttk.Label(
            self.center_frame,
            text="This button will interrupt all data processing and take you back to the calibration stage.",
            wraplength=self.center_frame.winfo_width() - 10,  # wrap near the frame width
            justify="left"
        ).pack(pady=(20, 4))
        
        stop_btn = ttk.Button(
            self.center_frame,
            text="STOP ALL PROCESSING",
            style="Danger.TButton",  # optional custom style if you define one
            command=self._stop_processing  # placeholder for later integration
        )
        stop_btn.pack(pady=(4, 20))

        # --- Wait for user input ---
        self.waiting_for_manual_name = True
        self.manual_name_selection = None
        while self.waiting_for_manual_name:
            self.center_frame.update_idletasks()
            self.center_frame.update()
            self.root.update_idletasks()
            self.root.update()


        return self.manual_name_selection

    def detect_circled_score(self, pil_image, score_positions, debug=False):
 
        img = np.array(pil_image.convert("L"))
        img = cv2.medianBlur(img, 5)

        circles = cv2.HoughCircles(
            img,
            cv2.HOUGH_GRADIENT,
            dp=1.2,
            minDist=30,
            param1=50,
            param2=30,
            minRadius=10,
            maxRadius=40
        )

        if circles is None:
            return None, 0

        circles = np.around(circles).astype(np.int32)[0]  # prevent overflow
        
        merged = []
        merged_indices = []
        distinct_centers = []

        
        for idx, (x, y, r) in enumerate(circles):
            merged_into_existing = False
            for m_idx, (mx, my, mr) in enumerate(merged):
                center_dist = np.hypot(x - mx, y - my)
                if center_dist < min(r, mr) * self.score_threshold.get():
                    merged[m_idx] = (
                        int((x + mx) / 2),
                        int((y + my) / 2),
                        int((r + mr) / 2)
                    )
                    merged_indices.append((idx, m_idx))
                    merged_into_existing = True
                    break
            if not merged_into_existing:
                merged.append((x, y, r))
                distinct_centers.append((x, y))

        # --- Determine closest score ---
        best_score = None
        min_dist = float("inf")
        for score, ref_x in score_positions.items():
            for (x, y) in distinct_centers:
                dist = abs(x - ref_x)
                if dist < min_dist:
                    min_dist = dist
                    best_score = score

        # --- Debug output ---
        if debug:
            debug_img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            for (x, y, r) in circles:
                cv2.circle(debug_img, (x, y), r, (0, 0, 255), 1)
            for (x, y, r) in merged:
                cv2.circle(debug_img, (x, y), r, (0, 255, 0), 2)
            for (x, y) in distinct_centers:
                cv2.circle(debug_img, (x, y), 12, (255, 0, 0), 2)
            if best_score is not None:
                cv2.putText(debug_img, f"Score: {best_score}", (10, 40),
                            cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 255, 0), 2)
            cv2.imwrite("debug_circles_debug.png", debug_img)
            print(f"DEBUG: {len(circles)} original, {len(merged)} merged, {len(distinct_centers)} final")

        return best_score, len(distinct_centers)

    def detect_circled_page_side(self, pil_image, debug=False):
        """
        Detect which side (A or B) is circled in the page side box.
        Uses the two x-positions stored in self.page_side_clicks.
        Returns "A", "B", or None if no circle detected.
        """
        if not hasattr(self, "page_side_clicks") or len(self.page_side_clicks) != 2:
            return None

        side_positions = {"A": self.page_side_clicks[0], "B": self.page_side_clicks[1]}

        # --- Convert to grayscale and blur ---
        img = np.array(pil_image.convert("L"))
        img = cv2.medianBlur(img, 5)

        # --- Detect circles ---
        circles = cv2.HoughCircles(
            img,
            cv2.HOUGH_GRADIENT,
            dp=1.2,
            minDist=20,
            param1=50,
            param2=30,
            minRadius=5,
            maxRadius=40
        )

        if circles is None:
            return None

        circles = np.around(circles).astype(np.int32)[0]

        # --- Merge close circles ---
        merged = []
        for x, y, r in circles:
            merged_into_existing = False
            for i, (mx, my, mr) in enumerate(merged):
                if np.hypot(x - mx, y - my) < min(r, mr) * self.score_threshold.get():
                    merged[i] = (int((x + mx)/2), int((y + my)/2), int((r + mr)/2))
                    merged_into_existing = True
                    break
            if not merged_into_existing:
                merged.append((x, y, r))

        # --- Determine closest side ---
        closest_side = None
        min_dist = float("inf")
        for side, ref_x in side_positions.items():
            for (x, y, r) in merged:
                dist = abs(x - ref_x)
                if dist < min_dist:
                    min_dist = dist
                    closest_side = side

        # --- Debug output ---
        if debug:
            debug_img = cv2.cvtColor(img, cv2.COLOR_GRAY2BGR)
            for (x, y, r) in merged:
                cv2.circle(debug_img, (x, y), r, (0, 255, 0), 2)
            cv2.putText(debug_img, f"Detected: {closest_side}", (10, 40),
                        cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
            cv2.imwrite("debug_page_side.png", debug_img)
            print(f"[DEBUG] Detected page side: {closest_side}")

        return closest_side


    def prompt_manual_score(self, full_page_image, cropped_image, student_name, topic_label, text_to_display):
        """Prompt user to manually select or type a score when multiple or no circles detected."""
        # --- Clear and update center frame ---
        pil_page_image = full_page_image
        
        for widget in self.center_frame.winfo_children():
            widget.destroy()

        ttk.Label(
            self.center_frame,
            text=text_to_display,
            font=("Arial", 18, "bold")
        ).pack(pady=10)

        ttk.Label(
            self.center_frame,
            text=f"Student Name: {student_name}",
            font=("Arial", 12)
        ).pack()

        ttk.Label(
            self.center_frame,
            text=f"Topic: {topic_label}",
            font=("Arial", 12)
        ).pack(pady=(0, 10))

        # --- Display the cropped score image in the center frame ---
        resized, scale = self._resize_image_to_fit(cropped_image, self.center_frame)


        img = ImageTk.PhotoImage(resized)
        img_label = ttk.Label(self.center_frame, image=img)
        img_label.image = img
        img_label.pack(pady=10)

        # Enable live resizing for the cropped score image
        self._bind_resize_event(cropped_image, self.center_frame, img_label, is_center=True)


        ttk.Label(
            self.center_frame,
            text="Click the circled score or type it manually:",
            wraplength=self.center_frame.winfo_width() - 20
        ).pack()

        score_var = tk.StringVar()
        entry = ttk.Entry(self.center_frame, textvariable=score_var, width=10, justify='center')
        entry.pack(pady=5)

        def on_click(event):
            # Get display image width (cropped, resized) and actual cropped image width
            display_w = img.width()  # width of what user clicks on (PhotoImage)
            orig_w, orig_h = cropped_image.size  # actual size of cropped image

            # Compute click position relative to original full-scale image
            click_ratio = event.x / display_w
            orig_x = int(click_ratio * orig_w)

            # Get the calibration data for this topic
            score_positions = self.calibration_data["score_calibrations"].get(topic_label, {})
            if not score_positions:
                messagebox.showwarning("Missing Calibration", f"No score calibration found for {topic_label}.")
                return

            # Find the score whose calibrated x-position is closest to where the user clicked
            closest_score = min(
                score_positions.keys(),
                key=lambda s: abs(score_positions[s] - orig_x)
            )

            # Update the score field
            score_var.set(str(closest_score))

            #print(f"[DEBUG] Clicked X={event.x}, mapped to {orig_x}, selected score={closest_score}")




        img_label.bind("<Button-1>", on_click)

        def confirm():
            val_raw = score_var.get().strip()
            if val_raw.lower() == "skip":
                self.manual_score_selection = "Skip"
            else:
                val = re.sub(r"[^\d\.]", "", val_raw)
                if not val:
                    messagebox.showwarning("Invalid Entry", "Please enter or click a valid score (or 'Skip').")
                    return
                self.manual_score_selection = val
            self.waiting_for_manual_score = False
            self._show_extraction_progress(f"Score saved for {student_name} - {topic_label}")

        ttk.Button(self.center_frame, text="Confirm", command=confirm).pack(pady=10)

        ttk.Button(
            self.center_frame,
            text="Skip This Score",
            command=lambda: self.skip_current_score(student_name, topic_label)
        ).pack(pady=(10, 15))

        # --- Stop Processing Label and Button ---
        ttk.Label(
            self.center_frame,
            text="This button will interrupt all data processing and take you back to the calibration stage.",
            wraplength=self.center_frame.winfo_width() - 10,  # wrap near the frame width
            justify="left"
        ).pack(pady=(20, 4))
        
        stop_btn = ttk.Button(
            self.center_frame,
            text="STOP ALL PROCESSING",
            style="Danger.TButton",  # optional custom style if you define one
            command=self._stop_processing  # placeholder for later integration
        )
        stop_btn.pack(pady=(4, 20))

        # --- Right frame: show full page scaled ---
        for widget in self.right_frame.winfo_children():
            widget.destroy()  # remove old preview

        self.right_frame.update_idletasks()
        frame_w = self.right_frame.winfo_width() or 800
        frame_h = self.right_frame.winfo_height() or 800

        # Initial scale by height
        resized, scale = self._resize_image_to_fit(pil_page_image, self.right_frame)



        canvas = tk.Canvas(self.right_frame, width=resized.width, height=resized.height, bg=self.bg_color)
        canvas.pack(fill="both", expand=True)
        photo = ImageTk.PhotoImage(resized)
        canvas.create_image(0, 0, anchor="nw", image=photo)

        # Store references
        self.right_canvas = canvas
        self.right_img = photo
        self.image_scale = scale

        self._bind_resize_event(pil_page_image, self.right_frame, canvas)

        # --- Set pause flags for main loop ---
        self.waiting_for_manual_score = True
        self.manual_score_selection = None

    def prompt_manual_page_side(self, pil_page_image, detected_side, expected_side):
        """
        Show full page and ask the user to confirm if the page is correct or skip.
        """
        self.waiting_for_manual_page_side = True

        # --- Clear frames ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # --- Center frame: heading + instructions ---
        ttk.Label(
            self.center_frame,
            text="Check Page Side",
            font=("Helvetica", 14, "bold")
        ).pack(pady=(10, 6))

        ttk.Label(
            self.center_frame,
            text=(
                f"This page got flagged as being on the wrong side.\n"
                f"Detected: {detected_side}, Expected: {expected_side}\n\n"
                "Please decide whether to save the data to the current topics, or skip the page.\n"
                "If you choose to skip the page, you may want to write down the data so you can add it to the gradebook properly."
            ),
            wraplength=self.center_frame.winfo_width() - 20
        ).pack(pady=(0, 10))

        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=6)

        # --- Buttons ---
        def process_page():
            self.manual_page_side_decision = "process"
            self.waiting_for_manual_page_side = False

            # Unbind the resize callback and clear canvas references
            self.right_frame.unbind("<Configure>")
            self.right_canvas = None
            self.right_img = None


        def skip_page():
            self.manual_page_side_decision = "skip"
            self.waiting_for_manual_page_side = False

            # Unbind the resize callback and clear stale canvas references
            self.right_frame.unbind("<Configure>")
            self.right_canvas = None
            self.right_img = None

            self.skip_current_page()  # now safe to call


        ttk.Button(btn_frame, text="Page is correct, process normally", command=process_page).pack(side="left", padx=4)
        ttk.Button(btn_frame, text="Skip this Page", command=skip_page).pack(side="left", padx=4)

        # --- Stop Processing Label and Button ---
        ttk.Label(
            self.center_frame,
            text="This button will interrupt all data processing and take you back to the calibration stage.",
            wraplength=self.center_frame.winfo_width() - 10,  # wrap near the frame width
            justify="left"
        ).pack(pady=(20, 4))
        
        stop_btn = ttk.Button(
            self.center_frame,
            text="STOP ALL PROCESSING",
            style="Danger.TButton",  # optional custom style if you define one
            command=self._stop_processing  # placeholder for later integration
        )
        stop_btn.pack(pady=(4, 20))

        # --- Right frame: show full page ---
        resized, scale = self._resize_image_to_fit(pil_page_image, self.right_frame)
        canvas = tk.Canvas(self.right_frame, width=resized.width, height=resized.height, bg=self.bg_color)
        canvas.pack(fill="both", expand=True)
        photo = ImageTk.PhotoImage(resized)
        canvas.create_image(0, 0, anchor="nw", image=photo)
        canvas.photo = photo
        self.right_canvas = canvas
        self.right_img = photo
        self.image_scale = scale

        # --- Optional: handle resizing ---
        def redraw_canvas(event=None):
            if getattr(self, "right_canvas", None) is None or not self.right_canvas.winfo_exists():
                return  # canvas no longer exists, skip
            resized_new, scale_new = self._resize_image_to_fit(pil_page_image, self.right_frame)
            self.right_canvas.config(width=resized_new.width, height=resized_new.height)
            photo_new = ImageTk.PhotoImage(resized_new)
            self.right_canvas.delete("all")
            self.right_canvas.create_image(0, 0, anchor="nw", image=photo_new)
            self.right_canvas.photo = photo_new
            self.image_scale = scale_new

        self.right_frame.bind("<Configure>", lambda e: self._debounce(redraw_canvas, 150, e))


    def show_data_preview_with_export(self):
        """
        Displays extracted data in the right frame with Treeview and scrollbars,
        and shows download/export buttons in the center frame.
        """
        # --- Clear center frame and add heading/text ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()

        ttk.Label(self.center_frame, text="Extraction Complete", style="Header.TLabel").pack(pady=(10,6))
        instructions = ("On the right is a preview of the data.\n"
                        "Use the following button to save your data:")
        tk.Label(self.center_frame, text=instructions,
                 wraplength=self.center_frame.winfo_width()-20,
                 justify="left").pack(pady=(0,10))

        # --- Buttons in center frame ---
        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=(0,10))

        def export_csv():
            file_path = filedialog.asksaveasfilename(
                title="Save CSV File",
                defaultextension=".csv",
                filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
            )
            if file_path:
                columns = ["Name"] + [v.get().strip() for v in self.topic_vars if v.get().strip()]
                with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
                    writer = csv.writer(f)
                    writer.writerow(columns)
                    for row_id in self.tree.get_children():
                        row_values = self.tree.item(row_id)["values"]
                        writer.writerow(normalize_score_row(row_values))
                messagebox.showinfo("Exported", f"CSV exported to:\n{file_path}")
                self.mark_step_done("download")

        ttk.Button(btn_frame, text="Download CSV File", command=export_csv).pack(side="left", padx=5)

        def export_excel():
            file_path = filedialog.asksaveasfilename(
                title="Save Excel File",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
            )
            if file_path:
                columns = ["Name"] + [v.get().strip() for v in self.topic_vars if v.get().strip()]
                rows = [
                    normalize_score_row(self.tree.item(row_id)["values"])
                    for row_id in self.tree.get_children()
                ]
                try:
                    pd.DataFrame(rows, columns=columns).to_excel(file_path, index=False)
                    messagebox.showinfo("Exported", f"Excel file exported to:\n{file_path}")
                    self.mark_step_done("download")
                except ImportError:
                    messagebox.showerror("Excel Export", "Install the 'openpyxl' package to export .xlsx files.")

        ttk.Button(btn_frame, text="Download Excel File", command=export_excel).pack(side="left", padx=5)
        # --- Horizontal separator ---
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=10)
        
        def update_gradebook():
            self.update_full_gradebook()
            self.update_gradebook_btn.state(["disabled"])
            self._show_gradebook_popup(additional_text = "Your local gradebook has been updated")

        if self.enable_gradebook_var.get():
            tk.Label(self.center_frame,
                text = "Update the optional local gradebook stored on this computer.\n",
                    wraplength = self.center_frame.winfo_width()-10,
                    justify="left").pack(pady=(8, 2))
        
            self.update_gradebook_btn = ttk.Button(
                btn_frame,
                text="Update Local Gradebook",
                command=update_gradebook
            )
            self.update_gradebook_btn.pack(side="left", padx=5)

        
        # --- Horizontal separator ---
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=10)

        # --- Coffee button ---
        tk.Label(self.center_frame, text="Like this program?").pack(pady=(5,2))
        def open_venmo():
            webbrowser.open("https://www.venmo.com/u/KevinPCassidy1981")
        ttk.Button(self.center_frame, text="Buy Kevin a Coffee", command=open_venmo).pack(pady=(0,10))

        def update_gsheets():
            if self.update_gsheet_from_extracted_data():
                self.google_sync_button.state(["disabled"])

        class_info = self.classes.get(self.class_combo.get(), {})
        if (self.google_sheets_enabled_var.get()
                and isinstance(class_info, dict)
                and class_info.get("source") == "google_sheet"
                and not class_info.get("needs_remapping")):
            self.google_sync_button = ttk.Button(
                self.center_frame,
                text="Sync with Google Sheet",
                command=update_gsheets,
            )
            self.google_sync_button.pack(pady=(20,10))

        # --- Right frame: clear and insert Treeview ---
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        columns = ["Name"] + [v.get().strip() for v in self.topic_vars if v.get().strip()]
        self.tree = ttk.Treeview(self.right_frame, columns=columns, show="headings", selectmode="none")

        # Style
        style = ttk.Style(self.right_frame)
        style.configure("Treeview", rowheight=22, font=("Arial", 10))
        style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
        self.tree.tag_configure("line1", background="#ffffff")
        self.tree.tag_configure("line2", background="#f0f0f0")

        # Set column widths dynamically
        tree_font = tkfont.Font(family="Arial", size=10)
        col_widths = {col: tree_font.measure(col) for col in columns}

        # Populate rows: include all students from the roster
        name_to_data = {entry["name"]: entry for entry in self.extracted_data}
        roster_names_from_csv = self.roster_names
        
        # Remove header row if present
        if roster_names_from_csv and roster_names_from_csv[0].strip().lower() == "name":
            roster_names_from_csv = roster_names_from_csv[1:]
        
        # Add any students not already in roster_names_from_csv at the end
        extra_names = [entry["name"] for entry in self.extracted_data if entry["name"] not in roster_names_from_csv]
        all_names_to_show = roster_names_from_csv + extra_names

        for i, name in enumerate(all_names_to_show):
            student_entry = name_to_data.get(name, {})
            topics_dict = student_entry.get("topics", {})
            row_values = [name] + [topics_dict.get(topic, "") for topic in columns[1:]]
            tag = "line1" if i % 2 == 0 else "line2"
            self.tree.insert("", "end", values=row_values, tags=(tag,))

            # Update column widths
            for col, val in zip(columns, row_values):
                width = tree_font.measure(str(val))
                if width > col_widths[col]:
                    col_widths[col] = width

        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=col_widths[col]+10, anchor="center")

        # Scrollbars
        vsb = ttk.Scrollbar(self.right_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.right_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.tree.pack(fill="both", expand=True)

        # --- Enable cell editing in the Treeview ---

        self.editing_entry = None

        def begin_edit(event):
            # Identify the row and column clicked
            row_id = self.tree.identify_row(event.y)
            col_id = self.tree.identify_column(event.x)
            if not row_id or not col_id:
                return

            col_index = int(col_id.replace("#", "")) - 1  # convert '#3' -> 2

            # Prevent editing the Name column if desired:
            # if col_index == 0:
            #     return

            # Get cell bounding box (x,y,width,height)
            x, y, w, h = self.tree.bbox(row_id, col_id)
            if w <= 0:
                return

            # Existing value
            old_value = self.tree.item(row_id)["values"][col_index]

            # Create entry overlay
            entry = tk.Entry(self.right_frame)
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, old_value)
            entry.focus()

            self.editing_entry = entry

            def finish_edit(*_):
                new_value = entry.get()

                # Update Treeview values
                values = list(self.tree.item(row_id)["values"])
                values[col_index] = new_value
                self.tree.item(row_id, values=values)

                entry.destroy()
                self.editing_entry = None
                
                # Re-enable the update button if it exists
                if hasattr(self, "update_gradebook_btn"):
                    self.update_gradebook_btn.state(["!disabled"])
                if hasattr(self, "google_sync_button"):
                    self.google_sync_button.state(["!disabled"])
                
                # ---- UPDATE extracted_data to match the edited value ----
                """Note - this is for my personal version to integrate extracted data for update to google sheets"""
                
                row_values = values  # already updated list: [name, topic1score, topic2score...]
                student_name = row_values[0]
                topic_name = columns[col_index]  # use your columns list

                # Find the correct entry in extracted_data
                for entry in self.extracted_data:
                    if entry["name"] == student_name:
                        # Make sure "topics" exists
                        if "topics" not in entry:
                            entry["topics"] = {}

                        # Update the topic value
                        entry["topics"][topic_name] = new_value
                        break
                """End update for my personal version"""                

            entry.bind("<Return>", finish_edit)
            entry.bind("<FocusOut>", finish_edit)

        # Bind double-click to start editing
        self.tree.bind("<ButtonRelease-1>", begin_edit)

        # Mouse wheel support
        def _on_mousewheel(event):
            try:
                self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except (tk.TclError, AttributeError):
                return "break"

        self.tree.bind("<MouseWheel>", _on_mousewheel)
        
                # Horizontal line
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=(15, 5))
        
        # Button to return to original center panel
        ttk.Button(self.center_frame, text="Home", command=lambda: self.reset_panels()).pack(pady=8)


    def _show_gradebook_popup(self, additional_text):
        """Pop-up after gradebook update with option to open the gradebook."""
        popup = tk.Toplevel(self.root)
        popup.title("Open Local Gradebook")
        popup.geometry("350x350")
        popup.grab_set()  # make modal
        
        update_text = additional_text
        ttk.Label(
            popup,
            text=(
                update_text+"\n\n"
                "You may open it directly to view or edit.\n"
                "Any edits you make will affect the actual gradebook file.\n\n"
                "You can access this gradebook again later via\n"
                "the Local Gradebook button under Preferences.\n\n"
                "Please note this is a .CSV file. To use any formulas or\n"
                "To adjust formatting, please save it as an excel file.\n\n"
            ),
            wraplength=320,
            justify="center"
        ).pack(pady=(20, 10))

        def open_gradebook():
            csv_path = getattr(self, "full_gradebook_path", None) or os.path.join(
                os.getcwd(),
                f"{self.view_class_combo.get().replace(' ', '_')}_gradebook.csv"
            )

            # --- Sync gradebook with roster before opening ---
            # --- Sync gradebook with roster before opening ---
            if os.path.exists(csv_path):
                roster_name = os.path.basename(csv_path).replace("_gradebook.csv", "")
                try:
                    df_existing = pd.read_csv(csv_path)
                    df_existing = self._sync_gradebook_with_roster(df_existing, roster_name)
                    df_existing.to_csv(csv_path, index=False)
                except PermissionError:
                    import tkinter.messagebox as messagebox
                    messagebox.showerror(
                        "Local Gradebook Update Error",
                        "Your gradebook is currently open in another program.\n"
                        "Please close the file and try again."
                    )
                    return  # stop further execution

                
            if csv_path and os.path.exists(csv_path):
                try:
                    # Open actual CSV in default program
                    try:
                        os.startfile(csv_path)  # Windows
                    except AttributeError:
                        if platform.system() == "Darwin":
                            subprocess.call(("open", csv_path))
                        else:
                            subprocess.call(("xdg-open", csv_path))
                    popup.destroy()

                except Exception as e:
                    messagebox.showerror("Error", f"Failed to open gradebook:\n{e}")

            else:
                messagebox.showwarning("Missing File", "The gradebook file could not be found.")

        ttk.Button(popup, text="Open Local Gradebook", command=open_gradebook).pack(pady=(0, 10))
        ttk.Button(popup, text="Close", command=popup.destroy).pack()


    def update_full_gradebook(self):
        """
        Append extracted data to the persistent gradebook for the selected roster.
        Updates existing students' scores and adds any new students.
        Keeps 'Name' as a column throughout.
        """

        # Determine gradebook file
        roster_name = self.class_combo.get().replace(" ", "_")  # replace spaces with underscores
        if roster_name == "-- Select Class --" or not roster_name:
            roster_name = "default_course"

        gradebook_name = f"{roster_name}_gradebook.csv"
        gradebook_path = os.path.join(os.getcwd(), gradebook_name)

        # Build current run’s data from the Treeview
        columns = ["Name"] + [v.get().strip() for v in self.topic_vars if v.get().strip()]
        new_data = [
            normalize_score_row(self.tree.item(row_id)["values"])
            for row_id in self.tree.get_children()
        ]
        df_new = pd.DataFrame(new_data, columns=columns)

        try:
            if os.path.exists(gradebook_path):
                # Load existing gradebook
                df_existing = pd.read_csv(gradebook_path)

                # Sync with roster (this now handles _from_extraction internally)
                df_existing = self._sync_gradebook_with_roster(df_existing, roster_name)

                # Ensure all columns exist
                for col in df_new.columns:
                    if col not in df_existing.columns:
                        df_existing[col] = ""

                # Update existing rows
                for i, row in df_new.iterrows():
                    name = row["Name"]
                    if name in df_existing["Name"].values:
                        # Update only the existing student
                        for col in df_new.columns[1:]:
                            df_existing.loc[df_existing["Name"] == name, col] = row[col]
                    else:
                        # Append new student
                        df_existing = pd.concat([df_existing, pd.DataFrame([row])], ignore_index=True)

                df_combined = df_existing

            else:
                df_combined = df_new

            # Save the gradebook
            df_combined.to_csv(gradebook_path, index=False)
            self.full_gradebook_path = gradebook_path
            print(f"[INFO] Full gradebook updated: {gradebook_path}")

        except PermissionError:
            messagebox.showerror(
                "Local Gradebook Update Error",
                "Your gradebook is already open on your computer, so the program cannot update it.\n"
                "Please close the file, then click 'Update Gradebook' again."
            )
            return


    def _sync_gradebook_with_roster(self, df_gradebook, roster_name):
        """
        Ensure the full gradebook rows match the current roster.
        Adds missing students in roster order and removes students not in roster.
        Extra rows from extraction are preserved at the bottom.
        """

        class_name = roster_name.replace('_', ' ')
        roster_path = self._resolve_class_path(class_name)

        if not roster_path or not os.path.exists(roster_path):
            print(f"[WARN] Roster file not found: {roster_path}")
            return df_gradebook

        try:
            df_roster = pd.DataFrame({"Name": read_roster_names(roster_path)})
        except ValueError as error:
            print(f"[ERROR] {error}: {roster_path}")
            return df_gradebook

        if "Name" not in df_roster.columns:
            print(f"[ERROR] 'Name' column missing from roster: {roster_path}")
            return df_gradebook

        # --- Mark rows that are not in the roster (from extraction) ---
        df_gradebook["_from_extraction"] = ~df_gradebook["Name"].isin(df_roster["Name"])

        # --- Keep only students who are in the roster or extra extraction rows ---
        df_gradebook = df_gradebook[df_gradebook["_from_extraction"] | df_gradebook["Name"].isin(df_roster["Name"])]

        # --- Add missing roster students ---
        missing_students = df_roster.loc[~df_roster["Name"].isin(df_gradebook["Name"]), "Name"]
        for student in missing_students:
            new_row = {"Name": student}
            for col in df_gradebook.columns:
                if col != "Name" and col != "_from_extraction":
                    new_row[col] = None
            df_gradebook = pd.concat([df_gradebook, pd.DataFrame([new_row])], ignore_index=True)

        # --- Reorder roster students, keep extraction-only rows at bottom ---
        df_gradebook = pd.concat([
            df_gradebook[df_gradebook["_from_extraction"] == False].sort_values(
                by="Name",
                key=lambda x: x.map({name: i for i, name in enumerate(df_roster["Name"])})
            ),
            df_gradebook[df_gradebook["_from_extraction"] == True]
        ]).drop(columns=["_from_extraction"], errors="ignore").reset_index(drop=True)

        return df_gradebook




    def _cleanup_temp_gradebook_copies(self):
        """Delete any temporary gradebook copies created during this session."""
        import os
        if hasattr(self, "_temp_gradebook_copies"):
            for temp_file in self._temp_gradebook_copies:
                try:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)
                        print(f"[INFO] Deleted temporary gradebook copy: {temp_file}")
                except Exception as e:
                    print(f"[WARN] Could not delete {temp_file}: {e}")

    def _on_view_gradebook(self):
        """Display the gradebook viewer UI with options to view, open, or delete."""
        # --- Clear center frame ---
        for widget in self.center_frame.winfo_children():
            widget.destroy()
            
        # --- Right frame: clear
        for widget in self.right_frame.winfo_children():
            widget.destroy()
        
        self._build_right_panel()

        # --- Heading ---
        ttk.Label(self.center_frame, text="View Local Gradebook", style="Header.TLabel").pack(pady=(10, 6))

        # --- Instructions ---
        instructions = "Select which class you would like to view."
        ttk.Label(self.center_frame, text=instructions,
                  wraplength=self.center_frame.winfo_width() - 20,
                  justify="left").pack(pady=(0, 10))

        # --- Class dropdown ---
        ttk.Label(self.center_frame, text="Select Class:", style="Bold.TLabel").pack(anchor="w", pady=(6, 2))
        class_names = ["-- Select Class --"] + list(self.classes.keys())
        self.view_class_combo = ttk.Combobox(self.center_frame, values=class_names, state="readonly")
        self.view_class_combo.current(0)
        self.view_class_combo.pack(fill="x", pady=(0, 10))

        # --- Button frame ---
        btn_frame = ttk.Frame(self.center_frame)
        btn_frame.pack(pady=(5, 10))

        # --- View Gradebook Button ---
        def view_gradebook():

            selected_class = self.view_class_combo.get().replace(" ", "_")
            if selected_class == "-- Select Class --" or not selected_class:
                messagebox.showwarning("No Class Selected", "Please select a class to view the gradebook.")
                return

            gradebook_file = os.path.join(os.getcwd(), f"{selected_class}_gradebook.csv")
            if not os.path.exists(gradebook_file):
                messagebox.showwarning("Missing File", f"No gradebook exists for class: {selected_class}")
                return
            # --- Sync gradebook with current roster ---
            df_existing = pd.read_csv(gradebook_file)
            df_existing = self._sync_gradebook_with_roster(df_existing, selected_class)
            df_existing.to_csv(gradebook_file, index=False)
            

            # --- Right frame: clear previous ---
            for widget in self.right_frame.winfo_children():
                widget.destroy()

            # Load CSV
            df = pd.read_csv(gradebook_file)
            columns = list(df.columns)

            self.tree = ttk.Treeview(self.right_frame, columns=columns, show="headings", selectmode="none")

            # Style
            style = ttk.Style(self.right_frame)
            style.configure("Treeview", rowheight=22, font=("Arial", 10))
            style.configure("Treeview.Heading", font=("Arial", 10, "bold"))
            self.tree.tag_configure("line1", background="#ffffff")
            self.tree.tag_configure("line2", background="#f0f0f0")

            # Populate rows
            for i, row in df.iterrows():
                row_values = ["" if pd.isna(row[col]) else row[col] for col in columns]
                tag = "line1" if i % 2 == 0 else "line2"
                self.tree.insert("", "end", values=row_values, tags=(tag,))

            # Set column widths dynamically
            tree_font = tkfont.Font(family="Arial", size=10)
            for col in columns:
                max_width = max([tree_font.measure(str(val)) for val in df[col]] + [tree_font.measure(col)])
                self.tree.heading(col, text=col)
                self.tree.column(col, width=max_width + 10, anchor="center")

            # Scrollbars
            vsb = ttk.Scrollbar(self.right_frame, orient="vertical", command=self.tree.yview)
            hsb = ttk.Scrollbar(self.right_frame, orient="horizontal", command=self.tree.xview)
            self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
            vsb.pack(side="right", fill="y")
            hsb.pack(side="bottom", fill="x")
            self.tree.pack(fill="both", expand=True)

            # --- Enable cell editing in the Treeview ---

            self.editing_entry = None

            def begin_edit(event):
                # Identify the row and column clicked
                row_id = self.tree.identify_row(event.y)
                col_id = self.tree.identify_column(event.x)
                if not row_id or not col_id:
                    return

                col_index = int(col_id.replace("#", "")) - 1  # convert '#3' -> 2

                # Prevent editing the Name column if desired:
                # if col_index == 0:
                #     return

                # Get cell bounding box (x,y,width,height)
                x, y, w, h = self.tree.bbox(row_id, col_id)
                if w <= 0:
                    return

                # Existing value
                old_value = self.tree.item(row_id)["values"][col_index]

                # Create entry overlay
                entry = tk.Entry(self.right_frame)
                entry.place(x=x, y=y, width=w, height=h)
                entry.insert(0, old_value)
                entry.focus()

                self.editing_entry = entry

                def finish_edit(*_):
                    new_value = entry.get()

                    # Update Treeview values
                    values = list(self.tree.item(row_id)["values"])
                    values[col_index] = new_value
                    self.tree.item(row_id, values=values)

                    entry.destroy()
                    self.editing_entry = None
                    
                    if self.enable_gradebook_var.get():
                        # Automatically update gradebook CSV
                        new_data = [self.tree.item(row)["values"] for row in self.tree.get_children()]
                        df = pd.DataFrame(new_data, columns=self.tree["columns"])
                        df.to_csv(gradebook_file, index=False)


                entry.bind("<Return>", finish_edit)
                entry.bind("<FocusOut>", finish_edit)

            # Bind double-click to start editing
            self.tree.bind("<ButtonRelease-1>", begin_edit)


            # Mousewheel support
            def _on_mousewheel(event):
                self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            self.tree.bind("<MouseWheel>", _on_mousewheel)

        ttk.Button(btn_frame, text="View Local Gradebook", command=view_gradebook).pack(side="left", padx=5)

        # --- Open Editable File Button ---
        ttk.Button(btn_frame, text="Open Editable File", command=lambda: self._show_gradebook_popup(additional_text = "This links to your most recently saved gradebook.")).pack(side="left", padx=5)

        # --- Delete Gradebook Button ---
        def delete_gradebook():
            selected_class = self.view_class_combo.get().replace(" ", "_")
            if selected_class == "-- Select Class --" or not selected_class:
                messagebox.showwarning("No Class Selected", "Please select a class to delete the gradebook.")
                return

            def confirm_delete():
                gradebook_file = os.path.join(os.getcwd(), f"{selected_class}_gradebook.csv")
                try:
                    if os.path.exists(gradebook_file):
                        os.remove(gradebook_file)
                    messagebox.showinfo("Deleted", f"Gradebook for class {selected_class} has been deleted.")
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to delete gradebook:\n{e}")
                confirm_popup.destroy()

            confirm_popup = tk.Toplevel(self.root)
            confirm_popup.title("Confirm Delete")
            confirm_popup.geometry("400x150")
            confirm_popup.grab_set()  # modal

            ttk.Label(confirm_popup, text="Please confirm you want to delete the gradebook from the program.",
                      wraplength=380, justify="center").pack(pady=(20, 10))

            btn_frame_popup = ttk.Frame(confirm_popup)
            btn_frame_popup.pack(pady=10)

            ttk.Button(btn_frame_popup, text="Cancel", command=confirm_popup.destroy).pack(side="left", padx=10)
            ttk.Button(btn_frame_popup, text="Confirm", command=confirm_delete).pack(side="left", padx=10)

        ttk.Button(btn_frame, text="Delete Local Gradebook", command=delete_gradebook).pack(side="left", padx=5)
        
        # Horizontal line
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=(15, 5))
        
        # Button to return to original center panel
        ttk.Button(self.center_frame, text="Home", command=lambda: self.reset_panels()).pack(pady=8)


        

    def reset_panels(self):
        self._clear_cached_images()

        # Destroy all children recursively
        for child in self.right_frame.winfo_children():
            child.destroy()

        # Rebuild right panel
        self._build_right_panel()

        # Destroy and rebuild center frame as usual
        for child in self.center_frame.winfo_children():
            child.destroy()
        self._build_center_panel()

    def _clear_cached_images(self):
        image_attrs = ["right_full_image"]
        for attr in image_attrs:
            if hasattr(self, attr):
                image = getattr(self, attr)
                if hasattr(image, "close"):
                    try:
                        image.close()
                    except Exception:
                        pass
                setattr(self, attr, None)

        self.right_photo = None
        self.right_img = None
        self.right_canvas = None
        self.current_page_index = None
        self.skipped_pages = []


    # ---------------- RIGHT PANEL ----------------
    def _build_right_panel(self):
        self.right_placeholder = ttk.Label(
            self.right_frame,
            text="Calibration / Data Preview Area",
            anchor="center",
            font=("Segoe UI", 14),
            relief="ridge",
            padding=40
        )
        self.right_placeholder.pack(fill="both", expand=True)
    
    def _show_class_list_panel(self):
        # Clear previous content
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        # Create panel for classes
        frame = ttk.Frame(self.right_frame)
        frame.pack(fill="both", expand=True)

        #Header
        ttk.Label(frame, text="Classes", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))


        # Treeview for classes
        columns = ("Class Name", "Source", "Roster / Tab")
        self.class_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.class_tree.heading("Class Name", text="Class Name")
        self.class_tree.heading("Source", text="Source")
        self.class_tree.heading("Roster / Tab", text="Roster / Tab")
        self.class_tree.pack(fill="both", expand=True, side="left")

        self.class_tree.bind("<Double-1>", self._preview_class_students)

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.class_tree.yview)
        self.class_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self._refresh_classes_tree()

    def _preview_class_students(self, event):
        # Get selected class
        selected = self.class_tree.selection()
        if not selected:
            return
        class_name = self.class_tree.item(selected[0], "values")[0]
        csv_path = self._resolve_class_path(class_name)
        if not csv_path or not os.path.exists(csv_path):
            messagebox.showerror("Error", f"CSV file for class '{class_name}' not found.")
            return

        # Create preview window
        win = tk.Toplevel(self.root)
        win.title(f"Preview: {class_name}")
        win.geometry(f"300x500+{int(self.root.winfo_screenwidth()/2 - 100)}+{int(self.root.winfo_screenheight()/4)}")

        # Treeview for student names
        tree_frame = ttk.Frame(win)
        tree_frame.pack(fill="both", expand=True)

        columns = ("Student Name",)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=20)
        tree.heading("Student Name", text="Student Name")
        tree.column("Student Name", width=180, anchor="w")
        tree.pack(side="left", fill="both", expand=True)

        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        try:
            for name in read_roster_names(csv_path):
                tree.insert("", "end", values=(name,))
        except ValueError as error:
            win.destroy()
            messagebox.showerror("Roster Format Error", str(error))
            return


        # Mousewheel binding
        def _on_mousewheel(event):
            try:
                tree = event.widget
                tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                # Widget was destroyed — ignore
                pass
            
            
        tree.bind_all("<MouseWheel>", _on_mousewheel)

    
    def _refresh_classes_tree(self):
        """Redraw all classes in the right Treeview."""
        if not hasattr(self, "class_tree"):
            return

        # Clear previous rows
        for row in self.class_tree.get_children():
            self.class_tree.delete(row)

        # Insert current classes
        for class_name, class_info in self.classes.items():
            if isinstance(class_info, dict):
                is_google = class_info.get("source") == "google_sheet"
                source = "Google Sheets" if is_google else "Local CSV"
                location = class_info.get("worksheet_title") if is_google else os.path.basename(class_info.get("roster_path", ""))
            else:
                source, location = "Local CSV", os.path.basename(class_info)
            self.class_tree.insert("", "end", values=(class_name, source, location))

    def _show_grading_scales_panel(self):
        # Clear right panel
        for widget in self.right_frame.winfo_children():
            widget.destroy()

        frame = ttk.Frame(self.right_frame)
        frame.pack(fill="both", expand=True)

        # Heading
        ttk.Label(frame, text="Grading Scales", font=("Segoe UI", 14, "bold")).pack(pady=(10,5))

        # Treeview
        columns = ("Scale Name", "Valid Scores")
        self.grading_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.grading_tree.heading("Scale Name", text="Scale Name")
        self.grading_tree.heading("Valid Scores", text="Valid Scores")
        self.grading_tree.pack(fill="both", expand=True, side="left")

        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=self.grading_tree.yview)
        self.grading_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        self._refresh_grading_tree()


        if not hasattr(self, "grading_tree"):
            return
        # Clear previous rows
        for row in self.grading_tree.get_children():
            self.grading_tree.delete(row)

        # Insert current grading scales
        if hasattr(self, "grading_scales"):
            for name, scores in self.grading_scales.items():
                score_str = ", ".join(
                    "Skip" if isinstance(s, str) and s.strip().lower() == "skip"
                    else (str(int(s)) if float(s).is_integer() else str(s))
                    for s in scores
                )
                self.grading_tree.insert("", "end", values=(name, score_str))

    def _refresh_grading_tree(self):
        if not hasattr(self, "grading_tree"):
            return

        # Clear existing rows
        for row in self.grading_tree.get_children():
            self.grading_tree.delete(row)

        # Insert current grading scales
        if hasattr(self, "grading_scales"):
            for name, scores in self.grading_scales.items():
                score_str = ", ".join(
                    "Skip" if isinstance(s, str) and s.strip().lower() == "skip"
                    else (str(int(s)) if float(s).is_integer() else str(s))
                    for s in scores
                )
                self.grading_tree.insert("", "end", values=(name, score_str))

    def _display_pdf_first_page(self):
        """Display the first page of the selected PDF as a scaled image on the right frame, 
        with live resize support that preserves calibration scaling."""
        pdf_path = self.pdf_path_var.get()
        if not pdf_path or not os.path.exists(pdf_path):
            ttk.Label(self.right_frame, text="No PDF selected.", foreground="red").pack(pady=20)
            return

        try:
            # --- Convert first page to image ---
            pages = convert_from_path(
                pdf_path,
                dpi=200,
                first_page=1,
                last_page=1,
                poppler_path=self.poppler_bin,
            )
            first_page = pages[0]
            self.right_full_image = first_page  # full-res reference

            # --- Clear frame and prepare canvas ---
            for widget in self.right_frame.winfo_children():
                widget.destroy()

            # --- Initial scaled render ---
            resized, scale = self._resize_image_to_fit(self.right_full_image, self.right_frame)
            self.image_scale = scale

            photo = ImageTk.PhotoImage(resized)
            canvas = tk.Canvas(self.right_frame, width=resized.width, height=resized.height, bg=self.bg_color)
            canvas.pack(fill="both", expand=True)
            canvas.create_image(0, 0, anchor="nw", image=photo)
            canvas.photo = photo  # prevent garbage collection

            # --- Store references ---
            self.right_canvas = canvas
            self.right_img = photo

            # --- Bind live resize updates (debounced) ---
            def debounced_resize(event):
                # skip if frame content changed (no longer showing image)
                if not hasattr(self, "right_full_image") or not hasattr(self, "right_canvas"):
                    return
                try:
                    resized_img, new_scale = self._resize_image_to_fit(self.right_full_image, self.right_frame)
                    self.image_scale = new_scale
                    new_photo = ImageTk.PhotoImage(resized_img)
                    self.right_canvas.delete("all")
                    self.right_canvas.create_image(0, 0, anchor="nw", image=new_photo)
                    self.right_canvas.photo = new_photo
                except Exception:
                    pass  # silently ignore if frame destroyed during resize

            self.right_frame.bind("<Configure>", lambda e: self._debounce(debounced_resize, 150, e))

        except Exception as e:
            ttk.Label(self.right_frame, text=f"Error loading PDF:\n{e}", foreground="red").pack(pady=20)


    def _enable_box_drawing(self, canvas, topic_name=None, color=None):
        """
        Allows the user to draw or adjust a single box on the canvas.
        Name box if topic_name=None, otherwise a topic score box.
        Stores coordinates in self.name_box or self.score_boxes[topic_name].
        Automatically handles resizing so drawn boxes stay aligned.
        """

        # --- Cleanup before enabling new drawing ---
        # Unbind resize events to prevent interference while drawing
        if hasattr(self, "_resize_after_id") and self._resize_after_id:
            self.right_frame.after_cancel(self._resize_after_id)
            self._resize_after_id = None
        self.right_frame.unbind("<Configure>")

        # Remove any existing box if present
        if topic_name is None and getattr(self, "name_box", None):
            if hasattr(self, "name_box_rect"):
                canvas.delete(self.name_box_rect)
            if hasattr(self, "name_box_label"):
                canvas.delete(self.name_box_label)
        elif topic_name and topic_name in self.score_boxes:
            rect_attr = f"score_box_rect_{topic_name}"
            label_attr = f"score_box_label_{topic_name}"
            if hasattr(self, rect_attr):
                canvas.delete(getattr(self, rect_attr))
            if hasattr(self, label_attr):
                canvas.delete(getattr(self, label_attr))

        if topic_name == "Page Side":
            color = "blue"
            label_text = "Page Side Area"
        elif topic_name is None:
            color = color or "red"
            label_text = "Name Box"
        else:
            color = color or "green"
            label_text = f"{topic_name} Score Area"


        # Temporary holder during drawing
        box_info = {'start': None, 'rect': None, 'label': None}

        # --- Drawing Handlers ---
        def on_press(event):
            box_info['start'] = (canvas.canvasx(event.x), canvas.canvasy(event.y))

        def on_drag(event):
            if box_info['start'] is None:
                return
            x0, y0 = box_info['start']
            x1, y1 = canvas.canvasx(event.x), canvas.canvasy(event.y)
            if box_info['rect']:
                canvas.delete(box_info['rect'])
            if box_info['label']:
                canvas.delete(box_info['label'])
            box_info['rect'] = canvas.create_rectangle(x0, y0, x1, y1, outline=color, width=2)
            box_info['label'] = canvas.create_text(
                (x0 + x1) // 2,
                min(y0, y1) - 10,
                text=label_text,
                fill=color,
                font=("TkDefaultFont", 10, "bold")
            )

        def on_release(event):
            if box_info['start'] is None:
                return
            x0, y0 = box_info['start']
            x1, y1 = canvas.canvasx(event.x), canvas.canvasy(event.y)
            coords = (int(min(x0, x1)), int(min(y0, y1)), int(max(x0, x1)), int(max(y0, y1)))

            scale_canvas = getattr(self, "image_scale", 1.0) or 1.0
            x0_full = int(coords[0] / scale_canvas)
            y0_full = int(coords[1] / scale_canvas)
            x1_full = int(coords[2] / scale_canvas)
            y1_full = int(coords[3] / scale_canvas)


            # Store results
            if topic_name is None:
                # Name box
                self.name_box = coords
                self.name_box_rect = box_info['rect']
                self.name_box_label = box_info['label']
                self.calibration_data["name_box"] = (x0_full, y0_full, x1_full, y1_full)
                self._update_next_button_state()

            elif topic_name == "Page Side":
                # Special Page Side box
                self.score_boxes[topic_name] = coords
                setattr(self, f"score_box_rect_{topic_name}", box_info['rect'])
                setattr(self, f"score_box_label_{topic_name}", box_info['label'])
                self.calibration_data["score_boxes"][topic_name] = (x0_full, y0_full, x1_full, y1_full)
                self._update_next_button_state()

                #print(f"[DEBUG] Page Side box saved: {self.calibration_data['score_boxes'][topic_name]}")
            else:
                # Regular score boxes
                self.score_boxes[topic_name] = coords
                setattr(self, f"score_box_rect_{topic_name}", box_info['rect'])
                setattr(self, f"score_box_label_{topic_name}", box_info['label'])
                self.calibration_data["score_boxes"][topic_name] = (x0_full, y0_full, x1_full, y1_full)
                self._update_next_button_state()

                #print(f"[DEBUG] Score box '{topic_name}' saved (full coords): {self.calibration_data['score_boxes'][topic_name]}")



            # Re-bind resize handler safely after drawing
            if self.right_full_image is not None and self.right_canvas is not None:
                self._bind_resize_event(self.right_full_image, self.right_frame, self.right_canvas)


        # --- Bind mouse events ---
        canvas.bind("<ButtonPress-1>", on_press)
        canvas.bind("<B1-Motion>", on_drag)
        canvas.bind("<ButtonRelease-1>", on_release)



    def _pil_image_for_canvas(self, pil_img, max_w=1000, max_h=1400):
        """Return a resized PIL image and the scale factor used (for canvas display)."""
        img = pil_img.copy()
        w, h = img.size
        scale = min(1.0, max_w / w, max_h / h)
        if scale < 1.0:
            img = img.resize((int(w*scale), int(h*scale)))
        return img, scale

    def _resize_image_to_fit(self, pil_image, frame):
        """Resize PIL image to fit inside the given frame while maintaining aspect ratio."""
        frame.update_idletasks()
        frame_w = frame.winfo_width() or 800
        frame_h = frame.winfo_height() or 800

        img_w, img_h = pil_image.size
        scale = min(frame_w / img_w, frame_h / img_h)
        new_w, new_h = int(img_w * scale), int(img_h * scale)

        resized = pil_image.resize((new_w, new_h))
        return resized, scale

    def _bind_resize_event(self, pil_image, frame, widget, is_center=False, redraw_overlays=None):
        """
        Bind a resize event with debounce to dynamically resize a displayed image.
        Also rescales any drawn boxes or overlays on the canvas.

        Args:
            pil_image (PIL.Image): Original unscaled image.
            frame (ttk.Frame): Parent frame to monitor for resize.
            widget (tk.Canvas | ttk.Label): Widget displaying the image.
            is_center (bool): If True, updates self.image_scale for calibration clicks.
            redraw_overlays (callable): Optional function(widget, scale) to redraw extra overlays.
        """
        resize_job = None  # holds ID of pending resize callback

        def do_resize():
            nonlocal resize_job
            resize_job = None  # clear pending job

            # Safety: skip if widget has been destroyed
            if not widget.winfo_exists() or not frame.winfo_exists():
                return

            try:
                resized, scale = self._resize_image_to_fit(pil_image, frame)
                photo = ImageTk.PhotoImage(resized)

                if isinstance(widget, tk.Canvas):
                    widget.delete("all")
                    widget.config(width=resized.width, height=resized.height)
                    widget.create_image(0, 0, anchor="nw", image=photo)
                    widget.image = photo

                    # --- redraw boxes if they exist ---
                    if hasattr(self, "calibration_data"):
                        # Name box
                        if "name_box" in self.calibration_data and self.calibration_data["name_box"]:
                            nb = self.calibration_data["name_box"]
                            x0, y0, x1, y1 = nb
                            x0_s, y0_s = int(x0 * scale), int(y0 * scale)
                            x1_s, y1_s = int(x1 * scale), int(y1 * scale)
                            rect = widget.create_rectangle(x0_s, y0_s, x1_s, y1_s, outline="red", width=2)
                            label = widget.create_text((x0_s + x1_s)//2, y0_s - 10,
                                                       text="Name Box", fill="red",
                                                       font=("TkDefaultFont", 10, "bold"))
                            self.name_box_rect, self.name_box_label = rect, label

                        # Score boxes
                        if "score_boxes" in self.calibration_data:
                            for topic, coords in self.calibration_data["score_boxes"].items():
                                x0, y0, x1, y1 = coords
                                x0_s, y0_s = int(x0 * scale), int(y0 * scale)
                                x1_s, y1_s = int(x1 * scale), int(y1 * scale)
                                rect = widget.create_rectangle(x0_s, y0_s, x1_s, y1_s, outline="green", width=2)
                                label = widget.create_text((x0_s + x1_s)//2, y0_s - 10,
                                                           text=f"{topic} Score Area", fill="green",
                                                           font=("TkDefaultFont", 10, "bold"))
                                setattr(self, f"score_box_rect_{topic}", rect)
                                setattr(self, f"score_box_label_{topic}", label)

                    # Optional user overlays
                    if callable(redraw_overlays):
                        redraw_overlays(widget, scale)

                else:
                    # For Label or other widgets
                    if widget.winfo_exists():
                        widget.config(image=photo)
                        widget.image = photo

                if is_center:
                    self.image_scale = scale

            except Exception as e:
                print(f"[DEBUG] Resize event failed: {e}")

        def on_resize(event):
            nonlocal resize_job
            if resize_job is not None:
                try:
                    frame.after_cancel(resize_job)
                except Exception:
                    pass
            resize_job = frame.after(200, do_resize)

        # Bind the resize event
        frame.bind("<Configure>", on_resize)

        
    def _debounce(self, func, delay=150, *args):
        """
        Debounce a callback — works even if QuizAppGUI isn't a Tk widget.
        """
        root = getattr(self, "root", None)
        if root is None:
            return func(*args)  # fallback if no root yet

        if not hasattr(self, "_debounce_jobs"):
            self._debounce_jobs = {}

        job_id = self._debounce_jobs.get(func)
        if job_id:
            try:
                root.after_cancel(job_id)
            except Exception:
                pass

        self._debounce_jobs[func] = root.after(delay, func, *args)

    #Personal update to Google Sheets:

    def get_gsheet_client(self, interactive=False, authorization_url_callback=None):
        """Return a user-authorized client; never use a distributed service account."""
        credentials = None
        if os.path.exists(self.token_file):
            try:
                credentials = Credentials.from_authorized_user_file(self.token_file, SCOPES)
            except (ValueError, json.JSONDecodeError):
                credentials = None
        if credentials and not credentials.has_scopes(SCOPES):
            credentials = None
        if credentials and credentials.expired and credentials.refresh_token:
            try:
                credentials.refresh(Request())
                atomic_write_json(self.token_file, json.loads(credentials.to_json()))
            except RefreshError:
                credentials = None
        if not credentials or not credentials.valid:
            if not interactive:
                raise PermissionError("Google authorization is missing or has expired.")
            if not os.path.exists(self.oauth_client_file):
                raise FileNotFoundError(
                    f"OAuth client configuration not found: {self.oauth_client_file}"
                )
            flow = InstalledAppFlow.from_client_secrets_file(self.oauth_client_file, SCOPES)
            browser_name = None
            if authorization_url_callback:
                class AuthorizationBrowser:
                    def open(self, url, new=0, autoraise=True):
                        authorization_url_callback(url)
                        return webbrowser.open(url, new=new, autoraise=autoraise)

                browser_name = f"quiz-processing-system-{id(flow)}"
                webbrowser.register(browser_name, None, AuthorizationBrowser())
            credentials = flow.run_local_server(
                port=0,
                open_browser=True,
                browser=browser_name,
                timeout_seconds=300,
                authorization_prompt_message=None,
            )
            atomic_write_json(self.token_file, json.loads(credentials.to_json()))
        return gspread.authorize(credentials)

    def _authorize_google_interactive(self):
        popup = tk.Toplevel(self.root)
        popup.title("Connect Google Sheets")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("720x270")
        ttk.Label(
            popup,
            text=("A browser window should open for Google authorization. If your browser window doesn't open, "
                  "copy and paste this link in a browser window for the Google Account you want to connect."),
            wraplength=670,
            justify="left",
        ).pack(padx=20, pady=(20, 10))
        authorization_url_var = tk.StringVar(value="Preparing authorization link…")
        link_entry = ttk.Entry(popup, textvariable=authorization_url_var, state="readonly")
        link_entry.pack(fill="x", padx=20, pady=5)
        status_var = tk.StringVar(value="Waiting for Google authorization…")
        ttk.Label(popup, textvariable=status_var).pack(pady=5)
        buttons = ttk.Frame(popup)
        buttons.pack(pady=10)

        def copy_link():
            url = authorization_url_var.get()
            if url.startswith("http"):
                self.root.clipboard_clear()
                self.root.clipboard_append(url)
                status_var.set("Authorization link copied. Paste it into the desired browser profile.")

        ttk.Button(buttons, text="Copy Link", command=copy_link).pack(side="left", padx=5)
        ttk.Button(
            buttons,
            text="Open in Default Browser",
            command=lambda: webbrowser.open(authorization_url_var.get()) if authorization_url_var.get().startswith("http") else None,
        ).pack(side="left", padx=5)
        ttk.Button(buttons, text="Close", command=popup.destroy).pack(side="left", padx=5)

        def show_url(url):
            self.root.after(0, lambda: authorization_url_var.set(url) if popup.winfo_exists() else None)

        def worker():
            try:
                self.get_gsheet_client(interactive=True, authorization_url_callback=show_url)
                self.root.after(0, lambda: self._finish_google_authorization(popup, status_var))
            except Exception as error:
                self.root.after(0, lambda error=error: self._show_google_authorization_failure(popup, status_var, error))

        threading.Thread(target=worker, daemon=True).start()

    def _finish_google_authorization(self, popup, status_var):
        self._stop_google_status_animation()
        self.google_session_connected = True
        self.google_connection_status_var.set("Google Sheets: Connected")
        self._update_google_controls_visibility()
        if hasattr(self, "google_authorize_button") and self.google_authorize_button.winfo_exists():
            self.google_authorize_button.configure(text="Reconnect / Change Google Account")
        if popup.winfo_exists():
            status_var.set("Google authorization completed successfully. You may close this window.")
        self.save_settings()

    def _show_google_authorization_failure(self, popup, status_var, error):
        if popup.winfo_exists():
            status_var.set(f"Authorization failed: {error}")

    def _reauthorize_google(self):
        if not self._run_when_dependency_ready(
            "google", self._reauthorize_google, "Preparing Google Sheets tools…"
        ):
            return
        if os.path.exists(self.token_file):
            if not messagebox.askyesno(
                "Reconnect Google Sheets",
                "Reconnect or change the Google account? Existing spreadsheet and class settings will be preserved.",
            ):
                return
            try:
                os.remove(self.token_file)
            except OSError as error:
                messagebox.showerror("Google Authorization Error", f"Could not replace the saved authorization:\n{error}")
                return
        self._authorize_google_interactive()

    def _start_google_startup_check(self):
        if not self.google_sheets_enabled_var.get():
            self._stop_google_status_animation()
            return
        self._start_google_status_animation("connecting")

        def worker():
            try:
                client = self.get_gsheet_client(interactive=False)
                title = ""
                if self.google_spreadsheet_id:
                    title = client.open_by_key(self.google_spreadsheet_id).title
                self.root.after(0, lambda: self._google_startup_success(title))
            except PermissionError:
                self.root.after(0, self._show_google_reconnect_prompt)
            except Exception as error:
                self.root.after(0, lambda error=error: self._show_google_startup_failure(error))
        threading.Thread(target=worker, daemon=True).start()

    def _google_startup_success(self, title):
        self._stop_google_status_animation()
        self.google_session_connected = True
        self.google_connection_status_var.set("Google Sheets: Connected")
        self._update_google_controls_visibility()
        if hasattr(self, "google_authorize_button") and self.google_authorize_button.winfo_exists():
            self.google_authorize_button.configure(text="Reconnect / Change Google Account")
        if title:
            self.google_sheet_title_var.set(title)
            self.save_settings()
            self._refresh_google_rosters(background=True, notify=False)

    def _show_google_startup_failure(self, error):
        self._stop_google_status_animation()
        self.google_session_connected = False
        self.google_connection_status_var.set(f"Google Sheets: Not Connected ({error})")
        self._update_google_controls_visibility()

    def _show_google_reconnect_prompt(self):
        self._stop_google_status_animation()
        self.google_session_connected = False
        self.google_connection_status_var.set("Google Sheets: Reconnection required")
        self._update_google_controls_visibility()
        if messagebox.askyesno(
            "Reconnect Google Sheets",
            "Google Sheets was previously connected, but its authorization has expired or was revoked. "
            "Your gradebook and class settings have been preserved. Reconnect now?",
        ):
            self._authorize_google_interactive()

    def _on_google_opt_in_changed(self):
        if self.google_sheets_enabled_var.get():
            self.google_sheets_enabled_var.set(False)
            self._show_google_opt_in_warning()
        else:
            self._stop_google_status_animation()
            self.google_session_connected = False
            self.google_connection_status_var.set("Google Sheets: Not Enabled")
            self.save_settings()
            self._update_google_controls_visibility()

    def _show_google_opt_in_warning(self):
        popup = tk.Toplevel(self.root)
        popup.title("About Google Sheets Integration")
        popup.transient(self.root)
        popup.grab_set()
        popup.geometry("650x760")
        ttk.Label(
            popup,
            text=("When you integrate with google sheets, you will manage all rosters through your own google sheet. "
                  "Please create a sheet, with one tab for each roster. Then, whenever you extract quiz scores, "
                  "it will update on your main roster page on the google sheet. You will still be able to download "
                  ".csv or .xlsx files of the current quiz grades."),
            wraplength=600,
            justify="left",
        ).pack(padx=20, pady=15)
        self._pack_google_help_image(popup, 500)
        buttons = ttk.Frame(popup)
        buttons.pack(pady=12)

        def cancel():
            self._stop_google_status_animation()
            self.google_sheets_enabled_var.set(False)
            self.google_session_connected = False
            self.google_connection_status_var.set("Google Sheets: Not Enabled")
            self.save_settings()
            self._update_google_controls_visibility()
            popup.destroy()

        def proceed():
            self._stop_google_status_animation()
            self.google_sheets_enabled_var.set(True)
            self.google_connection_status_var.set("Google Sheets: Not Connected")
            self.save_settings()
            self._update_google_controls_visibility()
            popup.destroy()

        ttk.Button(buttons, text="Cancel", command=cancel).pack(side="left", padx=8)
        ttk.Button(buttons, text="Proceed", command=proceed).pack(side="left", padx=8)
        popup.protocol("WM_DELETE_WINDOW", cancel)

    def _update_google_controls_visibility(self):
        controls_exist = (
            hasattr(self, "google_controls_frame")
            and self.google_controls_frame.winfo_exists()
        )
        if controls_exist:
            if self.google_sheets_enabled_var.get():
                self.google_controls_frame.pack(fill="x", padx=18)
            else:
                self.google_controls_frame.pack_forget()
        if hasattr(self, "google_create_button"):
            if self.google_create_button.winfo_exists():
                self.google_create_button.configure(
                    text=(
                        "Create New Google Sheets Gradebook"
                        if self.google_spreadsheet_id
                        else "Create Google Sheets Gradebook"
                    )
                )
        if hasattr(self, "google_open_button"):
            if self.google_open_button.winfo_exists():
                self.google_open_button.state(["!disabled"] if self.google_spreadsheet_id else ["disabled"])

        refresh_enabled = (
            self.google_sheets_enabled_var.get()
            and self.google_session_connected
            and bool(self.google_spreadsheet_id)
        )
        for attribute in ("google_refresh_button", "google_setup_refresh_button"):
            button = getattr(self, attribute, None)
            if button is not None and button.winfo_exists():
                button.state(["!disabled"] if refresh_enabled else ["disabled"])

    def _confirm_create_google_gradebook(self):
        if not self._run_when_dependency_ready(
            "google", self._confirm_create_google_gradebook, "Preparing Google Sheets tools…"
        ):
            return
        if self.google_spreadsheet_id:
            confirmed = messagebox.askyesno(
                "Create a New Google Sheets Gradebook?",
                f"Quiz Processing System is currently connected to:\n\n{self.google_sheet_title_var.get()}\n\n"
                "Creating a new gradebook will stop synchronization with the current spreadsheet. "
                "The existing Google spreadsheet will not be deleted or changed. Google-backed courses "
                "will need to be mapped to tabs in the new gradebook. Local classes and local gradebooks are unaffected.",
            )
            if not confirmed:
                return
        try:
            client = self.get_gsheet_client(interactive=False)
            known_titles = {item.get("name", "") for item in client.list_spreadsheet_files()}
            title = unique_gradebook_title(known_titles)
            spreadsheet = client.create(title)
            worksheet = spreadsheet.sheet1
            worksheet.update_title("Roster 1")
            worksheet.update(values=[["Name"]], range_name="A1", value_input_option="RAW")
            sample_error = None
            try:
                self._add_google_sample_worksheet(spreadsheet)
            except Exception as error:
                sample_error = error
            previous_id = self.google_spreadsheet_id
            self.google_spreadsheet_id = spreadsheet.id
            self.google_sheet_title_var.set(title)
            self.google_connection_status_var.set("Google Sheets: Connected")
            for class_info in self.classes.values():
                if isinstance(class_info, dict) and class_info.get("source") == "google_sheet":
                    if class_info.get("spreadsheet_id") == previous_id:
                        class_info["needs_remapping"] = True
            self._save_classes(redraw=False)
            self.save_settings()
            self._update_google_controls_visibility()
            self._show_google_sheets_help(created=True)
            if sample_error is not None:
                messagebox.showwarning(
                    "Sample Worksheet Not Added",
                    "The gradebook was created, but the SAMPLE worksheet could not be added:\n"
                    f"{sample_error}",
                )
        except PermissionError:
            messagebox.showinfo("Authorize Google Sheets", "Authorize Google Sheets before creating a gradebook.")
        except Exception as error:
            messagebox.showerror("Google Sheets Error", f"Could not create the gradebook:\n{error}")

    @staticmethod
    def _excel_rgb(color):
        """Convert an openpyxl ARGB color to a Google Sheets RGB color."""
        if color is None or color.type != "rgb" or not isinstance(color.rgb, str):
            return None
        rgb = color.rgb[-6:]
        try:
            red, green, blue = (int(rgb[index:index + 2], 16) / 255 for index in (0, 2, 4))
        except ValueError:
            return None
        return {"red": red, "green": green, "blue": blue}

    def _add_google_sample_worksheet(self, spreadsheet):
        """Copy the bundled Excel sample and its useful formatting into Sheets."""
        workbook_path = os.path.join(self.project_root, SAMPLE_WORKBOOK)
        openpyxl = importlib.import_module("openpyxl")
        workbook = openpyxl.load_workbook(workbook_path, data_only=False, read_only=False)
        source = workbook.worksheets[0]
        row_count = max(source.max_row, 1)
        column_count = max(source.max_column, 1)
        values = [
            [cell.value if cell.value is not None else "" for cell in row]
            for row in source.iter_rows(
                min_row=1, max_row=row_count, min_col=1, max_col=column_count
            )
        ]

        sample = spreadsheet.add_worksheet(
            title="SAMPLE", rows=row_count, cols=column_count, index=1
        )
        sample.update(values=values, range_name="A1", value_input_option="RAW")

        requests = []
        for row in source.iter_rows(
            min_row=1, max_row=row_count, min_col=1, max_col=column_count
        ):
            for cell in row:
                if not cell.has_style:
                    continue
                text_format = {
                    "fontFamily": cell.font.name,
                    "fontSize": cell.font.sz,
                    "bold": cell.font.bold,
                    "italic": cell.font.italic,
                }
                font_color = self._excel_rgb(cell.font.color)
                if font_color:
                    text_format["foregroundColor"] = font_color
                cell_format = {"textFormat": text_format}
                fill_color = self._excel_rgb(cell.fill.fgColor)
                if cell.fill.fill_type == "solid" and fill_color:
                    cell_format["backgroundColor"] = fill_color
                if cell.alignment.wrap_text:
                    cell_format["wrapStrategy"] = "WRAP"
                if cell.alignment.horizontal:
                    cell_format["horizontalAlignment"] = cell.alignment.horizontal.upper()
                if cell.alignment.vertical:
                    cell_format["verticalAlignment"] = cell.alignment.vertical.upper()
                requests.append({
                    "repeatCell": {
                        "range": {
                            "sheetId": sample.id,
                            "startRowIndex": cell.row - 1,
                            "endRowIndex": cell.row,
                            "startColumnIndex": cell.column - 1,
                            "endColumnIndex": cell.column,
                        },
                        "cell": {"userEnteredFormat": cell_format},
                        "fields": "userEnteredFormat",
                    }
                })

        for column_number in range(1, column_count + 1):
            letter = openpyxl.utils.get_column_letter(column_number)
            width = source.column_dimensions[letter].width
            if width is None:
                continue
            requests.append({
                "updateDimensionProperties": {
                    "range": {
                        "sheetId": sample.id,
                        "dimension": "COLUMNS",
                        "startIndex": column_number - 1,
                        "endIndex": column_number,
                    },
                    "properties": {"pixelSize": max(21, round(width * 7 + 5))},
                    "fields": "pixelSize",
                }
            })
        if requests:
            spreadsheet.batch_update({"requests": requests})

    def _open_google_gradebook(self):
        if not self.google_spreadsheet_id:
            messagebox.showinfo("No Google Gradebook", "Create a Google Sheets gradebook first.")
            return
        webbrowser.open(f"https://docs.google.com/spreadsheets/d/{quote(self.google_spreadsheet_id)}/edit")

    def _pack_google_help_image(self, parent, maximum):
        path = os.path.join(self.project_root, "reference", "google_sheet_roster_example.png")
        if self.dependency_states.get("pdf") != "ready":
            ttk.Label(parent, text="[Roster example image is still being prepared]").pack(pady=5)
            return
        try:
            image = Image.open(path)
            image.thumbnail((maximum, maximum))
            photo = ImageTk.PhotoImage(image)
            label = ttk.Label(parent, image=photo)
            label.image = photo
            label.pack(pady=5)
        except Exception:
            ttk.Label(parent, text="[Google Sheets roster example image]").pack(pady=5)

    def _show_google_sheets_help(self, created=False):
        popup = tk.Toplevel(self.root)
        popup.title("About Google Sheets")
        popup.transient(self.root)
        popup.geometry("700x820")
        ttk.Label(popup, text="Google Sheets Gradebook Setup", font=("Segoe UI", 16, "bold")).pack(pady=10)
        if created:
            ttk.Label(popup, text=f"Created: {self.google_sheet_title_var.get()}", style="Bold.TLabel").pack()
        instructions = (
            "Quiz Processing System created a private gradebook in your Google Drive. The SAMPLE tab contains an "
            "example roster and setup directions; you can map it as a practice class. Create one tab for each roster. "
            "Use 'Name' in cell A1 and enter one student name per row in column A. Rename the spreadsheet or tabs "
            "whenever you like; the program tracks their stable IDs. Do not merge cells in the roster or grade area. "
            "Existing topic headers are reused and new quiz topics are appended after the last used header. After "
            "editing rosters, click 'Refresh Rosters Now' or restart the program."
        )
        ttk.Label(popup, text=instructions, wraplength=650, justify="left").pack(padx=20, pady=10)
        self._pack_google_help_image(popup, 500)
        buttons = ttk.Frame(popup)
        buttons.pack(pady=10)
        ttk.Button(buttons, text="Open Google Sheets Gradebook", command=self._open_google_gradebook).pack(side="left", padx=6)
        ttk.Button(buttons, text="Close", command=popup.destroy).pack(side="left", padx=6)

    def _worksheet_by_id(self, spreadsheet, worksheet_id):
        return next((worksheet for worksheet in spreadsheet.worksheets() if worksheet.id == int(worksheet_id)), None)

    def _refresh_one_google_roster(self, class_name, class_info, client=None):
        """Atomically replace a Google class's local CSV cache."""
        try:
            client = client or self.get_gsheet_client(interactive=False)
            spreadsheet = client.open_by_key(class_info["spreadsheet_id"])
            worksheet = self._worksheet_by_id(spreadsheet, class_info["worksheet_id"])
            if worksheet is None:
                raise ValueError("The mapped worksheet tab no longer exists.")
            column = worksheet.col_values(1)
            headers = worksheet.row_values(1)
            if not column or column[0].strip().casefold() != "name":
                raise ValueError(f"Tab '{worksheet.title}' must have 'Name' in cell A1.")
            names = [name.strip() for name in column[1:] if name.strip()]
            duplicates = sorted({name for name in names if names.count(name) > 1})
            if duplicates:
                raise ValueError(f"Duplicate student name(s) in column A: {', '.join(duplicates)}")
            roster_path = self._resolve_class_path(class_name)
            write_roster_names(roster_path, names)
            class_info["worksheet_title"] = worksheet.title
            class_info["last_seen_headers"] = headers
            class_info["last_refreshed_at"] = datetime.now(timezone.utc).isoformat()
            class_info["needs_remapping"] = False
            return True
        except Exception as error:
            print(f"[WARN] Could not refresh Google roster '{class_name}': {error}")
            return False

    def _refresh_google_rosters(self, background=True, notify=False):
        if not self._run_when_dependency_ready(
            "google",
            lambda: self._refresh_google_rosters(background=background, notify=notify),
            "Preparing Google Sheets tools…",
        ):
            return
        records = [
            (name, info) for name, info in self.classes.items()
            if isinstance(info, dict)
            and info.get("source") == "google_sheet"
            and not info.get("needs_remapping")
        ]

        def refresh():
            succeeded = 0
            try:
                client = self.get_gsheet_client(interactive=False)
                for class_name, class_info in records:
                    succeeded += int(self._refresh_one_google_roster(class_name, class_info, client))
                self._save_classes(redraw=False)
            except Exception as error:
                print(f"[WARN] Google roster refresh failed: {error}")
            self.root.after(0, lambda: self._finish_roster_refresh(succeeded, len(records), notify))

        if background:
            threading.Thread(target=refresh, daemon=True).start()
        else:
            refresh()

    def _finish_roster_refresh(self, succeeded, total, notify):
        if total == 0:
            self.google_session_rosters_refreshed = not any(
                isinstance(info, dict) and info.get("source") == "google_sheet"
                for info in self.classes.values()
            )
            self.google_roster_status_var.set("No Google Sheets rosters are currently mapped.")
        elif succeeded == total:
            self.google_session_rosters_refreshed = True
            self.google_rosters_last_updated = datetime.now().astimezone().isoformat()
            self.google_roster_status_var.set(format_local_timestamp(self.google_rosters_last_updated))
            self.save_settings()
        else:
            previous = format_local_timestamp(self.google_rosters_last_updated)
            self.google_roster_status_var.set(f"Roster refresh incomplete: {succeeded} of {total} updated. {previous}")
        if notify:
            if succeeded == total:
                messagebox.showinfo("Google Rosters", f"Refreshed {succeeded} of {total} rosters.")
            else:
                messagebox.showwarning(
                    "Google Rosters",
                    f"Refreshed {succeeded} of {total} rosters. Existing local caches were preserved for failures.",
                )
    
    def update_gsheet_from_extracted_data(self):
        """Update existing topic columns and append only genuinely new topics."""
        class_name = self.class_combo.get()
        class_info = self.classes.get(class_name, {})
        if not isinstance(class_info, dict) or class_info.get("source") != "google_sheet":
            messagebox.showerror("Google Sheets Error", "The selected class is not mapped to Google Sheets.")
            return False
        try:
            client = self.get_gsheet_client(interactive=False)
            spreadsheet = client.open_by_key(class_info["spreadsheet_id"])
            worksheet = self._worksheet_by_id(spreadsheet, class_info["worksheet_id"])
            if worksheet is None:
                raise ValueError("The mapped worksheet tab no longer exists.")

            # Re-read live headers immediately before upload so concurrent edits are respected.
            headers = worksheet.row_values(1)
            if not headers or headers[0].strip().casefold() != "name":
                raise ValueError("The mapped tab must have 'Name' in cell A1.")
            names = [name.strip() for name in worksheet.col_values(1)[1:] if name.strip()]
            duplicates = sorted({name for name in names if names.count(name) > 1})
            if duplicates:
                raise ValueError(f"Duplicate student name(s) in column A: {', '.join(duplicates)}")
            if self.google_roster_snapshot is None or names != self.google_roster_snapshot:
                messagebox.showerror(
                    "Google Roster Changed",
                    "The Google Sheets roster changed after extraction began. No scores were written. "
                    "Refresh the roster and review the extracted grades before trying again.",
                )
                return False
            header_lookup = {value.strip().casefold(): index + 1 for index, value in enumerate(headers) if value.strip()}
            topic_names = [value.get().strip() for value in self.topic_vars if value.get().strip()]
            for topic in topic_names:
                key = topic.casefold()
                if key not in header_lookup:
                    headers.append(topic)
                    header_lookup[key] = len(headers)
                    header_cell = gspread.utils.rowcol_to_a1(1, len(headers))
                    worksheet.update(
                        values=[[topic]],
                        range_name=header_cell,
                        value_input_option="USER_ENTERED",
                    )

            row_lookup = {name: row for row, name in enumerate(names, start=2)}
            preview_rows = {
                str(self.tree.item(row_id)["values"][0]).strip(): self.tree.item(row_id)["values"]
                for row_id in self.tree.get_children()
            }
            updates = []
            for name, values in preview_rows.items():
                row_number = row_lookup.get(name)
                if not row_number:
                    continue  # The Google roster remains authoritative.
                for topic_index, topic in enumerate(topic_names, start=1):
                    value = values[topic_index] if topic_index < len(values) else ""
                    if isinstance(value, str) and value.strip().casefold() == "skip":
                        value = ""
                    else:
                        value = normalize_score_value(value)
                    updates.append({
                        "range": gspread.utils.rowcol_to_a1(row_number, header_lookup[topic.casefold()]),
                        "values": [[value]],
                    })
            if updates:
                worksheet.batch_update(updates, raw=True)
            class_info["last_seen_headers"] = headers
            self._save_classes(redraw=False)
            self._show_gsheets_popup(
                str1=f"Updated {len(topic_names)} topic column(s).",
                str2=f"Wrote {len(updates)} score cell(s).",
            )
            return True
        except Exception as error:
            messagebox.showerror("Google Sheets Error", f"Could not synchronize scores:\n{error}")
            return False


    #End personal update section

    #Google Sheet Updated Pop-up
    def _show_gsheets_popup(self, str1, str2):
        """Pop-up after gradebook update with option to open it."""
        popup = tk.Toplevel(self.root)
        popup.title("Google Sheets Updated")
        popup.geometry("430x190")
        popup.grab_set()  # make it modal

        ttk.Label(
            popup,
            text="Google Sheets was updated successfully.\n"
                + str1 + "\n"
                + str2 + "\n\n"
                + "It is now safe to edit your Google Sheet on the cloud as needed.",
            wraplength=400,
            justify="center"
        ).pack(pady=(20, 10))


        ttk.Button(popup, text="Close", command=popup.destroy).pack()






if __name__ == "__main__":
    root = tk.Tk()
    missing = missing_packaged_resources() if getattr(sys, "frozen", False) else []
    if missing:
        messagebox.showerror(
            "Quiz Processing System could not start",
            "This installation is incomplete. Reinstall the application.\n\nMissing:\n- "
            + "\n- ".join(missing),
            parent=root,
        )
        root.destroy()
    else:
        app = QuizAppGUI(root)
        root.mainloop()
