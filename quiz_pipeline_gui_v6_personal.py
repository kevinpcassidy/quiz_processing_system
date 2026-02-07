import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from tkinter import DoubleVar, BooleanVar
import json
import os
import shutil, tempfile
from tkinter import font as tkfont
from pdf2image import convert_from_path, pdfinfo_from_path
from PIL import Image, ImageEnhance, ImageTk
import pytesseract
from rapidfuzz import process
import pandas as pd
import re
import csv
import cv2
import numpy as np
import gspread
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
import io
import time
import webbrowser
import subprocess, platform
import threading


#Global Variables
STOP_PROCESSING = False
SETTINGS_FILE = os.path.join(os.getcwd(), "quiz_settings.json")

"""Next four lines are for personal version to update to google sheets"""
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
CREDENTIALS_FILE = 'credentials.json'
TOKEN_FILE = 'token.json'
SHEET_ID = '1ASYMjWn1JYevcal9jufg29DR5Q5P081rJTGNzUAZQJs'
"""End update to google sheets"""

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
        self.left_frame = ttk.Frame(self.paned, padding=(20, 12, 20, 12))
        self.center_frame = ttk.Frame(self.paned, padding=12)
        self.right_frame = ttk.Frame(self.paned, padding=12)
        self.paned.add(self.left_frame, weight=2)
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
        

    
        # Dictionary to hold classes and their CSV paths
        self.classes = {}
        self.rosters_dir = os.path.join(os.getcwd(), "rosters")
        os.makedirs(self.rosters_dir, exist_ok=True)
        self.classes_file = os.path.join(os.getcwd(), "saved_classes.json")
        self._load_classes()
        
        # Grading scales storage
        self.grading_file = os.path.join(os.getcwd(), "saved_grading_scales.json")
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
        self.gsheet_credentials_path = tk.StringVar(value="")
        
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
 
  
    # ---------------- UTILITY ----------------
    def _on_close(self):
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
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, "r") as f:
                    data = json.load(f)
                self.enable_side_detection.set(data.get("enable_side_detection", False))
                self.score_threshold.set(data.get("score_threshold", 3.2))
                self.enable_gradebook_var.set(data.get("enable_gradebook_var", False))
                self.gsheet_credentials_path.set(data.get("gsheet_credentials_path", ""))
            except Exception as e:
                print(f"[DEBUG] Failed to load settings: {e}")
        else:
            # Defaults if file doesn't exist
            self.enable_side_detection.set(False)
            self.score_threshold.set(3.2)
            self.enable_gradebook_var.set(False)
            self.gsheet_credentials_path.set("")

    def save_settings(self):
        """Save current settings to file."""
        try:
            data = {
                "enable_side_detection": self.enable_side_detection.get(),
                "score_threshold": self.score_threshold.get(),
                "enable_gradebook_var": self.enable_gradebook_var.get(),
                "gsheet_credentials_path": self.gsheet_credentials_path.get()
            }
            with open(SETTINGS_FILE, "w") as f:
                json.dump(data, f, indent=4)
        except Exception as e:
            print(f"[DEBUG] Failed to save settings: {e}")



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
            ("download", "Download CSV"),
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
        ttk.Button(self.left_frame, text="Set-up Classes", command=self._setup_classes_panel).pack(fill="x", pady=4)
        ttk.Button(self.left_frame, text="Set-up Grading Scale", command=self._open_grading_scale_setup).pack(fill="x", pady=4)
        if self.enable_gradebook_var:
            ttk.Button(self.left_frame, text="View Gradebook", command = self._on_view_gradebook).pack(fill="x", pady=4)
        ttk.Button(self.left_frame, text="Advanced", command=self.setup_advanced_pop_up).pack(fill="x", pady=4)

        ttk.Separator(self.left_frame, orient="horizontal").pack(fill="x", pady=(10,8))
        gsheet_frame = ttk.Frame(self.left_frame)
        gsheet_frame.pack(fill="x", pady=(0, 6))
        ttk.Label(
            gsheet_frame,
            text="Choose JSON file for Google Sheets",
        ).pack(side="left")
        ttk.Button(
            gsheet_frame,
            text="Browse",
            command=self._select_gsheet_credentials
        ).pack(side="right")




    # ---------------- SELECTION HANDLERS ----------------
    def _select_pdf(self):
        """Select PDF file and display its path."""
        file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if file_path:
            self.pdf_path_var.set(file_path)
            
            # --- Reset process tracking when a new PDF is selected ---
            self.completed_steps = {
                "pdf_selected": False,
                "roster_loaded": False,
                "names_verified": False,
                "scores_verified": False,
                "topics_saved": False,
                "calibrated": False,
                "csv_downloaded": False
            }

            # Reset dropdowns and selections to initial state
            if hasattr(self, "class_combo"):
                self.class_combo.current(0)
            if hasattr(self, "scale_combo"):
                self.scale_combo.current(0)
                        
            self._mark_topics_modified()            
            
            for step in self.progress_labels:
                self._set_check(self.progress_labels[step], is_done=False)

            # ✅ Now mark the PDF step complete again (after reset)
            self.mark_step_done("pdf")

    def _select_gsheet_credentials(self):
        file_path = filedialog.askopenfilename(
            title="Select Google Sheets Credentials",
            filetypes=[("JSON Files", "*.json"), ("All Files", "*.*")]
        )
        if file_path:
            self.gsheet_credentials_path.set(file_path)
            self.save_settings()

    def _get_gsheet_credentials_path(self):
        path = self.gsheet_credentials_path.get().strip()
        return path if path else CREDENTIALS_FILE


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

        # --- Enable Full Gradebook Update ---
        ttk.Label(content_frame, text="Track and Update a Full Gradebook", font=("Segoe UI", 14, "bold")).pack(pady=(10, 5))
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
            text="Enable Full Gradebook Update",
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
        # Section 2: Detecting Multiple Scores
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
        # Clear any existing widgets in the center frame
        for widget in self.center_frame.winfo_children():
            widget.destroy()
        
        # Heading at top
        ttk.Label(
            self.center_frame,
            text="Quiz Processing System",
            font=("TkDefaultFont", 20, "bold"),
            anchor="center"
        ).pack(pady=(30, 10))
        
        # Description below heading
        ttk.Label(
            self.center_frame,
            text="Welcome to the Quiz Processing System. \n"
                "Please begin by following the steps on the left frame.\n\n",
            wraplength=self.center_frame.winfo_width() - 20,
            justify="center"
        ).pack(pady=(0, 10))
        
        # Horizontal line
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=(0, 20))
        
        # YouTub heading
        ttk.Label(
            self.center_frame,
            text="YouTube Tutorials:",
            font=("TkDefaultFont", 12, "bold"),
            anchor="center"
        ).pack(pady=(0, 10))
        
        # YouTube clickable link
        link_label = tk.Label(
            self.center_frame,
            text="Come visit me at:\nhttps://www.youtube.com/@KevinsTeacherTech",
            fg="blue",
            cursor="hand2",
            justify="center"
        )
        link_label.pack(pady=(5, 20))
        
        def open_youtube_link(event):
            webbrowser.open("https://www.youtube.com/@KevinsTeacherTech")
        
        link_label.bind("<Button-1>", open_youtube_link)
        
        # Horizontal line
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=(0, 20))
        
        # Microsoft Word Templates heading
        ttk.Label(
            self.center_frame,
            text="Microsoft Word Templates:",
            font=("TkDefaultFont", 12, "bold"),
            anchor="center"
        ).pack(pady=(0, 10))
        
        # Template links
        templates = [
            ("One-Page Quiz Template", "https://docs.google.com/document/d/1EF0sel2g1I94xmV5VCxS2j-vzveeEm6P/export?format=docx"),
            ("Two-Page Quiz Template", "https://docs.google.com/document/d/1xk3f2LEKAum9tqkyix8UkZryegbYsVlA/export?format=docx")
        ]
        
        for name, url in templates:
            link = tk.Label(
                self.center_frame,
                text=name,
                fg="blue",
                cursor="hand2",
                justify="center"
            )
            link.pack(pady=2)
            link.bind("<Button-1>", lambda e, url=url: webbrowser.open(url))



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

        #Help Button
        ttk.Button(container, text="How to Setup Classes", command=self._show_setup_classes_help).pack(fill="x", pady=15)
        
        # Add/Delete buttons
        ttk.Button(container, text="Add Class", command=self._add_class).pack(fill="x", pady=5)
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

            # If Excel → convert to CSV
            if file_path.lower().endswith((".xlsx", ".xls")):
                try:
                    import pandas as pd
                    df = pd.read_excel(file_path)
                    df.to_csv(dest_path, index=False, encoding="utf-8-sig")
                except Exception as e:
                    messagebox.showerror("Conversion Error",
                        f"Could not convert Excel file:\n{e}")
                    return
            else:
                shutil.copy(file_path, dest_path)

            # Update classes dictionary and UI
            self.classes[class_name] = dest_path
            self._save_classes()
            self._refresh_classes_tree()
            self._refresh_class_combobox()

            popup.destroy()  # Close pop-up

        ttk.Button(popup, text="Save and Close", command=save_and_close).pack(pady=10)

        # Handle user closing the window without saving
        popup.protocol("WM_DELETE_WINDOW", popup.destroy)

    def _load_classes(self):
        """Load classes info from JSON if available"""
        if os.path.exists(self.classes_file):
            try:
                with open(self.classes_file, "r", encoding="utf-8") as f:
                    self.classes = json.load(f)
            except Exception as e:
                print("Error loading classes:", e)
                
    def _save_classes(self):
        """Save current classes dictionary to JSON"""
        try:
            with open(self.classes_file, "w", encoding="utf-8") as f:
                json.dump(self.classes, f, indent=2)
        except Exception as e:
            print("Error saving classes:", e)
        # After deleting and saving
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
        csv_path = self.classes.get(class_name)
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

        class_name, roster_file = self.class_tree.item(selected[0], "values")

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
        msg = f'Please confirm that you want to delete the class called "{class_name}" with roster "{roster_file}".'
        ttk.Label(popup, text=msg, wraplength=380, justify="center").pack(pady=(20, 10), padx=10)

        # Button frame
        btn_frame = ttk.Frame(popup)
        btn_frame.pack(pady=10)

        def confirm_delete():
            # Remove CSV file
            csv_path = self.classes.get(class_name)
            if csv_path and os.path.exists(csv_path):
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
        win_width = 450
        win_height = min(img.height + 400, 700)
        help_window.geometry(
            f"{win_width}x{win_height}+"
            f"{int((help_window.winfo_screenwidth()-win_width)/2)}+"
            f"{int(help_window.winfo_screenheight()/6)}"
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
        text = ("Use a gradebook program to export a CSV roster. \n\n"
                "We recommend creating a roster of all the classes of one subject, "
                "so if you teach two different subjects or grade level classes, each would be its own class.\n"
                "Here is a picture of a properly formatted CSV file.\n"
                "Lines 1-5 are students in period 1 of 8th grade math... lines 6-10 are period 2.\n"
                "Students are alphabetized to easily allow for copying/pasting into another file for your gradebook.\n\n")
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
            for v in score_vars:
                val = v.get().strip()
                if val:
                    try:
                        scores.append(float(val))
                    except ValueError:
                        messagebox.showwarning("Invalid Score", f"'{val}' is not a number.")
                        return
            scores = sorted(list(set(scores)))  # remove duplicates & sort

            # Convert floats that are whole numbers to int for display
            display_scores = [int(s) if s.is_integer() else s for s in scores]

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
                # Convert all scores to float (JSON may store as int)
                for key, scores in self.grading_scales.items():
                    self.grading_scales[key] = [float(s) for s in scores]
            except Exception as e:
                print("Error loading grading scales:", e)
                self.grading_scales = {}
        else:
            self.grading_scales = {}


    def _edit_grading_scale(self):
        # Ensure a scale is selected in the right panel treeview
        if not hasattr(self, "grading_tree") or not self.grading_tree.get_children():
            messagebox.showinfo("No Grading Scales", "There are no grading scales to edit.")
            return

        selected = self.grading_tree.selection()
        if not selected:
            messagebox.showinfo("No Selection", "Please select a grading scale from the right pane to edit.")
            return

        scale_name, scores_str = self.grading_tree.item(selected[0], "values")
        scores_list = [float(s) for s in scores_str.replace(" ", "").split(",")]

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

    def _on_run_calibration(self):
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
            pages = convert_from_path(pdf_path, dpi=200)
            if self.stop_processing:
                return

            # Pass result back to main thread safely
            self.root.after(0, lambda: self._on_pages_loaded(pages, start))

        except Exception as e:
            self.root.after(
                0,
                lambda: messagebox.showerror("PDF Error", str(e))
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
            with open(roster_file, newline="", encoding="utf-8-sig") as csvfile:
                reader = csv.reader(csvfile)
                self.roster_names = [row[0].strip() for row in reader if row]
            print(f"Loaded {len(self.roster_names)} students from {roster_file}")
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

        # You can add a short delay or directly call your calibration restart
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
            val = score_var.get().strip()
            val = re.sub(r"[^\d\.]", "", val)
            if not val:
                messagebox.showwarning("Invalid Entry", "Please enter or click a valid score.")
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
                        writer.writerow(row_values)
                messagebox.showinfo("Exported", f"CSV exported to:\n{file_path}")
                self.mark_step_done("download")

        ttk.Button(btn_frame, text="Download CSV File", command=export_csv).pack(side="left", padx=5)
        # --- Horizontal separator ---
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=10)
        
        def update_gradebook():
            self.update_full_gradebook()
            self.update_gradebook_btn.state(["disabled"])
            self._show_gradebook_popup(additional_text = "Your gradebook has been updated")

        if self.enable_gradebook_var.get():
            tk.Label(self.center_frame,
                    text = "If you want to save all data that has been extracted to your main gradebook records, you must click the following button to update.\n",
                    wraplength = self.center_frame.winfo_width()-10,
                    justify="left").pack(pady=(8, 2))
        
            self.update_gradebook_btn = ttk.Button(
                btn_frame,
                text="Update My Gradebook",
                command=lambda e: update_gradebook()
            )
            self.update_gradebook_btn.pack(side="left", padx=5)

        
        # --- Horizontal separator ---
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=10)

        # --- Coffee button ---
        tk.Label(self.center_frame, text="Like this program?").pack(pady=(5,2))
        def open_venmo():
            webbrowser.open("https://www.venmo.com/u/KevinPCassidy1981")
        ttk.Button(self.center_frame, text="Buy Kevin a Coffee", command=open_venmo).pack(pady=(0,10))

        """Next 8 lines or so are just for personal use gsheet update"""
        def update_gsheets():
            self.update_gsheet_from_extracted_data(sheet_id=SHEET_ID)
            self.update_gradebook_btn.state(["disabled"])

        self.update_gradebook_btn = ttk.Button(
            self.center_frame,
            text="Update to Google Sheets",
            command=update_gsheets        # runs ONLY when clicked
        )
        self.update_gradebook_btn.pack(pady=(20,10))


        """End update for personal gsheet update"""

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
        popup.title("Open Gradebook")
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
                "the Gradebook button under Preferences.\n\n"
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
                        "Gradebook Update Error",
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

        ttk.Button(popup, text="Open Gradebook", command=open_gradebook).pack(pady=(0, 10))
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
        new_data = [self.tree.item(row_id)["values"] for row_id in self.tree.get_children()]
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
                "Gradebook Update Error",
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

        roster_file_name = f"{roster_name.replace('_', ' ')}.csv"
        roster_path = os.path.join("rosters", roster_file_name)

        if not os.path.exists(roster_path):
            print(f"[WARN] Roster file not found: {roster_path}")
            return df_gradebook

        df_roster = pd.read_csv(roster_path, header=None)
        if str(df_roster.iloc[0, 0]).strip().lower() == "name":
            df_roster = pd.read_csv(roster_path)  # read normally with header
        else:
            df_roster.rename(columns={df_roster.columns[0]: "Name"}, inplace=True)

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
        ttk.Label(self.center_frame, text="View Gradebook", style="Header.TLabel").pack(pady=(10, 6))

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

        ttk.Button(btn_frame, text="View Gradebook", command=view_gradebook).pack(side="left", padx=5)

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

        ttk.Button(btn_frame, text="Delete Gradebook", command=delete_gradebook).pack(side="left", padx=5)
        
        # Horizontal line
        ttk.Separator(self.center_frame, orient="horizontal").pack(fill="x", pady=(15, 5))
        
        # Button to return to original center panel
        ttk.Button(self.center_frame, text="Home", command=lambda: self.reset_panels()).pack(pady=8)


        

    def reset_panels(self):
        # Destroy all children recursively
        for child in self.right_frame.winfo_children():
            child.destroy()

        # Rebuild right panel
        self._build_right_panel()

        # Destroy and rebuild center frame as usual
        for child in self.center_frame.winfo_children():
            child.destroy()
        self._build_center_panel()


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
        columns = ("Class Name", "Roster CSV")
        self.class_tree = ttk.Treeview(frame, columns=columns, show="headings", selectmode="browse")
        self.class_tree.heading("Class Name", text="Class Name")
        self.class_tree.heading("Roster CSV", text="Roster CSV")
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
        csv_path = self.classes.get(class_name)
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

        # Populate treeview from CSV
        with open(csv_path, "r", encoding="utf-8-sig") as f:
            for line in f:
                name = line.strip()
                while name.startswith('"') or name.startswith("'"):
                    name = name[1:]
                while name.endswith('"') or name.endswith("'"):
                    name = name[:-1]
                name = name.strip()
                if name:
                    tree.insert("", "end", values=(name,))


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
        for class_name, roster_path in self.classes.items():
            self.class_tree.insert("", "end", values=(class_name, os.path.basename(roster_path)))

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
                score_str = ", ".join(str(int(s)) if s.is_integer() else str(s) for s in scores)
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
                score_str = ", ".join(str(int(s)) if float(s).is_integer() else str(s) for s in scores)
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
            pages = convert_from_path(pdf_path, dpi=200, first_page=1, last_page=1)
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

    def get_gsheet_client(self):
        credentials_path = self._get_gsheet_credentials_path()
        if not os.path.exists(credentials_path):
            raise FileNotFoundError(
                f"Google Sheets credentials file not found: {credentials_path}"
            )
        return gspread.service_account(filename=credentials_path)
    
    def normalize_numeric_cells(self, data):
        """
        Converts numeric strings like '3', '2.5' to real numbers (floats or ints)
        Leaves everything else as strings. Safe for None values.
        """
        cleaned = []
        for row in data:
            new_row = []
            for cell in row:
                cell_str = str(cell).strip() if cell is not None else ""
                if re.fullmatch(r"\d+(\.\d+)?", cell_str):
                    num = float(cell_str)
                    if num.is_integer():
                        num = int(num)
                    new_row.append(num)
                else:
                    new_row.append(cell_str)
            cleaned.append(new_row)
        return cleaned


    def sanitize_for_upload(self, data):
        """Ensures everything written back is either a number or a clean string.
           Forces headers to be pure text so Sheets doesn't add a leading apostrophe.
        """
        result = []
        for r, row in enumerate(data):
            cleaned_row = []

            # --- HEADER ROW FIX ---
            if r == 0:
                # Force all headers to be simple text, preventing Google auto-numbering
                cleaned_row = [str(cell).strip() for cell in row]
                result.append(cleaned_row)
                continue
            # ----------------------

            for cell in row:
                if isinstance(cell, (int, float)):
                    cleaned_row.append(cell)
                else:
                    cleaned_row.append("" if cell is None else str(cell).strip())

            result.append(cleaned_row)

        return result


    def update_gsheet_from_extracted_data(self, sheet_id):
        """
        Append extracted data to a Google Sheet tab for the selected class.
        Creates new columns if topics are new and batch uploads the full dataset.
        Existing headers remain unchanged; new topics are appended with correct types.
        """
        tab_name = self.class_combo.get()
        try:
            client = self.get_gsheet_client()
        except (FileNotFoundError, ValueError) as e:
            messagebox.showerror("Google Sheets Error", str(e))
            return
        sheet = client.open_by_key(sheet_id).worksheet(tab_name)

        print("Pulling current sheet data...")
        all_values = sheet.get_all_values()
        all_values = self.normalize_numeric_cells(all_values)
        headers = all_values[0] if all_values else ['Name']
        existing_rows = all_values[1:] if len(all_values) > 1 else []

        # Build dict from existing sheet
        names = [row[0] for row in existing_rows]
        sheet_data = {name: {headers[i]: row[i] if i < len(row) else '' for i in range(len(headers))}
                      for name, row in zip(names, existing_rows)}

        new_topics = []

        # Process extracted data
        for student in self.extracted_data:
            if student is None:
                continue

            name = student['name']

            # Ensure student exists in sheet_data
            if name not in sheet_data:
                sheet_data[name] = {'Name': name}

            for topic, score in student['topics'].items():
                # Clean topic name
                topic_clean = str(topic).strip()

                # Add new topic if not in headers
                if topic_clean not in headers:
                    headers.append(topic_clean)
                    new_topics.append(topic_clean)

                # Clean score
                if score is None or score == "":
                    cleaned_score = ""
                else:
                    cleaned_score = re.sub(r"[^\d\.]", "", str(score).strip())
                    try:
                        if cleaned_score != "":
                            cleaned_score = float(cleaned_score)
                    except ValueError:
                        cleaned_score = ""

                sheet_data[name][topic_clean] = cleaned_score

        # Rebuild full sheet data
        updated_rows = [[sheet_data[name].get(h, '') for h in headers] for name in sheet_data]

        # --- UPLOAD HEADERS AND DATA SEPARATELY USING RAW ---
        print(f"Adding {len(new_topics)} new topics.")
        print(f"Writing full dataset with {len(updated_rows)} students and {len(headers)} columns...")
        self._show_gsheets_popup(str1=f"Added {len(new_topics)} new topics.",
                                 str2=f"Wrote full dataset with {len(updated_rows)} students and {len(headers)} columns.")

        # Upload headers as RAW to preserve types and prevent apostrophes
        sheet.update(values=[headers], range_name='1:1', value_input_option='RAW')

        # Upload all rows below header as RAW
        sheet.update(values=updated_rows, range_name='A2', value_input_option='RAW')

        print("Batch upload complete.")


    #End personal update section

    #Google Sheet Updated Pop-up
    def _show_gsheets_popup(self, str1, str2):
        """Pop-up after gradebook update with option to open it."""
        popup = tk.Toplevel(self.root)
        popup.title("Google Sheets Updated")
        popup.geometry("350x100")
        popup.grab_set()  # make it modal

        ttk.Label(
            popup,
            text="Google Sheets has been updated.\n"
                + str1 + "\n"
                + str2 + "\n",
            wraplength=320,
            justify="center"
        ).pack(pady=(20, 10))


        ttk.Button(popup, text="Close", command=popup.destroy).pack()






if __name__ == "__main__":
    root = tk.Tk()
    app = QuizAppGUI(root)
    root.mainloop()
