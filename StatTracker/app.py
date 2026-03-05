from __future__ import annotations

import os
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from PIL import Image, ImageGrab, ImageTk

from excel_writer import ExcelFileLockedError, append_stats
from stat_extractor import CANONICAL_FIELDS, extract_nickname, extract_stats_and_nickname


COLORS = {
    "bg": "#f3f5f8",
    "surface": "#ffffff",
    "surface_alt": "#eef2f7",
    "text": "#1f2933",
    "muted": "#5b6675",
    "border": "#d5dde8",
    "primary": "#0b66d0",
    "primary_active": "#0954aa",
    "success": "#157347",
    "success_active": "#12623c",
    "danger": "#b42318",
    "danger_active": "#8f1a12",
    "status_bg": "#e8edf3",
}


class StatTrackerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Stat Tracker OCR")
        self.geometry("1240x780")
        self.minsize(1050, 700)
        self.configure(bg=COLORS["bg"])

        self.selected_image: str | None = None
        self.excel_path: str | None = None
        self.preview_photo = None
        self.hotkey_listener = None
        self.hotkey_keyboard = None
        self.session_active = False
        self.capture_in_progress = False
        self.saved_rows = 0
        self.current_nickname: str | None = None
        self.window_title_var = tk.StringVar()
        self.value_vars = {field: tk.StringVar() for field in CANONICAL_FIELDS}

        self._configure_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure(
            ".",
            background=COLORS["bg"],
            foreground=COLORS["text"],
            font=("Segoe UI", 10),
        )

        style.configure("App.TFrame", background=COLORS["bg"])
        style.configure(
            "Card.TLabelframe",
            background=COLORS["surface"],
            borderwidth=1,
            relief="solid",
            bordercolor=COLORS["border"],
        )
        style.configure(
            "Card.TLabelframe.Label",
            background=COLORS["surface"],
            foreground=COLORS["text"],
            font=("Segoe UI Semibold", 10),
        )
        style.configure("Surface.TFrame", background=COLORS["surface"])
        style.configure("Panel.TFrame", background=COLORS["surface_alt"])
        style.configure("Card.TLabel", background=COLORS["surface"], foreground=COLORS["text"])
        style.configure("Muted.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])

        style.configure(
            "TEntry",
            fieldbackground="#ffffff",
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            lightcolor=COLORS["border"],
            darkcolor=COLORS["border"],
            insertcolor=COLORS["text"],
            padding=6,
        )
        style.map(
            "TEntry",
            bordercolor=[("focus", COLORS["primary"])],
            lightcolor=[("focus", COLORS["primary"])],
            darkcolor=[("focus", COLORS["primary"])],
        )

        style.configure(
            "Neutral.TButton",
            background=COLORS["surface_alt"],
            foreground=COLORS["text"],
            bordercolor=COLORS["border"],
            padding=(10, 6),
        )
        style.map(
            "Neutral.TButton",
            background=[("active", "#e2e8f0")],
        )

        style.configure(
            "Primary.TButton",
            background=COLORS["primary"],
            foreground="#ffffff",
            bordercolor=COLORS["primary"],
            padding=(12, 6),
        )
        style.map(
            "Primary.TButton",
            background=[("active", COLORS["primary_active"])],
            foreground=[("active", "#ffffff")],
        )

        style.configure(
            "Success.TButton",
            background=COLORS["success"],
            foreground="#ffffff",
            bordercolor=COLORS["success"],
            padding=(12, 6),
        )
        style.map(
            "Success.TButton",
            background=[("active", COLORS["success_active"])],
            foreground=[("active", "#ffffff")],
        )

        style.configure(
            "Danger.TButton",
            background=COLORS["danger"],
            foreground="#ffffff",
            bordercolor=COLORS["danger"],
            padding=(12, 6),
        )
        style.map(
            "Danger.TButton",
            background=[("active", COLORS["danger_active"])],
            foreground=[("active", "#ffffff")],
        )

        style.configure(
            "Status.TLabel",
            background=COLORS["status_bg"],
            foreground=COLORS["text"],
            font=("Segoe UI", 9),
            padding=(10, 6),
        )

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=14, style="App.TFrame")
        root.pack(fill="both", expand=True)
        root.columnconfigure(0, weight=1)
        root.columnconfigure(1, weight=1)
        root.rowconfigure(0, weight=1)

        left = ttk.LabelFrame(root, text="Capture", padding=12, style="Card.TLabelframe")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(2, weight=1)

        right = ttk.LabelFrame(root, text="Extracted Stats", padding=12, style="Card.TLabelframe")
        right.grid(row=0, column=1, sticky="nsew", padx=(8, 0))
        right.columnconfigure(0, weight=1)
        right.rowconfigure(0, weight=1)

        controls = ttk.Frame(left, style="Surface.TFrame")
        controls.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        ttk.Button(
            controls,
            text="Choose Screenshot",
            command=self.choose_image,
            style="Neutral.TButton",
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            controls, text="Extract Stats", command=self.extract, style="Neutral.TButton"
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            controls, text="Clear", command=self.clear_fields, style="Neutral.TButton"
        ).pack(side="left")
        ttk.Button(
            controls,
            text="Auto Capture + Save",
            command=self.auto_capture_and_save,
            style="Primary.TButton",
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            controls, text="Start Session", command=self.start_hotkey_session, style="Success.TButton"
        ).pack(side="left", padx=(8, 0))
        ttk.Button(
            controls, text="Stop Session", command=self.stop_hotkey_session, style="Danger.TButton"
        ).pack(side="left", padx=(8, 0))

        self.image_info_var = tk.StringVar(value="No screenshot selected.")
        ttk.Label(left, textvariable=self.image_info_var, style="Muted.TLabel").grid(
            row=1, column=0, sticky="ew", pady=(0, 8)
        )

        preview_frame = ttk.Frame(left, style="Panel.TFrame", padding=8)
        preview_frame.grid(row=2, column=0, sticky="nsew")
        preview_frame.columnconfigure(0, weight=1)
        preview_frame.rowconfigure(0, weight=1)
        self.preview_label = ttk.Label(preview_frame, style="Card.TLabel")
        self.preview_label.grid(row=0, column=0, sticky="nsew")

        excel_controls = ttk.Frame(left, style="Surface.TFrame")
        excel_controls.grid(row=3, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(
            excel_controls, text="Choose Excel File", command=self.choose_excel, style="Neutral.TButton"
        ).pack(side="left", padx=(0, 8))
        ttk.Button(
            excel_controls, text="Save Row to Excel", command=self.save_row, style="Primary.TButton"
        ).pack(side="left")

        self.excel_var = tk.StringVar(value="No Excel file selected.")
        ttk.Label(left, textvariable=self.excel_var, style="Muted.TLabel").grid(
            row=4, column=0, sticky="ew", pady=(8, 0)
        )

        window_row = ttk.Frame(left, style="Surface.TFrame")
        window_row.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        ttk.Label(window_row, text="Game Window Title Contains:", style="Card.TLabel").pack(
            side="left"
        )
        ttk.Entry(window_row, textvariable=self.window_title_var).pack(
            side="left", fill="x", expand=True, padx=(8, 0)
        )

        self.session_info_var = tk.StringVar(
            value="Session: stopped. Use Start Session then F8 in game (F9 stops)."
        )
        ttk.Label(left, textvariable=self.session_info_var, style="Muted.TLabel").grid(
            row=6, column=0, sticky="ew", pady=(8, 0)
        )

        tesseract_row = ttk.Frame(left, style="Surface.TFrame")
        tesseract_row.grid(row=7, column=0, sticky="ew", pady=(12, 0))
        ttk.Label(tesseract_row, text="Tesseract Path (optional):", style="Card.TLabel").pack(
            side="left"
        )
        self.tesseract_path_var = tk.StringVar()
        ttk.Entry(tesseract_row, textvariable=self.tesseract_path_var).pack(
            side="left", fill="x", expand=True, padx=(8, 0)
        )

        canvas = tk.Canvas(
            right,
            borderwidth=0,
            highlightthickness=0,
            background=COLORS["surface"],
        )
        scrollbar = ttk.Scrollbar(right, orient="vertical", command=canvas.yview)
        scrollable = ttk.Frame(canvas, style="Surface.TFrame")

        scrollable.bind(
            "<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.create_window((0, 0), window=scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for field in CANONICAL_FIELDS:
            row = ttk.Frame(scrollable, style="Surface.TFrame")
            row.pack(fill="x", pady=4)
            ttk.Label(row, text=field, width=34, style="Card.TLabel").pack(side="left")
            ttk.Entry(row, textvariable=self.value_vars[field]).pack(
                side="left", fill="x", expand=True
            )

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(self, textvariable=self.status_var, anchor="w", style="Status.TLabel").pack(
            fill="x", side="bottom"
        )

    def choose_image(self) -> None:
        path = filedialog.askopenfilename(
            title="Select screenshot",
            filetypes=[("Image files", "*.png *.jpg *.jpeg *.bmp *.webp")],
        )
        if not path:
            return

        self.selected_image = path
        self.current_nickname = None
        self.image_info_var.set(f"Screenshot: {path}")
        self._show_preview(path)
        self.status_var.set("Screenshot selected. Click 'Extract Stats'.")

    def _show_preview(self, path: str) -> None:
        image = Image.open(path)
        image.thumbnail((540, 540))
        self.preview_photo = ImageTk.PhotoImage(image)
        self.preview_label.configure(image=self.preview_photo)

    def choose_excel(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Choose Excel file",
            defaultextension=".xlsx",
            filetypes=[("Excel file", "*.xlsx")],
        )
        if not path:
            return
        self.excel_path = path
        self.excel_var.set(f"Excel file: {path}")
        self.status_var.set("Excel file selected.")

    def clear_fields(self) -> None:
        for var in self.value_vars.values():
            var.set("")
        self.current_nickname = None
        self.status_var.set("Fields cleared.")

    def extract(self) -> None:
        if not self.selected_image:
            messagebox.showwarning("No image", "Select a screenshot first.")
            return

        self._apply_tesseract_path()

        try:
            data, nickname = extract_stats_and_nickname(self.selected_image)
        except Exception as exc:
            messagebox.showerror("Extraction failed", str(exc))
            self.status_var.set("Extraction failed.")
            return

        self.current_nickname = nickname
        for field in CANONICAL_FIELDS:
            value = data.get(field)
            self.value_vars[field].set("" if value is None else str(value))

        self.status_var.set(
            f"Extracted {sum(1 for v in data.values() if v is not None)} fields. Review and save."
        )

    def _read_current_stats(self):
        stats = {}
        for field, var in self.value_vars.items():
            raw = var.get().strip()
            if not raw:
                continue
            cleaned = raw.replace(",", "")
            if not cleaned.isdigit():
                raise ValueError(f"Invalid number for '{field}': {raw}")
            stats[field] = int(cleaned)
        return stats

    def _apply_tesseract_path(self) -> None:
        tesseract_custom_path = self.tesseract_path_var.get().strip()
        if tesseract_custom_path:
            os.environ["TESSERACT_CMD"] = tesseract_custom_path

    def _safe_status(self, text: str) -> None:
        self.after(0, lambda: self.status_var.set(text))

    def _safe_session_info(self, text: str) -> None:
        self.after(0, lambda: self.session_info_var.set(text))

    def _safe_warning(self, title: str, text: str) -> None:
        self.after(0, lambda: messagebox.showwarning(title, text))

    def _safe_error(self, title: str, text: str) -> None:
        self.after(0, lambda: messagebox.showerror(title, text))

    def _looks_like_profile_stats(self, stats: dict) -> bool:
        required_markers = [
            "Units Killed",
            "Units Dead",
            "Units Healed",
            "Total Resources Gathered",
        ]
        marker_hits = sum(1 for key in required_markers if key in stats)
        return marker_hits >= 2 and len(stats) >= 6

    def start_hotkey_session(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("No Excel file", "Choose an Excel file first.")
            return
        if self.hotkey_listener is not None:
            self.status_var.set("Session already running. Press F8 to capture, F9 to stop.")
            return

        self._apply_tesseract_path()
        try:
            from pynput import keyboard  # type: ignore
        except Exception:
            messagebox.showerror(
                "Missing dependency",
                "Install dependencies again: pip install -r requirements.txt",
            )
            return

        self.hotkey_keyboard = keyboard
        self.hotkey_listener = keyboard.GlobalHotKeys(
            {
                "<f8>": self._on_hotkey_capture,
                "<f9>": self._on_hotkey_stop,
            }
        )
        self.hotkey_listener.start()
        self.session_active = True
        self._safe_status("Session started. In game, press F8 per profile. Press F9 to stop.")
        self._safe_session_info(
            "Session: running. Hotkeys -> F8 capture/save current profile, F9 stop session."
        )

    def stop_hotkey_session(self) -> None:
        if self.hotkey_listener is not None:
            try:
                self.hotkey_listener.stop()
            except Exception:
                pass
        self.hotkey_listener = None
        self.hotkey_keyboard = None
        self.session_active = False
        self.capture_in_progress = False
        self.status_var.set("Session stopped.")
        self.session_info_var.set(
            "Session: stopped. Use Start Session then F8 in game (F9 stops)."
        )

    def _on_hotkey_stop(self) -> None:
        self.after(0, self.stop_hotkey_session)

    def _on_hotkey_capture(self) -> None:
        self.after(0, self._capture_from_active_window)

    def _capture_from_active_window(self) -> None:
        if not self.session_active:
            return
        if self.capture_in_progress:
            self.status_var.set("Capture already running. Wait for completion.")
            return

        self.capture_in_progress = True
        self.status_var.set("Capturing active game window...")
        thread = threading.Thread(target=self._capture_worker, daemon=True)
        thread.start()

    def _capture_worker(self) -> None:
        try:
            image_path = self._grab_active_window_image()
            self.selected_image = image_path
            data, nickname = extract_stats_and_nickname(image_path)
            self.current_nickname = nickname

            if not self._looks_like_profile_stats(data):
                self._safe_status("Current window did not match profile stats layout.")
                return

            append_stats(self.excel_path, data, image_path, nickname=nickname)
            self.saved_rows += 1
            self.after(0, lambda: self._apply_extracted_ui(image_path, data, nickname))
            self._safe_status(
                f"Saved row #{self.saved_rows}. Keep browsing profiles and press F8 again."
            )
        except ExcelFileLockedError as exc:
            self._safe_warning(
                "Excel file is locked",
                str(exc),
            )
            self._safe_status("Excel file is locked. Close it, then press F8 again.")
        except Exception as exc:
            self._safe_error("Capture failed", str(exc))
            self._safe_status("Capture failed.")
        finally:
            self.capture_in_progress = False

    def _grab_active_window_image(self) -> str:
        try:
            import pygetwindow as gw  # type: ignore
        except Exception as exc:
            raise RuntimeError(
                "Missing dependency pygetwindow. Run: pip install -r requirements.txt"
            ) from exc

        active = gw.getActiveWindow()
        if active is None:
            raise RuntimeError("No active window found. Focus the game and press F8 again.")

        title = (active.title or "").strip()
        title_filter = self.window_title_var.get().strip().lower()
        if title_filter and title_filter not in title.lower():
            raise RuntimeError(
                f"Active window '{title}' does not match filter '{self.window_title_var.get()}'."
            )

        left = int(active.left)
        top = int(active.top)
        right = int(active.left + active.width)
        bottom = int(active.top + active.height)
        if right <= left or bottom <= top:
            raise RuntimeError("Active window has invalid size. Make sure it is visible.")

        capture_dir = Path.cwd() / "captures"
        capture_dir.mkdir(parents=True, exist_ok=True)
        image_path = capture_dir / f"profile_{datetime.now():%Y%m%d_%H%M%S}.png"

        screenshot = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
        screenshot.save(image_path)
        return str(image_path)

    def _apply_extracted_ui(self, image_path: str, data: dict, nickname: str | None = None) -> None:
        self.image_info_var.set(f"Screenshot: {image_path}")
        self.current_nickname = nickname
        self._show_preview(image_path)
        for field in CANONICAL_FIELDS:
            value = data.get(field)
            self.value_vars[field].set("" if value is None else str(value))

    def auto_capture_and_save(self) -> None:
        if not self.excel_path:
            messagebox.showwarning("No Excel file", "Choose an Excel file first.")
            return

        self._apply_tesseract_path()
        self.status_var.set(
            "Switch to the game profile screen. Capturing in 3 seconds..."
        )
        self.update_idletasks()
        self.after(3000, self._finish_auto_capture_and_save)

    def _finish_auto_capture_and_save(self) -> None:
        capture_dir = Path.cwd() / "captures"
        capture_dir.mkdir(parents=True, exist_ok=True)
        image_path = capture_dir / f"profile_{datetime.now():%Y%m%d_%H%M%S}.png"

        try:
            screenshot = ImageGrab.grab(all_screens=True)
            screenshot.save(image_path)
            self.selected_image = str(image_path)
            self.current_nickname = None
            self.image_info_var.set(f"Screenshot: {self.selected_image}")
            self._show_preview(self.selected_image)

            data, nickname = extract_stats_and_nickname(self.selected_image)
            self.current_nickname = nickname
            for field in CANONICAL_FIELDS:
                value = data.get(field)
                self.value_vars[field].set("" if value is None else str(value))

            if not self._looks_like_profile_stats(data):
                self.status_var.set("Capture did not match profile stats screen.")
                messagebox.showwarning(
                    "Profile not detected",
                    "Captured image does not look like the expected profile stats screen. "
                    "Open the profile page and try again.",
                )
                return

            append_stats(self.excel_path, data, self.selected_image, nickname=nickname)
        except ExcelFileLockedError as exc:
            messagebox.showwarning("Excel file is locked", str(exc))
            self.status_var.set("Excel file is locked. Close it and retry.")
            return
        except Exception as exc:
            messagebox.showerror("Auto capture failed", str(exc))
            self.status_var.set("Auto capture failed.")
            return

        self.status_var.set(f"Captured and saved row to: {self.excel_path}")
        messagebox.showinfo("Saved", "Profile captured, extracted, and saved to Excel.")

    def save_row(self) -> None:
        if not self.selected_image:
            messagebox.showwarning("No image", "Select a screenshot first.")
            return

        if not self.excel_path:
            messagebox.showwarning("No Excel file", "Choose an Excel file first.")
            return

        try:
            self._apply_tesseract_path()
            stats = self._read_current_stats()
            nickname = self.current_nickname
            if nickname is None and self.selected_image:
                nickname = extract_nickname(self.selected_image)
                self.current_nickname = nickname
            append_stats(self.excel_path, stats, self.selected_image, nickname=nickname)
        except ExcelFileLockedError as exc:
            messagebox.showwarning("Excel file is locked", str(exc))
            self.status_var.set("Excel file is locked. Close it and retry.")
            return
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))
            self.status_var.set("Save failed.")
            return

        self.status_var.set(f"Row appended to: {self.excel_path}")
        messagebox.showinfo("Saved", "Stats added to Excel.")

    def _on_close(self) -> None:
        self.stop_hotkey_session()
        self.destroy()


if __name__ == "__main__":
    app = StatTrackerApp()
    app.mainloop()
