import io
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from typing import Dict, Optional, Tuple

from PIL import Image, ImageDraw, ImageTk
import fitz  # PyMuPDF
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter


THEMES = {
    "dark": {
        "bg": "#050914",
        "bg_alt": "#0f172a",
        "card": "#111827",
        "muted": "#9ca3af",
        "text": "#f1f5f9",
        "subtext": "#cbd5f5",
        "accent": "#22c55e",
        "accent_hover": "#16a34a",
        "outline": "#1f2937",
        "shadow": "#020617",
    },
    "light": {
        "bg": "#f8fafc",
        "bg_alt": "#eef2ff",
        "card": "#ffffff",
        "muted": "#64748b",
        "text": "#0f172a",
        "subtext": "#475569",
        "accent": "#2563eb",
        "accent_hover": "#1d4ed8",
        "outline": "#d0d7e3",
        "shadow": "#cbd5f5",
    },
}

FONT_FAMILY = "Segoe UI"


def sanitize_filename(name: str) -> str:
    return re.sub(r"[\\/:\*\?\"<>\|]", "_", name)


def build_style(root: tk.Tk, palette: Dict[str, str]) -> None:
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("Card.TFrame", background=palette["card"])
    style.configure("Card.TLabel", background=palette["card"], foreground=palette["text"], font=(FONT_FAMILY, 11))
    style.configure("Heading.TLabel", background=palette["card"], foreground=palette["text"], font=(FONT_FAMILY, 14, "bold"))
    style.configure("Emphasis.TLabel", background=palette["card"], foreground=palette["accent"], font=(FONT_FAMILY, 11, "bold"))
    style.configure("Status.TLabel", background=palette["card"], foreground=palette["muted"], font=(FONT_FAMILY, 10))

    style.configure("Title.TLabel", background=palette["bg"], foreground=palette["text"], font=(FONT_FAMILY, 24, "bold"))
    style.configure("Subtitle.TLabel", background=palette["bg"], foreground=palette["muted"], font=(FONT_FAMILY, 11))

    style.configure("TLabel", background=palette["card"], foreground=palette["text"], font=(FONT_FAMILY, 11))
    style.configure("TButton", font=(FONT_FAMILY, 11), padding=8)
    style.configure(
        "Accent.TButton",
        font=(FONT_FAMILY, 11, "bold"),
        padding=10,
        foreground=palette["bg"],
        background=palette["accent"],
        borderwidth=0,
    )
    style.map(
        "Accent.TButton",
        background=[["active", palette["accent_hover"]], ["pressed", palette["accent"]]],
        foreground=[["disabled", palette["muted"]]],
    )

    style.configure("Surface.TFrame", background=palette["bg"])
    style.configure("Hero.TFrame", background=palette["bg"])
    style.configure("Metric.TLabel", background=palette["card"], foreground=palette["accent"], font=(FONT_FAMILY, 26, "bold"))
    style.configure("MetricCaption.TLabel", background=palette["card"], foreground=palette["muted"], font=(FONT_FAMILY, 10))

    style.configure("TCombobox", padding=6, font=(FONT_FAMILY, 11))
    style.map("TCombobox", fieldbackground=[["readonly", palette["card"]]], foreground=[["readonly", palette["text"]]])

    root.option_add("*TCombobox*Listbox*Font", (FONT_FAMILY, 11))
    root.option_add("*TCombobox*Listbox*Background", palette["card"])
    root.option_add("*TCombobox*Listbox*Foreground", palette["text"])


class AutoCertificateApp:
    """Interactive auto generator (no per-name review)."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Certificate Generator — Auto Mode")
        self.theme_name = tk.StringVar(value="dark")
        self.palette = THEMES[self.theme_name.get()]
        self.root.configure(bg=self.palette["bg"])
        self.root.minsize(920, 640)

        self.available_fonts = self.load_fonts()
        self.selected_font = tk.StringVar(value=next(iter(self.available_fonts)))

        self.pdf_path: Optional[str] = None
        self.names_path: Optional[str] = None
        self.names: list[str] = []
        self.template_pdf_bytes: Optional[bytes] = None

        self.pix: Optional[fitz.Pixmap] = None
        self.img: Optional[Image.Image] = None
        self.rect_info: Optional[dict] = None
        self.preview_photo: Optional[ImageTk.PhotoImage] = None

        self.output_dir: str = self.default_output_dir()

        self.pdf_status = tk.StringVar(value="No template selected")
        self.names_status = tk.StringVar(value="No names file selected")
        self.area_status = tk.StringVar(value="Area not selected")
        self.output_status = tk.StringVar(value=f"Output: {self.output_dir}")
        self.progress_status = tk.StringVar(value="Waiting to start…")
        self.names_metric = tk.StringVar(value="0")
        self.area_metric = tk.StringVar(value="Not selected")
        self.theme_button_text = tk.StringVar()

        self.progress_bar: Optional[ttk.Progressbar] = None
        self.log_text: Optional[tk.Text] = None
        self.preview_label: Optional[tk.Label] = None

        build_style(self.root, self.palette)
        self.update_theme_button_label()
        self.build_layout()
        self.refresh_metrics()
        self.root.mainloop()

    def update_theme_button_label(self) -> None:
        target = "Light" if self.theme_name.get() == "dark" else "Dark"
        self.theme_button_text.set(f"Switch to {target} Mode")

    def refresh_metrics(self) -> None:
        self.names_metric.set(str(len(self.names)))
        if self.rect_info:
            self.area_metric.set(f"{int(self.rect_info['rect_width'])} × {int(self.rect_info['rect_height'])} px")
        else:
            self.area_metric.set("Not selected")

    def default_output_dir(self) -> str:
        base = os.path.join(os.path.dirname(__file__), "output")
        os.makedirs(base, exist_ok=True)
        return base

    def append_log(self, message: str) -> None:
        if not self.log_text:
            return
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def clear_log(self) -> None:
        if self.log_text:
            self.log_text.delete("1.0", "end")

    def update_preview_snapshot(self) -> None:
        if not self.preview_label:
            return
        if not self.img:
            self.preview_label.configure(image="", text="Load a template to see a preview", anchor="center", fg=self.palette["muted"], bg=self.palette["card"])
            return

        preview_scale = min(620 / max(1, self.img.width), 620 / max(1, self.img.height), 1.0)
        render_img = self.img.copy()

        if self.rect_info:
            draw = ImageDraw.Draw(render_img)
            draw.rectangle(
                [
                    int(self.rect_info["tk_left"]),
                    int(self.rect_info["tk_top"]),
                    int(self.rect_info["tk_right"]),
                    int(self.rect_info["tk_bottom"]),
                ],
                outline=self.palette["accent"],
                width=3,
            )

        if abs(preview_scale - 1.0) > 1e-6:
            scaled_w = int(self.img.width * preview_scale)
            scaled_h = int(self.img.height * preview_scale)
            render_img = render_img.resize((scaled_w, scaled_h), Image.LANCZOS)

        self.preview_photo = ImageTk.PhotoImage(render_img)
        self.preview_label.configure(image=self.preview_photo, text="")
        self.preview_label.image = self.preview_photo

    def toggle_theme(self) -> None:
        new_theme = "light" if self.theme_name.get() == "dark" else "dark"
        self.theme_name.set(new_theme)
        self.palette = THEMES[new_theme]
        self.update_theme_button_label()
        self.rebuild_ui()

    def rebuild_ui(self) -> None:
        for child in self.root.winfo_children():
            child.destroy()
        build_style(self.root, self.palette)
        self.build_layout()
        self.refresh_metrics()
        self.update_ready_state()

    def load_fonts(self) -> Dict[str, str]:
        fonts_dir = os.path.join(os.path.dirname(__file__), "fonts")
        available_fonts: Dict[str, str] = {}

        if not os.path.exists(fonts_dir):
            os.makedirs(fonts_dir)

        for root_dir, _, files in os.walk(fonts_dir):
            family = os.path.basename(root_dir)
            for font_file in files:
                if font_file.lower().endswith((".ttf", ".otf")):
                    font_path = os.path.join(root_dir, font_file)
                    font_label = f"{family}/{os.path.splitext(font_file)[0]}" if family != "fonts" else os.path.splitext(font_file)[0]
                    try:
                        pdfmetrics.registerFont(TTFont(font_label, font_path))
                        available_fonts[font_label] = font_path
                    except Exception as exc:
                        print(f"Warning: Could not register font '{font_label}' at {font_path}: {exc}")

        if not available_fonts:
            messagebox.showerror("Font Error", "No fonts found in the 'fonts' directory. Add .ttf or .otf files and restart.")
            raise SystemExit(1)

        return available_fonts

    def build_layout(self) -> None:
        self.root.configure(bg=self.palette["bg"])

        container = ttk.Frame(self.root, style="Surface.TFrame")
        container.pack(fill="both", expand=True)

        canvas = tk.Canvas(container, bg=self.palette["bg"], highlightthickness=0, borderwidth=0)
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=v_scroll.set)
        canvas.pack(side="left", fill="both", expand=True)
        v_scroll.pack(side="right", fill="y")

        surface = ttk.Frame(canvas, style="Surface.TFrame")
        surface_id = canvas.create_window((0, 0), window=surface, anchor="nw")

        def _refresh_scroll_region(_event=None) -> None:
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _sync_surface_width(event) -> None:
            canvas.itemconfigure(surface_id, width=event.width)

        def _on_mousewheel(event) -> None:
            if event.widget.winfo_exists():
                canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        surface.bind("<Configure>", _refresh_scroll_region)
        canvas.bind("<Configure>", _sync_surface_width)
        surface.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", _on_mousewheel))
        surface.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))

        hero = ttk.Frame(surface, style="Hero.TFrame", padding=(32, 24))
        hero.pack(fill="x")
        hero.columnconfigure(0, weight=1)
        hero.columnconfigure(1, weight=0)

        hero_text = ttk.Frame(hero, style="Hero.TFrame")
        hero_text.grid(row=0, column=0, sticky="nsew")
        ttk.Label(hero_text, text="Certificate Automation", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            hero_text,
            text="Load your template, mark placement, and export polished certificates automatically.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(8, 0))

        metrics_frame = ttk.Frame(hero_text, style="Hero.TFrame")
        metrics_frame.pack(fill="x", pady=(20, 0))
        metrics_frame.columnconfigure(0, weight=1)
        metrics_frame.columnconfigure(1, weight=1)

        names_card = ttk.Frame(metrics_frame, style="Card.TFrame", padding=16)
        names_card.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        ttk.Label(names_card, textvariable=self.names_metric, style="Metric.TLabel").pack(anchor="w")
        ttk.Label(names_card, text="Names loaded", style="MetricCaption.TLabel").pack(anchor="w")

        area_card = ttk.Frame(metrics_frame, style="Card.TFrame", padding=16)
        area_card.grid(row=0, column=1, sticky="nsew")
        ttk.Label(area_card, textvariable=self.area_metric, style="Metric.TLabel").pack(anchor="w")
        ttk.Label(area_card, text="Placement window", style="MetricCaption.TLabel").pack(anchor="w")

        hero_actions = ttk.Frame(hero, style="Hero.TFrame")
        hero_actions.grid(row=0, column=1, sticky="ne")
        ttk.Button(hero_actions, textvariable=self.theme_button_text, style="Accent.TButton", command=self.toggle_theme).pack(anchor="e")

        content = ttk.Frame(surface, style="Surface.TFrame", padding=(28, 22))
        content.pack(fill="both", expand=True)
        content.columnconfigure(0, weight=2)
        content.columnconfigure(1, weight=1)
        content.rowconfigure(1, weight=1)

        steps_card = ttk.Frame(content, style="Card.TFrame", padding=24)
        steps_card.grid(row=0, column=0, rowspan=2, sticky="nsew", padx=(0, 18))
        steps_card.columnconfigure(0, weight=1)

        self.build_steps_panel(steps_card)

        preview_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        preview_card.grid(row=0, column=1, sticky="nsew", pady=(0, 14))
        preview_card.columnconfigure(0, weight=1)
        ttk.Label(preview_card, text="Template preview", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Separator(preview_card).grid(row=1, column=0, sticky="ew", pady=(8, 12))

        self.preview_label = tk.Label(preview_card, bg=self.palette["card"], bd=1, relief="solid")
        self.preview_label.grid(row=2, column=0, sticky="nsew")
        preview_card.rowconfigure(2, weight=1)
        ttk.Label(
            preview_card,
            text="Preview updates after loading a template and selecting placement.",
            style="Status.TLabel",
            wraplength=420,
        ).grid(row=3, column=0, sticky="w", pady=(10, 0))

        progress_card = ttk.Frame(content, style="Card.TFrame", padding=20)
        progress_card.grid(row=1, column=1, sticky="nsew")
        progress_card.columnconfigure(0, weight=1)

        ttk.Label(progress_card, text="Generation status", style="Heading.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Separator(progress_card).grid(row=1, column=0, sticky="ew", pady=(8, 10))

        summary_rows = [
            ("Template", self.pdf_status),
            ("Names", self.names_status),
            ("Placement", self.area_status),
            ("Output", self.output_status),
        ]
        for idx, (label_text, text_var) in enumerate(summary_rows, start=2):
            row_frame = ttk.Frame(progress_card, style="Card.TFrame")
            row_frame.grid(row=idx, column=0, sticky="ew", pady=2)
            row_frame.columnconfigure(1, weight=1)
            ttk.Label(row_frame, text=label_text, style="Emphasis.TLabel").grid(row=0, column=0, sticky="nw", padx=(0, 8))
            ttk.Label(row_frame, textvariable=text_var, style="Status.TLabel", wraplength=320).grid(row=0, column=1, sticky="w")

        self.progress_bar = ttk.Progressbar(progress_card, mode="determinate", maximum=1)
        self.progress_bar.grid(row=5, column=0, sticky="ew", pady=(14, 6))

        ttk.Label(progress_card, textvariable=self.progress_status, style="Status.TLabel", wraplength=360).grid(
            row=6, column=0, sticky="w"
        )

        self.log_text = tk.Text(progress_card, height=10, wrap="word", bg=self.palette["card"], fg=self.palette["text"], bd=1, relief="solid")
        self.log_text.grid(row=7, column=0, sticky="nsew", pady=(10, 0))
        self.log_text.configure(state="normal")
        progress_card.rowconfigure(7, weight=1)

        self.update_preview_snapshot()

    def build_steps_panel(self, container: ttk.Frame) -> None:
        ttk.Label(container, text="Workflow", style="Heading.TLabel").grid(row=0, column=0, sticky="w")

        files_section = ttk.Frame(container, style="Card.TFrame")
        files_section.grid(row=1, column=0, sticky="ew", pady=(12, 0))
        files_section.columnconfigure(1, weight=1)
        ttk.Label(files_section, text="1. Source files", style="Emphasis.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Button(files_section, text="Choose template PDF", style="Accent.TButton", command=self.choose_pdf).grid(row=1, column=0, padx=(0, 12), pady=4)
        ttk.Label(files_section, textvariable=self.pdf_status, style="Card.TLabel").grid(row=1, column=1, sticky="w")
        ttk.Button(files_section, text="Choose names TXT", command=self.choose_names).grid(row=2, column=0, padx=(0, 12), pady=4)
        ttk.Label(files_section, textvariable=self.names_status, style="Card.TLabel").grid(row=2, column=1, sticky="w")

        font_section = ttk.Frame(container, style="Card.TFrame")
        font_section.grid(row=2, column=0, sticky="ew", pady=(18, 0))
        font_section.columnconfigure(1, weight=1)
        ttk.Label(font_section, text="2. Typography", style="Emphasis.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Label(font_section, text="Certificate font", style="Card.TLabel").grid(row=1, column=0, sticky="w", pady=2)
        self.font_selector = ttk.Combobox(
            font_section,
            textvariable=self.selected_font,
            values=list(self.available_fonts.keys()),
            state="readonly",
        )
        self.font_selector.grid(row=1, column=1, sticky="ew", padx=(10, 0))
        ttk.Label(
            font_section,
            text="Drop new .ttf/.otf files into the fonts/ folder to expand this list instantly.",
            style="Status.TLabel",
            wraplength=420,
        ).grid(row=2, column=0, columnspan=2, sticky="w", pady=(8, 0))

        area_section = ttk.Frame(container, style="Card.TFrame")
        area_section.grid(row=3, column=0, sticky="ew", pady=(18, 0))
        area_section.columnconfigure(1, weight=1)
        ttk.Label(area_section, text="3. Placement", style="Emphasis.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Button(area_section, text="Select placement area", command=self.select_area).grid(row=1, column=0, padx=(0, 12), pady=4)
        ttk.Label(area_section, textvariable=self.area_status, style="Card.TLabel").grid(row=1, column=1, sticky="w")

        output_section = ttk.Frame(container, style="Card.TFrame")
        output_section.grid(row=4, column=0, sticky="ew", pady=(18, 0))
        output_section.columnconfigure(1, weight=1)
        ttk.Label(output_section, text="4. Output", style="Emphasis.TLabel").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Button(output_section, text="Choose output folder", command=self.choose_output_dir).grid(row=1, column=0, padx=(0, 12), pady=4)
        ttk.Label(output_section, textvariable=self.output_status, style="Card.TLabel", wraplength=420).grid(row=1, column=1, sticky="w")

        review_section = ttk.Frame(container, style="Card.TFrame")
        review_section.grid(row=5, column=0, sticky="ew", pady=(18, 0))
        review_section.columnconfigure(0, weight=1)
        ttk.Label(review_section, text="5. Generate", style="Emphasis.TLabel").grid(row=0, column=0, sticky="w", pady=(0, 6))
        self.start_button = ttk.Button(review_section, text="Generate certificates", style="Accent.TButton", command=self.start_generation)
        self.start_button.grid(row=1, column=0, sticky="ew")
        self.start_button.state(["disabled"])
        ttk.Label(review_section, textvariable=self.progress_status, style="Status.TLabel", wraplength=420).grid(
            row=2, column=0, sticky="w", pady=(10, 0)
        )

    def choose_pdf(self) -> None:
        path = filedialog.askopenfilename(title="Select Certificate Template PDF", filetypes=[["PDF Files", "*.pdf"]])
        if not path:
            return

        if not self.load_pdf_preview(path):
            return

        self.pdf_path = path
        self.area_status.set("Area not selected")
        self.rect_info = None
        self.pdf_status.set(f"Template: {os.path.basename(path)} ({self.pix.width} x {self.pix.height}px)")
        self.refresh_metrics()
        self.update_ready_state()
        self.append_log(f"Loaded template: {os.path.basename(path)}")
        self.update_preview_snapshot()

    def choose_names(self) -> None:
        path = filedialog.askopenfilename(title="Select Names TXT File", filetypes=[["Text Files", "*.txt"]])
        if not path:
            return

        with open(path, "r", encoding="utf-8") as handle:
            names = [line.strip() for line in handle.readlines() if line.strip()]

        if not names:
            messagebox.showwarning("Names", "The selected file has no names.")
            return

        self.names_path = path
        self.names = names
        self.names_status.set(f"{len(names)} name(s) ready from {os.path.basename(path)}")
        self.refresh_metrics()
        self.update_ready_state()
        self.append_log(f"Loaded {len(names)} names from {os.path.basename(path)}")

    def choose_output_dir(self) -> None:
        path = filedialog.askdirectory(title="Select output folder")
        if not path:
            return
        os.makedirs(path, exist_ok=True)
        self.output_dir = path
        self.output_status.set(f"Output: {path}")
        self.append_log(f"Output folder set to {path}")
        self.update_ready_state()

    def load_pdf_preview(self, path: str) -> bool:
        try:
            doc = fitz.open(path)
            page = doc.load_page(0)
            self.pix = page.get_pixmap()
            self.img = Image.frombytes("RGB", [self.pix.width, self.pix.height], self.pix.samples)
            return True
        except Exception as exc:
            messagebox.showerror("PDF Error", f"Unable to open the template.\n{exc}")
            return False
        finally:
            try:
                doc.close()
            except Exception:
                pass

    def select_area(self) -> None:
        if not self.pix or not self.img:
            messagebox.showinfo("Template required", "Choose a template PDF first.")
            return

        rect_start, rect_end = self.show_rectangle_selection(self.img, self.pix)
        if not rect_start or not rect_end:
            self.area_status.set("Area not selected — draw a rectangle and confirm")
            self.refresh_metrics()
            self.update_ready_state()
            return

        if rect_start == rect_end:
            self.area_status.set("Area not selected — rectangle has zero size")
            self.refresh_metrics()
            self.update_ready_state()
            return

        norm_start = (int(rect_start[0]), int(rect_start[1]))
        norm_end = (int(rect_end[0]), int(rect_end[1]))
        self.rect_info = self.calculate_rect(norm_start, norm_end, self.pix.height)
        rect_width = int(self.rect_info["rect_width"])
        rect_height = int(self.rect_info["rect_height"])
        self.area_status.set(f"Selected area: {rect_width} x {rect_height} px")
        self.refresh_metrics()
        self.update_ready_state()
        self.append_log(f"Placement area set: {rect_width} x {rect_height} px")
        self.update_preview_snapshot()

    def start_generation(self) -> None:
        if not self.pdf_path or not self.names or not self.rect_info or not self.output_dir:
            messagebox.showinfo("Missing info", "Please finish selecting files, font, area, and output folder first.")
            return

        os.makedirs(self.output_dir, exist_ok=True)

        self.clear_log()
        if self.progress_bar:
            self.progress_bar.configure(maximum=max(1, len(self.names)))
            self.progress_bar['value'] = 0
        self.progress_status.set("Generating certificates…")
        self.root.update_idletasks()

        with open(self.pdf_path, "rb") as template_file:
            self.template_pdf_bytes = template_file.read()

        for idx, name in enumerate(self.names, start=1):
            font_size = self.calculate_initial_font_size(name, self.selected_font.get())
            text_width = stringWidth(name, self.selected_font.get(), font_size)
            x_text = self.rect_info["x_left"] + (self.rect_info["rect_width"] - text_width) / 2
            y_text = self.rect_info["y_bottom"] + (self.rect_info["rect_height"] - font_size) / 2
            self.generate_certificate(name, self.selected_font.get(), font_size, x_text, y_text)
            self.progress_status.set(f"Saved {idx}/{len(self.names)}: {name}")
            if self.progress_bar:
                self.progress_bar['value'] = idx
            self.append_log(f"✓ {name} — font {self.selected_font.get()} @ {font_size}pt")
            self.root.update_idletasks()

        self.progress_status.set("All certificates generated automatically!")
        self.append_log("Completed all certificates.")

    def calculate_rect(self, rect_start: Tuple[int, int], rect_end: Tuple[int, int], pix_height: int) -> dict:
        x1, y1 = rect_start
        x2, y2 = rect_end
        pdf_x1, pdf_y1 = x1, pix_height - y1
        pdf_x2, pdf_y2 = x2, pix_height - y2

        rect_width = abs(pdf_x2 - pdf_x1)
        rect_height = abs(pdf_y2 - pdf_y1)
        x_left = min(pdf_x1, pdf_x2)
        y_bottom = min(pdf_y1, pdf_y2)

        return {
            "rect_width": rect_width,
            "rect_height": rect_height,
            "x_left": x_left,
            "y_bottom": y_bottom,
            "tk_left": min(x1, x2),
            "tk_top": min(y1, y2),
            "tk_right": max(x1, x2),
            "tk_bottom": max(y1, y2),
        }

    def calculate_initial_font_size(self, text_value: str, selected_font: str) -> int:
        max_size = int(min(self.rect_info["rect_height"], 120)) or 5
        size = max_size
        while size >= 5:
            text_width = stringWidth(text_value, selected_font, size)
            if text_width <= self.rect_info["rect_width"] and size <= self.rect_info["rect_height"]:
                return size
            size -= 1
        return 5

    def generate_certificate(self, name_text: str, font_choice: str, font_size: int, pdf_x: float, pdf_y: float) -> None:
        packet = io.BytesIO()
        c = canvas.Canvas(packet, pagesize=(self.pix.width, self.pix.height))
        c.setFont(font_choice, font_size)
        c.setFillColorRGB(1, 1, 1)
        c.drawString(pdf_x, pdf_y, name_text)
        c.save()

        packet.seek(0)
        template_reader = PdfReader(io.BytesIO(self.template_pdf_bytes))
        overlay_reader = PdfReader(packet)
        writer = PdfWriter()

        base_page = template_reader.pages[0]
        base_page.merge_page(overlay_reader.pages[0])
        writer.add_page(base_page)

        output_filename = os.path.join(self.output_dir, f"certificate_{sanitize_filename(name_text)}.pdf")
        with open(output_filename, "wb") as out_file:
            writer.write(out_file)

    def show_rectangle_selection(self, img: Image.Image, pix: fitz.Pixmap) -> Tuple[Optional[Tuple[int, int]], Optional[Tuple[int, int]]]:
        selection = {"start": None, "end": None, "coords": None}
        rect_id = None

        win = tk.Toplevel(self.root)
        win.title("Mark name placement area")
        win.configure(bg=self.palette["card"])

        window_width = min(pix.width + 60, win.winfo_screenwidth() - 40)
        window_height = min(pix.height + 180, win.winfo_screenheight() - 80)
        win.geometry(f"{window_width}x{window_height}")
        win.resizable(False, False)

        ttk.Label(win, text="Draw the rectangle where the name should appear", style="Heading.TLabel").pack(anchor="w", padx=20, pady=(16, 4))
        helper = ttk.Label(
            win,
            text="Click and drag to set the area, then press Confirm selection.",
            style="Status.TLabel",
        )
        helper.pack(anchor="w", padx=20)

        canvas_widget = tk.Canvas(
            win,
            width=pix.width,
            height=pix.height,
            bg="#0b1220",
            highlightthickness=2,
            highlightbackground=self.palette["outline"],
        )
        canvas_widget.pack(pady=10)
        tk_img = ImageTk.PhotoImage(img)
        canvas_widget.create_image(0, 0, anchor="nw", image=tk_img)

        confirm_btn = ttk.Button(win, text="Confirm selection", style="Accent.TButton")
        confirm_btn.pack(pady=(4, 12))
        confirm_btn.state(["disabled"])

        def on_mouse_down(event):
            nonlocal rect_id
            selection["start"] = (event.x, event.y)
            rect_id = canvas_widget.create_rectangle(event.x, event.y, event.x, event.y, outline=self.palette["accent"], width=2)

        def on_mouse_move(event):
            nonlocal rect_id
            if rect_id:
                canvas_widget.coords(rect_id, selection["start"][0], selection["start"][1], event.x, event.y)

        def on_mouse_up(event):
            selection["end"] = (event.x, event.y)
            if rect_id:
                selection["coords"] = canvas_widget.coords(rect_id)
            helper.configure(text=f"Selected from {selection['start']} to {selection['end']}. Click Confirm selection.")
            confirm_btn.state(["!disabled"])

        def on_confirm():
            if rect_id and not selection["coords"]:
                selection["coords"] = canvas_widget.coords(rect_id)
            win.destroy()

        def on_close():
            if rect_id and not selection["coords"]:
                selection["coords"] = canvas_widget.coords(rect_id)
            win.destroy()

        canvas_widget.bind("<Button-1>", on_mouse_down)
        canvas_widget.bind("<B1-Motion>", on_mouse_move)
        canvas_widget.bind("<ButtonRelease-1>", on_mouse_up)
        confirm_btn.configure(command=on_confirm)
        win.protocol("WM_DELETE_WINDOW", on_close)

        win.update_idletasks()
        x_pos = (win.winfo_screenwidth() // 2) - (window_width // 2)
        y_pos = (win.winfo_screenheight() // 2) - (window_height // 2)
        win.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

        win.grab_set()
        self.root.wait_window(win)

        if selection["start"] and selection["end"]:
            return selection["start"], selection["end"]
        if selection["coords"]:
            x1, y1, x2, y2 = selection["coords"]
            return (int(x1), int(y1)), (int(x2), int(y2))
        return None, None

    def update_ready_state(self) -> None:
        self.refresh_metrics()
        ready = bool(self.pdf_path and self.names and self.rect_info and self.output_dir)
        if ready:
            self.start_button.state(["!disabled"])
            self.progress_status.set("Ready to generate automatically")
        else:
            missing = []
            if not self.pdf_path:
                missing.append("template PDF")
            if not self.names:
                missing.append("names file")
            if not self.rect_info:
                missing.append("placement area")
            if not self.output_dir:
                missing.append("output folder")
            self.start_button.state(["disabled"])
            self.progress_status.set(f"Waiting: select {', '.join(missing)}")


if __name__ == "__main__":
    AutoCertificateApp()
