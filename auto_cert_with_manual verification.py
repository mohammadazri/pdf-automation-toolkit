"""Modernized certificate generator with guided, user-friendly flow."""

import io
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from typing import Dict, Optional, Tuple

from PIL import Image, ImageDraw, ImageFont, ImageTk
import fitz  # PyMuPDF
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter


PALETTE = {
    "bg": "#0f172a",
    "card": "#111827",
    "muted": "#9ca3af",
    "text": "#e5e7eb",
    "accent": "#22c55e",
    "accent_hover": "#16a34a",
    "outline": "#1f2937",
}

FONT_FAMILY = "Segoe UI"


def build_style(root: tk.Tk) -> None:
    """Apply a modern, minimal theme to ttk widgets."""

    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure("Card.TFrame", background=PALETTE["card"])
    style.configure("Card.TLabel", background=PALETTE["card"], foreground=PALETTE["text"], font=(FONT_FAMILY, 11))
    style.configure("Heading.TLabel", background=PALETTE["card"], foreground=PALETTE["text"], font=(FONT_FAMILY, 14, "bold"))
    style.configure("Emphasis.TLabel", background=PALETTE["card"], foreground=PALETTE["accent"], font=(FONT_FAMILY, 11, "bold"))
    style.configure("Status.TLabel", background=PALETTE["card"], foreground=PALETTE["muted"], font=(FONT_FAMILY, 10))

    style.configure("Title.TLabel", background=PALETTE["bg"], foreground=PALETTE["text"], font=(FONT_FAMILY, 22, "bold"))
    style.configure("Subtitle.TLabel", background=PALETTE["bg"], foreground=PALETTE["muted"], font=(FONT_FAMILY, 11))

    style.configure("TLabel", background=PALETTE["card"], foreground=PALETTE["text"], font=(FONT_FAMILY, 11))
    style.configure("TButton", font=(FONT_FAMILY, 11), padding=8)
    style.configure(
        "Accent.TButton",
        font=(FONT_FAMILY, 11, "bold"),
        padding=10,
        foreground=PALETTE["bg"],
        background=PALETTE["accent"],
        borderwidth=0,
    )
    style.map(
        "Accent.TButton",
        background=[("active", PALETTE["accent_hover"]), ("pressed", PALETTE["accent"])],
        foreground=[("disabled", PALETTE["muted"])],
    )

    style.configure("TCombobox", padding=6, font=(FONT_FAMILY, 11))
    style.map("TCombobox", fieldbackground=[("readonly", PALETTE["card"])], foreground=[("readonly", PALETTE["text"])])

    root.option_add("*TCombobox*Listbox*Font", (FONT_FAMILY, 11))
    root.option_add("*TCombobox*Listbox*Background", PALETTE["card"])
    root.option_add("*TCombobox*Listbox*Foreground", PALETTE["text"])


class CertificateApp:
    """Interactive, modern certificate generator."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Certificate Generator — Modern UI")
        self.root.configure(bg=PALETTE["bg"])
        self.root.minsize(820, 520)

        self.available_fonts = self.load_fonts()
        self.selected_font = tk.StringVar(value=next(iter(self.available_fonts)))

        self.pdf_path: Optional[str] = None
        self.names_path: Optional[str] = None
        self.names: list[str] = []
        self.template_pdf_bytes: Optional[bytes] = None

        self.pix: Optional[fitz.Pixmap] = None
        self.img: Optional[Image.Image] = None
        self.rect_info: Optional[dict] = None
        self.last_font_size: Optional[int] = None

        build_style(self.root)
        self.build_layout()
        self.root.mainloop()

    def load_fonts(self) -> Dict[str, str]:
        """Auto-register TTF/OTF fonts from the bundled fonts directory."""

        fonts_dir = os.path.join(os.path.dirname(__file__), "fonts")
        available_fonts: Dict[str, str] = {}

        if not os.path.exists(fonts_dir):
            os.makedirs(fonts_dir)

        for root_dir, _, files in os.walk(fonts_dir):
            family = os.path.basename(root_dir)
            for font_file in files:
                if font_file.lower().endswith((".ttf", ".otf")):
                    font_path = os.path.join(root_dir, font_file)
                    font_label = (
                        f"{family}/{os.path.splitext(font_file)[0]}" if family != "fonts" else os.path.splitext(font_file)[0]
                    )
                    try:
                        pdfmetrics.registerFont(TTFont(font_label, font_path))
                        available_fonts[font_label] = font_path
                    except Exception as exc:  # font registration failed
                        print(f"Warning: Could not register font '{font_label}' at {font_path}: {exc}")

        if not available_fonts:
            messagebox.showerror(
                "Font Error",
                "No fonts found in the 'fonts' directory. Add .ttf or .otf files and restart.",
            )
            raise SystemExit(1)

        return available_fonts

    def build_layout(self) -> None:
        """Compose the guided workflow UI."""

        header = tk.Frame(self.root, bg=PALETTE["bg"], padx=28, pady=18)
        header.pack(fill="x")

        ttk.Label(header, text="Certificate Automation", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            header,
            text="A guided, modern flow to pick files, choose fonts, mark placement, and review outputs.",
            style="Subtitle.TLabel",
        ).pack(anchor="w", pady=(4, 0))

        card = ttk.Frame(self.root, style="Card.TFrame", padding=22)
        card.pack(fill="both", expand=True, padx=22, pady=(0, 22))
        card.columnconfigure(0, weight=1)
        card.columnconfigure(1, weight=1)

        # Step 1 — file selection
        ttk.Label(card, text="1. Pick your files", style="Heading.TLabel").grid(row=0, column=0, sticky="w")

        file_row = ttk.Frame(card, style="Card.TFrame")
        file_row.grid(row=1, column=0, sticky="ew", pady=(8, 18))
        file_row.columnconfigure(1, weight=1)

        self.pdf_status = tk.StringVar(value="No template selected")
        self.names_status = tk.StringVar(value="No names file selected")

        ttk.Button(file_row, text="Choose template PDF", style="Accent.TButton", command=self.choose_pdf).grid(
            row=0, column=0, padx=(0, 12), pady=4
        )
        ttk.Label(file_row, textvariable=self.pdf_status, style="Card.TLabel").grid(row=0, column=1, sticky="w")

        ttk.Button(file_row, text="Choose names TXT", command=self.choose_names).grid(row=1, column=0, padx=(0, 12), pady=4)
        ttk.Label(file_row, textvariable=self.names_status, style="Card.TLabel").grid(row=1, column=1, sticky="w")

        # Step 2 — font selection
        ttk.Label(card, text="2. Pick your font", style="Heading.TLabel").grid(row=2, column=0, sticky="w")

        font_row = ttk.Frame(card, style="Card.TFrame")
        font_row.grid(row=3, column=0, sticky="ew", pady=(8, 18))
        font_row.columnconfigure(1, weight=1)

        ttk.Label(font_row, text="Certificate font", style="Card.TLabel").grid(row=0, column=0, sticky="w", pady=2)
        self.font_selector = ttk.Combobox(
            font_row,
            textvariable=self.selected_font,
            values=list(self.available_fonts.keys()),
            state="readonly",
        )
        self.font_selector.grid(row=0, column=1, sticky="ew", padx=(10, 0))

        ttk.Label(
            font_row,
            text="Fonts are auto-loaded from the 'fonts' folder. Add your own TTF/OTF files to expand the list.",
            style="Status.TLabel",
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(6, 0))

        # Step 3 — area selection
        ttk.Label(card, text="3. Mark the name area", style="Heading.TLabel").grid(row=4, column=0, sticky="w")

        area_row = ttk.Frame(card, style="Card.TFrame")
        area_row.grid(row=5, column=0, sticky="ew", pady=(8, 18))
        area_row.columnconfigure(1, weight=1)

        self.area_status = tk.StringVar(value="Area not selected")
        ttk.Button(area_row, text="Select placement area", command=self.select_area).grid(row=0, column=0, padx=(0, 12), pady=4)
        ttk.Label(area_row, textvariable=self.area_status, style="Card.TLabel").grid(row=0, column=1, sticky="w")

        # Step 4 — generate
        ttk.Label(card, text="4. Review & export", style="Heading.TLabel").grid(row=6, column=0, sticky="w")

        action_row = ttk.Frame(card, style="Card.TFrame")
        action_row.grid(row=7, column=0, sticky="ew", pady=(8, 0))
        action_row.columnconfigure(0, weight=1)

        self.start_button = ttk.Button(action_row, text="Start manual review", style="Accent.TButton", command=self.start_review)
        self.start_button.grid(row=0, column=0, sticky="ew")
        self.start_button.state(["disabled"])

        self.progress_status = tk.StringVar(value="Waiting to start…")
        ttk.Label(action_row, textvariable=self.progress_status, style="Status.TLabel").grid(row=1, column=0, sticky="w", pady=(8, 0))

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
        self.update_ready_state()

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

        print("[DEBUG] select_area invoked")
        rect_start, rect_end = self.show_rectangle_selection(self.img, self.pix)
        print("[DEBUG] select_area returned start=%s end=%s" % (rect_start, rect_end))
        if not rect_start or not rect_end:
            self.area_status.set("Area not selected — draw a rectangle and confirm")
            print("[DEBUG] Rectangle selection missing: rect_start=%s rect_end=%s" % (rect_start, rect_end))
            self.update_ready_state()
            return

        # Avoid zero-size selections
        if rect_start == rect_end:
            self.area_status.set("Area not selected — rectangle has zero size")
            print("[DEBUG] Rectangle selection zero-size: rect_start=%s rect_end=%s" % (rect_start, rect_end))
            self.update_ready_state()
            return

        # Normalize to ints in case of float coords from canvas
        norm_start = (int(rect_start[0]), int(rect_start[1]))
        norm_end = (int(rect_end[0]), int(rect_end[1]))
        self.rect_info = self.calculate_rect(norm_start, norm_end, self.pix.height)
        rect_width = int(self.rect_info["rect_width"])
        rect_height = int(self.rect_info["rect_height"])
        self.area_status.set(f"Selected area: {rect_width} x {rect_height} px")
        print("[DEBUG] Rectangle selection accepted: start=%s end=%s size=%sx%s" % (norm_start, norm_end, rect_width, rect_height))
        self.update_ready_state()

    def start_review(self) -> None:
        if not self.pdf_path or not self.names or not self.rect_info:
            messagebox.showinfo("Missing info", "Please finish selecting files, font, and area first.")
            return

        with open(self.pdf_path, "rb") as template_file:
            self.template_pdf_bytes = template_file.read()

        last_font = self.selected_font.get()
        self.last_font_size = None

        for name in self.names:
            initial_size = self.calculate_initial_font_size(name, last_font)
            if self.last_font_size:
                initial_size = max(5, min(int(self.rect_info["rect_height"]), self.last_font_size))

            review_result = self.review_name(name, last_font, initial_size)
            if not review_result:
                self.progress_status.set("Certificate generation cancelled by user.")
                break

            last_font = review_result["font"]
            self.last_font_size = review_result["size"]

            self.generate_certificate(
                name,
                review_result["font"],
                review_result["size"],
                review_result["x"],
                review_result["y"],
            )
            self.progress_status.set(f"Saved certificate for {name}")
        else:
            self.progress_status.set("All certificates generated with manual review!")

    def calculate_rect(self, rect_start: Tuple[int, int], rect_end: Tuple[int, int], pix_height: int) -> dict:
        print("[DEBUG] calculate_rect inputs start=%s end=%s pix_height=%s" % (rect_start, rect_end, pix_height))
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
        """Find the largest font size that fits the selected rectangle."""

        max_size = int(min(self.rect_info["rect_height"], 120)) or 5
        size = max_size
        while size >= 5:
            text_width = stringWidth(text_value, selected_font, size)
            if text_width <= self.rect_info["rect_width"] and size <= self.rect_info["rect_height"]:
                return size
            size -= 1
        return 5

    def review_name(self, name_text: str, default_font: str, default_size: int) -> Optional[dict]:
        """Manual review dialog with live preview."""

        result: dict = {}
        window = tk.Toplevel(self.root)
        window.title(f"Review Certificate — {name_text}")
        window.configure(bg=PALETTE["card"], padx=12, pady=12)
        window.columnconfigure(0, weight=1)
        window.columnconfigure(1, weight=0)
        window.rowconfigure(0, weight=1)

        preview_scale = min(1.0, 820 / max(1, self.pix.width), 560 / max(1, self.pix.height))
        display_width = int(self.pix.width * preview_scale)
        display_height = int(self.pix.height * preview_scale)

        use_original_size = abs(preview_scale - 1.0) < 1e-6
        preview_image = self.img if use_original_size else self.img.resize((display_width, display_height), Image.LANCZOS)
        preview_photo = ImageTk.PhotoImage(preview_image)

        preview_label = tk.Label(window, image=preview_photo, bd=2, relief="sunken")
        preview_label.grid(row=0, column=0, rowspan=6, sticky="nsew", padx=(0, 12))
        preview_label.image = preview_photo

        controls_frame = ttk.Frame(window, style="Card.TFrame")
        controls_frame.grid(row=0, column=1, sticky="n")
        controls_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(controls_frame, style="Card.TFrame")
        button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

        font_var = tk.StringVar(value=default_font)
        size_var = tk.IntVar(value=max(5, default_size))
        x_offset_limit = int(max(10, self.rect_info["rect_width"] / 2))
        y_offset_limit = int(max(10, self.rect_info["rect_height"] / 2))
        x_offset_var = tk.DoubleVar(value=0)
        y_offset_var = tk.DoubleVar(value=0)

        ttk.Label(controls_frame, text="Font", style="Card.TLabel").grid(row=1, column=0, sticky="w", padx=5, pady=5)
        font_selector = ttk.Combobox(
            controls_frame,
            textvariable=font_var,
            values=list(self.available_fonts.keys()),
            state="readonly",
        )
        font_selector.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
        font_selector.bind("<<ComboboxSelected>>", lambda _: update_preview())

        ttk.Label(controls_frame, text="Font size", style="Card.TLabel").grid(row=2, column=0, sticky="w", padx=5, pady=5)
        size_scale_max = max(int(self.rect_info["rect_height"]), default_size + 40, 10)
        size_scale = tk.Scale(
            controls_frame,
            from_=5,
            to=size_scale_max,
            orient="horizontal",
            variable=size_var,
            command=lambda _: update_preview(),
            bg=PALETTE["card"],
            troughcolor=PALETTE["outline"],
            highlightthickness=0,
            sliderrelief="flat",
        )
        size_scale.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(controls_frame, text="Horizontal offset", style="Card.TLabel").grid(row=3, column=0, sticky="w", padx=5, pady=5)
        x_scale = tk.Scale(
            controls_frame,
            from_=-x_offset_limit,
            to=x_offset_limit,
            orient="horizontal",
            resolution=1,
            variable=x_offset_var,
            command=lambda _: update_preview(),
            bg=PALETTE["card"],
            troughcolor=PALETTE["outline"],
            highlightthickness=0,
            sliderrelief="flat",
        )
        x_scale.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

        ttk.Label(controls_frame, text="Vertical offset", style="Card.TLabel").grid(row=4, column=0, sticky="w", padx=5, pady=5)
        y_scale = tk.Scale(
            controls_frame,
            from_=-y_offset_limit,
            to=y_offset_limit,
            orient="horizontal",
            resolution=1,
            variable=y_offset_var,
            command=lambda _: update_preview(),
            bg=PALETTE["card"],
            troughcolor=PALETTE["outline"],
            highlightthickness=0,
            sliderrelief="flat",
        )
        y_scale.grid(row=4, column=1, sticky="ew", padx=5, pady=5)

        def compute_positions() -> Tuple[str, int, float, float]:
            current_font = font_var.get()
            current_size = max(5, int(size_var.get()))
            text_width = stringWidth(name_text, current_font, current_size)
            base_x = self.rect_info["x_left"] + (self.rect_info["rect_width"] - text_width) / 2
            base_y = self.rect_info["y_bottom"] + (self.rect_info["rect_height"] - current_size) / 2
            adjusted_x = base_x + float(x_offset_var.get())
            adjusted_y = base_y + float(y_offset_var.get())
            return current_font, current_size, adjusted_x, adjusted_y

        def update_preview() -> None:
            try:
                current_font, current_size, adjusted_x, adjusted_y = compute_positions()
                working_image = self.img.copy()
                draw = ImageDraw.Draw(working_image)
                draw.rectangle(
                    [
                        int(self.rect_info["tk_left"]),
                        int(self.rect_info["tk_top"]),
                        int(self.rect_info["tk_right"]),
                        int(self.rect_info["tk_bottom"]),
                    ],
                    outline=PALETTE["accent"],
                    width=2,
                )

                font_path = self.available_fonts[current_font]
                pil_font = ImageFont.truetype(font_path, current_size)
                draw.text(
                    (int(adjusted_x), int(self.pix.height - adjusted_y)),
                    name_text,
                    font=pil_font,
                    fill=(255, 255, 255),
                    anchor="ls",
                )

                display_image = working_image
                if not use_original_size:
                    width = int(self.pix.width * preview_scale)
                    height = int(self.pix.height * preview_scale)
                    display_image = working_image.resize((width, height), Image.LANCZOS)

                preview = ImageTk.PhotoImage(display_image)
                preview_label.configure(image=preview)
                preview_label.image = preview
            except OSError as error:
                messagebox.showerror("Font Error", f"Unable to load font file.\n{error}", parent=window)

        def on_next() -> None:
            current_font, current_size, adjusted_x, adjusted_y = compute_positions()
            result["font"] = current_font
            result["size"] = current_size
            result["x"] = adjusted_x
            result["y"] = adjusted_y
            window.destroy()

        def on_cancel() -> None:
            result.clear()
            window.destroy()

        ttk.Button(button_frame, text="Save & Next", style="Accent.TButton", command=on_next).pack(
            side="left", padx=5
        )
        ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=5)

        controls_frame.columnconfigure(1, weight=1)

        window.bind("<Return>", lambda _: on_next())
        window.bind("<Escape>", lambda _: on_cancel())
        window.protocol("WM_DELETE_WINDOW", on_cancel)
        window.update_idletasks()
        update_preview()
        window.grab_set()
        self.root.wait_window(window)

        return result if result else None

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

        output_filename = f"certificate_{name_text.replace(' ', '_')}.pdf"
        with open(output_filename, "wb") as out_file:
            writer.write(out_file)

    def show_rectangle_selection(self, img: Image.Image, pix: fitz.Pixmap) -> Tuple[Optional[Tuple[int, int]], Optional[Tuple[int, int]]]:
        """Let the user draw a rectangle; return last drawn rectangle even if window is closed."""

        selection = {"start": None, "end": None, "coords": None}
        rect_id = None

        win = tk.Toplevel(self.root)
        win.title("Mark name placement area")
        win.configure(bg=PALETTE["card"])

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
            highlightbackground=PALETTE["outline"],
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
            rect_id = canvas_widget.create_rectangle(event.x, event.y, event.x, event.y, outline=PALETTE["accent"], width=2)
            print("[DEBUG] Rectangle mouse_down start=%s" % (selection["start"],))

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
            print("[DEBUG] Rectangle mouse_up end=%s coords=%s" % (selection["end"], selection["coords"]))

        def on_confirm():
            if rect_id and not selection["coords"]:
                selection["coords"] = canvas_widget.coords(rect_id)
            print("[DEBUG] Rectangle confirm coords=%s start=%s end=%s" % (selection["coords"], selection["start"], selection["end"]))
            win.destroy()

        def on_close():
            if rect_id and not selection["coords"]:
                selection["coords"] = canvas_widget.coords(rect_id)
            print("[DEBUG] Rectangle close coords=%s start=%s end=%s" % (selection["coords"], selection["start"], selection["end"]))
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

        # Prefer exact start/end, but fall back to canvas coords if needed
        if selection["start"] and selection["end"]:
            print("[DEBUG] Rectangle return start/end=%s %s" % (selection["start"], selection["end"]))
            return selection["start"], selection["end"]
        if selection["coords"]:
            x1, y1, x2, y2 = selection["coords"]
            print("[DEBUG] Rectangle return coords=%s" % (selection["coords"],))
            return (int(x1), int(y1)), (int(x2), int(y2))
        print("[DEBUG] Rectangle return None (no selection)")
        return None, None

    def update_ready_state(self) -> None:
        ready = bool(self.pdf_path and self.names and self.rect_info)
        if ready:
            self.start_button.state(["!disabled"])
            self.progress_status.set("Ready to review and export")
            print("[DEBUG] Ready state: ENABLED (pdf=%s names=%s rect=%s)" % (bool(self.pdf_path), bool(self.names), bool(self.rect_info)))
        else:
            missing = []
            if not self.pdf_path:
                missing.append("template PDF")
            if not self.names:
                missing.append("names file")
            if not self.rect_info:
                missing.append("placement area")
            self.start_button.state(["disabled"])
            self.progress_status.set(f"Waiting: select {', '.join(missing)}")
            print("[DEBUG] Ready state: DISABLED missing=%s" % missing)


if __name__ == "__main__":
    CertificateApp()
