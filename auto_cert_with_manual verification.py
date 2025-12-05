import io
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PyPDF2 import PdfReader, PdfWriter

# -------------------- Select Files --------------------
pdf_path = filedialog.askopenfilename(title="Select Certificate Template PDF", filetypes=[("PDF Files", "*.pdf")])
txt_path = filedialog.askopenfilename(title="Select Names TXT File", filetypes=[("Text Files", "*.txt")])

with open(txt_path, 'r') as f:
    names = [line.strip() for line in f.readlines() if line.strip()]

# -------------------- Register Fonts --------------------

# Automatically load all TTF/OTF fonts from the 'fonts' directory
fonts_dir = os.path.join(os.path.dirname(__file__), "fonts")
available_fonts = {}

if not os.path.exists(fonts_dir):
    os.makedirs(fonts_dir)

# Recursively search for font files in all subfolders
for root, dirs, files in os.walk(fonts_dir):
    family = os.path.basename(root)
    for font_file in files:
        if font_file.lower().endswith(('.ttf', '.otf')):
            font_path = os.path.join(root, font_file)
            font_label = f"{family}/{os.path.splitext(font_file)[0]}" if family != "fonts" else os.path.splitext(font_file)[0]
            try:
                pdfmetrics.registerFont(TTFont(font_label, font_path))
                available_fonts[font_label] = font_path
            except Exception as e:
                print(f"Warning: Could not register font '{font_label}' at {font_path}: {e}")

if not available_fonts:
    messagebox.showerror("Font Error", f"No fonts found in '{fonts_dir}' or its subfolders. Please add .ttf or .otf font files.")
    raise SystemExit(1)

current_font_choice = next(iter(available_fonts))
    
# -------------------- Draw Rectangle --------------------
doc = fitz.open(pdf_path)
page = doc.load_page(0)

pix = page.get_pixmap()
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
doc.close()

root = tk.Tk()
root.title("Draw rectangle for name area")

canvas_widget = tk.Canvas(root, width=pix.width, height=pix.height)
canvas_widget.pack()
tk_img = ImageTk.PhotoImage(img)
canvas_widget.create_image(0, 0, anchor="nw", image=tk_img)

rect_start = None
rect_end = None
rect_id = None

def on_mouse_down(event):
    global rect_start, rect_id
    rect_start = (event.x, event.y)
    rect_id = canvas_widget.create_rectangle(event.x, event.y, event.x, event.y, outline="red", width=2)

def on_mouse_move(event):
    global rect_id
    if rect_id:
        canvas_widget.coords(rect_id, rect_start[0], rect_start[1], event.x, event.y)

def on_mouse_up(event):
    global rect_end
    rect_end = (event.x, event.y)
    messagebox.showinfo("Rectangle Selected", f"Rectangle drawn from {rect_start} to {rect_end}\nClose window to generate certificates.")

canvas_widget.bind("<Button-1>", on_mouse_down)
canvas_widget.bind("<B1-Motion>", on_mouse_move)
canvas_widget.bind("<ButtonRelease-1>", on_mouse_up)
root.mainloop()
try:
    if root.winfo_exists():
        root.destroy()
except tk.TclError:
    pass

if not rect_start or not rect_end:
    print("No rectangle drawn. Exiting.")
    raise SystemExit(0)

# Convert Tkinter coords to PDF coords
x1, y1 = rect_start
x2, y2 = rect_end
pdf_x1, pdf_y1 = x1, pix.height - y1
pdf_x2, pdf_y2 = x2, pix.height - y2

rect_width = abs(pdf_x2 - pdf_x1)
rect_height = abs(pdf_y2 - pdf_y1)
x_left = min(pdf_x1, pdf_x2)
y_bottom = min(pdf_y1, pdf_y2)

tk_rect_left = min(x1, x2)
tk_rect_top = min(y1, y2)
tk_rect_right = max(x1, x2)
tk_rect_bottom = max(y1, y2)


def calculate_initial_font_size(text_value: str, selected_font: str) -> int:
    """Determine the largest font size that fits within the selected rectangle."""
    max_size = int(min(rect_height, 120)) or 5
    size = max_size
    while size >= 5:
        text_width = stringWidth(text_value, selected_font, size)
        if text_width <= rect_width and size <= rect_height:
            return size
        size -= 1
    return 5


def review_name(name_text: str, default_font: str, default_size: int):
    """Allow manual review of placement and font before exporting the certificate."""
    result = {}
    window = tk.Tk()
    window.title(f"Review Certificate: {name_text}")
    window.configure(padx=12, pady=12)
    window.columnconfigure(0, weight=1)
    window.columnconfigure(1, weight=0)
    window.rowconfigure(0, weight=1)

    preview_scale = min(1.0, 800 / max(1, pix.width), 550 / max(1, pix.height))
    display_width = int(pix.width * preview_scale)
    display_height = int(pix.height * preview_scale)

    use_original_size = abs(preview_scale - 1.0) < 1e-6
    preview_image = img if use_original_size else img.resize((display_width, display_height), Image.LANCZOS)
    preview_photo = ImageTk.PhotoImage(preview_image)

    preview_label = tk.Label(window, image=preview_photo, bd=2, relief="sunken")
    preview_label.grid(row=0, column=0, rowspan=6, sticky="nsew", padx=(0, 12))
    preview_label.image = preview_photo

    controls_frame = ttk.Frame(window)
    controls_frame.grid(row=0, column=1, sticky="n")
    controls_frame.columnconfigure(1, weight=1)

    button_frame = ttk.Frame(controls_frame)
    button_frame.grid(row=0, column=0, columnspan=2, sticky="ew", pady=(0, 8))

    font_var = tk.StringVar(value=default_font)
    size_var = tk.IntVar(value=max(5, default_size))
    x_offset_limit = int(max(10, rect_width / 2))
    y_offset_limit = int(max(10, rect_height / 2))
    x_offset_var = tk.DoubleVar(value=0)
    y_offset_var = tk.DoubleVar(value=0)

    ttk.Label(controls_frame, text="Font").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    font_selector = ttk.Combobox(controls_frame, textvariable=font_var, values=list(available_fonts.keys()), state="readonly")
    font_selector.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    font_selector.bind("<<ComboboxSelected>>", lambda _: update_preview())

    ttk.Label(controls_frame, text="Font Size").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    size_scale_max = max(int(rect_height), default_size + 40, 10)
    size_scale = tk.Scale(controls_frame, from_=5, to=size_scale_max, orient="horizontal", variable=size_var, command=lambda _: update_preview())
    size_scale.grid(row=2, column=1, sticky="ew", padx=5, pady=5)

    ttk.Label(controls_frame, text="Horizontal Offset").grid(row=3, column=0, sticky="w", padx=5, pady=5)
    x_scale = tk.Scale(controls_frame, from_=-x_offset_limit, to=x_offset_limit, orient="horizontal", resolution=1, variable=x_offset_var, command=lambda _: update_preview())
    x_scale.grid(row=3, column=1, sticky="ew", padx=5, pady=5)

    ttk.Label(controls_frame, text="Vertical Offset").grid(row=4, column=0, sticky="w", padx=5, pady=5)
    y_scale = tk.Scale(controls_frame, from_=-y_offset_limit, to=y_offset_limit, orient="horizontal", resolution=1, variable=y_offset_var, command=lambda _: update_preview())
    y_scale.grid(row=4, column=1, sticky="ew", padx=5, pady=5)

    def compute_positions():
        current_font = font_var.get()
        current_size = max(5, int(size_var.get()))
        text_width = stringWidth(name_text, current_font, current_size)
        base_x = x_left + (rect_width - text_width) / 2
        base_y = y_bottom + (rect_height - current_size) / 2
        adjusted_x = base_x + float(x_offset_var.get())
        adjusted_y = base_y + float(y_offset_var.get())
        return current_font, current_size, adjusted_x, adjusted_y

    def update_preview():
        try:
            current_font, current_size, adjusted_x, adjusted_y = compute_positions()
            working_image = img.copy()
            draw = ImageDraw.Draw(working_image)
            draw.rectangle([
                int(tk_rect_left),
                int(tk_rect_top),
                int(tk_rect_right),
                int(tk_rect_bottom),
            ], outline="red", width=2)
            font_path = available_fonts[current_font]
            pil_font = ImageFont.truetype(font_path, current_size)
            draw.text((int(adjusted_x), int(pix.height - adjusted_y)), name_text, font=pil_font, fill=(255, 255, 255), anchor="ls")

            display_image = working_image
            if not use_original_size:
                width = int(pix.width * preview_scale)
                height = int(pix.height * preview_scale)
                display_image = working_image.resize((width, height), Image.LANCZOS)

            preview = ImageTk.PhotoImage(display_image)
            preview_label.configure(image=preview)
            preview_label.image = preview
        except OSError as error:
            messagebox.showerror("Font Error", f"Unable to load font file.\n{error}", parent=window)

    def on_next():
        current_font, current_size, adjusted_x, adjusted_y = compute_positions()
        result["font"] = current_font
        result["size"] = current_size
        result["x"] = adjusted_x
        result["y"] = adjusted_y
        window.destroy()

    def on_cancel():
        result.clear()
        window.destroy()

    ttk.Button(button_frame, text="Save & Next", command=on_next).pack(side="left", padx=5)
    ttk.Button(button_frame, text="Cancel", command=on_cancel).pack(side="left", padx=5)

    controls_frame.columnconfigure(1, weight=1)

    window.bind("<Return>", lambda _: on_next())
    window.bind("<Escape>", lambda _: on_cancel())
    window.protocol("WM_DELETE_WINDOW", on_cancel)
    window.update_idletasks()
    update_preview()
    window.mainloop()

    return result if result else None


def generate_certificate(name_text: str, font_choice: str, font_size: int, pdf_x: float, pdf_y: float):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(pix.width, pix.height))
    c.setFont(font_choice, font_size)
    c.setFillColorRGB(1, 1, 1)
    c.drawString(pdf_x, pdf_y, name_text)
    c.save()

    packet.seek(0)
    template_reader = PdfReader(io.BytesIO(template_pdf_bytes))
    overlay_reader = PdfReader(packet)
    writer = PdfWriter()

    base_page = template_reader.pages[0]
    base_page.merge_page(overlay_reader.pages[0])
    writer.add_page(base_page)

    output_filename = f"certificate_{name_text.replace(' ', '_')}.pdf"
    with open(output_filename, "wb") as out_file:
        writer.write(out_file)
    print(f"Saved: {output_filename}")

# -------------------- Generate Certificates --------------------
with open(pdf_path, "rb") as template_file:
    template_pdf_bytes = template_file.read()

last_selected_font = current_font_choice
last_font_size = None

for name in names:
    initial_size = calculate_initial_font_size(name, last_selected_font)
    if last_font_size:
        initial_size = max(5, min(int(rect_height), last_font_size))

    review_result = review_name(name, last_selected_font, initial_size)

    if not review_result:
        print("Certificate generation cancelled by user.")
        break

    last_selected_font = review_result["font"]
    last_font_size = review_result["size"]

    generate_certificate(name, review_result["font"], review_result["size"], review_result["x"], review_result["y"])
else:
    print("All certificates generated with manual review!")
