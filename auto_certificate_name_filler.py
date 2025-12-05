import io
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from PyPDF2 import PdfReader, PdfWriter
import re


# -------------------- Select Files --------------------
pdf_path = filedialog.askopenfilename(title="Select Certificate Template PDF", filetypes=[("PDF Files", "*.pdf")])
txt_path = filedialog.askopenfilename(title="Select Names TXT File", filetypes=[("Text Files", "*.txt")])

with open(txt_path, 'r') as f:
    names = [line.strip() for line in f.readlines() if line.strip()]

def sanitize_filename(name):
    # Replace invalid filename characters with underscore
    return re.sub(r'[\/\\\:\*\?\"\<\>\|]', '_', name)

# -------------------- Register Fonts --------------------
# Register Arial Bold
pdfmetrics.registerFont(TTFont('Arial-Bold', r'C:\Users\moham\Downloads\Compressed\arial\ARIALBD.TTF'))

# Use it for names
font_name = 'Arial-Bold'
    
# -------------------- Draw Rectangle --------------------
doc = fitz.open(pdf_path)
page = doc.load_page(0)

pix = page.get_pixmap()
img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)

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

if not rect_start or not rect_end:
    print("No rectangle drawn. Exiting.")
    exit()

# Convert Tkinter coords to PDF coords
x1, y1 = rect_start
x2, y2 = rect_end
pdf_x1, pdf_y1 = x1, pix.height - y1
pdf_x2, pdf_y2 = x2, pix.height - y2

rect_width = abs(pdf_x2 - pdf_x1)
rect_height = abs(pdf_y2 - pdf_y1)
x_left = min(pdf_x1, pdf_x2)
y_bottom = min(pdf_y1, pdf_y2)

# -------------------- Generate Certificates --------------------
for name in names:
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(pix.width, pix.height))

    # Start with a large font and shrink until it fits
    font_size = 100
    while True:
        text_width = stringWidth(name, font_name, font_size)
        if text_width <= rect_width and font_size <= rect_height:
            break
        font_size -= 1
        if font_size < 5:
            break  # minimum font

    c.setFont(font_name, font_size)
    c.setFillColorRGB(1, 1, 1)  # White color

    # Center text inside rectangle
    x_text = x_left + (rect_width - text_width)/2
    y_text = y_bottom + (rect_height - font_size)/2
    c.drawString(x_text, y_text, name)
    c.save()

    packet.seek(0)

    template = PdfReader(open(pdf_path, "rb"))
    overlay = PdfReader(packet)
    writer = PdfWriter()

    page = template.pages[0]
    page.merge_page(overlay.pages[0])
    writer.add_page(page)

    output_filename = f"certificate_{sanitize_filename(name)}.pdf"
    with open(output_filename, "wb") as out_file:
        writer.write(out_file)
    print(f"Saved: {output_filename}")

print("All certificates generated with dynamic font inside rectangle!")
