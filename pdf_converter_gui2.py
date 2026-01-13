import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from pdf2image import convert_from_path
from PIL import Image, ImageEnhance

def pdf_to_images():
pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
if not pdf_file:
return

```
save_folder = filedialog.askdirectory(title="Select Folder to Save Images")
if not save_folder:
    return

img_format = simpledialog.askstring("Image Format", "Enter image format (png/jpg):", initialvalue="png")
if img_format.lower() not in ["png", "jpg", "jpeg"]:
    messagebox.showerror("Error", "Invalid image format!")
    return

dpi = simpledialog.askinteger("DPI", "Enter DPI (e.g., 300):", initialvalue=300)
upscale = simpledialog.askfloat("Upscale Factor", "Enter upscale factor (e.g., 1.5 or 2.0):", initialvalue=1.5)

# Ask for poppler path
poppler_path = filedialog.askdirectory(title="Select Poppler 'bin' folder")
if not poppler_path:
    messagebox.showerror("Error", "Poppler path is required to convert PDF to images!")
    return

try:
    # Ask user if they want full PDF or a page range
    page_option = simpledialog.askstring("Page Range", "Enter page range (e.g., 1-3) or leave empty for all pages:")

    pages = convert_from_path(pdf_file, dpi=dpi, poppler_path=poppler_path)

    # Filter pages if range provided
    if page_option:
        try:
            start_str, end_str = page_option.split("-")
            start_page = max(int(start_str), 1)
            end_page = min(int(end_str), len(pages))
            pages = pages[start_page-1:end_page]
        except Exception:
            messagebox.showerror("Error", "Invalid page range format. Using all pages.")

    progress_win = tk.Toplevel(root)
    progress_win.title("Converting PDF → Images")
    tk.Label(progress_win, text="Converting PDF pages to images...").pack(pady=10)
    progress = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
    progress.pack(pady=10)
    progress["maximum"] = len(pages)

    base_name = os.path.splitext(os.path.basename(pdf_file))[0]

    for i, page in enumerate(pages, start=1):
        width, height = page.size
        page = page.resize((int(width*upscale), int(height*upscale)), Image.LANCZOS)
        page = ImageEnhance.Sharpness(page).enhance(1.5)
        page = ImageEnhance.Contrast(page).enhance(1.2)
        page = ImageEnhance.Color(page).enhance(1.1)

        img_path = os.path.join(save_folder, f"{base_name}_page_{i}.{img_format.lower()}")
        page.save(img_path)
        progress["value"] = i
        progress_win.update_idletasks()

    progress_win.destroy()
    messagebox.showinfo("Success", f"Saved {len(pages)} images in {save_folder}")
except Exception as e:
    messagebox.showerror("Error", str(e))
```

# -------------------- GUI --------------------

root = tk.Tk()
root.title("📄 PDF → Image Converter")
root.geometry("450x220")
root.configure(bg="white")

tk.Label(root, text="PDF → Image Converter", font=("Arial", 16, "bold"), bg="white").pack(pady=20)
tk.Button(root, text="Select PDF and Convert", command=pdf_to_images, width=35, height=2).pack(pady=20)

root.mainloop()
