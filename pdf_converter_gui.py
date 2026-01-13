import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
from pdf2docx import Converter
import pdfplumber
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
from docx import Document


# -------------------- Conversion Functions --------------------

def convert_pdf_to_word(pdf_file):
    try:
        save_path = filedialog.asksaveasfilename(
            initialfile=os.path.splitext(os.path.basename(pdf_file))[0] + ".docx",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx")]
        )
        if not save_path:
            return "Conversion cancelled"

        cv = Converter(pdf_file)
        cv.convert(save_path)
        cv.close()
        return save_path

    except Exception as e:
        return f"Error converting to Word: {e}"


def convert_pdf_to_excel(pdf_file):
    try:
        save_path = filedialog.asksaveasfilename(
            initialfile=os.path.splitext(os.path.basename(pdf_file))[0] + ".xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return "Conversion cancelled"

        all_tables = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table[1:], columns=table[0])
                    all_tables.append(df)

        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df.to_excel(save_path, index=False)
            return save_path
        else:
            return "No tables found"

    except Exception as e:
        return f"Error converting to Excel: {e}"


# NEW: Word → Excel converter
def convert_word_to_excel(docx_file):
    try:
        save_path = filedialog.asksaveasfilename(
            initialfile=os.path.splitext(os.path.basename(docx_file))[0] + ".xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if not save_path:
            return "Conversion cancelled"

        document = Document(docx_file)
        all_tables = []

        for table in document.tables:
            data = []
            keys = None

            for i, row in enumerate(table.rows):
                text_row = [cell.text.strip() for cell in row.cells]
                if i == 0:
                    keys = text_row
                else:
                    data.append(text_row)

            if keys and data:
                df = pd.DataFrame(data, columns=keys)
                all_tables.append(df)

        if all_tables:
            final_df = pd.concat(all_tables, ignore_index=True)
            final_df.to_excel(save_path, index=False)
            return save_path
        else:
            return "No tables found in Word"

    except Exception as e:
        return f"Error converting Word to Excel: {e}"


# -------------------- PDF Split Functions --------------------

def split_pdf():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        return

    save_folder = filedialog.askdirectory(title="Select Folder to Save Pages")
    if not save_folder:
        return

    try:
        reader = PdfReader(pdf_file)
        total_pages = len(reader.pages)

        progress_win = tk.Toplevel(root)
        progress_win.title("Splitting PDF")
        tk.Label(progress_win, text="Splitting PDF...").pack(pady=10)
        progress = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10)
        progress["maximum"] = total_pages

        for i, page in enumerate(reader.pages, start=1):
            writer = PdfWriter()
            writer.add_page(page)
            page_path = os.path.join(save_folder, f"{os.path.splitext(os.path.basename(pdf_file))[0]}_page_{i}.pdf")
            with open(page_path, "wb") as f:
                writer.write(f)

            progress["value"] = i
            progress_win.update_idletasks()

        progress_win.destroy()
        messagebox.showinfo("Success", f"PDF split into {total_pages} pages.")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def split_pdf_by_number():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        return

    save_folder = filedialog.askdirectory(title="Select Folder to Save Split PDFs")
    if not save_folder:
        return

    try:
        reader = PdfReader(pdf_file)
        total_pages = len(reader.pages)

        num_pages_str = simpledialog.askstring("Pages per Split", "Enter number of pages per split:")
        if not num_pages_str or not num_pages_str.isdigit():
            messagebox.showwarning("Invalid Input", "Please enter a valid number.")
            return

        num_pages = int(num_pages_str)

        progress_win = tk.Toplevel(root)
        progress_win.title("Splitting PDF")
        tk.Label(progress_win, text="Splitting PDF...").pack(pady=10)
        progress = ttk.Progressbar(progress_win, orient="horizontal", length=300, mode="determinate")
        progress.pack(pady=10)

        total_splits = (total_pages + num_pages - 1) // num_pages
        progress["maximum"] = total_splits

        split_count = 0

        for start in range(0, total_pages, num_pages):
            writer = PdfWriter()
            for i in range(start, min(start + num_pages, total_pages)):
                writer.add_page(reader.pages[i])

            split_count += 1
            split_path = os.path.join(save_folder, f"{os.path.splitext(os.path.basename(pdf_file))[0]}_part_{split_count}.pdf")
            with open(split_path, "wb") as f:
                writer.write(f)

            progress["value"] = split_count
            progress_win.update_idletasks()

        progress_win.destroy()
        messagebox.showinfo("Success", f"PDF split into {split_count} files.")

    except Exception as e:
        messagebox.showerror("Error", str(e))


def split_pdf_by_range():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        return

    save_folder = filedialog.askdirectory(title="Select Folder to Save Split PDFs")
    if not save_folder:
        return

    try:
        reader = PdfReader(pdf_file)
        total_pages = len(reader.pages)

        start_page_str = simpledialog.askstring("Start Page", f"Enter start page (1-{total_pages}):")
        end_page_str = simpledialog.askstring("End Page", f"Enter end page (1-{total_pages}):")

        if not start_page_str.isdigit() or not end_page_str.isdigit():
            messagebox.showwarning("Invalid Input", "Enter valid page numbers.")
            return

        start_page = int(start_page_str)
        end_page = int(end_page_str)

        if start_page < 1 or end_page > total_pages or start_page > end_page:
            messagebox.showwarning("Invalid Range", "Invalid page range.")
            return

        writer = PdfWriter()
        for i in range(start_page - 1, end_page):
            writer.add_page(reader.pages[i])

        save_path = os.path.join(
            save_folder,
            f"{os.path.splitext(os.path.basename(pdf_file))[0]}_pages_{start_page}_to_{end_page}.pdf"
        )

        with open(save_path, "wb") as f:
            writer.write(f)

        messagebox.showinfo("Success", f"Saved: {save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# -------------------- Remove Pages --------------------

def remove_pages_from_pdf():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if not pdf_file:
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
    if not save_path:
        return

    try:
        reader = PdfReader(pdf_file)
        writer = PdfWriter()
        total_pages = len(reader.pages)

        remove_option = simpledialog.askstring(
            "Remove Pages",
            "Enter pages/ranges to remove (e.g., 2,5,7-9):"
        )

        remove_list = []
        parts = remove_option.split(",")

        for part in parts:
            if "-" in part:
                start, end = part.split("-")
                remove_list.extend(range(int(start) - 1, int(end)))
            else:
                remove_list.append(int(part) - 1)

        for i in range(total_pages):
            if i not in remove_list:
                writer.add_page(reader.pages[i])

        with open(save_path, "wb") as f:
            writer.write(f)

        messagebox.showinfo("Success", f"Saved: {save_path}")

    except Exception as e:
        messagebox.showerror("Error", str(e))


# -------------------- Merge PDFs --------------------

def merge_pdfs():
    pdf_files = filedialog.askopenfilenames(
        title="Select PDFs to Merge (Select at least 2)",
        filetypes=[("PDF Files", "*.pdf")]
    )

    # Validation: minimum 2 PDFs required
    if not pdf_files or len(pdf_files) < 2:
        messagebox.showwarning(
            "Merge PDFs",
            "Please select at least TWO PDF files to merge."
        )
        return

    # Ask user for output file name (clear message)
    output_name = simpledialog.askstring(
        "Save Merged PDF",
        "Enter name for the merged PDF file (without .pdf):"
    )

    if not output_name:
        messagebox.showwarning(
            "Merge PDFs",
            "No file name provided. Merge cancelled."
        )
        return

    save_path = filedialog.asksaveasfilename(
        initialfile=f"{output_name}.pdf",
        defaultextension=".pdf",
        filetypes=[("PDF Files", "*.pdf")]
    )

    if not save_path:
        return

    try:
        writer = PdfWriter()

        # Calculate total pages for progress bar
        total_pages = sum(len(PdfReader(pdf).pages) for pdf in pdf_files)

        progress_win = tk.Toplevel(root)
        progress_win.title("Merging PDFs")
        tk.Label(
            progress_win,
            text="Merging selected PDF files...\nPlease wait.",
            font=("Arial", 10)
        ).pack(pady=10)

        progress = ttk.Progressbar(
            progress_win,
            orient="horizontal",
            length=320,
            mode="determinate"
        )
        progress.pack(pady=10)
        progress["maximum"] = total_pages

        count = 0

        # Merge PDFs in selected order
        for pdf in pdf_files:
            reader = PdfReader(pdf)
            for page in reader.pages:
                writer.add_page(page)
                count += 1
                progress["value"] = count
                progress_win.update_idletasks()

        with open(save_path, "wb") as f:
            writer.write(f)

        progress_win.destroy()

        messagebox.showinfo(
            "Merge Successful",
            f"PDFs merged successfully!\n\nSaved as:\n{save_path}"
        )

    except Exception as e:
        messagebox.showerror("Merge Error", str(e))

# -------------------- GUI --------------------

root = tk.Tk()
root.title("PDF Converter & Tools")
root.geometry("450x700")
root.configure(bg="white")  # FIX: no transparent

tk.Label(root, text="📄 PDF Converter & Tool", font=("Arial", 18, "bold"), bg="white").pack(pady=15)

tk.Button(root, text="PDF → Word", command=lambda: run_pdf_to_word(), width=35, height=2).pack(pady=5)
tk.Button(root, text="PDF → Excel", command=lambda: run_pdf_to_excel(), width=35, height=2).pack(pady=5)
tk.Button(root, text="Word → Excel", command=lambda: run_word_to_excel(), width=35, height=2).pack(pady=5)

tk.Button(root, text="Split PDF (Each Page)", command=split_pdf, width=35, height=2).pack(pady=5)
tk.Button(root, text="Split PDF by Number", command=split_pdf_by_number, width=35, height=2).pack(pady=5)
tk.Button(root, text="Split PDF by Range", command=split_pdf_by_range, width=35, height=2).pack(pady=5)

tk.Button(root, text="Remove Pages", command=remove_pages_from_pdf, width=35, height=2).pack(pady=5)
tk.Button(root, text="Merge PDFs", command=merge_pdfs, width=35, height=2).pack(pady=5)


def run_pdf_to_word():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_file:
        result = convert_pdf_to_word(pdf_file)
        messagebox.showinfo("Result", result)


def run_pdf_to_excel():
    pdf_file = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
    if pdf_file:
        result = convert_pdf_to_excel(pdf_file)
        messagebox.showinfo("Result", result)


def run_word_to_excel():
    docx_file = filedialog.askopenfilename(filetypes=[("Word Files", "*.docx")])
    if docx_file:
        result = convert_word_to_excel(docx_file)
        messagebox.showinfo("Result", result)


root.mainloop()
