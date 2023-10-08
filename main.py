import fitz  
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox
import PyPDF2
import os
class PDFManipulatorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Manipulator")
        self.root.geometry("400x550")

        self.selected_files = []

        self.create_gui()

    def create_gui(self):
        ttk.Label(self.root, text="PDF Manipulator", font=("Helvetica", 16)).pack(pady=10)

        ttk.Button(self.root, text="Select PDFs to Merge", command=self.select_merge_pdfs).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Split", command=self.select_split_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Encrypt", command=self.select_encrypt_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Decrypt", command=self.select_decrypt_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Watermark", command=self.select_watermark_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Rotate Pages", command=self.select_rotate_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Delete Pages", command=self.select_delete_pages_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Reorder Pages", command=self.select_reorder_pages_pdf).pack(pady=10)
        ttk.Button(self.root, text="Select PDF to Extract Pages", command=self.select_extract_pages_pdf).pack(pady=10)

        self.merge_button = ttk.Button(self.root, text="Merge PDFs", state=tk.DISABLED, command=self.merge_pdfs)
        self.merge_button.pack(pady=10)

        self.split_button = ttk.Button(self.root, text="Split PDF", state=tk.DISABLED, command=self.split_pdf)
        self.split_button.pack(pady=10)

        self.encrypt_button = ttk.Button(self.root, text="Encrypt PDF", state=tk.DISABLED, command=self.encrypt_pdf)
        self.encrypt_button.pack(pady=10)

        self.decrypt_button = ttk.Button(self.root, text="Decrypt PDF", state=tk.DISABLED, command=self.decrypt_pdf)
        self.decrypt_button.pack(pady=10)

        self.watermark_button = ttk.Button(self.root, text="Watermark PDF", state=tk.DISABLED, command=self.watermark_pdf)
        self.watermark_button.pack(pady=10)

        self.rotate_button = ttk.Button(self.root, text="Rotate Pages", state=tk.DISABLED, command=self.rotate_pages)
        self.rotate_button.pack(pady=10)

        self.delete_pages_button = ttk.Button(self.root, text="Delete Pages", state=tk.DISABLED, command=self.delete_pages)
        self.delete_pages_button.pack(pady=10)

        self.reorder_pages_button = ttk.Button(self.root, text="Reorder Pages", state=tk.DISABLED, command=self.reorder_pages)
        self.reorder_pages_button.pack(pady=10)

        self.extract_pages_button = ttk.Button(self.root, text="Extract Pages", state=tk.DISABLED, command=self.extract_pages)
        self.extract_pages_button.pack(pady=10)

        self.preview_button = ttk.Button(self.root, text="Preview PDF", state=tk.DISABLED, command=self.preview_pdf)
        self.preview_button.pack(pady=10)

        self.status_label = ttk.Label(self.root, text="")
        self.status_label.pack(pady=10)
      def merge_pdfs(output_name, location, pdf1, pdf2):
    paths = [pdf1, pdf2]
    location = os.path.join(location, output_name)
    
    pdf_writer = fitz.open()
    
    for path in paths:
        pdf_reader = fitz.open(path)
        pdf_writer.insert_pdf(pdf_reader)
    
    pdf_writer.save(location)
    pdf_writer.close()
    
    return location

def create_split(pdf, output_name, location, pages):
    location = os.path.join(location, output_name)
    pdf_reader = fitz.open(pdf)
    pdf_writer = fitz.open()
    
    if ',' in pages:
        page_numbers = [int(page) for page in pages.split(',')]
    elif '-' in pages:
        start, end = map(int, pages.split('-'))
        page_numbers = range(start, end + 1)
    else:
        page_numbers = [int(pages)]
    
    for page_number in page_numbers:
        if 0 < page_number <= len(pdf_reader):
            pdf_writer.insert_pdf(pdf_reader, from_page=page_number - 1, to_page=page_number - 1)
    
    pdf_writer.save(location)
    pdf_writer.close()
    
    return location
def select_merge_pdfs(self):
        self.selected_files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        if len(self.selected_files) >= 2:
            self.merge_button["state"] = tk.NORMAL
        else:
            self.merge_button["state"] = tk.DISABLED

    def merge_pdfs(self):
        output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if output_path:
            pdf_writer = fitz.open()
            for pdf_file in self.selected_files:
                pdf_reader = fitz.open(pdf_file)
                pdf_writer.insert_pdf(pdf_reader)
                pdf_reader.close()
            pdf_writer.save(output_path)
            pdf_writer.close()
            messagebox.showinfo("Info", "PDFs merged successfully!")
            self.status_label.config(text=f"Merged PDF saved to: {output_path}")

    def select_split_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.split_button["state"] = tk.NORMAL
            self.preview_button["state"] = tk.NORMAL

    def split_pdf(self):
        page_range = filedialog.askstring("Page Range", "Enter page range (e.g., 2,4,6-8):")
        if page_range:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_reader = fitz.open(self.selected_files)
                pdf_writer = fitz.open()
                page_numbers = self.parse_page_range(page_range, len(pdf_reader))
                for page_number in page_numbers:
                    pdf_writer.insert_pdf(pdf_reader, from_page=page_number - 1, to_page=page_number - 1)
                pdf_writer.save(output_path)
                pdf_writer.close()
                messagebox.showinfo("Info", "PDF split successfully!")
                self.status_label.config(text=f"Split PDF saved to: {output_path}")

    def preview_pdf(self):
        try:
            os.startfile(self.selected_files)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to preview PDF: {e}")

    def parse_page_range(self, page_range, max_pages):
        pages = []
        for part in page_range.split(","):
            if "-" in part:
                start, end = map(int, part.split("-"))
                if start <= end and 1 <= start <= max_pages and 1 <= end <= max_pages:
                    pages.extend(range(start, end + 1))
            else:
                page = int(part)
                if 1 <= page <= max_pages:
                    pages.append(page)
        return pages
 def select_encrypt_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.encrypt_button["state"] = tk.NORMAL

    def encrypt_pdf(self):
        password = filedialog.askstring("Password", "Enter password for encryption:")
        if password:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_writer = PyPDF2.PdfFileWriter()
                pdf_reader = PyPDF2.PdfFileReader(self.selected_files)
                for page_num in range(pdf_reader.getNumPages()):
                    page = pdf_reader.getPage(page_num)
                    pdf_writer.addPage(page)
                pdf_writer.encrypt(password)
                with open(output_path, "wb") as encrypted_pdf:
                    pdf_writer.write(encrypted_pdf)
                messagebox.showinfo("Info", "PDF encrypted successfully!")
                self.status_label.config(text=f"Encrypted PDF saved to: {output_path}")

    def select_decrypt_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.decrypt_button["state"] = tk.NORMAL

    def decrypt_pdf(self):
        password = filedialog.askstring("Password", "Enter password for decryption:")
        if password:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_reader = PyPDF2.PdfFileReader(open(self.selected_files, "rb"))
                if pdf_reader.decrypt(password):
                    pdf_writer = PyPDF2.PdfFileWriter()
                    for page_num in range(pdf_reader.getNumPages()):
                        page = pdf_reader.getPage(page_num)
                        pdf_writer.addPage(page)
                    with open(output_path, "wb") as decrypted_pdf:
                        pdf_writer.write(decrypted_pdf)
                    messagebox.showinfo("Info", "PDF decrypted successfully!")
                    self.status_label.config(text=f"Decrypted PDF saved to: {output_path}")
                else:
                    messagebox.showerror("Error", "Password incorrect. PDF decryption failed.")

    def select_watermark_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.watermark_button["state"] = tk.NORMAL

    def watermark_pdf(self):
        watermark_text = filedialog.askstring("Watermark Text", "Enter watermark text:")
        if watermark_text:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_reader = PyPDF2.PdfFileReader(open(self.selected_files, "rb"))
                pdf_writer = PyPDF2.PdfFileWriter()
                for page_num in range(pdf_reader.getNumPages()):
                    page = pdf_reader.getPage(page_num)
                    page.mergePage(PyPDF2.pdf.PageObject.createTextString(watermark_text))
                    pdf_writer.addPage(page)
                with open(output_path, "wb") as watermarked_pdf:
                    pdf_writer.write(watermarked_pdf)
                messagebox.showinfo("Info", "PDF watermarked successfully!")
                self.status_label.config(text=f"Watermarked PDF saved to: {output_path}")

    def extract_text(self):
        text_output_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if text_output_path:
            try:
                pdf_document = fitz.open(self.selected_files[0])
                extracted_text = ""
                for page_number in range(len(pdf_document)):
                    page = pdf_document.load_page(page_number)
                    extracted_text += page.get_text()
                with open(text_output_path, "w", encoding="utf-8") as text_file:
                    text_file.write(extracted_text)
                messagebox.showinfo("Info", "Text extracted successfully!")
                self.status_label.config(text=f"Extracted text saved to: {text_output_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to extract text: {e}")

    def rotate_pages(self):
        rotation_degrees = filedialog.askinteger("Rotate Pages", "Enter rotation angle (90, 180, or 270 degrees):")
        if rotation_degrees in [90, 180, 270]:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_document = fitz.open(self.selected_files[0])
                pdf_document_rotated = fitz.open()
                for page_number in range(len(pdf_document)):
                    page = pdf_document.load_page(page_number)
                    page.setRotation(rotation_degrees)
                    pdf_document_rotated.insert_pdf(pdf_document, from_page=page_number, to_page=page_number)
                pdf_document_rotated.save(output_path)
                pdf_document_rotated.close()
                messagebox.showinfo("Info", "Pages rotated successfully!")
                self.status_label.config(text=f"Rotated PDF saved to: {output_path}")
               def select_delete_pages_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.delete_pages_button["state"] = tk.NORMAL

    def delete_pages(self):
        page_range = filedialog.askstring("Delete Pages", "Enter page range to delete (e.g., 2,4,6-8):")
        if page_range:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_reader = fitz.open(self.selected_files)
                pdf_writer = fitz.open()
                page_numbers = self.parse_page_range(page_range, len(pdf_reader))
                for page_number in range(len(pdf_reader)):
                    if page_number + 1 not in page_numbers:
                        pdf_writer.insert_pdf(pdf_reader, from_page=page_number, to_page=page_number)
                pdf_writer.save(output_path)
                pdf_writer.close()
                messagebox.showinfo("Info", "Pages deleted successfully!")
                self.status_label.config(text=f"Deleted Pages PDF saved to: {output_path}")
 def reorder_pages(self):
        page_order = filedialog.askstring("Reorder Pages", "Enter page order (e.g., 2,1,3,4):")
        if page_order:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_document = fitz.open(self.selected_files[0])
                pdf_document_reordered = fitz.open()
                page_order_list = self.parse_page_range(page_order, len(pdf_document))
                for page_number in page_order_list:
                    pdf_document_reordered.insert_pdf(pdf_document, from_page=page_number - 1, to_page=page_number - 1)
                pdf_document_reordered.save(output_path)
                pdf_document_reordered.close()
                messagebox.showinfo("Info", "Pages reordered successfully!")
                self.status_label.config(text=f"Reordered PDF saved to: {output_path}")

    def select_extract_pages_pdf(self):
        self.selected_files = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.selected_files:
            self.extract_pages_button["state"] = tk.NORMAL

    def extract_pages(self):
        page_range = filedialog.askstring("Extract Pages", "Enter page range to extract (e.g., 2,4,6-8):")
        if page_range:
            output_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
            if output_path:
                pdf_reader = fitz.open(self.selected_files)
                pdf_writer = fitz.open()
                page_numbers = self.parse_page_range(page_range, len(pdf_reader))
                for page_number in page_numbers:
                    pdf_writer.insert_pdf(pdf_reader, from_page=page_number - 1, to_page=page_number - 1)
                pdf_writer.save(output_path)
                pdf_writer.close()
                messagebox.showinfo("Info", "Pages extracted successfully!")
                self.status_label.config(text=f"Extracted Pages PDF saved to: {output_path}")
              if __name__ == "__main__":
    root = tk.Tk()
    app = PDFManipulatorApp(root)
    root.mainloop()
