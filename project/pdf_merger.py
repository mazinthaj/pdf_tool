import PyPDF2
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
from PIL import Image, ImageTk
import re
import os
import fitz  # PyMuPDF
import cv2
import numpy as np

class PDFMergerApp:
    def __init__(self):
        self.downloads_dir = os.path.join(os.path.expanduser('~'), 'Downloads')
        
        self.files = []
        self.page_deletion_formula = None

        self.window = tk.Tk()
        self.window.title("PDF Merger")
        self.window.geometry('500x800')
        self.window.configure(bg="#ffffff")

        self.header_frame = tk.Frame(self.window, bg="#ffcc00", height=100)
        self.header_frame.pack(fill="x")
        
        self.header_image = Image.open(r"dhl_log.png")
        self.header_image = self.header_image.resize((152, 26), Image.Resampling.LANCZOS)
        self.header_image = ImageTk.PhotoImage(self.header_image)
        
        self.header_image_label = tk.Label(self.header_frame, image=self.header_image, bg="#ffcc00")
        self.header_image_label.pack(pady=10)

        self.title_label = tk.Label(self.window, text="PDF Merger Tool", font=("Arial", 16, "bold"), fg="#666666", bg="#ffffff")
        self.title_label.pack(pady=10)

        background_frame = tk.Frame(self.window, bg="#ffffff")
        background_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.add_file_button = tk.Button(background_frame, bg="#d40511", fg="#ffffff", text="Add PDF File", command=self.choose_pdf_file, relief="flat", borderwidth=0, font=("Arial", 10, "bold"), padx=10, pady=5, cursor="hand2", width=20)
        self.add_file_button.pack(pady=5)
        self.add_button_hover_effect(self.add_file_button)

        self.file_list_label = tk.Label(background_frame, bg="#ffffff", text="Selected PDF files:", pady=10, font=("Arial", 10, "bold"), fg="#666666")
        self.file_list_label.pack()

        self.file_listbox = tk.Listbox(background_frame, selectmode=tk.SINGLE, width=80, height=10)
        self.file_listbox.pack(pady=5)

        self.page_deletion_label = tk.Label(background_frame, bg="#ffffff", text="Page Deletion Formula:\n\nEXAMPLE: [1: 2, 3, 5-8][2: 3, 5-20]", pady=5, font=("Arial", 10, "bold"), fg="#666666")
        self.page_deletion_label.pack()

        self.page_deletion_entry = tk.Entry(background_frame, width=80, font=("Arial", 10))
        self.page_deletion_entry.pack(pady=5)

        self.remove_blank_pages_var = tk.BooleanVar()
        self.remove_blank_pages_checkbox = tk.Checkbutton(
            background_frame, 
            text="Remove Blank Pages", 
            variable=self.remove_blank_pages_var, 
            bg="#ffffff", 
            font=("Arial", 10)
        )
        self.remove_blank_pages_checkbox.pack(pady=5)

        self.move_up_button = tk.Button(background_frame, bg="#d40511", fg="#ffffff", text="⬆️    Move Up", command=self.move_up, relief="flat", borderwidth=0, font=("Arial", 10, "bold"), padx=10, pady=5, cursor="hand2", width=20)
        self.move_up_button.pack(pady=5)
        self.add_button_hover_effect(self.move_up_button)

        self.move_down_button = tk.Button(background_frame, bg="#d40511", fg="#ffffff", text="⬇️Move Down", command=self.move_down, relief="flat", borderwidth=0, font=("Arial", 10, "bold"), padx=10, pady=5, cursor="hand2", width=20)
        self.move_down_button.pack(pady=5)
        self.add_button_hover_effect(self.move_down_button)

        self.delete_button = tk.Button(background_frame, bg="#d40511", fg="#ffffff", text="🗑️        Delete", command=self.delete, relief="flat", borderwidth=0, font=("Arial", 10, "bold"), padx=10, pady=5, cursor="hand2", width=20)
        self.delete_button.pack(pady=5)
        self.add_button_hover_effect(self.delete_button)

        self.process_button = tk.Button(background_frame, bg="#d40511", fg="#ffffff", text="Process Files", font=("Arial", 12, "bold"), padx=10, pady=10, command=self.process_files, relief="flat", borderwidth=0, cursor="hand2", width=20)
        self.process_button.pack(pady=30)
        self.add_button_hover_effect(self.process_button)

        self.window.mainloop()

    def add_button_hover_effect(self, button):
        """ Adds hover effect to buttons. """
        def on_enter(event):
            button.config(bg="#b2050f")

        def on_leave(event):
            button.config(bg="#d40511")

        button.bind("<Enter>", on_enter)
        button.bind("<Leave>", on_leave)

    def parse_page_deletion_formula(self, formula):
        """
        Parse the page deletion formula with new flexible format.
        
        Formula format: 
        - [1: 1, 2, 5-6, 8, 10, 20-25]
        - [1: 2, 3, 5-8][2: 3, 5-20]
        
        Returns a dictionary of page deletions or raises a ValueError
        """
        if not formula:
            return {}

        # Remove whitespaces
        formula = formula.replace(' ', '')
        
        # Regex to match PDF index and its deletion pages
        pdf_pattern = r'\[(\d+):((?:\d+(?:-\d+)?(?:,\s*)?)+)\]'
        pdf_matches = re.findall(pdf_pattern, formula)
        
        if not pdf_matches or formula != ''.join([f'[{m[0]}:{m[1]}]' for m in pdf_matches]):
            raise ValueError("Invalid formula format")
        
        # Create a dictionary to track page deletions
        page_deletions = {}
        used_pages = {}
        
        for pdf_index_str, deletion_pages in pdf_matches:
            pdf_index = int(pdf_index_str) - 1  # Convert to 0-based index
            
            if pdf_index not in used_pages:
                used_pages[pdf_index] = set()
            
            deletion_set = set()

            for page_spec in deletion_pages.split(','):
                if '-' in page_spec:
                    start, end = map(int, page_spec.split('-'))
                    
                    # Validate page range
                    if start > end:
                        raise ValueError(f"Invalid page range: {page_spec}")
                    
                    # Add full range (inclusive)
                    range_pages = set(range(start, end + 1))
                    deletion_set.update(range_pages)
                else:
                    deletion_set.add(int(page_spec))
            
            # Check for overlapping page deletions
            if deletion_set.intersection(used_pages[pdf_index]):
                raise ValueError("Overlapping page deletions")
            
            used_pages[pdf_index].update(deletion_set)
            
            # Store page deletions (convert to 0-based index for PyPDF2)
            page_deletions[pdf_index] = [page - 1 for page in deletion_set]
        
        return page_deletions

    def is_blank_page(self, page):
        """
        Check if a PDF page is blank using image processing.
        
        :param page: PyMuPDF page object
        :return: Boolean indicating if the page is blank
        """
        # Convert PDF page to an image
        pix = page.get_pixmap()
        img = np.frombuffer(pix.samples, dtype=np.uint8).reshape(pix.height, pix.width, pix.n)
        
        # Convert to grayscale
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        
        # Apply a threshold to get a binary image
        _, thresh = cv2.threshold(gray, 240, 255, cv2.THRESH_BINARY)
        
        # Check the percentage of white pixels
        white_pixels = np.sum(thresh == 255)
        total_pixels = thresh.size
        white_ratio = white_pixels / total_pixels
        
        return white_ratio > 0.97

    def merge_pdfs(self, pdf_list, output_path):
        """ 
        Merge PDFs with optional page deletion and blank page removal 
        """
        # Get page deletion formula
        page_deletions = self.parse_page_deletion_formula(
            self.page_deletion_entry.get()
        )

        # Open the output document
        output_doc = fitz.open()

        for index, pdf_path in enumerate(pdf_list):
            # Open the PDF
            pdf_reader = PyPDF2.PdfReader(open(pdf_path, 'rb'))
            
            for page_num, page in enumerate(pdf_reader.pages):
                # Skip pages if specified in deletion formula
                if index in page_deletions and page_num in page_deletions[index]:
                    continue
                
                # Convert PyPDF2 page to PyMuPDF
                temp_pdf = fitz.open()
                temp_pdf.insert_pdf(fitz.open(pdf_path), from_page=page_num, to_page=page_num)
                
                # If blank page removal is on, check the page
                if not self.remove_blank_pages_var.get() or not self.is_blank_page(temp_pdf[0]):
                    output_doc.insert_pdf(temp_pdf, from_page=0, to_page=0)
                
                temp_pdf.close()

        # Save the merged PDF
        output_doc.save(output_path)
        output_doc.close()

    def choose_pdf_file(self):
        file_paths = filedialog.askopenfilenames(
            title="Select PDF files",
            filetypes=[("PDF files", "*.pdf")]
        )
        if file_paths:
            for f in file_paths:
                self.files.append(f)
                self.file_listbox.insert(tk.END, f)

    def move_up(self):
        selected_index = self.file_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        if index == 0:
            return
        self.files[index], self.files[index - 1] = self.files[index - 1], self.files[index]

        self.update_file_listbox()
        self.file_listbox.select_set(index - 1)

    def move_down(self):
        selected_index = self.file_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        if index == len(self.files) - 1:
            return
        self.files[index], self.files[index + 1] = self.files[index + 1], self.files[index]
        self.update_file_listbox()
        self.file_listbox.select_set(index + 1)

    def delete(self):
        selected_index = self.file_listbox.curselection()
        if not selected_index:
            return
        index = selected_index[0]
        self.files.pop(index)
        self.update_file_listbox()

    def update_file_listbox(self):
        self.file_listbox.delete(0, tk.END)
        for file in self.files:
            self.file_listbox.insert(tk.END, file)

    def process_files(self):
        if not self.files:
            messagebox.showwarning("Warning", "Please select at least one PDF file before processing.")
            return

        try:
            if self.page_deletion_entry.get():
                self.parse_page_deletion_formula(self.page_deletion_entry.get())
        except ValueError as e:
            messagebox.showerror("Invalid Formula", str(e))
            return

        output_file_name = simpledialog.askstring("Output File Name", "Enter the name for the merged PDF file:")
        if not output_file_name:
            messagebox.showwarning("Warning", "You must provide a name for the output file.")
            return

        output_file_path = os.path.join(self.downloads_dir, f"{output_file_name}.pdf")
        
        if os.path.exists(output_file_path):
            if not messagebox.askyesno("File Exists", f"The file {output_file_name}.pdf already exists in Downloads. Do you want to replace it?"):
                return

        self.merge_pdfs(self.files, output_file_path)
        messagebox.showinfo("Success", f"Merged PDF saved in Downloads as {output_file_name}.pdf")

if __name__ == "__main__":
    app = PDFMergerApp()
