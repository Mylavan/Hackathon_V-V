from Main_Helper import Main_Helper_Entry 
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading

import platform
import subprocess

class PDFLoaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Folder Analyzer")
        self.root.geometry("600x450")

        self.selected_folder_path = tk.StringVar()

        # Folder selection
        tk.Label(root, text="Selected Folder:").pack(pady=5)
        self.folder_entry = tk.Entry(root, textvariable=self.selected_folder_path, width=70, state='readonly')
        self.folder_entry.pack(pady=5)
        tk.Button(root, text="Browse Folder", command=self.browse_folder).pack(pady=5)

        # PDF list
        self.list_label = tk.Label(root, text="PDF Files in Folder:")
        self.list_label.pack(pady=5)
        self.pdf_listbox = tk.Listbox(root, width=80, height=10)
        self.pdf_listbox.pack(pady=5)

        # Process button
        self.process_button = tk.Button(root, text="Process PDFs", command=self.start_pdf_processing)
        self.process_button.pack(pady=10)

        # Progress bar and label
        self.progress_label = tk.Label(root, text="")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(root, orient="horizontal", length=400, mode="indeterminate")

        self.pdf_files = []

    def browse_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder Containing PDFs")
        if folder_path:
            self.selected_folder_path.set(folder_path)
            self.load_pdfs(folder_path)

    def load_pdfs(self, folder):
        self.pdf_files = [f for f in os.listdir(folder) if f.lower().endswith(".pdf")]
        self.pdf_listbox.delete(0, tk.END)

        if not self.pdf_files:
            self.pdf_listbox.insert(tk.END, "No PDF files found in selected folder.")
        else:
            for pdf in self.pdf_files:
                self.pdf_listbox.insert(tk.END, pdf)

    def start_pdf_processing(self):
        if not self.pdf_files:
            messagebox.showwarning("No PDFs", "No PDF files found to process.")
            return

        folder = self.selected_folder_path.get()
        self.progress_label.config(text="Analyzing all PDFs...")
        self.progress_bar.pack(pady=5)
        self.progress_bar.start(10)

        threading.Thread(target=self.analyze_all_pdfs, args=(folder,), daemon=True).start()

    def analyze_all_pdfs(self, folder):
         
        self.process_pdf_file( folder)

        
    def is_excel_open(self, file_path):
        """Check if Excel file is open (locked)."""
        if not os.path.exists(file_path):
            return False  # file not present = not open
        
        try:
            with open(file_path, "r+b"):
                return False  # file is not locked
        except PermissionError:
            return True  # file is open/locked
         
    def process_pdf_file(self, folder_path):
        print(f"From folder: {folder_path}")
        Excel_file_path = os.path.join(folder_path, "Quick_Summary.xlsx")
        print(f"Excel file path: {Excel_file_path}")
        if self.is_excel_open(Excel_file_path):
            messagebox.showerror(
                "File in Use",
                f"The Excel file is currently open:\n{folder_path}\n\nPlease close it and try again."
            )
            self.root.after(1000, self.root.destroy)
            return  # Stop processing if file is open

        else:
            Main_Helper_Entry(folder_path)  # your main processing function

        self.root.after(0, self.analysis_complete)


    def analysis_complete(self):
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.config(text="All PDFs analyzed successfully!")
        messagebox.showinfo("Done", "All PDFs in the folder have been analyzed.")
        self.open_excel_in_folder(self.selected_folder_path.get())
        # Automatically close the UI after a short delay (e.g., 1 second)
        self.root.after(1000, self.root.destroy)

    def open_excel_in_folder(self, folder_path):
        """Find and open the first Excel file in the folder."""
        excel_files = [f for f in os.listdir(folder_path) if f.lower().endswith(".xlsx")]
        if not excel_files:
            messagebox.showinfo("No Excel File", "No Excel (.xlsx) file found in the folder.")
            return

        excel_path = os.path.join(folder_path, excel_files[0])
        try:
            if platform.system() == "Windows":
                os.startfile(excel_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.call(["open", excel_path])
            else:  # Linux and others
                subprocess.call(["xdg-open", excel_path])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open Excel file:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFLoaderApp(root)
    root.mainloop()
