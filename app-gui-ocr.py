# pdf_to_text_gui.py
"""
PDF to Text Converter - GUI Version menggunakan tkinter
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
import os
from pathlib import Path
import sys

# Import PDF libraries
try:
    import PyPDF2
    import pdfplumber
except ImportError:
    messagebox.showerror("Error", "Install required libraries:\npip install PyPDF2 pdfplumber")
    sys.exit(1)

class PDFConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF to Text Converter")
        self.root.geometry("800x600")
        self.root.resizable(True, True)
        
        # Variables
        self.selected_files = []
        self.extraction_method = tk.StringVar(value="pdfplumber")
        self.output_folder = tk.StringVar()
        self.progress_var = tk.DoubleVar()
        self.status_var = tk.StringVar(value="Ready")
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup user interface"""
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="PDF to Text Converter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File selection section
        files_frame = ttk.LabelFrame(main_frame, text="Select PDF Files", padding="10")
        files_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        files_frame.columnconfigure(0, weight=1)
        
        # File list
        self.file_listbox = tk.Listbox(files_frame, height=6, selectmode=tk.EXTENDED)
        self.file_listbox.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # File buttons
        btn_frame = ttk.Frame(files_frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E))
        
        ttk.Button(btn_frame, text="Add Files", command=self.add_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Add Folder", command=self.add_folder).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Clear List", command=self.clear_files).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(btn_frame, text="Remove Selected", command=self.remove_selected).pack(side=tk.LEFT)
        
        # Settings section
        settings_frame = ttk.LabelFrame(main_frame, text="Settings", padding="10")
        settings_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Extraction method
        ttk.Label(settings_frame, text="Extraction Method:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        method_combo = ttk.Combobox(settings_frame, textvariable=self.extraction_method, 
                                   values=["pdfplumber", "PyPDF2"], state="readonly", width=15)
        method_combo.grid(row=0, column=1, sticky=tk.W, padx=(0, 20))
        
        # Output folder
        ttk.Label(settings_frame, text="Output Folder:").grid(row=0, column=2, sticky=tk.W, padx=(0, 10))
        ttk.Entry(settings_frame, textvariable=self.output_folder, width=30).grid(row=0, column=3, sticky=(tk.W, tk.E), padx=(0, 5))
        ttk.Button(settings_frame, text="Browse", command=self.browse_output_folder).grid(row=0, column=4)
        
        settings_frame.columnconfigure(3, weight=1)
        
        # Convert button
        convert_frame = ttk.Frame(main_frame)
        convert_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        
        self.convert_btn = ttk.Button(convert_frame, text="Convert to Text", 
                                     command=self.start_conversion, style="Accent.TButton")
        self.convert_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        ttk.Button(convert_frame, text="Preview Single File", 
                  command=self.preview_single_file).pack(side=tk.LEFT)
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress", padding="10")
        progress_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_bar = ttk.Progressbar(progress_frame, variable=self.progress_var, 
                                           maximum=100, length=400)
        self.progress_bar.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 5))
        
        self.status_label = ttk.Label(progress_frame, textvariable=self.status_var)
        self.status_label.grid(row=1, column=0, sticky=tk.W)
        
        # Results section
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="10")
        results_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        self.results_text = scrolledtext.ScrolledText(results_frame, height=10, wrap=tk.WORD)
        self.results_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure main frame row weights
        main_frame.rowconfigure(5, weight=1)
        
    def add_files(self):
        """Add PDF files"""
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]
        )
        
        for file_path in files:
            if file_path not in self.selected_files:
                self.selected_files.append(file_path)
                self.file_listbox.insert(tk.END, os.path.basename(file_path))
        
        self.update_status(f"Added {len(files)} files. Total: {len(self.selected_files)} files")
    
    def add_folder(self):
        """Add all PDF files from a folder"""
        folder = filedialog.askdirectory(title="Select Folder Containing PDF Files")
        
        if folder:
            pdf_files = [f for f in os.listdir(folder) if f.lower().endswith('.pdf')]
            added_count = 0
            
            for pdf_file in pdf_files:
                file_path = os.path.join(folder, pdf_file)
                if file_path not in self.selected_files:
                    self.selected_files.append(file_path)
                    self.file_listbox.insert(tk.END, pdf_file)
                    added_count += 1
            
            self.update_status(f"Added {added_count} files from folder. Total: {len(self.selected_files)} files")
    
    def clear_files(self):
        """Clear all selected files"""
        self.selected_files.clear()
        self.file_listbox.delete(0, tk.END)
        self.update_status("File list cleared")
    
    def remove_selected(self):
        """Remove selected files from list"""
        selected_indices = self.file_listbox.curselection()
        
        # Remove in reverse order to avoid index shifting
        for index in reversed(selected_indices):
            self.file_listbox.delete(index)
            del self.selected_files[index]
        
        self.update_status(f"Removed {len(selected_indices)} files. Total: {len(self.selected_files)} files")
    
    def browse_output_folder(self):
        """Browse for output folder"""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
    
    def update_status(self, message):
        """Update status message"""
        self.status_var.set(message)
        self.root.update_idletasks()
    
    def extract_text_pdfplumber(self, pdf_path):
        """Extract text using pdfplumber"""
        text = ""
        try:
            with pdfplumber.open(pdf_path) as pdf:
                for page_num, page in enumerate(pdf.pages, 1):
                    page_text = page.extract_text()
                    if page_text:
                        text += f"\n--- Page {page_num} ---\n{page_text}\n"
            return text
        except Exception as e:
            raise Exception(f"pdfplumber error: {e}")
    
    def extract_text_pypdf2(self, pdf_path):
        """Extract text using PyPDF2"""
        text = ""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                for page_num, page in enumerate(pdf_reader.pages, 1):
                    page_text = page.extract_text()
                    text += f"\n--- Page {page_num} ---\n{page_text}\n"
            return text
        except Exception as e:
            raise Exception(f"PyPDF2 error: {e}")
    
    def get_pdf_info(self, pdf_path):
        """Get PDF information"""
        try:
            with open(pdf_path, 'rb') as file:
                pdf_reader = PyPDF2.PdfReader(file)
                return {
                    'pages': len(pdf_reader.pages),
                    'size': os.path.getsize(pdf_path),
                    'title': pdf_reader.metadata.get('/Title', 'N/A') if pdf_reader.metadata else 'N/A'
                }
        except:
            return {'pages': 'Unknown', 'size': 0, 'title': 'N/A'}
    
    def convert_single_file(self, pdf_path, output_dir):
        """Convert single PDF file"""
        filename = os.path.basename(pdf_path)
        
        try:
            # Get PDF info
            pdf_info = self.get_pdf_info(pdf_path)
            
            # Extract text based on method
            if self.extraction_method.get() == "pdfplumber":
                text = self.extract_text_pdfplumber(pdf_path)
            else:
                text = self.extract_text_pypdf2(pdf_path)
            
            if not text.strip():
                return f"❌ {filename}: No text found (possibly scanned PDF)"
            
            # Generate output filename
            output_filename = f"{Path(pdf_path).stem}.txt"
            output_path = os.path.join(output_dir, output_filename)
            
            # Add header with file info
            header = f"PDF: {filename}\n"
            header += f"Pages: {pdf_info['pages']}\n"
            header += f"Size: {pdf_info['size']:,} bytes\n"
            header += f"Method: {self.extraction_method.get()}\n"
            header += "=" * 50 + "\n\n"
            
            # Save file
            with open(output_path, 'w', encoding='utf-8') as f:
                f.write(header + text)
            
            return f"✅ {filename}: {len(text)} characters extracted"
            
        except Exception as e:
            return f"❌ {filename}: {str(e)}"
    
    def start_conversion(self):
        """Start conversion process in separate thread"""
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select PDF files first!")
            return
        
        if not self.output_folder.get():
            # Set default output folder
            self.output_folder.set(os.path.join(os.getcwd(), "extracted_texts"))
        
        # Create output directory
        os.makedirs(self.output_folder.get(), exist_ok=True)
        
        # Disable convert button
        self.convert_btn.config(state="disabled")
        
        # Start conversion thread
        thread = threading.Thread(target=self.run_conversion)
        thread.daemon = True
        thread.start()
    
    def run_conversion(self):
        """Run conversion process"""
        results = []
        total_files = len(self.selected_files)
        
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"Converting {total_files} files...\n\n")
        
        for i, pdf_path in enumerate(self.selected_files):
            # Update progress
            progress = (i / total_files) * 100
            self.progress_var.set(progress)
            self.update_status(f"Processing {i+1}/{total_files}: {os.path.basename(pdf_path)}")
            
            # Convert file
            result = self.convert_single_file(pdf_path, self.output_folder.get())
            results.append(result)
            
            # Update results display
            self.results_text.insert(tk.END, result + "\n")
            self.results_text.see(tk.END)
            self.root.update_idletasks()
        
        # Complete
        self.progress_var.set(100)
        self.update_status("Conversion completed!")
        
        # Show summary
        success_count = sum(1 for r in results if r.startswith("✅"))
        self.results_text.insert(tk.END, f"\n{'='*50}\n")
        self.results_text.insert(tk.END, f"SUMMARY:\n")
        self.results_text.insert(tk.END, f"Total files: {total_files}\n")
        self.results_text.insert(tk.END, f"Successful: {success_count}\n")
        self.results_text.insert(tk.END, f"Failed: {total_files - success_count}\n")
        self.results_text.insert(tk.END, f"Output folder: {self.output_folder.get()}\n")
        
        # Re-enable convert button
        self.convert_btn.config(state="normal")
        
        # Show completion message
        messagebox.showinfo("Completed", 
                           f"Conversion completed!\n\n"
                           f"Successful: {success_count}/{total_files}\n"
                           f"Output folder: {self.output_folder.get()}")
    
    def preview_single_file(self):
        """Preview text from a single selected file"""
        if not self.selected_files:
            messagebox.showwarning("Warning", "Please select PDF files first!")
            return
        
        # Show file selection dialog
        if len(self.selected_files) == 1:
            selected_file = self.selected_files[0]
        else:
            # Let user choose from selected files
            file_names = [os.path.basename(f) for f in self.selected_files]
            dialog = FileSelectionDialog(self.root, file_names)
            if dialog.result is None:
                return
            selected_file = self.selected_files[dialog.result]
        
        # Extract and show preview
        try:
            self.update_status("Extracting text for preview...")
            
            if self.extraction_method.get() == "pdfplumber":
                text = self.extract_text_pdfplumber(selected_file)
            else:
                text = self.extract_text_pypdf2(selected_file)
            
            # Show preview window
            PreviewWindow(self.root, os.path.basename(selected_file), text)
            
            self.update_status("Preview completed")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract text:\n{e}")
            self.update_status("Preview failed")

class FileSelectionDialog:
    """Dialog untuk memilih file dari list"""
    def __init__(self, parent, file_names):
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Select File to Preview")
        self.dialog.geometry("400x300")
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Center dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 100, parent.winfo_rooty() + 100))
        
        # Listbox
        ttk.Label(self.dialog, text="Select file to preview:").pack(pady=10)
        
        listbox = tk.Listbox(self.dialog, height=10)
        listbox.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        for name in file_names:
            listbox.insert(tk.END, name)
        
        # Buttons
        btn_frame = ttk.Frame(self.dialog)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Preview", 
                  command=lambda: self.select_file(listbox.curselection())).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Cancel", 
                  command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def select_file(self, selection):
        if selection:
            self.result = selection[0]
            self.dialog.destroy()

class PreviewWindow:
    """Window untuk preview text"""
    def __init__(self, parent, filename, text):
        self.window = tk.Toplevel(parent)
        self.window.title(f"Preview: {filename}")
        self.window.geometry("800x600")
        self.window.transient(parent)
        
        # Text widget
        text_widget = scrolledtext.ScrolledText(self.window, wrap=tk.WORD)
        text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Show text
        if text.strip():
            text_widget.insert(1.0, text)
        else:
            text_widget.insert(1.0, "No text found in this PDF file.\n\nThis might be a scanned PDF that requires OCR.")
        
        text_widget.config(state=tk.DISABLED)
        
        # Buttons
        btn_frame = ttk.Frame(self.window)
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="Save to File", 
                  command=lambda: self.save_text(text, filename)).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Close", 
                  command=self.window.destroy).pack(side=tk.LEFT, padx=5)
    
    def save_text(self, text, filename):
        output_file = filedialog.asksaveasfilename(
            title="Save Text As",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialvalue=f"{Path(filename).stem}.txt"
        )
        
        if output_file:
            try:
                with open(output_file, 'w', encoding='utf-8') as f:
                    f.write(text)
                messagebox.showinfo("Success", f"Text saved to:\n{output_file}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save file:\n{e}")

def main():
    """Main function"""
    root = tk.Tk()
    
    # Set theme (optional)
    try:
        root.tk.call("source", "azure.tcl")
        root.tk.call("set_theme", "light")
    except:
        pass  # Theme not available, use default
    
    app = PDFConverterGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()