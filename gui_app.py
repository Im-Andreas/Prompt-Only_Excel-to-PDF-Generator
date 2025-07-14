import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import shutil
from pathlib import Path
import threading
from main import create_professor_pie_charts, read_excel, cleanup_temp_folder

class ProfessorReportGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Professor Evaluation Report Generator")
        self.root.geometry("800x700")  # Larger default size
        self.root.minsize(600, 500)  # Minimum window size
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.excel_file_path = tk.StringVar()
        self.selected_professor = tk.StringVar()
        self.professors_list = []
        self.data = None
        
        # Configure scaling for high DPI displays
        self.root.tk.call('tk', 'scaling', 1.2)
        
        # Create main frame with better padding
        main_frame = ttk.Frame(root, padding="25")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights for responsive design
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        for i in range(5):  # Configure all rows to expand
            main_frame.rowconfigure(i, weight=1)
        
        # Title with responsive font
        title_label = ttk.Label(main_frame, text="Professor Evaluation Report Generator", 
                               font=('Arial', 18, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 25), sticky=(tk.W, tk.E))
        
        # Step 1: Excel File Selection
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Select Excel File", padding="15")
        step1_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        step1_frame.columnconfigure(1, weight=1)
        
        ttk.Label(step1_frame, text="Excel File:", font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.file_entry = ttk.Entry(step1_frame, textvariable=self.excel_file_path, state='readonly', font=('Arial', 10))
        self.file_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))
        
        self.browse_button = ttk.Button(step1_frame, text="Browse", command=self.browse_file)
        self.browse_button.grid(row=0, column=2, padx=(0, 5))
        
        # File status
        self.file_status = ttk.Label(step1_frame, text="No file selected", foreground="red", font=('Arial', 9))
        self.file_status.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # Step 2: Professor Selection
        step2_frame = ttk.LabelFrame(main_frame, text="Step 2: Select Professor", padding="15")
        step2_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        step2_frame.columnconfigure(1, weight=1)
        
        ttk.Label(step2_frame, text="Professor:", font=('Arial', 10)).grid(row=0, column=0, sticky=tk.W, padx=(0, 15))
        
        self.professor_combo = ttk.Combobox(step2_frame, textvariable=self.selected_professor, 
                                          state='readonly', font=('Arial', 10))
        self.professor_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 15))
        
        # Professor selection status
        self.prof_status = ttk.Label(step2_frame, text="Please select an Excel file first", 
                                   foreground="orange", font=('Arial', 9))
        self.prof_status.grid(row=1, column=0, columnspan=3, pady=(10, 0), sticky=(tk.W, tk.E))
        
        # Step 3: Generate Reports
        step3_frame = ttk.LabelFrame(main_frame, text="Step 3: Generate Reports", padding="15")
        step3_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 15))
        step3_frame.columnconfigure(0, weight=1)
        step3_frame.rowconfigure(2, weight=1)  # Make status text area expandable
        
        self.generate_button = ttk.Button(step3_frame, text="Generate PDF Report(s)", 
                                        command=self.generate_reports, state='disabled')
        self.generate_button.grid(row=0, column=0, pady=15, sticky=(tk.W, tk.E))
        
        # Progress bar
        self.progress = ttk.Progressbar(step3_frame, mode='indeterminate')
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Status text with better sizing
        self.status_text = tk.Text(step3_frame, height=10, font=('Consolas', 9), 
                                 state='disabled', wrap=tk.WORD, bg='#f8f8f8')
        self.status_text.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for status text
        scrollbar = ttk.Scrollbar(step3_frame, orient="vertical", command=self.status_text.yview)
        scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
        self.status_text.configure(yscrollcommand=scrollbar.set)
        
        # Step 4: Download Location Info
        step4_frame = ttk.LabelFrame(main_frame, text="Step 4: Download Location", padding="15")
        step4_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 15))
        
        downloads_path = str(Path.home() / "Downloads")
        ttk.Label(step4_frame, text=f"PDF reports will be saved to: {downloads_path}", 
                 font=('Arial', 10), wraplength=750).pack(fill=tk.X)
        
        # Initialize UI state
        self.update_ui_state()
    
    def browse_file(self):
        """Open file dialog to select Excel file"""
        file_types = [
            ('Excel files', '*.xlsx *.xls'),
            ('All files', '*.*')
        ]
        
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=file_types
        )
        
        if filename:
            self.excel_file_path.set(filename)
            self.load_excel_file(filename)
    
    def load_excel_file(self, filename):
        """Load Excel file and extract professor names"""
        try:
            # Copy file to assets folder with correct name
            assets_dir = "assets"
            os.makedirs(assets_dir, exist_ok=True)
            
            # Copy and rename file
            target_file = os.path.join(assets_dir, "QuestionPro-SR-RawData.xlsx")
            shutil.copy2(filename, target_file)
            
            # Load the data
            self.data = read_excel(target_file)
            
            # Extract professor names
            professors = self.data['Level 2'].dropna().unique()
            professors = sorted([prof for prof in professors if pd.notna(prof)])
            
            # Update professor list
            self.professors_list = ["All Professors"] + list(professors)
            self.professor_combo['values'] = self.professors_list
            self.selected_professor.set("All Professors")  # Default selection
            
            # Update status
            self.file_status.config(text=f"✓ File loaded successfully ({len(professors)} professors found)", 
                                  foreground="green")
            self.prof_status.config(text=f"✓ {len(professors)} professors available for selection", 
                                  foreground="green")
            
            self.log_status(f"Excel file loaded successfully: {os.path.basename(filename)}")
            self.log_status(f"Found {len(professors)} professors in the data")
            self.log_status("File copied to assets folder as QuestionPro-SR-RawData.xlsx")
            
        except Exception as e:
            self.file_status.config(text=f"✗ Error loading file: {str(e)}", foreground="red")
            self.prof_status.config(text="Please select a valid Excel file", foreground="red")
            self.log_status(f"Error loading file: {str(e)}")
            messagebox.showerror("Error", f"Failed to load Excel file:\n{str(e)}")
        
        self.update_ui_state()
    
    def update_ui_state(self):
        """Update UI elements based on current state"""
        if self.data is not None and len(self.professors_list) > 0:
            self.generate_button.config(state='normal')
        else:
            self.generate_button.config(state='disabled')
    
    def log_status(self, message):
        """Add message to status log"""
        self.status_text.config(state='normal')
        self.status_text.insert(tk.END, f"{message}\n")
        self.status_text.see(tk.END)
        self.status_text.config(state='disabled')
        self.root.update()
    
    def generate_reports(self):
        """Generate PDF reports based on selection"""
        if not self.data is not None:
            messagebox.showerror("Error", "Please load an Excel file first")
            return
        
        if not self.selected_professor.get():
            messagebox.showerror("Error", "Please select a professor")
            return
        
        # Disable button and start progress
        self.generate_button.config(state='disabled')
        self.progress.start()
        
        # Run generation in separate thread to prevent UI freezing
        thread = threading.Thread(target=self._generate_reports_thread)
        thread.daemon = True
        thread.start()
    
    def _generate_reports_thread(self):
        """Thread function for report generation"""
        try:
            selected = self.selected_professor.get()
            
            if selected == "All Professors":
                self.log_status("Starting generation for all professors...")
                specific_professor = None
                total_professors = len([p for p in self.professors_list if p != "All Professors"])
                self.log_status(f"Will generate reports for {total_professors} professors")
                self.log_status("⚠ This may take several minutes for large datasets...")
            else:
                self.log_status(f"Starting generation for professor: {selected}")
                specific_professor = selected
                total_professors = 1
            
            # Create output directory if it doesn't exist
            os.makedirs("output", exist_ok=True)
            
            # Define progress callback function
            def progress_callback(message):
                # Update UI on main thread
                self.root.after(0, lambda: self.log_status(message))
            
            # Generate reports with progress feedback
            create_professor_pie_charts(self.data, specific_professor, progress_callback)
            
            # Move generated PDFs to Downloads folder
            downloads_dir = Path.home() / "Downloads"
            output_dir = Path("output")
            
            moved_files = []
            if output_dir.exists():
                for pdf_file in output_dir.glob("*.pdf"):
                    target_path = downloads_dir / pdf_file.name
                    
                    # Handle file conflicts
                    counter = 1
                    original_target = target_path
                    while target_path.exists():
                        stem = original_target.stem
                        suffix = original_target.suffix
                        target_path = downloads_dir / f"{stem}_{counter}{suffix}"
                        counter += 1
                    
                    shutil.move(str(pdf_file), str(target_path))
                    moved_files.append(target_path.name)
                    self.root.after(0, lambda f=pdf_file.name: self.log_status(f"✓ Moved {f} to Downloads folder"))
            
            # Clean up
            cleanup_temp_folder()
            
            # Update UI on main thread
            self.root.after(0, self._generation_complete, moved_files, total_professors)
            
        except Exception as e:
            self.root.after(0, self._generation_error, str(e))
    
    def _generation_complete(self, moved_files, total_professors):
        """Called when generation is complete"""
        self.progress.stop()
        self.generate_button.config(state='normal')
        
        if moved_files:
            self.log_status(f"\n✓ SUCCESS! Generated {len(moved_files)} PDF report(s)")
            self.log_status(f"Files saved to Downloads folder:")
            for file in moved_files:
                self.log_status(f"  - {file}")
            
            messagebox.showinfo("Success", 
                              f"Successfully generated {len(moved_files)} PDF report(s)!\n\n"
                              f"Files saved to your Downloads folder:\n" + 
                              "\n".join([f"• {file}" for file in moved_files[:5]]) +
                              (f"\n... and {len(moved_files)-5} more" if len(moved_files) > 5 else ""))
        else:
            self.log_status("⚠ No PDF files were generated")
            messagebox.showwarning("Warning", "No PDF files were generated. Please check your data.")
    
    def _generation_error(self, error_message):
        """Called when generation encounters an error"""
        self.progress.stop()
        self.generate_button.config(state='normal')
        
        self.log_status(f"\n✗ ERROR: {error_message}")
        messagebox.showerror("Error", f"Failed to generate reports:\n{error_message}")

def main():
    """Create and run the GUI application"""
    # Create necessary directories
    os.makedirs("temp", exist_ok=True)
    os.makedirs("assets", exist_ok=True)
    os.makedirs("output", exist_ok=True)
    
    # Create and run GUI
    root = tk.Tk()
    app = ProfessorReportGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
