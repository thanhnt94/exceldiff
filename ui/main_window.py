import customtkinter as ctk
from tkinter import filedialog
import openpyxl
import os
import threading
from core.comparator import ExcelComparator
from core.data_types import DiffResult

# Modern theme
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

class MainWindow(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        self.title("ExcelDiff - File Comparison Tool")
        self.geometry("700x450")
        self.resizable(False, False)
        
        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=0)  # Header
        self.grid_rowconfigure(1, weight=1)  # Content
        self.grid_rowconfigure(2, weight=0)  # Status bar
        
        self._create_header()
        self._create_content()
        self._create_status_bar()
        
    def _create_header(self):
        """Create header with app title."""
        header = ctk.CTkFrame(self, fg_color="#0f3460", corner_radius=0, height=50)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_propagate(False)
        
        title = ctk.CTkLabel(header, text="🔍 ExcelDiff", font=ctk.CTkFont(size=20, weight="bold"))
        title.pack(side="left", padx=20, pady=10)
        
        subtitle = ctk.CTkLabel(header, text="File Comparison Tool", font=ctk.CTkFont(size=12), text_color="#aaa")
        subtitle.pack(side="left", pady=10)
        
    def _create_content(self):
        """Create main content area."""
        content = ctk.CTkFrame(self, fg_color="transparent")
        content.grid(row=1, column=0, sticky="nsew", padx=30, pady=20)
        content.grid_columnconfigure(1, weight=1)
        
        # File A
        ctk.CTkLabel(content, text="Original File (Base):", font=ctk.CTkFont(size=13)).grid(row=0, column=0, sticky="w", pady=(0, 5))
        
        file_a_frame = ctk.CTkFrame(content, fg_color="transparent")
        file_a_frame.grid(row=1, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        file_a_frame.grid_columnconfigure(0, weight=1)
        
        self.entry_file_a = ctk.CTkEntry(file_a_frame, placeholder_text="Select Excel file...", height=35)
        self.entry_file_a.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        self.btn_browse_a = ctk.CTkButton(file_a_frame, text="📁", width=40, height=35, command=self.browse_file_a)
        self.btn_browse_a.grid(row=0, column=1, padx=(0, 10))
        
        self.combo_sheet_a = ctk.CTkComboBox(file_a_frame, width=120, height=35, values=["Select File First"])
        self.combo_sheet_a.grid(row=0, column=2)
        
        # File B
        ctk.CTkLabel(content, text="Modified File:", font=ctk.CTkFont(size=13)).grid(row=2, column=0, sticky="w", pady=(0, 5))
        
        file_b_frame = ctk.CTkFrame(content, fg_color="transparent")
        file_b_frame.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, 15))
        file_b_frame.grid_columnconfigure(0, weight=1)
        
        self.entry_file_b = ctk.CTkEntry(file_b_frame, placeholder_text="Select Excel file...", height=35)
        self.entry_file_b.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        self.btn_browse_b = ctk.CTkButton(file_b_frame, text="📁", width=40, height=35, command=self.browse_file_b)
        self.btn_browse_b.grid(row=0, column=1, padx=(0, 10))
        
        self.combo_sheet_b = ctk.CTkComboBox(file_b_frame, width=120, height=35, values=["Select File First"])
        self.combo_sheet_b.grid(row=0, column=2)
        
        # Output folder
        ctk.CTkLabel(content, text="Output Folder:", font=ctk.CTkFont(size=13)).grid(row=4, column=0, sticky="w", pady=(0, 5))
        
        output_frame = ctk.CTkFrame(content, fg_color="transparent")
        output_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 25))
        output_frame.grid_columnconfigure(0, weight=1)
        
        self.entry_output = ctk.CTkEntry(output_frame, placeholder_text="Output folder...", height=35)
        self.entry_output.insert(0, os.getcwd())
        self.entry_output.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        
        self.btn_browse_output = ctk.CTkButton(output_frame, text="📁", width=40, height=35, command=self.browse_output)
        self.btn_browse_output.grid(row=0, column=1, padx=(0, 10))
        
        self.btn_open_output = ctk.CTkButton(output_frame, text="Open Folder", width=120, height=35, 
                                              fg_color="#2d4a3e", hover_color="#3d5a4e",
                                              command=lambda: os.startfile(self.entry_output.get()) if os.path.isdir(self.entry_output.get()) else None)
        self.btn_open_output.grid(row=0, column=2)
        
        # Compare button
        self.btn_compare = ctk.CTkButton(
            content, 
            text="🔄 Compare Files", 
            font=ctk.CTkFont(size=15, weight="bold"),
            height=45,
            fg_color="#e94560",
            hover_color="#c73e54",
            command=self.run_comparison
        )
        self.btn_compare.grid(row=6, column=0, columnspan=3, pady=(0, 10))
        
    def _create_status_bar(self):
        """Create status bar with copyright."""
        status_bar = ctk.CTkFrame(self, fg_color="#16213e", corner_radius=0, height=35)
        status_bar.grid(row=2, column=0, sticky="ew")
        status_bar.grid_propagate(False)
        status_bar.grid_columnconfigure(0, weight=1)
        
        self.lbl_status = ctk.CTkLabel(status_bar, text="Ready", font=ctk.CTkFont(size=11), text_color="#888")
        self.lbl_status.pack(side="left", padx=15)
        
        copyright_lbl = ctk.CTkLabel(status_bar, text="© KNT15083", font=ctk.CTkFont(size=11), text_color="#666")
        copyright_lbl.pack(side="right", padx=15)
        
    def browse_file_a(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.entry_file_a.delete(0, "end")
            self.entry_file_a.insert(0, path)
            self.load_sheets(path, self.combo_sheet_a)
            
    def browse_file_b(self):
        path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if path:
            self.entry_file_b.delete(0, "end")
            self.entry_file_b.insert(0, path)
            self.load_sheets(path, self.combo_sheet_b)
            
    def browse_output(self):
        path = filedialog.askdirectory()
        if path:
            self.entry_output.delete(0, "end")
            self.entry_output.insert(0, path)
            
    def load_sheets(self, filepath, combo_box):
        def _load():
            try:
                wb = openpyxl.load_workbook(filepath, read_only=True, keep_links=False)
                sheets = wb.sheetnames
                wb.close()
                self.after(0, lambda: combo_box.configure(values=sheets))
                if sheets:
                    self.after(0, lambda: combo_box.set(sheets[0]))
            except Exception as e:
                self.after(0, lambda: self.update_status(f"Error: {e}", "red"))
        threading.Thread(target=_load, daemon=True).start()
        
    def update_status(self, text, color="#888"):
        self.lbl_status.configure(text=text, text_color=color)
        
    def run_comparison(self):
        file_a = self.entry_file_a.get()
        file_b = self.entry_file_b.get()
        
        if not file_a or not file_b:
            self.update_status("Please select both files.", "#e94560")
            return
            
        self.update_status("Comparing files...", "#f0a500")
        self.btn_compare.configure(state="disabled")
        
        sheet_a = self.combo_sheet_a.get()
        sheet_b = self.combo_sheet_b.get()
        
        threading.Thread(target=self._compare_thread, args=(file_a, file_b, sheet_a, sheet_b), daemon=True).start()
        
    def _compare_thread(self, file_a, file_b, sheet_a, sheet_b):
        try:
            self.after(0, lambda: self.update_status("Loading files...", "#f0a500"))
            comparator = ExcelComparator(file_a, file_b, sheet_a, sheet_b)
            
            self.after(0, lambda: self.update_status("Analyzing differences...", "#f0a500"))
            result = comparator.compare()
            
            self.after(0, lambda: self.update_status("Generating report...", "#f0a500"))
            
            output_dir = self.entry_output.get()
            if not output_dir or not os.path.isdir(output_dir):
                output_dir = os.getcwd()
                
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_path = os.path.join(output_dir, f"ExcelDiff_Report_{timestamp}.xlsx")
            
            from reporting.visual_reporter import VisualReporter
            reporter = VisualReporter(file_a, file_b, output_path, sheet_a, sheet_b)
            reporter.generate(result)
            
            self.after(0, lambda: self._on_complete(output_path, len(result.items)))
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.after(0, lambda: self.update_status(f"Error: {e}", "#e94560"))
            self.after(0, lambda: self.btn_compare.configure(state="normal"))
            
    def _on_complete(self, output_path, diff_count):
        self.btn_compare.configure(state="normal")
        filename = os.path.basename(output_path)
        self.update_status(f"✅ Done! {diff_count} differences found. Report: {filename}", "#4ade80")
        
        # Open the report file
        try:
            os.startfile(output_path)
        except:
            pass

if __name__ == "__main__":
    app = MainWindow()
    app.mainloop()
