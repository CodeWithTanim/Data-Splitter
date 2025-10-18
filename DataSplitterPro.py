import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from tkinter.font import Font
import pandas as pd
import xml.etree.ElementTree as ET
import openpyxl
from openpyxl import load_workbook
from ttkthemes import ThemedStyle
import math

class DataSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Splitter Pro")
        self.root.geometry("800x700")  # Increased height to accommodate new feature
        self.root.minsize(700, 600)
        
        # Apply a modern theme
        self.style = ThemedStyle(self.root)
        self.style.set_theme("arc")
        
        # Custom fonts
        self.title_font = Font(family="Segoe UI", size=16, weight="bold")
        self.subtitle_font = Font(family="Segoe UI", size=12)
        self.button_font = Font(family="Segoe UI", size=10, weight="bold")
        
        # Configure styles
        self.style.configure("Title.TLabel", font=self.title_font, foreground="#333")
        self.style.configure("Subtitle.TLabel", font=self.subtitle_font, foreground="#555")
        self.style.configure("TButton", font=self.button_font, padding=6)
        self.style.configure("TFrame", background=self.style.lookup("TFrame", "background"))
        
        # Create UI elements
        self.create_widgets()
        
    def create_widgets(self):
        # Main container
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header
        self.header_frame = ttk.Frame(self.main_frame)
        self.header_frame.pack(fill=tk.X, pady=(0, 20))
        
        self.title_label = ttk.Label(
            self.header_frame,
            text="Data Splitter Pro",
            style="Title.TLabel"
        )
        self.title_label.pack(side=tk.LEFT)
        
        self.version_label = ttk.Label(
            self.header_frame,
            text="v1.0",
            style="Subtitle.TLabel"
        )
        self.version_label.pack(side=tk.RIGHT)
        
        # Input section
        self.input_frame = ttk.LabelFrame(
            self.main_frame,
            text="Input File",
            padding=(15, 10)
        )
        self.input_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.file_path = tk.StringVar()
        self.file_type = tk.StringVar(value="csv")
        
        ttk.Label(self.input_frame, text="File Path:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.path_entry = ttk.Entry(self.input_frame, textvariable=self.file_path, width=50)
        self.path_entry.grid(row=0, column=1, sticky=tk.EW, padx=(0, 10))
        
        self.browse_btn = ttk.Button(
            self.input_frame,
            text="Browse",
            command=self.browse_file,
            width=10
        )
        self.browse_btn.grid(row=0, column=2, sticky=tk.E)
        
        ttk.Label(self.input_frame, text="File Type:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.file_type_combo = ttk.Combobox(
            self.input_frame,
            textvariable=self.file_type,
            values=["csv", "xlsx", "xml", "txt"],
            state="readonly",
            width=10
        )
        self.file_type_combo.grid(row=1, column=1, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        # Split options
        self.options_frame = ttk.LabelFrame(
            self.main_frame,
            text="Split Options",
            padding=(15, 10)
        )
        self.options_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.split_method = tk.StringVar(value="equal_parts")
        
        ttk.Radiobutton(
            self.options_frame,
            text="Split into equal parts",
            variable=self.split_method,
            value="equal_parts"
        ).grid(row=0, column=0, sticky=tk.W, padx=(0, 20))
        
        ttk.Radiobutton(
            self.options_frame,
            text="Split by rows per file",
            variable=self.split_method,
            value="rows_per_file"
        ).grid(row=0, column=1, sticky=tk.W)
        
        self.parts_label = ttk.Label(self.options_frame, text="Number of parts:")
        self.parts_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.parts_entry = ttk.Spinbox(
            self.options_frame,
            from_=2,
            to=100,
            width=10
        )
        self.parts_entry.grid(row=1, column=0, sticky=tk.W, padx=(100, 0), pady=(10, 0))
        self.parts_entry.set(2)
        
        self.rows_label = ttk.Label(self.options_frame, text="Rows per file:")
        self.rows_label.grid(row=1, column=1, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        
        self.rows_entry = ttk.Spinbox(
            self.options_frame,
            from_=1,
            to=100000,
            width=10
        )
        self.rows_entry.grid(row=1, column=1, sticky=tk.W, padx=(100, 0), pady=(10, 0))
        self.rows_entry.set(1000)
        
        # Custom Text Section
        self.text_frame = ttk.LabelFrame(
            self.main_frame,
            text="Custom Text to Add to Each File",
            padding=(15, 10)
        )
        self.text_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.add_custom_text = tk.BooleanVar(value=False)
        self.custom_text_check = ttk.Checkbutton(
            self.text_frame,
            text="Add custom text to each split file",
            variable=self.add_custom_text,
            command=self.toggle_custom_text
        )
        self.custom_text_check.grid(row=0, column=0, sticky=tk.W, pady=(0, 10))
        
        # Text area for custom text with scrollbar
        self.text_area_frame = ttk.Frame(self.text_frame)
        self.text_area_frame.grid(row=1, column=0, sticky=tk.EW, pady=(5, 0))
        
        self.custom_text_area = scrolledtext.ScrolledText(
            self.text_area_frame,
            height=4,
            width=70,
            wrap=tk.WORD
        )
        self.custom_text_area.pack(fill=tk.BOTH, expand=True)
        
        # Add some sample text as placeholder
        self.custom_text_area.insert(tk.END, "Email Header Line 1\nEmail Header Line 2")
        
        ttk.Label(
            self.text_frame, 
            text="Enter the text you want to add at the beginning of each split file (one line per row)",
            foreground="#666",
            font=("Segoe UI", 9)
        ).grid(row=2, column=0, sticky=tk.W, pady=(5, 0))
        
        # Output options
        self.output_frame = ttk.LabelFrame(
            self.main_frame,
            text="Output Options",
            padding=(15, 10)
        )
        self.output_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(self.output_frame, text="Output Format:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.output_format = tk.StringVar(value="csv")
        self.format_combo = ttk.Combobox(
            self.output_frame,
            textvariable=self.output_format,
            values=["csv", "xlsx", "xml", "txt"],
            state="readonly",
            width=10
        )
        self.format_combo.grid(row=0, column=0, sticky=tk.W, padx=(100, 0))
        
        ttk.Label(self.output_frame, text="Output Directory:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(10, 0))
        self.output_dir = tk.StringVar()
        self.dir_entry = ttk.Entry(self.output_frame, textvariable=self.output_dir, width=50)
        self.dir_entry.grid(row=1, column=0, sticky=tk.EW, padx=(100, 10), pady=(10, 0))
        
        self.dir_btn = ttk.Button(
            self.output_frame,
            text="Browse",
            command=self.browse_directory,
            width=10
        )
        self.dir_btn.grid(row=1, column=1, sticky=tk.E, pady=(10, 0))
        
        # Action buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.split_btn = ttk.Button(
            self.button_frame,
            text="Split Data",
            command=self.split_data,
            style="Accent.TButton"
        )
        self.split_btn.pack(side=tk.RIGHT, padx=(10, 0))
        
        self.clear_btn = ttk.Button(
            self.button_frame,
            text="Clear",
            command=self.clear_form
        )
        self.clear_btn.pack(side=tk.RIGHT)
        
        # Status bar
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.status_label = ttk.Label(
            self.status_frame,
            text="Ready",
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        self.status_label.pack(fill=tk.X)
        
        # Configure grid weights
        self.input_frame.columnconfigure(1, weight=1)
        self.output_frame.columnconfigure(0, weight=1)
        self.text_frame.columnconfigure(0, weight=1)
        
        # Bind events
        self.split_method.trace_add("write", self.update_ui)
        self.update_ui()
        self.toggle_custom_text()
    
    def toggle_custom_text(self):
        if self.add_custom_text.get():
            self.custom_text_area.config(state=tk.NORMAL)
        else:
            self.custom_text_area.config(state=tk.DISABLED)
    
    def update_ui(self, *args):
        if self.split_method.get() == "equal_parts":
            self.parts_label.config(state=tk.NORMAL)
            self.parts_entry.config(state=tk.NORMAL)
            self.rows_label.config(state=tk.DISABLED)
            self.rows_entry.config(state=tk.DISABLED)
        else:
            self.parts_label.config(state=tk.DISABLED)
            self.parts_entry.config(state=tk.DISABLED)
            self.rows_label.config(state=tk.NORMAL)
            self.rows_entry.config(state=tk.NORMAL)
    
    def browse_file(self):
        filetypes = (
            ("CSV Files", "*.csv"),
            ("Excel Files", "*.xlsx"),
            ("XML Files", "*.xml"),
            ("Text Files", "*.txt"),
            ("All Files", "*.*")
        )
        
        filename = filedialog.askopenfilename(
            title="Select a file",
            filetypes=filetypes
        )
        
        if filename:
            self.file_path.set(filename)
            ext = filename.split(".")[-1].lower()
            if ext in ["csv", "xlsx", "xml", "txt"]:
                self.file_type.set(ext)
            self.update_status(f"Selected file: {os.path.basename(filename)}")
    
    def browse_directory(self):
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_dir.set(directory)
            self.update_status(f"Output directory: {directory}")
    
    def clear_form(self):
        self.file_path.set("")
        self.output_dir.set("")
        self.file_type.set("csv")
        self.output_format.set("csv")
        self.split_method.set("equal_parts")
        self.parts_entry.set(2)
        self.rows_entry.set(1000)
        self.add_custom_text.set(False)
        self.custom_text_area.delete(1.0, tk.END)
        self.custom_text_area.insert(tk.END, "Email Header Line 1\nEmail Header Line 2")
        self.toggle_custom_text()
        self.update_status("Ready")
    
    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()
    
    def get_custom_text_lines(self):
        """Get the custom text lines from the text area"""
        if not self.add_custom_text.get():
            return []
        
        text_content = self.custom_text_area.get(1.0, tk.END).strip()
        if not text_content:
            return []
        
        # Split by lines and remove empty lines
        lines = [line for line in text_content.split('\n') if line.strip()]
        return lines
    
    def add_custom_text_to_dataframe(self, df, text_lines):
        """Add custom text lines as new rows at the beginning of the DataFrame"""
        if not text_lines:
            return df
        
        # Create a new DataFrame with the custom text
        custom_data = []
        for line in text_lines:
            # Create a row where all columns contain the custom text
            row_data = {col: line for col in df.columns}
            custom_data.append(row_data)
        
        custom_df = pd.DataFrame(custom_data)
        
        # Concatenate the custom text DataFrame with the original DataFrame
        result_df = pd.concat([custom_df, df], ignore_index=True)
        return result_df
    
    def add_custom_text_to_csv(self, file_path, text_lines):
        """Add custom text lines to a CSV file"""
        if not text_lines:
            return
        
        # Read the original content
        with open(file_path, 'r', encoding='utf-8') as f:
            original_content = f.read()
        
        # Write custom text lines followed by original content
        with open(file_path, 'w', encoding='utf-8') as f:
            for line in text_lines:
                # For CSV, we need to handle the line properly
                # If the line contains commas, we should quote it
                if ',' in line:
                    f.write(f'"{line}"\n')
                else:
                    f.write(f'{line}\n')
            f.write(original_content)
    
    def add_custom_text_to_txt(self, file_path, text_lines):
        """Add custom text lines to a TXT file"""
        if not text_lines:
            return
        
        # Read the original content
        with open(file_path, 'r', encoding='utf-8') as f:
            original_content = f.read()
        
        # Write custom text lines followed by original content
        with open(file_path, 'w', encoding='utf-8') as f:
            for line in text_lines:
                f.write(f'{line}\n')
            f.write(original_content)
    
    def read_txt_file(self, file_path):
        """Read a text file and convert to DataFrame"""
        data = []
        with open(file_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Treat each line as a row with a single column
        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if line:  # Skip empty lines
                data.append({'Line': line, 'Line_Number': line_num})
        
        return pd.DataFrame(data)
    
    def write_txt_file(self, df, output_file):
        """Write DataFrame to text file"""
        with open(output_file, 'w', encoding='utf-8') as f:
            for _, row in df.iterrows():
                # For TXT files, we write the content from the 'Line' column if it exists
                # Otherwise, we join all columns with tab separator
                if 'Line' in df.columns:
                    f.write(f"{row['Line']}\n")
                else:
                    # Join all column values with tab separator
                    line = '\t'.join(str(val) for val in row.values)
                    f.write(f"{line}\n")
    
    def split_data(self):
        input_file = self.file_path.get()
        output_dir = self.output_dir.get()
        
        if not input_file:
            messagebox.showerror("Error", "Please select an input file")
            return
        
        if not output_dir:
            messagebox.showerror("Error", "Please select an output directory")
            return
        
        try:
            file_type = self.file_type.get()
            output_format = self.output_format.get()
            custom_text_lines = self.get_custom_text_lines()
            
            # Read input file based on type
            if file_type == "csv":
                df = pd.read_csv(input_file)
            elif file_type == "xlsx":
                df = pd.read_excel(input_file)
            elif file_type == "xml":
                tree = ET.parse(input_file)
                root = tree.getroot()
                # Simple XML to DataFrame conversion (may need customization)
                data = []
                for child in root:
                    row = {}
                    for subchild in child:
                        row[subchild.tag] = subchild.text
                    data.append(row)
                df = pd.DataFrame(data)
            elif file_type == "txt":
                df = self.read_txt_file(input_file)
            
            total_rows = len(df)
            
            if self.split_method.get() == "equal_parts":
                num_parts = int(self.parts_entry.get())
                rows_per_file = math.ceil(total_rows / num_parts)
            else:
                rows_per_file = int(self.rows_entry.get())
                num_parts = math.ceil(total_rows / rows_per_file)
            
            if rows_per_file < 1:
                messagebox.showerror("Error", "Rows per file must be at least 1")
                return
            
            if num_parts < 1:
                messagebox.showerror("Error", "Number of parts must be at least 1")
                return
            
            # Create output directory if it doesn't exist
            os.makedirs(output_dir, exist_ok=True)
            
            base_filename = os.path.splitext(os.path.basename(input_file))[0]
            
            self.update_status(f"Splitting {total_rows} rows into {num_parts} parts...")
            
            for i in range(num_parts):
                start_idx = i * rows_per_file
                end_idx = min((i + 1) * rows_per_file, total_rows)
                part_df = df.iloc[start_idx:end_idx]
                
                output_file = os.path.join(
                    output_dir,
                    f"{base_filename}_part{i+1}.{output_format}"
                )
                
                # Add custom text if enabled
                if custom_text_lines:
                    if output_format == "csv":
                        # For CSV, we'll write the file first then prepend the custom text
                        part_df.to_csv(output_file, index=False)
                        self.add_custom_text_to_csv(output_file, custom_text_lines)
                    elif output_format == "txt":
                        # For TXT, we'll write the file first then prepend the custom text
                        self.write_txt_file(part_df, output_file)
                        self.add_custom_text_to_txt(output_file, custom_text_lines)
                    elif output_format in ["xlsx", "xml"]:
                        # For Excel and XML, we'll modify the DataFrame first
                        part_df = self.add_custom_text_to_dataframe(part_df, custom_text_lines)
                        
                        if output_format == "xlsx":
                            part_df.to_excel(output_file, index=False)
                        elif output_format == "xml":
                            root = ET.Element("root")
                            for _, row in part_df.iterrows():
                                record = ET.SubElement(root, "record")
                                for col, val in row.items():
                                    ET.SubElement(record, col).text = str(val)
                            tree = ET.ElementTree(root)
                            tree.write(output_file, encoding="utf-8", xml_declaration=True)
                else:
                    # No custom text to add
                    if output_format == "csv":
                        part_df.to_csv(output_file, index=False)
                    elif output_format == "xlsx":
                        part_df.to_excel(output_file, index=False)
                    elif output_format == "xml":
                        root = ET.Element("root")
                        for _, row in part_df.iterrows():
                            record = ET.SubElement(root, "record")
                            for col, val in row.items():
                                ET.SubElement(record, col).text = str(val)
                        tree = ET.ElementTree(root)
                        tree.write(output_file, encoding="utf-8", xml_declaration=True)
                    elif output_format == "txt":
                        self.write_txt_file(part_df, output_file)
                
                self.update_status(f"Created part {i+1}/{num_parts}: {os.path.basename(output_file)}")
            
            custom_text_msg = f" with {len(custom_text_lines)} custom text lines" if custom_text_lines else ""
            messagebox.showinfo("Success", f"Successfully split data into {num_parts} files{custom_text_msg}")
            self.update_status(f"Completed. Split into {num_parts} files in {output_dir}")
        
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.update_status(f"Error: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = DataSplitterApp(root)
    root.mainloop()