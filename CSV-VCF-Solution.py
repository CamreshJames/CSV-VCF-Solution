import tkinter as tk
from tkinter import filedialog, ttk, messagebox, scrolledtext
import pandas as pd # type: ignore
import os
import re
import csv
from pathlib import Path

class CombinedCSVApp:
    def __init__(self, master):
        self.master = master
        master.title("CSV VCF Solution")
        master.geometry("900x700")
        master.configure(bg="#F0F4F8")  # Light blue-gray background

        # Set the icon (make sure you have this icon file)
        icon_path = "logojmc.png"
        icon_image = tk.PhotoImage(file=icon_path)
        master.wm_iconphoto(True, icon_image)

        self.dfs = []  # List to hold multiple DataFrames for the editor
        self.setup_ui()

    def setup_ui(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.configure_styles()

        # Create a notebook (tabbed interface)
        self.notebook = ttk.Notebook(self.master)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create frames for each tab
        self.editor_frame = ttk.Frame(self.notebook)
        self.merger_frame = ttk.Frame(self.notebook)
        self.maker_merger_frame = ttk.Frame(self.notebook)
        self.converter_frame = ttk.Frame(self.notebook)
        self.how_to_use_frame = ttk.Frame(self.notebook)

        # Add the frames to the notebook
        self.notebook.add(self.editor_frame, text="CSV Editor")
        self.notebook.add(self.merger_frame, text="CSV Merger")
        self.notebook.add(self.maker_merger_frame, text="CSV Maker")
        self.notebook.add(self.converter_frame, text="CSV/VCF Converter")
        self.notebook.add(self.how_to_use_frame, text="How to Use")

        # Setup each tab
        self.setup_editor_tab()
        self.setup_merger_tab()
        self.setup_maker_merger_tab()
        self.setup_converter_tab()
        self.setup_how_to_use_tab()

        # Creator label
        creator_frame = ttk.Frame(self.master)
        creator_frame.pack(fill=tk.X, padx=10, pady=5)
        creator_label = ttk.Label(creator_frame, text="Created by CamreshJames CNJM Technologies INC.", style="Creator.TLabel")
        creator_label.pack(side=tk.RIGHT)

    def configure_styles(self):
        self.style.configure("TButton", padding=10, font=("Segoe UI", 11), background="#4A90E2", foreground="white")
        self.style.map("TButton", background=[('active', '#2980B9')])
        self.style.configure("TLabel", font=("Segoe UI", 11), background="#F0F4F8")
        self.style.configure("TFrame", background="#F0F4F8")
        self.style.configure("Treeview", font=("Segoe UI", 10), rowheight=25)
        self.style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        self.style.configure("TNotebook", background="#F0F4F8")
        self.style.configure("TNotebook.Tab", padding=[10, 5], font=("Segoe UI", 11))
        self.style.configure("Creator.TLabel", font=("Segoe UI", 9), foreground="#555")

    def setup_editor_tab(self):
        # File selection
        ttk.Button(self.editor_frame, text="Load CSV Files", command=self.load_csv_files).pack(pady=10)
        
        # Treeview for displaying CSV contents
        tree_frame = ttk.Frame(self.editor_frame)
        tree_frame.pack(pady=10, padx=10, expand=True, fill=tk.BOTH)

        self.tree = ttk.Treeview(tree_frame, show="headings")
        self.tree.pack(side=tk.LEFT, expand=True, fill=tk.BOTH)

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscroll=scrollbar.set)

        # Name change frame
        change_frame = ttk.Frame(self.editor_frame)
        change_frame.pack(pady=10, padx=10, fill=tk.X)

        ttk.Label(change_frame, text="New Name Prefix:").grid(row=0, column=0, padx=5, pady=5, sticky="e")
        self.name_prefix_entry = ttk.Entry(change_frame, width=30)
        self.name_prefix_entry.grid(row=0, column=1, padx=5, pady=5)

        ttk.Label(change_frame, text="Starting Index:").grid(row=0, column=2, padx=5, pady=5, sticky="e")
        self.start_index_entry = ttk.Entry(change_frame, width=10)
        self.start_index_entry.grid(row=0, column=3, padx=5, pady=5)

        ttk.Button(change_frame, text="Change Names", command=self.change_names).grid(row=0, column=4, padx=5, pady=5)

        # Save button
        ttk.Button(self.editor_frame, text="Save Merged CSV", command=self.save_csv).pack(pady=10)

    def setup_merger_tab(self):
        label_instruction = ttk.Label(self.merger_frame, text="Select CSV files to merge:", style="TLabel")
        label_instruction.pack(pady=(10, 5))

        self.files_text = scrolledtext.ScrolledText(self.merger_frame, wrap=tk.WORD, width=50, height=10)
        self.files_text.pack(pady=(0, 10), padx=10, expand=True, fill=tk.BOTH)

        btn_select = ttk.Button(self.merger_frame, text="Select Files", command=self.select_files_for_merge)
        btn_select.pack(pady=(0, 10))

        btn_merge = ttk.Button(self.merger_frame, text="Merge Files", command=self.merge_files)
        btn_merge.pack(pady=(0, 10))

        self.merger_result_label = ttk.Label(self.merger_frame, text="", wraplength=540, justify="center")
        self.merger_result_label.pack(pady=(10, 0))

    def setup_converter_tab(self):
        self.current_mode = tk.StringVar(value="csv2vcf")

        mode_frame = ttk.Frame(self.converter_frame)
        mode_frame.pack(fill=tk.X, pady=10)
        
        ttk.Radiobutton(mode_frame, text="CSV to VCF", variable=self.current_mode, value="csv2vcf", command=self.update_converter_ui).pack(side=tk.LEFT, padx=10)
        ttk.Radiobutton(mode_frame, text="VCF to CSV", variable=self.current_mode, value="vcf2csv", command=self.update_converter_ui).pack(side=tk.LEFT, padx=10)

        self.converter_label = ttk.Label(self.converter_frame, text="Select CSV files to convert to VCF:", style="TLabel")
        self.converter_label.pack(pady=(10, 5))

        self.btn_select_files = ttk.Button(self.converter_frame, text="Select Files", command=self.select_files_for_conversion)
        self.btn_select_files.pack(pady=(0, 10))

        self.listbox_files = tk.Listbox(self.converter_frame, width=80, height=8)
        self.listbox_files.pack(pady=(0, 10), padx=10, expand=True, fill=tk.BOTH)

        scrollbar = ttk.Scrollbar(self.converter_frame, orient="vertical", command=self.listbox_files.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox_files.config(yscrollcommand=scrollbar.set)

        output_frame = ttk.Frame(self.converter_frame)
        output_frame.pack(fill=tk.X, pady=10)

        ttk.Label(output_frame, text="Output directory:", style="TLabel").pack(side=tk.LEFT, padx=(0, 5))
        self.entry_output_dir = ttk.Entry(output_frame, width=50)
        self.entry_output_dir.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        ttk.Button(output_frame, text="Browse", command=self.select_output_directory).pack(side=tk.RIGHT)

        ttk.Button(self.converter_frame, text="Convert Files", command=self.convert_files).pack(pady=(10, 20))

        self.converter_result_label = ttk.Label(self.converter_frame, text="", wraplength=640, justify="left")
        self.converter_result_label.pack(pady=(10, 0))

    def setup_how_to_use_tab(self):
        text_widget = scrolledtext.ScrolledText(self.how_to_use_frame, wrap=tk.WORD, width=80, height=30, font=("Segoe UI", 11))
        text_widget.pack(padx=10, pady=10, expand=True, fill=tk.BOTH)

        how_to_use_text = """
         How to Use the CSV Editor, Merger, Maker, and Converter Application

 CSV Editor Tab

1. Load CSV Files: 
   - Click 'Load CSV Files' to select and load one or multiple CSV files.
   - The contents will be displayed in a table format for easy viewing.

2. Edit Contents: 
   - View and examine the loaded CSV data in the table.
   - (Note: Direct editing in the table is not supported in the current version)

3. Change Names: 
   - Enter a 'New Name Prefix' in the provided field.
   - Specify the 'Starting Index' for numbering.
   - Click 'Change Names' to apply changes to the first column entries.

4. Save: 
   - Click 'Save Merged CSV' to save your edited and merged CSV file.
   - Choose a location and filename for the output file.

 CSV Merger Tab

1. Select Files: 
   - Click 'Select Files' to choose multiple CSV files for merging.
   - Selected files will be listed in the text area.

2. Merge: 
   - Click 'Merge Files' to combine the selected files into one CSV file.
   - Files are merged in the order they appear in the list.

3. Save: 
   - You'll be prompted to choose a location and filename for the merged file.

 CSV Maker Tab

1. Enter Phone Numbers:
   - Type or paste phone numbers into the text area, one per line.

2. Set Name Prefix:
   - Enter a prefix for contact names in the 'Name Prefix' field.
   - Default is "BET GROUP 1".

3. Set Starting Index:
   - Enter a number in the 'Start Index' field to begin numbering.
   - Default is "1".

4. Convert to CSV:
   - Click 'Convert to CSV' to create a new CSV file with the entered data.
   - Choose a location and filename for the new CSV file.

5. Concatenate (Optional):
   - After creating a CSV, the 'Concatenate with another CSV' button becomes active.
   - Use this to combine the newly created CSV with an existing one.

 CSV/VCF Converter Tab

1. Choose Mode: 
   - Select either 'CSV to VCF' or 'VCF to CSV' conversion mode.

2. Select Files: 
   - Click 'Select Files' to choose files for conversion.
   - Selected files will be listed in the box.

3. Set Output: 
   - Click 'Browse' to choose the output directory for converted files.

4. Convert: 
   - Click 'Convert Files' to perform the conversion.
   - Converted files will be saved in the specified output directory.

 General Tips

- Ensure consistent headers in CSV files for accurate merging and conversion.
- For large files, the application might take a moment to process. Please be patient.
- Always verify the output after conversion or merging.
- When using the CSV Maker, ensure phone numbers are in a consistent format.
- The CSV Editor supports viewing multiple files at once, but saves them as a single merged file.
- In the Converter tab, make sure the input CSV files have 'Name' and 'Phone' columns for VCF conversion.
- Regularly save your work, especially when dealing with large datasets.

 If you encounter any issues or have suggestions for improvement, please let us know through:
    camreshjames@gmail.com or https://cnjmtechnologies.com/

 yours faithfully:
     Camresh James CNJM Technologies INC.
            https://cnjmtechnologies.com/
        """

        text_widget.insert(tk.END, how_to_use_text)
        text_widget.config(state=tk.DISABLED)

    def setup_maker_merger_tab(self):
        label_instruction = ttk.Label(self.maker_merger_frame, text="Enter phone numbers (one per line):", style="TLabel")
        label_instruction.pack(pady=(10, 5))

        self.phone_numbers_text = scrolledtext.ScrolledText(self.maker_merger_frame, wrap=tk.WORD, width=50, height=8)
        self.phone_numbers_text.pack(pady=(0, 10), padx=10, expand=True, fill=tk.BOTH)

        name_frame = ttk.Frame(self.maker_merger_frame)
        name_frame.pack(fill=tk.X, pady=5)

        ttk.Label(name_frame, text="Name Prefix:", style="TLabel").pack(side=tk.LEFT, padx=(0, 5))
        self.name_prefix_entry = ttk.Entry(name_frame, width=30)
        self.name_prefix_entry.pack(side=tk.LEFT, expand=True, fill=tk.X)
        self.name_prefix_entry.insert(0, "BET GROUP 1")

        index_frame = ttk.Frame(self.maker_merger_frame)
        index_frame.pack(fill=tk.X, pady=5)

        ttk.Label(index_frame, text="Start Index:", style="TLabel").pack(side=tk.LEFT, padx=(0, 5))
        self.start_index_entry = ttk.Entry(index_frame, width=10)
        self.start_index_entry.pack(side=tk.LEFT)
        self.start_index_entry.insert(0, "1")

        btn_convert = ttk.Button(self.maker_merger_frame, text="Convert to CSV", command=self.convert_to_csv)
        btn_convert.pack(pady=(10, 5))

        self.btn_concatenate = ttk.Button(self.maker_merger_frame, text="Concatenate with another CSV", command=self.concatenate_csv, state=tk.DISABLED)
        self.btn_concatenate.pack(pady=5)

        self.maker_merger_result_label = ttk.Label(self.maker_merger_frame, text="", wraplength=540, justify="center")
        self.maker_merger_result_label.pack(pady=(10, 0))

    def convert_to_csv(self):
        phone_numbers = self.phone_numbers_text.get("1.0", tk.END).strip().split('\n')
        name_prefix = self.name_prefix_entry.get().strip()
        start_index = int(self.start_index_entry.get())

        # Prepare the data
        data = []
        for i, phone in enumerate(phone_numbers):
            name = f"{name_prefix} {str(start_index + i).zfill(3)}"
            data.append({"name": name, "phone": phone.strip()})

        # Convert to DataFrame
        df = pd.DataFrame(data)

        # Save to CSV
        output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if output_path:
            df.to_csv(output_path, index=False)
            self.update_merger_result_maker(f"CSV file saved successfully as:\n{output_path}", "success")
            
            # Enable concatenate button after creating the first CSV
            self.btn_concatenate.config(state=tk.NORMAL)
        else:
            self.update_merger_result_maker("CSV file was not saved.", "error")

    def concatenate_csv(self):
        # Ask user to select the second CSV file
        second_csv_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if not second_csv_path:
            self.update_merger_result_maker("No file selected for concatenation.", "error")
            return

        try:
            # Read the second CSV
            df_second = pd.read_csv(second_csv_path)

            # Read the first CSV (the one we just created)
            first_csv_path = self.merger_result_label.cget("text").split("\n")[-1]
            df_first = pd.read_csv(first_csv_path)

            # Concatenate the DataFrames
            df_concatenated = pd.concat([df_first, df_second], ignore_index=True)

            # Save the concatenated DataFrame
            output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
            if output_path:
                df_concatenated.to_csv(output_path, index=False)
                self.update_merger_result_maker(f"Concatenated CSV file saved successfully as:\n{output_path}", "success")
            else:
                self.update_merger_result_maker("Concatenated CSV file was not saved.", "error")

        except Exception as e:
            self.update_merger_result_maker(f"Error during concatenation: {str(e)}", "error")

    def update_merger_result_maker(self, message, status):
        self.maker_merger_result_label.config(text=message)
        if status == "error":
            self.maker_merger_result_label.configure(style="Error.TLabel")
        elif status == "success":
            self.maker_merger_result_label.configure(style="Success.TLabel")

    # Converter methods
    def select_files_for_conversion(self):
        file_type = "CSV" if self.current_mode.get() == "csv2vcf" else "VCF"
        file_extension = "*.csv" if self.current_mode.get() == "csv2vcf" else "*.vcf"
        file_paths = filedialog.askopenfilenames(filetypes=[(f"{file_type} Files", file_extension)])
        if file_paths:
            self.listbox_files.delete(0, tk.END)
            for file_path in file_paths:
                self.listbox_files.insert(tk.END, file_path)

    def select_output_directory(self):
        directory = filedialog.askdirectory()
        if directory:
            self.entry_output_dir.delete(0, tk.END)
            self.entry_output_dir.insert(0, directory)

    def convert_files(self):
        output_dir = self.entry_output_dir.get()
        if not output_dir:
            messagebox.showwarning("Warning", "Please select an output directory first.")
            return

        file_paths = self.listbox_files.get(0, tk.END)
        if not file_paths:
            messagebox.showwarning("Warning", "Please select files to convert first.")
            return

        try:
            converted_files = []
            for file_path in file_paths:
                if self.current_mode.get() == "csv2vcf":
                    output_path = self.convert_csv_to_vcf(file_path, output_dir)
                else:
                    output_path = self.convert_vcf_to_csv(file_path, output_dir)
                converted_files.append(output_path)
            
            self.converter_result_label.config(text=f"Conversion successful!\nFiles saved in:\n{output_dir}\n\nConverted files:\n" + "\n".join(str(f) for f in converted_files))
        except Exception as e:
            messagebox.showerror("Error", f"Error during conversion: {str(e)}")

    def update_converter_ui(self):
        if self.current_mode.get() == "csv2vcf":
            self.converter_label.config(text="Select CSV files to convert to VCF:")
            self.btn_select_files.config(text="Select CSV Files")
        else:
            self.converter_label.config(text="Select VCF files to convert to CSV:")
            self.btn_select_files.config(text="Select VCF Files")

    def create_vcard(self, name, phone_numbers):
        vcard = (
            "BEGIN:VCARD\n"
            "VERSION:2.1\n"
            f"N:;{name};;;\n"
            f"FN:{name}\n"
        )
        for phone in phone_numbers:
            vcard += f"TEL;CELL:{phone}\n"
        vcard += "END:VCARD\n"
        return vcard

    def convert_csv_to_vcf(self, csv_path, output_dir):
        contacts = []
        with open(csv_path, mode='r') as file:
            csv_reader = csv.reader(file)
            next(csv_reader)  # Skip header row
            
            for row in csv_reader:
                if len(row) >= 2:
                    name = row[0].strip()
                    phone_numbers = [num.strip() for num in row[1:] if num.strip()]
                    contacts.append((name, phone_numbers))
        
        vcards = [self.create_vcard(name, phone_numbers) for name, phone_numbers in contacts]
        
        output_path = Path(output_dir) / Path(csv_path).with_suffix('.vcf').name
        with open(output_path, mode='w') as file:
            file.writelines(vcards)
        
        return output_path

    def convert_vcf_to_csv(self, vcf_path, output_dir):
            contacts = []
            with open(vcf_path, mode='r') as file:
                lines = file.readlines()
                contact = {}
                for line in lines:
                    if line.startswith("FN:"):
                        contact["name"] = line.split(":", 1)[1].strip()
                    elif line.startswith("TEL;"):
                        if "phone" not in contact:
                            contact["phone"] = []
                        contact["phone"].append(line.split(":", 1)[1].strip())
                    elif line.startswith("END:VCARD"):
                        if "name" in contact and "phone" in contact:
                            contacts.append((contact["name"], contact["phone"]))
                        contact = {}

            output_path = Path(output_dir) / Path(vcf_path).with_suffix('.csv').name
            with open(output_path, mode='w', newline='') as file:
                csv_writer = csv.writer(file)
                csv_writer.writerow(["Name", "Phone Number 1", "Phone Number 2"])
                for contact in contacts:
                    row = [contact[0]] + contact[1][:2] + [''] * (2 - len(contact[1]))
                    csv_writer.writerow(row)

            return output_path

    # CSV Editor methods
    def load_csv_files(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if file_paths:
            self.dfs = []
            for file_path in file_paths:
                try:
                    df = pd.read_csv(file_path)
                    self.dfs.append(df)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to load {file_path}: {str(e)}")
            
            if self.dfs:
                self.update_treeview()

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        
        if not self.dfs:
            return

        columns = list(self.dfs[0].columns)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for df in self.dfs:
            for index, row in df.iterrows():
                self.tree.insert("", "end", values=list(row))

    def change_names(self):
        if not self.dfs:
            messagebox.showwarning("Warning", "Please load CSV files first.")
            return

        prefix = self.name_prefix_entry.get()
        try:
            start_index = int(self.start_index_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Starting Index must be a number.")
            return

        name_column = self.dfs[0].columns[0]
        total_rows = sum(len(df) for df in self.dfs)

        new_names = [f"{prefix} {str(i).zfill(3)}" for i in range(start_index, start_index + total_rows)]

        current_index = 0
        for df in self.dfs:
            df_length = len(df)
            df[name_column] = new_names[current_index:current_index + df_length]
            current_index += df_length

        self.update_treeview()

    def save_csv(self):
        if not self.dfs:
            messagebox.showwarning("Warning", "No data to save. Please load and edit CSV files first.")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
        if file_path:
            merged_df = pd.concat(self.dfs, ignore_index=True)
            merged_df.to_csv(file_path, index=False)
            messagebox.showinfo("Success", f"Merged file saved successfully to {file_path}")

    # CSV Merger methods
    def select_files_for_merge(self):
        file_paths = filedialog.askopenfilenames(filetypes=[("CSV Files", "*.csv")])
        if file_paths:
            self.files_text.delete('1.0', tk.END)
            for file in file_paths:
                self.files_text.insert(tk.END, file + '\n')

    def merge_files(self):
        file_paths = self.files_text.get('1.0', tk.END).strip().split('\n')
        if not file_paths:
            self.update_merger_result("No files selected for merging.", "error")
            return

        try:
            sorted_files = sorted(file_paths, key=lambda x: int(re.findall(r'\d+', os.path.basename(x))[-1]))
            df_list = [pd.read_csv(file) for file in sorted_files]
            df_merged = pd.concat(df_list, ignore_index=True)

            output_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV Files", "*.csv")])
            if output_path:
                df_merged.to_csv(output_path, index=False)
                self.update_merger_result(f"Merged CSV file saved successfully as:\n{output_path}", "success")
            else:
                self.update_merger_result("Merged CSV file was not saved.", "error")
        except Exception as e:
            self.update_merger_result(f"Error during merging: {str(e)}", "error")

    def update_merger_result(self, message, status):
        self.merger_result_label.config(text=message)
        if status == "error":
            self.merger_result_label.configure(foreground="#D32F2F")
        elif status == "success":
            self.merger_result_label.configure(foreground="#388E3C")

if __name__ == "__main__":
    root = tk.Tk()
    app = CombinedCSVApp(root)
    root.mainloop()