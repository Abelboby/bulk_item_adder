import pandas as pd
import random
import string
from faker import Faker
import os
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk
import re
import openpyxl
import win32com.client

fake = Faker()

UNIT_CHOICES = [
    ('Millimeter', 'Millimeter'),
    ('Centimeter', 'Centimeter'),
    ('Meter', 'Meter'),
    ('Kilometer', 'Kilometer'),
    ('Milligram', 'Milligram'),
    ('Gram', 'Gram'),
    ('Kilogram', 'Kilogram'),
    ('Ton', 'Ton'),
    ('Millimeter Square', 'Millimeter Square'),
    ('Centimeter Square', 'Centimeter Square'),
    ('Meter Square', 'Meter Square'),
    ('Kilometer Square', 'Kilometer Square'),
    ('Milliliter', 'Milliliter'),
    ('Liter', 'Liter'),
    ('Piece', 'Piece'),
    ('Box', 'Box'),
    ('Bag', 'Bag'),
    ('Set', 'Set')
]

class ProductTemplateGeneratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bulk Item Adder Wizard")
        self.root.geometry("600x400")
        self.template_file = None
        self.existing_data = {}
        self.brand_codes = []
        self.category_codes = []
        self.tax_codes = []
        self.selected_brands = []
        self.selected_categories = []
        self.selected_tax_codes = []
        self.num_items = 1000
        self.barcode_option = 1
        self.include_images = True
        self.progress = None
        self.progress_var = tk.DoubleVar()
        self.init_welcome()

    def clear_frame(self):
        for widget in self.root.winfo_children():
            widget.destroy()

    def init_welcome(self):
        self.clear_frame()
        tk.Label(self.root, text="Bulk Item Adder", font=("Arial", 20)).pack(pady=30)
        tk.Label(self.root, text="Welcome! This wizard will help you generate bulk item Excel files.").pack(pady=10)
        tk.Button(self.root, text="Start", command=self.init_file_select, width=20).pack(pady=30)

    def init_file_select(self):
        self.clear_frame()
        tk.Label(self.root, text="Step 1: Select Item Template File", font=("Arial", 14)).pack(pady=20)
        tk.Button(self.root, text="Choose File", command=self.select_template_file).pack(pady=10)
        self.file_label = tk.Label(self.root, text="No file selected.")
        self.file_label.pack(pady=10)
        tk.Button(self.root, text="Back", command=self.init_welcome).pack(side=tk.LEFT, padx=20, pady=20)

    def select_template_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Item Template File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return
        self.template_file = file_path
        self.file_label.config(text=os.path.basename(file_path))
        if self.load_existing_data():
            self.root.after(500, self.init_code_selection)

    def load_existing_data(self):
        try:
            excel_file = pd.ExcelFile(self.template_file)
            required_sheets = ['Item Template', 'Branch Codes', 'Category Codes', 'Brand Codes', 'Tax Codes']
            missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
            if missing_sheets:
                messagebox.showerror("Missing Sheets", f"Missing required sheets: {', '.join(missing_sheets)}")
                return False
            for sheet_name in excel_file.sheet_names:
                self.existing_data[sheet_name] = pd.read_excel(self.template_file, sheet_name=sheet_name)
            if 'Category Codes' in self.existing_data:
                cat_df = self.existing_data['Category Codes']
                if 'category_code' in cat_df.columns:
                    self.category_codes = cat_df['category_code'].dropna().tolist()
            if 'Brand Codes' in self.existing_data:
                brand_df = self.existing_data['Brand Codes']
                if 'brand_code' in brand_df.columns:
                    self.brand_codes = brand_df['brand_code'].dropna().tolist()
            if 'Tax Codes' in self.existing_data:
                tax_df = self.existing_data['Tax Codes']
                if 'tax_code' in tax_df.columns:
                    self.tax_codes = tax_df['tax_code'].dropna().tolist()
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Error loading template file: {str(e)}")
            return False

    def init_code_selection(self):
        self.clear_frame()
        tk.Label(self.root, text="Step 2: Select Codes", font=("Arial", 14)).pack(pady=10)
        # Brand codes
        tk.Label(self.root, text="Select Brand Codes (Ctrl+Click for multi-select):").pack()
        self.brand_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, exportselection=0, height=6)
        for code in self.brand_codes:
            self.brand_listbox.insert(tk.END, code)
        self.brand_listbox.pack(pady=2)
        self.brand_new_entry = tk.Entry(self.root)
        self.brand_new_entry.pack()
        tk.Button(self.root, text="Add Brand Code", command=self.add_brand_code).pack(pady=2)
        # Category codes
        tk.Label(self.root, text="Select Category Codes (Ctrl+Click for multi-select):").pack()
        self.category_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, exportselection=0, height=6)
        for code in self.category_codes:
            self.category_listbox.insert(tk.END, code)
        self.category_listbox.pack(pady=2)
        self.category_new_entry = tk.Entry(self.root)
        self.category_new_entry.pack()
        tk.Button(self.root, text="Add Category Code", command=self.add_category_code).pack(pady=2)
        # Tax codes
        tk.Label(self.root, text="Select Tax Codes (Ctrl+Click for multi-select):").pack()
        self.tax_listbox = tk.Listbox(self.root, selectmode=tk.MULTIPLE, exportselection=0, height=6)
        for code in self.tax_codes:
            self.tax_listbox.insert(tk.END, code)
        self.tax_listbox.pack(pady=2)
        self.tax_new_entry = tk.Entry(self.root)
        self.tax_new_entry.pack()
        tk.Button(self.root, text="Add Tax Code", command=self.add_tax_code).pack(pady=2)
        tk.Button(self.root, text="Next", command=self.init_options).pack(pady=10)
        tk.Button(self.root, text="Back", command=self.init_file_select).pack(side=tk.LEFT, padx=20, pady=10)

    def add_brand_code(self):
        code = self.brand_new_entry.get().strip()
        if code and code not in self.brand_codes:
            self.brand_codes.append(code)
            self.brand_listbox.insert(tk.END, code)
            self.brand_new_entry.delete(0, tk.END)

    def add_category_code(self):
        code = self.category_new_entry.get().strip()
        if code and code not in self.category_codes:
            self.category_codes.append(code)
            self.category_listbox.insert(tk.END, code)
            self.category_new_entry.delete(0, tk.END)

    def add_tax_code(self):
        code = self.tax_new_entry.get().strip()
        if code and code not in self.tax_codes:
            self.tax_codes.append(code)
            self.tax_listbox.insert(tk.END, code)
            self.tax_new_entry.delete(0, tk.END)

    def init_options(self):
        self.selected_brands = [self.brand_codes[i] for i in self.brand_listbox.curselection()] or self.brand_codes
        self.selected_categories = [self.category_codes[i] for i in self.category_listbox.curselection()] or self.category_codes
        self.selected_tax_codes = [self.tax_codes[i] for i in self.tax_listbox.curselection()] or self.tax_codes
        self.clear_frame()
        tk.Label(self.root, text="Step 3: Options", font=("Arial", 14)).pack(pady=10)
        # Number of items
        tk.Label(self.root, text="Number of items to generate:").pack()
        self.num_items_var = tk.StringVar(value="1000")
        tk.Entry(self.root, textvariable=self.num_items_var).pack(pady=2)
        # Barcode options
        tk.Label(self.root, text="Barcode Options:").pack()
        self.barcode_var = tk.IntVar(value=1)
        tk.Radiobutton(self.root, text="All items have barcodes", variable=self.barcode_var, value=1).pack(anchor=tk.W)
        tk.Radiobutton(self.root, text="Only random items have barcodes", variable=self.barcode_var, value=2).pack(anchor=tk.W)
        tk.Radiobutton(self.root, text="No barcodes", variable=self.barcode_var, value=3).pack(anchor=tk.W)
        # Image option
        self.image_var = tk.BooleanVar(value=True)
        tk.Checkbutton(self.root, text="Include random image URLs", variable=self.image_var).pack(anchor=tk.W)
        tk.Button(self.root, text="Generate", command=self.start_generation).pack(pady=10)
        tk.Button(self.root, text="Back", command=self.init_code_selection).pack(side=tk.LEFT, padx=20, pady=10)

    def start_generation(self):
        try:
            self.num_items = int(self.num_items_var.get())
            if self.num_items <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Invalid Input", "Please enter a valid positive number for items.")
            return
        self.barcode_option = self.barcode_var.get()
        self.include_images = self.image_var.get()
        self.init_progress()
        self.root.after(100, self.generate_items)

    def init_progress(self):
        self.clear_frame()
        tk.Label(self.root, text="Generating Items...", font=("Arial", 14)).pack(pady=20)
        self.progress = ttk.Progressbar(self.root, variable=self.progress_var, maximum=self.num_items)
        self.progress.pack(fill=tk.X, padx=40, pady=30)
        self.progress_var.set(0)

    def generate_items(self):
        try:
            new_item_data = self.create_product_data(
                self.selected_brands, self.selected_categories, self.selected_tax_codes,
                self.num_items, self.barcode_option, self.include_images
            )
            self.save_updated_template(new_item_data)
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

    def generate_product_name(self, used_names=None):
        if used_names is None:
            used_names = set()
        adjectives = ['Premium', 'Deluxe', 'Standard', 'Professional', 'Classic', 'Modern', 'Advanced', 'Basic', 
                     'Elite', 'Superior', 'Ultimate', 'Enhanced', 'Optimized', 'Innovative', 'Smart', 'Precision',
                     'HeavyDuty', 'Compact', 'Portable', 'Industrial', 'Commercial', 'Residential', 'Universal',
                     'HighPerformance', 'EnergyEfficient', 'EcoFriendly', 'Waterproof', 'Durable', 'Lightweight']
        nouns = ['Widget', 'Device', 'Component', 'Tool', 'Product', 'Item', 'Gadget', 'Equipment', 'System',
                'Apparatus', 'Instrument', 'Mechanism', 'Assembly', 'Unit', 'Module', 'Controller', 'Sensor',
                'Adapter', 'Connector', 'Terminal', 'Switch', 'Generator', 'Processor', 'Monitor', 'Display',
                'Interface', 'Hub', 'Router', 'Converter', 'Transformer', 'Regulator', 'Filter', 'Amplifier']
        categories = ['Pro', 'Max', 'Plus', 'Lite', 'X', 'XL', 'Mini', 'Micro', 'Ultra', 'Super', 'Turbo', 'Neo',
                     'Edge', 'Core', 'Prime', 'Elite', 'Force', 'Rapid', 'Swift', 'Quantum', 'Digital', 'Smart']
        max_attempts = 1000
        for attempt in range(max_attempts):
            adjective = random.choice(adjectives)
            noun = random.choice(nouns)
            number = str(random.randint(100, 9999))
            category = random.choice(categories) if random.random() < 0.4 else ""
            if category:
                name = f"{adjective} {noun} {category} {number}"
            else:
                name = f"{adjective} {noun} {number}"
            name = re.sub(r'[^A-Za-z0-9 ]+', '', name)
            if re.search(r'[A-Za-z]', name) and name not in used_names:
                used_names.add(name)
                return name
        name = "Item " + str(random.randint(100000, 999999))
        used_names.add(name)
        return name

    def create_product_data(self, brands, categories, tax_codes, num_items, barcode_option, include_images):
        data = []
        used_names = set()
        for i in range(num_items):
            if (i + 1) % 10 == 0:
                self.progress_var.set(i + 1)
                self.root.update_idletasks()
            product_name = self.generate_product_name(used_names)
            item_data = {
                'name': product_name,
                'description': f"High-quality {product_name.lower()} with excellent features and durability.",
                'bar_qr_code': self.generate_barcode(barcode_option, i),
                'brand_code': random.choice(brands),
                'category_code': random.choice(categories),
                'image_url': self.generate_image_url() if include_images else "",
                'tax_code': random.choice(tax_codes),
                'hsn_code': self.generate_random_hsn(),
                'unit': random.choice(UNIT_CHOICES)[0]
            }
            data.append(item_data)
        self.progress_var.set(num_items)
        self.root.update_idletasks()
        return pd.DataFrame(data)

    def generate_barcode(self, barcode_option, item_index):
        if barcode_option == 1:
            return self.generate_random_code("BC", 10)
        elif barcode_option == 2:
            return self.generate_random_code("BC", 10) if random.random() < 0.7 else ""
        else:
            return ""

    def generate_random_code(self, prefix="", length=6):
        random_part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))
        return f"{prefix}{random_part}" if prefix else random_part

    def generate_random_hsn(self):
        return ''.join(random.choices(string.digits, k=8))

    def generate_image_url(self):
        width = random.choice([300, 400, 500])
        height = random.choice([300, 400, 500])
        return f"https://picsum.photos/{width}/{height}?random={random.randint(1, 1000)}"

    def generate_output_filename(self, original_filename, batch_number, total_batches, items_in_batch, total_items):
        return f"item_template_{batch_number}.xlsx"

    def split_data_into_batches(self, data_df, batch_size=1000):
        batches = []
        for i in range(0, len(data_df), batch_size):
            batch = data_df.iloc[i:i + batch_size].copy()
            batches.append(batch)
        return batches

    def excel_open_and_save(self, filepath):
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(filepath)
        wb.Save()
        wb.Close()
        excel.Quit()

    def save_updated_template(self, new_item_data):
        try:
            batches = self.split_data_into_batches(new_item_data, batch_size=1000)
            total_batches = len(batches)
            saved_files = []
            for batch_number, batch_data in enumerate(batches, 1):
                output_filename = self.generate_output_filename(
                    self.template_file, 
                    batch_number, 
                    total_batches, 
                    len(batch_data), 
                    len(new_item_data)
                )
                original_dir = os.path.dirname(self.template_file)
                output_path = os.path.join(original_dir, output_filename)
                updated_data = self.existing_data.copy()
                updated_data['Item Template'] = batch_data
                with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
                    for sheet_name, df in updated_data.items():
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                self.excel_open_and_save(output_path)
                saved_files.append(output_path)
            self.show_summary(new_item_data, saved_files)
        except Exception as e:
            messagebox.showerror("Error", f"Error saving file: {str(e)}")

    def show_summary(self, item_data, saved_files):
        self.clear_frame()
        tk.Label(self.root, text="Generation Completed!", font=("Arial", 16)).pack(pady=20)
        tk.Label(self.root, text=f"Total products generated: {len(item_data)}").pack()
        tk.Label(self.root, text=f"Files created: {len(saved_files)}").pack()
        tk.Label(self.root, text=f"Brands used: {len(item_data['brand_code'].unique())}").pack()
        tk.Label(self.root, text=f"Categories used: {len(item_data['category_code'].unique())}").pack()
        tk.Label(self.root, text=f"Tax codes used: {len(item_data['tax_code'].unique())}").pack()
        tk.Label(self.root, text=f"Items with barcodes: {len(item_data[item_data['bar_qr_code'] != ''])}").pack()
        tk.Label(self.root, text=f"Items with image URLs: {len(item_data[item_data['image_url'] != ''])}").pack()
        tk.Label(self.root, text="Generated Files:").pack(pady=5)
        for i, file_path in enumerate(saved_files, 1):
            tk.Label(self.root, text=f"{i}. {os.path.basename(file_path)}").pack()
        tk.Button(self.root, text="Finish", command=self.root.quit).pack(pady=20)


def main():
    root = tk.Tk()
    app = ProductTemplateGeneratorGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()