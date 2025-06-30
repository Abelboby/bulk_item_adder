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

class ProductTemplateGenerator:
    def __init__(self):
        self.template_file = None
        self.existing_data = {}
        self.brand_codes = []
        self.category_codes = []
        self.tax_codes = []
        
    def select_template_file(self):
        """Open file picker to select the template file"""
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        
        print("Please select your item_template.xlsx file...")
        
        file_path = filedialog.askopenfilename(
            title="Select Item Template File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        root.destroy()
        
        if not file_path:
            print("No file selected. Exiting...")
            return None
            
        self.template_file = file_path
        print(f"Selected file: {os.path.basename(file_path)}")
        return file_path
    
    def load_existing_data(self):
        """Load existing data from all sheets in the template file"""
        try:
            # Read all sheets
            excel_file = pd.ExcelFile(self.template_file)
            required_sheets = ['Item Template', 'Branch Codes', 'Category Codes', 'Brand Codes', 'Tax Codes']
            missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_file.sheet_names]
            if missing_sheets:
                print(f"Missing required sheets: {', '.join(missing_sheets)}")
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
            print(f"Error loading template file: {str(e)}")
            return False
    
    def display_existing_codes(self):
        """Display existing codes to user"""
        if self.brand_codes:
            print(f"Brand Codes ({len(self.brand_codes)}):")
            for i, code in enumerate(self.brand_codes, 1):
                print(f"   {i:2d}. {code}")
        if self.category_codes:
            print(f"Category Codes ({len(self.category_codes)}):")
            for i, code in enumerate(self.category_codes, 1):
                print(f"   {i:2d}. {code}")
        if self.tax_codes:
            print(f"Tax Codes ({len(self.tax_codes)}):")
            for i, code in enumerate(self.tax_codes, 1):
                print(f"   {i:2d}. {code}")
    
    def get_user_selections(self):
        """Get user selections and allow adding new codes"""
        self.display_existing_codes()
        
        # Brand codes selection
        selected_brands = self.select_codes("Brand", self.brand_codes, 'Brand Codes')
        
        # Category codes selection
        selected_categories = self.select_codes("Category", self.category_codes, 'Category Codes')
        
        # Tax codes selection
        selected_tax_codes = self.select_codes("Tax", self.tax_codes, 'Tax Codes')
        
        # Get number of items
        num_items = self.get_number_of_items()
        
        # Barcode options
        barcode_option = self.get_barcode_option()
        
        # Image URL option
        include_images = self.get_image_option()
        
        return selected_brands, selected_categories, selected_tax_codes, num_items, barcode_option, include_images
    
    def select_codes(self, code_type, existing_codes, sheet_name):
        print(f"{code_type} Code Selection:")
        print("1. Use all existing codes")
        print("2. Select specific codes")
        print("3. Add new codes")
        print("4. Mix of existing and new codes")
        while True:
            try:
                choice = int(input(f"Choose option for {code_type} codes (1-4): "))
                if choice in [1, 2, 3, 4]:
                    break
                print("Please enter 1, 2, 3, or 4")
            except ValueError:
                print("Please enter a valid number")
        selected_codes = []
        if choice == 1:
            selected_codes = existing_codes.copy()
        elif choice == 2:
            print(f"Select {code_type.lower()} codes (enter numbers separated by comma):")
            for i, code in enumerate(existing_codes, 1):
                print(f"   {i}. {code}")
            while True:
                try:
                    selections = input("Enter selections (e.g., 1,3,5): ").strip()
                    indices = [int(x.strip()) - 1 for x in selections.split(',')]
                    selected_codes = [existing_codes[i] for i in indices if 0 <= i < len(existing_codes)]
                    if selected_codes:
                        break
                    print("Please select at least one valid code")
                except (ValueError, IndexError):
                    print("Please enter valid numbers separated by commas")
        elif choice == 3:
            new_codes_input = input(f"New {code_type.lower()} codes (comma separated): ").strip()
            new_codes = [code.strip() for code in new_codes_input.split(',') if code.strip()]
            if new_codes:
                selected_codes = new_codes
                self.add_new_codes_to_sheet(new_codes, sheet_name, f"{code_type.lower()}_code")
            else:
                selected_codes = existing_codes.copy()
        elif choice == 4:
            print(f"Select existing {code_type.lower()} codes (numbers comma separated, or 'all'): ")
            for i, code in enumerate(existing_codes, 1):
                print(f"   {i}. {code}")
            selections_input = input("Enter selections: ").strip().lower()
            if selections_input == 'all':
                selected_codes = existing_codes.copy()
            else:
                try:
                    indices = [int(x.strip()) - 1 for x in selections_input.split(',')]
                    selected_codes = [existing_codes[i] for i in indices if 0 <= i < len(existing_codes)]
                except (ValueError, IndexError):
                    selected_codes = []
            new_codes_input = input(f"New {code_type.lower()} codes (comma separated): ").strip()
            new_codes = [code.strip() for code in new_codes_input.split(',') if code.strip()]
            if new_codes:
                selected_codes.extend(new_codes)
                self.add_new_codes_to_sheet(new_codes, sheet_name, f"{code_type.lower()}_code")
        return selected_codes if selected_codes else existing_codes.copy()
    
    def add_new_codes_to_sheet(self, new_codes, sheet_name, code_column):
        """Add new codes to the respective sheet"""
        if sheet_name in self.existing_data:
            df = self.existing_data[sheet_name]
            
            for code in new_codes:
                # Check if code already exists
                if code not in df[code_column].values:
                    new_row = {code_column: code, 'name': f"New {code}"}
                    df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)
            
            self.existing_data[sheet_name] = df
    
    def get_number_of_items(self):
        """Get number of items to generate"""
        while True:
            try:
                num_input = input("\nEnter number of items to generate (default: 1000): ").strip()
                if not num_input:
                    return 1000
                num_items = int(num_input)
                if num_items > 0:
                    return num_items
                print("Please enter a positive number")
            except ValueError:
                print("Please enter a valid number")
    
    def get_barcode_option(self):
        """Get barcode generation option"""
        print("\n🔹 Barcode Options:")
        print("1. All items have barcodes")
        print("2. Only random items have barcodes")
        print("3. No barcodes")
        
        while True:
            try:
                choice = int(input("Choose barcode option (1-3): "))
                if choice in [1, 2, 3]:
                    return choice
                print("Please enter 1, 2, or 3")
            except ValueError:
                print("Please enter a valid number")
    
    def get_image_option(self):
        """Get image URL option"""
        include_images = input("\nInclude random image URLs? (y/n, default: y): ").strip().lower()
        return include_images != 'n'
    
    def generate_random_code(self, prefix="", length=6):
        """Generate a random alphanumeric code"""
        random_part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=length))
        return f"{prefix}{random_part}" if prefix else random_part
    
    def generate_random_hsn(self):
        """Generate a random HSN code (8 digits)"""
        return ''.join(random.choices(string.digits, k=8))
    
    def generate_product_name(self, used_names=None):
        """Generate a unique random product name with only letters, numbers, spaces, and at least one letter"""
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
            # Remove any non-alphanumeric and non-space characters
            name = re.sub(r'[^A-Za-z0-9 ]+', '', name)
            # Ensure at least one letter
            if re.search(r'[A-Za-z]', name) and name not in used_names:
                used_names.add(name)
                return name
        # Fallback: use a default name
        name = "Item " + str(random.randint(100000, 999999))
        used_names.add(name)
        return name
    
    def generate_image_url(self):
        """Generate a random placeholder image URL"""
        width = random.choice([300, 400, 500])
        height = random.choice([300, 400, 500])
        return f"https://picsum.photos/{width}/{height}?random={random.randint(1, 1000)}"
    
    def generate_barcode(self, barcode_option, item_index):
        """Generate barcode based on user option"""
        if barcode_option == 1:  # All items have barcodes
            return self.generate_random_code("BC", 10)
        elif barcode_option == 2:  # Random items have barcodes
            return self.generate_random_code("BC", 10) if random.random() < 0.7 else ""
        else:  # No barcodes
            return ""
    
    def create_product_data(self, brands, categories, tax_codes, num_items, barcode_option, include_images):
        print(f"Generating {num_items} unique product items...")
        data = []
        used_names = set()
        for i in range(num_items):
            if (i + 1) % 100 == 0:
                print(f"   Generated {i + 1}/{num_items} items...")
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
        return pd.DataFrame(data)
    
    def generate_output_filename(self, original_filename, batch_number, total_batches, items_in_batch, total_items):
        """Return output file name as item_template_1.xlsx, item_template_2.xlsx, etc."""
        return f"item_template_{batch_number}.xlsx"
    
    def split_data_into_batches(self, data_df, batch_size=1000):
        """Split data into batches of specified size"""
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
        print(f"Preparing to save {len(new_item_data)} items...")
        try:
            batches = self.split_data_into_batches(new_item_data, batch_size=1000)
            total_batches = len(batches)
            print(f"Creating {total_batches} file(s) (max 1000 items per file)")
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
                # Open and save in Excel to normalize the file
                self.excel_open_and_save(output_path)
                saved_files.append(output_path)
                print(f"   Saved batch {batch_number}/{total_batches}: {output_filename} ({len(batch_data)} items)")
            self.display_summary(new_item_data, saved_files)
        except Exception as e:
            print(f"Error saving file: {str(e)}")
    
    def display_summary(self, item_data, saved_files):
        print(f"Template generation completed successfully!")
        print(f"Summary:")
        print(f"   - Total products generated: {len(item_data)}")
        print(f"   - Unique product names: {len(item_data['name'].unique())}")
        print(f"   - Files created: {len(saved_files)}")
        print(f"   - Brands used: {len(item_data['brand_code'].unique())}")
        print(f"   - Categories used: {len(item_data['category_code'].unique())}")
        print(f"   - Tax codes used: {len(item_data['tax_code'].unique())}")
        print(f"   - Items with barcodes: {len(item_data[item_data['bar_qr_code'] != ''])}")
        print(f"   - Items with image URLs: {len(item_data[item_data['image_url'] != ''])}")
        print(f"Generated Files:")
        for i, file_path in enumerate(saved_files, 1):
            file_size = os.path.getsize(file_path) / 1024  # Size in KB
            print(f"   {i}. {os.path.basename(file_path)} ({file_size:.1f} KB)")
            print(f"      {file_path}")
        print(f"Sample data (first 3 rows):")
        print(item_data.head(3)[['name', 'brand_code', 'category_code', 'tax_code']].to_string(index=False))
    
    def run(self):
        try:
            if not self.select_template_file():
                return
            if not self.load_existing_data():
                return
            brands, categories, tax_codes, num_items, barcode_option, include_images = self.get_user_selections()
            new_item_data = self.create_product_data(brands, categories, tax_codes, num_items, barcode_option, include_images)
            self.save_updated_template(new_item_data)
        except KeyboardInterrupt:
            print("Operation cancelled by user.")
        except Exception as e:
            print(f"An unexpected error occurred: {str(e)}")
            print("Please check your template file and try again.")

def main():
    """Main function"""
    try:
        import pandas as pd
        import openpyxl
        import tkinter
    except ImportError as e:
        print(f"❌ Missing required library: {e}")
        print("Please install required libraries:")
        print("pip install pandas openpyxl faker")
        return
    
    generator = ProductTemplateGenerator()
    generator.run()

if __name__ == "__main__":
    main()