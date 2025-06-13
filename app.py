import re
import json
import csv
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from PyPDF2 import PdfReader

class POExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PO Data Extractor")
        self.root.geometry("600x400")

        # GUI Elements
        tk.Label(root, text="Select PO PDF:").pack(pady=5)
        self.file_entry = tk.Entry(root, width=50)
        self.file_entry.pack(pady=5)
        tk.Button(root, text="Browse", command=self.browse_file).pack(pady=5)
        tk.Button(root, text="Extract Data", command=self.extract_data).pack(pady=20)

        self.result_text = tk.Text(root, height=10, width=70)
        self.result_text.pack(pady=10)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("PDF Files", "*.pdf")],
            title="Select Purchase Order PDF"
        )
        if file_path:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, file_path)

    def extract_data(self):
        pdf_path = self.file_entry.get().strip()
        if not pdf_path or not os.path.isfile(pdf_path):
            messagebox.showerror("Error", "Please select a valid PDF file.")
            return

        try:
            data = self.extract_po_data(pdf_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to extract data: {e}")
            return

        # Display JSON in text widget
        self.result_text.delete(1.0, tk.END)
        json_str = json.dumps(data, indent=2)
        self.result_text.insert(tk.END, json_str)

        # Save JSON and CSV
        base = os.path.splitext(pdf_path)[0]
        json_file = base + "_extracted.json"
        csv_file = base + "_extracted.csv"

        try:
            with open(json_file, 'w') as jf:
                json.dump(data, jf, indent=2)

            self.save_csv(data, csv_file)

            messagebox.showinfo("Success", f"Data saved to:\n{json_file}\n{csv_file}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save files: {e}")

    def extract_po_data(self, pdf_path):
        reader = PdfReader(pdf_path)
        text = "\n".join(page.extract_text() or "" for page in reader.pages)

        header_patterns = {
            'po_number':     r'P\/O No\s*:? *([A-Z0-9/]+)',
            'date':          r'Date\s*:?\s*(\d{2}/\d{2}/\d{4})',
            'supplier':      r'To\s*:?\s*(.*?)(?=\n\S)',
            'delivery_date': r'Delivery Date\s*:?\s*(\d{2}/\d{2}/\d{4})',
            'grand_total':   r'Grand Total\s*:?\s*([\d,]+\.\d{2})',
            'vat':           r'SVAT18\s*18\.00%\s*([\d,]+\.\d{2})'
        }

        result = {}
        for field, pat in header_patterns.items():
            m = re.search(pat, text, re.DOTALL)
            if m:
                val = m.group(1).strip()
                if field == 'supplier':
                    val = ' '.join(val.split())
                result[field] = val

        block_m = re.search(r'Seq\s*No(.*?)Sub Total', text, re.DOTALL)
        products = []
        if block_m:
            raw_block = block_m.group(1)
            row_regex = re.compile(
                r'(\d+)\s+'                  # Seq No
                r'([A-Z0-9/]+)\s+'           # Product Code
                r'(.*?)\s+'                  # Description
                r'([\d,\.]+)\s+NOS\s+'     # Quantity
                r'([\d\.]+)\s+'            # Unit Price
                r'([\d,]+\.\d{2})',        # Total Value
                re.DOTALL
            )
            for m in row_regex.finditer(raw_block):
                _, code, desc, qty, unit_price, total_val = m.groups()
                desc_clean = ' '.join(desc.split())
                products.append({
                    'product_code': code,
                    'description':  desc_clean,
                    'quantity':     qty,
                    'unit_price':   unit_price,
                    'total_value':  total_val
                })

        if products:
            result['products'] = products

        return result

    def save_csv(self, data, csv_path):
        # Prepare CSV columns
        header_fields = ['po_number', 'date', 'supplier', 'delivery_date', 'grand_total', 'vat']
        product_fields = ['product_code', 'description', 'quantity', 'unit_price', 'total_value']
        all_fields = header_fields + product_fields

        with open(csv_path, 'w', newline='') as cf:
            writer = csv.DictWriter(cf, fieldnames=all_fields)
            writer.writeheader()
            products = data.get('products', [])
            if products:
                for prod in products:
                    row = {field: data.get(field, '') for field in header_fields}
                    row.update(prod)
                    writer.writerow(row)
            else:
                # If no products, write just header data
                row = {field: data.get(field, '') for field in header_fields}
                writer.writerow(row)

if __name__ == '__main__':
    root = tk.Tk()
    app = POExtractorApp(root)
    root.mainloop()
