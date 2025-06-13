import re
import json
from PyPDF2 import PdfReader

def extract_po_data(pdf_path):
    reader = PdfReader(pdf_path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)

    header_patterns = {
        'po_number':     r'P\/O No\s*:? *([A-Z0-9/]+)',
        'date':          r'Date\s*:\s*(\d{2}/\d{2}/\d{4})',
        'supplier':      r'To\s*:\s*(.*?)(?=\n\S)',   
        'delivery_date': r'Delivery Date\s*:\s*(\d{2}/\d{2}/\d{4})',
        'grand_total':   r'Grand Total\s*([\d,]+\.\d{2})',
        'vat':           r'SVAT18\s*18\.00%\s*([\d,]+\.\d{2})'
    }

    result = {}
    for field, pat in header_patterns.items():
        m = re.search(pat, text, re.DOTALL)
        if not m:
            continue
        val = m.group(1).strip()
        if field == 'supplier':
            
            val = ' '.join(val.split())
        result[field] = val

    block_m = re.search(r'Seq\s*No(.*?)Sub Total', text, re.DOTALL)
    products = []
    if block_m:
        raw_block = block_m.group(1)

        
        row_regex = re.compile(
            r'(\d+)\s+'                  # 1) Seq No
            r'([A-Z0-9/]+)\s+'           # 2) Product Code
            r'(.*?)\s+'                  # 3) Description (multiline)
            r'([\d,\.]+)\s+NOS\s+'       # 4) Quantity (allows decimals)
            r'([\d\.]+)\s+'              # 5) Unit Price USD
            r'([\d,]+\.\d{2})',          # 6) Total Value USD
            re.DOTALL
        )

        for m in row_regex.finditer(raw_block):
            seq, code, desc, qty, unit_price, total_val = m.groups()
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

if __name__ == '__main__':
    pdf_path = "PO 2712.pdf"
    data = extract_po_data(pdf_path)
    print("Extracted Data:")
    print(json.dumps(data, indent=2))

    with open("po_extracted.json", "w") as f:
        json.dump(data, f, indent=2)
