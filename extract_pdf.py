import fitz
import re
from openpyxl import Workbook

# Open the PDF
doc = fitz.open("Angebot_36227159-001.pdf")

lines = []
for page in doc:
    lines.extend(page.get_text().split("\n"))

# Prepare list to hold extracted rows
items = []
i = 0

while i < len(lines):
    line = lines[i]
    if line.startswith("POS"):
        try:
            pos = re.search(r"POS\s+(\d+)", line).group(1)
            article_code = lines[i + 3].strip()
            desc_lines = [lines[i + 4].strip()]

            # Quantity & unit price
            qty_price_line = lines[i + 5]
            total_price_line = lines[i + 6]
            qty_match = re.search(r"(\d+,\d{3})\s+ST\s+([\d,]+)", qty_price_line)
            if qty_match:
                quantity = qty_match.group(1)
                unit_price = qty_match.group(2)
            else:
                quantity = ""
                unit_price = ""

            total_price = re.search(r"([\d,]+)", total_price_line).group(1)

            # Additional description lines
            j = i + 7
            while j < len(lines) and not lines[j].startswith("POS"):
                desc_lines.append(lines[j].strip())
                j += 1

            full_description = " ".join(desc_lines)
            words = full_description.split()
            short_description = " ".join(words[:10])

            items.append((pos, article_code, short_description, quantity, unit_price, total_price))
            i = j
        except Exception as e:
            print(f"Skipping line {i}: {e}")
            i += 1
    else:
        i += 1

# Export to Excel
wb = Workbook()
ws = wb.active
ws.title = "Extracted Items"
ws.append(["POS", "Article Code", "Description", "Quantity", "Unit Price", "Total Price"])

for item in items:
    ws.append(item)

xlsx_filename = "extracted_items.xlsx"
wb.save(xlsx_filename)
print(f"✅ Data exported to '{xlsx_filename}' with short descriptions.")

