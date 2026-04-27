import re
import pdfplumber
from django.http import HttpResponse
from django.shortcuts import render
from django.core.files.storage import FileSystemStorage
from openpyxl import Workbook
import os

#Comma and Fullstop for Celine
def swap_separators(value):
    if not value:
        return value
    return value.replace(".", "TEMP").replace(",", ".").replace("TEMP", ",")

#HSCode Extract for Celine
def extract_all_hs_codes(pdf_path):
    HS_X0 = 574  # adjust as per your PDF
    HS_X1 = 614
    hs_codes = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()

            # Find HS Code header
            hs_header_y = None
            for i, word in enumerate(words):
                text = word['text'].strip().upper()
                if text == "HS" and i + 1 < len(words):
                    next_word = words[i + 1]['text'].strip().upper()
                    if next_word == "CODE":
                        hs_header_y = max(word['bottom'], words[i + 1]['bottom'])
                        break
            if hs_header_y is None:
                continue  # no header on this page

            # Collect HS codes below header within column
            for word in words:
                x0, y0 = word['x0'], word['top']
                text = word['text'].strip()
                if y0 > hs_header_y and HS_X0 <= x0 <= HS_X1 and re.fullmatch(r"\d{6,8}", text):
                    hs_codes.append({
                        "page": page_number,
                        "hs_code": text,
                        "top": y0,
                        "bottom": word['bottom']
                    })

    return hs_codes


#Extract Invoice For Celine
def extract_invoice_numbers(pdf_path):
    # Rectangle coordinates for invoice number area
    INV_X0 = 600
    INV_X1 = 1000
    INV_Y0 = 50
    INV_Y1 = 100

    invoice_numbers = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            words = page.extract_words()

            # Keep only words inside the rectangle
            rect_words = [w for w in words if INV_X0 <= w['x0'] <= INV_X1 and INV_Y0 <= w['top'] <= INV_Y1]

            if not rect_words:
                continue

            # Sort by top, then left to form lines
            rect_words.sort(key=lambda w: (w['top'], w['x0']))

            # Combine words into a single line
            line_text = " ".join(w['text'] for w in rect_words).strip()

            # Look for "DOCUMENT N° :" and capture what comes after
            match = re.search(r"DOCUMENT\s*N°?\s*:?(.+)", line_text, re.IGNORECASE)
            if match:
                invoice_number = match.group(1).strip()  # this is "1164466 RJ"
                invoice_numbers.append({
                    "page": page_number,
                    "invoice_number": invoice_number,
                    "top": rect_words[0]['top'],
                    "bottom": rect_words[-1]['bottom']
                })

    return invoice_numbers



#Common DropDown


def extract_pdf_data(request):
    if request.method == "POST" and request.FILES.get("pdf"):
        mode = request.POST.get("mode")
        pdf_file = request.FILES["pdf"]

        # ===============================
        # SAVE PDF TO SERVER
        # ===============================
        fs = FileSystemStorage(location='media/pdfs')
        os.makedirs(fs.location, exist_ok=True)
        saved_pdf_name = fs.save(pdf_file.name, pdf_file)
        saved_pdf_path = fs.path(saved_pdf_name)

        # ===============================
        # EXTRACT FULL TEXT FROM PDF
        # ===============================
        full_text = ""
        with pdfplumber.open(saved_pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

        extracted_data = []

        # ===============================
        # HAKKO CORPORATION MODE
        # ===============================
        if mode == "Hakko_Corporation":
            blocks = re.split(r"\*FREE OF CHARGE", full_text, flags=re.IGNORECASE)
            for block in blocks:
                block = block.strip()
                if not block:
                    continue
                # Extract COO
                coo_match = re.search(r"MADE IN\s+([A-Z]+)", block, re.IGNORECASE)
                coo = coo_match.group(1).upper() if coo_match else ""
                # Extract Quantity, Unit, Unit Price, Amount from last table line
                table_line_match = re.findall(
                    r"(\d+(?:\.\d+)?)\s+(\w+)\s*\(\s*([\d,]+\.\d+)\s*\)\s*\(\s*([\d,]+\.\d+)\s*\)",
                    block,
                    re.MULTILINE
                )
                if table_line_match:
                    quantity, unit, unit_price, amount = table_line_match[-1]
                    try:
                        quantity = float(quantity.replace(",", ""))
                    except:
                        quantity = 0.0
                    try:
                        unit_price = float(unit_price.replace(",", ""))
                    except:
                        unit_price = 0.0
                    try:
                        amount = float(amount.replace(",", ""))
                    except:
                        amount = 0.0
                else:
                    quantity = 0.0
                    unit = ""
                    unit_price = 0.0
                    amount = 0.0

                extracted_data.append({
                    "COO": coo,
                    "Quantity": quantity,
                    "Unit": unit,
                    "Unit Price": unit_price,
                    "Amount": amount,
                    "HS Code": "" 
                })

  
        # ===============================
        # CELINE MODE
        # ===============================
 
        elif mode == "Celine":
            hs_codes_positions = extract_all_hs_codes(saved_pdf_path)  # your existing function
            invoice_numbers_positions = extract_invoice_numbers(saved_pdf_path)
            # print("invoice_numbers_positions:",invoice_numbers_positions)
            lines = full_text.splitlines()

            seen_invoice_numbers = set() 
            hs_index = 0  # pointer to assign HS codes one by one
            inv_index = 0
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                # Extract COO
                coo_match = re.search(r"Made in\s+([A-Z]{2})", line, re.IGNORECASE)
                if not coo_match:
                    continue

                coo = coo_match.group(1).upper()

                # Remove "Made in XX"
                remaining_text = re.sub(
                    r"Made in\s+[A-Z]{2}", "", line, flags=re.IGNORECASE
                ).strip()

                parts = remaining_text.split()

                quantity = parts[0] if len(parts) >= 1 else ""
                unit = parts[1] if len(parts) >= 2 else ""
                unit_price = parts[2] if len(parts) >= 3 else ""
                #print("unit_price",unit_price)
                unit_price_1 = swap_separators(unit_price)
                #print("unit_price_1",unit_price_1)
                amount = parts[3] if len(parts) >= 4 else ""

                # Assign HS Code in order
                hs_code = ""
                if hs_index < len(hs_codes_positions):
                    hs_code = hs_codes_positions[hs_index]['hs_code']
                    hs_index += 1  # move to next HS code for the next line

                # invoice_number = ""
                # if inv_index < len(invoice_numbers_positions):
                #     invoice_number = invoice_numbers_positions[inv_index]['invoice_number']
                #     inv_index += 1
                #     print("invoice_number:",invoice_number)
                    invoice_number = ""
                    while inv_index < len(invoice_numbers_positions):
                        potential_invoice = invoice_numbers_positions[inv_index]['invoice_number']
                        inv_index += 1
                        if potential_invoice not in seen_invoice_numbers:
                            invoice_number = potential_invoice
                            seen_invoice_numbers.add(invoice_number)
                            break

                extracted_data.append({
                    "COO": coo,
                    "Quantity": quantity,
                    "Unit": unit,
                    "Unit Price": "",
                    "Amount": unit_price_1, 
                    "HS Code": hs_code,
                    "Invoice Number": invoice_number 

                })

                
        # ===============================
        # MARINETRANS MODE
        # ===============================

        elif mode == "Marinetrans":
            print('hello')

        
        # ===============================
        # CREATE EXCEL FILE
        # ===============================
        excel_fs = FileSystemStorage(location='media/excels')
        os.makedirs(excel_fs.location, exist_ok=True)
        excel_filename = os.path.splitext(pdf_file.name)[0] + f"_{mode}_trimmed.xlsx"
        excel_path = os.path.join(excel_fs.location, excel_filename)

        wb = Workbook()
        ws = wb.active
        ws.title = f"{mode} Trimmed Data"
        if mode == "Celine":
            ws.append(["COO", "Quantity", "Unit", "Unit Price", "Amount", "HS Code","Invoice Number"])
        else:
            ws.append(["COO", "Quantity", "Unit", "Unit Price", "Amount"])

        for item in extracted_data:
            if mode == "Celine":
                ws.append([
                    item["COO"],
                    item["Quantity"],
                    item["Unit"],
                    item["Unit Price"],
                    item["Amount"],
                    item["HS Code"],
                    item["Invoice Number"]
                ])
            else:
                ws.append([
                    item["COO"],
                    item["Quantity"],
                    item["Unit"],
                    item["Unit Price"],
                    item["Amount"]
                ])
        wb.save(excel_path)

        with open(excel_path, "rb") as f:
            response = HttpResponse(
                f.read(),
                content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            response["Content-Disposition"] = f'attachment; filename="{excel_filename}"'
            return response

    # GET request or no file uploaded
    return render(request, "index.html")


#nnr_checking

import re
import pdfplumber
from django.http import JsonResponse


def nnrchecking(request):
    pdf_path = r"C:\Users\Admin\Downloads\ilovepdf_merged (1)_merged.pdf"
    full_text = ""

    # Step 1: Read PDF
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    # Step 2: Regex to find complete line before "Made In"
    pattern = r"([^\r\n]+)\s*[\r\n]+\s*Made In\s+([A-Za-z]+)"
    matches = re.findall(pattern, full_text)

    results = []

    for index, (line_data, country) in enumerate(matches, start=1):
        # Remove leading numbers/special characters until first letter
        clean_line_match = re.search(r"[A-Za-z].*", line_data)
        if clean_line_match:
            clean_line_text = clean_line_match.group()
        else:
            clean_line_text = line_data.strip()

        # Split description and the rest
        first_number_match = re.search(r"\d", clean_line_text)
        if first_number_match:
            num_index = first_number_match.start()
            description = clean_line_text[:num_index].strip()
            rest = clean_line_text[num_index:].split()
        else:
            description = clean_line_text.strip()
            rest = []

        # Map the remaining fields safely
        quantity = rest[0] if len(rest) > 0 else ""
        unit = rest[1] if len(rest) > 1 else ""
        price = rest[2] if len(rest) > 2 else ""
        value1 = rest[3] if len(rest) > 3 else ""
        value2 = rest[4] if len(rest) > 4 else ""
        value3 = rest[5] if len(rest) > 5 else ""

        results.append({
            "serial_number": index,
            "description": description,
            "quantity": quantity,
            "unit": unit,
            "price": price,
            "country": country.strip()
        })

    return JsonResponse(results, safe=False)



#Marinetrans Checking
import pdfplumber
import re
from django.http import JsonResponse

def extract_hscode_lines(request):
    pdf_path = r"C:\Users\Admin\Downloads\6100000183 CIPL (3).pdf"
    data = []

    # Regex to find "HS Code" header in table
    header_pattern = re.compile(r"\bHS\s*Code\b", re.IGNORECASE)
    
    # Regex to match actual HS Codes: 6 digits + space + 2 digits
    hscode_pattern = re.compile(r"^\d{6}\s\d{2}$")

    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):

            tables = page.extract_tables()
            if not tables:
                continue

            for table_index, table in enumerate(tables):

                # Check if table has HS Code header
                table_text = "\n".join(
                    [" ".join([cell if cell else "" for cell in row]) for row in table]
                )

                if header_pattern.search(table_text):
                    # Extract only rows that have a valid HS Code
                    for row in table:
                        for idx, cell in enumerate(row):
                            if cell:
                                cell_text = cell.strip()
                                if hscode_pattern.match(cell_text):
                                    # Take HSCode and all data **after it**
                                    row_after_hscode = [c if c else "" for c in row[idx+1:]]
                                    
                                    row_dict = {
                                        "HSCode": cell_text,
                                        "RowData": row_after_hscode
                                    }
                                    data.append(row_dict)
                                    break  # Only take first HS Code per row

    return JsonResponse({"data": data})


    
import re
from django.http import JsonResponse
import pdfplumber

def extract_customer_po_data(request):
    pdf_path = r"C:\Users\Admin\Downloads\NNRSG0235696.pdf"

    try:
        po_data = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = text.splitlines()
                capture = False

                for i in range(len(lines)):
                    line = lines[i].strip()
                    if not line:
                        continue

                    # Start capturing after header
                    if not capture and re.search(
                        r"Item\s+Customer'?s?\s+P/O\s+No\.?",
                        line,
                        re.IGNORECASE
                    ):
                        capture = True
                        continue

                    if capture:
                        po_match = re.match(r"^\d+\s+([A-Z]+[A-Z0-9-]+)\s+(.*)", line)

                        if po_match:
                            po_number = po_match.group(1)
                            main_row = po_match.group(2).split()

                            # -----------------------------
                            # 🔥 FULL PROCESSING LOGIC HERE
                            # -----------------------------

                            quantity_index = None

                            # Detect quantity (like 3,000)
                            for idx, value in enumerate(main_row):
                                if re.match(r"^\d{1,3}(,\d{3})*$", value):
                                    quantity_index = idx
                                    break

                            if quantity_index is not None and len(main_row) >= quantity_index + 3:

                                description = " ".join(main_row[:quantity_index])
                                quantity = main_row[quantity_index]
                                unit_price = main_row[quantity_index + 1]

                                raw_amount = main_row[quantity_index + 2]
                                amount = re.sub(r"[^\d.,]", "", raw_amount)

                                # Take only immediate next line
                                customer_pn = ""
                                toshiba_pn = ""

                                if i + 1 < len(lines):
                                    next_line = lines[i + 1].strip()
                                    parts = next_line.split()

                                    if len(parts) >= 1:
                                        customer_pn = parts[0]

                                    if len(parts) >= 2:
                                        toshiba_pn = parts[1].replace("RoHS", "")

                                po_data.append({
                                    "PO": po_number,
                                    "Description": description,
                                    "Quantity": quantity,
                                    "UnitPrice": unit_price,
                                    "Amount": amount,
                                    "CustomerPN": customer_pn,
                                    "ToshibaPN": toshiba_pn
                                })

        return JsonResponse({"po_data": po_data})

    except FileNotFoundError:
        return JsonResponse({"error": "PDF file not found."}, status=404)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)



#
import re
from django.http import JsonResponse
import pdfplumber

def extract_package_data(request):
    pdf_path = r"C:\Users\Admin\Downloads\NNRSG0235696.pdf"

    try:
        package_data = []

        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if not text:
                    continue

                lines = [line.strip() for line in text.splitlines() if line.strip()]
                capture = False
                i = 0

                while i < len(lines):
                    line = lines[i]

                    # Start capturing after header containing "Package"
                    if not capture and "Package" in line and "Customer" in line and "Part" in line:
                        capture = True
                        i += 1
                        continue

                    if capture:
                        # Detect package start: number at start of line
                        package_match = re.match(r"^(\d+)\s*(.*)", line)
                        if package_match:
                            package_number = package_match.group(1)
                            package_lines = []

                            # Capture all lines until next package number or end
                            j = i + 1
                            while j < len(lines):
                                next_line = lines[j]
                                next_match = re.match(r"^\d+\s*", next_line)
                                if next_match:
                                    break
                                package_lines.append(next_line)
                                j += 1

                            # Extract fields
                            customer_part = package_lines[0] if len(package_lines) > 0 else ""
                            coo = package_lines[1] if len(package_lines) > 1 else ""

                            if customer_part.strip() or coo.strip():
                                # Use regex to split fields reliably
                                part_pattern = re.match(r"(\S+)\s+(.+)\s+(\d+)\s+([\d\.]+)\s+([\d\.]+)", customer_part)
                                if part_pattern:
                                    part_number = part_pattern.group(1)
                                    description = part_pattern.group(2)
                                    quantity = part_pattern.group(3)
                                    weight = part_pattern.group(4)
                                    unit_price = part_pattern.group(5)
                                else:
                                    parts = customer_part.split()
                                    part_number = parts[0] if len(parts) > 0 else ""
                                    description = " ".join(parts[1:-3]) if len(parts) > 4 else ""
                                    quantity = parts[-3] if len(parts) > 2 else ""
                                    weight = parts[-2] if len(parts) > 1 else ""
                                    unit_price = parts[-1] if len(parts) > 0 else ""

                                # Clean COO field
                                coo_clean = coo.upper().replace("MADE IN ", "").strip()

                                package_data.append({
                                    "PartNumber": part_number,
                                    "Description": description,
                                    "Quantity": quantity,
                                    "Weight": weight,
                                    "UnitPrice": unit_price,
                                    "Coo": coo_clean
                                })

                            i = j
                            continue

                    i += 1

        return JsonResponse({"package_data": package_data})

    except FileNotFoundError:
        return JsonResponse({"error": "PDF file not found."}, status=404)
    except Exception as e:
        return JsonResponse({"error": str(e)}, status=500)




 #nnr_TOSHIBA_checking

# import re
# from django.http import JsonResponse
# import pdfplumber

# def extract_customer_po_data(request):
#     pdf_path = r"C:\Users\Admin\Downloads\NNRSG0235503 (1).pdf"

#     try:
#         po_data = []

#         with pdfplumber.open(pdf_path) as pdf:
#             for page_number, page in enumerate(pdf.pages, start=1):
#                 # Extract all text lines with positions
#                 lines = page.extract_text().splitlines()
#                 capture = False
#                 current_po = None
#                 row_data = []

#                 for line in lines:
#                     line = line.strip()
#                     if not line:
#                         continue

#                     # Start capturing after "Item Customer's P/O No."
#                     if re.search(r"Item\s+Customer'?s?\s+P/O\s+No\.", line, re.IGNORECASE):
#                         capture = True
#                         continue

#                     if capture:
#                         # Match numbered PO line e.g., "1 PO-APN1-0042898 ..."
#                         po_match = re.match(r"\d+\s+(PO-[A-Z0-9-]+)", line)
#                         if po_match:
#                             # Save previous PO if exists
#                             if current_po:
#                                 po_data.append({
#                                     "PO": current_po,
#                                     "RowData": row_data
#                                 })

#                             current_po = po_match.group(1)
#                             row_data = [line]  # start new row
#                         else:
#                             # Add subsequent lines under current PO
#                             if current_po:
#                                 row_data.append(line)

#                 # Save last PO on page
#                 if current_po:
#                     po_data.append({
#                         "PO": current_po,
#                         "RowData": row_data
#                     })

#         return JsonResponse({"po_data": po_data})

#     except FileNotFoundError:
#         return JsonResponse({"error": "PDF file not found."}, status=404)
#     except Exception as e:
#         return JsonResponse({"error": str(e)}, status=500)


        # ws1 = wb.active
        # ws1.title = "PO Data"
        # ws1.append(["PO", "Description", "Quantity", "UnitPrice", "Amount", "CustomerPN", "ToshibaPN"])
        # for item in po_data:
        #     ws1.append([
        #         item["PO"],
        #         item["Description"],
        #         item["Quantity"],
        #         item["UnitPrice"],
        #         item["Amount"],
        #         item["CustomerPN"],
        #         item["ToshibaPN"]
        #     ])

        # # Sheet 2: Package Data
        # ws2 = wb.create_sheet(title="Package Data")
        # ws2.append(["PartNumber", "Description", "Quantity", "Weight", "UnitPrice", "Coo"])
        # for item in package_data:
        #     ws2.append([
        #         item["PartNumber"],
        #         item["Description"],
        #         item["Quantity"],
        #         item["Weight"],
        #         item["UnitPrice"],
        #         item["Coo"]
        #     ])


# linux checking
import re
import pdfplumber
from django.http import JsonResponse

def linux_checking(request):
    pdf_path = r"C:\Users\Admin\Downloads\313023900 - SUZANNE COYLE.pdf"
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = full_text.splitlines()

    # Skip lines until column header, stop at "Shipping Charge"
    header_found = False
    start_processing = []
    for line in lines:
        line = line.strip()
        # Detect table header
        if not header_found and ("Commodity" in line and "Full Description of" in line and "No. of" in line):
            header_found = True
            continue
        if header_found:
            # Stop at any line containing "Shipping Charge"
            if "Shipping Charge" in line:
                break
            start_processing.append(line)

    current_row = None

    for line in start_processing:
        if not line:
            continue

        # Check if line starts with HS code
        hs_match = re.match(r"^(\d{4,10})\s*(.*)", line)
        if hs_match:
            # Append previous row(s)
            if current_row:
                numeric_parts = current_row['numeric_part']
                for i in range(0, len(numeric_parts), 2):
                    unit_value = numeric_parts[i] if i < len(numeric_parts) else "0"
                    total_value = numeric_parts[i+1] if i+1 < len(numeric_parts) else "0"
                    description = " ".join(current_row['description_lines']).strip()
                    extracted_data.append({
                        "HS Code": current_row['HS Code'],
                        "Description": description,
                        "Unit Value": unit_value,
                        "Total Value": total_value
                    })

            hs_code = hs_match.group(1)
            rest_of_line = hs_match.group(2)

            # Extract numeric values as strings; anything non-numeric is description
            parts = rest_of_line.split()
            numeric_parts = []
            description_part = []
            for part in parts:
                # Match digits with optional comma and decimal
                if re.match(r"^\d[\d,]*\.\d+$", part):
                    numeric_parts.append(part)
                else:
                    description_part.append(part)

            current_row = {
                "HS Code": hs_code,
                "description_lines": description_part,
                "numeric_part": numeric_parts  # Keep as list
            }

        else:
            # Line continuation for description
            if current_row:
                current_row['description_lines'].append(line)

    # Append last row(s)
    if current_row:
        numeric_parts = current_row['numeric_part']
        for i in range(0, len(numeric_parts), 2):
            unit_value = numeric_parts[i] if i < len(numeric_parts) else "0"
            total_value = numeric_parts[i+1] if i+1 < len(numeric_parts) else "0"
            description = " ".join(current_row['description_lines']).strip()
            extracted_data.append({
                "HS Code": current_row['HS Code'],
                "Description": description,
                "Unit Value": unit_value,
                "Total Value": total_value
            })

    return JsonResponse({"data": extracted_data})




import re
import pdfplumber
from django.http import JsonResponse

def linux_checking(request):
    pdf_path = r"C:\Users\Admin\Downloads\313023900 - SUZANNE COYLE.pdf"
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = full_text.splitlines()

    # Skip lines until column header, stop at "Shipping Charge"
    header_found = False
    start_processing = []
    for line in lines:
        if not header_found and ("Commodity" in line and "Full Description of" in line and "No. of" in line):
            header_found = True
            continue
        if header_found:
            if "Shipping Charge" in line:
                break
            start_processing.append(line.strip())

    current_row = None

    for line in start_processing:
        if not line:
            continue

        # Check if line starts with HS code
        hs_match = re.match(r"^(\d{4,10})\s*(.*)", line)
        if hs_match:
            # Append previous row(s)
            if current_row:
                numeric_parts = current_row['numeric_part']
                # Extract No. of Items, Unit Value, Total Value
                no_of_items = int(numeric_parts[0]) if len(numeric_parts) > 0 else 0
                unit_value = float(numeric_parts[1].replace(",", "")) if len(numeric_parts) > 1 else 0.0
                total_value = float(numeric_parts[2].replace(",", "")) if len(numeric_parts) > 2 else 0.0
                description = " ".join(current_row['description_lines']).strip()
                extracted_data.append({
                    "HS Code": current_row['HS Code'],
                    "Description": description,
                    "No. of Items": no_of_items,
                    "Unit Value": unit_value,
                    "Total Value": total_value
                })

            hs_code = hs_match.group(1)
            rest_of_line = hs_match.group(2)

            # Extract numeric values; anything else is description
            parts = rest_of_line.split()
            numeric_parts = []
            description_part = []
            for part in parts:
                # Match integers or floats
                if re.match(r"^\d+(\.\d+)?$", part.replace(",", "")):
                    numeric_parts.append(part)
                else:
                    description_part.append(part)

            current_row = {
                "HS Code": hs_code,
                "description_lines": description_part,
                "numeric_part": numeric_parts
            }

        else:
            # Line continuation for description
            if current_row:
                current_row['description_lines'].append(line)

    # Append last row
    if current_row:
        numeric_parts = current_row['numeric_part']
        no_of_items = int(numeric_parts[0]) if len(numeric_parts) > 0 else 0
        unit_value = float(numeric_parts[1].replace(",", "")) if len(numeric_parts) > 1 else 0.0
        total_value = float(numeric_parts[2].replace(",", "")) if len(numeric_parts) > 2 else 0.0
        description = " ".join(current_row['description_lines']).strip()
        extracted_data.append({
            "HS Code": current_row['HS Code'],
            "Description": description,
            "No. of Items": no_of_items,
            "Unit Value": unit_value,
            "Total Value": total_value
        })

    return JsonResponse({"data": extracted_data})


# linux checking
import re
import pdfplumber
from django.http import JsonResponse

def linux_checking(request):
    pdf_path = r"C:\Users\Admin\Downloads\313023900 - SUZANNE COYLE.pdf"
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = full_text.splitlines()

    # Skip lines until column header, stop at "Shipping Charge"
    header_found = False
    start_processing = []
    for line in lines:
        if not header_found and ("Commodity" in line and "Full Description of" in line and "No. of" in line):
            header_found = True
            continue
        if header_found:
            if "Shipping Charge" in line:
                break
            start_processing.append(line.strip())

    current_row = None

    for line in start_processing:
        if not line:
            continue

        # Check if line starts with HS code
        hs_match = re.match(r"^(\d{4,10})\s*(.*)", line)
        if hs_match:
            # Append previous row(s)
            if current_row:
                numeric_parts = current_row['numeric_part']
                no_of_items = numeric_parts[0] if len(numeric_parts) > 0 else "0"
                unit_value = numeric_parts[1] if len(numeric_parts) > 1 else "0"
                total_value = numeric_parts[2] if len(numeric_parts) > 2 else "0"
                coo = numeric_parts[3] if len(numeric_parts) > 3 else ""
                description = " ".join(current_row['description_lines']).strip()
                extracted_data.append({
                    "HS Code": current_row['HS Code'],
                    "Description": description,
                    "No. of Items": no_of_items,
                    "Unit Value": unit_value,
                    "Total Value": total_value,
                    "COO": coo
                })

            hs_code = hs_match.group(1)
            rest_of_line = hs_match.group(2)

            # Extract numeric values and COO; everything else is description
            parts = rest_of_line.split()
            numeric_parts = []
            description_part = []
            for part in parts:
                # Match integers, floats, or country code (all strings)
                if re.match(r"^\d[\d,]*\.\d+$", part) or re.match(r"^\d+$", part) or re.match(r"^[A-Z]{2,3}$", part):
                    numeric_parts.append(part)
                else:
                    description_part.append(part)

            current_row = {
                "HS Code": hs_code,
                "description_lines": description_part,
                "numeric_part": numeric_parts  # Keep as list of strings
            }

        else:
            # Line continuation for description
            if current_row:
                current_row['description_lines'].append(line)

    # Append last row
    if current_row:
        numeric_parts = current_row['numeric_part']
        no_of_items = numeric_parts[0] if len(numeric_parts) > 0 else "0"
        unit_value = numeric_parts[1] if len(numeric_parts) > 1 else "0"
        total_value = numeric_parts[2] if len(numeric_parts) > 2 else "0"
        coo = numeric_parts[3] if len(numeric_parts) > 3 else ""
        description = " ".join(current_row['description_lines']).strip()
        extracted_data.append({
            "HS Code": current_row['HS Code'],
            "Description": description,
            "No. of Items": no_of_items,
            "Unit Value": unit_value,
            "Total Value": total_value,
            "COO": coo
        })

    return JsonResponse({"data": extracted_data})

import re
import pdfplumber
from django.http import JsonResponse

def remove_duplicates(text):
    """
    Remove consecutive duplicate words/phrases in a string.
    Example: "CAMBER SOFT WASH CAMBER SOFT WASH NAPKIN S/4 NAPKIN S/4"
    becomes: "CAMBER SOFT WASH NAPKIN S/4"
    """
    words = text.split()
    deduped = []
    for word in words:
        if not deduped or word != deduped[-1]:
            deduped.append(word)
    return " ".join(deduped)

def linux_checking(request):
    pdf_path = r"C:\Users\Admin\Downloads\312410383 - BENZ MONG.pdf"
    extracted_data = []

    # 1️⃣ Extract all text from PDF
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = full_text.splitlines()
    current_row = None
    start_processing = False  # Flag to start when header is found

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Start processing after header row
        if not start_processing:
            if ("Commodity" in line and "Full Description of" in line and "No. of" in line):
                start_processing = True
            continue  # skip header line itself

        # Stop processing at Shipping Charge
        if "Shipping Charge" in line:
            break

        # Check if line starts with HS Code
        hs_match = re.match(r"^(\d{4,10})\s+(.*)", line)
        if hs_match:
            # Save previous row
            if current_row:
                # Deduplicate description before saving
                current_row['Description'] = remove_duplicates(current_row['Description'])
                extracted_data.append(current_row)

            # Start new product row
            hs_code = hs_match.group(1)
            rest_of_line = hs_match.group(2).strip()
            tokens = rest_of_line.split()

            # COO = last token
            coo = tokens[-1] if tokens else ""

            # Numeric values = before COO
            numeric_tokens = [t for t in tokens[:-1] if re.match(r"^\d+(\.\d+)?$", t)]
            no_of_items = numeric_tokens[0] if len(numeric_tokens) >= 1 else "0"
            unit_value = numeric_tokens[1] if len(numeric_tokens) >= 2 else "0"
            total_value = numeric_tokens[2] if len(numeric_tokens) >= 3 else "0"

            # Description = all tokens before first numeric
            first_num_index = next((i for i, t in enumerate(tokens) if re.match(r"^\d+(\.\d+)?$", t)), len(tokens))
            description_tokens = tokens[:first_num_index]
            description = " ".join(description_tokens)

            # Initialize current row
            current_row = {
                "HS Code": hs_code,
                "Description": description,
                "Product Composition": "",
                "No. of Items": no_of_items,
                "Unit Value": unit_value,
                "Total Value": total_value,
                "COO": coo
            }
        else:
            # Additional description/product composition
            if current_row:
                current_row['Description'] += " " + line

    # Append last row
    if current_row:
        current_row['Description'] = remove_duplicates(current_row['Description'])
        extracted_data.append(current_row)

    return JsonResponse({"data": extracted_data})


    

# import re
# import pdfplumber
# from django.http import JsonResponse

# def remove_duplicates(text):
#     """
#     Remove consecutive duplicate words/phrases in a string.
#     Example: "CAMBER SOFT WASH CAMBER SOFT WASH NAPKIN S/4 NAPKIN S/4"
#     becomes: "CAMBER SOFT WASH NAPKIN S/4"
#     """
#     words = text.split()
#     i = 0
#     while i < len(words):
#         # Try sequences from longest possible to 1
#         for size in range(len(words)//2, 0, -1):
#             if i + 2*size <= len(words):
#                 first = words[i:i+size]
#                 second = words[i+size:i+2*size]
#                 if first == second:
#                     # Remove the repeated second sequence
#                     del words[i+size:i+2*size]
#                     # restart checking from same position
#                     i = max(i - size, 0)
#                     break
#         else:
#             i += 1
#     return " ".join(words)

# def linux_checking(request):
#     pdf_path = r"C:\Users\Admin\Downloads\313023900 - SUZANNE COYLE.pdf"
#     extracted_data = []

#     # 1️⃣ Extract all text from PDF
#     with pdfplumber.open(pdf_path) as pdf:
#         full_text = ""
#         for page in pdf.pages:
#             text = page.extract_text()
#             if text:
#                 full_text += text + "\n"

#     lines = full_text.splitlines()
#     current_row = None
#     start_processing = False  # Flag to start when header is found

#     for line in lines:
#         line = line.strip()
#         if not line:
#             continue

#         # ---------------------------
#         # Start processing after header row
#         # ---------------------------
#         if not start_processing:
#             if ("Commodity" in line and "Full Description of" in line and "No. of" in line):
#                 start_processing = True
#             continue  # skip header line itself

#         # ---------------------------
#         # Stop processing at Shipping Charge
#         # ---------------------------
#         if "Shipping Charge" in line:
#             break

#         # ---------------------------
#         # Process HS Code rows
#         # ---------------------------
#         hs_match = re.match(r"^(\d{4,10})\s+(.*)", line)
#         if hs_match:
#             # Save previous row
#             if current_row:
#                 extracted_data.append(current_row)

#             # Start new product row
#             hs_code = hs_match.group(1)
#             rest_of_line = hs_match.group(2).strip()
#             tokens = rest_of_line.split()

#             # COO = last token
#             coo = tokens[-1] if tokens else ""

#             # Numeric values = before COO
#             numeric_tokens = [t for t in tokens[:-1] if re.match(r"^\d+(\.\d+)?$", t)]
#             no_of_items = numeric_tokens[0] if len(numeric_tokens) >= 1 else "0"
#             unit_value = numeric_tokens[1] if len(numeric_tokens) >= 2 else "0"
#             total_value = numeric_tokens[2] if len(numeric_tokens) >= 3 else "0"

#             # Description = all tokens before first numeric
#             first_num_index = next((i for i, t in enumerate(tokens) if re.match(r"^\d+(\.\d+)?$", t)), len(tokens))
#             description_tokens = tokens[:first_num_index]
#             description = " ".join(description_tokens)

#             # Initialize current row
#             current_row = {
#                 "HS Code": hs_code,
#                 "Description": description,
#                 "Product Composition": "",
#                 "No. of Items": no_of_items,
#                 "Unit Value": unit_value,
#                 "Total Value": total_value,
#                 "COO": coo
#             }
#         else:
#             # This line is additional description/product composition
#             if current_row:
#                current_row['Description'] += " " + line
#     # Append last row
#     if current_row:
#             current_row['Description'] = remove_duplicates(current_row['Description'])
#             extracted_data.append(current_row)

#     return JsonResponse({"data": extracted_data})


import re
import pdfplumber
from django.http import JsonResponse

def remove_repeated_phrases(text):
    """
    Remove repeated multi-word sequences in a string.
    Example:
    "CAMBER SOFT WASH CAMBER SOFT WASH NAPKIN S/4 NAPKIN S/4"
    -> "CAMBER SOFT WASH NAPKIN S/4"
    """
    words = text.split()
    i = 0
    while i < len(words):
        for size in range(len(words)//2, 0, -1):
            if i + 2*size <= len(words):
                first = words[i:i+size]
                second = words[i+size:i+2*size]
                if first == second:
                    del words[i+size:i+2*size]
                    # After deletion, check again from current position
                    i = max(i - size, 0)
                    break
        else:
            i += 1
    return " ".join(words)

def linux_checking(request):
    pdf_path = r"C:\Users\Admin\Downloads\313023900 - SUZANNE COYLE.pdf"
    extracted_data = []

    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                full_text += text + "\n"

    lines = full_text.splitlines()
    current_row = None
    start_processing = False

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Start after header row
        if not start_processing:
            if ("Commodity" in line and "Full Description of" in line and "No. of" in line):
                start_processing = True
            continue

        # Stop at Shipping Charge
        if "Shipping Charge" in line:
            break

        # HS Code row
        hs_match = re.match(r"^(\d{4,10})\s+(.*)", line)
        if hs_match:
            if current_row:
                # Deduplicate before saving
                current_row['Description'] = remove_repeated_phrases(current_row['Description'])
                extracted_data.append(current_row)

            hs_code = hs_match.group(1)
            rest_of_line = hs_match.group(2).strip()
            tokens = rest_of_line.split()

            coo = tokens[-1] if tokens else ""
            numeric_tokens = [t for t in tokens[:-1] if re.match(r"^\d+(\.\d+)?$", t)]
            no_of_items = numeric_tokens[0] if len(numeric_tokens) >= 1 else "0"
            unit_value = numeric_tokens[1] if len(numeric_tokens) >= 2 else "0"
            total_value = numeric_tokens[2] if len(numeric_tokens) >= 3 else "0"

            first_num_index = next((i for i, t in enumerate(tokens) if re.match(r"^\d+(\.\d+)?$", t)), len(tokens))
            description_tokens = tokens[:first_num_index]
            description = " ".join(description_tokens)

            current_row = {
                "HS Code": hs_code,
                "Description": description,
                "Product Composition": "",
                "No. of Items": no_of_items,
                "Unit Value": unit_value,
                "Total Value": total_value,
                "COO": coo
            }
        else:
            if current_row:
                current_row['Description'] += " " + line

    # Append last row
    if current_row:
        current_row['Description'] = remove_repeated_phrases(current_row['Description'])
        extracted_data.append(current_row)

    return JsonResponse({"data": extracted_data})