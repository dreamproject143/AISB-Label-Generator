import os
import tempfile
from flask import Flask, render_template, request, send_file, jsonify
import pdfplumber
import re
import openpyxl
from openpyxl.styles import Font, Border, Side, Alignment
from openpyxl.drawing.image import Image
from openpyxl.worksheet.pagebreak import Break # <-- Added this for page breaks

app = Flask(__name__)

# Logo path
logo_file_path = os.path.join(os.path.dirname(__file__), 'logo.png')


def draw_label_in_excel(sheet, start_row, data):
    # --- Define Styles (matched with your local code) ---
    bold_font_large = Font(name='Arial', size=16, bold=True)
    bold_font_medium = Font(name='Arial', size=12, bold=True)
    bold_font_small = Font(name='Arial', size=11, bold=True)
    bold_font_qty = Font(name='Arial', size=14, bold=True) # <-- Added new font style for QTY

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left_no_wrap = Alignment(horizontal='left', vertical='center', wrap_text=False)

    # --- Changed border style from 'thick' to 'medium' ---
    medium_side = Side(style='medium')
    thin_side = Side(style='thin')
    border_medium_all = Border(left=medium_side, right=medium_side, top=medium_side, bottom=medium_side)

    # --- Adjusted Column Widths and Row Heights (matched with your local code) ---
    sheet.column_dimensions['A'].width = 25
    sheet.column_dimensions['B'].width = 12
    sheet.column_dimensions['C'].width = 20
    sheet.column_dimensions['D'].width = 25
    
    sheet.row_dimensions[start_row].height = 38
    sheet.row_dimensions[start_row + 1].height = 38
    sheet.row_dimensions[start_row + 2].height = 22
    sheet.row_dimensions[start_row + 3].height = 28
    sheet.row_dimensions[start_row + 4].height = 21
    sheet.row_dimensions[start_row + 5].height = 38

    # --- Merge Cells (no changes here) ---
    sheet.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=1)
    sheet.merge_cells(start_row=start_row, start_column=2, end_row=start_row, end_column=4)
    sheet.merge_cells(start_row=start_row + 1, start_column=1, end_row=start_row + 1, end_column=1)
    sheet.merge_cells(start_row=start_row + 1, start_column=2, end_row=start_row + 1, end_column=3)
    sheet.merge_cells(start_row=start_row + 2, start_column=1, end_row=start_row + 2, end_column=1)
    sheet.merge_cells(start_row=start_row + 2, start_column=2, end_row=start_row + 2, end_column=3)
    sheet.merge_cells(start_row=start_row + 3, start_column=1, end_row=start_row + 3, end_column=1)
    sheet.merge_cells(start_row=start_row + 3, start_column=2, end_row=start_row + 3, end_column=3)
    sheet.merge_cells(start_row=start_row + 4, start_column=1, end_row=start_row + 5, end_column=1)
    sheet.merge_cells(start_row=start_row + 4, start_column=2, end_row=start_row + 5, end_column=2)
    sheet.merge_cells(start_row=start_row + 4, start_column=3, end_row=start_row + 5, end_column=3)

    # --- Place Data and Apply Styles (matched with your local code) ---
    cell = sheet.cell(row=start_row, column=2, value="JOONHEE ENGINEERING SDN. BHD.")
    cell.font = Font(name='Arial', size=18, bold=True)
    cell.alignment = align_left_no_wrap

    cell = sheet.cell(row=start_row + 1, column=1, value="JOONHEE ENGINEERING")
    cell.font = bold_font_medium
    cell.alignment = align_center

    cell = sheet.cell(row=start_row + 1, column=2, value="→")
    cell.font = Font(name='Arial', size=55, bold=True)
    cell.alignment = align_center

    qty_text = f"AISB\nQTY: {data['qty']}"
    qty_cell = sheet.cell(row=start_row + 1, column=4, value=qty_text)
    qty_cell.font = bold_font_qty # <-- Using the new font style
    qty_cell.alignment = align_center

    sheet.cell(row=start_row + 2, column=1, value="PART NO.").font = bold_font_small
    sheet.cell(row=start_row + 2, column=2, value="KANBAN NO.").font = bold_font_small
    sheet.cell(row=start_row + 2, column=4, value="ISSUE DATE :").font = bold_font_small
    sheet.cell(row=start_row + 4, column=1, value="PART NAME").font = bold_font_small
    sheet.cell(row=start_row + 4, column=2, value="COLOR CODE").font = bold_font_small
    sheet.cell(row=start_row + 4, column=4, value="DELIVERY DATE :").font = bold_font_small

    sheet.cell(row=start_row + 3, column=1, value=data['part_no']).font = bold_font_large
    sheet.cell(row=start_row + 3, column=2, value=data['kanban_no']).font = bold_font_large
    sheet.cell(row=start_row + 3, column=4, value=data['issue_date']).font = bold_font_large
    sheet.cell(row=start_row + 5, column=4, value=data['delivery_date']).font = bold_font_large

    part_name_cell = sheet.cell(row=start_row + 4, column=1)
    part_name_cell.value = data['part_name']
    part_name_cell.font = bold_font_medium

    for r in range(start_row + 2, start_row + 6):
        for c in range(1, 5):
            sheet.cell(row=r, column=c).alignment = align_center

    # --- Apply Borders (using medium style) ---
    for r in range(start_row, start_row + 6):
        for c in range(1, 5):
            sheet.cell(row=r, column=c).border = border_medium_all

    # --- Divider Lines (using medium style) ---
    for col in range(1, 5):
        sheet.cell(row=start_row + 2, column=col).border = Border(bottom=thin_side, left=medium_side, right=medium_side, top=medium_side)
        sheet.cell(row=start_row + 4, column=col).border = Border(bottom=thin_side, left=medium_side, right=medium_side, top=medium_side)

    if os.path.exists(logo_file_path):
        img = Image(logo_file_path)
        img.height = 50
        img.width = 85
        sheet.add_image(img, f'A{start_row}')


def process_pdfs(files):
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Labels"
    
    # --- Added Page Setup for A4 Printing (from your local code) ---
    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4
    sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT
    sheet.page_margins.top = 0
    sheet.page_margins.bottom = 0
    sheet.page_margins.left = 0
    sheet.page_margins.right = 0
    sheet.sheet_view.showGridLines = False

    # --- Step 1: Gather all labels from all PDFs into a single list ---
    all_labels_info = []
    for file in files:
        with pdfplumber.open(file) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

        issue_date_match = re.search(r"Issue Date\s*:\s*(\d{2}-\w{3}-\d{4})", full_text)
        delivery_date_match = re.search(r"Delivery Date\s*:\s*(\d{2}-\w{3}-\d{4})", full_text)
        issue_date = issue_date_match.group(1) if issue_date_match else "N/A"
        delivery_date = delivery_date_match.group(1) if delivery_date_match else "N/A"

        lines = [l.strip() for l in full_text.split("\n") if l.strip()]
        item_lines = [l for l in lines if re.match(r'^\d+\s+[A-Z0-9\-]+\s+.*EA$', l)]

        for i, line in enumerate(item_lines):
            parts = line.split()
            std_pkg_qty = parts[-3]
            description_tokens = parts[2:-3]

            if len(parts) > 2 and re.match(r'[A-Z]-\d{3,}', parts[2]):
                part_no = parts[2]
            else:
                internal_code = parts[1]
                match = re.search(r'([A-Z])[A-Z0-9]*-?([0-9]{3,})', internal_code)
                if match:
                    part_no = f"{match.group(1)}-{match.group(2)}"
                else:
                    part_no = internal_code

            if description_tokens and description_tokens[0] == part_no:
                description_tokens = description_tokens[1:]
            description = " ".join(description_tokens)

            start_idx = lines.index(line) + 1
            end_idx = lines.index(item_lines[i + 1]) if i + 1 < len(item_lines) else len(lines)
            block_lines = lines[start_idx:end_idx]
            kanban_cards = []
            for bl in block_lines:
                kanban_cards.extend(re.findall(r'\b\d{10}\b', bl))

            for card in kanban_cards:
                label_info = {
                    "part_no": part_no,
                    "part_name": description,
                    "qty": std_pkg_qty,
                    "kanban_no": card,
                    "issue_date": issue_date,
                    "delivery_date": delivery_date
                }
                # Instead of drawing here, we add it to our list
                all_labels_info.append(label_info)

    # --- Step 2: Draw all collected labels with page break logic ---
    current_excel_row = 2
    total_labels = len(all_labels_info)
    for i, label_data in enumerate(all_labels_info):
        draw_label_in_excel(sheet, current_excel_row, label_data)
        current_excel_row += 7  # Add a 1-row gap for cutting
        
        # This adds a page break after every 4th label
        if (i + 1) % 4 == 0 and (i + 1) < total_labels:
            sheet.row_breaks.append(Break(id=current_excel_row - 1))

    # --- Save and return the file ---
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    workbook.save(temp_file.name)
    temp_file.close()
    return temp_file.name


@app.route('/')
def home():
    # Assuming your index.html is in a 'templates' folder
    return render_template('index.html')


@app.route('/upload', methods=['POST'])
def upload_and_process():
    if 'files' not in request.files:
        return jsonify({"success": False, "message": "No files uploaded"})

    files = request.files.getlist('files')
    pdf_files = [file for file in files if file.filename.endswith('.pdf')]

    if not pdf_files:
        return jsonify({"success": False, "message": "No valid PDF files"})

    excel_path = process_pdfs(pdf_files)
    return send_file(excel_path, as_attachment=True, download_name="Generated_Labels.xlsx")


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
