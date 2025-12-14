from flask import Flask, render_template, request, jsonify, send_file
import firebase_admin
from firebase_admin import credentials, storage
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Font, Alignment
from werkzeug.middleware.proxy_fix import ProxyFix
from openpyxl.styles import Border, Side
from PIL import Image
import io
import os
from datetime import datetime
import json
import base64
import uuid
from PIL import Image as PILImage
from openpyxl.styles import PatternFill  # Add this line
import os
from dotenv import load_dotenv
import gc  # Add this line at the top with other imports

app = Flask(__name__)
app.wsgi_app = ProxyFix(app.wsgi_app)

app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB limit
load_dotenv()

# Firebase Configuration
firebase_config = {
    "apiKey": os.getenv('FIREBASE_API_KEY'),
    "authDomain": os.getenv('FIREBASE_AUTH_DOMAIN'),
    "projectId": os.getenv('FIREBASE_PROJECT_ID'),
    "storageBucket": os.getenv('FIREBASE_STORAGE_BUCKET'),
    "messagingSenderId": os.getenv('FIREBASE_MESSAGING_SENDER_ID'),
    "appId": os.getenv('FIREBASE_APP_ID')
}

# Initialize Firebase
if not firebase_admin._apps:
    cred = credentials.Certificate(os.getenv('SERVICE_ACCOUNT_KEY_PATH', 'serviceAccountKey.json'))
    firebase_admin.initialize_app(cred, {
        'storageBucket': firebase_config['storageBucket']
    })

bucket = storage.bucket()

# Template path
TEMPLATE_PATH = 'templates/excel_template.xlsx'

def calculate_row_height(text, font_size=12, cell_width=50):
    """Calculate appropriate row height based on text length"""
    if not text:
        return 15  # Default height for empty cells
    
    # Approximate characters per line based on cell width
    chars_per_line = cell_width * 1.5  # Rough estimate
    
    # Calculate number of lines needed
    lines = len(text) / chars_per_line
    if lines < 1:
        lines = 1
    
    # Calculate height (font_size * 1.5 per line + padding)
    height = (font_size * 1.5 * lines) + 5
    
    # Minimum height
    if height < 15:
        height = 15
    
    # Maximum height (to avoid extremely tall rows)
    if height > 150:
        height = 150
    
    return height

def unmerge_and_write(ws, cell_address, value):
    """Safely write to a cell, handling merged cells"""
    try:
        # Check if cell is part of a merged range
        cell = ws[cell_address]
        for merged_range in list(ws.merged_cells.ranges):  # Convert to list to avoid modification during iteration
            if cell.coordinate in merged_range:
                # Unmerge the range
                ws.unmerge_cells(str(merged_range))
                # Write to the top-left cell of the merged range
                top_left_cell = merged_range.start_cell
                top_left_cell.value = value
                # Re-merge the cells
                ws.merge_cells(str(merged_range))
                return
        # If not merged, write directly
        cell.value = value
    except Exception as e:
        print(f"Error writing to {cell_address}: {e}")
        try:
            ws[cell_address].value = value
        except:
            pass
@app.route('/')
def index():
    return render_template('index.html')

@app.route('/check-existing', methods=['POST'])
def check_existing():
    """Check if SOL ID and Visit No already exists"""
    data = request.json
    sol_id = data.get('sol_id')
    visit_no = data.get('visit_no')
    
    # Check in Firebase Storage
    blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/data.json"
    blob = bucket.blob(blob_path)
    
    if blob.exists():
        # Download existing data
        existing_data = json.loads(blob.download_as_string())
        return jsonify({
            'exists': True,
            'data': existing_data  # Return ALL form data
        })
    
    return jsonify({'exists': False})
@app.route('/get-existing-images', methods=['POST'])
def get_existing_images():
    """Get existing images for a visit"""
    data = request.json
    sol_id = data.get('sol_id')
    visit_no = data.get('visit_no')
    photo_count = int(data.get('photo_count', 0))
    
    images = {}
    
    for i in range(photo_count):
        try:
            blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/photo_{i}.jpg"
            blob = bucket.blob(blob_path)
            
            if blob.exists():
                # Get image as base64
                image_bytes = blob.download_as_bytes()
                base64_image = base64.b64encode(image_bytes).decode('utf-8')
                images[f'photo_image_{i}'] = base64_image
        except Exception as e:
            print(f"Error loading image {i}: {e}")
    
    return jsonify({'images': images})

@app.route('/get-project-name', methods=['POST'])
def get_project_name():
    """Get project name based on SOL ID"""
    data = request.json
    sol_id = data.get('sol_id')
    
    # Search for any visit with this SOL ID
    prefix = f"ICICI_Site_Progress_Report/{sol_id}/"
    blobs = bucket.list_blobs(prefix=prefix, max_results=1)
    
    for blob in blobs:
        if 'data.json' in blob.name:
            existing_data = json.loads(blob.download_as_string())
            return jsonify({
                'found': True,
                'project_name': existing_data.get('project_name', '')
            })
    
    return jsonify({'found': False})

@app.route('/get-image-from-firebase', methods=['POST'])
def get_image_from_firebase():
    """Fetch a single image from Firebase as base64"""
    try:
        data = request.json
        sol_id = data.get('sol_id')
        visit_no = data.get('visit_no')
        image_type = data.get('image_type')  # 'work', 'qual', 'make'
        image_index = data.get('image_index')
        
        blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/{image_type}_{image_index}.jpg"
        blob = bucket.blob(blob_path)
        
        if blob.exists():
            image_bytes = blob.download_as_bytes()
            base64_image = base64.b64encode(image_bytes).decode('utf-8')
            return jsonify({'image_base64': base64_image})
        else:
            return jsonify({'image_base64': None}), 404
            
    except Exception as e:
        print(f"Error fetching image: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/submit-report', methods=['POST'])
def submit_report():
    """Process and save the complete report"""
    try:
        # Get form data
        form_data = request.form.to_dict()
        files = request.files
        
        sol_id = form_data['sol_id']
        visit_no = form_data['visit_no']
        project_name = form_data['project_name']
        
        # Check if this is an update (existing visit)
        blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/data.json"
        existing_blob = bucket.blob(blob_path)
        is_update = existing_blob.exists()
        
        # Load template
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        
        # Fill Progress Report Sheet
        fill_progress_report(wb, form_data)
        
        # Save images to Firebase (before filling photographs sheet)
        save_images_to_firebase(sol_id, visit_no, form_data, files)
        
        # Fill Site Visit Photographs Sheet
        fill_photographs_sheet(wb, form_data, files, sol_id, visit_no)
        
        # Fill Quality and Critical Challenges Sheet
        fill_quality_sheet(wb, form_data)
        
        # Save to memory
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        
       # Upload to Firebase with metadata to enable download in console
        report_blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/report.xlsx"
        report_blob = bucket.blob(report_blob_path)
        
        # ✅ Set metadata with content disposition to enable download
        report_blob.metadata = {
            'firebaseStorageDownloadTokens': str(uuid.uuid4())  # Generate unique token
        }
        report_blob.content_disposition = 'attachment; filename="report.xlsx"'
        
        report_blob.upload_from_file(
            excel_buffer, 
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # ✅ ADD THIS: Make the file publicly accessible
        report_blob.make_public()
        
       
        
        # Save JSON data (OVERWRITE existing)
        json_blob = bucket.blob(blob_path)
        json_blob.upload_from_string(json.dumps(form_data), content_type='application/json')
        
        action = "updated" if is_update else "created"
        print(f"Report {action} successfully for SOL ID: {sol_id}, Visit: {visit_no}")
        
        # Return download
        excel_buffer.seek(0)
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=f"{project_name}_Visit_{visit_no}_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print(f"Submit report error: {str(e)}")
        return jsonify({'error': str(e)}), 500

def save_images_to_firebase(sol_id, visit_no, form_data, files):
    """Save uploaded images to Firebase Storage - MEMORY OPTIMIZED"""
    
    def process_and_upload_image(file_key, image_file, blob_path):
        """Helper to process single image and clean up memory"""
        try:
            img = Image.open(image_file)
            
            # ✅ Aggressive resize to reduce memory
            max_size = (800, 600)
            img.thumbnail(max_size, Image.Resampling.LANCZOS)
            
            img_buffer = io.BytesIO()
            img.save(img_buffer, format='JPEG', quality=75, optimize=True)  # Lower quality
            img_buffer.seek(0)
            
            blob = bucket.blob(blob_path)
            blob.upload_from_file(img_buffer, content_type='image/jpeg')
            
            # ✅ Immediate cleanup
            img.close()
            img_buffer.close()
            del img
            del img_buffer
            gc.collect()  # Force garbage collection
            
            return True
        except Exception as e:
            print(f"Error saving {file_key}: {e}")
            return False
    
    # Work Progress Images
    work_count = int(form_data.get('work_count', 0))
    for i in range(work_count):
        file_key = f'work_image_{i}'
        if file_key in files:
            blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/work_{i}.jpg"
            process_and_upload_image(file_key, files[file_key], blob_path)
    
    # Quality Images
    quality_count = int(form_data.get('quality_count', 0))
    for i in range(quality_count):
        file_key = f'qual_image_{i}'
        if file_key in files:
            blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/qual_{i}.jpg"
            process_and_upload_image(file_key, files[file_key], blob_path)
    
    # Make/Model Images
    make_count = int(form_data.get('make_count', 0))
    for i in range(make_count):
        file_key = f'make_image_{i}'
        if file_key in files:
            blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/make_{i}.jpg"
            process_and_upload_image(file_key, files[file_key], blob_path)
def fill_progress_report(wb, data):
    """Fill the Progress Report sheet"""
    ws = wb['Progress Report']
    
    # Fill header information
    unmerge_and_write(ws, 'B2', data.get('project_name', ''))  # Project name
    
    # Branch Area should show "Branch Area: X" format
    branch_area = data.get('branch_area', '')
    unmerge_and_write(ws, 'G2', f'Branch Area: {branch_area}')   # Branch Area with label
    
    unmerge_and_write(ws, 'B3', data.get('branch_code', ''))   # Branch Code
    unmerge_and_write(ws, 'E3', data.get('date_of_visit', '')) # Date of visit (goes in E3)
    
    # Visit no should show "Visit No: X" format
    visit_no = data.get('visit_no', '')
    unmerge_and_write(ws, 'G3', f'Visit No: {visit_no}')      # Visit no with label
    
    # Get current visit number to know which column to fill
    current_visit = int(data.get('visit_no', '1'))
    
    # Fill Non-Internal Work + Civil items (rows 9-19: all items from a to k)
    # Fill Non-Internal Work + Civil items (rows 9-19: all items from a to k)
    non_internal_and_civil_items = [
        'demolition', 'block_work', 'internal_plaster', 'rcc_wall', 'pcc_work', 'pop_punning',
        'waterproofing', 'flooring', 'dado', 'painting', 'plumbing'
    ]

    for idx, item in enumerate(non_internal_and_civil_items, start=9):
        for visit in range(1, 5):
            cell_key = f"{item}_visit{visit}"
            if cell_key in data and data[cell_key]:
                col = chr(ord('B') + visit)  # Visit 1 = C, Visit 2 = D, etc.
                value = data[cell_key]
                if value and str(value).strip():
                    cell = ws[f'{col}{idx}']
                    cell.value = float(value) / 100
                    cell.number_format = '0%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')  # ✅ RIGHT align

        remark_key = f"{item}_remark"
        if remark_key in data:
            unmerge_and_write(ws, f'G{idx}', data[remark_key])
    
    
    
    # Fill Carpentry Work
    carpentry_items = ['partition', 'paneling', 'door_window', 'false_ceiling', 
                       'loose_furniture', 'alum_skirting', 'window_blind', 'signage']
    for idx, item in enumerate(carpentry_items, start=22):
        for visit in range(1, 5):
            cell_key = f"{item}_visit{visit}"
            if cell_key in data and data[cell_key]:
                col = chr(ord('B') + visit)
                value = data[cell_key]
                if value and str(value).strip():
                    cell = ws[f'{col}{idx}']
                    cell.value = float(value) / 100
                    cell.number_format = '0%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        remark_key = f"{item}_remark"
        if remark_key in data:
            unmerge_and_write(ws, f'G{idx}', data[remark_key])  # Remarks in column G
    
    # Fill Electrical Work
    electrical_items = ['pipe_conduit', 'raceway', 'wiring', 'fixtures', 'main_lt']
    for idx, item in enumerate(electrical_items, start=32):
        for visit in range(1, 5):
            cell_key = f"{item}_visit{visit}"
            if cell_key in data and data[cell_key]:
                col = chr(ord('B') + visit)
                value = data[cell_key]
                if value and str(value).strip():
                    cell = ws[f'{col}{idx}']
                    cell.value = float(value) / 100
                    cell.number_format = '0%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        remark_key = f"{item}_remark"
        if remark_key in data:
            unmerge_and_write(ws, f'G{idx}', data[remark_key])  # Remarks in column G

    # Fill HVAC Work
    hvac_items = ['hvac_indoor', 'hvac_wiring', 'hvac_outdoor']
    for idx, item in enumerate(hvac_items, start=39):  # Adjust row number based on your Excel template
        for visit in range(1, 5):
            cell_key = f"{item}_visit{visit}"
            if cell_key in data and data[cell_key]:
                col = chr(ord('B') + visit)
                value = data[cell_key]
                if value and str(value).strip():
                    cell = ws[f'{col}{idx}']
                    cell.value = float(value) / 100
                    cell.number_format = '0%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        remark_key = f"{item}_remark"
        if remark_key in data:
            unmerge_and_write(ws, f'G{idx}', data[remark_key])
    
    # Fill CMS Work
    cms_items = ['cms_pipe', 'cms_wiring', 'cms_fixture']
    for idx, item in enumerate(cms_items, start=44):  # Adjust row number based on your Excel template
        for visit in range(1, 5):
            cell_key = f"{item}_visit{visit}"
            if cell_key in data and data[cell_key]:
                col = chr(ord('B') + visit)
                value = data[cell_key]
                if value and str(value).strip():
                    cell = ws[f'{col}{idx}']
                    cell.value = float(value) / 100
                    cell.number_format = '0%'
                    cell.alignment = Alignment(horizontal='right', vertical='center')
        
        remark_key = f"{item}_remark"
        if remark_key in data:
            unmerge_and_write(ws, f'G{idx}', data[remark_key])

        
    current_row = 48  # Start at row 48 (after one blank row)

# Get the actual count of "Others" items from form data
    other_count = int(data.get('other_count', 0))

# Define border style FIRST (before using it)
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )

    if other_count > 0:  # Only add "Others" section if user has items
    # Insert section header row at 41
        ws.insert_rows(current_row, 1)
    
    # Add "V" in column A
        ws[f'A{current_row}'] = 'VI'
        ws[f'A{current_row}'].font = Font(name='Calibri', size=11, bold=True)
        ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

        # Add "Others" in column B ONLY (like "Electrical" format)
        ws[f'B{current_row}'] = 'Others'
        ws[f'B{current_row}'].font = Font(name='Calibri', size=11, bold=True)
        ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

        # Add borders to row 41
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            ws[f'{col}{current_row}'].border = thin_border
    
        current_row += 1  # Move to next row for first item (row 42)
    
    # Now add user's "Other" items with a, b, c... numbering
        actual_item_count = 0
        for i in range(other_count):
            item_name = data.get(f'other_item_{i}', '')
            if item_name and item_name.strip():
            # Insert a new row
                ws.insert_rows(current_row, 1)

                # Add serial number (a, b, c, d...)
                ws[f'A{current_row}'] = chr(97 + actual_item_count)
                ws[f'A{current_row}'].font = Font(name='Calibri', size=11)
                ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

                # Write item name in column B (NOT BOLD, NORMAL)
                ws[f'B{current_row}'] = item_name
                ws[f'B{current_row}'].font = Font(name='Calibri', size=11)
                ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center')

                # Fill percentages for ALL visits (1 to 4)
                for visit in range(1, 5):
                    cell_key = f"other_item_{i}_visit{visit}"
                    if cell_key in data and data[cell_key]:
                        col = chr(ord('B') + visit)  # Visit 1=C, 2=D, 3=E, 4=F
                        value = data[cell_key]
                        if value and str(value).strip():
                            cell = ws[f'{col}{current_row}']
                            cell.value = float(value) / 100
                            cell.number_format = '0%'
                            cell.alignment = Alignment(horizontal='right', vertical='center')
                    
                    
            
            # Fill remark in column G (NOT BOLD, NORMAL)
                remark_key = f"other_item_{i}_remark"
                if remark_key in data:
                    ws[f'G{current_row}'] = data[remark_key]
                    ws[f'G{current_row}'].font = Font(name='Calibri', size=11)
                    ws[f'G{current_row}'].alignment = Alignment(horizontal='left', vertical='center')
            
            # Add borders to this row
                for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
                    ws[f'{col}{current_row}'].border = thin_border
            
                actual_item_count += 1
                current_row += 1

# Leave ONE BLANK ROW after Others section
    current_row += 1

# NOW insert "Prepared By" and "Checked By" row
    ws.insert_rows(current_row, 1)

    # Merge E and F for "Prepared By:"
    ws.merge_cells(f'E{current_row}:F{current_row}')
    ws[f'E{current_row}'] = 'Prepared By:'
    ws[f'E{current_row}'].font = Font(name='Calibri', size=11, bold=True)
    ws[f'E{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

    # "Checked By:" in column G
    ws[f'G{current_row}'] = 'Checked By:'
    ws[f'G{current_row}'].font = Font(name='Calibri', size=11, bold=True)
    ws[f'G{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

    # Add borders to Prepared By row
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        ws[f'{col}{current_row}'].border = thin_border

    # Add one more row for actual names (Ram and Tanush)
    current_row += 1
    ws.insert_rows(current_row, 1)

    # Merge E and F for prepared by name
    ws.merge_cells(f'E{current_row}:F{current_row}')
    ws[f'E{current_row}'] = data.get('prepared_by', '')
    ws[f'E{current_row}'].font = Font(name='Calibri', size=11, bold=True)
    ws[f'E{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

    # Checked by name in column G
    ws[f'G{current_row}'] = data.get('checked_by', '')
    ws[f'G{current_row}'].font = Font(name='Calibri', size=11, bold=True)
    ws[f'G{current_row}'].alignment = Alignment(horizontal='center', vertical='center')

    # Add borders to names row (FINAL ROW OF TABLE)
    for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
        ws[f'{col}{current_row}'].border = thin_border

    # ✅✅✅ NEW CODE: Remove all borders and formatting from rows below ✅✅✅
    # Clear any existing formatting from rows after the table ends
    max_clear_row = current_row + 10000  # Clear next 20 rows to be safe

    for clear_row in range(current_row + 1, max_clear_row + 1):
        for col in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            cell = ws[f'{col}{clear_row}']
            cell.border = Border()  # Remove all borders
            cell.value = None  # Clear any value
            cell.font = Font(name='Calibri', size=11)  # Reset to default font
            cell.alignment = Alignment(horizontal='general', vertical='bottom')  # Default alignment

    # Unmerge any merged cells below the table
    for merged_range in list(ws.merged_cells.ranges):
        # If merged range starts after our table, unmerge it
        if merged_range.min_row > current_row:
            ws.unmerge_cells(str(merged_range))



def fill_photographs_sheet(wb, data, files, sol_id, visit_no):
    """Fill Site visit Photographs - DYNAMIC INSERTION (NO HARDCODED ROWS)"""
    ws = wb['Site visit Photographs']
    
    # Define border style
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    # ✅ SET EXACT COLUMN WIDTHS (per user specifications)
    ws.column_dimensions['A'].width = 31.89  # Sr. No. (294 px)
    ws.column_dimensions['B'].width = 50.22  # Critical Milestones (459 px)
    ws.column_dimensions['C'].width = 50.22  # Description (459 px)
    ws.column_dimensions['D'].width = 60  # Photographs (540 px)
    ws.column_dimensions['E'].width = 46.89  # Remarks (429 px)
    
    # ✅ DELETE ALL ROWS FROM ROW 9 ONWARDS (clear template placeholders)
    max_row = ws.max_row
    if max_row >= 9:
        ws.delete_rows(9, max_row - 8)
    
    current_row = 9  # Start fresh after "Work Progress" header (row 8)
    
    # ====================
    # 1. WORK PROGRESS SECTION
    # ====================
    work_count = int(data.get('work_count', 0))
    
    if work_count > 0:
        for i in range(work_count):
            ws.insert_rows(current_row, 1)
            ws.row_dimensions[current_row].height = 249.60
            
            # Add borders
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws[f'{col}{current_row}'].border = thin_border
            
            # Serial Number
            ws[f'A{current_row}'].value = i + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=20, bold=True)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            # Milestone
            ws[f'B{current_row}'].value = data.get(f'work_milestone_{i}', '')
            ws[f'B{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # Description
            ws[f'C{current_row}'].value = data.get(f'work_desc_{i}', '')
            ws[f'C{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'C{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # Remarks
            ws[f'E{current_row}'].value = data.get(f'work_remark_{i}', '')
            ws[f'E{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'E{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # IMAGE
            file_key = f'work_image_{i}'
            img_stream = None
            
            if file_key in files and files[file_key].filename != '':
                try:
                    files[file_key].seek(0)
                    img_stream = files[file_key]
                except:
                    pass
            
            if img_stream is None:
                try:
                    blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/work_{i}.jpg"
                    blob = bucket.blob(blob_path)
                    if blob.exists():
                        img_bytes = blob.download_as_bytes()
                        img_stream = io.BytesIO(img_bytes)
                except Exception as e:
                    print(f"Error fetching work image {i}: {e}")
            
            if img_stream:
                try:
                    pil_img = PILImage.open(img_stream)
                    pil_img = pil_img.resize((430, 320), PILImage.Resampling.LANCZOS)
                    
                    img_byte_arr = io.BytesIO()
                    pil_img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    
                    xl_img = XLImage(img_byte_arr)
                    xl_img.anchor = f'D{current_row}'
                    ws.add_image(xl_img)
                    
                    ws[f'D{current_row}'].value = ""
                    pil_img.close()
                    img_byte_arr.close()
                    del pil_img
                    del img_byte_arr
                except Exception as e:
                    print(f"Error inserting work image {i}: {e}")
                    ws[f'D{current_row}'].value = "Error loading image"
                    ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            else:
                ws[f'D{current_row}'].value = "[No Image]"
                ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            current_row += 1
    
    # ====================
    # 2. QUALITY OBSERVATION SECTION
    # ====================
    quality_count = int(data.get('quality_count', 0))
    
    if quality_count > 0:
        # Insert section title row
        ws.insert_rows(current_row, 1)
        ws.merge_cells(f'B{current_row}:E{current_row}')
        ws[f'B{current_row}'] = 'Quality Observation'
        ws.row_dimensions[current_row].height = 28.20  # Title row: 49 pixels
        ws[f'B{current_row}'].font = Font(name='Calibri', size=20, bold=True)
        ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}{current_row}'].border = thin_border
        
        current_row += 1
        
        # Insert header row
        ws.insert_rows(current_row, 1)
        ws[f'A{current_row}'] = 'Sr. No.'
        ws[f'B{current_row}'] = 'Items'
        ws[f'C{current_row}'] = 'Description'
        ws[f'D{current_row}'] = 'Photographs'
        ws[f'E{current_row}'] = 'Remarks'
        
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}{current_row}'].font = Font(name='Calibri', size=20, bold=True)
            ws[f'{col}{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'{col}{current_row}'].border = thin_border
        
        current_row += 1
        
        # Insert quality entries
        for i in range(quality_count):
            ws.insert_rows(current_row, 1)
            ws.row_dimensions[current_row].height = 249.60
            
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws[f'{col}{current_row}'].border = thin_border
            
            ws[f'A{current_row}'].value = i + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=20, bold=True)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            ws[f'B{current_row}'].value = data.get(f'qual_item_{i}', '')
            ws[f'B{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            ws[f'C{current_row}'].value = data.get(f'qual_desc_{i}', '')
            ws[f'C{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'C{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            ws[f'E{current_row}'].value = data.get(f'qual_remark_{i}', '')
            ws[f'E{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'E{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # IMAGE
            file_key = f'qual_image_{i}'
            img_stream = None
            
            if file_key in files and files[file_key].filename != '':
                try:
                    files[file_key].seek(0)
                    img_stream = files[file_key]
                except:
                    pass
            
            if img_stream is None:
                try:
                    blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/qual_{i}.jpg"
                    blob = bucket.blob(blob_path)
                    if blob.exists():
                        img_bytes = blob.download_as_bytes()
                        img_stream = io.BytesIO(img_bytes)
                except Exception as e:
                    print(f"Error fetching qual image {i}: {e}")
            
            if img_stream:
                try:
                    pil_img = PILImage.open(img_stream)
                    pil_img = pil_img.resize((430, 320), PILImage.Resampling.LANCZOS)
                    
                    img_byte_arr = io.BytesIO()
                    pil_img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    
                    xl_img = XLImage(img_byte_arr)
                    xl_img.anchor = f'D{current_row}'
                    ws.add_image(xl_img)
                    
                    ws[f'D{current_row}'].value = ""
                    pil_img.close()
                    img_byte_arr.close()
                    del pil_img
                    del img_byte_arr
                except Exception as e:
                    print(f"Error inserting qual image {i}: {e}")
                    ws[f'D{current_row}'].value = "Error loading image"
                    ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            else:
                ws[f'D{current_row}'].value = "[No Image]"
                ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            current_row += 1
    
    # ====================
    # 3. MAKE/MODEL INSPECTION SECTION
    # ====================
    make_count = int(data.get('make_count', 0))
    
    if make_count > 0:
        # Insert section title
        ws.insert_rows(current_row, 1)
        ws.merge_cells(f'B{current_row}:E{current_row}')
        ws[f'B{current_row}'] = 'Make / Model Inspection'
        ws.row_dimensions[current_row].height = 28.20  # Title row: 49 pixels
        ws[f'B{current_row}'].font = Font(name='Calibri', size=20, bold=True)
        ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
        
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}{current_row}'].border = thin_border
        
        current_row += 1
        
        # Insert header row
        ws.insert_rows(current_row, 1)
        ws[f'A{current_row}'] = 'Sr. No.'
        ws[f'B{current_row}'] = 'Items'
        ws[f'C{current_row}'] = 'Make/Model Observed'
        ws[f'D{current_row}'] = 'Photographs'
        ws[f'E{current_row}'] = 'Remarks'
        
        for col in ['A', 'B', 'C', 'D', 'E']:
            ws[f'{col}{current_row}'].font = Font(name='Calibri', size=20, bold=True)
            ws[f'{col}{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'{col}{current_row}'].border = thin_border
        
        current_row += 1
        
        # Insert make/model entries
        for i in range(make_count):
            ws.insert_rows(current_row, 1)
            ws.row_dimensions[current_row].height = 249.60
            
            for col in ['A', 'B', 'C', 'D', 'E']:
                ws[f'{col}{current_row}'].border = thin_border
            
            ws[f'A{current_row}'].value = i + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=20, bold=True)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            ws[f'B{current_row}'].value = data.get(f'make_item_{i}', '')
            ws[f'B{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            ws[f'C{current_row}'].value = data.get(f'make_observed_{i}', '')
            ws[f'C{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'C{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            ws[f'E{current_row}'].value = data.get(f'make_remark_{i}', '')
            ws[f'E{current_row}'].font = Font(name='Calibri', size=24)
            ws[f'E{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            
            # IMAGE
            file_key = f'make_image_{i}'
            img_stream = None
            
            if file_key in files and files[file_key].filename != '':
                try:
                    files[file_key].seek(0)
                    img_stream = files[file_key]
                except:
                    pass
            
            if img_stream is None:
                try:
                    blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/make_{i}.jpg"
                    blob = bucket.blob(blob_path)
                    if blob.exists():
                        img_bytes = blob.download_as_bytes()
                        img_stream = io.BytesIO(img_bytes)
                except Exception as e:
                    print(f"Error fetching make image {i}: {e}")
            
            if img_stream:
                try:
                    pil_img = PILImage.open(img_stream)
                    pil_img = pil_img.resize((430, 320), PILImage.Resampling.LANCZOS)
                    
                    img_byte_arr = io.BytesIO()
                    pil_img.save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    
                    xl_img = XLImage(img_byte_arr)
                    xl_img.anchor = f'D{current_row}'
                    ws.add_image(xl_img)
                    
                    ws[f'D{current_row}'].value = ""
                    pil_img.close()
                    img_byte_arr.close()
                    del pil_img
                    del img_byte_arr
                except Exception as e:
                    print(f"Error inserting make image {i}: {e}")
                    ws[f'D{current_row}'].value = "Error loading image"
                    ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            else:
                ws[f'D{current_row}'].value = "[No Image]"
                ws[f'D{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            
            current_row += 1


def fill_quality_sheet(wb, data):
    """Fill the Quality and Critical Challenges sheet - REMOVES EMPTY ROWS"""
    ws = wb['Quality and Critical challenges']
    
    # Define border style
    thin_border = Border(
        left=Side(style='thin', color='000000'),
        right=Side(style='thin', color='000000'),
        top=Side(style='thin', color='000000'),
        bottom=Side(style='thin', color='000000')
    )
    
    # Define Green color (Accent 6, Lighter 80%) - RGB: E2EFDA
    green_fill = PatternFill(start_color='E2EFDA', end_color='E2EFDA', fill_type='solid')
    
    # Fill header information (rows 2-9) with Calibri 12
    ws['B2'] = data.get('quality_project_title', '')
    ws['B2'].font = Font(name='Calibri', size=12)
    
    ws['B3'] = data.get('quality_branch_area', '')
    ws['B3'].font = Font(name='Calibri', size=12)
    
    ws['B4'] = data.get('civil_vendor', '')
    ws['B4'].font = Font(name='Calibri', size=12)
    
    ws['B5'] = data.get('hvac_vendor', '')
    ws['B5'].font = Font(name='Calibri', size=12)
    
    ws['B6'] = data.get('project_start_date', '')
    ws['B6'].font = Font(name='Calibri', size=12)
    
    ws['B7'] = data.get('planned_handover_date', '')
    ws['B7'].font = Font(name='Calibri', size=12)
    
    ws['B8'] = data.get('actual_handover_date', '')
    ws['B8'].font = Font(name='Calibri', size=12)
    
    ws['B9'] = data.get('quality_date_of_visit', '')
    ws['B9'].font = Font(name='Calibri', size=12)
    
    # Start building dynamic sections
    current_row = 10
    
    # ===== SECTION HEADER: Quality Observations =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Quality Observations'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Quality Observations data
    filled_count = 0
    for i in range(6):
        obs_value = data.get(f'quality_observation_{i}', '').strip()
        if obs_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = obs_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            ws.row_dimensions[current_row].height = None  # Auto height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
    
    # ===== SECTION HEADER: Site Delay Reasons =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Site Delay Reasons'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Site Delay Reasons data
    filled_count = 0
    for i in range(6):
        delay_value = data.get(f'site_delay_reason_{i}', '').strip()
        if delay_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = delay_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            row_height = calculate_row_height(delay_value, font_size=12, cell_width=50)
            ws.row_dimensions[current_row].height = row_height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
    
    # ===== SECTION HEADER: Collaborative Challenges/Inputs =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Collaborative Challenges/Inputs'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Collaborative Challenges data
    filled_count = 0
    for i in range(6):
        challenge_value = data.get(f'collaborative_challenge_{i}', '').strip()
        if challenge_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = challenge_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            row_height = calculate_row_height(challenge_value, font_size=12, cell_width=50)
            ws.row_dimensions[current_row].height = row_height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
    
    # ===== SECTION HEADER: Criticalities =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Criticalities'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Criticalities data
    filled_count = 0
    for i in range(6):
        critical_value = data.get(f'criticality_{i}', '').strip()
        if critical_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = critical_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            row_height = calculate_row_height(critical_value, font_size=12, cell_width=50)
            ws.row_dimensions[current_row].height = row_height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
    
    # ===== SECTION HEADER: Hindrances =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Hindrances'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Hindrances data
    filled_count = 0
    for i in range(6):
        hindrance_value = data.get(f'hindrance_{i}', '').strip()
        if hindrance_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = hindrance_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            row_height = calculate_row_height(hindrance_value, font_size=12, cell_width=50)
            ws.row_dimensions[current_row].height = row_height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
    
    # ===== SECTION HEADER: Others =====
    ws[f'A{current_row}'] = 'Sr. No.'
    ws[f'A{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'A{current_row}'].fill = green_fill
    ws[f'A{current_row}'].border = thin_border
    
    ws.merge_cells(f'B{current_row}:D{current_row}')
    ws[f'B{current_row}'] = 'Others'
    ws[f'B{current_row}'].font = Font(name='Times New Roman', size=12, bold=True)
    ws[f'B{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
    ws[f'B{current_row}'].fill = green_fill
    ws[f'B{current_row}'].border = thin_border
    ws[f'C{current_row}'].border = thin_border
    ws[f'D{current_row}'].border = thin_border
    
    current_row += 1
    
    # Fill Others data
    filled_count = 0
    for i in range(6):
        other_value = data.get(f'other_{i}', '').strip()
        if other_value:
            ws[f'A{current_row}'] = filled_count + 1
            ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'A{current_row}'].alignment = Alignment(horizontal='center', vertical='center')
            ws[f'A{current_row}'].border = thin_border
            
            ws[f'B{current_row}'] = other_value
            ws[f'B{current_row}'].font = Font(name='Calibri', size=12)
            ws[f'B{current_row}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
            row_height = calculate_row_height(other_value, font_size=12, cell_width=50)
            ws.row_dimensions[current_row].height = row_height
            
            for col in ['A', 'B', 'C', 'D']:
                ws[f'{col}{current_row}'].border = thin_border
            
            filled_count += 1
            current_row += 1
    
    if filled_count == 0:
        ws[f'A{current_row}'] = 1
        ws[f'A{current_row}'].font = Font(name='Calibri', size=12)
        ws[f'B{current_row}'] = ''
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{current_row}'].border = thin_border
        current_row += 1
def update_visit_percentages(sol_id, current_visit, data):
    """Update percentages across all visits for the same SOL ID"""
    # This function copies percentages from current visit to previous visits
    # and retrieves percentages from previous visits to add to current visit
    
    for visit in range(1, 5):
        if visit != int(current_visit):
            try:
                blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit}/report.xlsx"
                blob = bucket.blob(blob_path)
                
                if blob.exists():
                    # Download and update the existing visit file
                    excel_buffer = io.BytesIO()
                    blob.download_to_file(excel_buffer)
                    excel_buffer.seek(0)
                    
                    wb = openpyxl.load_workbook(excel_buffer)
                    # Update with current visit percentages
                    # Implementation details...
                    
                    # Save back
                    output_buffer = io.BytesIO()
                    wb.save(output_buffer)
                    output_buffer.seek(0)
                    blob.upload_from_file(output_buffer, content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            except Exception as e:
                print(f"Error updating visit {visit}: {e}")
                pass

@app.route('/generate-report', methods=['POST'])
def generate_report():
    try:
        print("=" * 50)
        print("📥 STARTING REPORT GENERATION")
        print("=" * 50)
        
        # 1. Get form data
        form_data = request.form.to_dict()
        files = request.files
        
        sol_id = form_data.get('sol_id')
        visit_no = form_data.get('visit_no')
        project_name = form_data.get('project_name')
        
        print(f"📋 SOL ID: {sol_id}")
        print(f"📋 Visit: {visit_no}")
        print(f"📋 Project: {project_name}")
        
        if not sol_id or not visit_no or not project_name:
            print("❌ Missing required fields")
            return jsonify({'error': 'Missing required fields (SOL ID, Visit No, or Project Name)'}), 400

        # 2. Count images
        work_count = int(form_data.get('work_count', 0))
        quality_count = int(form_data.get('quality_count', 0))
        make_count = int(form_data.get('make_count', 0))
        
        print(f"📸 Images: Work={work_count}, Quality={quality_count}, Make={make_count}")
        
        # ✅ SKIP FIREBASE UPLOAD FOR NOW (to test if this is the issue)
        # save_images_to_firebase(sol_id, visit_no, form_data, files)
        
        # 3. Load template
        print("📄 Loading Excel template...")
        if not os.path.exists(TEMPLATE_PATH):
            print(f"❌ Template not found at: {TEMPLATE_PATH}")
            return jsonify({'error': 'Excel template not found'}), 500
            
        wb = openpyxl.load_workbook(TEMPLATE_PATH)
        print("✅ Template loaded")
        
        # 4. Fill sheets
        print("✍️ Filling Progress Report...")
        fill_progress_report(wb, form_data)
        print("✅ Progress Report done")
        
        print("📸 Filling Photographs Sheet...")
        fill_photographs_sheet(wb, form_data, files, sol_id, visit_no)
        print("✅ Photographs done")
        
        print("✅ Filling Quality Sheet...")
        fill_quality_sheet(wb, form_data)
        print("✅ Quality Sheet done")
        
        # 5. Save to memory
        print("💾 Saving Excel to memory...")
        excel_buffer = io.BytesIO()
        wb.save(excel_buffer)
        excel_buffer.seek(0)
        print("✅ Excel saved to buffer")
        
        # 6. Save JSON to Firebase
        print("☁️ Saving JSON to Firebase...")
        try:
            blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/data.json"
            bucket.blob(blob_path).upload_from_string(json.dumps(form_data), content_type='application/json')
            print("✅ JSON saved to Firebase")
        except Exception as fb_error:
            print(f"⚠️ Firebase JSON upload failed: {fb_error}")
            # Continue anyway - don't fail the whole request
        
        # 7. Cleanup
        wb.close()
        print("✅ Workbook closed")
        
        print("=" * 50)
        print("✅ REPORT GENERATION COMPLETE")
        print("=" * 50)
        
        # 8. Return file
        return send_file(
            excel_buffer,
            as_attachment=True,
            download_name=f"{project_name}_Visit_{visit_no}_Draft.xlsx",
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
    except Exception as e:
        print("=" * 50)
        print("❌ CRITICAL ERROR IN GENERATE-REPORT")
        print("=" * 50)
        print(f"Error Type: {type(e).__name__}")
        print(f"Error Message: {str(e)}")
        
        import traceback
        print("Full Traceback:")
        print(traceback.format_exc())
        print("=" * 50)
        
        return jsonify({
            'error': f'Server error: {str(e)}',
            'type': type(e).__name__
        }), 500

@app.route('/upload-final-report', methods=['POST'])
def upload_final_report():
    """Upload user's completed Excel file with photos to Firebase"""
    try:
        if 'excel_file' not in request.files:
            return jsonify({'error': 'No file uploaded'}), 400
        
        excel_file = request.files['excel_file']
        sol_id = request.form.get('sol_id')
        visit_no = request.form.get('visit_no')
        project_name = request.form.get('project_name')
        form_data_json = request.form.get('form_data_json')
        
        if not excel_file or not sol_id or not visit_no:
            return jsonify({'error': 'Missing required data'}), 400
        
        # Read the uploaded Excel file
        excel_buffer = io.BytesIO()
        excel_file.save(excel_buffer)
        excel_buffer.seek(0)
        
        # Upload Excel to Firebase
        report_blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/report_final.xlsx"
        report_blob = bucket.blob(report_blob_path)
        
        report_blob.metadata = {
            'firebaseStorageDownloadTokens': str(uuid.uuid4())
        }
        report_blob.content_disposition = f'attachment; filename="{project_name}_Visit_{visit_no}_Final.xlsx"'
        
        excel_buffer.seek(0)
        report_blob.upload_from_file(
            excel_buffer, 
            content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        report_blob.make_public()
        
        # Save/update the form data JSON
        if form_data_json:
            json_blob_path = f"ICICI_Site_Progress_Report/{sol_id}/Visit_{visit_no}/data.json"
            json_blob = bucket.blob(json_blob_path)
            json_blob.upload_from_string(form_data_json, content_type='application/json')
        
        print(f"✅ Final report uploaded for SOL ID: {sol_id}, Visit: {visit_no}")
        
        return jsonify({
            'success': True,
            'message': 'Final report saved to Firebase successfully',
            'firebase_url': report_blob.public_url
        })
        
    except Exception as e:
        print(f"Upload error: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(debug=False, host='0.0.0.0', port=port)
