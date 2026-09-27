import os
import io
import re
import json
import base64
import time
import urllib.request
import urllib.parse
from collections import defaultdict
from flask import Flask, render_template, request, jsonify, send_file, Response
from flask_cors import CORS
import pandas as pd
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import ddddocr
import tempfile

app = Flask(__name__, template_folder='templates', static_folder='static')
app.config['TEMPLATES_AUTO_RELOAD'] = True
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max file size
CORS(app)

UPLOAD_FOLDER = os.path.join(tempfile.gettempdir(), 'dtdc_uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs('templates', exist_ok=True)
os.makedirs('static', exist_ok=True)

USER_AGENT = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36'

# Initialize OCR engine once at startup
print("Initializing Automated Captcha OCR Engine...")
ocr = ddddocr.DdddOcr(show_ad=False)
print("OCR Engine Ready!")


def fetch_dtdc_captcha():
    """Fetch a fresh captcha from DTDC."""
    url = f"https://www.dtdc.com/wp-json/custom/v1/generate-captcha?t={int(os.times().elapsed * 1000)}"
    req = urllib.request.Request(
        url,
        headers={
            'User-Agent': USER_AGENT,
            'Referer': 'https://www.dtdc.com/track-your-shipment/'
        }
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode('utf-8'))
        return {
            'key': data.get('key'),
            'image': data.get('image')
        }


def validate_dtdc_captcha(key, value):
    """Validate captcha and obtain DTDC Track Token."""
    url = "https://www.dtdc.com/wp-json/custom/v1/captcha/validate"
    payload = json.dumps({"captchaKey": key, "captchaValue": value}).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=payload,
        headers={
            'Content-Type': 'application/json',
            'User-Agent': USER_AGENT,
            'Referer': 'https://www.dtdc.com/track-your-shipment/'
        }
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        return json.loads(resp.read().decode('utf-8'))


def get_verified_token_auto(max_attempts=8):
    """Automatically fetch, OCR-solve, and validate captcha with DTDC."""
    for attempt in range(max_attempts):
        try:
            c = fetch_dtdc_captcha()
            img_bytes = base64.b64decode(c['image'])
            pred = ocr.classification(img_bytes)
            cleaned = pred.upper().replace(' ', '')
            
            # Validate with DTDC
            val_res = validate_dtdc_captcha(c['key'], cleaned)
            if val_res.get('success') and val_res.get('token'):
                return val_res.get('token')
            time.sleep(0.2)
        except Exception:
            pass
    return None


def pull_dtdc_details(numbers, token):
    """Fetch encrypted payload for list of tracking numbers."""
    url = "https://www.dtdc.com/wp-json/custom/v1/tracking/pull-details"
    payload = json.dumps({"trackNumbers": numbers}).encode('utf-8')
    headers = {
        'Content-Type': 'application/json',
        'User-Agent': USER_AGENT,
        'Referer': 'https://www.dtdc.com/track-your-shipment/'
    }
    if token:
        headers['X-DTDC-Track-Token'] = token

    req = urllib.request.Request(url, data=payload, headers=headers)
    with urllib.request.urlopen(req, timeout=30) as resp:
        return json.loads(resp.read().decode('utf-8'))


def post_to_trackshipment(redirect_url, payload_d):
    """POST payload to trackshipment and retrieve rendered HTML."""
    data = urllib.parse.urlencode({'d': payload_d}).encode('utf-8')
    req = urllib.request.Request(
        redirect_url,
        data=data,
        headers={
            'User-Agent': USER_AGENT,
            'Referer': 'https://www.dtdc.com/track-your-shipment/',
            'Content-Type': 'application/x-www-form-urlencoded'
        }
    )
    with urllib.request.urlopen(req, timeout=35) as resp:
        return resp.read().decode('utf-8', errors='ignore')


def parse_dtdc_html(html_content, requested_numbers=None):
    """Parse DTDC HTML tracking page and extract structured tracking details."""
    soup = BeautifulSoup(html_content, 'html.parser')
    results = []

    # CASE 1: Multiple shipments view (accordion items)
    cards = soup.find_all('div', class_=re.compile(r'accordion-item'))
    
    if cards:
        for card in cards:
            item = {
                "awb": "",
                "reference_no": "",
                "edd": "",
                "status": "",
                "status_type": "pending",
                "origin": "",
                "destination": "",
                "latest_update": "",
                "latest_time": ""
            }

            header = card.find('div', class_=re.compile(r'accordion-header|accord-row'))
            if header:
                labels = header.find_all('p', class_='label')
                for lbl in labels:
                    txt = lbl.get_text(strip=True).lower()
                    val_el = lbl.find_next_sibling('p', class_='value')
                    val = val_el.get_text(strip=True) if val_el else ""

                    if 'shipment' in txt or 'awb' in txt or 'consignment' in txt:
                        item['awb'] = val
                    elif 'reference' in txt:
                        item['reference_no'] = val
                    elif 'edd' in txt:
                        item['edd'] = val

                status_el = header.find('span', class_=re.compile(r'status'))
                if status_el:
                    item['status'] = re.sub(r'\s+', ' ', status_el.get_text(strip=True))
                    classes = status_el.get('class', [])
                    for c in classes:
                        if c != 'status':
                            item['status_type'] = c.lower()

            body = card.find('div', class_=re.compile(r'accordion-body'))
            if body:
                loc_items = body.find_all('div', class_='location-item')
                for loc in loc_items:
                    lbl = loc.find('span', class_='location-label')
                    val = loc.find('span', class_='location-value')
                    if lbl and val:
                        l_txt = lbl.get_text(strip=True).lower()
                        v_txt = re.sub(r'\s+', ' ', val.get_text(strip=True))
                        if 'origin' in l_txt:
                            item['origin'] = v_txt
                        elif 'destination' in l_txt:
                            item['destination'] = v_txt

                tl_first = body.find('div', class_=re.compile(r'timeline-item'))
                if tl_first:
                    time_span = tl_first.find('span', class_=re.compile(r'timeline-date'))
                    desc_div = tl_first.find('div', class_=re.compile(r'timeline-description'))
                    status_span = tl_first.find('span', class_=re.compile(r'timeline-status'))
                    if time_span:
                        item['latest_time'] = time_span.get_text(strip=True)
                    if desc_div:
                        item['latest_update'] = desc_div.get_text(strip=True)
                    elif status_span:
                        item['latest_update'] = status_span.get_text(strip=True)

            if item['awb']:
                results.append(item)

    # CASE 2: Single shipment view
    if not results:
        awb_sec = soup.find('div', class_=re.compile(r'awb-section'))
        awb_val = ""
        if awb_sec:
            m = re.search(r'([A-Za-z0-9]{8,15})', awb_sec.get_text(strip=True))
            if m:
                awb_val = m.group(1)

        if not awb_val and requested_numbers:
            awb_val = requested_numbers[0]

        if awb_val:
            item = {
                "awb": awb_val,
                "reference_no": "-",
                "edd": "-",
                "status": "",
                "status_type": "pending",
                "origin": "-",
                "destination": "-",
                "latest_update": "-",
                "latest_time": "-"
            }

            ref_el = soup.find(class_=re.compile(r'ref-number|refNumber'))
            if ref_el:
                item['reference_no'] = ref_el.get_text(strip=True)

            badge = soup.select_one('div.status-badge:not(.status-badge-main)')
            if not badge:
                badge = soup.find('span', class_=re.compile(r'timeline-status'))
            if badge:
                item['status'] = re.sub(r'\s+', ' ', badge.get_text(strip=True))

            edd_title = soup.find(class_=re.compile(r'main-status-title|edd-date'))
            if edd_title:
                edd_text = edd_title.get_text(strip=True)
                m_edd = re.search(r'(?:Delivery by|Delivered on)\s*(.*)', edd_text, re.I)
                if m_edd:
                    item['edd'] = m_edd.group(1).strip()
                else:
                    item['edd'] = edd_text

            loc_items = soup.find_all('div', class_=re.compile(r'location-item'))
            for loc in loc_items:
                lbl = loc.find(class_=re.compile(r'location-label'))
                val = loc.find(class_=re.compile(r'location-value'))
                if lbl and val:
                    l_txt = lbl.get_text(strip=True).lower()
                    v_txt = re.sub(r'\s+', ' ', val.get_text(strip=True))
                    if 'origin' in l_txt:
                        item['origin'] = v_txt
                    elif 'destination' in l_txt:
                        item['destination'] = v_txt

            timeline = soup.find(class_=re.compile(r'timeline'))
            if timeline:
                tl_first = timeline.find('div', class_=re.compile(r'timeline-item'))
                if tl_first:
                    time_span = tl_first.find('span', class_=re.compile(r'timeline-date'))
                    desc_div = tl_first.find('div', class_=re.compile(r'timeline-description'))
                    status_span = tl_first.find('span', class_=re.compile(r'timeline-status'))
                    if time_span:
                        item['latest_time'] = time_span.get_text(strip=True)
                    if desc_div:
                        item['latest_update'] = desc_div.get_text(strip=True)
                    elif status_span:
                        item['latest_update'] = status_span.get_text(strip=True)

            if item['status'] or item['origin'] != '-':
                results.append(item)

    # Categorize all parsed items
    for item in results:
        st_lower = item['status'].lower()
        if 'delivered' in st_lower:
            item['category'] = 'Delivered'
        elif 'out for delivery' in st_lower:
            item['category'] = 'Out for Delivery'
        elif any(k in st_lower for k in ['in transit', 'dispatched', 'arrived', 'transit', 'reached', 'on the way']):
            item['category'] = 'In Transit'
        elif any(k in st_lower for k in ['booked', 'picked up', 'manifest']):
            item['category'] = 'Booked'
        elif any(k in st_lower for k in ['undelivered', 'rto', 'return', 'failed', 'cancelled', 'exception']):
            item['category'] = 'Issue/RTO'
        else:
            item['category'] = 'In Transit' if item['status'] else 'Not Found'

    if requested_numbers:
        found_awbs = {r['awb'].upper() for r in results}
        for num in requested_numbers:
            if num.upper() not in found_awbs:
                results.append({
                    "awb": num,
                    "reference_no": "-",
                    "edd": "-",
                    "status": "No record / Invalid AWB",
                    "status_type": "not_found",
                    "category": "Not Found",
                    "origin": "-",
                    "destination": "-",
                    "latest_update": "No tracking record found on DTDC",
                    "latest_time": "-"
                })

    return results


@app.after_request
def add_header(response):
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, post-check=0, pre-check=0, max-age=0'
    response.headers['Pragma'] = 'no-cache'
    response.headers['Expires'] = '-1'
    return response


@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/captcha', methods=['GET'])
def get_captcha():
    try:
        data = fetch_dtdc_captcha()
        return jsonify({'success': True, 'key': data['key'], 'image': data['image']})
    except Exception as e:
        return jsonify({'success': False, 'error': f'Failed to fetch captcha: {str(e)}'}), 500


@app.route('/api/track-stream', methods=['POST'])
def track_stream():
    """Stream live tracking results in safe 25-number chunks with automated captcha solving."""
    data = request.json or {}
    raw_numbers = data.get('numbers', [])

    clean_nums = []
    for n in raw_numbers:
        cleaned = re.sub(r'[^a-zA-Z0-9]', '', str(n).strip())
        if cleaned and cleaned not in clean_nums:
            clean_nums.append(cleaned)

    if not clean_nums:
        return jsonify({'success': False, 'error': 'No valid tracking numbers provided.'}), 400

    def generate_events():
        total = len(clean_nums)
        CHUNK_SIZE = 25
        completed = 0
        all_accumulated = []

        for i in range(0, total, CHUNK_SIZE):
            chunk = clean_nums[i:i + CHUNK_SIZE]
            batch_num = (i // CHUNK_SIZE) + 1
            total_batches = (total + CHUNK_SIZE - 1) // CHUNK_SIZE

            # Step 1: Automatically solve captcha & get verified token
            yield f"data: {json.dumps({'type': 'status', 'message': f'Verifying security token for batch {batch_num} of {total_batches}...', 'completed': completed, 'total': total})}\n\n"

            token = get_verified_token_auto(max_attempts=8)
            if not token:
                yield f"data: {json.dumps({'type': 'error', 'message': f'Could not verify security token for batch {batch_num}. Retrying...'})}\n\n"
                token = get_verified_token_auto(max_attempts=8)
                if not token:
                    continue

            # Step 2: Query DTDC for this chunk
            yield f"data: {json.dumps({'type': 'status', 'message': f'Tracking batch {batch_num} of {total_batches} ({len(chunk)} parcels)...', 'completed': completed, 'total': total})}\n\n"

            try:
                pull_res = pull_dtdc_details(chunk, token)
                if pull_res.get('success') and pull_res.get('redirect') and pull_res.get('payload'):
                    html = post_to_trackshipment(pull_res['redirect'], pull_res['payload'])
                    chunk_results = parse_dtdc_html(html, requested_numbers=chunk)
                else:
                    chunk_results = parse_dtdc_html("", requested_numbers=chunk)
            except Exception as e:
                print(f"Error tracking chunk: {e}")
                chunk_results = parse_dtdc_html("", requested_numbers=chunk)

            completed += len(chunk)
            all_accumulated.extend(chunk_results)

            # Yield progress event with new batch results
            payload = {
                'type': 'progress',
                'completed': completed,
                'total': total,
                'percent': int((completed / total) * 100),
                'new_results': chunk_results
            }
            yield f"data: {json.dumps(payload)}\n\n"
            time.sleep(0.3)

        # Finished all batches
        yield f"data: {json.dumps({'type': 'done', 'total': total, 'count': len(all_accumulated)})}\n\n"

    return Response(generate_events(), mimetype='text/event-stream')


@app.route('/api/upload-excel', methods=['POST'])
def upload_excel():
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': 'No file uploaded.'}), 400

    file = request.files['file']
    if not file.filename:
        return jsonify({'success': False, 'error': 'Empty filename.'}), 400

    filename = file.filename
    ext = os.path.splitext(filename)[1].lower()

    if ext not in ['.xlsx', '.xls', '.csv']:
        return jsonify({'success': False, 'error': 'Please upload an Excel (.xlsx, .xls) or CSV file.'}), 400

    try:
        temp_path = os.path.join(UPLOAD_FOLDER, filename)
        file.save(temp_path)

        wb = openpyxl.load_workbook(temp_path, read_only=True)
        sheet_names = wb.sheetnames

        chosen_sheet = sheet_names[0]
        for s in ['MIA-2', 'DAILY ENTRY', 'TPT NAGAR']:
            if s in sheet_names:
                chosen_sheet = s
                break

        ws = wb[chosen_sheet]
        
        date_dockets = defaultdict(list)
        all_dockets = []
        curr_date = "Unknown"

        for row in ws.iter_rows(values_only=True):
            row_date = None
            row_dockets = []
            for cell in row:
                if not cell:
                    continue
                c_str = str(cell).strip()
                if re.match(r'^\d{2}\.\d{2}\.\d{2,4}$', c_str):
                    row_date = c_str
                elif re.match(r'^[A-Z]{1,4}[0-9]{6,12}$', c_str, re.I):
                    row_dockets.append(c_str)

            if row_date:
                curr_date = row_date

            for d in row_dockets:
                if d not in all_dockets:
                    all_dockets.append(d)
                if d not in date_dockets[curr_date]:
                    date_dockets[curr_date].append(d)

        wb.close()

        valid_dates = [d for d in date_dockets.keys() if d != "Unknown"]
        latest_date = valid_dates[-1] if valid_dates else None
        
        date_options = []

        # Option 1: Latest Date
        if latest_date:
            date_options.append({
                'label': f"Latest Date: {latest_date} ({len(date_dockets[latest_date])} dockets)",
                'value': latest_date,
                'count': len(date_dockets[latest_date]),
                'numbers': date_dockets[latest_date]
            })

        # Option 2: Last 3 Days
        if len(valid_dates) >= 2:
            last_3 = valid_dates[-3:]
            comb_3 = []
            for d in last_3:
                comb_3.extend(date_dockets[d])
            date_options.append({
                'label': f"Last 3 Days ({len(comb_3)} dockets)",
                'value': 'last_3',
                'count': len(comb_3),
                'numbers': comb_3
            })

        # Option 3: Last 7 Days
        if len(valid_dates) >= 4:
            last_7 = valid_dates[-7:]
            comb_7 = []
            for d in last_7:
                comb_7.extend(date_dockets[d])
            date_options.append({
                'label': f"Last 7 Days ({len(comb_7)} dockets)",
                'value': 'last_7',
                'count': len(comb_7),
                'numbers': comb_7
            })

        # Option 4: Individual recent dates
        for d in reversed(valid_dates[-7:]):
            if d != latest_date:
                date_options.append({
                    'label': f"Date {d} ({len(date_dockets[d])} dockets)",
                    'value': d,
                    'count': len(date_dockets[d]),
                    'numbers': date_dockets[d]
                })

        default_numbers = date_options[0]['numbers'] if date_options else all_dockets[:50]

        return jsonify({
            'success': True,
            'filename': filename,
            'sheets': sheet_names,
            'chosen_sheet': chosen_sheet,
            'total_dockets_in_sheet': len(all_dockets),
            'date_options': date_options,
            'numbers': default_numbers,
            'count': len(default_numbers),
            'latest_date': latest_date
        })

    except Exception as e:
        return jsonify({'success': False, 'error': f'Failed to parse file: {str(e)}'}), 500


@app.route('/api/export-excel', methods=['POST'])
def export_excel():
    data = request.json or {}
    results = data.get('results', [])
    mode = data.get('mode', 'summary')  # 'summary' (fast instant) or 'merged' (merge original file)

    if not results:
        return jsonify({'success': False, 'error': 'No data to export.'}), 400

    status_map = {r['awb'].strip().upper(): r for r in results}

    if mode == 'merged':
        files_dir = UPLOAD_FOLDER if (os.path.exists(UPLOAD_FOLDER) and os.listdir(UPLOAD_FOLDER)) else 'uploads'
        original_files = [f for f in os.listdir(files_dir) if 'daily entry' in f.lower() or 'autosaved' in f.lower()] if os.path.exists(files_dir) else []
        base_file = None
        if original_files:
            base_file = os.path.join(files_dir, original_files[-1])
        elif os.path.exists("DAILY ENTRY 2026.xlsx"):
            base_file = "DAILY ENTRY 2026.xlsx"

        if base_file and os.path.exists(base_file):
            try:
                print(f"Merging live status into {base_file}...", flush=True)
                wb = openpyxl.load_workbook(base_file)
                header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
                header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
                status_font = Font(name="Calibri", size=10, bold=True)
                regular_font = Font(name="Calibri", size=10)

                fill_delivered = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                font_delivered = Font(name="Calibri", size=10, bold=True, color="006100")
                fill_transit = PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid")
                fill_ofd = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
                fill_issue = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

                for sname in ['MIA-2', 'DAILY ENTRY']:
                    if sname not in wb.sheetnames:
                        continue
                    ws = wb[sname]
                    status_col = 11
                    edd_col = 12
                    act_col = 13
                    time_col = 14

                    ws.cell(row=1, column=status_col, value="Live Status").fill = header_fill
                    ws.cell(row=1, column=status_col).font = header_font
                    ws.cell(row=1, column=edd_col, value="EDD").fill = header_fill
                    ws.cell(row=1, column=edd_col).font = header_font
                    ws.cell(row=1, column=act_col, value="Latest Activity").fill = header_fill
                    ws.cell(row=1, column=act_col).font = header_font
                    ws.cell(row=1, column=time_col, value="Activity Date/Time").fill = header_fill
                    ws.cell(row=1, column=time_col).font = header_font

                    for row in ws.iter_rows(min_row=2):
                        row_idx = row[0].row
                        row_docket = None
                        docket_cell = None
                        for cell in row[:10]:
                            if cell.value:
                                val_str = str(cell.value).strip()
                                if re.match(r'^[A-Z]{1,4}[0-9]{6,12}$', val_str, re.I):
                                    row_docket = val_str.upper()
                                    docket_cell = cell
                                    break

                        if row_docket and row_docket in status_map:
                            info = status_map[row_docket]
                            c_status = ws.cell(row=row_idx, column=status_col, value=info.get('status', ''))
                            c_status.font = status_font
                            cat = info.get('category', '')
                            if cat == 'Delivered':
                                c_status.fill = fill_delivered
                                c_status.font = font_delivered
                                # Highlight the docket number cell green when delivered!
                                if docket_cell is not None:
                                    docket_cell.fill = fill_delivered
                                    docket_cell.font = font_delivered
                            elif cat == 'Out for Delivery':
                                c_status.fill = fill_ofd
                            elif cat == 'In Transit':
                                c_status.fill = fill_transit
                            elif cat == 'Issue/RTO':
                                c_status.fill = fill_issue

                            ws.cell(row=row_idx, column=edd_col, value=info.get('edd', '-')).font = regular_font
                            ws.cell(row=row_idx, column=act_col, value=info.get('latest_update', '-')).font = regular_font
                            ws.cell(row=row_idx, column=time_col, value=info.get('latest_time', '-')).font = regular_font

                out_path = os.path.join(UPLOAD_FOLDER, 'temp_merged_export.xlsx')
                wb.save(out_path)
                wb.close()
                print("Merge complete! Sending file...", flush=True)

                return send_file(
                    out_path,
                    as_attachment=True,
                    download_name="DAILY_ENTRY_2026_LIVE_STATUS.xlsx",
                    mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                print("Failed to merge into original:", e, flush=True)

    # Standalone summary workbook fallback (Fast, 0.05s)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DTDC Tracking Status"
    ws.views.sheetView[0].showGridLines = True

    header_fill = PatternFill(start_color="1A365D", end_color="1A365D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    headers = [
        "S.No", "Consignment / AWB No", "Reference No", "Category", 
        "Live Status", "EDD", "Origin", "Destination", "Latest Activity", "Activity Date/Time"
    ]

    ws.append(headers)
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
    ws.row_dimensions[1].height = 26

    cat_colors = {
        'Delivered': PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
        'Out for Delivery': PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid"),
        'In Transit': PatternFill(start_color="E0F2FE", end_color="E0F2FE", fill_type="solid"),
        'Booked': PatternFill(start_color="F1F5F9", end_color="F1F5F9", fill_type="solid"),
        'Issue/RTO': PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid"),
        'Not Found': PatternFill(start_color="F3F4F6", end_color="F3F4F6", fill_type="solid")
    }

    for idx, r in enumerate(results, start=1):
        row_num = idx + 1
        cat = r.get('category', 'In Transit')
        row_data = [
            idx, r.get('awb', ''), r.get('reference_no', '-'), cat,
            r.get('status', ''), r.get('edd', '-'), r.get('origin', '-'),
            r.get('destination', '-'), r.get('latest_update', '-'), r.get('latest_time', '-')
        ]
        ws.append(row_data)
        ws.row_dimensions[row_num].height = 20
        for col_idx in range(1, len(headers) + 1):
            cell = ws.cell(row=row_num, column=col_idx)
            cell.font = data_font
            cell.border = thin_border
            if col_idx in [1, 2, 3, 4, 6]:
                cell.alignment = center_align
            else:
                cell.alignment = left_align
            if col_idx == 4 and cat in cat_colors:
                cell.fill = cat_colors[cat]
                if cat == 'Delivered':
                    cell.font = Font(name="Calibri", size=10, bold=True, color="006100")
            # Highlight docket number in green when delivered
            if col_idx == 2 and cat == 'Delivered':
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                cell.font = Font(name="Calibri", size=10, bold=True, color="006100")

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 14)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return send_file(
        buf,
        as_attachment=True,
        download_name="DTDC_Tracking_Summary.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f"Starting DTDC Bulk Tracker Dashboard on http://0.0.0.0:{port} ...")
    app.run(host='0.0.0.0', port=port, debug=False)

