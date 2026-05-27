from flask import Flask, request, render_template, send_file, jsonify
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.drawing.spreadsheet_drawing import TwoCellAnchor, AnchorMarker
from PIL import Image as PILImage
import io, os, json, shutil
from datetime import datetime

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(BASE_DIR, 'template.xlsx')
SAVE_DIR = os.path.join(BASE_DIR, 'saved_data')
META_FILE = os.path.join(SAVE_DIR, 'meta.json')

os.makedirs(SAVE_DIR, exist_ok=True)

LVSBS = [
    {'id': '1', 'name': 'SS5-1-LVSB1', 'sheet': 'No.1_SS5-1-LVSB1', 'pf': 1.00, 'total': 8,  'closed': 5, 'open': 3},
    {'id': '2', 'name': 'SS5-1-LVSB2', 'sheet': 'No.2_SS5-1-LVSB2', 'pf': 0.99, 'total': 10, 'closed': 6, 'open': 4},
    {'id': '3', 'name': 'SS5-2-LVSB1', 'sheet': 'No.3_SS5-2-LVSB1', 'pf': 1.00, 'total': 8,  'closed': 5, 'open': 3},
    {'id': '4', 'name': 'SS5-2-LVSB2', 'sheet': 'No.4_SS5-2-LVSB2', 'pf': 0.99, 'total': 10, 'closed': 6, 'open': 4},
    {'id': '5', 'name': 'SS5-3-LVSB1', 'sheet': 'No.5_SS5-3-LVSB1', 'pf': 0.95, 'total': 8,  'closed': 6, 'open': 2},
]

# Exact anchor values copied from template — ตรงกรอบแดงทุก slot
SLOT_LAYOUT = [
    dict(fr_row=9,  fr_col=0,  fr_rowOff=168965, fr_colOff=341244, to_row=19, to_col=7,  to_rowOff=86590, to_colOff=278000,  w=981, h=981),
    dict(fr_row=9,  fr_col=7,  fr_rowOff=168965, fr_colOff=402949, to_row=19, to_col=14, to_rowOff=86590, to_colOff=368115,  w=981, h=981),
    dict(fr_row=19, fr_col=0,  fr_rowOff=151646, fr_colOff=341244, to_row=29, to_col=7,  to_rowOff=69272, to_colOff=278000,  w=981, h=981),
    dict(fr_row=19, fr_col=7,  fr_rowOff=151646, fr_colOff=402949, to_row=29, to_col=14, to_rowOff=69272, to_colOff=368115,  w=981, h=981),
    dict(fr_row=29, fr_col=0,  fr_rowOff=151646, fr_colOff=341244, to_row=39, to_col=7,  to_rowOff=69272, to_colOff=278000,  w=981, h=981),
    dict(fr_row=29, fr_col=7,  fr_rowOff=151646, fr_colOff=402949, to_row=39, to_col=14, to_rowOff=69272, to_colOff=368115,  w=981, h=981),
    dict(fr_row=41, fr_col=0,  fr_rowOff=173039, fr_colOff=341243, to_row=51, to_col=7,  to_rowOff=90664, to_colOff=277999,  w=981, h=981),
    dict(fr_row=41, fr_col=7,  fr_rowOff=173039, fr_colOff=402948, to_row=51, to_col=14, to_rowOff=90664, to_colOff=368114,  w=981, h=981),
    dict(fr_row=51, fr_col=0,  fr_rowOff=155720, fr_colOff=341243, to_row=61, to_col=7,  to_rowOff=73346, to_colOff=277999,  w=981, h=981),
    dict(fr_row=51, fr_col=7,  fr_rowOff=155720, fr_colOff=402948, to_row=61, to_col=14, to_rowOff=73346, to_colOff=368114,  w=981, h=981),
    dict(fr_row=61, fr_col=0,  fr_rowOff=155720, fr_colOff=341243, to_row=71, to_col=7,  to_rowOff=73346, to_colOff=277999,  w=981, h=981),
    dict(fr_row=61, fr_col=7,  fr_rowOff=155720, fr_colOff=402948, to_row=71, to_col=14, to_rowOff=73346, to_colOff=368114,  w=981, h=981),
]

def load_meta():
    if os.path.exists(META_FILE):
        with open(META_FILE, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {'month_th': 'เมษายน', 'year_th': '2568', 'lvsbs': {}}

def save_meta(meta):
    with open(META_FILE, 'w', encoding='utf-8') as f:
        json.dump(meta, f, ensure_ascii=False, indent=2)

def img_dir(lid):
    d = os.path.join(SAVE_DIR, f'lvsb_{lid}')
    os.makedirs(d, exist_ok=True)
    return d

def resize_image_square(img_bytes, size=981):
    """Resize & crop image to exact square matching template slot size"""
    img = PILImage.open(io.BytesIO(img_bytes))
    if img.mode in ('RGBA', 'P'):
        img = img.convert('RGB')
    # Crop to square from center
    w, h = img.size
    min_dim = min(w, h)
    left = (w - min_dim) // 2
    top = (h - min_dim) // 2
    img = img.crop((left, top, left + min_dim, top + min_dim))
    img = img.resize((size, size), PILImage.LANCZOS)
    out = io.BytesIO()
    img.save(out, format='JPEG', quality=90)
    out.seek(0)
    return out.read()

def insert_image_exact(ws, img_path, slot):
    """Insert image using exact anchor from template"""
    xl_img = XLImage(img_path)
    xl_img.width  = slot['w']
    xl_img.height = slot['h']

    anchor = TwoCellAnchor()
    anchor.editAs = 'twoCell'
    anchor._from = AnchorMarker(
        col=slot['fr_col'], colOff=slot['fr_colOff'],
        row=slot['fr_row'], rowOff=slot['fr_rowOff']
    )
    anchor.to = AnchorMarker(
        col=slot['to_col'], colOff=slot['to_colOff'],
        row=slot['to_row'], rowOff=slot['to_rowOff']
    )
    xl_img.anchor = anchor
    ws.add_image(xl_img)

@app.route('/')
def index():
    meta = load_meta()
    return render_template('index.html', lvsbs=LVSBS, meta=meta)

@app.route('/save_lvsb', methods=['POST'])
def save_lvsb():
    lid = request.form.get('lid')
    if not lid:
        return jsonify({'ok': False, 'error': 'missing lid'})

    meta = load_meta()
    if 'lvsbs' not in meta:
        meta['lvsbs'] = {}

    meta['month_th'] = request.form.get('month_th', meta.get('month_th', ''))
    meta['year_th']  = request.form.get('year_th',  meta.get('year_th', ''))

    entry = meta['lvsbs'].get(lid, {})
    entry['date']     = request.form.get('date', '')
    entry['kwh_prev'] = request.form.get('kwh_prev', '')
    entry['kwh_curr'] = request.form.get('kwh_curr', '')
    entry['saved']    = True

    d = img_dir(lid)
    saved_imgs = list(entry.get('images', []))

    files = request.files.getlist('photos')
    new_imgs = [f for f in files if f and f.filename]

    for f in new_imgs:
        idx = len(saved_imgs)
        if idx >= 12:
            break
        fname = f'{idx:02d}_{f.filename.replace("/","_")}.jpg'
        fpath = os.path.join(d, fname)
        img_bytes = resize_image_square(f.read())
        with open(fpath, 'wb') as out:
            out.write(img_bytes)
        saved_imgs.append(fname)

    entry['images'] = saved_imgs
    meta['lvsbs'][lid] = entry
    save_meta(meta)

    return jsonify({'ok': True, 'img_count': len(saved_imgs), 'entry': entry})

@app.route('/delete_image', methods=['POST'])
def delete_image():
    data = request.get_json()
    lid   = data.get('lid')
    fname = data.get('fname')
    meta  = load_meta()
    entry = meta.get('lvsbs', {}).get(lid, {})
    imgs  = entry.get('images', [])

    fpath = os.path.join(img_dir(lid), fname)
    if os.path.exists(fpath):
        os.remove(fpath)
    if fname in imgs:
        imgs.remove(fname)
    entry['images'] = imgs
    meta['lvsbs'][lid] = entry
    save_meta(meta)
    return jsonify({'ok': True, 'img_count': len(imgs)})

@app.route('/get_state')
def get_state():
    meta = load_meta()
    result = {}
    for lid, entry in meta.get('lvsbs', {}).items():
        imgs = [f for f in entry.get('images', [])
                if os.path.exists(os.path.join(img_dir(lid), f))]
        entry['images'] = imgs
        result[lid] = entry
    meta['lvsbs'] = result
    return jsonify(meta)

@app.route('/image/<lid>/<fname>')
def serve_image(lid, fname):
    fpath = os.path.join(img_dir(lid), fname)
    if not os.path.exists(fpath):
        return 'not found', 404
    return send_file(fpath, mimetype='image/jpeg')

@app.route('/reset', methods=['POST'])
def reset():
    if os.path.exists(SAVE_DIR):
        shutil.rmtree(SAVE_DIR)
    os.makedirs(SAVE_DIR, exist_ok=True)
    return jsonify({'ok': True})

@app.route('/export')
def export():
    meta     = load_meta()
    month_th = meta.get('month_th', '')
    year_th  = meta.get('year_th', '')

    wb = openpyxl.load_workbook(TEMPLATE_PATH)

    ws_main = wb['Main']
    ws_main['A1'] = f'แบบแผนงานสำรวจ  LVSB  ประจำเดือน  {month_th}   {year_th}'

    for lvsb in LVSBS:
        lid   = lvsb['id']
        entry = meta.get('lvsbs', {}).get(lid, {})
        ws    = wb[lvsb['sheet']]

        # วันที่
        date_val = entry.get('date', '')
        if date_val:
            try:
                d = datetime.strptime(date_val, '%Y-%m-%d')
                ws['C7'] = d.strftime('%d/%m/%Y')
            except:
                ws['C7'] = date_val

        # kWh ใน Main sheet
        row_idx = int(lid) + 3
        for col, key in [(7, 'kwh_prev'), (8, 'kwh_curr')]:
            val = entry.get(key, '')
            if val:
                try:    ws_main.cell(row=row_idx, column=col).value = float(val)
                except: ws_main.cell(row=row_idx, column=col).value = val

        # รูปภาพ — เก็บ logo (image 0) แทนที่รูป placeholder ด้วยรูปจริง
        logo = ws._images[0] if ws._images else None
        ws._images = []
        if logo:
            ws._images.append(logo)

        imgs = entry.get('images', [])
        for i, fname in enumerate(imgs[:12]):
            fpath = os.path.join(img_dir(lid), fname)
            if not os.path.exists(fpath):
                continue
            try:
                insert_image_exact(ws, fpath, SLOT_LAYOUT[i])
            except Exception as e:
                print(f"Image error slot {i} {fname}: {e}")

    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    fname = f'Energy_LVSB_{month_th}_{year_th}.xlsx'
    return send_file(out, as_attachment=True, download_name=fname,
                     mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

if __name__ == '__main__':
    print("\n" + "="*55)
    print("  ⚡ LVSB Energy Report App")
    print("  เปิดเบราว์เซอร์แล้วไปที่: http://localhost:5050")
    print("  ข้อมูลบันทึกอัตโนมัติใน saved_data/")
    print("="*55 + "\n")
    app.run(debug=False, port=5050)
