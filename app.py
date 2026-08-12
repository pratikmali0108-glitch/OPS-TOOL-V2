from flask import Flask, render_template, request, send_file, redirect, url_for, jsonify
import os
import json
from datetime import datetime
from werkzeug.utils import secure_filename
import tempfile
from pathlib import Path
# Original modules
from OMJ import process_omj_file
from SHEFI import process_shefi_file
# Previously added modules
from ambition import process_ambition_file, _reset_client_style_cache
from craft import process_craft_file
from hk import process_hk_file
from fsa import process_fsa_file
from jjl import process_jjl_file
from obu import process_obu_file
from rbl import process_rbl_file
from anaya import process_anaya_file
from uneek import process_uneek_file
from JU import process_ju_excel_file
# New customer modules
from AAM import process_aam_file
from BDLDHI import process_bdldhi_file
from DCT import process_dct_file
from MOR import process_mor_file
from NGL import process_ngl_file
from PC2 import process_pc2_file
from PCB import process_pcb_file
from SGI import process_sgi_file
from vimco_v2 import process_vimco_v2_file
from gati_excel_output import enhance_xlsx_with_template, write_gati_excel
# SKU Extractor modules
import re
import csv
import io
import pandas as pd
from bs4 import BeautifulSoup

# SKU Extractor constants and helpers
SKU_COLUMN_ALIASES = {"sku #", "sku#", "sku"}      # normalised (lower-stripped)

def _normalise(name: str) -> str:
    return name.strip().lower()

def _is_html_xls(filepath: str) -> bool:
    """Returns True when the .xls file is actually an HTML document."""
    try:
        with open(filepath, "rb") as fh:
            header = fh.read(16)
        return header.lstrip().startswith(b"<")
    except OSError:
        return False

def extract_sku_from_csv(filepath: str) -> list[dict]:
    """
    Reads a CSV/TSV file, finds SKU column (SKU #, SKU#, or SKU),
    and returns a list of dicts with keys: file, folder, sku.
    """
    rows = []
    encodings = ["utf-8", "windows-1252", "latin-1"]

    raw = None
    for enc in encodings:
        try:
            with open(filepath, encoding=enc) as fh:
                raw = fh.read()
            break
        except UnicodeDecodeError:
            continue

    if raw is None:
        print(f"  [WARN] Could not decode {filepath}")
        return rows

    # strip Excel "sep=<char>" hint line if present
    lines = raw.splitlines()
    sep = ","
    start = 0
    if lines and lines[0].startswith("sep="):
        # do NOT strip — separator may itself be whitespace (e.g. a tab)
        sep_char = lines[0][4:]
        if sep_char:
            sep = sep_char
        start = 1
        raw = "\n".join(lines[start:])

    try:
        df = pd.read_csv(io.StringIO(raw), sep=sep, dtype=str, low_memory=False)
    except Exception as exc:
        print(f"  [WARN] pandas failed on {filepath}: {exc}")
        return rows

    # find matching SKU column
    sku_col = None
    for col in df.columns:
        if _normalise(col) in SKU_COLUMN_ALIASES:
            sku_col = col
            break

    if sku_col is None:
        print(f"  [WARN] No SKU column found in {filepath}  (cols: {list(df.columns)[:8]})")
        return rows

    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    for val in df[sku_col].dropna():
        val = str(val).strip()
        if val:
            rows.append({"folder": folder, "file": fname, "sku": val})

    return rows

def extract_sku_from_html_xls(filepath: str) -> list[dict]:
    """
    Parses an HTML-disguised XLS file, locates the PODETAIL table
    (the one that contains a 'SKU' column header), and extracts SKU values.
    """
    rows = []
    encodings = ["windows-1252", "utf-8", "latin-1"]

    soup = None
    for enc in encodings:
        try:
            with open(filepath, encoding=enc) as fh:
                content = fh.read()
            soup = BeautifulSoup(content, "html.parser")
            break
        except (UnicodeDecodeError, Exception):
            continue

    if soup is None:
        print(f"  [WARN] Could not parse {filepath}")
        return rows

    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    for table in soup.find_all("table"):
        table_rows = table.find_all("tr")
        if not table_rows:
            continue

        # look for a header row that contains a 'SKU' cell
        header_idx = None
        sku_col_idx = None
        for i, tr in enumerate(table_rows):
            cells = tr.find_all("td")
            texts = [c.get_text(strip=True) for c in cells]
            for j, t in enumerate(texts):
                if _normalise(t) in SKU_COLUMN_ALIASES:
                    header_idx  = i
                    sku_col_idx = j
                    break
            if header_idx is not None:
                break

        if header_idx is None or sku_col_idx is None:
            continue   # this table has no SKU column

        # extract data rows below the header
        for tr in table_rows[header_idx + 1:]:
            cells = tr.find_all("td")
            if sku_col_idx >= len(cells):
                continue
            val = cells[sku_col_idx].get_text(strip=True)
            # skip blank, total/summary rows
            if not val or not re.match(r"^\d+", val):
                continue
            rows.append({"folder": folder, "file": fname, "sku": val})

    if not rows:
        print(f"  [WARN] No SKU data found in {filepath}")

    return rows

def extract_sku_from_binary_xls(filepath: str) -> list[dict]:
    try:
        import xlrd
    except ImportError:
        print("  [WARN] xlrd not installed; cannot read binary XLS.")
        return []

    rows = []
    folder = os.path.basename(os.path.dirname(filepath))
    fname  = os.path.basename(filepath)

    try:
        wb = xlrd.open_workbook(filepath)
    except Exception as exc:
        print(f"  [WARN] xlrd failed on {filepath}: {exc}")
        return rows

    for sheet in wb.sheets():
        # scan every row for a header containing SKU alias
        header_row_idx = None
        sku_col_idx = None
        for r in range(min(sheet.nrows, 30)):
            vals = [str(sheet.cell_value(r, c)).strip() for c in range(sheet.ncols)]
            for c, v in enumerate(vals):
                if _normalise(v) in SKU_COLUMN_ALIASES:
                    header_row_idx = r
                    sku_col_idx    = c
                    break
            if header_row_idx is not None:
                break

        if header_row_idx is None:
            continue

        for r in range(header_row_idx + 1, sheet.nrows):
            val = str(sheet.cell_value(r, sku_col_idx)).strip()
            if val and val not in ("", "0.0", "0"):
                rows.append({"folder": folder, "file": fname, "sku": val})

    return rows

def extract_sku_from_file(filepath: str) -> list[dict]:
    """Extract SKUs from a single file based on its type"""
    ext = os.path.splitext(filepath)[1].lower()

    if ext in (".csv",):
        return extract_sku_from_csv(filepath)

    if ext in (".xls", ".xlsx", ".xlsm"):
        if _is_html_xls(filepath):
            return extract_sku_from_html_xls(filepath)
        else:
            return extract_sku_from_binary_xls(filepath)

    return []
import importlib.util as _ilu
_shefi_new_spec = _ilu.spec_from_file_location(
    'shefi_dhaval',
    os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SHEFI_PO_DHAVAL', 'shefi.py')
)
_shefi_new_mod = _ilu.module_from_spec(_shefi_new_spec)
_shefi_new_spec.loader.exec_module(_shefi_new_mod)
process_shefi_new_file = _shefi_new_mod.process_shefi_new_file
del _shefi_new_spec, _shefi_new_mod, _ilu

app = Flask(__name__)
app.config['SECRET_KEY'] = 'your-secret-key-here'
app.config['UPLOAD_FOLDER'] = tempfile.gettempdir()
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'pdf', 'csv'}

# --- Order Stats Tracking ---
STATS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'order_stats.json')

def load_stats():
    if os.path.exists(STATS_FILE):
        try:
            with open(STATS_FILE, 'r') as f:
                return json.load(f)
        except Exception:
            pass
    return {'customers': {}, 'total_files': 0, 'total_orders': 0}

def save_stats(stats):
    try:
        with open(STATS_FILE, 'w') as f:
            json.dump(stats, f, indent=2)
    except Exception:
        pass

def record_processing(customer_name, files_count, rows_count):
    stats = load_stats()
    cust = stats['customers'].setdefault(customer_name, {
        'files_processed': 0, 'orders_processed': 0, 'last_processed': None
    })
    cust['files_processed'] += files_count
    cust['orders_processed'] += int(rows_count or 0)
    cust['last_processed'] = datetime.now().strftime('%Y-%m-%d %H:%M')
    stats['total_files'] = stats.get('total_files', 0) + files_count
    stats['total_orders'] = stats.get('total_orders', 0) + int(rows_count or 0)
    save_stats(stats)
# --- End Stats Tracking ---

def log_customer_wise_breakdown_daily(stats):
    """
    Write a daily Excel snapshot of the customer-wise breakdown table.

    The snapshot is derived from `stats` (order_stats.json) and is written to:
    OPS_Tool_01042026/SHEFI_PO_DHAVAL/Customer-wise_Breakdown_YYYY-MM-DD.xlsx
    """
    try:
        from datetime import datetime as _dt
        import pandas as pd
    except Exception:
        # If pandas is missing, fail silently (dashboard should still work).
        return

    log_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'SHEFI_PO_DHAVAL')
    os.makedirs(log_dir, exist_ok=True)

    today = _dt.now().strftime('%Y-%m-%d')
    if stats.get('last_breakdown_logged_date') == today:
        return

    customers = stats.get('customers') or {}
    total_orders = int(stats.get('total_orders') or 0)

    rows = []
    # Match the dashboard semantics (orders_processed + files_processed + computed share).
    for name, data in customers.items():
        orders_processed = int((data or {}).get('orders_processed') or 0)
        files_processed = int((data or {}).get('files_processed') or 0)
        last_processed = (data or {}).get('last_processed')
        ip_address = (data or {}).get('ip_address')
        share = round((orders_processed / total_orders) * 100, 1) if total_orders > 0 else 0
        rows.append({
            'Logged Date': today,
            'Customer': name,
            'Orders Processed': orders_processed,
            'Files Processed': files_processed,
            'Share (%)': share,
            'Last Processed': last_processed or 'â\u201d',
            'IP Address': ip_address or 'â\u201d',
        })

    df = pd.DataFrame(rows)
    # Ensure columns exist even when there is no data yet.
    expected_cols = ['Logged Date', 'Customer', 'Orders Processed', 'Files Processed', 'Share (%)', 'Last Processed', 'IP Address']
    if df.empty:
        df = pd.DataFrame(columns=expected_cols)
    else:
        df = df[expected_cols]
        df = df.sort_values('Orders Processed', ascending=False)

    out_path = os.path.join(log_dir, f'Customer-wise_Breakdown_{today}.xlsx')
    df.to_excel(out_path, index=False)

    stats['last_breakdown_logged_date'] = today
    save_stats(stats)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def _apply_gati_template(output_path, csd_remarks=None):
    """Add CSD (OPS form) from sheettoadd.xlsx to an xlsx output file."""
    if output_path and str(output_path).lower().endswith('.xlsx'):
        enhance_xlsx_with_template(output_path, csd_remarks=csd_remarks)

@app.route('/')
def index():
    stats = load_stats()
    log_customer_wise_breakdown_daily(stats)
    return render_template('index.html', stats=stats)

@app.route('/reset-dashboard', methods=['POST'])
def reset_dashboard():
    # Reset dashboard stats but keep "SHEFI New PO" counters intact.
    stats = load_stats()
    customers = stats.get('customers') or {}

    kept_customer_name = 'SHEFI New PO'
    kept = customers.get(kept_customer_name)

    new_customers = {}
    if kept:
        new_customers[kept_customer_name] = kept

    # Recompute totals from kept customer (so share % remains correct).
    total_files = int((kept or {}).get('files_processed') or 0)
    total_orders = int((kept or {}).get('orders_processed') or 0)

    save_stats({
        'customers': new_customers,
        'total_files': total_files,
        'total_orders': total_orders,
        'last_breakdown_logged_date': None,
    })
    return ('', 204)

@app.route('/omj')
def omj_tool():
    return render_template('index_omj.html', cs_count=_get_cs_count('OMJ'))

@app.route('/omj-add-style', methods=['POST'])
def omj_add_style():
    return _handle_add_style('OMJ', 'index_omj.html', 'omj_add_style')

@app.route('/shefi')
def shefi_tool():
    return render_template('index_shefi.html', cs_count=_get_cs_count('SHEFI'))

@app.route('/shefi-add-style', methods=['POST'])
def shefi_add_style():
    return _handle_add_style('SHEFI', 'index_shefi.html', 'shefi_add_style')

def _ambition_cs_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Ambition_CS.xlsx')


def _ambition_cs_count():
    try:
        df = pd.read_excel(_ambition_cs_path())
        return int(df['Client Style No'].dropna().shape[0])
    except Exception:
        return None


@app.context_processor
def inject_ambition_cs_count():
    """Make cs_count available in every template automatically."""
    return {'cs_count': _ambition_cs_count()}


# New tool pages
@app.route('/ambition')
def ambition_tool():
    return render_template('index_ambition.html')


@app.route('/ambition-add-style', methods=['POST'])
def ambition_add_style():
    raw = (request.form.get('style_codes') or '').strip()
    if not raw:
        return render_template('index_ambition.html', cs_error='No style codes entered.')

    # Accept newline-separated or comma-separated (or both)
    entries = [e.strip() for e in re.split(r'[\n,]+', raw) if e.strip()]
    if not entries:
        return render_template('index_ambition.html', cs_error='No valid style codes found.')

    cs_path = _ambition_cs_path()
    try:
        try:
            df_cs = pd.read_excel(cs_path)
            if 'Client Style No' not in df_cs.columns:
                df_cs = pd.DataFrame(columns=['Client Style No'])
        except Exception:
            df_cs = pd.DataFrame(columns=['Client Style No'])

        existing = set(str(v).strip() for v in df_cs['Client Style No'].dropna())
        new_entries = [e for e in entries if e not in existing]
        duplicates = [e for e in entries if e in existing]

        if new_entries:
            new_rows = pd.DataFrame({'Client Style No': new_entries})
            df_cs = pd.concat([df_cs, new_rows], ignore_index=True)
            df_cs.to_excel(cs_path, index=False)
            _reset_client_style_cache()

        msg_parts = []
        if new_entries:
            msg_parts.append(f'{len(new_entries)} code(s) added: {", ".join(new_entries)}')
        if duplicates:
            msg_parts.append(f'{len(duplicates)} already existed: {", ".join(duplicates)}')

        return render_template('index_ambition.html', cs_success=' | '.join(msg_parts))
    except Exception as exc:
        return render_template('index_ambition.html', cs_error=f'Failed to update master: {exc}')

@app.route('/craft')
def craft_tool():
    return render_template('index_craft_hk.html', cs_count=_get_cs_count('Craft'))

@app.route('/craft-add-style', methods=['POST'])
def craft_add_style():
    return _handle_add_style('Craft', 'index_craft_hk.html', 'craft_add_style')

@app.route('/hk')
def hk_tool():
    return render_template('index_hk.html', cs_count=_get_cs_count('HK'))

@app.route('/hk-add-style', methods=['POST'])
def hk_add_style():
    return _handle_add_style('HK', 'index_hk.html', 'hk_add_style')

@app.route('/fsa')
def fsa_tool():
    return render_template('index_fsa.html', cs_count=_get_cs_count('FSA'))

@app.route('/fsa-add-style', methods=['POST'])
def fsa_add_style():
    return _handle_add_style('FSA', 'index_fsa.html', 'fsa_add_style')

@app.route('/jjl')
def jjl_tool():
    return render_template('index_jjl.html', cs_count=_get_cs_count('JJL'))

@app.route('/jjl-add-style', methods=['POST'])
def jjl_add_style():
    return _handle_add_style('JJL', 'index_jjl.html', 'jjl_add_style')

@app.route('/obu')
def obu_tool():
    return render_template('index_obu.html', cs_count=_get_cs_count('OBU'))

@app.route('/obu-add-style', methods=['POST'])
def obu_add_style():
    return _handle_add_style('OBU', 'index_obu.html', 'obu_add_style')

@app.route('/rbl')
def rbl_tool():
    return render_template('index_rbl.html', cs_count=_get_cs_count('RBL'))

@app.route('/rbl-add-style', methods=['POST'])
def rbl_add_style():
    return _handle_add_style('RBL', 'index_rbl.html', 'rbl_add_style')

@app.route('/anaya')
def anaya_tool():
    return render_template('index_anaya.html', cs_count=_get_cs_count('Anaya'))

@app.route('/anaya-add-style', methods=['POST'])
def anaya_add_style():
    return _handle_add_style('Anaya', 'index_anaya.html', 'anaya_add_style')

@app.route('/uneek')
def uneek_tool():
    return render_template('index_uneek.html', cs_count=_get_cs_count('UNEEK'))

@app.route('/uneek-add-style', methods=['POST'])
def uneek_add_style():
    return _handle_add_style('UNEEK', 'index_uneek.html', 'uneek_add_style')

@app.route('/ju')
def ju_tool():
    return render_template('index_ju.html', cs_count=_get_cs_count('JU'))

@app.route('/ju-add-style', methods=['POST'])
def ju_add_style():
    return _handle_add_style('JU', 'index_ju.html', 'ju_add_style')

@app.route('/extract', methods=['POST'])
def extract_skus():
    """Handle file uploads and extract SKUs"""
    try:
        # Get uploaded files
        files = request.files.getlist('files')
        paths = request.form.getlist('paths')
        
        if not files:
            return jsonify({'error': 'No files uploaded'}), 400
        
        all_results = []
        logs = []
        
        for file, path in zip(files, paths):
            if file.filename == '':
                continue
                
            # Save uploaded file temporarily
            filename = secure_filename(file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(filepath)
            
            try:
                # Extract SKUs from the file
                extracted = extract_sku_from_file(filepath)
                all_results.extend(extracted)
                logs.append({
                    'msg': f'Processed {filename}: {len(extracted)} SKUs found',
                    'type': 'ok'
                })
            except Exception as e:
                logs.append({
                    'msg': f'Error processing {filename}: {str(e)}',
                    'type': 'warn'
                })
            finally:
                # Clean up temporary file
                try:
                    os.remove(filepath)
                except:
                    pass
        
        return jsonify({
            'rows': all_results,
            'logs': logs
        })
        
    except Exception as e:
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/sku_extractor')
def sku_extractor_tool():
    return render_template('index_ju.html')

@app.route('/aam')
def aam_tool():
    return render_template('index_aam.html', cs_count=_get_cs_count('AAM'))

@app.route('/aam-add-style', methods=['POST'])
def aam_add_style():
    return _handle_add_style('AAM', 'index_aam.html', 'aam_add_style')

@app.route('/bhakti_dharam')
def bhakti_dharam_tool():
    return redirect(url_for('bdldhi_tool'))

@app.route('/bhakti-add-style', methods=['POST'])
def bhakti_add_style():
    return _handle_add_style('Bhakti', 'index_bhakti_dharam.html', 'bhakti_add_style')

@app.route('/bdldhi')
def bdldhi_tool():
    return render_template('index_bhakti_dharam.html', cs_count=_get_cs_count('Bhakti'))

@app.route('/bdl')
def bdl_tool():
    return render_template('index_bdl.html', cs_count=_get_cs_count('BDL'))

@app.route('/bdl-add-style', methods=['POST'])
def bdl_add_style():
    return _handle_add_style('BDL', 'index_bdl.html', 'bdl_add_style')

@app.route('/dct')
def dct_tool():
    return render_template('index_dct.html', cs_count=_get_cs_count('DCT'))

@app.route('/dct-add-style', methods=['POST'])
def dct_add_style():
    return _handle_add_style('DCT', 'index_dct.html', 'dct_add_style')

@app.route('/mor')
def mor_tool():
    return render_template('index_mor.html', cs_count=_get_cs_count('MOR'))

@app.route('/mor-add-style', methods=['POST'])
def mor_add_style():
    return _handle_add_style('MOR', 'index_mor.html', 'mor_add_style')

@app.route('/ngl')
def ngl_tool():
    return render_template('index_ngl.html', cs_count=_get_cs_count('NGL'))

@app.route('/ngl-add-style', methods=['POST'])
def ngl_add_style():
    return _handle_add_style('NGL', 'index_ngl.html', 'ngl_add_style')

@app.route('/pc2')
def pc2_tool():
    return render_template('index_pc2.html', cs_count=_get_cs_count('PC2'))

@app.route('/pc2-add-style', methods=['POST'])
def pc2_add_style():
    return _handle_add_style('PC2', 'index_pc2.html', 'pc2_add_style')

@app.route('/pcb')
def pcb_tool():
    return render_template('index_pcb.html', cs_count=_get_cs_count('PCB'))

@app.route('/pcb-add-style', methods=['POST'])
def pcb_add_style():
    return _handle_add_style('PCB', 'index_pcb.html', 'pcb_add_style')

@app.route('/sgi')
def sgi_tool():
    return render_template('index_sgi.html', cs_count=_get_cs_count('SGI'))

@app.route('/sgi-add-style', methods=['POST'])
def sgi_add_style():
    return _handle_add_style('SGI', 'index_sgi.html', 'sgi_add_style')

@app.route('/vimco')
def vimco_tool():
    return render_template('index_vimco.html', cs_count=_get_cs_count('VIMCO'))

@app.route('/vimco-add-style', methods=['POST'])
def vimco_add_style():
    return _handle_add_style('VIMCO', 'index_vimco.html', 'vimco_add_style')

@app.route('/process-omj', methods=['POST'])
def process_omj():
    try:
        # Support both single file ('file') and multiple files ('files')
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']  # Filter empty files
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template('index_omj.html', error='No file selected')
        
        if not files:
            return render_template('index_omj.html', error='No file selected')
        
        # Check if separate processing is requested
        process_separately = request.form.get('separate') == 'true'
        
        # Save uploaded files
        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)
        
        if not valid_files:
            return render_template('index_omj.html', error='No valid Excel files found')
        
        try:
            download_urls = []
            all_dataframes = []
            combined_output_filename = None
            
            if process_separately or len(valid_files) == 1:
                # Process each file separately
                for filepath in valid_files:
                    success, output_path, error, df = process_omj_file(filepath, app.config['UPLOAD_FOLDER'])
                    
                    if success:
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                        if df is not None:
                            all_dataframes.append(df)
                    else:
                        return render_template('index_omj.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
            else:
                # Combine all files
                for filepath in valid_files:
                    success, output_path, error, df = process_omj_file(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        all_dataframes.append(df)
                    else:
                        # Clean up on error
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template('index_omj.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                # Combine all dataframes
                import pandas as pd
                combined_df = pd.concat(all_dataframes, ignore_index=True)
                
                # Generate combined output filename
                combined_output_filename = 'OMJ_CASTING_PO_Cleaned_Combined.csv'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], combined_output_filename)
                combined_df.to_csv(output_path, index=False)
                
                download_urls.append({
                    'url': url_for('download_file', filename=combined_output_filename),
                    'filename': combined_output_filename,
                    'rows': len(combined_df)
                })
                all_dataframes = [combined_df]
                
                success_msg = f'Successfully processed and combined {len(valid_files)} file(s)!'
            
            # Clean up input files
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            
            # Check for missing style codes
            missing_style_codes = []
            if "OMJ" in CUSTOMER_CS_CONFIG:
                cs_filename = CUSTOMER_CS_CONFIG["OMJ"]
                cs_list = _load_client_style_list(cs_filename)
                missing_style_codes = _check_missing_style_codes(all_dataframes, cs_list)
            
            # Record processing stats
            record_processing('OMJ', len(valid_files), sum(len(df) for df in all_dataframes))
            
            # Prepare template variables
            template_vars = {
                'success': success_msg,
                'cs_count': _get_cs_count('OMJ'),
            }
            if process_separately or len(valid_files) == 1:
                if len(download_urls) > 1:
                    template_vars['download_urls'] = download_urls
                else:
                    template_vars['download_url'] = download_urls[0]['url']
            else:
                template_vars['download_url'] = url_for('download_file', filename=combined_output_filename)

            if missing_style_codes:
                template_vars['error'] = f'client style master not created for the style(s): {", ".join(missing_style_codes)}'

            return render_template('index_omj.html', **template_vars)
        
        except Exception as proc_error:
            # Clean up input files on processing error
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            raise proc_error
    
    except Exception as e:
        return render_template('index_omj.html', error=f'Error processing file: {str(e)}')

@app.route('/process-shefi', methods=['POST'])
def process_shefi():
    try:
        # Support both single file ('file') and multiple files ('files')
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']  # Filter empty files
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template('index_shefi.html', error='No file selected')
        
        if not files:
            return render_template('index_shefi.html', error='No file selected')
        
        # Check if separate processing is requested
        process_separately = request.form.get('separate') == 'true'
        
        # Save uploaded files
        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)
        
        if not valid_files:
            return render_template('index_shefi.html', error='No valid Excel files found')
        
        try:
            download_urls = []
            all_dataframes = []
            combined_output_filename = None
            
            if process_separately or len(valid_files) == 1:
                # Process each file separately
                for filepath in valid_files:
                    success, output_path, error, df = process_shefi_file(filepath, app.config['UPLOAD_FOLDER'])
                    
                    if success:
                        _apply_gati_template(output_path)
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                        if df is not None:
                            all_dataframes.append(df)
                    else:
                        return render_template('index_shefi.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
            else:
                # Combine all files
                for filepath in valid_files:
                    success, output_path, error, df = process_shefi_file(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        all_dataframes.append(df)
                    else:
                        # Clean up on error
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template('index_shefi.html', 
                                             error=f'Error processing {os.path.basename(filepath)}: {error}')
                
                # Combine all dataframes
                import pandas as pd
                combined_df = pd.concat(all_dataframes, ignore_index=True)
                
                # Generate combined output filename
                combined_output_filename = 'GATI_FORMAT_SHEFI_CLEAN_Combined.xlsx'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], combined_output_filename)
                write_gati_excel(combined_df, output_path)
                
                download_urls.append({
                    'url': url_for('download_file', filename=combined_output_filename),
                    'filename': combined_output_filename,
                    'rows': len(combined_df)
                })
                all_dataframes = [combined_df]
                
                success_msg = f'Successfully processed and combined {len(valid_files)} file(s)!'
            
            # Clean up input files
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            
            # Check for missing style codes
            missing_style_codes = []
            if "SHEFI" in CUSTOMER_CS_CONFIG:
                cs_filename = CUSTOMER_CS_CONFIG["SHEFI"]
                cs_list = _load_client_style_list(cs_filename)
                missing_style_codes = _check_missing_style_codes(all_dataframes, cs_list)
            
            # Record processing stats
            record_processing('SHEFI', len(valid_files), sum(len(df) for df in all_dataframes))
            
            # Prepare template variables
            template_vars = {
                'success': success_msg,
                'cs_count': _get_cs_count('SHEFI'),
            }
            if process_separately or len(valid_files) == 1:
                if len(download_urls) > 1:
                    template_vars['download_urls'] = download_urls
                else:
                    template_vars['download_url'] = download_urls[0]['url']
            else:
                template_vars['download_url'] = url_for('download_file', filename=combined_output_filename)

            if missing_style_codes:
                template_vars['error'] = f'client style master not created for the style(s): {", ".join(missing_style_codes)}'

            return render_template('index_shefi.html', **template_vars)
        
        except Exception as proc_error:
            # Clean up input files on processing error
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            raise proc_error
    
    except Exception as e:
        return render_template('index_shefi.html', error=f'Error processing file: {str(e)}')

@app.route('/download/<filename>')
def download_file(filename):
    return send_file(
        os.path.join(app.config['UPLOAD_FOLDER'], filename),
        as_attachment=True,
        download_name=filename
    )

# Generic processor utility for newly added tools
def _handle_generic_processing(request, template_name, processor_func, output_ext_default, customer_name=None, apply_gati_template=True, get_csd_remarks=None):
    try:
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template(template_name, error='No file selected')

        if not files:
            return render_template(template_name, error='No file selected')

        process_separately = request.form.get('separate') == 'true'

        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)

        if not valid_files:
            return render_template(template_name, error='No valid files found')

        try:
            download_urls = []
            all_dataframes = []
            combined_output_filename = None
            if process_separately or len(valid_files) == 1:
                for filepath in valid_files:
                    success, output_path, error, df = processor_func(filepath, app.config['UPLOAD_FOLDER'])
                    if success:
                        if apply_gati_template and output_ext_default == 'xlsx':
                            # Get remarks if get_csd_remarks is provided
                            csd_remarks = None
                            if get_csd_remarks is not None and df is not None:
                                csd_remarks = get_csd_remarks(df)
                            _apply_gati_template(output_path, csd_remarks=csd_remarks)
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                        if df is not None:
                            all_dataframes.append(df)
                    else:
                        return render_template(template_name, error=f'Error processing {os.path.basename(filepath)}: {error}')
            else:
                import pandas as pd
                dataframes = []
                for filepath in valid_files:
                    success, output_path, error, df = processor_func(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        dataframes.append(df)
                    else:
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template(template_name, error=f'Error processing {os.path.basename(filepath)}: {error}')

                if not dataframes:
                    for fp in valid_files:
                        try:
                            os.remove(fp)
                        except:
                            pass
                    return render_template(template_name, error='No dataframes produced to combine')

                combined_df = pd.concat(dataframes, ignore_index=True)
                combined_output_filename = f'combined_output.{output_ext_default}'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], combined_output_filename)
                if output_ext_default == 'xlsx':
                    if apply_gati_template:
                        csd_remarks = None
                        if get_csd_remarks is not None:
                            csd_remarks = get_csd_remarks(combined_df)
                        write_gati_excel(combined_df, output_path, csd_remarks=csd_remarks)
                    else:
                        combined_df.to_excel(output_path, index=False)
                else:
                    combined_df.to_csv(output_path, index=False)
                
                download_urls.append({
                    'url': url_for('download_file', filename=combined_output_filename),
                    'filename': combined_output_filename,
                    'rows': len(combined_df)
                })
                all_dataframes = [combined_df]

            # Clean up input files
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass
            
            # Check for missing style codes if customer has CS config
            missing_style_codes = []
            if customer_name and customer_name in CUSTOMER_CS_CONFIG:
                cs_filename = CUSTOMER_CS_CONFIG[customer_name]
                cs_list = _load_client_style_list(cs_filename)
                exact_only = customer_name == "VIMCO"
                missing_style_codes = _check_missing_style_codes(all_dataframes, cs_list, exact_only=exact_only)

            # Record processing stats
            if customer_name:
                if process_separately or len(valid_files) == 1:
                    total_rows = sum(d.get('rows', 0) or 0 for d in download_urls)
                else:
                    total_rows = len(all_dataframes[0])
                record_processing(customer_name, len(valid_files), total_rows)

            # Prepare preview data
            import pandas as pd
            preview_data = []
            for i, df in enumerate(all_dataframes):
                if df is not None:
                    # Limit preview to first 20 rows
                    preview_df = df.head(20)
                    preview_data.append({
                        'index': i,
                        'columns': list(preview_df.columns),
                        'rows': preview_df.fillna('').to_dict('records')
                    })

            # Prepare template variables
            success_msg = f'Successfully processed {len(valid_files)} file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
            template_vars = {
                'success': success_msg,
                'preview_data': preview_data
            }

            if customer_name and customer_name in CUSTOMER_CS_CONFIG:
                template_vars['cs_count'] = _get_cs_count(customer_name)

            if process_separately or len(valid_files) == 1:
                if len(download_urls) > 1:
                    template_vars['download_urls'] = download_urls
                else:
                    template_vars['download_url'] = download_urls[0]['url']
            else:
                template_vars['download_url'] = url_for('download_file', filename=combined_output_filename)

            if missing_style_codes:
                template_vars['error'] = f'client style master not created for the style(s): {", ".join(missing_style_codes)}'

            return render_template(template_name, **template_vars)

        except Exception as proc_error:
            for fp in valid_files:
                try:
                    os.remove(fp)
                except:
                    pass
            raise proc_error
    except Exception as e:
        return render_template(template_name, error=f'Error processing file: {str(e)}')


def _is_style_code_valid(built_code: str, cs_list: list, exact_only: bool = False) -> bool:
    """Check if a style code is valid. If exact_only is True, only exact matches are valid."""
    if not built_code or not cs_list:
        return True  # If no code or no list, assume valid

    # 1. Exact match (always checked)
    if built_code in cs_list:
        return True

    if exact_only:
        return False

    # 2. Master entry starts with built_code
    for cs in cs_list:
        if cs.startswith(built_code):
            return True

    # 3. Try inserting 'IN' before the metal/tone suffix
    import re
    _in_pat = re.compile(
        r'^(.+-)(\d+(?:\.\d+)?)((?:WG|YG|RG|AG|PT|W|Y|R|P).*)$',
        re.IGNORECASE,
    )
    m = _in_pat.match(built_code)
    if m:
        with_in = m.group(1) + m.group(2) + 'IN' + m.group(3).upper()
        # 3a. Exact match on IN-variant
        if with_in in cs_list:
            return True
        # 3b. Prefix match on IN-variant
        for cs in cs_list:
            if cs.startswith(with_in):
                return True

    return False

# Customer configuration: maps customer key (used in routes/templates) to CS filename
CUSTOMER_CS_CONFIG = {
    "OMJ": "OMJ_CS.xlsx",
    "SHEFI": "SHEFI_CS.xlsx",
    "Ambition": "Ambition_CS.xlsx",
    "Craft": "CRAFT_CS.xlsx",
    "HK": "HK_CS.xlsx",
    "FSA": "FSA_CS.xlsx",
    "JJL": "JJL_CS.xlsx",
    "OBU": "OBU_CS.xlsx",
    "RBL": "RBL_CS.xlsx",
    "Anaya": "ANAYA_CS.xlsx",
    "UNEEK": "UNEEK_CS.xlsx",
    "JU": "JU_CS.xlsx",
    "AAM": "AAM_CS.xlsx",
    "Bhakti": "DHI_CS.xlsx",
    "Bhakti & Dharam": "DHI_CS.xlsx",
    "Bhakti Diamond LLC": "DHI_CS.xlsx",
    "BDL": "BDL_CS.xlsx",
    "DCT": "DCT_CS.xlsx",
    "MOR": "MOR_CS.xlsx",
    "NGL": "NGL_CS.xlsx",
    "PC2": "PC2_CS.xlsx",
    "PCB": "PCB_CS.xlsx",
    "SGI": "SGI_CS.xlsx",
    "VIMCO": "VIMCO_CS.xlsx",
    "DHI":"DHI_CS.xlsx"
}

def _get_cs_path(customer_key):
    """Get the full path to the CS file for a given customer key."""
    if customer_key not in CUSTOMER_CS_CONFIG:
        return None
    filename = CUSTOMER_CS_CONFIG[customer_key]
    # Check if CS file is in CS_220526 directory (for most customers) or root (for Ambition)
    cs_220526_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'CS_220526', filename)
    if os.path.exists(cs_220526_path):
        return cs_220526_path
    # Fallback to root directory
    root_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)
    return root_path

def _get_cs_count(customer_key):
    """Get the number of valid entries in the CS file for a given customer key."""
    cs_path = _get_cs_path(customer_key)
    if not cs_path:
        return None
    try:
        df = pd.read_excel(cs_path)
        # Check for common column names
        for col in ['Client Style No', 'Style_No', 'Style No', 'STYLE NO']:
            if col in df.columns:
                return int(df[col].dropna().shape[0])
        # If no known column, just count non‑empty rows
        return int(df.dropna(how='all').shape[0])
    except Exception:
        return None

def _handle_add_style(customer_key, template_name, route_name):
    """Generic handler for adding style codes to a customer's CS master file."""
    raw = (request.form.get('style_codes') or '').strip()
    if not raw:
        return render_template(template_name, cs_error='No style codes entered.', 
                             success=None, error=None, download_url=None, download_urls=None)

    # Accept newline-separated or comma-separated (or both)
    entries = [e.strip() for e in re.split(r'[\n,]+', raw) if e.strip()]
    if not entries:
        return render_template(template_name, cs_error='No valid style codes found.',
                             success=None, error=None, download_url=None, download_urls=None)

    cs_path = _get_cs_path(customer_key)
    if not cs_path:
        return render_template(template_name, cs_error='No CS file configured for this customer.',
                             success=None, error=None, download_url=None, download_urls=None)

    try:
        try:
            df_cs = pd.read_excel(cs_path)
            # Check for common column names
            col_name = None
            for col in ['Client Style No', 'Style_No', 'Style No', 'STYLE NO']:
                if col in df_cs.columns:
                    col_name = col
                    break
            if not col_name:
                df_cs = pd.DataFrame(columns=['Client Style No'])
                col_name = 'Client Style No'
        except Exception:
            df_cs = pd.DataFrame(columns=['Client Style No'])
            col_name = 'Client Style No'

        existing = set(str(v).strip() for v in df_cs[col_name].dropna())
        new_entries = [e for e in entries if e not in existing]
        duplicates = [e for e in entries if e in existing]

        if new_entries:
            new_rows = pd.DataFrame({col_name: new_entries})
            df_cs = pd.concat([df_cs, new_rows], ignore_index=True)
            df_cs.to_excel(cs_path, index=False)
            _reset_client_style_cache()

        msg_parts = []
        if new_entries:
            msg_parts.append(f'{len(new_entries)} code(s) added: {", ".join(new_entries)}')
        if duplicates:
            msg_parts.append(f'{len(duplicates)} already existed: {", ".join(duplicates)}')

        return render_template(template_name, cs_success=' | '.join(msg_parts),
                             success=None, error=None, download_url=None, download_urls=None)
    except Exception as exc:
        return render_template(template_name, cs_error=f'Failed to update master: {exc}',
                             success=None, error=None, download_url=None, download_urls=None)

def _load_client_style_list(cs_filename):
    """Load client style list from CS file in CS_220526 directory."""
    try:
        cs_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "CS_220526")
        cs_path = os.path.join(cs_dir, cs_filename)
        if not os.path.exists(cs_path):
            return []
        if cs_filename == "VIMCO_CS.xlsx":
            # VIMCO_CS.xlsx has headers at row 2
            df_cs = pd.read_excel(cs_path, header=2)
            # Clean up column names
            df_cs.columns = ['Select', 'Client Style No', 'ItemSize', 'Style No', 'Client Code', 'Category']
        else:
            df_cs = pd.read_excel(cs_path)
        col = "Client Style No"
        if col in df_cs.columns:
            return [
                str(v).strip() for v in df_cs[col].dropna()
                if str(v).strip() and str(v).strip().upper() != "NAN"
            ]
    except Exception:
        pass
    return []

def _check_missing_style_codes(df_list, cs_list, exact_only: bool = False):
    """Check all StyleCode columns in DataFrames for missing codes."""
    missing = set()
    for df in df_list:
        if df is not None and "StyleCode" in df.columns:
            for code in df["StyleCode"].dropna():
                code_str = str(code).strip()
                if code_str and not _is_style_code_valid(code_str, cs_list, exact_only=exact_only):
                    missing.add(code_str)
    return sorted(list(missing))


@app.route('/process-ambition', methods=['POST'])
def process_ambition():
    size_prefix = (request.form.get('size_prefix') or 'UP').strip().upper()

    def _proc(path, out_dir):
        return process_ambition_file(path, out_dir, size_prefix=size_prefix)

    # Custom handler for Ambition with style code validation
    try:
        if 'files' in request.files:
            files = request.files.getlist('files')
            files = [f for f in files if f.filename != '']
        elif 'file' in request.files:
            file = request.files['file']
            files = [file] if file.filename != '' else []
        else:
            return render_template('index_ambition.html', error='No file selected')

        if not files:
            return render_template('index_ambition.html', error='No file selected')

        process_separately = request.form.get('separate') == 'true'

        valid_files = []
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                valid_files.append(filepath)

        if not valid_files:
            return render_template('index_ambition.html', error='No valid files found')

        try:
            download_urls = []
            all_dataframes = []
            all_output_paths = []
            combined_output_filename = None
            if process_separately or len(valid_files) == 1:
                for filepath in valid_files:
                    success, output_path, error_msg, df = _proc(filepath, app.config['UPLOAD_FOLDER'])
                    if success:
                        _apply_gati_template(output_path)
                        output_filename = os.path.basename(output_path)
                        download_urls.append({
                            'url': url_for('download_file', filename=output_filename),
                            'filename': output_filename,
                            'rows': len(df) if df is not None else 0
                        })
                        if df is not None:
                            all_dataframes.append(df)
                        all_output_paths.append(output_path)
                    else:
                        return render_template('index_ambition.html', error=f'Error processing {os.path.basename(filepath)}: {error_msg}')
            else:
                for filepath in valid_files:
                    success, output_path, error_msg, df = _proc(filepath, app.config['UPLOAD_FOLDER'])
                    if success and df is not None:
                        all_dataframes.append(df)
                    else:
                        for fp in valid_files:
                            try:
                                os.remove(fp)
                            except:
                                pass
                        return render_template('index_ambition.html', error=f'Error processing {os.path.basename(filepath)}: {error_msg}')

                if not all_dataframes:
                    for fp in valid_files:
                        try:
                            os.remove(fp)
                        except:
                            pass
                    return render_template('index_ambition.html', error='No dataframes produced to combine')

                combined_df = pd.concat(all_dataframes, ignore_index=True)
                combined_output_filename = 'combined_output.xlsx'
                output_path = os.path.join(app.config['UPLOAD_FOLDER'], combined_output_filename)
                write_gati_excel(combined_df, output_path)
                download_urls.append({
                    'url': url_for('download_file', filename=combined_output_filename),
                    'filename': combined_output_filename,
                    'rows': len(combined_df)
                })
                all_dataframes = [combined_df]  # Replace with combined df for style checking
                all_output_paths.append(output_path)

            # Clean up input files
            for filepath in valid_files:
                try:
                    os.remove(filepath)
                except:
                    pass

            # Now check style codes against Ambition_CS.xlsx
            missing_style_codes = set()
            try:
                cs_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Ambition_CS.xlsx')
                df_cs = pd.read_excel(cs_path)
                col = 'Client Style No'
                if col in df_cs.columns:
                    cs_list = [
                        str(v).strip() for v in df_cs[col].dropna()
                        if str(v).strip() and str(v).strip().upper() != 'NAN'
                    ]
                else:
                    cs_list = []

                for df in all_dataframes:
                    if 'StyleCode' in df.columns:
                        for code in df['StyleCode'].dropna():
                            code_str = str(code).strip()
                            if code_str and not _is_style_code_valid(code_str, cs_list):
                                missing_style_codes.add(code_str)
            except Exception:
                pass  # If something goes wrong reading the master, just skip the check

            # Record processing stats
            record_processing('Ambition', len(valid_files), sum(len(df) for df in all_dataframes))

            # Prepare template variables
            success_msg = 'Successfully processed file(s)!' if len(valid_files) > 1 else 'File processed successfully!'
            template_vars = {
                'success': success_msg,
            }

            if process_separately or len(valid_files) == 1:
                if len(download_urls) > 1:
                    template_vars['download_urls'] = download_urls
                else:
                    template_vars['download_url'] = download_urls[0]['url']
            else:
                template_vars['download_url'] = url_for('download_file', filename=combined_output_filename)

            if missing_style_codes:
                template_vars['error'] = f'client style master not created for the style(s): {", ".join(sorted(missing_style_codes))}'

            return render_template('index_ambition.html', **template_vars)

        except Exception as proc_error:
            for fp in valid_files:
                try:
                    os.remove(fp)
                except:
                    pass
            raise proc_error
    except Exception as e:
        return render_template('index_ambition.html', error=f'Error processing file: {str(e)}')


@app.route('/process-craft', methods=['POST'])
def process_craft():
    # Read user inputs from form
    size_prefix = (request.form.get('size_prefix') or 'US').strip()
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()

    # Bind arguments via a wrapper so generic handler can call with (path, out_dir)
    def _proc(path, out_dir):
        return process_craft_file(path, out_dir, size_prefix=size_prefix, default_priority=default_priority)

    return _handle_generic_processing(request, 'index_craft_hk.html', _proc, 'xlsx', 'Craft')


@app.route('/process-hk', methods=['POST'])
def process_hk():
    size_prefix = (request.form.get('size_prefix') or 'US').strip()
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()

    def _proc(path, out_dir):
        return process_hk_file(path, out_dir, size_prefix=size_prefix, default_priority=default_priority)

    return _handle_generic_processing(request, 'index_hk.html', _proc, 'xlsx', 'HK')


@app.route('/process-fsa', methods=['POST'])
def process_fsa():
    default_priority = (request.form.get('default_priority') or 'REG').strip().upper()
    stamp_var = (request.form.get('stamp_var') or '').strip().lower()  # '' or 'lgd'

    def _proc(path, out_dir):
        return process_fsa_file(path, out_dir, default_priority=default_priority, default_stamp_var=stamp_var)

    return _handle_generic_processing(request, 'index_fsa.html', _proc, 'xlsx', 'FSA')


@app.route('/process-jjl', methods=['POST'])
def process_jjl():
    default_priority = (request.form.get('default_priority') or 'REG').strip()
    diamond_quality = (request.form.get('diamond_quality') or 'REG').strip()
    def _proc(path, out_dir):
        return process_jjl_file(path, out_dir, default_priority=default_priority, default_diamond_quality=diamond_quality)
    return _handle_generic_processing(request, 'index_jjl.html', _proc, 'xlsx', 'JJL')


@app.route('/process-obu', methods=['POST'])
def process_obu():
    def _get_obu_csd_remarks(df):
        if df is not None and not df.empty and 'CustomerProductionInstruction' in df.columns:
            return str(df.iloc[0]['CustomerProductionInstruction']).strip()
        return None
    return _handle_generic_processing(request, 'index_obu.html', process_obu_file, 'xlsx', 'OBU', get_csd_remarks=_get_obu_csd_remarks)


@app.route('/process-rbl', methods=['POST'])
def process_rbl():
    end_customer_name = (request.form.get('end_customer_name') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    def _proc(path, out_dir):
        return process_rbl_file(path, out_dir, end_customer_name=end_customer_name, priority_value=priority)
    return _handle_generic_processing(request, 'index_rbl.html', _proc, 'xlsx', 'RBL')


@app.route('/process-anaya', methods=['POST'])
def process_anaya():
    tone = (request.form.get('tone') or 'Y').strip().upper()
    def _proc(path, out_dir):
        return process_anaya_file(path, out_dir, tone=tone)
    return _handle_generic_processing(request, 'index_anaya.html', _proc, 'csv', 'Anaya')


@app.route('/process-uneek', methods=['POST'])
def process_uneek():
    po_value = (request.form.get('po_value') or '').strip()
    item_no = (request.form.get('item_no') or '').strip()
    base_serial_start_raw = (request.form.get('base_serial_start') or '').strip()
    style_code = (request.form.get('style_code') or '').strip()
    item_size = (request.form.get('item_size') or '').strip()

    # Safely convert base_serial_start to int if provided
    base_serial_start = None
    if base_serial_start_raw:
        try:
            base_serial_start = int(base_serial_start_raw)
        except ValueError:
            base_serial_start = None

    def _proc(path, out_dir):
        return process_uneek_file(
            path,
            out_dir,
            po_value=po_value,
            item_no=item_no,
            base_serial_start=base_serial_start,
            style_code_input=style_code,
            item_size_input=item_size,
        )

    return _handle_generic_processing(request, 'index_uneek.html', _proc, 'xlsx', 'UNEEK')


@app.route('/process-ju', methods=['POST'])
def process_ju():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority   = (request.form.get('priority') or 'REG').strip().upper()

    def _proc(path, out_dir):
        return process_ju_excel_file(path, out_dir, item_po_no=item_po_no, priority=priority)

    return _handle_generic_processing(request, 'index_ju.html', _proc, 'xlsx', 'JU')


@app.route('/process-aam', methods=['POST'])
def process_aam():
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_aam_file(path, out_dir, priority_value=priority)
    
    return _handle_generic_processing(request, 'index_aam.html', _proc, 'xlsx', 'AAM')


@app.route('/process-bdldhi', methods=['POST'])
def process_bdldhi():
    recycled     = request.form.get('recycled', 'non-recycled').strip().lower() == 'recycled'
    order_group  = (request.form.get('order_group') or '').strip()
    priority     = (request.form.get('priority') or '-5').strip()

    def _proc(path, out_dir):
        return process_bdldhi_file(path, out_dir,
                                   recycled=recycled,
                                   order_group=order_group,
                                   priority=priority)

    return _handle_generic_processing(request, 'index_bhakti_dharam.html', _proc, 'xlsx', 'Bhakti Diamond LLC')


@app.route('/process-dct', methods=['POST'])
def process_dct():
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_dct_file(path, out_dir, priority=priority)
    
    return _handle_generic_processing(request, 'index_dct.html', _proc, 'csv', 'DCT')


@app.route('/process-mor', methods=['POST'])
def process_mor():
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    
    def _proc(path, out_dir):
        return process_mor_file(path, out_dir, item_po_no=item_po_no, priority_value=priority)
    
    return _handle_generic_processing(request, 'index_mor.html', _proc, 'xlsx', 'MOR')


@app.route('/process-ngl', methods=['POST'])
def process_ngl():
    order_qty = (request.form.get('order_qty') or '').strip()
    item_po_no = (request.form.get('item_po_no') or '').strip()
    priority = (request.form.get('priority') or '').strip()
    additional_after_dia = (request.form.get('additional_after_dia') or '').strip()
    
    def _proc(path, out_dir):
        return process_ngl_file(path, out_dir, order_qty=order_qty, item_po_no=item_po_no,
                               priority=priority, additional_after_dia=additional_after_dia)
    
    return _handle_generic_processing(request, 'index_ngl.html', _proc, 'csv', 'NGL')


@app.route('/process-pc2', methods=['POST'])
def process_pc2():
    return _handle_generic_processing(request, 'index_pc2.html', process_pc2_file, 'xlsx', 'PC2')


@app.route('/process-pcb', methods=['POST'])
def process_pcb():
    priority = (request.form.get('priority') or '').strip()
    skuno = (request.form.get('skuno') or '').strip()
    
    def _proc(path, out_dir):
        return process_pcb_file(path, out_dir, priority_value=priority, skuno_value=skuno)
    
    return _handle_generic_processing(request, 'index_pcb.html', _proc, 'csv', 'PCB')


@app.route('/process-sgi', methods=['POST'])
def process_sgi():
    cust_order_no = (request.form.get('cust_order_no') or '').strip()
    
    def _proc(path, out_dir):
        return process_sgi_file(path, out_dir, cust_order_no=cust_order_no)
    
    return _handle_generic_processing(request, 'index_sgi.html', _proc, 'csv', 'SGI')


@app.route('/process-vimco', methods=['POST'])
def process_vimco():
    order_group = (request.form.get('order_group') or '').strip()
    priority = (request.form.get('priority') or '5 day').strip()
    
    def _proc(path, out_dir):
        return process_vimco_v2_file(path, out_dir, order_group=order_group, priority=priority)
    
    return _handle_generic_processing(request, 'index_vimco.html', _proc, 'xlsx', 'VIMCO')


@app.route('/shefi-new')
def shefi_new_tool():
    return render_template('index_shefi_new.html')


def _shefi_cs_path():
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shefi_cs.xlsx')


@app.route('/shefi-new-add-style', methods=['POST'])
def shefi_new_add_style():
    style_nos = request.form.getlist('style_no')
    csms = request.form.getlist('csm')
    n = max(len(style_nos), len(csms))
    pairs: list[tuple[str, str]] = []
    for i in range(n):
        sn = (style_nos[i] if i < len(style_nos) else '').strip()
        csm = (csms[i] if i < len(csms) else '').strip()
        if sn:
            pairs.append((sn, csm))

    if not pairs:
        return render_template('index_shefi_new.html', cs_error='Enter at least one Style_No in the table.')

    cs_path = _shefi_cs_path()
    try:
        try:
            df_cs = pd.read_excel(cs_path)
            if 'Style_No' not in df_cs.columns:
                df_cs = pd.DataFrame(columns=['Style_No', 'CSM'])
        except Exception:
            df_cs = pd.DataFrame(columns=['Style_No', 'CSM'])

        if 'CSM' not in df_cs.columns:
            df_cs['CSM'] = ''

        # Normalise Style_No for lookup (strip)
        key_series = df_cs['Style_No'].astype(str).str.strip()
        existing_map = {k: idx for idx, k in enumerate(key_series) if k and k.lower() != 'nan'}

        added: list[str] = []
        updated: list[str] = []
        skipped_dup_in_form: list[str] = []
        seen_in_form: set[str] = set()

        for sn, csm in pairs:
            if sn in seen_in_form:
                skipped_dup_in_form.append(sn)
                continue
            seen_in_form.add(sn)

            if sn in existing_map:
                idx = existing_map[sn]
                df_cs.at[idx, 'CSM'] = csm
                updated.append(sn)
            else:
                row = {c: '' for c in df_cs.columns}
                row['Style_No'] = sn
                row['CSM'] = csm
                df_cs = pd.concat([df_cs, pd.DataFrame([row])], ignore_index=True)
                existing_map[sn] = len(df_cs) - 1
                added.append(sn)

        df_cs.to_excel(cs_path, index=False)

        msg_parts = []
        if added:
            msg_parts.append(f'{len(added)} row(s) appended: {", ".join(added)}')
        if updated:
            msg_parts.append(f'{len(updated)} existing style(s) CSM updated: {", ".join(updated)}')
        if skipped_dup_in_form:
            msg_parts.append(
                f'Skipped duplicate Style_No in form: {", ".join(skipped_dup_in_form)}'
            )

        return render_template('index_shefi_new.html', cs_success=' | '.join(msg_parts))
    except Exception as exc:
        return render_template('index_shefi_new.html', cs_error=f'Failed to update shefi_cs.xlsx: {exc}')


@app.route('/process-shefi-new', methods=['POST'])
def process_shefi_new():
    return _handle_generic_processing(request, 'index_shefi_new.html',
                                      process_shefi_new_file, 'xlsx', 'SHEFI New PO',
                                      apply_gati_template=False)


if __name__ == '__main__':
    #app.run(debug=True)
    app.run(debug=True, host='0.0.0.0', port=5003)