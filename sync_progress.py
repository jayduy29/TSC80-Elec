"""
Script đồng bộ dữ liệu hàng tuần tự động từ TẤT CẢ các file tiến độ tuần
(TCS80-C459-ELECTRICAL PROGRESS-W*.xlsx / ELECTRICCAL PROGRESS-W*.xlsx
và file master TCS80-C459-ELECTRICAL PROGRESS.xlsx) trong folder
vào C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx, prog_data.json, và index.html (D.cbl).

Sử dụng:
  python sync_progress.py                  # Tự động quét tất cả các file tuần trong folder + file master
  python sync_progress.py --week W48       # Chỉ định tuần cho file master
  python sync_progress.py --no-push        # Không tự động git push
"""

import os
import sys
import json
import re
import glob
import zipfile
import shutil
import subprocess
from datetime import date, datetime, timedelta
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font
from openpyxl.formatting.rule import FormulaRule

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
MASTER_FILE = os.path.join(BASE_DIR, 'TCS80-C459-ELECTRICAL PROGRESS.xlsx')
OVERVIEW_FILE = os.path.join(BASE_DIR, 'C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx')
JSON_OUT = os.path.join(BASE_DIR, 'prog_data.json')
HTML_FILE = os.path.join(BASE_DIR, 'index.html')

CONTR_PLANS = {
    'HDC': {'cut': 30259, 'pull': 30259, 'conn': 2694},
    'NAUTICS': {'cut': 26381, 'pull': 26381, 'conn': 2006},
    'ST': {'cut': 8593, 'pull': 8593, 'conn': 554},
    'TOTAL': {'cut': 65233, 'pull': 65233, 'conn': 5254}
}

ROW_MAP = {
    'HDC': {'cut': 7, 'pull': 9, 'conn': 11},
    'NAUTICS': {'cut': 13, 'pull': 15, 'conn': 17},
    'ST': {'cut': 19, 'pull': 21, 'conn': 23},
    'TOTAL': {'cut': 25, 'pull': 27, 'conn': 29}
}

ROW_PLANS = {
    7: 30259, 9: 30259, 11: 2694,
    13: 26381, 15: 26381, 17: 2006,
    19: 8593, 21: 8593, 23: 554,
    25: 65233, 27: 65233, 29: 5254
}

QTY_ROWS = [7, 9, 11, 13, 15, 17, 19, 21, 23]
TOTAL_QTY_ROWS = {25: (7, 13, 19), 27: (9, 15, 21), 29: (11, 17, 23)}
PCT_ROWS = [8, 10, 12, 14, 16, 18, 20, 22, 24, 26, 28, 30]

# Màu nền từng khối nhà thầu để ẩn chữ #N/A ở các tuần tương lai chưa có số liệu
BAND_BG_COLORS = [
    ((8, 10, 12), 'FFEBF1F5'),    # HDC
    ((14, 16, 18), 'FFF4F1EA'),   # NAUTICS
    ((20, 22, 24), 'FFEAF2EC'),   # ST
    ((26, 28, 30), 'FFD9E2EC'),   # TOTAL
]

EXCEL_ERRORS = {'#VALUE!', '#REF!', '#N/A', '#NAME?', '#DIV/0!', '#NUM!', '#NULL!'}
WEEKLY_FILE_RE = re.compile(r'TCS80-C459-ELECTRI[A-Z]*\s+PROGRESS[-_\s]*W(\d{1,2})\.xlsx$', re.IGNORECASE)


def parse_cell_num(val):
    """Chuyển giá trị ô sang số. Trả về None nếu ô bị lỗi Excel (#VALUE!, #REF!...), 0 nếu trống/'-'."""
    if val is None:
        return 0
    s = str(val).strip()
    if s == '' or s == '-':
        return 0
    if s.upper() in EXCEL_ERRORS:
        return None
    try:
        return float(s.replace(',', ''))
    except Exception:
        return None


def to_int(val, default=0):
    if val is None:
        return default
    try:
        return int(float(str(val).strip()))
    except Exception:
        return default


def get_current_report_week():
    today = date.today()
    # Dữ liệu tuần mới được tính từ Thứ 5 hàng tuần (Thứ 5 đến Thứ 4 tuần sau cùng 1 chu kỳ)
    offset = (today.weekday() - 3) % 7
    report_thu = today - timedelta(days=offset)
    return f"W{report_thu.isocalendar()[1]:02d}"


def read_tong_sheet(excel_path):
    """Đọc sheet TỔNG từ một file tiến độ tuần. Các ô lỗi trả về None để không ghi đè mất dữ liệu cũ."""
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        tong_name = None
        for s in wb.sheetnames:
            if s.strip().upper() in ('TỔNG', 'TONG'):
                tong_name = s
                break
        if not tong_name:
            raise ValueError(f"Sheet 'TỔNG' không tồn tại trong {excel_path}")
        ws = wb[tong_name]

        data = {
            'HDC': {'cut': 0, 'pull': 0, 'conn': 0},
            'NAUTICS': {'cut': 0, 'pull': 0, 'conn': 0},
            'ST': {'cut': 0, 'pull': 0, 'conn': 0},
            'TOTAL': {'cut': 0, 'pull': 0, 'conn': 0}
        }

        for row in ws.iter_rows(min_row=1, max_row=20, max_col=10, values_only=True):
            c1 = str(row[0] or '').strip().lower()
            if not c1 or 'total' in c1 or 'tổng' in c1:
                continue
            if 'cut' in c1 or 'cắt' in c1:
                key = 'cut'
            elif 'pull' in c1 or 'kéo' in c1:
                key = 'pull'
            elif 'connect' in c1 or 'đấu' in c1:
                key = 'conn'
            else:
                continue

            data['HDC'][key] = parse_cell_num(row[1] if len(row) > 1 else None)
            data['NAUTICS'][key] = parse_cell_num(row[3] if len(row) > 3 else None)
            data['ST'][key] = parse_cell_num(row[5] if len(row) > 5 else None)
            data['TOTAL'][key] = parse_cell_num(row[7] if len(row) > 7 else None)

        return data
    finally:
        wb.close()


def discover_weekly_files(folder_dir, master_file=None, master_week=None):
    """
    Quét toàn bộ thư mục để tìm các file báo cáo tuần (TCS80-C459-ELECTRI* PROGRESS-W##.xlsx)
    và kết hợp với file master hiện tại (TCS80-C459-ELECTRICAL PROGRESS.xlsx).
    Trả về dict: { 'W37': (filepath, mtime), ..., 'W48': (filepath, mtime) }
    """
    week_files = {}
    for fname in os.listdir(folder_dir):
        if fname.startswith('~$') or not fname.lower().endswith('.xlsx'):
            continue
        m = WEEKLY_FILE_RE.match(fname)
        if not m:
            continue
        wcode = f"W{int(m.group(1)):02d}"
        fpath = os.path.join(folder_dir, fname)
        mtime = os.path.getmtime(fpath)
        # Nếu có 2 file trùng tuần (vd lỗi chính tả ELECTRICCAL vs ELECTRICAL), ưu tiên file mới hơn
        if wcode not in week_files or mtime > week_files[wcode][1]:
            week_files[wcode] = (fpath, mtime)

    # Kết hợp file master cho tuần hiện tại nếu file master mới hơn hoặc tuần hiện tại chưa có file riêng
    if master_file and os.path.exists(master_file):
        mw = master_week or get_current_report_week()
        m_mtime = os.path.getmtime(master_file)
        if mw not in week_files or m_mtime >= week_files[mw][1]:
            week_files[mw] = (master_file, m_mtime)

    return week_files


def ensure_overview_formulas_and_styles(wb_ov):
    """
    Khôi phục và chuẩn hóa toàn bộ công thức & định dạng trong cả 2 sheet của
    C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx:
    - Sheet 'Electrical Progress':
        * Các dòng Tổng dự án (25, 27, 29): công thức SUM động
        * Các dòng % (8, 10, ..., 30): công thức IF(..., NA(), Qty/$E$row) để biểu đồ không tụt về 0
        * Định dạng số #,##0 và 0.0% cho toàn bộ 24 tuần (F..AC)
        * Conditional Formatting ẩn chữ #N/A cho các tuần tương lai chưa có dữ liệu
    - Sheet 'Tiến độ phần điện':
        * Toàn bộ F7:AC30 link động qua công thức từ 'Electrical Progress'
    """
    if 'Electrical Progress' not in wb_ov.sheetnames:
        return

    ws_en = wb_ov['Electrical Progress']
    week_cols = []
    for c in range(1, ws_en.max_column + 1):
        val = str(ws_en.cell(5, c).value or '').strip().upper()
        if re.match(r'^W\d{2}$', val):
            week_cols.append((c, val))

    if not week_cols:
        return

    first_col_l = get_column_letter(week_cols[0][0])
    last_col_l = get_column_letter(week_cols[-1][0])

    # 1. Chuẩn hóa công thức và định dạng trên sheet 'Electrical Progress'
    for c, _ in week_cols:
        col_l = get_column_letter(c)
        for r in QTY_ROWS:
            ws_en.cell(r, c).number_format = '#,##0'
        for tr, (r1, r2, r3) in TOTAL_QTY_ROWS.items():
            cell = ws_en.cell(tr, c)
            cell.value = f'=IF(COUNT({col_l}{r1},{col_l}{r2},{col_l}{r3})>0,SUM({col_l}{r1},{col_l}{r2},{col_l}{r3}),"")'
            cell.number_format = '#,##0'
        for pr in PCT_ROWS:
            cell = ws_en.cell(pr, c)
            cell.value = f'=IF(OR({col_l}{pr-1}="",ISBLANK({col_l}{pr-1})),NA(),{col_l}{pr-1}/$E${pr-1})'
            cell.number_format = '0.0%'

    # 2. Chuẩn hóa công thức trên sheet 'Tiến độ phần điện'
    if 'Tiến độ phần điện' in wb_ov.sheetnames:
        ws_vi = wb_ov['Tiến độ phần điện']
        for c, _ in week_cols:
            col_l = get_column_letter(c)
            for r in QTY_ROWS + list(TOTAL_QTY_ROWS.keys()):
                cell = ws_vi.cell(r, c)
                cell.value = f"=IF('Electrical Progress'!{col_l}{r}=\"\",\"\",'Electrical Progress'!{col_l}{r})"
                cell.number_format = '#,##0'
            for pr in PCT_ROWS:
                cell = ws_vi.cell(pr, c)
                cell.value = f"=IF(ISNA('Electrical Progress'!{col_l}{pr}),NA(),IF('Electrical Progress'!{col_l}{pr}=\"\",\"\",'Electrical Progress'!{col_l}{pr}))"
                cell.number_format = '0.0%'

    # 3. Thêm Conditional Formatting ẩn #N/A cho các tuần trống ở cả 2 sheet
    for sname in ('Electrical Progress', 'Tiến độ phần điện'):
        if sname not in wb_ov.sheetnames:
            continue
        ws = wb_ov[sname]
        # Xóa CF cũ nếu có để tránh trùng lặp
        ws.conditional_formatting._cf_rules.clear()
        for rows_group, bg_hex in BAND_BG_COLORS:
            hide_font = Font(color=bg_hex)
            for r in rows_group:
                rng = f"{first_col_l}{r}:{last_col_l}{r}"
                rule = FormulaRule(formula=[f"ISNA({first_col_l}{r})"], font=hide_font)
                ws.conditional_formatting.add(rng, rule)


def atomic_save_workbook(wb, target_path):
    """Lưu file Excel qua file tạm + kiểm tra zip hợp lệ rồi mới thay thế để tránh hỏng file (BadZipFile)."""
    tmp_path = target_path + '.tmp'
    try:
        wb.save(tmp_path)
        # Kiểm tra file tạm là file zip hợp lệ
        with zipfile.ZipFile(tmp_path, 'r') as zf:
            if zf.testzip() is not None:
                raise IOError("File tạm sau khi lưu bị lỗi cấu trúc zip!")
        os.replace(tmp_path, target_path)
        return True
    except PermissionError:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        raise
    except Exception:
        if os.path.exists(tmp_path):
            try:
                os.remove(tmp_path)
            except Exception:
                pass
        raise


def recalculate_with_excel_com(xlsx_path):
    """Mở ngầm bằng Excel COM (nếu có Excel trên Windows) để tính toán lại và lưu cache công thức."""
    ps_script = f"""
$ErrorActionPreference = 'Stop'
$excel = $null
$wb = $null
try {{
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.AskToUpdateLinks = $false
    $wb = $excel.Workbooks.Open('{xlsx_path}')
    $excel.CalculateFullRebuild()
    $wb.Save()
    $wb.Close($false)
    $wb = $null
    $excel.Quit()
    $excel = $null
    Write-Output 'OK'
}} catch {{
    if ($wb -ne $null) {{ try {{ $wb.Close($false) }} catch {{}} }}
    if ($excel -ne $null) {{ try {{ $excel.Quit() }} catch {{}} }}
    Write-Error $_
    exit 1
}}
"""
    try:
        res = subprocess.run(
            ['powershell', '-NoProfile', '-Command', ps_script],
            capture_output=True, text=True, timeout=45
        )
        if res.returncode == 0 and 'OK' in res.stdout:
            return True
    except Exception:
        pass
    return False


def update_overview_from_folder(week_data_map):
    """
    Cập nhật toàn bộ các tuần tìm thấy trong folder vào C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx,
    giữ nguyên công thức động, định dạng và biểu đồ.
    Trả về (wb_ov, resolved_weeks_values) để xuất prog_data.json chính xác.
    """
    if not os.path.exists(OVERVIEW_FILE):
        raise FileNotFoundError(f"{OVERVIEW_FILE} không tìm thấy!")

    # Kiểm tra file có bị lỗi BadZipFile không; nếu hỏng thì thử khôi phục từ git HEAD
    try:
        wb_ov = openpyxl.load_workbook(OVERVIEW_FILE)
    except zipfile.BadZipFile:
        print("⚠️ Phát hiện file Overview bị lỗi cấu trúc zip, đang khôi phục bản sạch từ Git...")
        subprocess.run(['git', '-C', BASE_DIR, 'checkout', 'HEAD', '--', 'C459_ELECTRICAL_PROGRESS_OVERVIEW.xlsx'], check=True)
        wb_ov = openpyxl.load_workbook(OVERVIEW_FILE)

    ws_en = wb_ov['Electrical Progress']
    week_col_map = {}
    for c in range(1, ws_en.max_column + 1):
        val = str(ws_en.cell(5, c).value or '').strip().upper()
        if re.match(r'^W\d{2}$', val):
            week_col_map[val] = c

    # Cập nhật 9 dòng số lượng của 3 nhà thầu trên sheet 'Electrical Progress'
    updated_weeks = []
    for wcode, data in sorted(week_data_map.items()):
        if wcode not in week_col_map:
            print(f"⚠️ Bỏ qua {wcode}: không có cột tương ứng trong Overview")
            continue
        col_idx = week_col_map[wcode]
        for contr in ('HDC', 'NAUTICS', 'ST'):
            for item_k in ('cut', 'pull', 'conn'):
                r = ROW_MAP[contr][item_k]
                new_val = data[contr][item_k]
                if new_val is None:
                    # Giữ nguyên số cũ hợp lệ trong Overview nếu file nguồn bị lỗi #VALUE!
                    existing = ws_en.cell(r, col_idx).value
                    if isinstance(existing, (int, float)):
                        continue
                    ws_en.cell(r, col_idx).value = 0
                else:
                    ws_en.cell(r, col_idx).value = int(round(new_val))
        updated_weeks.append(wcode)

    # Khôi phục toàn bộ công thức SUM, %, link sang sheet tiếng Việt và Conditional Formatting
    ensure_overview_formulas_and_styles(wb_ov)

    # Thu thập bảng giá trị thực tế cho mọi tuần (dùng cho prog_data.json)
    col_values = {}
    for wcode, c in week_col_map.items():
        vals = {}
        has_any = False
        for r in QTY_ROWS:
            v = ws_en.cell(r, c).value
            if isinstance(v, (int, float)):
                vals[r] = int(round(v))
                has_any = True
            else:
                vals[r] = None
        if has_any:
            for tr, (r1, r2, r3) in TOTAL_QTY_ROWS.items():
                vals[tr] = (vals[r1] or 0) + (vals[r2] or 0) + (vals[r3] or 0)
            for pr in PCT_ROWS:
                pv = vals.get(pr - 1)
                plan_v = ROW_PLANS.get(pr - 1, 0)
                vals[pr] = round(pv / plan_v, 6) if (pv is not None and plan_v > 0) else None
        col_values[c] = vals

    saved = False
    try:
        atomic_save_workbook(wb_ov, OVERVIEW_FILE)
        saved = True
        print(f"✅ Đã cập nhật an toàn {os.path.basename(OVERVIEW_FILE)} ({len(updated_weeks)} tuần: {', '.join(updated_weeks)})")
    except PermissionError:
        print(f"⚠️ Lưu ý: File {os.path.basename(OVERVIEW_FILE)} đang mở trong Excel, bỏ qua lưu đè Excel nhưng vẫn đồng bộ JSON & Web bình thường!")
    except Exception as e:
        print(f"⚠️ Lỗi lưu {os.path.basename(OVERVIEW_FILE)}: {e}")

    if saved:
        if recalculate_with_excel_com(OVERVIEW_FILE):
            print("✅ Đã tính toán lại toàn bộ công thức và cập nhật cache bằng Excel.")
        else:
            print("ℹ️ Bỏ qua bước recalculate bằng Excel COM (công thức sẽ tự tính khi mở file trong Excel).")

    return wb_ov, col_values


def generate_prog_json(wb_ov, col_values):
    """Xuất prog_data.json từ cấu trúc Overview + giá trị đã tính sẵn (không bị dính chuỗi công thức)."""
    sheets_json = []
    now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")

    for sname in wb_ov.sheetnames:
        ws = wb_ov[sname]
        rows = []
        for r in range(1, ws.max_row + 1):
            row_vals = []
            for c in range(1, ws.max_column + 1):
                if c in col_values and r in (QTY_ROWS + list(TOTAL_QTY_ROWS.keys()) + PCT_ROWS):
                    val = col_values[c].get(r)
                elif c == 5 and r in PCT_ROWS:
                    val = 1
                else:
                    raw = ws.cell(r, c).value
                    if isinstance(raw, str) and raw.startswith('='):
                        # Link cột E (Total Plan) ở sheet Tiếng Việt
                        if c == 5 and r in ROW_PLANS:
                            val = ROW_PLANS[r]
                        else:
                            val = ''
                    else:
                        val = raw

                if val is None or val == '':
                    row_vals.append('')
                elif isinstance(val, float):
                    if val.is_integer():
                        row_vals.append(str(int(val)))
                    else:
                        row_vals.append(f"{val:.6g}")
                else:
                    row_vals.append(str(val))
            if any(x != '' for x in row_vals):
                rows.append(row_vals)
        sheets_json.append({
            "name": sname,
            "rows": rows,
            "synced_at": now_str
        })

    tmp_json = JSON_OUT + '.tmp'
    with open(tmp_json, 'w', encoding='utf-8') as f:
        json.dump(sheets_json, f, ensure_ascii=False, indent=2)
    os.replace(tmp_json, JSON_OUT)

    print(f"✅ Đã tạo {os.path.basename(JSON_OUT)} ({len(sheets_json)} sheets)")


def sync_cables_to_html(excel_path):
    if not os.path.exists(HTML_FILE):
        return 0
    wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
    try:
        if 'Consolidated' not in wb.sheetnames:
            print("⚠️ Sheet 'Consolidated' không có, bỏ qua đồng bộ cáp vào index.html")
            return 0

        ws = wb['Consolidated']
        rows_iter = ws.iter_rows(values_only=True)
        header = [str(v or '').strip() for v in next(rows_iter)[:25]]

        col_contr = header.index('Contractor')
        col_cid = header.index('Cable ID')
        col_cut = header.index('Cut')
        col_pulled = header.index('Pulled')
        col_conn = header.index('Connection')
        col_megger = header.index('CWC- Megger')
        col_fa = header.index('From Device Area')
        col_ta = header.index('To Device Area')

        contr_map = {'HDC': 0, 'NAUTICS': 1, 'ST': 2}

        new_cbl = {}
        for row in rows_iter:
            cid = row[col_cid] if len(row) > col_cid else None
            if not cid:
                continue
            cid = str(cid).strip()

            c_name = str((row[col_contr] if len(row) > col_contr else '') or '').strip().upper()
            c_idx = contr_map.get(c_name, 3)

            cut_v = to_int(row[col_cut] if len(row) > col_cut else 0, 0)
            pulled_v = to_int(row[col_pulled] if len(row) > col_pulled else 0, 0)
            conn_v = to_int(row[col_conn] if len(row) > col_conn else 0, 0)
            megger_v = to_int(row[col_megger] if len(row) > col_megger else 0, 0)

            fa = str((row[col_fa] if len(row) > col_fa else '') or '').strip()
            ta = str((row[col_ta] if len(row) > col_ta else '') or '').strip()

            new_cbl[cid] = [c_idx, pulled_v, cut_v, conn_v, fa, ta, megger_v]
    finally:
        wb.close()

    with open(HTML_FILE, 'r', encoding='utf-8') as f:
        html = f.read()

    pos1 = html.find('"cbl":{')
    if pos1 == -1:
        print('⚠️ Không tìm thấy "cbl":{ trong index.html')
        return 0
    pos2 = html.find('}};', pos1)
    if pos2 == -1:
        print('⚠️ Không tìm thấy kết thúc cbl trong index.html')
        return 0

    cur_cbl = json.loads(html[pos1+6:pos2+1])
    diff_count = 0
    for cid, nval in new_cbl.items():
        if cur_cbl.get(cid) != nval:
            diff_count += 1
            cur_cbl[cid] = nval

    if diff_count > 0:
        new_cbl_str = json.dumps(cur_cbl, separators=(',', ':'))
        new_html = html[:pos1+6] + new_cbl_str + html[pos2+1:]
        with open(HTML_FILE, 'w', encoding='utf-8') as f:
            f.write(new_html)
        print(f"✅ Đã cập nhật {diff_count} dây thay đổi trạng thái vào index.html")
    else:
        print("ℹ️ Trạng thái cáp trên bản vẽ index.html đã hoàn toàn khớp (0 thay đổi)")

    return diff_count


def git_commit_and_push(target_week):
    try:
        sw_path = os.path.join(BASE_DIR, 'sw.js')
        if os.path.exists(sw_path):
            try:
                with open(sw_path, 'r', encoding='utf-8') as f_sw:
                    sw_c = f_sw.read()
                m_sw = re.search(r"CACHE_NAME = 'tsc80-pwa-v(\d+\.\d+)'", sw_c)
                if m_sw:
                    cur_v = float(m_sw.group(1))
                    new_v = f"{cur_v + 0.1:.1f}"
                    sw_c = sw_c.replace(m_sw.group(0), f"CACHE_NAME = 'tsc80-pwa-v{new_v}'")
                    with open(sw_path, 'w', encoding='utf-8') as f_sw:
                        f_sw.write(sw_c)
                    subprocess.run(['git', '-C', BASE_DIR, 'add', 'sw.js'], check=False)
            except Exception as e:
                print('⚠️ Không thể tăng cache sw.js:', e)

        subprocess.run(
            ['git', '-C', BASE_DIR, 'add',
             os.path.basename(OVERVIEW_FILE),
             os.path.basename(JSON_OUT),
             os.path.basename(HTML_FILE),
             os.path.basename(MASTER_FILE)],
            check=True
        )
        st = subprocess.run(['git', '-C', BASE_DIR, 'diff', '--cached', '--name-only'], capture_output=True, text=True)
        if not st.stdout.strip():
            print("ℹ️ Không có thay đổi mới nào so với bản trên GitHub. Bỏ qua commit & push.")
            return
        now_str = datetime.now().strftime('%d/%m/%Y %H:%M')
        msg = f"feat(prog): cap nhat tien do tuan {target_week} luc {now_str}"
        subprocess.run(['git', '-C', BASE_DIR, 'commit', '-m', msg], check=False)
        res = subprocess.run(['git', '-C', BASE_DIR, 'push', 'origin', 'main'], capture_output=True, text=True)
        if res.returncode == 0:
            print("🚀 Đã tự động đẩy lên GitHub (nhánh main)! Web và điện thoại sẽ nhận bản mới sau 30s.")
        else:
            print("⚠️ Git push gặp lỗi (có thể kiểm tra mạng):", res.stderr)
    except Exception as e:
        print("⚠️ Lỗi git commit/push:", e)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Đồng bộ tiến độ hàng tuần từ tất cả các file Excel trong folder")
    parser.add_argument('--file', default=MASTER_FILE, help="Đường dẫn file Excel master tiến độ tuần")
    parser.add_argument('--week', default=None, help="Mã tuần cho file master (ví dụ W48). Mặc định tự tính.")
    parser.add_argument('--no-push', action='store_true', help="Không tự động git push")
    args = parser.parse_args()

    file_path = args.file if os.path.isabs(args.file) else os.path.join(BASE_DIR, args.file)
    target_week = args.week
    if not target_week:
        m = re.search(r'W(\d{1,2})', os.path.basename(file_path), re.I)
        if m:
            target_week = f"W{int(m.group(1)):02d}"
        else:
            target_week = get_current_report_week()

    print(f"📂 Đang quét tất cả các file tiến độ tuần trong: {BASE_DIR}")
    week_files = discover_weekly_files(BASE_DIR, master_file=file_path, master_week=target_week)
    if not week_files:
        print("❌ Không tìm thấy file tiến độ tuần nào trong folder!")
        sys.exit(1)

    week_data_map = {}
    for wcode in sorted(week_files.keys()):
        fpath, _ = week_files[wcode]
        try:
            data = read_tong_sheet(fpath)
            week_data_map[wcode] = data
            hdc, nau, st = data['HDC'], data['NAUTICS'], data['ST']
            print(f"  ✔ {wcode} <- {os.path.basename(fpath)} | "
                  f"HDC({hdc['cut']},{hdc['pull']},{hdc['conn']}) "
                  f"NAU({nau['cut']},{nau['pull']},{nau['conn']}) "
                  f"ST({st['cut']},{st['pull']},{st['conn']})")
        except Exception as e:
            print(f"  ⚠️ Bỏ qua {wcode} ({os.path.basename(fpath)}): {e}")

    print(f"\n🎯 Cập nhật Overview cho {len(week_data_map)} tuần (tuần hiện tại: {target_week})")
    wb_ov, col_values = update_overview_from_folder(week_data_map)
    generate_prog_json(wb_ov, col_values)

    if os.path.exists(file_path):
        sync_cables_to_html(file_path)

    if not args.no_push:
        print("\n📤 Đang đẩy thay đổi lên GitHub...")
        git_commit_and_push(target_week)

    print(f"\n🎉 HOÀN THÀNH ĐỒNG BỘ ({len(week_data_map)} tuần)!")


if __name__ == '__main__':
    main()
