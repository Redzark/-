import streamlit as st
import openpyxl
import io
import uuid
import re
import traceback
from openpyxl.styles import Alignment

# ============================================================================
# 1. 기초 데이터 및 설정
# ============================================================================
MAT_START_ROW = 12
MAT_STEP = 4
LAB_START_ROW = 55
EXP_START_ROW = 75

YEARLY_LABOR_RATES = {2025: 25700, 2026: 30000, 2027: 31600, 2028: 33200, 2029: 34800}
DIRECT_EXP_TABLE = {50: 2042, 70: 2248, 100: 2735, 120: 2819, 150: 3230, 170: 3219, 220: 4404, 250: 4861, 300: 6210, 350: 6349, 450: 8204, 500: 7940, 550: 9228, 600: 10009, 650: 11482, 700: 11488, 750: 13458, 850: 14604, 900: 15154, 1050: 17575, 1300: 21270, 1600: 23872, 1800: 27671, 2000: 27671, 2200: 33488, 2300: 33488, 2400: 33488, 2500: 33488, 3000: 48003}
MATERIAL_DATA = {
    "무도장 TPO": {"coeff": 2.58, "f12": "무도장 TPO", "f13": "MS220-19 TYPE B-2"},
    "도장용 TPO": {"coeff": 3.56, "f12": "도장용 TPO", "f13": "MS220-19 TYPE B-1"},
    "ASA": {"coeff": 3.11, "f12": "ASA-022 TYPE B", "f13": "MS225-22"},
    "도금용 ABS": {"coeff": 3.62, "f12": "도금용 ABS", "f13": "MS225-20"},
    "도장용 ABS": {"coeff": 3.62, "f12": "도장용 ABS", "f13": "MS225-18 TYPE C"},
    "PP": {"coeff": 2.58, "f12": "PP", "f13": "General"}
}
DRY_CYCLE_MAP = {50:10, 70:11, 100:12, 120:13, 150:14, 170:14, 220:15, 280:16, 350:19, 450:21, 500:21, 550:21, 600:22, 650:22, 700:23, 750:23, 850:26, 900:26, 1050:26, 1300:28, 1600:30, 1800:31, 2000:32, 2200:36, 2300:37, 2400:37, 2500:38, 3000:44}

# ============================================================================
# 2. 로직 함수 (안전장치 강화)
# ============================================================================
def safe_float(value, default=0.0):
    try:
        if value is None: return default
        s_val = str(value).strip().upper()
        if not s_val: return default
        
        for sep in ['\n', '(', '\r']:
            if sep in s_val: s_val = s_val.split(sep)[0].strip()
        
        if "/" in s_val:
            parts = s_val.split("/")
            if parts[0].strip(): s_val = parts[0]
        
        clean_val = re.sub(r"[^0-9.]", "", s_val)
        if not clean_val: return default
        if clean_val == ".": return default
        
        return float(clean_val)
    except: return default

def get_loss_rate(real_vol):
    if real_vol <= 3000: return 0.049
    elif real_vol <= 5000: return 0.032
    elif real_vol <= 10000: return 0.019
    elif real_vol <= 20000: return 0.010
    elif real_vol <= 40000: return 0.006
    elif real_vol <= 80000: return 0.006
    elif real_vol <= 100000: return 0.005
    elif real_vol <= 200000: return 0.005
    elif real_vol <= 300000: return 0.003
    elif real_vol <= 400000: return 0.002
    elif real_vol <= 600000: return 0.002
    elif real_vol <= 800000: return 0.002
    else: return 0.001 

def get_lot_size(L, W, H, real_vol):
    max_dim = max(L, W, H)
    idx = 1 if max_dim <= 100 else (3 if max_dim > 1500 else 2)
    return 5000 if idx==1 else (3000 if idx==2 else 1500)

def get_manpower(ton, mat_name):
    if "도금" in mat_name: return 0.5 if ton <= 150 else 1.0
    return 0.5 if ton < 650 else 1.0

def get_setup_time(ton):
    if ton <= 150: return 25
    elif ton < 650: return 30
    elif ton < 2000: return 35
    else: return 45

def get_sr_rate_value(w, c):
    total = w * c
    if total <= 3: return 240
    elif total <= 5: return 160
    elif total <= 10: return 90
    elif total <= 30: return 35
    elif total <= 50: return 22
    elif total <= 200: return 8
    elif total <= 1500: return 5
    elif total <= 3000: return 4
    else: return 3

def get_machine_factor(ton):
    if ton < 150: return 0.9
    elif ton < 300: return 1.0
    elif ton < 650: return 1.05
    elif ton < 1300: return 1.1
    elif ton < 1800: return 1.2
    else: return 1.3

def get_depth_factor(h):
    if h <= 50: return 0.8
    elif h <= 100: return 0.9
    elif h <= 150: return 0.95
    elif h <= 200: return 1.0
    elif h <= 250: return 1.05
    elif h <= 300: return 1.1
    elif h <= 400: return 1.15
    else: return 1.2

def safe_write(ws, coord, value):
    try: ws[coord] = value
    except Exception: pass

# ============================================================================
# 3. PART LIST 파싱 함수
# ============================================================================
def extract_header_info(ws):
    extracted = {"car": "", "vol": 0}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=150, values_only=True)):
        for j, cell in enumerate(row):
            if not cell: continue
            s_val = str(cell).replace(" ","").upper()
            if "차종" in s_val or "PROJECT" in s_val:
                for k in range(j + 1, len(row)):
                    if row[k]: extracted["car"] = str(row[k]).strip(); break
            if "생산대수" in s_val or "VOLUME" in s_val or "볼륨" in s_val or "생산량" in s_val:
                for k in range(j, min(j + 10, len(row))): 
                    val = safe_float(row[k])
                    if val > 0: extracted["vol"] = val; break 
    return extracted

def parse_part_list_dynamic(file):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        
        parsed_items = []
        header_info = extract_header_info(ws) # 여기서 header_info로 받음
        
        all_rows = list(ws.iter_rows(values_only=True))
        header_row_index = -1
        col_map = {'part_no': 7, 'name': 8, 'qty': [], 'mat': 22, 'ton': 23, 'cav': 24, 'L':10, 'W':11, 'H':12} 
        
        for i, r in enumerate(all_rows):
            row_str = " ".join([str(x) for x in r if x]).replace(" ", "").upper()
            if "PARTNO" in row_str or "품번" in row_str:
                header_row_index = i
                row1 = r
                row2 = all_rows[i+1] if i+1 < len(all_rows) else [None]*len(r)
                
                for idx, cell in enumerate(row1):
                    if not cell: continue
                    c_val = str(cell).upper().replace(" ", "").replace("\n", "")
                    if "PARTNO" in c_val or "품번" in c_val: col_map['part_no'] = idx
                    elif "PARTNAME" in c_val or "품명" in c_val: col_map['name'] = idx
                    elif "MATERIAL" in c_val or "재질" in c_val: col_map['mat'] = idx
                    elif "THICK" in c_val or "두께" in c_val: col_map['thick'] = idx
                    elif "WEIGHT" in c_val or "중량" in c_val: col_map['weight'] = idx
                    if "QTY" in c_val or "수량" in c_val or "USG" in c_val:
                        if idx not in col_map['qty']: col_map['qty'].append(idx)
                            
                for idx, cell in enumerate(row2):
                    if not cell: continue
                    c_val = str(cell).upper().replace(" ", "").replace("\n", "")
                    if "가로" in c_val or "L" == c_val or "LENGTH" in c_val: col_map['L'] = idx
                    elif "세로" in c_val or "W" == c_val or "WIDTH" in c_val: col_map['W'] = idx
                    elif "깊이" in c_val or "높이" in c_val or "H" == c_val: col_map['H'] = idx
                    elif "TON" in c_val or "톤" in c_val: col_map['ton'] = idx
                    elif "C/V" in c_val or "CAV" in c_val: col_map['cav'] = idx
                    elif "THICK" in c_val or "두께" in c_val: col_map['thick'] = idx
                    elif "WEIGHT" in c_val or "중량" in c_val: col_map['weight'] = idx
                    if "QTY" in c_val or "수량" in c_val:
                        if idx not in col_map['qty']: col_map['qty'].append(idx)
                break

        if header_row_index == -1: header_row_index = 5 

        for i in range(header_row_index + 1, len(all_rows)):
            r = list(all_rows[i])
            if len(r) < 100: r.extend([None] * (100 - len(r)))
            
            p_idx = col_map.get('part_no', 7)
            if p_idx >= len(r) or not r[p_idx]: continue
            p_no_str = str(r[p_idx]).strip()
            clean_p_no = p_no_str.replace(" ", "").upper()
            
            if "PARTNO" in clean_p_no or "품번" in clean_p_no: continue
            if "비고" in clean_p_no or "REMARK" in clean_p_no: continue
            
            n_idx = col_map.get('name', 8)
            rem_val = str(r[n_idx + 1] if n_idx + 1 < len(r) and r[n_idx+1] else "")
            
            t_idx = col_map.get('ton', 28)
            m_idx = col_map.get('mat', 27) 
            raw_ton = r[t_idx] if t_idx < len(r) else None
            raw_mat = r[m_idx] if m_idx < len(r) else None
            
            if not safe_float(raw_ton) and (not raw_mat or str(raw_mat).strip() == ""):
                continue

            usage = 1.0
            if col_map.get('qty'):
                max_q = 0
                for q_idx in col_map['qty']:
                    if q_idx < len(r):
                        v = safe_float(r[q_idx])
                        if v > max_q: max_q = v
                if max_q > 0: usage = max_q

            p_name = str(r[n_idx]).strip() if n_idx < len(r) and r[n_idx] else ""
            
            l = safe_float(r[col_map.get('L', 13)])
            w = safe_float(r[col_map.get('W', 14)])
            h = safe_float(r[col_map.get('H', 15)])
            
            t_col = col_map.get('thick')
            t = safe_float(r[t_col]) if t_col and t_col < len(r) else 2.5
            if t == 0: t = 2.5

            w_col = col_map.get('weight')
            weight_val = safe_float(r[w_col]) if w_col and w_col < len(r) else 0.0

            mapped_mat = "무도장 TPO"
            if raw_mat:
                s_mat = str(raw_mat).upper()
                for key in MATERIAL_DATA.keys():
                    if key in s_mat:
                        mapped_mat = key; break
                if "PP" in s_mat and mapped_mat == "무도장 TPO": mapped_mat = "PP"

            ton = int(safe_float(raw_ton, default=1300))
            
            cv_idx = col_map.get('cav', t_idx + 1)
            raw_cav = str(r[cv_idx]) if cv_idx < len(r) else "1"
            
            if "/" in raw_cav:
                try:
                    cav = int(sum(safe_float(x) for x in raw_cav.split('/') if x.strip()))
                except:
                    cav = int(safe_float(raw_cav, default=1))
            else:
                cav = int(safe_float(raw_cav, default=1))
            
            if cav < 1: cav = 1

            item = {
                "id": str(uuid.uuid4()),
                "level": "사출제품",
                "no": p_no_str,
                "name": p_name,
                "remarks": rem_val,
                "opt_rate": 100.0,
                "usage": usage,
                "L": l, "W": w, "H": h, "thick": t,
                "weight": weight_val,
                "mat": mapped_mat,
                "ton": ton,
                "cavity": cav,
                "price": 2000
            }
            parsed_items.append(item)

        # ▼▼▼ [수정완료] extracted_info -> header_info로 수정 ▼▼▼
        return parsed_items, header_info 

    except Exception as e:
        st.error(f"데이터 분석 중 오류 발생: {e}")
        st.code(traceback.format_exc())
        return [], {}

# ============================================================================
# 4. 엑셀 생성 함수
# ============================================================================
def generate_excel_batch_sheets(common, items, sel_year):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        template_sheet = wb.active
        template_sheet.title = "Master_Template"
    except: return None

    align_center = Alignment(horizontal='center', vertical='center')

    for item in items:
        safe_title = str(item['no']).replace("/", "_").replace("*", "")[:30]
        
        if "비고" in safe_title or "REMARK" in safe_title: continue

        target_sheet = wb.copy_worksheet(template_sheet)
        target_sheet.title = safe_title
        ws = target_sheet

        safe_write(ws, "N3", common['car'])
        safe_write(ws, "C3", item['no'])    
        safe_write(ws, "C4", item['name']) 
        
        curr_m = MAT_START_ROW
        
        item_usage = item.get('usage', 1.0)
        real_vol = common['base_vol'] * (item['opt_rate'] / 100) * item_usage
        
        loss_val = get_loss_rate(real_vol)

        safe_write(ws, f"B{curr_m}", item['name'])
        safe_write(ws, f"B{curr_m+1}", item['no'])
        
        mat_info = MATERIAL_DATA.get(item['mat'], MATERIAL_DATA["무도장 TPO"])
        try: ws.merge_cells(f"F{curr_m}:G{curr_m}"); ws.merge_cells(f"F{curr_m+1}:G{curr_m+1}")
        except: pass

        safe_write(ws, f"F{curr_m}", mat_info['f12'])
        safe_write(ws, f"F{curr_m+1}", mat_info['f13'])
        if ws[f"F{curr_m}"]: ws[f"F{curr_m}"].alignment = align_center
        if ws[f"F{curr_m+1}"]: ws[f"F{curr_m+1}"].alignment = align_center
        
        safe_write(ws, f"D{curr_m}", real_vol) 
        if ws[f"D{curr_m}"]: ws[f"D{curr_m}"].number_format = '#,##0'
        
        safe_write(ws, f"J{curr_m}", item['weight']/1000); safe_write(ws, f"K{curr_m}", item['price'])
        safe_write(ws, f"H{curr_m}", 1.0); safe_write(ws, f"I{curr_m}", "kg")
        if ws[f"I{curr_m}"]: ws[f"I{curr_m}"].alignment = align_center
        safe_write(ws, f"L{curr_m}", f"=(J{curr_m}*(1+{loss_val}))*K{curr_m}*H{curr_m}")
        
        sr_val = get_sr_rate_value(item['weight'], item['cavity'])
        safe_write(ws, f"J{curr_m+1}", f"=J{curr_m} * {sr_val} / 100")
        safe_write(ws, f"K{curr_m+1}", 87); safe_write(ws, f"H{curr_m+1}", 1.0)
        safe_write(ws, f"I{curr_m+1}", "kg")
        if ws[f"I{curr_m+1}"]: ws[f"I{curr_m+1}"].alignment = align_center
        safe_write(ws, f"L{curr_m+1}", f"=J{curr_m+1}*K{curr_m+1}*H{curr_m+1}")

        l_row, e_row = LAB_START_ROW, EXP_START_ROW
        setup, lot = get_setup_time(item['ton']), get_lot_size(item['L'], item['W'], item['H'], real_vol)
        mp, l_rate, e_rate = get_manpower(item['ton'], item['mat']), YEARLY_LABOR_RATES[sel_year], DIRECT_EXP_TABLE.get(item['ton'], 7940)
        
        safe_write(ws, f"B{l_row}", item['name'])
        safe_write(ws, f"F{l_row}", setup)
        if ws[f"F{l_row}"]: ws[f"F{l_row}"].alignment = align_center
        safe_write(ws, f"G{l_row}", lot)
        if ws[f"G{l_row}"]: ws[f"G{l_row}"].alignment = align_center
        
        safe_write(ws, f"H{l_row}", item['cavity']); safe_write(ws, f"I{l_row}", mp); safe_write(ws, f"K{l_row}", l_rate)
        safe_write(ws, f"E{l_row}", 1.0)

        mf, hf, dry = get_machine_factor(item['ton']), get_depth_factor(item.get('H', 100)), DRY_CYCLE_MAP.get(item['ton'], 44)
        ct_formula = f"={dry}+(4.396*((SUM(J{curr_m}:J{curr_m+1})*H{l_row})*1000)^0.1477)+({MATERIAL_DATA.get(item['mat'], MATERIAL_DATA['무도장 TPO'])['coeff']}*{item.get('thick', 2.5)}^2*{mf}*{hf})"
        if item['mat'] == "도금용 ABS": ct_formula += "+15"
        
        safe_write(ws, f"J{l_row}", ct_formula)
        safe_write(ws, f"L{l_row}", f"=(J{l_row}*1.1/H{l_row}+F{l_row}*60/G{l_row})*I{l_row}*K{l_row}/3600*E{l_row}") 

        safe_write(ws, f"B{e_row}", item['name'])
        safe_write(ws, f"F{e_row}", setup); safe_write(ws, f"G{e_row}", lot); safe_write(ws, f"H{e_row}", item['cavity'])
        safe_write(ws, f"I{e_row}", item['ton'])
        if ws[f"I{e_row}"]: ws[f"I{e_row}"].number_format = '#,##0"T"'
        safe_write(ws, f"J{e_row}", f"=J{l_row}"); safe_write(ws, f"K{e_row}", e_rate)
        safe_write(ws, f"E{e_row}", 1.0) 
        safe_write(ws, f"L{e_row}", f"=(J{l_row}*1.1/H{e_row}+F{e_row}*60/G{e_row})*K{e_row}/3600*(1+0.64)")

    if "Master_Template" in wb.sheetnames: wb.remove(wb["Master_Template"])
    output = io.BytesIO(); wb.save(output); return output.getvalue()

def generate_excel_final_mapping(common, items, sel_year):
    try:
        wb = openpyxl.load_workbook("template.xlsx"); ws = wb.active
    except: return None
    safe_write(ws, "N3", common['car']); safe_write(ws, "C3", common['no']); safe_write(ws, "C4", common['name'])
    align_center = Alignment(horizontal='center', vertical='center')
    for i, item in enumerate(items):
        curr_m = MAT_START_ROW + (i * MAT_STEP)
        real_vol = common['base_vol'] * (item['opt_rate'] / 100)
        loss_val = get_loss_rate(real_vol)
        safe_write(ws, f"B{curr_m}", item['name']); safe_write(ws, f"B{curr_m+1}", item['no'])
        mat_info = MATERIAL_DATA.get(item['mat'], MATERIAL_DATA["무도장 TPO"])
        try: ws.merge_cells(f"F{curr_m}:G{curr_m}"); ws.merge_cells(f"F{curr_m+1}:G{curr_m+1}")
        except: pass
        safe_write(ws, f"F{curr_m}", mat_info['f12']); safe_write(ws, f"F{curr_m+1}", mat_info['f13'])
        if ws[f"F{curr_m}"]: ws[f"F{curr_m}"].alignment = align_center
        if ws[f"F{curr_m+1}"]: ws[f"F{curr_m+1}"].alignment = align_center
        safe_write(ws, f"D{curr_m}", real_vol)
        if ws[f"D{curr_m}"]: ws[f"D{curr_m}"].number_format = '#,##0'
        safe_write(ws, f"J{curr_m}", item['weight']/1000); safe_write(ws, f"K{curr_m}", item['price'])
        safe_write(ws, f"H{curr_m}", 1.0); safe_write(ws, f"I{curr_m}", "kg")
        if ws[f"I{curr_m}"]: ws[f"I{curr_m}"].alignment = align_center
        safe_write(ws, f"L{curr_m}", f"=(J{curr_m}*(1+{loss_val}))*K{curr_m}*H{curr_m}")
        sr_val = get_sr_rate_value(item['weight'], item['cavity'])
        safe_write(ws, f"J{curr_m+1}", f"=J{curr_m} * {sr_val} / 100")
        safe_write(ws, f"K{curr_m+1}", 87); safe_write(ws, f"H{curr_m+1}", 1.0)
        safe_write(ws, f"I{curr_m+1}", "kg")
        if ws[f"I{curr_m+1}"]: ws[f"I{curr_m+1}"].alignment = align_center
        safe_write(ws, f"L{curr_m+1}", f"=J{curr_m+1}*K{curr_m+1}*H{curr_m+1}")
    inject_items = [x for x in items if "사출" in x['level']]
    for i, item in enumerate(inject_items):
        l_row, e_row = LAB_START_ROW + i, EXP_START_ROW + i
        real_vol = common['base_vol'] * (item['opt_rate'] / 100)
        setup, lot = get_setup_time(item['ton']), get_lot_size(item['L'], item['W'], item['H'], real_vol)
        mp, l_rate, e_rate = get_manpower(item['ton'], item['mat']), YEARLY_LABOR_RATES[sel_year], DIRECT_EXP_TABLE.get(item['ton'], 7940)
        safe_write(ws, f"B{l_row}", item['name']); safe_write(ws, f"F{l_row}", setup)
        if ws[f"F{l_row}"]: ws[f"F{l_row}"].alignment = align_center
        safe_write(ws, f"G{l_row}", lot)
        if ws[f"G{l_row}"]: ws[f"G{l_row}"].alignment = align_center
        safe_write(ws, f"H{l_row}", item['cavity']); safe_write(ws, f"I{l_row}", mp); safe_write(ws, f"K{l_row}", l_rate)
        safe_write(ws, f"E{l_row}", 1.0)
        mf, hf, dry = get_machine_factor(item['ton']), get_depth_factor(item.get('H', 100)), DRY_CYCLE_MAP.get(item['ton'], 44)
        m_row = MAT_START_ROW + (items.index(item) * MAT_STEP)
        ct_formula = f"={dry}+(4.396*((SUM(J{m_row}:J{m_row+1})*H{l_row})*1000)^0.1477)+({MATERIAL_DATA.get(item['mat'], MATERIAL_DATA['무도장 TPO'])['coeff']}*{item.get('thick', 2.5)}^2*{mf}*{hf})"
        if item['mat'] == "도금용 ABS": ct_formula += "+15"
        safe_write(ws, f"J{l_row}", ct_formula)
        safe_write(ws, f"L{l_row}", f"=(J{l_row}*1.1/H{l_row}+F{l_row}*60/G{l_row})*I{l_row}*K{l_row}/3600*E{l_row}") 
        safe_write(ws, f"B{e_row}", item['name']); safe_write(ws, f"F{e_row}", setup); safe_write(ws, f"G{e_row}", lot); safe_write(ws, f"H{e_row}", item['cavity'])
        safe_write(ws, f"I{e_row}", item['ton'])
        if ws[f"I{e_row}"]: ws[f"I{e_row}"].number_format = '#,##0"T"'
        safe_write(ws, f"J{e_row}", f"=J{l_row}"); safe_write(ws, f"K{e_row}", e_rate)
        safe_write(ws, f"E{e_row}", 1.0) 
        safe_write(ws, f"L{e_row}", f"=(J{l_row}*1.1/H{e_row}+F{e_row}*60/G{e_row})*K{e_row}/3600*(1+0.64)")
    output = io.BytesIO(); wb.save(output); return output.getvalue()

# ============================================================================
# 5. Streamlit UI
# ============================================================================
st.set_page_config(page_title="원가계산서(초안)", layout="wide")
st.title("원가계산서(사출품목 초안 버전)")
st.warning("⚠️ 사출품목만 가능하니 사출품목 外 계산은 할 수 없습니다.")

if 'bom_master' not in st.session_state: st.session_state.bom_master = []
if 'common_car' not in st.session_state: st.session_state.common_car = ""
if 'common_vol' not in st.session_state: st.session_state.common_vol = 0
if 'excel_data' not in st.session_state: st.session_state.excel_data = None

mode = st.radio("작업 모드 선택", ["단품 계산", "ASSY(수동 입력)", "PART LIST 엑셀 업로드"], horizontal=True)

if mode == "PART LIST 엑셀 업로드":
    st.markdown("### 📂 1단계: 파일 업로드")
    uploaded_file = st.file_uploader("PART LIST 엑셀 파일(.xlsx)을 올려주세요.", type=["xlsx", "xls"])
    
    if uploaded_file:
        if st.button("🔄 데이터 불러오기", type="primary"):
            new_items, info = parse_part_list_dynamic(uploaded_file)
            if new_items:
                st.session_state.bom_master = new_items 
                if info.get('car'): st.session_state.common_car = info['car']
                
                if info.get('vol'): 
                    st.session_state.common_vol = info['vol']
                    st.session_state['common_vol_in'] = int(info['vol'])
                
                st.success(f"✅ {len(new_items)}개 품목(사출품)을 성공적으로 불러왔습니다!")
            else:
                st.error("데이터를 읽지 못했습니다. 엑셀 양식을 확인해주세요.")

st.markdown("---")
with st.expander("📌 정보 확인 및 수정", expanded=True):
    c1, c2, c3, c4 = st.columns(4)
    car = st.text_input("차종", value=st.session_state.common_car, key="common_car_in")
    p_no = st.text_input("대표 품번", value="", key="p_no_in")
    p_name = st.text_input("대표 품명", value="", key="p_name_in")
    base_vol = st.number_input("기본 Volume (대)", value=int(st.session_state.common_vol), key="common_vol_in")

    if mode == "단품 계산" and not st.session_state.bom_master:
         st.session_state.bom_master = [{"id":str(uuid.uuid4()), "level":"사출제품", "no":"", "name":"", "opt_rate":100.0, "L":0.0, "W":0.0, "H":0.0, "thick":2.5, "weight":0.0, "mat":"무도장 TPO", "ton":1300, "cavity":1, "price":2000}]
    elif mode == "ASSY(수동 입력)":
        if st.button("➕ 품목 추가"):
            st.session_state.bom_master.append({"id":str(uuid.uuid4()), "level":"사출제품", "no":"", "name":"", "opt_rate":100.0, "L":0.0, "W":0.0, "H":0.0, "thick":2.5, "weight":0.0, "mat":"무도장 TPO", "ton":1300, "cavity":1, "price":2000})

    for i, item in enumerate(st.session_state.bom_master):
        uid = item['id']
        with st.container(border=True):
            c = st.columns([2,2,2,1.5,0.5])
            item['level'] = c[0].selectbox("구분", ["사출제품", "부자재"], key=f"lv_{uid}")
            item['no'] = c[1].text_input("품번", value=item['no'], key=f"n_{uid}")
            item['name'] = c[2].text_input("품명", value=item['name'], key=f"nm_{uid}")
            item['opt_rate'] = c[3].number_input("옵션율(%)", value=item['opt_rate'], key=f"op_{uid}")
            
            usage_val = item.get('usage', 1.0)
            c[4].write(f"Qty: {usage_val}")

            if mode == "ASSY(수동 입력)":
                if c[4].button("🗑️", key=f"d_{uid}"): st.session_state.bom_master.pop(i); st.rerun()
            
            if "사출" in item['level']:
                r = st.columns(5)
                item['L'] = r[0].number_input("L", value=item['L'], key=f"l_{uid}")
                item['W'] = r[1].number_input("W", value=item['W'], key=f"w_{uid}")
                item['H'] = r[2].number_input("H", value=item['H'], key=f"h_{uid}")
                item['thick'] = r[3].number_input("T", value=item['thick'], key=f"t_{uid}")
                item['weight'] = r[4].number_input("중량(g)", value=item['weight'], key=f"g_{uid}")
                r2 = st.columns(3)
                mat_idx = 0
                if item['mat'] in MATERIAL_DATA: mat_idx = list(MATERIAL_DATA.keys()).index(item['mat'])
                item['mat'] = r2[0].selectbox("소재", list(MATERIAL_DATA.keys()), index=mat_idx, key=f"ma_{uid}")
                ton_idx = 20
                if item['ton'] in DIRECT_EXP_TABLE: ton_idx = list(DIRECT_EXP_TABLE.keys()).index(item['ton'])
                item['ton'] = r2[1].selectbox("Ton", list(DIRECT_EXP_TABLE.keys()), index=ton_idx, key=f"to_{uid}")
                current_cav = item.get('cavity', 1)
                if current_cav < 1: current_cav = 1
                item['cavity'] = r2[2].number_input("Cav", min_value=1, value=int(current_cav), key=f"ca_{uid}")
            item['price'] = st.number_input("단가", value=item['price'], key=f"pr_{uid}")

st.markdown("### 💰 엑셀 산출")
if st.button("엑셀 파일 생성하기 (Click)", type="primary", use_container_width=True):
    if mode == "PART LIST 엑셀 업로드":
        st.session_state.excel_data = generate_excel_batch_sheets(
            {"car":car, "base_vol":base_vol}, 
            st.session_state.bom_master, 
            2026
        )
    else:
        p_no_val = st.session_state.get('p_no_in', '')
        p_name_val = st.session_state.get('p_name_in', '')
        st.session_state.excel_data = generate_excel_final_mapping(
            {"car":car, "no":p_no_val, "name":p_name_val, "base_vol":base_vol, "is_assy":(mode!="단품 계산")}, 
            st.session_state.bom_master, 
            2026
        )
    
    if st.session_state.excel_data:
        st.success("생성이 완료되었습니다. 아래 버튼을 눌러 다운로드하세요.")

if st.session_state.excel_data:
    st.download_button(
        label="📥 결과물 다운로드 (Download)",
        data=st.session_state.excel_data,
        file_name="Cost_Calculation_Result.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
