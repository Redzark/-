import streamlit as st
import openpyxl
import io
import uuid
import re
import traceback
import zipfile
from copy import copy
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ============================================================================
# 1. 기초 데이터 및 설정
# ============================================================================
MAT_START_ROW = 12
MAT_STEP = 4
LAB_START_ROW = 55
EXP_START_ROW = 75
TEMPLATE_HEIGHT = 85 

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
# 2. 로직 함수
# ============================================================================
def safe_float(value, default=0.0):
    try:
        if value is None: return default
        s_val = str(value).strip().upper()
        if not s_val or s_val in [".", "-", ""]: return default
        if "/" in s_val: s_val = s_val.split("/")[0]
        clean_val = re.sub(r"[^0-9.]", "", s_val)
        return float(clean_val) if clean_val else default
    except: return default

def get_loss_rate(real_vol):
    if real_vol <= 3000: return 0.049
    elif real_vol <= 5000: return 0.032
    elif real_vol <= 10000: return 0.019
    elif real_vol <= 20000: return 0.010
    elif real_vol <= 40000: return 0.006
    elif real_vol <= 80000: return 0.006
    else: return 0.005 

def get_lot_size(L, W, H, real_vol):
    return 3000

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
    if total <= 10: return 90
    elif total <= 50: return 22
    else: return 5

def get_machine_factor(ton):
    if ton < 150: return 0.9
    elif ton < 650: return 1.05
    else: return 1.3

def get_depth_factor(h):
    if h <= 100: return 0.9
    else: return 1.1

def safe_write(ws, coord, value):
    try: ws[coord] = value
    except: pass

# ============================================================================
# 3. 파싱 함수 (강력한 매트릭스 스캔)
# ============================================================================
def normalize_header(s):
    if not s: return ""
    return re.sub(r'[^A-Z0-9가-힣]', '', str(s).upper())

def extract_header_info(ws):
    extracted = {"car": "", "vol": 0}
    for i, row in enumerate(ws.iter_rows(min_row=1, max_row=50, values_only=True)):
        for j, cell in enumerate(row):
            s_val = normalize_header(cell)
            if "PROJECT" in s_val or "차종" in s_val:
                if j+1 < len(row): extracted["car"] = str(row[j+1])
            if "VOLUME" in s_val or "생산대수" in s_val:
                if j+1 < len(row): extracted["vol"] = safe_float(row[j+1])
    return extracted

def parse_part_list_matrix(file):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        header_info = extract_header_info(ws)
        all_rows = list(ws.iter_rows(values_only=True))
        
        # 1. 헤더 행 찾기
        header_row_index = -1
        for i in range(min(30, len(all_rows))):
            r = all_rows[i]
            row_norm = "".join([normalize_header(x) for x in r])
            if "PARTNO" in row_norm or "품번" in row_norm:
                header_row_index = i
                break
        if header_row_index == -1: header_row_index = 5

        # 2. 컬럼 매핑
        col_map = {'lv': -1, 'part_no': -1, 'name': -1, 'qty_cols': [], 'ton': -1, 'mat': -1}
        header_row = all_rows[header_row_index]
        for idx, cell in enumerate(header_row):
            s = normalize_header(cell)
            if "LV" in s: col_map['lv'] = idx
            elif "PARTNO" in s or "품번" in s: col_map['part_no'] = idx
            elif "PARTNAME" in s or "품명" in s: col_map['name'] = idx
            elif "TON" in s or "톤" in s: col_map['ton'] = idx
            elif "MATERIAL" in s or "재질" in s: col_map['mat'] = idx
        
        # Qty 컬럼 자동 감지 (데이터 기반)
        qty_candidates = []
        start_check_col = col_map['name'] + 1 if col_map['name'] != -1 else 5
        for c in range(start_check_col, len(header_row)):
            has_data = False
            for r in range(header_row_index + 1, min(header_row_index + 50, len(all_rows))):
                val = str(all_rows[r][c]).strip()
                if val == '1' or val == '●' or val == '1.0':
                    has_data = True
                    break
            if has_data: qty_candidates.append(c)
        col_map['qty_cols'] = qty_candidates
        
        # 추가 매핑
        extra_headers = all_rows[header_row_index+1] if header_row_index+1 < len(all_rows) else []
        col_map.update({'L': -1, 'W': -1, 'H': -1, 'thick': -1, 'weight': -1, 'cav': -1})
        for r_search in [header_row, extra_headers]:
            for idx, cell in enumerate(r_search):
                s = normalize_header(cell)
                if s in ['L', 'LENGTH', '가로']: col_map['L'] = idx
                elif s in ['W', 'WIDTH', '세로']: col_map['W'] = idx
                elif s in ['H', 'HEIGHT', '높이']: col_map['H'] = idx
                elif "THICK" in s or "두께" in s: col_map['thick'] = idx
                elif "WEIGHT" in s or "중량" in s: col_map['weight'] = idx
                elif "CAV" in s or "CV" in s: col_map['cav'] = idx

        # 3. 데이터 파싱
        assy_dict = {} 
        active_parents = {c: None for c in col_map['qty_cols']}
        debug_log = [f"✅ 헤더 위치: {header_row_index+1}행", f"ℹ️ 감지된 수량 열: {len(qty_candidates)}개"]
        lv_col = col_map['lv'] if col_map['lv'] != -1 else 1
        
        for i in range(header_row_index + 1, len(all_rows)):
            r = list(all_rows[i])
            if len(r) < 50: r.extend([None] * (50 - len(r)))
            
            lv_val = str(r[lv_col]).strip()
            is_root = ("●" in lv_val or "1" == lv_val or "LV1" in normalize_header(lv_val))
            
            for q_col in col_map['qty_cols']:
                u_val = safe_float(r[q_col])
                if is_root:
                    if u_val > 0:
                        p_idx, n_idx = col_map['part_no'], col_map['name']
                        raw_no = str(r[p_idx]).strip() if p_idx != -1 and r[p_idx] else ""
                        raw_name = str(r[n_idx]).strip() if n_idx != -1 and r[n_idx] else f"Unknown_{i}"
                        if not raw_no or "ASSY" in raw_no.upper() or "필요" in raw_no:
                            base_name = raw_name.replace("/", "_").replace("*", "")[:30]
                            col_letter = openpyxl.utils.get_column_letter(q_col + 1)
                            key_name = f"{base_name}_{col_letter}"
                        else:
                            key_name = raw_no.replace("/", "_").replace("*", "")
                        active_parents[q_col] = key_name
                        if key_name not in assy_dict: assy_dict[key_name] = []
                    else:
                        active_parents[q_col] = None
                
                curr_parent = active_parents[q_col]
                if curr_parent and u_val > 0:
                    t_idx, m_idx = col_map['ton'], col_map['mat']
                    has_ton = (t_idx != -1 and safe_float(r[t_idx]) > 0)
                    has_mat = (m_idx != -1 and r[m_idx] and str(r[m_idx]).strip())
                    
                    if has_ton or has_mat:
                        item = {
                            "id": str(uuid.uuid4()),
                            "no": str(r[col_map['part_no']]).strip() if col_map['part_no'] != -1 else "",
                            "name": str(r[col_map['name']]).strip() if col_map['name'] != -1 else "",
                            "usage": u_val,
                            "mat": str(r[col_map['mat']]).strip() if has_mat else "무도장 TPO",
                            "ton": int(safe_float(r[col_map['ton']], 1300)),
                            "cavity": int(safe_float(r[col_map['cav']], 1)),
                            "L": safe_float(r[col_map['L']]) if col_map['L'] != -1 else 0,
                            "W": safe_float(r[col_map['W']]) if col_map['W'] != -1 else 0,
                            "H": safe_float(r[col_map['H']]) if col_map['H'] != -1 else 0,
                            "thick": safe_float(r[col_map['thick']], 2.5),
                            "weight": safe_float(r[col_map['weight']]),
                            "price": 2000, "opt_rate": 100.0, "remarks": ""
                        }
                        exists = False
                        for ex in assy_dict[curr_parent]:
                            if ex['no'] == item['no'] and ex['name'] == item['name']: exists = True; break
                        if not exists: assy_dict[curr_parent].append(item)

        final_dict = {k: v for k, v in assy_dict.items() if v}
        return final_dict, header_info, debug_log

    except Exception as e:
        return {}, {}, [f"ERROR: {str(e)}"]

# ============================================================================
# 4. 엑셀 생성 함수 (안정성 강화)
# ============================================================================
def generate_excel_file_stacked(common, items, sel_year):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        template_ws = wb.active
    except: return None

    ws_sum = wb.create_sheet("Summary", 0)
    headers = ["NO", "PART NO", "PART NAME", "USAGE", "MATERIAL", "TON", "CV", "WEIGHT"]
    for i, h in enumerate(headers, 1): ws_sum.cell(1, i, h).font = Font(bold=True)
    for idx, item in enumerate(items, 1):
        row = [idx, item['no'], item['name'], item['usage'], item['mat'], item['ton'], item['cavity'], item['weight']]
        for c, v in enumerate(row, 1): ws_sum.cell(idx+1, c, v)

    ws_main = wb.create_sheet("Calculation", 1)
    offset = 0
    temp_rows = list(template_ws.iter_rows(max_row=TEMPLATE_HEIGHT))
    
    for item in items:
        for r_idx, row in enumerate(temp_rows):
            tgt_r = offset + r_idx + 1
            for c_idx, cell in enumerate(row):
                tgt_c = c_idx + 1
                new_cell = ws_main.cell(tgt_r, tgt_c, cell.value)
                if cell.has_style:
                    new_cell.font = copy(cell.font)
                    new_cell.border = copy(cell.border)
                    new_cell.fill = copy(cell.fill)
                    new_cell.number_format = cell.number_format
                    new_cell.alignment = copy(cell.alignment)
        
        for rng in template_ws.merged_cells.ranges:
            min_col, min_row, max_col, max_row = rng.bounds
            ws_main.merge_cells(start_row=offset+min_row, start_column=min_col,
                                end_row=offset+max_row, end_column=max_col)

        def w(rc, val):
            c_char = re.match(r"([A-Z]+)", rc).group(1)
            r_num = int(re.match(r"[A-Z]+([0-9]+)", rc).group(2))
            col_idx = openpyxl.utils.column_index_from_string(c_char)
            ws_main.cell(offset + r_num, col_idx, val)

        w("N3", common['car'])
        w("C3", item['no'])
        w("C4", item['name'])
        
        m_row = MAT_START_ROW
        real_vol = common['base_vol'] * (item['opt_rate']/100) * item['usage']
        loss = get_loss_rate(real_vol)
        w(f"B{m_row}", item['name'])
        w(f"B{m_row+1}", item['no'])
        w(f"D{m_row}", real_vol)
        w(f"J{m_row}", item['weight']/1000)
        w(f"K{m_row}", item['price'])
        w(f"L{m_row}", f"=(J{offset+m_row}*(1+{loss}))*K{offset+m_row}")
        
        l_row = LAB_START_ROW
        setup = get_setup_time(item['ton'])
        lot = get_lot_size(item['L'], item['W'], item['H'], real_vol)
        mp = get_manpower(item['ton'], item['mat'])
        l_rate = YEARLY_LABOR_RATES[sel_year]
        w(f"B{l_row}", item['name'])
        w(f"F{l_row}", setup)
        w(f"G{l_row}", lot)
        w(f"H{l_row}", item['cavity'])
        w(f"I{l_row}", mp)
        w(f"K{l_row}", l_rate)
        
        mf = get_machine_factor(item['ton'])
        hf = get_depth_factor(item['H'])
        dry = DRY_CYCLE_MAP.get(item['ton'], 40)
        j_curr = f"J{offset+m_row}"
        j_curr_next = f"J{offset+m_row+1}"
        h_l = f"H{offset+l_row}"
        coeff = MATERIAL_DATA.get(item['mat'], {}).get('coeff', 2.58)
        ct_formula = f"={dry}+(4.396*((SUM({j_curr}:{j_curr_next})*{h_l})*1000)^0.1477)+({coeff}*{item['thick']}^2*{mf}*{hf})"
        if item['mat'] == "도금용 ABS": ct_formula += "+15"
        w(f"J{l_row}", ct_formula)
        w(f"L{l_row}", f"=(J{offset+l_row}*1.1/{h_l}+F{offset+l_row}*60/G{offset+l_row})*I{offset+l_row}*K{offset+l_row}/3600")

        e_row = EXP_START_ROW
        e_rate = DIRECT_EXP_TABLE.get(item['ton'], 5000)
        w(f"B{e_row}", item['name'])
        w(f"I{e_row}", item['ton'])
        w(f"J{e_row}", f"=J{offset+l_row}")
        w(f"K{e_row}", e_rate)
        w(f"L{e_row}", f"=(J{offset+l_row}*1.1/H{offset+e_row}+F{offset+e_row}*60/G{offset+e_row})*K{offset+e_row}/3600*(1+0.64)")

        offset += (TEMPLATE_HEIGHT + 2)

    if "Master_Template" in wb.sheetnames: wb.remove(wb["Master_Template"])
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()

# ============================================================================
# 5. UI
# ============================================================================
st.set_page_config(page_title="원가계산서(통합)", layout="wide")
st.title("원가계산서 (단품/수동 + ASSY 통합본)")

if 'manual_items' not in st.session_state: st.session_state.manual_items = []
if 'assy_dict' not in st.session_state: st.session_state.assy_dict = {}
if 'common_car' not in st.session_state: st.session_state.common_car = ""
if 'common_vol' not in st.session_state: st.session_state.common_vol = 0

mode = st.radio("작업 모드 선택", ["단품 계산", "ASSY(수동 입력)", "PART LIST 엑셀 업로드(Matrix)"], horizontal=True)

if mode in ["단품 계산", "ASSY(수동 입력)"]:
    st.info("💡 직접 데이터를 입력하여 계산서를 만듭니다.")
    c1, c2, c3 = st.columns(3)
    car = c1.text_input("차종", value=st.session_state.common_car)
    base_vol = c2.number_input("기본 Volume (대)", value=int(st.session_state.common_vol) if st.session_state.common_vol else 0)

    if mode == "단품 계산" and not st.session_state.manual_items:
        st.session_state.manual_items = [{"id":str(uuid.uuid4()), "level":"사출제품", "no":"", "name":"", "opt_rate":100.0, "usage":1.0, "L":0.0, "W":0.0, "H":0.0, "thick":2.5, "weight":0.0, "mat":"무도장 TPO", "ton":1300, "cavity":1, "price":2000}]
    
    if mode == "ASSY(수동 입력)":
        if st.button("➕ 품목 추가"):
            st.session_state.manual_items.append({"id":str(uuid.uuid4()), "level":"사출제품", "no":"", "name":"", "opt_rate":100.0, "usage":1.0, "L":0.0, "W":0.0, "H":0.0, "thick":2.5, "weight":0.0, "mat":"무도장 TPO", "ton":1300, "cavity":1, "price":2000})

    for i, item in enumerate(st.session_state.manual_items):
        uid = item['id']
        with st.container(border=True):
            cols = st.columns([2, 2, 2, 1, 1, 0.5])
            item['no'] = cols[0].text_input("품번", value=item['no'], key=f"n_{uid}")
            item['name'] = cols[1].text_input("품명", value=item['name'], key=f"nm_{uid}")
            item['opt_rate'] = cols[2].number_input("옵션율(%)", value=item['opt_rate'], key=f"op_{uid}")
            item['usage'] = cols[3].number_input("Qty", value=item['usage'], key=f"us_{uid}")
            if mode == "ASSY(수동 입력)":
                if cols[5].button("🗑️", key=f"d_{uid}"): st.session_state.manual_items.pop(i); st.rerun()
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
            ton_keys = list(DIRECT_EXP_TABLE.keys())
            ton_idx = ton_keys.index(item['ton']) if item['ton'] in ton_keys else ton_keys.index(1300)
            item['ton'] = r2[1].selectbox("Ton", ton_keys, index=ton_idx, key=f"to_{uid}")
            item['cavity'] = r2[2].number_input("Cav", min_value=1, value=int(item['cavity']), key=f"ca_{uid}")
            item['price'] = st.number_input("단가(참고용)", value=item['price'], key=f"pr_{uid}")

    if st.button("엑셀 생성", type="primary"):
        excel_bytes = generate_excel_file_stacked({"car":car, "base_vol":base_vol}, st.session_state.manual_items, 2026)
        if excel_bytes: st.download_button("📥 다운로드", excel_bytes, "Manual_Cost.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

else:
    st.info("💡 엑셀을 올리면 [모든 기둥 + 모든 Lv.1 그룹]을 자동 분석하여 ZIP으로 줍니다.")
    uploaded_file = st.file_uploader("PART LIST 파일 업로드", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🔄 분석 시작"):
            assy_data, info, debug_log = parse_part_list_matrix(uploaded_file)
            with st.expander("🔍 분석 리포트 (눌러서 확인)", expanded=True):
                for log in debug_log: st.write(log)
            if assy_data:
                st.session_state.assy_dict = assy_data
                st.session_state.common_car = info.get('car', '')
                st.session_state.common_vol = info.get('vol', 0)
                st.success(f"✅ 총 {len(assy_data)}개의 ASSY 파일이 생성될 예정입니다!")
            else: st.error("데이터 없음 (Qt'y 헤더 또는 톤수/재질 확인 필요)")

    if st.session_state.assy_dict:
        c1, c2 = st.columns(2)
        car = c1.text_input("차종", value=st.session_state.common_car, key="m_car")
        base_vol = c2.number_input("기본 Volume", value=int(st.session_state.common_vol), key="m_vol")
        st.markdown("---")
        for name, items in st.session_state.assy_dict.items():
            with st.expander(f"📦 {name} ({len(items)} parts)"):
                for it in items: st.write(f"- {it['no']} ({it['name']})")
        if st.button("ZIP 다운로드 (ASSY별 통합 파일)", type="primary"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for name, items in st.session_state.assy_dict.items():
                    xb = generate_excel_file_stacked({"car":car, "base_vol":base_vol}, items, 2026)
                    if xb: zf.writestr(f"{name}_통합계산서.xlsx", xb)
            st.download_button("📥 ZIP 받기", zip_buffer.getvalue(), "Integrated_Cost_Set.zip", "application/zip")
