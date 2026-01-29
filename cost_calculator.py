import streamlit as st
import openpyxl
import io
import uuid
import re
import traceback
import zipfile
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill

# ============================================================================
# 1. 기초 데이터 및 설정 (사장님 기준 절대 유지)
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
# 2. 로직 함수 (1/1=1, 1/1=2 유지, A3 보존, 100칸 패딩)
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
# 3. PART LIST 파싱 함수 (매트릭스 구조 자동 분해)
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

def parse_part_list_matrix(file):
    try:
        wb = openpyxl.load_workbook(file, data_only=True)
        ws = wb.active
        header_info = extract_header_info(ws)
        all_rows = list(ws.iter_rows(values_only=True))
        
        header_row_index = -1
        # 컬럼 매핑 초기화
        col_map = {'part_no': 7, 'name': 8, 'qty_cols': [], 'mat': 22, 'ton': 23, 'cav': 24, 'L':10, 'W':11, 'H':12} 
        
        # 헤더 찾기
        for i, r in enumerate(all_rows):
            row_str = " ".join([str(x) for x in r if x]).replace(" ", "").upper()
            if "PARTNO" in row_str or "품번" in row_str:
                header_row_index = i
                row1 = r
                row2 = all_rows[i+1] if i+1 < len(all_rows) else [None]*len(r)
                
                # Qty(수량) 컬럼이 어디어디 있는지 몽땅 찾음 (J열, K열, L열...)
                for idx, cell in enumerate(row1):
                    if not cell: continue
                    c_val = str(cell).upper().replace(" ", "").replace("\n", "")
                    if "PARTNO" in c_val or "품번" in c_val: col_map['part_no'] = idx
                    elif "PARTNAME" in c_val or "품명" in c_val: col_map['name'] = idx
                    elif "MATERIAL" in c_val or "재질" in c_val: col_map['mat'] = idx
                    elif "THICK" in c_val or "두께" in c_val: col_map['thick'] = idx
                    elif "WEIGHT" in c_val or "중량" in c_val: col_map['weight'] = idx
                    if "QTY" in c_val or "수량" in c_val or "USG" in c_val:
                        if idx not in col_map['qty_cols']: col_map['qty_cols'].append(idx)
                            
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
                        if idx not in col_map['qty_cols']: col_map['qty_cols'].append(idx)
                break

        if header_row_index == -1: header_row_index = 5 

        # [핵심] 기둥(Column)별로 쪼개기
        assy_dict = {} 
        
        for q_col in col_map['qty_cols']:
            # 1. 이 기둥의 주인(ASSY 품번) 찾기
            # 해당 열에서 가장 위에 있는 '1'을 가진 품목이 대장(ASSY)이라고 가정
            assy_name = f"ASSY_Type_{q_col}" 
            for i in range(header_row_index + 1, len(all_rows)):
                r = list(all_rows[i])
                if len(r) > q_col and safe_float(r[q_col]) > 0:
                    temp_no = str(r[col_map['part_no']]).strip()
                    if temp_no and "None" not in temp_no:
                        # 파일명으로 쓸 거니까 특수문자 제거
                        assy_name = temp_no.replace("/", "_").replace("*", "")
                        break
            
            # 2. 이 기둥에 속한(1이 찍힌) 부품들 싹 긁어모으기
            items_in_assy = []
            for i in range(header_row_index + 1, len(all_rows)):
                r = list(all_rows[i])
                if len(r) < 100: r.extend([None] * (100 - len(r))) # 100칸 패딩 (안전장치)
                
                # 이 기둥(q_col)에 숫자가 없으면 내 부품 아님 -> 패스
                u_val_raw = safe_float(r[q_col]) 
                if u_val_raw <= 0: continue

                p_idx = col_map.get('part_no', 7)
                if not r[p_idx]: continue
                p_no_str = str(r[p_idx]).strip()
                clean_p_no = p_no_str.replace(" ", "").upper()
                if "PARTNO" in clean_p_no or "품번" in clean_p_no: continue
                if "비고" in clean_p_no or "REMARK" in clean_p_no: continue
                
                # 사출품인지 확인 (톤수/재질)
                t_idx = col_map.get('ton', 28)
                m_idx = col_map.get('mat', 27) 
                raw_ton = r[t_idx] if t_idx < len(r) else None
                raw_mat = r[m_idx] if m_idx < len(r) else None
                
                if not safe_float(raw_ton) and (not raw_mat or str(raw_mat).strip() == ""):
                    continue

                # 데이터 추출
                n_idx = col_map.get('name', 8)
                rem_val = str(r[n_idx + 1] if n_idx + 1 < len(r) and r[n_idx+1] else "")
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
                        if key in s_mat: mapped_mat = key; break
                    if "PP" in s_mat and mapped_mat == "무도장 TPO": mapped_mat = "PP"

                ton = int(safe_float(raw_ton, default=1300))
                
                # Cavity 1/1 -> 2
                cv_idx = col_map.get('cav', t_idx + 1)
                raw_cav = str(r[cv_idx]) if cv_idx < len(r) else "1"
                if "/" in raw_cav:
                    try: cav = int(sum(safe_float(x) for x in raw_cav.split('/') if x.strip()))
                    except: cav = int(safe_float(raw_cav, default=1))
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
                    "usage": u_val_raw, 
                    "L": l, "W": w, "H": h, "thick": t,
                    "weight": weight_val,
                    "mat": mapped_mat,
                    "ton": ton,
                    "cavity": cav,
                    "price": 2000
                }
                items_in_assy.append(item)
            
            # 3. 결과 저장 (ASSY 이름 : 부품 리스트)
            if items_in_assy:
                if assy_name in assy_dict: assy_name = f"{assy_name}_{q_col}"
                assy_dict[assy_name] = items_in_assy

        return assy_dict, header_info

    except Exception as e:
        st.error(f"분석 중 오류 발생: {e}")
        st.code(traceback.format_exc())
        return {}, {}

# ============================================================================
# 4. 엑셀 생성 함수 (집계표 + 상세시트 포함한 '통합 엑셀' 생성)
# ============================================================================
def generate_excel_file(common, items, sel_year):
    try:
        wb = openpyxl.load_workbook("template.xlsx")
        template_sheet = wb.active
        template_sheet.title = "Master_Template"
    except: return None

    # [1] 집계표(Summary) 시트 생성 (맨 앞장)
    ws_summary = wb.create_sheet("ASSY_Summary", 0)
    
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="36486b", end_color="36486b", fill_type="solid")
    align_center = Alignment(horizontal='center', vertical='center')
    thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))

    headers = ["NO", "PART NO", "PART NAME", "USAGE", "MATERIAL", "TON", "CAVITY", "WEIGHT(g)", "NOTE"]
    for col_idx, h_text in enumerate(headers, 1):
        cell = ws_summary.cell(row=1, column=col_idx, value=h_text)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = align_center
        cell.border = thin_border
    
    for idx, item in enumerate(items, 1):
        row_num = idx + 1
        data = [idx, item['no'], item['name'], item['usage'], item['mat'], item['ton'], item['cavity'], item['weight'], item['remarks']]
        for col_idx, val in enumerate(data, 1):
            cell = ws_summary.cell(row=row_num, column=col_idx, value=val)
            cell.alignment = align_center
            cell.border = thin_border
    
    ws_summary.column_dimensions['B'].width = 25
    ws_summary.column_dimensions['C'].width = 35
    ws_summary.column_dimensions['E'].width = 15

    # [2] 상세 시트 생성 (부품 하나당 시트 하나씩)
    for item in items:
        safe_title = str(item['no']).replace("/", "_").replace("*", "")[:30]
        if "비고" in safe_title or "REMARK" in safe_title: continue
        if not safe_title: safe_title = "No_Name"

        target_sheet = wb.copy_worksheet(template_sheet)
        target_sheet.title = safe_title
        ws = target_sheet

        safe_write(ws, "N3", common['car'])
        safe_write(ws, "C3", item['no'])    
        safe_write(ws, "C4", item['name']) 
        
        # [철칙] A3 셀 건드리지 않음
        
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

# ============================================================================
# 5. Streamlit UI (통합)
# ============================================================================
st.set_page_config(page_title="원가계산서(통합)", layout="wide")
st.title("원가계산서 (단품/수동 + ASSY 자동분해)")

if 'manual_items' not in st.session_state: st.session_state.manual_items = []
if 'assy_dict' not in st.session_state: st.session_state.assy_dict = {}
if 'common_car' not in st.session_state: st.session_state.common_car = ""
if 'common_vol' not in st.session_state: st.session_state.common_vol = 0
if 'excel_data' not in st.session_state: st.session_state.excel_data = None

mode = st.radio("작업 모드 선택", ["단품 계산", "ASSY(수동 입력)", "PART LIST 엑셀 업로드(Matrix)"], horizontal=True)

# [MODE 1 & 2] 단품 및 수동 입력
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
                if cols[5].button("🗑️", key=f"d_{uid}"): 
                    st.session_state.manual_items.pop(i)
                    st.rerun()

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

    if st.button("엑셀 생성 (Single File)", type="primary"):
        excel_bytes = generate_excel_file({"car":car, "base_vol":base_vol}, st.session_state.manual_items, 2026)
        if excel_bytes:
            st.download_button("📥 다운로드", excel_bytes, "Manual_Cost.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# [MODE 3] PART LIST 엑셀 업로드 (Matrix)
else:
    st.info("💡 엑셀을 올리면 기둥(Column)별로 분리 + 'ASSY 집계표'가 포함된 파일을 생성합니다.")
    
    uploaded_file = st.file_uploader("PART LIST 파일 업로드", type=["xlsx", "xls"])
    if uploaded_file:
        if st.button("🔄 분석 시작"):
            assy_data, info = parse_part_list_matrix(uploaded_file)
            if assy_data:
                st.session_state.assy_dict = assy_data
                st.session_state.common_car = info.get('car', '')
                st.session_state.common_vol = info.get('vol', 0)
                st.success(f"✅ {len(assy_data)}개 ASSY 분리 완료!")
            else:
                st.error("데이터 없음 (톤수/재질 확인)")

    if st.session_state.assy_dict:
        c1, c2 = st.columns(2)
        car = c1.text_input("차종", value=st.session_state.common_car, key="m_car")
        base_vol = c2.number_input("기본 Volume", value=int(st.session_state.common_vol), key="m_vol")
        
        st.markdown("---")
        for name, items in st.session_state.assy_dict.items():
            with st.expander(f"📦 {name} ({len(items)} items)"):
                for it in items: st.write(f"- {it['no']} ({it['name']})")
        
        if st.button("ZIP 다운로드 (All in One)", type="primary"):
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zf:
                for name, items in st.session_state.assy_dict.items():
                    xb = generate_excel_file({"car":car, "base_vol":base_vol}, items, 2026)
                    if xb: zf.writestr(f"{name}.xlsx", xb)
            st.download_button("📥 ZIP 받기", zip_buffer.getvalue(), "Cost_Set.zip", "application/zip")
