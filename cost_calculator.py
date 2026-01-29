import streamlit as st
import openpyxl
import io
import uuid
import re
import traceback
import zipfile
from openpyxl.styles import Alignment

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
# 2. 로직 함수 (안전장치 & 계산식 절대 유지)
# ============================================================================
def safe_float(value, default=0.0):
    try:
        if value is None: return default
        s_val = str(value).strip().upper()
        if not s_val: return default
        
        # 괄호, 줄바꿈 제거
        for sep in ['\n', '(', '\r']:
            if sep in s_val: s_val = s_val.split(sep)[0].strip()
        
        # '/'가 있으면 앞의 숫자만 가져옴 (U/S 1/1 -> 1 유지)
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
# 3. PART LIST 파싱 함수 (Matrix 대응: 기둥별 자동 분리)
# ============================================================================
def extract_header_info(ws):
    extracted = {"car": "", "vol": 0}
    # [유지] 범위 150행까지 넉넉하게
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
        
        # 1. 공통 정보 (차종, 볼륨)
        header_info = extract_header_info(ws)
        all_rows = list(ws.iter_rows(values_only=True))
        
        # 2. 헤더 찾기 (PART NO)
        header_row_index = -1
        col_map = {'part_no': 7, 'name': 8, 'qty_cols': [], 'mat': 22, 'ton': 23, 'cav': 24, 'L':10, 'W':11, 'H':12} 
        
        for i, r in enumerate(all_rows):
            row_str = " ".join([str(x) for x in r if x]).replace(" ", "").upper()
            if "PARTNO" in row_str or "품번" in row_str:
                header_row_index = i
                row1 = r
                row2 = all_rows[i+1] if i+1 < len(all_rows) else [None]*len(r)
                
                # 윗줄 (Qty 열 찾기)
                for idx, cell in enumerate(row1):
                    if not cell: continue
                    c_val = str(cell).upper().replace(" ", "").replace("\n", "")
                    if "PARTNO" in c_val or "품번" in c_val: col_map['part_no'] = idx
                    elif "PARTNAME" in c_val or "품명" in c_val: col_map['name'] = idx
                    elif "MATERIAL" in c_val or "재질" in c_val: col_map['mat'] = idx
                    elif "THICK" in c_val or "두께" in c_val: col_map['thick'] = idx
                    elif "WEIGHT" in c_val or "중량" in c_val: col_map['weight'] = idx
                    # [핵심] QTY 컬럼 모두 수집
                    if "QTY" in c_val or "수량" in c_val or "USG" in c_val:
                        if idx not in col_map['qty_cols']: col_map['qty_cols'].append(idx)
                            
                # 아랫줄 (Qty 열 추가 확인 및 스펙 찾기)
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

        # 3. [핵심] 각 기둥(Qty 열)별로 나누어 담기
        assy_dict = {} # {"ASSY명": [부품리스트], ...}

        for q_col in col_map['qty_cols']:
            # 3-1. 이 기둥의 주인(ASSY 이름) 찾기
            assy_name = f"ASSY_Type_{q_col}" 
            # (옵션) 해당 열 최상단에 있는 품번을 ASSY명으로 쓰기
            for i in range(header_row_index + 1, len(all_rows)):
                r = list(all_rows[i])
                if len(r) > q_col and safe_float(r[q_col]) > 0:
                    # 해당 열에 수량이 있는 첫 번째 놈의 품번을 파일명으로
                    temp_no = str(r[col_map['part_no']]).strip()
                    if temp_no and "None" not in temp_no:
                        assy_name = temp_no.replace("/", "_").replace("*", "")
                        break
            
            # 3-2. 부품 긁어모으기
            items_in_assy = []
            
            for i in range(header_row_index + 1, len(all_rows)):
                r = list(all_rows[i])
                # [유지] 100칸 패딩 (안전장치)
                if len(r) < 100: r.extend([None] * (100 - len(r)))
                
                # 해당 열(q_col)에 수량이 없으면 이 ASSY 부품 아님 -> 스킵
                u_val_raw = safe_float(r[q_col]) # 1/1 -> 1 (유지)
                if u_val_raw <= 0: continue

                # 파싱 시작
                p_idx = col_map.get('part_no', 7)
                if not r[p_idx]: continue
                p_no_str = str(r[p_idx]).strip()
                clean_p_no = p_no_str.replace(" ", "").upper()
                if "PARTNO" in clean_p_no or "품번" in clean_p_no: continue
                if "비고" in clean_p_no or "REMARK" in clean_p_no: continue
                
                n_idx = col_map.get('name', 8)
                rem_val = str(r[n_idx + 1] if n_idx + 1 < len(r) and r[n_idx+1] else "")
                
                # 사출품 여부 판단
                t_idx = col_map.get('ton', 28)
                m_idx = col_map.get('mat', 27) 
                raw_ton = r[t_idx] if t_idx < len(r) else None
                raw_mat = r[m_idx] if m_idx < len(r) else None
                
                if not safe_float(raw_ton) and (not raw_mat or str(raw_mat).strip() == ""):
                    continue

                # 데이터 추출 (기존 로직 100% 동일)
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
                
                # [유지] Cavity 1/1 -> 2 로직
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
                    "usage": u_val_raw, # 해당 기둥의 수량 사용
                    "L": l, "W": w, "H": h, "thick": t,
                    "weight": weight_val,
                    "mat": mapped_mat,
                    "ton": ton,
                    "cavity": cav,
                    "price": 2000
                }
                items_in_assy.append(item)
            
            # 해당 ASSY에 부품이 있으면 저장
            if items_in_assy:
                # 중복 이름 방지
                if assy_name in assy_dict: assy_name = f"{assy_name}_{q_col}"
                assy_dict[assy_name] = items_in_assy

        return assy_dict, header_info

    except Exception as e:
        st.error(f"분석 중 오류 발생: {e}")
        st.code(traceback.format_exc())
        return {}, {}

# ============================================================================
# 4. 엑셀 생성 함수 (단일 파일 생성용 - 내부 로직 완전 동일)
# ============================================================================
def create_excel_bytes(common, items, sel_year):
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
        
        # [유지] A3 절대 건드리지 않음
        
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
# 5. Streamlit UI (ZIP 다운로드 기능 추가)
# ============================================================================
st.set_page_config(page_title="원가계산서(Matrix)", layout="wide")
st.title("원가계산서 (다중 ASSY 자동 분할)")
st.warning("⚠️ PART LIST 엑셀을 올리면, 기둥(Column)별로 ASSY를 자동 인식하여 분리합니다.")

if 'assy_dict' not in st.session_state: st.session_state.assy_dict = {}
if 'common_car' not in st.session_state: st.session_state.common_car = ""
if 'common_vol' not in st.session_state: st.session_state.common_vol = 0

uploaded_file = st.file_uploader("PART LIST 엑셀 파일(.xlsx)을 올려주세요.", type=["xlsx", "xls"])

if uploaded_file:
    if st.button("🔄 데이터 불러오기", type="primary"):
        with st.spinner("엑셀 분석 및 ASSY 분리 중..."):
            assy_data, info = parse_part_list_matrix(uploaded_file)
            
            if assy_data:
                st.session_state.assy_dict = assy_data
                if info.get('car'): st.session_state.common_car = info['car']
                if info.get('vol'): st.session_state.common_vol = info['vol']
                st.success(f"✅ 총 {len(assy_data)}개의 ASSY를 찾아냈습니다!")
            else:
                st.error("데이터를 찾을 수 없습니다.")

st.markdown("---")

if st.session_state.assy_dict:
    c1, c2 = st.columns(2)
    car = c1.text_input("차종", value=st.session_state.common_car)
    base_vol = c2.number_input("기본 Volume (대)", value=int(st.session_state.common_vol))
    
    st.markdown("### 📋 감지된 ASSY 목록")
    for name, items in st.session_state.assy_dict.items():
        with st.expander(f"📦 {name} (부품 {len(items)}개)"):
            for it in items:
                st.write(f"- {it['no']} : {it['name']} (Qty:{it['usage']}, C/V:{it['cavity']}, Ton:{it['ton']})")

    st.markdown("---")
    st.markdown("### 💰 엑셀 일괄 생성")
    
    if st.button("모든 ASSY 계산서 ZIP으로 다운로드", type="primary", use_container_width=True):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for assy_name, items in st.session_state.assy_dict.items():
                excel_bytes = create_excel_bytes({"car":car, "base_vol":base_vol}, items, 2026)
                if excel_bytes:
                    zf.writestr(f"{assy_name}_원가계산서.xlsx", excel_bytes)
        
        st.download_button(
            label="📥 ZIP 파일 다운로드 (Click)",
            data=zip_buffer.getvalue(),
            file_name=f"{car}_원가계산서_모음.zip",
            mime="application/zip",
            use_container_width=True
        )
