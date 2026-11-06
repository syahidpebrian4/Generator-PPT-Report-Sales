import streamlit as st
import openpyxl
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
import io
import datetime
import os

st.set_page_config(page_title="Generator PPT Sales Report", layout="centered")

# Inisialisasi Session State agar file download tidak hilang saat diklik
if "generated_ppts" not in st.session_state:
    st.session_state["generated_ppts"] = {}

# Custom Styling CSS UI Card Putih
st.markdown("""
    <style>
    .main {
        background-color: #f7f9fa;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
        max-width: 850px;
    }
    .card-container {
        background-color: #ffffff;
        padding: 30px;
        border-radius: 10px;
        box-shadow: 0px 4px 12px rgba(0, 0, 0, 0.05);
        border: 1px solid #e1e4e8;
        margin-bottom: 20px;
    }
    .main-title {
        color: #d9251d;
        font-weight: 700;
        text-align: center;
        margin-bottom: 25px;
        font-size: 26px;
    }
    .section-header {
        background-color: #d9251d;
        color: white;
        padding: 8px 15px;
        font-weight: bold;
        border-radius: 4px;
        margin-top: 15px;
        margin-bottom: 15px;
        font-size: 14px;
        letter-spacing: 0.5px;
    }
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="card-container">', unsafe_allow_html=True)
st.markdown('<div class="main-title">Generator PPT Sales Report</div>', unsafe_allow_html=True)

# SECTION 1: PERIODE TANGGAL REPORT
st.markdown('<div class="section-header">1. PERIODE REPORT</div>', unsafe_allow_html=True)

col_date1, col_date2 = st.columns(2)
with col_date1:
    date_last_week = st.date_input("Tanggal Last Week *", datetime.date.today() - datetime.timedelta(days=7))
with col_date2:
    date_this_week = st.date_input("Tanggal This Week *", datetime.date.today())

str_date_lw = date_last_week.strftime("%b %d")
str_date_tw = date_this_week.strftime("%b %d")

# Tanggal khusus Slide Cover (Ukuran Font 20pt)
str_cover_date = date_this_week.strftime("%d %B %Y")

# Function mendapatkan suffix tanggal (1st, 2nd, 3rd, 4th, dst)
def get_day_with_suffix(d):
    day = d.day
    if 11 <= day <= 13:
        return f"{day}th"
    else:
        suffixes = {1: "st", 2: "nd", 3: "rd"}
        return f"{day}{suffixes.get(day % 10, 'th')}"

str_mtd_header = f"Performance by Customer y26 {date_this_week.strftime('%b')} MTD {get_day_with_suffix(date_this_week)}"

# SECTION 2: UPLOAD FILE EXCEL
st.markdown('<div class="section-header">2. UPLOAD FILE SOURCE EXCEL (4 FILE)</div>', unsafe_allow_html=True)

col_up1, col_up2 = st.columns(2)
with col_up1:
    excel_file_sales_lw = st.file_uploader("Sales LAST WEEK (.xlsx)", type=["xlsx", "xls"])
with col_up2:
    excel_file_sales_tw = st.file_uploader("Sales THIS WEEK (.xlsx)", type=["xlsx", "xls"])

col_up3, col_up4 = st.columns(2)
with col_up3:
    excel_file_enduser = st.file_uploader("Sales Value (.xlsx)", type=["xlsx", "xls"])
with col_up4:
    excel_file_top5 = st.file_uploader("TOP 50 GAP By Prod (.xlsx)", type=["xlsx", "xls"])

st.markdown('</div>', unsafe_allow_html=True)

TEMPLATE_PPT_PATH = "TEMPLATE.pptx"

REGION_MAP = {
    "Region 1": [
        {"code": "06001", "name": "Pasar Rebo"}, {"code": "06003", "name": "Kelapa Gading"},
        {"code": "06006", "name": "Ciputat"}, {"code": "06007", "name": "Alam Sutera"},
        {"code": "06010", "name": "Medan"}, {"code": "06014", "name": "Palembang"},
        {"code": "06015", "name": "Pekanbaru"}, {"code": "06021", "name": "Jatake"},
        {"code": "06022", "name": "Serang"}, {"code": "06029", "name": "Batam"},
        {"code": "06031", "name": "Lampung"}, {"code": "06039", "name": "Serpong"}
    ],
    "Region 2": [
        {"code": "06004", "name": "Meruya"}, {"code": "06005", "name": "Bandung"},
        {"code": "06008", "name": "Cibitung"}, {"code": "06018", "name": "Bekasi"},
        {"code": "06023", "name": "Cikarang"}, {"code": "06024", "name": "Cirebon"},
        {"code": "06026", "name": "Bogor"}, {"code": "06027", "name": "Tasikmalaya"},
        {"code": "06030", "name": "Pakansari"}, {"code": "06034", "name": "Kerawang"},
        {"code": "06036", "name": "Cimahi"}, {"code": "06038", "name": "Tegal"}
    ],
    "Region 3": [
        {"code": "06002", "name": "Sidoarjo"}, {"code": "06009", "name": "Denpasar"},
        {"code": "06011", "name": "Semarang"}, {"code": "06013", "name": "Makasar"},
        {"code": "06016", "name": "Yogyakarta"}, {"code": "06017", "name": "Banjarmasin"},
        {"code": "06019", "name": "Solo"}, {"code": "06020", "name": "Balikpapan"},
        {"code": "06028", "name": "Mastrip"}, {"code": "06032", "name": "Samarinda"},
        {"code": "06033", "name": "Manado"}, {"code": "06037", "name": "Mataram"}
    ]
}

CAT_MAPPING = {
    "Fresh Food": ["31", "32", "33", "34", "35"],
    "Meal Solution": ["80", "82"],
    "Dry Food": ["17", "21", "11", "26", "27", "14", "19"],
    "Non Food Int": ["86", "87", "88", "76", "77", "64", "68", "51", "57", "58", "85", "13", "62", "71", "60", "61", "63", "67"]
}

# FORMAT ANGKA: KOMA PAKE TITIK (.), RIBUAN PAKE KOMA (,)
def format_cell_value(val, is_sales=False):
    if val is None or str(val).strip() == "":
        return ""
    elif isinstance(val, (int, float)):
        is_negative = val < 0
        abs_val = abs(val)
        if is_sales:
            formatted_num = f"{round(abs_val):,}"
        else:
            if abs_val < 100 and "." in str(val):
                formatted_num = f"{abs_val:.1f}"
            else:
                formatted_num = f"{abs_val:,.1f}"
                if formatted_num.endswith(".0"):
                    formatted_num = formatted_num[:-2]
        return f"▲{formatted_num}" if is_negative else formatted_num
    else:
        val_str = str(val).strip()
        if val_str.startswith("-"):
            val_str = "▲" + val_str[1:]
        return val_str

def format_integer_value(val):
    if val is None or str(val).strip() == "":
        return ""
    try:
        num = round(float(val))
        return f"{num:,}"
    except Exception:
        return str(val).strip()

def extract_store_sales_raw(ws, store_code):
    clean_target = str(int(store_code))
    found_row = -1
    
    for row_idx in range(1, ws.max_row + 1):
        cell_val = str(ws.cell(row=row_idx, column=2).value or "")
        if clean_target in cell_val:
            found_row = row_idx
            break
            
    if found_row == -1:
        for row_idx in range(1, ws.max_row + 1):
            row_str = " ".join([str(ws.cell(row=row_idx, column=c).value or "") for c in range(1, 10)])
            if clean_target in row_str:
                found_row = row_idx
                break

    if found_row == -1:
        return None, None

    start_row = found_row + 4
    target_cols = [4, 5, 6, 7, 10, 13, 16, 22, 25, 28, 31]
    
    formatted_data = []
    numeric_data = []
    
    for r in range(7):
        curr_row = start_row + r
        row_fmt = []
        row_num = []
        for idx, c in enumerate(target_cols):
            val = ws.cell(row=curr_row, column=c).value
            row_fmt.append(format_cell_value(val, is_sales=(idx == 0)))
            try:
                row_num.append(float(val) if val is not None else 0.0)
            except Exception:
                row_num.append(0.0)
        formatted_data.append(row_fmt)
        numeric_data.append(row_num)
        
    return formatted_data, numeric_data

def extract_store_data_rowdata(df_raw, store_code):
    clean_code_5 = str(store_code).zfill(5)
    clean_code_num = str(int(store_code))
    
    col_a = df_raw.iloc[:, 0].astype(str).str.strip().str.zfill(5)
    col_b = df_raw.iloc[:, 1].astype(str).str.strip().str.split('.').str[0].str.zfill(2)
    col_aw = pd.to_numeric(df_raw.iloc[:, 48], errors='coerce').fillna(0)
    col_ax = pd.to_numeric(df_raw.iloc[:, 49], errors='coerce').fillna(0)
    
    store_mask = (col_a == clean_code_5) | (col_a == clean_code_num.zfill(5))
    
    if not store_mask.any():
        return None

    enduser_rows = []
    total_act = 0
    total_ly = 0

    for div_name, cat_list in CAT_MAPPING.items():
        cat_list_formatted = [c.zfill(2) for c in cat_list]
        full_mask = store_mask & (col_b.isin(cat_list_formatted))
        
        act_val = col_aw[full_mask].sum() / 1_000_000.0
        ly_val = col_ax[full_mask].sum() / 1_000_000.0
        
        gr_val = ((act_val - ly_val) / ly_val * 100) if ly_val != 0 else 0
        gap_val = act_val - ly_val
        
        total_act += act_val
        total_ly += ly_val
        
        enduser_rows.append([
            format_integer_value(act_val),
            format_integer_value(ly_val),
            format_cell_value(gr_val),
            format_cell_value(gap_val)
        ])

    total_gr = ((total_act - total_ly) / total_ly * 100) if total_ly != 0 else 0
    total_gap = total_act - total_ly
    
    enduser_rows.append([
        format_integer_value(total_act),
        format_integer_value(total_ly),
        format_cell_value(total_gr),
        format_cell_value(total_gap)
    ])

    return enduser_rows

def get_top5_products_str(df_top5, store_code, cust_grp_code):
    if df_top5 is None:
        return ""
        
    clean_code_5 = str(store_code).zfill(5)
    clean_code_num = str(int(store_code))
    cust_code_str = str(cust_grp_code).strip()
    
    col_str = df_top5.iloc[:, 0].astype(str).str.strip()
    col_grp = df_top5.iloc[:, 6].astype(str).str.strip().str.split('.').str[0]
    
    mask = ((col_str == clean_code_5) | (col_str == clean_code_num)) & (col_grp == cust_code_str)
    df_filtered = df_top5[mask].copy()
    
    if df_filtered.empty:
        return ""
        
    prod_names = df_filtered.iloc[:, 2].astype(str).values
    gap_vals = df_filtered.iloc[:, 16].values

    top_items = []
    for p_name, g_val in zip(prod_names[:5], gap_vals[:5]):
        p_clean = p_name.strip()
        if g_val is None or str(g_val).strip() == "":
            top_items.append(p_clean)
            continue
            
        try:
            num = float(str(g_val).replace(",", "").replace("▲", "-"))
            if abs(num) >= 100000:
                num_mio = num / 1_000_000.0
            else:
                num_mio = num
                
            abs_mio = abs(num_mio)
            formatted_num = f"{abs_mio:,.0f}"
            sign = "▲" if num < 0 else ""
            top_items.append(f"{p_clean} {sign}{formatted_num} mio")
        except Exception:
            top_items.append(f"{p_clean} {g_val} mio")
        
    return ", ".join(top_items)

def set_cell_text_center(cell, text, font_size=12):
    cell.text = str(text)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in cell.text_frame.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(font_size)
        for r in p.runs:
            r.font.size = Pt(font_size)

def set_header_text_formatted(shape, text_content, font_size=12):
    shape.text_frame.text = text_content
    for p in shape.text_frame.paragraphs:
        p.font.size = Pt(font_size)
        p.font.bold = True
        for r in p.runs:
            r.font.size = Pt(font_size)
            r.font.bold = True

def set_store_name_text(shape, store_name):
    tf = shape.text_frame
    tf.text = f"■ Store {store_name}"
    for p in tf.paragraphs:
        p.font.name = "Aptos"
        p.font.size = Pt(18)
        p.font.bold = True
        for r in p.runs:
            r.font.name = "Aptos"
            r.font.size = Pt(18)
            r.font.bold = True

def set_region_slide_title(shape, text_content):
    tf = shape.text_frame
    tf.text = text_content
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    for p in tf.paragraphs:
        p.alignment = PP_ALIGN.CENTER
        p.font.size = Pt(36)
        p.font.bold = True
        for r in p.runs:
            r.font.size = Pt(36)
            r.font.bold = True

def update_cover_date(prs, date_str):
    slide_cover = prs.slides[0]
    for shape in slide_cover.shapes:
        if shape.has_text_frame:
            text = shape.text_frame.text.strip()
            if any(month in text for month in ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December", "202"]) or text.isdigit():
                shape.text_frame.text = date_str
                for p in shape.text_frame.paragraphs:
                    p.alignment = PP_ALIGN.CENTER
                    p.font.size = Pt(20)
                    for r in p.runs:
                        r.font.size = Pt(20)

def add_analysis_paragraph(tf, header_title, gap_val, gr_val, top5_str):
    if len(tf.paragraphs) > 0 and len(tf.paragraphs[0].text.strip()) == 0:
        p_head = tf.paragraphs[0]
    else:
        p_head = tf.add_paragraph()
        
    p_head.text = f"• {header_title} {gap_val} Mio vs LY, G/R {gr_val}%"
    p_head.line_spacing = 1.0
    p_head.space_before = Pt(2)
    p_head.space_after = Pt(1)
    p_head.font.size = Pt(10)
    p_head.font.bold = True
    p_head.font.color.rgb = RGBColor(0, 0, 0)
    
    if top5_str:
        p_body = tf.add_paragraph()
        p_body.text = top5_str
        p_body.line_spacing = 1.0
        p_body.space_before = Pt(0)
        p_body.space_after = Pt(6)
        p_body.font.size = Pt(10)
        p_body.font.bold = False
        p_body.font.color.rgb = RGBColor(0, 0, 0)

# TOMBOL GENERATE
if st.button("Generate File PPT", type="primary", use_container_width=True):
    if not excel_file_sales_lw or not excel_file_sales_tw or not excel_file_enduser or not excel_file_top5:
        st.error("Harap upload KEEMPAT file Excel terlebih dahulu!")
    elif not os.path.exists(TEMPLATE_PPT_PATH):
        st.error(f"File template '{TEMPLATE_PPT_PATH}' tidak ditemukan di folder project!")
    else:
        with st.spinner("Sedang mengolah ke-4 data mentah & membuat file PPT..."):
            wb_sales_lw = openpyxl.load_workbook(excel_file_sales_lw, data_only=True)
            ws_sales_lw = wb_sales_lw["PPT"] if "PPT" in wb_sales_lw.sheetnames else wb_sales_lw.active
            
            wb_sales_tw = openpyxl.load_workbook(excel_file_sales_tw, data_only=True)
            ws_sales_tw = wb_sales_tw["PPT"] if "PPT" in wb_sales_tw.sheetnames else wb_sales_tw.active
            
            try:
                df_rowdata = pd.read_excel(excel_file_enduser, sheet_name="Rowdata", header=None)
            except Exception:
                df_rowdata = pd.read_excel(excel_file_enduser, sheet_name=0, header=None)

            try:
                df_top5 = pd.read_excel(excel_file_top5, sheet_name="By_Prod")
            except Exception:
                df_top5 = pd.read_excel(excel_file_top5, sheet_name=0)
            
            st.session_state["generated_ppts"] = {}
            blank_layout = Presentation(TEMPLATE_PPT_PATH).slide_layouts[6]

            # 1. PROCESS PPT PER REGION (REG 1, REG 2, REG 3)
            for reg in ["Region 1", "Region 2", "Region 3"]:
                prs = Presentation(TEMPLATE_PPT_PATH)
                update_cover_date(prs, str_cover_date)

                if len(prs.slides) >= 2:
                    for shape in prs.slides[1].shapes:
                        if shape.has_text_frame and "Region" in shape.text:
                            set_region_slide_title(shape, reg)

                stores = REGION_MAP[reg]
                for store in stores:
                    sales_fmt_lw, sales_num_lw = extract_store_sales_raw(ws_sales_lw, store["code"])
                    sales_fmt_tw, sales_num_tw = extract_store_sales_raw(ws_sales_tw, store["code"])
                    enduser_data = extract_store_data_rowdata(df_rowdata, store["code"])
                    
                    if sales_fmt_tw and sales_fmt_lw:
                        # Slide Perbandingan
                        temp_prs_cmp = Presentation(TEMPLATE_PPT_PATH)
                        slide_cmp = temp_prs_cmp.slides[2]
                        for shape in slide_cmp.shapes:
                            if shape.has_text_frame:
                                text_content = shape.text_frame.text
                                if "Store" in text_content or "PASAR REBO" in text_content:
                                    set_store_name_text(shape, store['name'])
                                elif "Performance by Customer" in text_content or "MTD" in text_content:
                                    set_header_text_formatted(shape, f"■ {str_mtd_header} (This Week vs Last Week)", font_size=12)

                        for shape in slide_cmp.shapes:
                            if shape.has_table:
                                table = shape.table
                                if len(table.columns) >= 13:
                                    for grp_idx in range(7):
                                        base_row = (grp_idx * 3) + 2
                                        if base_row < len(table.rows):
                                            set_cell_text_center(table.cell(base_row, 2), str_date_lw, font_size=12)
                                            for c_idx, val in enumerate(sales_fmt_lw[grp_idx]):
                                                set_cell_text_center(table.cell(base_row, c_idx + 3), str(val), font_size=12)
                                        if base_row + 1 < len(table.rows):
                                            set_cell_text_center(table.cell(base_row + 1, 2), str_date_tw, font_size=12)
                                            for c_idx, val in enumerate(sales_fmt_tw[grp_idx]):
                                                set_cell_text_center(table.cell(base_row + 1, c_idx + 3), str(val), font_size=12)
                                        if base_row + 2 < len(table.rows):
                                            for c_idx in range(11):
                                                val_tw = sales_num_tw[grp_idx][c_idx]
                                                val_lw = sales_num_lw[grp_idx][c_idx]
                                                gap_calc = val_tw - val_lw
                                                set_cell_text_center(table.cell(base_row + 2, c_idx + 3), format_cell_value(gap_calc, is_sales=(c_idx == 0)), font_size=12)

                        new_slide_cmp = prs.slides.add_slide(blank_layout)
                        for shape in slide_cmp.shapes:
                            new_slide_cmp.shapes._spTree.append(shape.element)

                        # Slide Detail Store
                        temp_prs_dtl = Presentation(TEMPLATE_PPT_PATH)
                        slide_dtl = temp_prs_dtl.slides[3]
                        for shape in slide_dtl.shapes:
                            if shape.has_text_frame:
                                text_content = shape.text_frame.text
                                if "Store" in text_content or "PASAR REBO" in text_content:
                                    set_store_name_text(shape, store['name'])
                                elif "Performance by Customer" in text_content or "MTD" in text_content:
                                    set_header_text_formatted(shape, f"■ {str_mtd_header}", font_size=12)

                        for shape in slide_dtl.shapes:
                            if shape.has_table:
                                table = shape.table
                                if len(table.columns) >= 10:
                                    for r_idx, row_data in enumerate(sales_fmt_tw):
                                        for c_idx, val in enumerate(row_data):
                                            row_target = r_idx + 2
                                            col_target = c_idx + 2
                                            if row_target < len(table.rows) and col_target < len(table.columns):
                                                set_cell_text_center(table.cell(row_target, col_target), str(val), font_size=12)
                                elif len(table.columns) < 10 and enduser_data:
                                    for r_idx, row_data in enumerate(enduser_data):
                                        for c_idx, val in enumerate(row_data):
                                            row_target = r_idx + 2
                                            col_target = c_idx + 1
                                            if row_target < len(table.rows) and col_target < len(table.columns):
                                                set_cell_text_center(table.cell(row_target, col_target), str(val), font_size=8)

                        top5_retail = get_top5_products_str(df_top5, store["code"], 20)
                        top5_horeka = get_top5_products_str(df_top5, store["code"], 30)
                        top5_service = get_top5_products_str(df_top5, store["code"], 40)
                        top5_enduser = get_top5_products_str(df_top5, store["code"], 50)

                        gap_retail = sales_fmt_tw[0][2] if len(sales_fmt_tw) > 0 else ""
                        gr_retail = sales_fmt_tw[0][3] if len(sales_fmt_tw) > 0 else ""
                        gap_horeka = sales_fmt_tw[1][2] if len(sales_fmt_tw) > 1 else ""
                        gr_horeka = sales_fmt_tw[1][3] if len(sales_fmt_tw) > 1 else ""
                        gap_service = sales_fmt_tw[2][2] if len(sales_fmt_tw) > 2 else ""
                        gr_service = sales_fmt_tw[2][3] if len(sales_fmt_tw) > 2 else ""
                        gap_enduser = sales_fmt_tw[4][2] if len(sales_fmt_tw) > 4 else ""
                        gr_enduser = sales_fmt_tw[4][3] if len(sales_fmt_tw) > 4 else ""

                        for shape in slide_dtl.shapes:
                            if shape.has_text_frame:
                                tf = shape.text_frame
                                text_full = tf.text
                                if "GAP Retail" in text_full or "GAP Horeka" in text_full or "GAP Service" in text_full:
                                    tf.clear()
                                    add_analysis_paragraph(tf, "GAP Retail", gap_retail, gr_retail, top5_retail)
                                    add_analysis_paragraph(tf, "GAP Horeka", gap_horeka, gr_horeka, top5_horeka)
                                    add_analysis_paragraph(tf, "GAP Service", gap_service, gr_service, top5_service)
                                elif "Sales" in text_full and ("Mio vs LY" in text_full or "(angka)" in text_full):
                                    tf.clear()
                                    add_analysis_paragraph(tf, "Sales", gap_enduser, gr_enduser, top5_enduser)

                        new_slide_dtl = prs.slides.add_slide(blank_layout)
                        for shape in slide_dtl.shapes:
                            new_slide_dtl.shapes._spTree.append(shape.element)

                for _ in range(2):
                    rId = prs.slides._sldIdLst[2].rId
                    prs.part.drop_rel(rId)
                    del prs.slides._sldIdLst[2]

                output_ppt = io.BytesIO()
                prs.save(output_ppt)
                output_ppt.seek(0)
                
                st.session_state["generated_ppts"][reg] = {
                    "bytes": output_ppt.getvalue(),
                    "filename": f"Report_{reg}_{datetime.date.today().strftime('%Y%m%d')}.pptx"
                }

            # 2. PROCESS FILE PPT GABUNGAN ALL REGION
            prs_all = Presentation(TEMPLATE_PPT_PATH)
            update_cover_date(prs_all, str_cover_date)

            for reg in ["Region 1", "Region 2", "Region 3"]:
                # Slide Pemisah Region (36pt Rata Tengah)
                temp_prs_reg = Presentation(TEMPLATE_PPT_PATH)
                slide_reg = temp_prs_reg.slides[1]
                for shape in slide_reg.shapes:
                    if shape.has_text_frame and "Region" in shape.text:
                        set_region_slide_title(shape, reg)
                
                new_slide_reg = prs_all.slides.add_slide(blank_layout)
                for shape in slide_reg.shapes:
                    new_slide_reg.shapes._spTree.append(shape.element)
                
                # Slide Toko-Toko
                stores = REGION_MAP[reg]
                for store in stores:
                    sales_fmt_lw, sales_num_lw = extract_store_sales_raw(ws_sales_lw, store["code"])
                    sales_fmt_tw, sales_num_tw = extract_store_sales_raw(ws_sales_tw, store["code"])
                    enduser_data = extract_store_data_rowdata(df_rowdata, store["code"])
                    
                    if sales_fmt_tw and sales_fmt_lw:
                        # Slide Perbandingan
                        temp_prs_cmp = Presentation(TEMPLATE_PPT_PATH)
                        slide_cmp = temp_prs_cmp.slides[2]
                        for shape in slide_cmp.shapes:
                            if shape.has_text_frame:
                                text_content = shape.text_frame.text
                                if "Store" in text_content or "PASAR REBO" in text_content:
                                    set_store_name_text(shape, store['name'])
                                elif "Performance by Customer" in text_content or "MTD" in text_content:
                                    set_header_text_formatted(shape, f"■ {str_mtd_header} (This Week vs Last Week)", font_size=12)

                        for shape in slide_cmp.shapes:
                            if shape.has_table:
                                table = shape.table
                                if len(table.columns) >= 13:
                                    for grp_idx in range(7):
                                        base_row = (grp_idx * 3) + 2
                                        if base_row < len(table.rows):
                                            set_cell_text_center(table.cell(base_row, 2), str_date_lw, font_size=12)
                                            for c_idx, val in enumerate(sales_fmt_lw[grp_idx]):
                                                set_cell_text_center(table.cell(base_row, c_idx + 3), str(val), font_size=12)
                                        if base_row + 1 < len(table.rows):
                                            set_cell_text_center(table.cell(base_row + 1, 2), str_date_tw, font_size=12)
                                            for c_idx, val in enumerate(sales_fmt_tw[grp_idx]):
                                                set_cell_text_center(table.cell(base_row + 1, c_idx + 3), str(val), font_size=12)
                                        if base_row + 2 < len(table.rows):
                                            for c_idx in range(11):
                                                val_tw = sales_num_tw[grp_idx][c_idx]
                                                val_lw = sales_num_lw[grp_idx][c_idx]
                                                gap_calc = val_tw - val_lw
                                                set_cell_text_center(table.cell(base_row + 2, c_idx + 3), format_cell_value(gap_calc, is_sales=(c_idx == 0)), font_size=12)

                        new_slide_cmp = prs_all.slides.add_slide(blank_layout)
                        for shape in slide_cmp.shapes:
                            new_slide_cmp.shapes._spTree.append(shape.element)

                        # Slide Detail Store
                        temp_prs_dtl = Presentation(TEMPLATE_PPT_PATH)
                        slide_dtl = temp_prs_dtl.slides[3]
                        for shape in slide_dtl.shapes:
                            if shape.has_text_frame:
                                text_content = shape.text_frame.text
                                if "Store" in text_content or "PASAR REBO" in text_content:
                                    set_store_name_text(shape, store['name'])
                                elif "Performance by Customer" in text_content or "MTD" in text_content:
                                    set_header_text_formatted(shape, f"■ {str_mtd_header}", font_size=12)

                        for shape in slide_dtl.shapes:
                            if shape.has_table:
                                table = shape.table
                                if len(table.columns) >= 10:
                                    for r_idx, row_data in enumerate(sales_fmt_tw):
                                        for c_idx, val in enumerate(row_data):
                                            row_target = r_idx + 2
                                            col_target = c_idx + 2
                                            if row_target < len(table.rows) and col_target < len(table.columns):
                                                set_cell_text_center(table.cell(row_target, col_target), str(val), font_size=12)
                                elif len(table.columns) < 10 and enduser_data:
                                    for r_idx, row_data in enumerate(enduser_data):
                                        for c_idx, val in enumerate(row_data):
                                            row_target = r_idx + 2
                                            col_target = c_idx + 1
                                            if row_target < len(table.rows) and col_target < len(table.columns):
                                                set_cell_text_center(table.cell(row_target, col_target), str(val), font_size=8)

                        top5_retail = get_top5_products_str(df_top5, store["code"], 20)
                        top5_horeka = get_top5_products_str(df_top5, store["code"], 30)
                        top5_service = get_top5_products_str(df_top5, store["code"], 40)
                        top5_enduser = get_top5_products_str(df_top5, store["code"], 50)

                        gap_retail = sales_fmt_tw[0][2] if len(sales_fmt_tw) > 0 else ""
                        gr_retail = sales_fmt_tw[0][3] if len(sales_fmt_tw) > 0 else ""
                        gap_horeka = sales_fmt_tw[1][2] if len(sales_fmt_tw) > 1 else ""
                        gr_horeka = sales_fmt_tw[1][3] if len(sales_fmt_tw) > 1 else ""
                        gap_service = sales_fmt_tw[2][2] if len(sales_fmt_tw) > 2 else ""
                        gr_service = sales_fmt_tw[2][3] if len(sales_fmt_tw) > 2 else ""
                        gap_enduser = sales_fmt_tw[4][2] if len(sales_fmt_tw) > 4 else ""
                        gr_enduser = sales_fmt_tw[4][3] if len(sales_fmt_tw) > 4 else ""

                        for shape in slide_dtl.shapes:
                            if shape.has_text_frame:
                                tf = shape.text_frame
                                text_full = tf.text
                                if "GAP Retail" in text_full or "GAP Horeka" in text_full or "GAP Service" in text_full:
                                    tf.clear()
                                    add_analysis_paragraph(tf, "GAP Retail", gap_retail, gr_retail, top5_retail)
                                    add_analysis_paragraph(tf, "GAP Horeka", gap_horeka, gr_horeka, top5_horeka)
                                    add_analysis_paragraph(tf, "GAP Service", gap_service, gr_service, top5_service)
                                elif "Sales" in text_full and ("Mio vs LY" in text_full or "(angka)" in text_full):
                                    tf.clear()
                                    add_analysis_paragraph(tf, "Sales", gap_enduser, gr_enduser, top5_enduser)

                        new_slide_dtl = prs_all.slides.add_slide(blank_layout)
                        for shape in slide_dtl.shapes:
                            new_slide_dtl.shapes._spTree.append(shape.element)

            # Hapus slide template awal (Index 1, 2, 3)
            for _ in range(3):
                rId = prs_all.slides._sldIdLst[1].rId
                prs_all.part.drop_rel(rId)
                del prs_all.slides._sldIdLst[1]

            output_ppt = io.BytesIO()
            prs_all.save(output_ppt)
            output_ppt.seek(0)
            
            st.session_state["generated_ppts"] = {
                "ALL Region (Gabungan)": {
                    "bytes": output_ppt.getvalue(),
                    "filename": f"Report_ALL_REGIONS_{datetime.date.today().strftime('%Y%m%d')}.pptx"
                },
                **st.session_state["generated_ppts"]
            }

# MENAMPILKAN TOMBOL DOWNLOAD KETIKA SELESAI GENERATE
if st.session_state["generated_ppts"]:
    st.success("✅ Generate Berhasil! Silakan unduh file:")
    for reg_name, ppt_info in st.session_state["generated_ppts"].items():
        st.download_button(
            label=f"📥 Download PPTX {reg_name}",
            data=ppt_info["bytes"],
            file_name=ppt_info["filename"],
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            key=f"dl_btn_{reg_name}"
        )
