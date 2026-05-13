import streamlit as st
import pandas as pd
import numpy as np
import os
import io
from datetime import datetime
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# ==========================================
# 1. Page Config & Smart Colors
# ==========================================
st.set_page_config(page_title="Pipes Pricing Tool", layout="wide", page_icon="🏗️")

st.markdown("""
<style>
    .main-container { background-color: #0077b5; padding: 2rem; border-radius: 15px; margin-bottom: 2rem; }
    .main-header { font-size: 3.5rem; color: #ffffff !important; font-weight: 900; margin-bottom: 0.2rem; font-family: sans-serif; }
    .sub-header { font-size: 1.4rem; color: rgba(255, 255, 255, 0.9) !important; font-weight: 600; text-transform: uppercase; }
    .stNumberInput label p, .stTextInput label p, .stSelectbox label p, .stRadio label p { font-size: 1.3rem !important; font-weight: 800 !important; color: var(--text-color) !important; }
    .stRadio div[role='radiogroup'] label div p { font-size: 1.4rem !important; font-weight: bold !important; color: var(--text-color) !important; }
    .stButton>button { width: 100%; border-radius: 8px; font-weight: bold; height: 3.5em; font-size: 1.2rem; background-color: #0077b5; color: white; border: none; }
    .stButton>button:hover { background-color: #005f91; }
    @media print {
        .main-container { background-color: white !important; color: black !important; }
        .main-header { color: black !important; }
        .sub-header { color: #333 !important; }
        .stButton, footer, header { display: none !important; }
    }
</style>
<div class="main-container">
    <div class="main-header">HDPE & uPVC Pipe Pricing Tool</div>
    <div class="sub-header">CREATED BY Eng. Ahmed Sabra</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. Session State
# ==========================================
if 'quote_list' not in st.session_state:
    st.session_state.quote_list = []

# ==========================================
# 3. Sidebar
# ==========================================
st.sidebar.header("⚙️ Settings")
material_type = st.sidebar.radio("Select Material:", ("HDPE", "uPVC"), index=0)
st.sidebar.markdown("---")
st.sidebar.info("**Eng. Ahmed Sabra**\n\n📞 +201148777463")

# ==========================================
# 4. Data Loader (BUG FIX: DataType Handling)
# ==========================================
data_file = 'data.xlsx'

@st.cache_data(ttl=30)
def load_data(file_path, sheet_name):
    try:
        xl = pd.ExcelFile(file_path)
        sheet_map = {str(name).strip().upper(): name for name in xl.sheet_names}
        target_sheet = sheet_map.get(str(sheet_name).strip().upper())
        
        if target_sheet:
            df = pd.read_excel(file_path, sheet_name=target_sheet)
            df.columns = [str(c).strip() for c in df.columns]
            
            # 1. حماية القطر والوزن وتأكيد إنهم أرقام
            if 'Diameter' in df.columns:
                df['Diameter'] = pd.to_numeric(df['Diameter'], errors='coerce').fillna(0)
            if 'Weight' in df.columns:
                df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(0)
            
            # 2. تحويل باقي الأعمدة لنصوص وحل مشكلة الـ (-)
            for col in df.columns:
                if col not in ['Diameter', 'Weight']:
                    df[col] = df[col].astype(str).str.strip().str.upper()
                    df[col] = df[col].replace(['NAN', 'NAT', 'NULL', 'NONE', '<NA>'], '-')
                    df[col].fillna("-", inplace=True)
                    
            return df, None
        return None, f"Sheet '{sheet_name}' not found."
    except Exception as e:
        return None, f"Error: {str(e)}"

def create_pdf(dataframe):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=18)
    elements = []
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(name='Title', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=24, alignment=1, spaceAfter=20, textColor=colors.black)
    elements.append(Paragraph(f"Pipe Quotation: {material_type}", title_style))
    
    date_str = datetime.now().strftime("%Y-%m-%d")
    elements.append(Paragraph(f"Date: {date_str}", styles['Normal']))
    elements.append(Spacer(1, 20))

    print_df = dataframe.copy()
    data = [print_df.columns.to_list()] + print_df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 11),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('ROWBACKGROUNDS', (1, 0), (-1, -1), [colors.white, colors.whitesmoke]),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 40))
    footer_style = ParagraphStyle(name='Footer', parent=styles['Normal'], alignment=1, fontSize=10)
    elements.append(Paragraph("<b>CREATED BY Eng. Ahmed Sabra | Contact: +201148777463</b>", footer_style))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ==========================================
# 5. Application UI
# ==========================================
df = None
error_msg = None

if os.path.exists(data_file):
    df, error_msg = load_data(data_file, material_type)

if df is None:
    st.warning(f"⚠️ {error_msg if error_msg else 'Database file not found.'}")
    uploaded = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded:
        df, error_msg = load_data(uploaded, material_type)

if df is not None:
    if material_type == "HDPE":
        allowed_cols = ['PN', 'SDR']
        spec_cols = [c for c in df.columns if c in allowed_cols]
    else:
        base_cols = ['Diameter', 'Weight']
        spec_cols = [c for c in df.columns if c not in base_cols]

    tab1, tab2 = st.tabs(["1️⃣ Price Calculator (Forward)", "2️⃣ Reverse Analysis (Backward)"])

    with tab1:
        st.subheader("1. Calculate Meter Price")
        c1, c2 = st.columns([1, 2])
        with c1:
            ton_price = st.number_input("Ton Price (EGP):", min_value=0.0, step=500.0)
            dia_unit = st.radio("Input Unit:", ["mm", "Inch"], horizontal=True)
        with c2:
            dia_input_str = st.text_input("Diameters (Comma separated, e.g. 110, 200):")

        user_specs = {}
        if spec_cols:
            cols = st.columns(len(spec_cols))
            for idx, col in enumerate(spec_cols):
                with cols[idx]:
                    vals = [x for x in sorted(df[col].unique().tolist(), key=str) if x != "-"]
                    vals.insert(0, "-")
                    user_specs[col] = st.selectbox(f"Select {col}", vals)

        if st.button("Calculate Price 🚀", type="primary"):
            if ton_price > 0 and dia_input_str:
                try:
                    raw_dias = dia_input_str.replace(" ", ",").split(",")
                    target_dias = [float(x) for x in raw_dias if x.strip() != ""]
                    
                    batch_results = []
                    all_dias_db = sorted(df['Diameter'].unique().tolist())

                    for d_in in target_dias:
                        target_mm = d_in * 25.4 if dia_unit == "Inch" else d_in
                        if all_dias_db:
                            actual_dia = all_dias_db[(np.abs(np.asarray(all_dias_db) - target_mm)).argmin()]
                            mask = (df['Diameter'] == actual_dia)
                            for k, v in user_specs.items():
                                if v != "-": mask &= (df[k] == v)
                            
                            row = df[mask]
                            if not row.empty:
                                w = row.iloc[0]['Weight']
                                if w > 0:
                                    p = (ton_price / 1000) * w
                                    item = {"Material": material_type, "Diameter": actual_dia, "Weight": w, "Price": round(p, 2)}
                                    for col in spec_cols: item[col] = row.iloc[0][col]
                                    batch_results.append(item)
                    
                    if batch_results:
                        st.session_state.current_batch = batch_results
                    else: st.warning("No matches found.")
                except: st.error("Invalid input.")

        if 'current_batch' in st.session_state and st.session_state.current_batch:
            st.dataframe(pd.DataFrame(st.session_state.current_batch), use_container_width=True)
            c_add, c_clr = st.columns([1, 4])
            with c_add:
                if st.button("Add to Final List"):
                    st.session_state.quote_list.extend(st.session_state.current_batch)
                    full = pd.DataFrame(st.session_state.quote_list)
                    valid_keys = [k for k in (['Material', 'Diameter'] + spec_cols) if k in full.columns]
                    full = full.sort_values(by=valid_keys, ascending=True)
                    st.session_state.quote_list = full.to_dict('records')
                    del st.session_state.current_batch
                    st.rerun()
            with c_clr:
                if st.button("Clear Preview"):
                    del st.session_state.current_batch
                    st.rerun()

        st.markdown("---")
        if len(st.session_state.quote_list) > 0:
            st.markdown("### 📋 Final Quotation List")
            final_df = pd.DataFrame(st.session_state.quote_list)
            st.dataframe(final_df, use_container_width=True)
            
            cp, cc = st.columns(2)
            with cp:
                pdf = create_pdf(final_df)
                st.download_button("📄 Download PDF (Landscape)", pdf, "Quotation.pdf", "application/pdf", type="primary")
            with cc:
                if st.button("🗑️ Clear All"):
                    st.session_state.quote_list = []
                    st.rerun()

    with tab2:
        st.subheader("2. Analyze Offer (Find Ton Price)")
        c1, c2 = st.columns(2)
        op = c1.number_input("Offer Meter Price (EGP):", min_value=0.0)
        rd = c2.selectbox("Select Diameter (mm):", sorted(df['Diameter'].unique().tolist()) if not df.empty else [])
        
        rev_specs = {}
        if spec_cols:
            cols = st.columns(len(spec_cols))
            for idx, col in enumerate(spec_cols):
                with cols[idx]:
                    vals = [x for x in sorted(df[col].unique().tolist(), key=str) if x != "-"]
                    vals.insert(0, "-")
                    rev_specs[col] = st.selectbox(f"Select {col}", vals, key=f"t2_{col}")

        if st.button("Analyze Offer 🔍", type="secondary"):
            if op > 0:
                mask = (df['Diameter'] == rd)
                for k, v in rev_specs.items():
                    if v != "-": mask &= (df[k] == v)
                
                row = df[mask]
                if row.empty:
                    st.warning("❌ Item not found.")
                else:
                    unique_weights = [w for w in row['Weight'].unique() if w > 0]
                    if len(unique_weights) == 0:
                        st.error("⚠️ Weight is 0.")
                    elif len(unique_weights) == 1:
                        w = unique_weights[0]
                        est_ton = (op / w) * 1000
                        st.success(f"🏭 Estimated Ton Price: **{est_ton:,.2f} EGP**")
                    else:
                        st.info("💡 Found multiple possible pipes:")
                        results_table = row.copy()
                        results_table['Calculated Ton Price'] = (op / results_table['Weight']) * 1000
                        display_cols = ['Diameter', 'Weight', 'Calculated Ton Price'] + [c for c in spec_cols if c in results_table.columns]
                        st.dataframe(results_table[display_cols].style.format({'Calculated Ton Price': '{:,.2f}', 'Weight': '{:.3f}'}))
            else:
                st.warning("Please enter an offer price.")
