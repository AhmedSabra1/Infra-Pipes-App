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
st.set_page_config(page_title="Professional BOQ Pricing Tool", layout="wide", page_icon="🏗️")

st.markdown("""
<style>
    .main-container { background-color: #0077b5; padding: 2rem; border-radius: 15px; margin-bottom: 2rem; }
    .main-header { font-size: 3.5rem; color: #ffffff !important; font-weight: 900; margin-bottom: 0.2rem; font-family: sans-serif; }
    .sub-header { font-size: 1.4rem; color: rgba(255, 255, 255, 0.9) !important; font-weight: 600; text-transform: uppercase; }
    .stNumberInput label p, .stTextInput label p, .stSelectbox label p, .stRadio label p { font-size: 1.2rem !important; font-weight: bold !important; color: var(--text-color) !important; }
    .grand-total { font-size: 2rem; font-weight: 900; color: #28a745; text-align: right; padding: 10px; border-top: 3px solid #28a745; margin-top: 20px;}
    @media print {
        .main-container { background-color: white !important; color: black !important; }
        .main-header { color: black !important; }
        .sub-header { color: #333 !important; }
        .stButton, footer, header { display: none !important; }
    }
</style>
<div class="main-container">
    <div class="main-header">Pro BOQ Pipe Pricing Tool</div>
    <div class="sub-header">CREATED BY ENG. AHMED SABRA</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# 2. Session State Initialization
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
# 4. Core Logic & Exporters
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
            if 'Diameter' in df.columns: df['Diameter'] = pd.to_numeric(df['Diameter'], errors='coerce').fillna(0)
            if 'Weight' in df.columns: df['Weight'] = pd.to_numeric(df['Weight'], errors='coerce').fillna(0)
            for col in df.columns:
                if col not in ['Diameter', 'Weight']:
                    df[col] = df[col].astype(str).str.strip().str.upper().replace(['NAN', 'NAT', 'NULL', '<NA>'], '-')
                    df[col].fillna("-", inplace=True)
            return df, None
        return None, f"Sheet '{sheet_name}' not found."
    except Exception as e:
        return None, f"Error: {str(e)}"

def format_currency(val):
    return f"{val:,.2f}"

def prepare_export_df():
    df_export = pd.DataFrame(st.session_state.quote_list)
    if not df_export.empty:
        df_export['Unit Price (EGP/m)'] = df_export['Unit Price (EGP/m)'].apply(lambda x: round(float(x), 2))
        df_export['Total (EGP)'] = df_export['Total (EGP)'].apply(lambda x: round(float(x), 2))
    return df_export

def create_pdf(dataframe, grand_total):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=30, bottomMargin=18)
    elements, styles = [], getSampleStyleSheet()
    
    title_style = ParagraphStyle(name='Title', parent=styles['Title'], fontName='Helvetica-Bold', fontSize=22, alignment=1, spaceAfter=15)
    elements.append(Paragraph(f"BOQ Quotation: {material_type}", title_style))
    elements.append(Paragraph(f"Date: {datetime.now().strftime('%Y-%m-%d')}", styles['Normal']))
    elements.append(Spacer(1, 15))

    print_df = dataframe.copy()
    print_df['Unit Price (EGP/m)'] = print_df['Unit Price (EGP/m)'].apply(format_currency)
    print_df['Total (EGP)'] = print_df['Total (EGP)'].apply(format_currency)
    
    data = [print_df.columns.to_list()] + print_df.values.tolist()
    empty_row = [""] * (len(print_df.columns) - 2)
    data.append(empty_row + ["GRAND TOTAL:", f"{grand_total:,.2f} EGP"])

    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('GRID', (0, 0), (-1, -2), 0.5, colors.black),
        ('FONTNAME', (-2, -1), (-1, -1), 'Helvetica-Bold'),
        ('TEXTCOLOR', (-2, -1), (-1, -1), colors.darkgreen),
        ('ALIGN', (-2, -1), (-1, -1), 'RIGHT'),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 30))
    elements.append(Paragraph("<b>CREATED BY Eng. Ahmed Sabra | Contact: +201148777463</b>", ParagraphStyle(name='Footer', alignment=1)))

    doc.build(elements)
    buffer.seek(0)
    return buffer

def create_excel(dataframe, grand_total):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='BOQ')
        worksheet = writer.sheets['BOQ']
        last_row = len(dataframe) + 2
        worksheet.cell(row=last_row, column=len(dataframe.columns)-1, value="GRAND TOTAL:")
        worksheet.cell(row=last_row, column=len(dataframe.columns), value=grand_total)
    return output.getvalue()

# ==========================================
# 5. Application UI
# ==========================================
df, error_msg = None, None
if os.path.exists(data_file): df, error_msg = load_data(data_file, material_type)
if df is None:
    st.warning(f"⚠️ {error_msg if error_msg else 'File not found.'}")
    uploaded = st.sidebar.file_uploader("Upload Excel", type=["xlsx"])
    if uploaded: df, error_msg = load_data(uploaded, material_type)

if df is not None:
    spec_cols = [c for c in df.columns if c in ['PN', 'SDR']] if material_type == "HDPE" else [c for c in df.columns if c not in ['Diameter', 'Weight']]

    tab1, tab2 = st.tabs(["📋 1. BOQ Builder (Forward Pricing)", "🕵️ 2. Reverse Analysis"])

    with tab1:
        st.subheader("Add Pipes to BOQ")
        with st.container(border=True):
            c1, c2, c3 = st.columns([1.5, 1.5, 1])
            with c1:
                price_unit = st.radio("Price Per:", ["Ton", "Kg"], horizontal=True)
                base_price = st.number_input(f"Base Price (EGP/{price_unit}):", min_value=0.0, step=500.0 if price_unit=="Ton" else 5.0)
            with c2:
                dia_unit = st.radio("Diameter Unit:", ["mm", "Inch"], horizontal=True)
                dia_input_str = st.text_input("Diameters (e.g. 110, 200):")
            with c3:
                st.write("") # spacing
                batch_qty = st.number_input("Quantity (m):", min_value=1.0, value=1.0, step=10.0, help="Tip: You can add 1m now and edit quantities easily in the table below!")

            user_specs = {}
            if spec_cols:
                cols = st.columns(len(spec_cols))
                for idx, col in enumerate(spec_cols):
                    with cols[idx]:
                        vals = ["-"] + [x for x in sorted(df[col].unique().tolist(), key=str) if x != "-"]
                        user_specs[col] = st.selectbox(f"Select {col}", vals)

            if st.button("➕ Calculate & Add to BOQ", type="primary"):
                if base_price > 0 and dia_input_str:
                    try:
                        raw_dias = [float(x.strip()) for x in dia_input_str.replace(" ", ",").split(",") if x.strip()]
                        all_dias_db = sorted(df['Diameter'].unique().tolist())
                        
                        price_per_kg = base_price / 1000 if price_unit == "Ton" else base_price

                        for d_in in raw_dias:
                            target_mm = d_in * 25.4 if dia_unit == "Inch" else d_in
                            if all_dias_db:
                                actual_dia = all_dias_db[(np.abs(np.asarray(all_dias_db) - target_mm)).argmin()]
                                mask = (df['Diameter'] == actual_dia)
                                for k, v in user_specs.items():
                                    if v != "-": mask &= (df[k] == v)
                                
                                row = df[mask]
                                if not row.empty:
                                    w = float(row.iloc[0]['Weight'])
                                    if w > 0:
                                        unit_price = price_per_kg * w
                                        total_price = unit_price * batch_qty
                                        
                                        item = {
                                            "Material": material_type,
                                            "Diameter": actual_dia
                                        }
                                        for col in spec_cols: item[col] = row.iloc[0][col]
                                        item["Weight (kg/m)"] = w
                                        item["Unit Price (EGP/m)"] = unit_price
                                        item["Quantity (m)"] = batch_qty
                                        item["Total (EGP)"] = total_price
                                        
                                        st.session_state.quote_list.append(item)
                        # The fix for the error: Catching Exception so it doesn't block Rerun
                        st.rerun()
                    except Exception as e:
                        if type(e).__name__ != 'RerunException':
                            st.error("❌ Please make sure you entered numbers only for the diameters.")
                else: st.warning("⚠️ Please enter price and diameters.")

        st.markdown("---")
        st.subheader("📄 Print Preview & BOQ Editor")
        
        if len(st.session_state.quote_list) > 0:
            current_df = pd.DataFrame(st.session_state.quote_list)
            
            with st.expander("🛠️ Bulk Update Prices (Apply new price to entire list)"):
                uc1, uc2, uc3 = st.columns([1, 1, 2])
                up_unit = uc1.radio("New Price Per:", ["Ton", "Kg"], horizontal=True, key="up_u")
                new_base_price = uc2.number_input("New Base Price:", min_value=0.0, step=100.0)
                if uc3.button("🔄 Update All Items"):
                    if new_base_price > 0:
                        pkg = new_base_price / 1000 if up_unit == "Ton" else new_base_price
                        for item in st.session_state.quote_list:
                            item['Unit Price (EGP/m)'] = item['Weight (kg/m)'] * pkg
                            item['Total (EGP)'] = item['Unit Price (EGP/m)'] * item['Quantity (m)']
                        st.success("✅ Prices Updated!")
                        st.rerun()

            st.caption("💡 TIP: Edit the 'Quantity (m)' directly in the table below! The totals will update instantly.")
            
            disabled_cols = ["Material", "Diameter", "Weight (kg/m)", "Unit Price (EGP/m)", "Total (EGP)"] + spec_cols
            
            edited_df = st.data_editor(
                current_df, 
                num_rows="dynamic", 
                disabled=disabled_cols,
                use_container_width=True,
                hide_index=True
            )
            
            if not edited_df.equals(current_df):
                edited_df['Total (EGP)'] = edited_df['Unit Price (EGP/m)'] * edited_df['Quantity (m)']
                st.session_state.quote_list = edited_df.to_dict('records')
                st.rerun()

            grand_total = edited_df['Total (EGP)'].sum()
            st.markdown(f"<div class='grand-total'>💰 GRAND TOTAL: {grand_total:,.2f} EGP</div>", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True)
            export_df = prepare_export_df()
            ex1, ex2, ex3 = st.columns(3)
            with ex1:
                pdf_file = create_pdf(export_df, grand_total)
                st.download_button("📄 Download PDF", pdf_file, f"BOQ_{material_type}.pdf", "application/pdf", type="primary")
            with ex2:
                excel_file = create_excel(export_df, grand_total)
                st.download_button("📊 Download Excel", excel_file, f"BOQ_{material_type}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
            with ex3:
                if st.button("🗑️ Clear BOQ List"):
                    st.session_state.quote_list = []
                    st.rerun()
        else:
            st.info("Your BOQ list is empty. Add pipes from the section above.")

    with tab2:
        st.subheader("🕵️ Find Supplier's Ton Price")
        c1, c2 = st.columns(2)
        op = c1.number_input("Offer Meter Price (EGP/m):", min_value=0.0)
        rd = c2.selectbox("Select Diameter (mm):", sorted(df['Diameter'].unique().tolist()) if not df.empty else [])
        
        rev_specs = {}
        if spec_cols:
            cols = st.columns(len(spec_cols))
            for idx, col in enumerate(spec_cols):
                with cols[idx]:
                    vals = ["-"] + [x for x in sorted(df[col].unique().tolist(), key=str) if x != "-"]
                    rev_specs[col] = st.selectbox(f"Select {col}", vals, key=f"t2_{col}")

        if st.button("🔍 Analyze Offer", type="secondary"):
            if op > 0:
                mask = (df['Diameter'] == rd)
                for k, v in rev_specs.items():
                    if v != "-": mask &= (df[k] == v)
                
                row = df[mask]
                if row.empty: st.warning("❌ Item not found.")
                else:
                    w_vals = [w for w in row['Weight'].unique() if w > 0]
                    if not w_vals: st.error("⚠️ Weight is 0.")
                    elif len(w_vals) == 1:
                        w = w_vals[0]
                        est_ton = (op / w) * 1000
                        st.success(f"🏭 Estimated Base Price: **{est_ton:,.2f} EGP / Ton** (or {op/w:,.2f} EGP/Kg)")
                    else:
                        st.info("💡 Found multiple pipes:")
                        res = row.copy()
                        res['Est. Ton Price'] = (op / res['Weight']) * 1000
                        st.dataframe(res)
            else: st.warning("Please enter a price.")
