import io
import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
import google.generativeai as genai

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Smart AI Excel Summarizer",
    page_icon="🧠",
    layout="wide"
)

# ---------------------------------------------------------
# 2. CUSTOM CSS: HIDE STREAMLIT HEADER & WATERMARK
# ---------------------------------------------------------
hide_streamlit_style = """
            <style>
            #MainMenu {visibility: hidden;}
            header {visibility: hidden;}
            footer {visibility: hidden;}
            
            .created-by {
                text-align: right;
                color: #6c757d;
                font-size: 0.9rem;
                font-weight: bold;
                margin-top: -10px;
                margin-bottom: 20px;
                border-bottom: 1px solid #e9ecef;
                padding-bottom: 8px;
            }
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Watermark / Author Name
st.markdown('<div class="created-by">Created by iqbalmantam</div>', unsafe_allow_html=True)

# ---------------------------------------------------------
# 3. INITIALIZATION GEMINI API FROM SECRETS
# ---------------------------------------------------------
api_key = st.secrets.get("GEMINI_API_KEY", None)

def get_ai_insight(df_summary_str, context_info):
    """Fungsi untuk mengirim data ringkasan ke Gemini AI."""
    if not api_key:
        return "⚠️ API Key tidak ditemukan di Streamlit Secrets. Pastikan GEMINI_API_KEY sudah terpasang di Secrets."
    
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash')
        
        prompt = f"""
        Kamu adalah seorang Senior Data Analyst. Analisis data ringkasan berikut dari sebuah file Excel.
        
        Konteks Data: {context_info}
        Ringkasan Data:
        {df_summary_str}
        
        Tolong buatkan Executive Summary yang cerdas dan profesional dalam bahasa Indonesia:
        1. **Temuan Utama (Key Takeaways)**: Poin-poin paling krusial.
        2. **Anomali / Tren Menarik**: Pola unik, lonjakan, atau penurunan signifikan.
        3. **Rekomendasi Bisnis**: Tindakan konkret yang sebaiknya diambil manajemen berdasarkan data ini.
        Format respons menggunakan Markdown yang rapi dan lugas.
        """
        response = model.generate_content(prompt)
        return response.text
    except Exception as e:
        return f"Gagal mendapatkan respon AI: {str(e)}"

# ---------------------------------------------------------
# 4. HELPER EXPORT PDF
# ---------------------------------------------------------
def generate_smart_pdf(title, ai_insight, df_summary):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Title
    pdf.set_font("Helvetica", style="B", size=15)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.ln(5)
    
    # AI Summary
    pdf.set_font("Helvetica", style="B", size=11)
    pdf.cell(0, 8, "AI Executive Summary", ln=True)
    pdf.set_font("Helvetica", size=8.5)
    clean_text = ai_insight.replace("*", "").replace("#", "")
    pdf.multi_cell(0, 5, clean_text)
    pdf.ln(6)
    
    # Table Data
    pdf.set_font("Helvetica", style="B", size=10)
    pdf.cell(0, 8, "Data Summary Table", ln=True)
    
    pdf.set_font("Helvetica", style="B", size=8)
    cols = df_summary.columns.tolist()
    col_width = 180 / max(len(cols), 1)
    
    pdf.set_fill_color(230, 230, 230)
    for col in cols:
        pdf.cell(col_width, 7, str(col)[:18], border=1, align="C", fill=True)
    pdf.ln()
    
    pdf.set_font("Helvetica", size=8)
    for _, row in df_summary.head(30).iterrows():
        for col in cols:
            pdf.cell(col_width, 6, str(row[col])[:20], border=1, align="L")
        pdf.ln()
        
    pdf_buffer = io.BytesIO()
    pdf_bytes = pdf.output()
    pdf_buffer.write(pdf_bytes)
    pdf_buffer.seek(0)
    return pdf_buffer

# ---------------------------------------------------------
# 5. JUDUL APLIKASI & UPLOAD DATA
# ---------------------------------------------------------
st.title("🧠 Smart AI Excel Summarizer & Executive Insights")
st.caption("Upload file Excel mentah. Sistem akan membuat Pivot Table kustom, grafik, dan AI akan menyusun Executive Summary otomatis.")

st.sidebar.header("📁 Upload File Excel")
uploaded_file = st.sidebar.file_uploader("Pilih File (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df = pd.read_csv(uploaded_file)
    else:
        excel_file = pd.ExcelFile(uploaded_file)
        selected_sheet = st.sidebar.selectbox("Pilih Sheet", excel_file.sheet_names)
        df = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()

    tab1, tab2, tab3 = st.tabs(["🤖 AI Executive Summary", "🔀 Custom Pivot Table", "📌 Raw Data Preview"])

    # --- TAB 1: AI EXECUTIVE SUMMARY ---
    with tab1:
        st.subheader("🤖 Analisis Naratif Berbasis Gemini AI")
        if categorical_cols and numeric_cols:
            c1, c2 = st.columns(2)
            with c1:
                group_col = st.selectbox("Dimensi Kategori:", categorical_cols)
            with c2:
                val_col = st.selectbox("Metrik Utama (Angka):", numeric_cols)
                
            summary_df = df.groupby(group_col)[val_col].agg(['sum', 'mean', 'count']).reset_index()
            summary_df = summary_df.sort_values(by='sum', ascending=False)
            
            if st.button("🚀 Hasikan AI Executive Summary", type="primary"):
                with st.spinner("Gemini AI sedang membaca pola data..."):
                    ctx = f"Analisis kategori '{group_col}' terhadap metrik '{val_col}'."
                    data_str = summary_df.head(15).to_string(index=False)
                    res = get_ai_insight(data_str, ctx)
                    
                    st.session_state['res'] = res
                    st.session_state['data'] = summary_df

            if 'res' in st.session_state:
                st.markdown("---")
                st.markdown(st.session_state['res'])
                
                pdf_bytes = generate_smart_pdf(
                    title=f"AI Report: {val_col} per {group_col}",
                    ai_insight=st.session_state['res'],
                    df_summary=st.session_state['data']
                )
                st.download_button(
                    label="📄 Download Executive Report (PDF)",
                    data=pdf_bytes,
                    file_name="AI_Executive_Report.pdf",
                    mime="application/pdf"
                )
        else:
            st.warning("Membutuhkan minimal 1 kolom kategori dan 1 kolom angka.")

    # --- TAB 2: CUSTOM PIVOT STUDIO ---
    with tab2:
        st.subheader("🔀 Custom Pivot Studio")
        p_rows = st.multiselect("Rows (Baris):", df.columns.tolist(), default=[categorical_cols[0]] if categorical_cols else [])
        p_vals = st.selectbox("Values (Nilai):", numeric_cols if numeric_cols else df.columns.tolist())
        p_agg = st.selectbox("Fungsi Agregasi:", ["sum", "mean", "count", "min", "max"])
        
        if p_rows and p_vals:
            p_res = pd.pivot_table(df, index=p_rows, values=p_vals, aggfunc=p_agg).reset_index()
            st.dataframe(p_res, use_container_width=True)
            
            fig = px.bar(p_res, x=p_rows[0], y=p_vals, title=f"{p_agg.upper()} {p_vals} per {p_rows[0]}")
            st.plotly_chart(fig, use_container_width=True)

    # --- TAB 3: PREVIEW DATA ---
    with tab3:
        st.subheader("Raw Data Preview")
        st.write(f"Total Baris: **{len(df)}** | Total Kolom: **{len(df.columns)}**")
        st.dataframe(df, use_container_width=True)

else:
    st.info("💡 Silakan upload file Excel di sidebar kiri untuk memulai.")
