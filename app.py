import io
import streamlit as st
import pandas as pd
import plotly.express as px
from fpdf import FPDF
from groq import Groq

# ---------------------------------------------------------
# 1. KONFIGURASI HALAMAN
# ---------------------------------------------------------
st.set_page_config(
    page_title="Smart AI Excel Summarizer Studio",
    page_icon="🧠",
    layout="wide"
)

# ---------------------------------------------------------
# 2. CUSTOM CSS: HIDE ALL HEADER CONTENT (CLEAN LOOK)
# ---------------------------------------------------------
hide_streamlit_style = """
            <style>
            header[data-testid="stHeader"] {
                display: none !important;
            }
            #MainMenu {visibility: hidden;}
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
# 3. INITIALIZATION GROQ API FROM SECRETS
# ---------------------------------------------------------
api_key = st.secrets.get("GROQ_API_KEY", None)

def get_ai_insight(df_summary_str, context_info):
    """Fungsi Executive Summary berbasis Groq Llama 3.3."""
    if not api_key:
        return "⚠️ GROQ_API_KEY tidak ditemukan di Streamlit Secrets."
    try:
        client = Groq(api_key=api_key)
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
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Gagal mendapatkan respon AI: {str(e)}"

def ask_data_chat(df_preview_str, user_query):
    """Fungsi Tanya Jawab / Chat dengan Data Excel."""
    if not api_key:
        return "⚠️ GROQ_API_KEY tidak terdeteksi."
    try:
        client = Groq(api_key=api_key)
        prompt = f"""
        Kamu adalah Asisten Analis Data. Jawab pertanyaan pengguna berdasarkan sampel data Excel berikut.
        
        Sampel Data / Summary:
        {df_preview_str}
        
        Pertanyaan Pengguna: {user_query}
        
        Berikan jawaban yang singkat, tepat, akurat, dan mudah dipahami.
        """
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Gagal menjawab pertanyaan: {str(e)}"

# ---------------------------------------------------------
# 4. HELPER EXPORT PDF & EXCEL
# ---------------------------------------------------------
def generate_smart_pdf(title, ai_insight, df_summary):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Helvetica", style="B", size=15)
    pdf.cell(0, 10, title, ln=True, align="C")
    pdf.ln(5)
    
    pdf.set_font("Helvetica", style="B", size=11)
    pdf.cell(0, 8, "AI Executive Summary", ln=True)
    pdf.set_font("Helvetica", style="B", size=8.5)
    clean_text = ai_insight.replace("*", "").replace("#", "")
    pdf.multi_cell(0, 5, clean_text)
    pdf.ln(6)
    
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

def convert_df_to_excel(df):
    """Export ke File Excel (.xlsx) Rapi."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary_Result')
    output.seek(0)
    return output

# ---------------------------------------------------------
# 5. JUDUL APLIKASI & UPLOAD DATA LANGSUNG DI HALAMAN UTAMA
# ---------------------------------------------------------
st.title("🧠 Smart AI Excel Summarizer Pro Studio")
st.caption("Upload file Excel mentah, bersihkan data, filter interaktif, obrolkan data dengan AI, buat pivot & grafik kustom, lalu unduh laporan PDF/Excel.")

# UPLOAD FILE DI HALAMAN UTAMA (TIDAK PERLU SIDEBAR)
st.write("---")
uploaded_file = st.file_uploader("📁 Pilih & Upload File Excel (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    if uploaded_file.name.endswith('.csv'):
        df_raw = pd.read_csv(uploaded_file)
    else:
        excel_file = pd.ExcelFile(uploaded_file)
        selected_sheet = st.selectbox("📄 Pilih Sheet Excel", excel_file.sheet_names)
        df_raw = pd.read_excel(uploaded_file, sheet_name=selected_sheet)

    # --- EXPANDER UTILS: DATA CLEANING & FILTER ---
    with st.expander("🛠️ Pengaturan Filter & Data Cleaning (Opsional)", expanded=False):
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            auto_clean = st.checkbox("🧹 Aktifkan Auto-Cleaning (Hapus Duplikat & Trim Spasi)", value=True)
            
        df = df_raw.copy()
        if auto_clean:
            df = df.drop_duplicates()
            str_cols = df.select_dtypes(include=['object']).columns
            for col in str_cols:
                df[col] = df[col].astype(str).str.strip()

        categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
        numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
        
        datetime_cols = []
        for col in df.columns:
            if 'date' in col.lower() or 'tanggal' in col.lower() or pd.api.types.is_datetime64_any_dtype(df[col]):
                try:
                    df[col] = pd.to_datetime(df[col])
                    datetime_cols.append(col)
                except:
                    pass

        filtered_df = df.copy()
        
        with col_c2:
            if datetime_cols:
                date_col_selected = st.selectbox("Filter Kolom Tanggal:", datetime_cols)
                min_date = filtered_df[date_col_selected].min().date()
                max_date = filtered_df[date_col_selected].max().date()
                date_range = st.date_input("Rentang Tanggal:", [min_date, max_date])
                if len(date_range) == 2:
                    start_d, end_d = date_range
                    filtered_df = filtered_df[
                        (filtered_df[date_col_selected].dt.date >= start_d) & 
                        (filtered_df[date_col_selected].dt.date <= end_d)
                    ]

    if 'filtered_df' not in locals():
        filtered_df = df_raw.copy()
        categorical_cols = filtered_df.select_dtypes(include=['object', 'category']).columns.tolist()
        numeric_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()

    # --- TAB NAVIGASI UTAMA ---
    st.write("---")
    tab1, tab2, tab3, tab4 = st.tabs([
        "🤖 AI Executive Summary", 
        "💬 Talk to Data (Q&A)", 
        "🔀 Custom Pivot & Charts Studio", 
        "📌 Data Preview & Export"
    ])

    # TAB 1: AI EXECUTIVE SUMMARY
    with tab1:
        st.subheader("🤖 Analisis Naratif Berbasis Groq AI (Llama 3.3)")
        if categorical_cols and numeric_cols:
            c1, c2 = st.columns(2)
            with c1:
                group_col = st.selectbox("Dimensi Kategori:", categorical_cols)
            with c2:
                val_col = st.selectbox("Metrik Utama (Angka):", numeric_cols)
                
            summary_df = filtered_df.groupby(group_col)[val_col].agg(['sum', 'mean', 'count']).reset_index()
            summary_df = summary_df.sort_values(by='sum', ascending=False)
            
            if st.button("🚀 Hasikan AI Executive Summary", type="primary"):
                with st.spinner("Groq AI sedang menganalisis data terfilter..."):
                    ctx = f"Analisis kategori '{group_col}' terhadap metrik '{val_col}' (Total data: {len(filtered_df)} baris)."
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

    # TAB 2: CHAT / TALK TO DATA
    with tab2:
        st.subheader("💬 Tanya Jawab Interaktif dengan Data Excel")
        st.caption("Ajukan pertanyaan bebas tentang data Anda (contoh: 'Siapa Top 3 Customer terbanyak?', 'Berapa total penjualan?')")
        
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
            
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                
        user_prompt = st.chat_input("Ketik pertanyaan Anda tentang data ini...")
        if user_prompt:
            st.session_state.chat_history.append({"role": "user", "content": user_prompt})
            with st.chat_message("user"):
                st.markdown(user_prompt)
                
            with st.chat_message("assistant"):
                with st.spinner("Membaca data dan berpikir..."):
                    data_sample = filtered_df.describe(include='all').to_string()
                    ai_response = ask_data_chat(data_sample, user_prompt)
                    st.markdown(ai_response)
                    
            st.session_state.chat_history.append({"role": "assistant", "content": ai_response})

    # TAB 3: CHART SELECTOR & PIVOT STUDIO
    with tab3:
        st.subheader("🔀 Custom Pivot & Dynamic Chart Studio")
        
        c_p1, c_p2, c_p3 = st.columns(3)
        with c_p1:
            p_rows = st.multiselect("Rows (Baris):", filtered_df.columns.tolist(), default=[categorical_cols[0]] if categorical_cols else [])
        with c_p2:
            p_vals = st.selectbox("Values (Nilai):", numeric_cols if numeric_cols else filtered_df.columns.tolist())
        with c_p3:
            p_agg = st.selectbox("Fungsi Agregasi:", ["sum", "mean", "count", "min", "max"])
            
        if p_rows and p_vals:
            p_res = pd.pivot_table(filtered_df, index=p_rows, values=p_vals, aggfunc=p_agg).reset_index()
            st.dataframe(p_res, use_container_width=True)
            
            st.divider()
            st.write("### 📊 Visualisasi Grafik")
            chart_type = st.radio("Pilih Jenis Grafik:", ["Bar Chart", "Line Chart", "Pie Chart", "Scatter Plot"], horizontal=True)
            
            if chart_type == "Bar Chart":
                fig = px.bar(p_res, x=p_rows[0], y=p_vals, title=f"{p_agg.upper()} {p_vals} per {p_rows[0]}")
            elif chart_type == "Line Chart":
                fig = px.line(p_res, x=p_rows[0], y=p_vals, markers=True, title=f"Tren {p_vals} per {p_rows[0]}")
            elif chart_type == "Pie Chart":
                fig = px.pie(p_res, names=p_rows[0], values=p_vals, title=f"Proporsi {p_vals} per {p_rows[0]}")
            elif chart_type == "Scatter Plot":
                fig = px.scatter(p_res, x=p_rows[0], y=p_vals, size=p_vals, title=f"Sebaran {p_vals} per {p_rows[0]}")
                
            st.plotly_chart(fig, use_container_width=True)

    # TAB 4: PREVIEW DATA & EXPORT EXCEL
    with tab4:
        st.subheader("📌 Data Preview & Export Excel")
        st.write(f"Total Baris Data Terfilter: **{len(filtered_df)}** (Dari Total Raw: **{len(df_raw)}**)")
        st.dataframe(filtered_df, use_container_width=True)
        
        st.divider()
        st.write("### 📥 Unduh Data Hasil Cleaning & Filter")
        
        col_ex1, col_ex2 = st.columns(2)
        with col_ex1:
            csv_data = filtered_df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Unduh Data (CSV)",
                data=csv_data,
                file_name="data_filtered.csv",
                mime="text/csv",
                use_container_width=True
            )
        with col_ex2:
            excel_data = convert_df_to_excel(filtered_df)
            st.download_button(
                label="📊 Unduh Data Rapi (.xlsx)",
                data=excel_data,
                file_name="data_filtered.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

else:
    st.info("💡 Silakan klik tombol 'Pilih & Upload File Excel' di atas untuk memulai.")
