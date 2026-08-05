import io
import re
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
# 2. CUSTOM CSS: HIDE ALL HEADER CONTENT
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
# 3. CACHED DATA LOADING & CLEANING
# ---------------------------------------------------------
@st.cache_data(show_spinner=False)
def load_excel_data(file_bytes, file_name, sheet_name=None):
    try:
        file_obj = io.BytesIO(file_bytes)
        if file_name.endswith('.csv'):
            df = pd.read_csv(file_obj)
        else:
            excel = pd.ExcelFile(file_obj)
            selected = sheet_name if sheet_name else excel.sheet_names[0]
            df = excel.parse(selected)
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception as e:
        st.error(f"Gagal membaca file: {str(e)}")
        return pd.DataFrame()

@st.cache_data(show_spinner=False)
def clean_data_advanced(df, do_trim=True, do_upper=False, do_lower=False, drop_dups=True):
    df_clean = df.copy()
    
    # Konversi seluruh nama kolom ke tipe String
    df_clean.columns = [str(col).strip() for col in df_clean.columns]
    
    # Penanganan Kolom Duplikat yang Aman
    cols = pd.Series(df_clean.columns)
    for dup in cols[cols.duplicated()].unique():
        dup_mask = cols == dup
        dup_indices = cols[dup_mask].index
        for idx_count, idx in enumerate(dup_indices):
            if idx_count > 0:
                cols.iloc[idx] = f"{dup}_{idx_count}"
    df_clean.columns = cols

    if drop_dups:
        df_clean = df_clean.drop_duplicates()
    
    str_cols = df_clean.select_dtypes(include=['object']).columns
    for col in str_cols:
        if do_trim:
            df_clean[col] = df_clean[col].astype(str).str.strip()
        if do_upper:
            df_clean[col] = df_clean[col].astype(str).str.upper()
        elif do_lower:
            df_clean[col] = df_clean[col].astype(str).str.lower()

    # Konversi otomatis kolom angka yang bermasalah (Pembersihan format IDR/USD, koma, titik)
    for col in df_clean.columns:
        if str(col).lower() in ['qty', 'quantity', 'jumlah', 'weight', 'sku qty', 'price', 'total', 'harga']:
            if df_clean[col].dtype == object:
                # Menghilangkan mata uang dan karakter non-numerik kecuali minus, titik, koma
                cleaned_series = df_clean[col].astype(str).str.replace(r'[^\d.,-]', '', regex=True)
                # Standarisasi format koma/titik ribuan
                cleaned_series = cleaned_series.str.replace(',', '.', regex=False)
                df_clean[col] = pd.to_numeric(cleaned_series, errors='coerce')
            else:
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

    # Helper kolom angka jika data kualitatif tanpa angka murni
    num_cols = df_clean.select_dtypes(include=['number']).columns
    if len(num_cols) == 0:
        df_clean['Jumlah_Data'] = 1
            
    return df_clean

# ---------------------------------------------------------
# 4. INITIALIZATION GROQ API & AI ENGINE
# ---------------------------------------------------------
api_key = st.secrets.get("GROQ_API_KEY", None)

if "summary_cache" not in st.session_state:
    st.session_state.summary_cache = {}

def get_ai_insight(df_summary_str, context_info, cache_key=None):
    if cache_key and cache_key in st.session_state.summary_cache:
        return st.session_state.summary_cache[cache_key]

    if not api_key:
        return "⚠️ GROQ_API_KEY tidak ditemukan di Streamlit Secrets."
    try:
        client = Groq(api_key=api_key)
        prompt = f"""
        Anda adalah Senior Business Intelligence Consultant & Data Strategist profesional.
        Analisis data ringkasan berikut secara objektif, teliti, dan logis.
        
        Konteks Data: {context_info}
        Ringkasan Data:
        {df_summary_str}
        
        PENTING HARUS DIPERHATIKAN:
        - Perhatikan angka secara seksama. Angka total atau rata-rata yang lebih besar HARUS dinyatakan lebih tinggi secara benar dan konsisten.
        - Jangan salah membandingkan dua nilai (misalnya: angka 58.000 lebih besar daripada 46.498, maka 58.000 adalah yang tertinggi).
        
        Susunlah analisis eksekutif profesional sesuai jenis datanya (Inventaris, Penjualan, Operasional, SDM, atau Keuangan):
        1. **Executive Summary**: Gambaran umum performa data & agregasi utama (pastikan perbandingan angka 100% akurat).
        2. **Key Insights & Trends**: Temuan krusial, porsi kontribusi terbesar, atau perbedaan signifikan antara frekuensi (count) dan volume per entri (mean).
        3. **Outliers & Anomalies**: Lonjakan, konsentrasi yang tidak seimbang, atau potensi kendala.
        4. **Strategic Recommendations**: Langkah operasional/bisnis konkret yang disarankan.
        5. **Management Conclusion**: Kesimpulan ringkas untuk manajemen.
        
        Gunakan bahasa Indonesia profesional, terstruktur, dan akurat secara matematis.
        """
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",
        )
        res_text = chat_completion.choices[0].message.content
        if cache_key:
            st.session_state.summary_cache[cache_key] = res_text
        return res_text
    except Exception as e:
        return f"Gagal mendapatkan respon AI: {str(e)}"

def prepare_advanced_data_context(df):
    categorical_cols = df.select_dtypes(include=['object', 'category']).columns.tolist()
    numeric_cols = df.select_dtypes(include=['number']).columns.tolist()
    
    ctx = [f"TOTAL ROWS: {len(df)} | TOTAL COLS: {len(df.columns)}"]
    
    if categorical_cols:
        for cat in categorical_cols[:5]:
            top_counts = df[cat].value_counts().head(10).reset_index()
            top_counts.columns = [str(cat), 'Jumlah']
            ctx.append(f"\n--- DISTRIBUSI TOP 10 {str(cat).upper()} ---\n")
            ctx.append(top_counts.to_string(index=False))

    if numeric_cols:
        stats = df[numeric_cols[:3]].describe().T[['mean', 'min', 'max', '50%']]
        ctx.append("\n--- STATISTIK HIMPUNAN ANGKA UTAMA ---\n")
        ctx.append(stats.to_string())
                
    return "\n".join(ctx)

def ask_data_chat(df, user_query):
    if not api_key:
        return "⚠️ GROQ_API_KEY tidak terdeteksi."
    try:
        client = Groq(api_key=api_key)
        
        full_context = prepare_advanced_data_context(df)
        sample_rows = df.head(15).to_string(index=False)
        
        prompt = f"""
        Anda adalah Asisten Data Analyst Profesional. Jawab pertanyaan pengguna berdasarkan data Excel di bawah ini.
        
        === ANALISIS PRE-COMPUTED ===
        {full_context}
        
        === SAMPEL 15 BARIS PERTAMA ===
        {sample_rows}
        
        Pertanyaan Pengguna: {user_query}
        
        Petunjuk Jawaban:
        - Jawab secara akurat, lugas, dan berikan tabel Markdown jika menyajikan ranking/daftar.
        - Perhatikan akurasi angka secara matematis. Jika informasi tidak tersedia di sampel/agregasi di atas, sampaikan keterbatasannya dengan sopan.
        - Sesuaikan konteks jawaban dengan jenis datanya (Penjualan, Inventaris/Inbound, SDM, dll).
        """
        
        chat_completion = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.3-70b-versatile",
        )
        return chat_completion.choices[0].message.content
    except Exception as e:
        return f"Gagal menjawab pertanyaan: {str(e)}"

# ---------------------------------------------------------
# 5. HELPER EXPORT PDF & EXCEL
# ---------------------------------------------------------
def clean_latin_text(text):
    if not text:
        return ""
    replacements = {
        '—': '-', '–': '-', '“': '"', '”': '"', '‘': "'", '’': "'",
        '…': '...', '•': '*', '™': 'TM', '®': '(R)', '©': '(C)'
    }
    for orig, repl in replacements.items():
        text = text.replace(orig, repl)
    return text.encode('latin-1', 'ignore').decode('latin-1')

def generate_smart_pdf(title, ai_insight, df_summary):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    pdf.set_font("Helvetica", style="B", size=14)
    pdf.cell(0, 10, clean_latin_text(title), ln=True, align="C")
    pdf.ln(5)
    
    pdf.set_font("Helvetica", style="B", size=11)
    pdf.cell(0, 8, "AI Executive Summary", ln=True)
    pdf.set_font("Helvetica", size=9)
    clean_text = clean_latin_text(ai_insight.replace("*", "").replace("#", ""))
    pdf.multi_cell(0, 5, clean_text)
    pdf.ln(6)
    
    pdf.set_font("Helvetica", style="B", size=10)
    pdf.cell(0, 8, "Data Summary Table", ln=True)
    
    pdf.set_font("Helvetica", style="B", size=8)
    cols = df_summary.columns.tolist()
    col_width = 180 / max(len(cols), 1)
    
    pdf.set_fill_color(230, 230, 230)
    for col in cols:
        pdf.cell(col_width, 7, clean_latin_text(str(col))[:18], border=1, align="C", fill=True)
    pdf.ln()
    
    pdf.set_font("Helvetica", size=8)
    for _, row in df_summary.head(30).iterrows():
        for col in cols:
            pdf.cell(col_width, 6, clean_latin_text(str(row[col]))[:20], border=1, align="L")
        pdf.ln()
        
    pdf_buffer = io.BytesIO()
    pdf_bytes = bytes(pdf.output())
    pdf_buffer.write(pdf_bytes)
    pdf_buffer.seek(0)
    return pdf_buffer

def convert_df_to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Summary_Result')
    output.seek(0)
    return output

# ---------------------------------------------------------
# 6. JUDUL APLIKASI & UPLOAD DATA
# ---------------------------------------------------------
st.title("🧠 Smart AI Excel Summarizer Pro Studio")
st.caption("Upload file Excel mentah, bersihkan data, filter interaktif, obrolkan data dengan AI, buat pivot & grafik kustom, lalu unduh laporan PDF/Excel.")

st.write("---")
uploaded_file = st.file_uploader("📁 Pilih & Upload File Excel (.xlsx / .xls / .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file:
    file_bytes = uploaded_file.getvalue()
    progress_bar = st.progress(0, text="Membaca file Excel...")
    
    sheet_name = None
    if not uploaded_file.name.endswith('.csv'):
        excel_meta = pd.ExcelFile(io.BytesIO(file_bytes))
        if len(excel_meta.sheet_names) > 1:
            sheet_name = st.selectbox("📄 Pilih Sheet Excel", excel_meta.sheet_names)
            
    progress_bar.progress(50, text="Memuat dataset ke memori...")
    df_raw = load_excel_data(file_bytes, uploaded_file.name, sheet_name)
    progress_bar.progress(100, text="Selesai!")
    progress_bar.empty()

    if df_raw.empty:
        st.error("Data kosong atau gagal dimuat.")
        st.stop()

    # --- PENGATURAN FILTER & CLEANING ---
    with st.expander("🛠️ Pengaturan Filter & Data Cleaning (Opsional)", expanded=False):
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            st.write("**:broom: Auto Data Cleaning Options**")
            auto_clean = st.checkbox("Aktifkan Cleaning Otomatis", value=True)
            do_trim = st.checkbox("Trim Spaces (Hapus Spasi Ekstra)", value=True)
            drop_dups = st.checkbox("Hapus Baris Duplikat", value=True)
            
        df = clean_data_advanced(df_raw, do_trim=do_trim, drop_dups=drop_dups) if auto_clean else df_raw.copy()

        datetime_cols = []
        for col in df.columns:
            if 'date' in str(col).lower() or 'tanggal' in str(col).lower() or pd.api.types.is_datetime64_any_dtype(df[col]):
                try:
                    df[col] = pd.to_datetime(df[col], errors='coerce')
                    if df[col].notna().any():
                        datetime_cols.append(col)
                except Exception:
                    pass

        filtered_df = df.copy()
        
        with col_c2:
            st.write("**:mag: Dynamic Multi-Column Filters**")
            if datetime_cols:
                date_col_selected = st.selectbox("Filter Kolom Tanggal:", datetime_cols)
                valid_dates = filtered_df[date_col_selected].dropna()
                if not valid_dates.empty:
                    min_date = valid_dates.min().date()
                    max_date = valid_dates.max().date()
                    date_range = st.date_input("Rentang Tanggal:", [min_date, max_date])
                    if len(date_range) == 2:
                        start_d, end_d = date_range
                        filtered_df = filtered_df[
                            (filtered_df[date_col_selected].dt.date >= start_d) & 
                            (filtered_df[date_col_selected].dt.date <= end_d)
                        ]

            cat_cols_raw = df.select_dtypes(include=['object', 'category']).columns.tolist()
            if cat_cols_raw:
                for cat_col in cat_cols_raw[:5]:
                    unique_opts = list(df[cat_col].dropna().unique())
                    selected_opts = st.multiselect(f"Filter {cat_col}:", unique_opts, default=[])
                    if selected_opts:
                        filtered_df = filtered_df[filtered_df[cat_col].isin(selected_opts)]

    # Pemisahan kolom kategori & numerik secara ketat
    categorical_cols = filtered_df.select_dtypes(include=['object', 'category']).columns.tolist()
    numeric_cols = filtered_df.select_dtypes(include=['number']).columns.tolist()

    # Filter tambahan: Eliminasi kolom ID/Kode Unik bertipe integer
    valid_numeric_cols = []
    for c in numeric_cols:
        c_str = str(c).lower()
        if 'id' not in c_str and 'code' not in c_str and 'no' not in c_str:
            valid_numeric_cols.append(c)
        elif filtered_df[c].nunique() < len(filtered_df):
            valid_numeric_cols.append(c)
            
    if not valid_numeric_cols:
        valid_numeric_cols = numeric_cols if numeric_cols else ['Jumlah_Data']

    # ---------------------------------------------------------
    # DASHBOARD KPI CARDS & STATISTIK CEPAT
    # ---------------------------------------------------------
    st.write("### 📌 Executive KPI Dashboard")
    kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)
    
    kpi_col1.metric("Total Baris Data", f"{len(filtered_df):,}")
    kpi_col2.metric("Total Kolom", f"{len(filtered_df.columns)}")
    kpi_col3.metric("Missing Values", f"{filtered_df.isna().sum().sum():,}")
    kpi_col4.metric("Baris Duplikat", f"{filtered_df.duplicated().sum():,}")
    
    main_num = valid_numeric_cols[0] if valid_numeric_cols else None
    if main_num and main_num in filtered_df.columns:
        kpi_col5.metric(f"Total {main_num}", f"{filtered_df[main_num].sum():,.0f}")
    else:
        kpi_col5.metric("Kolom Numerik", f"{len(valid_numeric_cols)}")

    # ---------------------------------------------------------
    # ONE CLICK ANALYZE
    # ---------------------------------------------------------
    st.write("")
    if st.button("⚡ One-Click Auto Analyze & Generate Full Report", type="primary", use_container_width=True):
        if len(filtered_df) == 0:
            st.error("Data terfilter kosong. Harap sesuaikan kembali opsi filter Anda.")
        elif categorical_cols and valid_numeric_cols:
            with st.spinner("Menjalankan analisis otomatis 360 derajat..."):
                auto_cat = categorical_cols[0]
                auto_num = valid_numeric_cols[0]
                auto_sum = filtered_df.groupby(auto_cat)[auto_num].agg(['sum', 'mean', 'count']).reset_index().sort_values(by='sum', ascending=False)
                
                ckey = f"{auto_cat}_{auto_num}_oneclick_{len(filtered_df)}"
                auto_res = get_ai_insight(auto_sum.head(15).to_string(index=False), f"Analisis Otomatis Dimensi '{auto_cat}' terhadap Metrik '{auto_num}'", cache_key=ckey)
                
                st.session_state['res'] = auto_res
                st.session_state['data'] = auto_sum
                st.success("✅ One-Click Analysis Selesai! Hasil dapat dilihat pada tab 'AI Executive Summary'.")
        else:
            st.warning("Membutuhkan minimal 1 kolom kategori dan 1 kolom numerik untuk One-Click Analyze.")

    # ---------------------------------------------------------
    # TAB NAVIGASI UTAMA
    # ---------------------------------------------------------
    st.write("---")
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🤖 AI Executive Summary", 
        "💬 Talk to Data (Q&A)", 
        "🔀 Custom Pivot & Charts", 
        "📊 Data Profiling & Quality",
        "📌 Raw Data & Export"
    ])

    # --- TAB 1: AI EXECUTIVE SUMMARY ---
    with tab1:
        st.subheader("🤖 Analisis Naratif Berbasis AI")
        if len(filtered_df) == 0:
            st.warning("Data terfilter kosong. Harap sesuaikan opsi filter Anda.")
        elif categorical_cols and valid_numeric_cols:
            c1, c2 = st.columns(2)
            with c1:
                group_col = st.selectbox("Dimensi Kategori (Kolom Teks):", categorical_cols, key="tab1_cat")
            with c2:
                default_idx = 0
                for idx, col_name in enumerate(valid_numeric_cols):
                    if str(col_name).lower() in ['qty', 'quantity', 'jumlah']:
                        default_idx = idx
                        break
                val_col = st.selectbox("Metrik Utama (Kolom Angka):", valid_numeric_cols, index=default_idx, key="tab1_num")
                
            summary_df = filtered_df.groupby(group_col)[val_col].agg(['sum', 'mean', 'count']).reset_index()
            summary_df = summary_df.sort_values(by='sum', ascending=False)
            
            cache_key = f"{group_col}_{val_col}_{len(filtered_df)}"
            
            if st.button("🚀 Hasilkan AI Executive Summary", type="secondary"):
                with st.spinner("Groq AI sedang menganalisis data..."):
                    ctx = f"Analisis Kategori '{group_col}' berdasarkan Metrik '{val_col}' (Total data: {len(filtered_df)} baris)."
                    data_str = summary_df.head(15).to_string(index=False)
                    res = get_ai_insight(data_str, ctx, cache_key=cache_key)
                    
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

    # --- TAB 2: CHAT / TALK TO DATA ---
    with tab2:
        st.subheader("💬 Tanya Jawab Interaktif dengan Data Excel")
        st.caption("Ajukan pertanyaan bebas tentang data Anda (contoh: 'Vendor Name mana yang kirim Qty terbanyak?', 'Berapa total Qty per Product Name?')")
        
        if "chat_history" not in st.session_state:
            st.session_state.chat_history = []
            
        if len(st.session_state.chat_history) > 20:
            st.session_state.chat_history = st.session_state.chat_history[-20:]
            
        for message in st.session_state.chat_history:
            with st.chat_message(message["role"]):
                st.markdown(message["content"])
                
        user_prompt = st.chat_input("Ketik pertanyaan Anda tentang data ini...")
        if user_prompt:
            st.session_state.chat_history.append({"role": "user", "content": user_prompt})
            with st.chat_message("user"):
                st.markdown(user_prompt)
                
            with st.chat_message("assistant"):
                with st.spinner("Membaca data dan menghitung..."):
                    ai_response = ask_data_chat(filtered_df, user_prompt)
                    st.markdown(ai_response)
                    
            st.session_state.chat_history.append({"role": "assistant", "content": ai_response})

    # --- TAB 3: CHART SELECTOR & PIVOT STUDIO ---
    with tab3:
        st.subheader("🔀 Custom Pivot & Dynamic Chart Studio")
        
        c_p1, c_p2, c_p3 = st.columns(3)
        with c_p1:
            p_rows = st.multiselect("Rows (Baris/Kategori):", categorical_cols if categorical_cols else filtered_df.columns.tolist(), default=[categorical_cols[0]] if categorical_cols else [])
        with c_p2:
            p_vals = st.selectbox("Values (Nilai Metrik):", valid_numeric_cols if valid_numeric_cols else filtered_df.columns.tolist())
        with c_p3:
            p_agg = st.selectbox("Fungsi Agregasi:", ["sum", "count", "mean", "min", "max"])
            
        if p_rows and p_vals:
            try:
                p_res = pd.pivot_table(filtered_df, index=p_rows, values=p_vals, aggfunc=p_agg).reset_index()
                st.dataframe(p_res, use_container_width=True)
                
                st.divider()
                st.write("### 📊 Visualisasi Grafik")
                
                available_charts = ["Bar Chart", "Line Chart", "Pie Chart"]
                chart_type = st.radio("Pilih Jenis Grafik:", available_charts, horizontal=True)
                
                fig = None
                if chart_type == "Bar Chart":
                    fig = px.bar(p_res, x=p_rows[0], y=p_vals, title=f"{p_agg.upper()} {p_vals} per {p_rows[0]}")
                elif chart_type == "Line Chart":
                    fig = px.line(p_res, x=p_rows[0], y=p_vals, markers=True, title=f"Tren {p_vals} per {p_rows[0]}")
                elif chart_type == "Pie Chart":
                    if p_res[p_vals].sum() <= 0:
                        st.warning("⚠️ Total nilai agregasi kurang dari atau sama dengan 0. Pie Chart tidak dapat ditampilkan.")
                    else:
                        if len(p_res) > 20:
                            st.warning("⚠️ Kategori lebih dari 20 item. Pie Chart otomatis membatasi ke Top 10 kategori terbesar.")
                            p_res_pie = p_res.sort_values(by=p_vals, ascending=False).head(10)
                        else:
                            p_res_pie = p_res
                        fig = px.pie(p_res_pie, names=p_rows[0], values=p_vals, title=f"Proporsi {p_vals} per {p_rows[0]}")
                    
                if fig:
                    st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Gagal membuat Pivot Table: {str(e)}.")

    # --- TAB 4: DATA PROFILING & MISSING VALUE REPORT ---
    with tab4:
        st.subheader("📊 Data Profiling & Quality Report")
        
        prof_c1, prof_c2 = st.columns(2)
        with prof_c1:
            st.write("**:mag: Missing Value Report**")
            null_df = pd.DataFrame({
                'Kolom': [str(c) for c in filtered_df.columns],
                'Missing Count': filtered_df.isna().sum().values,
                'Missing (%)': (filtered_df.isna().sum().values / max(len(filtered_df), 1) * 100).round(2)
            }).sort_values(by='Missing Count', ascending=False)
            st.dataframe(null_df, use_container_width=True)
            
        with prof_c2:
            st.write("**:info: Dataset Overview & Memory**")
            dtype_df = pd.DataFrame({
                'Kolom': [str(c) for c in filtered_df.columns],
                'Data Type': filtered_df.dtypes.astype(str).values,
                'Unique Values': [filtered_df[c].nunique() for c in filtered_df.columns]
            })
            st.dataframe(dtype_df, use_container_width=True)
            st.caption(f"Total Memori Digunakan: **{filtered_df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB**")

    # --- TAB 5: PREVIEW DATA & EXPORT EXCEL ---
    with tab5:
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
