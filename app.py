import streamlit as st
import pandas as pd
import plotly.express as px
import glob
import urllib.parse
from PIL import Image

# --- FUNGSI MENCARI & MEMBACA LOGO JABATAN ---
def dapatkan_logo():
    senarai_gambar = glob.glob("*.jpg") + glob.glob("*.png") + glob.glob("*.jpeg") + glob.glob("*.JPG") + glob.glob("*.PNG")
    if len(senarai_gambar) > 0:
        try:
            return Image.open(senarai_gambar[0])
        except:
            return "🏢"
    return "🏢"

logo_jabatan = dapatkan_logo()

# 1. Konfigurasi Halaman (Dipaksa untuk sentiasa 'expanded' secara lalai)
st.set_page_config(
    page_title="Dashboard Aset JPNS", 
    page_icon=logo_jabatan, 
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- SUNTIKAN CSS KORPORAT ---
st.markdown("""
    <style>
    [data-testid="stToolbar"] {visibility: hidden !important;}
    [data-testid="stAppDeployButton"] {visibility: hidden !important;}
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    
    /* PENYELAMAT BUTANG SIDEBAR: Memastikan ikon > sentiasa di lapisan paling hadapan dan boleh diklik */
    [data-testid="collapsedControl"] { z-index: 999999 !important; }
    
    .main-title { color: #2C3E50; font-size: 32px; font-weight: 800; line-height: 1.2; letter-spacing: 0.5px; margin-bottom: 0px; }
    .sub-title { color: #7F8C8D; font-size: 16px; font-weight: 500; margin-top: 5px; margin-bottom: 25px;}
    .section-header { color: #34495E; font-size: 18px; font-weight: 700; border-bottom: 2px solid #E5E7E9; padding-bottom: 8px; margin-top: 25px; margin-bottom: 15px;}
    
    [data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; box-shadow: 0px 4px 10px rgba(0,0,0,0.05); }
    [data-testid="stMetricValue"] { font-size: 26px !important; color: #1b5e20; }
    </style>
""", unsafe_allow_html=True)

# --- BAHAGIAN TAJUK DAN LOGO ---
col_logo, col_tajuk = st.columns([1, 9])
with col_logo:
    if isinstance(logo_jabatan, Image.Image):
        st.image(logo_jabatan, width=90)
    else:
        st.info("Logo")
with col_tajuk:
    st.markdown('<div class="main-title">SISTEM DASHBOARD ASET BANGUNAN</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">JABATAN PERHUTANAN NEGERI SEMBILAN (JPNS)</div>', unsafe_allow_html=True)

def proses_multilink_drive(val):
    val_str = str(val).strip()
    if val_str == "-" or val_str == "" or val_str.lower() == "nan": return []
    links = [l.strip() for l in val_str.split(",")]
    processed_links = []
    for url in links:
        if "drive.google.com" in url:
            try:
                file_id = None
                if "/file/d/" in url: file_id = url.split("/file/d/")[1].split("/")[0]
                elif "id=" in url: file_id = url.split("id=")[1].split("&")[0]
                if file_id: processed_links.append(f"https://lh3.googleusercontent.com/d/{file_id}")
            except: pass
    return processed_links

# 2. Fungsi Membaca Data
def load_all_data_combined():
    try:
        xls = pd.ExcelFile('data.xlsx')
        all_sheets_list = []
        for sheet in xls.sheet_names:
            df_raw = pd.read_excel(xls, sheet_name=sheet, header=None)
            header_idx = 3 
            for i in range(min(15, len(df_raw))):
                row_vals = [str(x).lower() for x in df_raw.iloc[i].values]
                if any(katakunci in val for val in row_vals for katakunci in ['perkara', 'daerah', 'aset', 'bangunan', 'fasiliti', 'keterangan', 'item']):
                    header_idx = i
                    break
            
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=header_idx)
            df = df.dropna(how='all').fillna("-")
            df.columns = [str(c).strip() for c in df.columns]
            
            daerah_sivil_col = None; daerah_pentadbiran_col = None; status_col = None
            gps_col = None; gambar_perkara_col = None; lokasi_col = None; image_cols = []

            for c in df.columns:
                c_low = c.lower()
                if 'sivil' in c_low: daerah_sivil_col = c
                elif 'pentadbiran' in c_low: daerah_pentadbiran_col = c
                elif c_low == 'daerah' and not daerah_sivil_col: daerah_sivil_col = c
                elif 'status' in c_low and 'kefungsian' in c_low: status_col = c
                elif 'gps' in c_low or 'kedudukan' in c_low: gps_col = c
                elif any(kata in c_low for kata in ['perkara', 'aset', 'bangunan', 'fasiliti', 'nama', 'keterangan', 'butiran', 'item']): 
                    if not gambar_perkara_col: gambar_perkara_col = c 
                elif 'lokasi' in c_low: lokasi_col = c
                elif 'gambar' in c_low and 'perkara' not in c_low: image_cols.append(c)

            if not daerah_sivil_col:
                for c in df.columns:
                    if 'daerah' in c.lower() and 'pentadbiran' not in c.lower():
                        daerah_sivil_col = c
                        break

            new_df = pd.DataFrame()
            if gambar_perkara_col: new_df['Nama_Fasiliti'] = df[gambar_perkara_col].astype(str).str.strip()
            else: continue 
                
            saringan_nama = new_df['Nama_Fasiliti'].str.lower()
            mask_sah = ~(saringan_nama.isin(['nan', 'none', '-', '', 'null'])) & (new_df['Nama_Fasiliti'].str.len() > 2) & (~saringan_nama.str.startswith('http'))
            new_df = new_df[mask_sah]
            
            valid_indices = new_df.index
            if len(valid_indices) == 0: continue
                
            new_df['Kategori_Fasiliti'] = str(sheet).strip().title()
            
            if daerah_sivil_col: new_df['Daerah Sivil'] = df.loc[valid_indices, daerah_sivil_col].astype(str).str.strip().str.title()
            else: new_df['Daerah Sivil'] = "Tidak Dinyatakan"
                
            if daerah_pentadbiran_col: new_df['Daerah Pentadbiran'] = df.loc[valid_indices, daerah_pentadbiran_col].astype(str).str.strip().str.upper()
            else: new_df['Daerah Pentadbiran'] = "Tidak Dinyatakan"
                
            if lokasi_col: new_df['Lokasi'] = df.loc[valid_indices, lokasi_col].astype(str).str.strip().str.title()
            else: new_df['Lokasi'] = "-"

            if status_col:
                status_series = df.loc[valid_indices, status_col].astype(str)
                new_df['Status_Bersih'] = status_series.apply(lambda x: x.split('/')[0].strip().title() if '/' in x else x.strip().title())
                new_df['Jenis_Bangunan'] = status_series.apply(lambda x: x.split('/')[1].strip().title() if '/' in x else "-")
            else:
                new_df['Status_Bersih'] = "-"; new_df['Jenis_Bangunan'] = "-"

            if gps_col:
                gps_series = df.loc[valid_indices, gps_col].astype(str)
                def bina_link_maps(gps):
                    gps = gps.strip()
                    if gps.lower().startswith('http'): return gps 
                    elif gps.lower() not in ['nan', 'none', '', '-']: return f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(gps)}"
                    return None
                new_df['Pautan_Peta'] = gps_series.apply(bina_link_maps)
            else: new_df['Pautan_Peta'] = None

            def proses_links_dari_row(row_index):
                all_processed = []
                for col in image_cols:
                    val_str = str(df.loc[row_index, col]).strip()
                    if val_str == "-" or val_str == "" or val_str.lower() == "nan": continue
                    if "," in val_str: links = [l.strip() for l in val_str.split(",")]
                    elif ";" in val_str: links = [l.strip() for l in val_str.split(";")]
                    else: links = [val_str]
                    for url in links:
                        if "drive.google.com" in url:
                            try:
                                file_id = None
                                if "/file/d/" in url: file_id = url.split("/file/d/")[1].split("/")[0]
                                elif "id=" in url: file_id = url.split("id=")[1].split("&")[0]
                                if file_id: all_processed.append(f"https://lh3.googleusercontent.com/d/{file_id}")
                            except: pass
                return all_processed

            new_df['Senarai_Imej'] = [proses_links_dari_row(idx) for idx in valid_indices]
            if not new_df.empty: all_sheets_list.append(new_df)
                
        if len(all_sheets_list) > 0: return pd.concat(all_sheets_list, ignore_index=True)
        return None
    except Exception as e:
        st.error(f"Ralat Sistem: {e}")
        return None

df_master = load_all_data_combined()

if df_master is not None and not df_master.empty:
    
    # 3. SIDEBAR KORPORAT
    st.sidebar.markdown('<div class="section-header">Panel Tapisan Aset</div>', unsafe_allow_html=True)
    
    senarai_pentadbiran = sorted(df_master["Daerah Pentadbiran"].unique().tolist())
    pentadbiran_terpilih = st.sidebar.selectbox(
        "Daerah Pentadbiran:",
        options=["■ SEMUA PENTADBIRAN"] + senarai_pentadbiran,
        index=0
    )
    
    if pentadbiran_terpilih != "■ SEMUA PENTADBIRAN":
        df_sementara_kategori = df_master[df_master["Daerah Pentadbiran"] == pentadbiran_terpilih]
    else:
        df_sementara_kategori = df_master.copy()
        
    senarai_kategori_dinamik = sorted(df_sementara_kategori["Kategori_Fasiliti"].unique().tolist())
    
    kategori_terpilih = st.sidebar.selectbox(
        "Kategori Fasiliti:",
        options=["■ SEMUA KATEGORI"] + senarai_kategori_dinamik,
        index=0
    )
    
    df_filtered = df_master.copy()
    if pentadbiran_terpilih != "■ SEMUA PENTADBIRAN":
        df_filtered = df_filtered[df_filtered["Daerah Pentadbiran"] == pentadbiran_terpilih]
    if kategori_terpilih != "■ SEMUA KATEGORI":
        df_filtered = df_filtered[df_filtered["Kategori_Fasiliti"] == kategori_terpilih]

    t_kategori = kategori_terpilih.replace("■ ", "")
    t_pentadbiran = pentadbiran_terpilih.replace("■ ", "")

    # 4. RINGKASAN EKSEKUTIF (4 Kotak)
    total_aset = len(df_filtered)
    aset_baik = df_filtered['Status_Bersih'].str.contains('Baik', case=False, na=False).sum()
    
    m1, m2, m3, m4 = st.columns(4)
    with m1: st.metric(label="Jumlah Keseluruhan", value=f"{total_aset} Buah")
    with m2: st.metric(label="Daerah Pentadbiran", value=t_pentadbiran)
    with m3: st.metric(label="Kategori Terpilih", value=t_kategori)
    with m4: st.metric(label="Kondisi Baik", value=f"{aset_baik} Unit")

    # 5. VISUALISASI GRAFIK 
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown('<div class="section-header">Taburan Aset Mengikut Daerah Sivil</div>', unsafe_allow_html=True)
        kiraan_sivil = df_filtered["Daerah Sivil"].value_counts().reset_index()
        kiraan_sivil.columns = ['Daerah Sivil', 'Jumlah']
        if not kiraan_sivil.empty and total_aset > 0:
            fig_bar = px.bar(kiraan_sivil, x='Daerah Sivil', y='Jumlah', text='Jumlah', color='Daerah Sivil', template='plotly_white')
            fig_bar.update_traces(textposition='outside')
            fig_bar.update_layout(showlegend=False, xaxis_title=None, yaxis_title=None, margin=dict(t=10, b=10, l=0, r=0))
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Tiada data untuk dipaparkan.")

    with col2:
        st.markdown('<div class="section-header">Pecahan Status Kondisi Keseluruhan</div>', unsafe_allow_html=True)
        kiraan_status = df_filtered["Status_Bersih"].value_counts().reset_index()
        kiraan_status.columns = ['Status Kondisi', 'Jumlah']
        if not kiraan_status.empty and total_aset > 0:
            fig_status = px.bar(kiraan_status, x='Jumlah', y='Status Kondisi', orientation='h', text='Jumlah', color='Status Kondisi', template='plotly_white')
            fig_status.update_traces(textposition='outside')
            fig_status.update_layout(showlegend=False, yaxis={'categoryorder':'total ascending'}, xaxis_title=None, yaxis_title=None, margin=dict(t=10, b=10, l=0, r=0))
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Tiada data untuk dipaparkan.")

    # 6. JADUAL PERINCIAN
    st.markdown('<div class="section-header">Log Perincian Pangkalan Data Aset</div>', unsafe_allow_html=True)
    
    lajur_paparan = ['Nama_Fasiliti', 'Kategori_Fasiliti', 'Lokasi', 'Daerah Sivil', 'Daerah Pentadbiran', 'Status_Bersih', 'Jenis_Building', 'Pautan_Peta']
    # Membetulkan ralat nama kekunci lajur jenis bangunan pangkalan data sedia ada
    if 'Jenis_Bangunan' in df_filtered.columns:
        lajur_paparan = ['Nama_Fasiliti', 'Kategori_Fasiliti', 'Lokasi', 'Daerah Sivil', 'Daerah Pentadbiran', 'Status_Bersih', 'Jenis_Bangunan', 'Pautan_Peta']
        
    lajur_paparan = [kolum for kolum in lajur_paparan if kolum in df_filtered.columns]
    
    if total_aset > 0:
        df_display = df_filtered[lajur_paparan].copy()
        df_display = df_display.astype(str).replace(['nan', 'None', '<NA>', ''], '-')
        if 'Pautan_Peta' in df_display.columns:
            df_display['Pautan_Peta'] = df_display['Pautan_Peta'].replace('-', None)
            
        df_display.insert(0, 'Bil', range(1, len(df_display) + 1))
        
        st.dataframe(
            df_display,
            column_config={
                "Bil": st.column_config.NumberColumn("Bil.", width="small"),
                "Nama_Fasiliti": st.column_config.TextColumn("Nama Fasiliti / Aset"),
                "Kategori_Fasiliti": st.column_config.TextColumn("Kategori"),
                "Daerah Sivil": st.column_config.TextColumn("Daerah Sivil"),
                "Daerah Pentadbiran": st.column_config.TextColumn("Pentadbiran"),
                "Status_Bersih": st.column_config.TextColumn("Status Kondisi"),
                "Jenis_Bangunan": st.column_config.TextColumn("Jenis Bangunan"),
                "Pautan_Peta": st.column_config.LinkColumn("Peta Lokasi", display_text="Buka Peta")
            },
            use_container_width=True, hide_index=True
        )
    else:
        st.warning("Rekod data tidak lengkap untuk paparan jadual ini.")

    # --- GALERI ALBUM ---
    st.markdown('<div class="section-header">Galeri Pemerhatian Aset</div>', unsafe_allow_html=True)
    
    if total_aset > 0:
        senarai_nama_aset = sorted(df_filtered['Nama_Fasiliti'].unique().tolist())
        aset_dipilih = st.selectbox("Sila pilih aset dari senarai untuk paparan gambar:", options=senarai_nama_aset)
        
        df_aset_tunggal = df_filtered[df_filtered['Nama_Fasiliti'] == aset_dipilih]
        if not df_aset_tunggal.empty:
            senarai_imej = []
            for _, r in df_aset_tunggal.iterrows():
                if isinstance(r['Senarai_Imej'], list):
                    senarai_imej.extend(r['Senarai_Imej'])
            
            senarai_imej = list(dict.fromkeys(senarai_imej))
            
            if len(senarai_imej) > 0:
                num_images = len(senarai_imej)
                cols = st.columns(min(num_images, 3))
                for idx, img_url in enumerate(senarai_imej):
                    col_idx = idx % 3
                    with cols[col_idx]:
                        st.image(img_url, caption=f"Imej {idx+1}: {aset_dipilih}", use_container_width=True)
            else:
                st.info("Tiada rekod imej digital dijumpai untuk fasiliti ini di dalam pangkalan data.")
    else:
        st.info("Pilih sekurang-kurangnya satu aset untuk paparan galeri.")

else:
    st.error("Perhatian: Pangkalan data kosong. Sila pastikan fail Excel telah dikemaskini dengan betul.")