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

# 1. Konfigurasi Halaman 
st.set_page_config(page_title="Sistem Dashboard Aset Bangunan JPNS", page_icon=logo_jabatan, layout="wide")

st.markdown("""
    <style>
    /* 1. Sembunyikan Ikon Kucing (GitHub) & Toolbar Atas */
    [data-testid="stToolbar"] {visibility: hidden !important;}
    
    /* 2. Sembunyikan Menu Tiga Garis (Hamburger Menu) */
    #MainMenu {visibility: hidden !important;}
    
    /* 3. Sembunyikan tulisan 'Made with Streamlit' di bawah */
    footer {visibility: hidden !important;}
    
    /* Style Tajuk Utama Anda */
    .main-title { color: #1b5e20; font-size: 38px; font-weight: bold; line-height: 1.2; }
    .sub-title { color: #4e342e; font-size: 18px; margin-top: 5px; }
    </style>
""", unsafe_allow_html=True)

# --- BAHAGIAN TAJUK DAN LOGO ---
col_logo, col_tajuk = st.columns([1, 8])
with col_logo:
    if isinstance(logo_jabatan, Image.Image):
        st.image(logo_jabatan, width=110)
    else:
        st.info("Logo")
with col_tajuk:
    st.markdown('<div class="main-title">Sistem Dashboard Interaktif Aset Bangunan</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-title">Jabatan Perhutanan Negeri Sembilan (JPNS)</div>', unsafe_allow_html=True)

st.markdown("---")

# FUNGSI PINTAR IMED: Memproses multilink drive
def proses_multilink_drive(val):
    val_str = str(val).strip()
    if val_str == "-" or val_str == "" or val_str.lower() == "nan":
        return []
    links = [l.strip() for l in val_str.split(",")]
    processed_links = []
    for url in links:
        if "drive.google.com" in url:
            try:
                file_id = None
                if "/file/d/" in url: file_id = url.split("/file/d/")[1].split("/")[0]
                elif "id=" in url: file_id = url.split("id=")[1].split("&")[0]
                if file_id:
                    processed_links.append(f"https://lh3.googleusercontent.com/d/{file_id}")
            except:
                pass
    return processed_links

# 2. Fungsi Membaca Data
def load_all_data_combined():
    try:
        xls = pd.ExcelFile('data.xlsx')
        all_sheets_list = []
        
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=3)
            df = df.dropna(how='all')
            df = df.fillna("-")
            
            # Bersihkan nama lajur asal Excel
            df.columns = [str(c).strip() for c in df.columns]
            
            daerah_sivil_col = None
            daerah_pentadbiran_col = None
            status_col = None
            gps_col = None
            gambar_perkara_col = None
            lokasi_col = None
            image_cols = []

            for c in df.columns:
                c_low = c.lower()
                if 'sivil' in c_low: 
                    daerah_sivil_col = c
                elif 'pentadbiran' in c_low: 
                    daerah_pentadbiran_col = c
                elif c_low == 'daerah' and not daerah_sivil_col: 
                    daerah_sivil_col = c
                elif 'status' in c_low and 'kefungsian' in c_low: 
                    status_col = c
                elif 'gps' in c_low or 'kedudukan' in c_low: 
                    gps_col = c
                elif 'gambar' in c_low and 'perkara' in c_low: 
                    gambar_perkara_col = c
                elif 'lokasi' in c_low: 
                    lokasi_col = c
                elif 'gambar' in c_low: 
                    image_cols.append(c)

            if not daerah_sivil_col:
                for c in df.columns:
                    if 'daerah' in c.lower() and 'pentadbiran' not in c.lower():
                        daerah_sivil_col = c
                        break

            new_df = pd.DataFrame()
            
            if gambar_perkara_col:
                new_df['Gambar / Perkara'] = df[gambar_perkara_col].astype(str).str.strip()
            else:
                continue 
                
            new_df = new_df[(new_df['Gambar / Perkara'] != "-") & (new_df['Gambar / Perkara'] != "") & (~new_df['Gambar / Perkara'].str.lower().isin(["nan", "none"]))]
            valid_indices = new_df.index
            
            if len(valid_indices) == 0:
                continue
                
            new_df['Kawasan_Sheet'] = sheet
            
            if daerah_sivil_col:
                new_df['Daerah Sivil'] = df.loc[valid_indices, daerah_sivil_col].astype(str).str.strip().str.title()
            else:
                new_df['Daerah Sivil'] = "Tidak Dinyatakan"
                
            if daerah_pentadbiran_col:
                new_df['Daerah Pentadbiran'] = df.loc[valid_indices, daerah_pentadbiran_col].astype(str).str.strip().str.upper()
            else:
                new_df['Daerah Pentadbiran'] = "Tidak Dinyatakan"
                
            if lokasi_col:
                new_df['Lokasi'] = df.loc[valid_indices, lokasi_col].astype(str).str.strip()
            else:
                new_df['Lokasi'] = "-"

            if status_col:
                status_series = df.loc[valid_indices, status_col].astype(str)
                new_df['Status_Bersih'] = status_series.apply(lambda x: x.split('/')[0].strip() if '/' in x else x)
                new_df['Jenis_Bangunan'] = status_series.apply(lambda x: x.split('/')[1].strip() if '/' in x else "-")
            else:
                new_df['Status_Bersih'] = "-"
                new_df['Jenis_Bangunan'] = "-"

            if gps_col:
                gps_series = df.loc[valid_indices, gps_col].astype(str)
                def bina_link_maps(gps):
                    gps = gps.strip()
                    if gps.lower().startswith('http'): return gps 
                    elif gps.lower() not in ['nan', 'none', '', '-']:
                        return f"https://www.google.com/maps/search/?api=1&query={urllib.parse.quote(gps)}"
                    return None
                new_df['Pautan_Peta'] = gps_series.apply(bina_link_maps)
            else:
                new_df['Pautan_Peta'] = None

            def proses_links_dari_row(row_index):
                all_processed = []
                for col in image_cols:
                    val_str = str(df.loc[row_index, col]).strip()
                    if val_str == "-" or val_str == "" or val_str.lower() == "nan":
                        continue
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

            if not new_df.empty:
                all_sheets_list.append(new_df)
                
        if len(all_sheets_list) > 0:
            combined_df = pd.concat(all_sheets_list, ignore_index=True)
            return combined_df
        return None
    except Exception as e:
        st.error(f"Ralat teknikal pembacaan Excel: {e}")
        return None

df_master = load_all_data_combined()

if df_master is not None and not df_master.empty:
    # 3. SIDEBAR: TAPISAN DAERAH PENTADBIRAN
    st.sidebar.header("🔍 Tapisan Profil JPNS")
    
    senarai_pentadbiran = sorted(df_master["Daerah Pentadbiran"].unique().tolist())
    pentadbiran_terpilih = st.sidebar.selectbox(
        "❖ Pilih Daerah Pentadbiran:",
        options=["⊞ SEMUA DAERAH PENTADBIRAN"] + senarai_pentadbiran,
        index=0
    )
    
    # Proses Logik Penapisan
    df_filtered = df_master.copy()
    if pentadbiran_terpilih != "⊞ SEMUA DAERAH PENTADBIRAN":
        df_filtered = df_filtered[df_filtered["Daerah Pentadbiran"] == pentadbiran_terpilih]

    t_pentadbiran = pentadbiran_terpilih.replace("⊞ ", "")
    tajuk_kawasan = f"Seluruh Negeri Sembilan" if "SEMUA" in t_pentadbiran else f"Pentadbiran: {t_pentadbiran}"

    # 4. Ringkasan Eksekutif Dinamik
    total_aset = len(df_filtered)
    aset_baik = df_filtered['Status_Bersih'].str.contains('Baik', case=False, na=False).sum()
    
    m1, m2, m3 = st.columns(3)
    with m1: st.metric(label="■ Jumlah Keseluruhan Aset", value=f"{total_aset} Buah")
    with m2: st.metric(label="■ Daerah Pentadbiran Dipapar", value=t_pentadbiran)
    with m3: st.metric(label="■ Kondisi Berstatus Baik", value=f"{aset_baik} Unit")
        
    st.markdown("---")

    # 5. Visualisasi Grafik Komprehensif
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown("### 📈 Taburan Aset Mengikut Daerah Sivil")
        kiraan_sivil = df_filtered["Daerah Sivil"].value_counts().reset_index()
        kiraan_sivil.columns = ['Daerah Sivil', 'Jumlah']
        if not kiraan_sivil.empty and total_aset > 0:
            fig_bar = px.bar(kiraan_sivil, x='Daerah Sivil', y='Jumlah', text='Jumlah', color='Daerah Sivil', color_discrete_sequence=px.colors.sequential.Greens_r, template='plotly_white')
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Tiada statistik grafik dipaparkan.")

    with col2:
        st.markdown("### 📊 Status Kondisi & Kefungsian Semasa")
        kiraan_status = df_filtered["Status_Bersih"].value_counts().reset_index()
        kiraan_status.columns = ['Status Kondisi', 'Jumlah']
        if not kiraan_status.empty and total_aset > 0:
            fig_status = px.bar(kiraan_status, x='Jumlah', y='Status Kondisi', orientation='h', text='Jumlah', color='Status Kondisi', color_discrete_sequence=px.colors.sequential.YlGnBu_r, template='plotly_white')
            fig_status.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Tiada statistik status dipaparkan.")

    # 6. Jadual Perincian Utama
    st.markdown(f"### 📋 Log Perincian Data: {tajuk_kawasan}")
    
    tab1, tab2 = st.tabs(["🗂️ Jadual Pangkalan Data Penuh", "🏢 Ringkasan Jenis Bangunan"])
    
    with tab1:
        lajur_paparan = ['Gambar / Perkara', 'Lokasi', 'Daerah Sivil', 'Daerah Pentadbiran', 'Status_Bersih', 'Jenis_Bangunan', 'Pautan_Peta']
        lajur_paparan = [kolum for kolum in lajur_paparan if kolum in df_filtered.columns]
        
        if total_aset > 0:
            df_display = df_filtered[lajur_paparan].copy()
            df_display = df_display.astype(str).replace(['nan', 'None', '<NA>'], '-')
            if 'Pautan_Peta' in df_display.columns:
                df_display['Pautan_Peta'] = df_display['Pautan_Peta'].replace('-', None)
                
            df_display.insert(0, 'Bil.', range(1, len(df_display) + 1))
            
            st.dataframe(
                df_display,
                column_config={
                    "Bil.": st.column_config.NumberColumn("No.", width="small"),
                    # --- DI SINI KITA KUKUHKAN NAMA PENUH TANPA UNDERSCORE ---
                    "Status_Bersih": st.column_config.TextColumn("Status Kondisi"),
                    "Jenis_Bangunan": st.column_config.TextColumn("Kategori Bangunan"),
                    "Pautan_Peta": st.column_config.LinkColumn("🗺️ Peta", display_text="Buka Peta 📍")
                },
                use_container_width=True, hide_index=True
            )
        else:
            st.warning("Tiada rekod data dijumpai untuk paparan jadual ini.")
        
    with tab2:
        if total_aset > 0:
            st.table(df_filtered["Jenis_Bangunan"].value_counts().reset_index().rename(columns={'index':'Kategori', 'Jenis_Bangunan':'Jumlah'}))
        else:
            st.info("Tiada ringkasan kategori.")

    # --- BAHAGIAN GALERI ALBUM GABUNGAN PREMIUM DI BAWAH JADUAL ---
    st.markdown("---")
    st.markdown("### 📸 Galeri Album Foto Aset Pilihan")
    st.markdown("*Sila pilih nama bangunan di bawah untuk melihat semua koleksi gambar.*")
    
    if total_aset > 0:
        senarai_nama_aset = sorted(df_filtered['Gambar / Perkara'].unique().tolist())
        aset_dipilih = st.selectbox("❖ Pilih Bangunan / Perkara Kaji:", options=senarai_nama_aset)
        
        df_aset_tunggal = df_filtered[df_filtered['Gambar / Perkara'] == aset_dipilih]
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
                        st.image(img_url, caption=f"Gambar {idx+1}: {aset_dipilih}", use_container_width=True)
            else:
                st.info("ℹ️ Tiada fail imej dimasukkan untuk aset ini di dalam Excel atau pautan Google Drive tidak sah.")
    else:
        st.info("Tiada aset tersedia untuk paparan galeri.")

else:
    st.error("Ralat Utama: Fail 'data.xlsx' kosong atau tiada data lajur 'Daerah Sivil' atau 'Daerah Pentadbiran' yang sah dijumpai.")