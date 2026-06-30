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

# FUNGSI PINTAR: Memecah & memproses berbilang pautan Google Drive yang dipisahkan dengan koma
def proses_multilink_drive(val):
    val_str = str(val).strip()
    if val_str == "-" or val_str == "" or val_str.lower() == "nan":
        return []
    
    # Pecahkan mengikut tanda koma jika pengguna letak banyak link
    links = [l.strip() for l in val_str.split(",")]
    processed_links = []
    for url in links:
        if "drive.google.com" in url:
            try:
                file_id = None
                if "/file/d/" in url:
                    file_id = url.split("/file/d/")[1].split("/")[0]
                elif "id=" in url:
                    file_id = url.split("id=")[1].split("&")[0]
                
                if file_id:
                    processed_links.append(f"https://lh3.googleusercontent.com/d/{file_id}")
            except:
                pass
    return processed_links

# 2. Fungsi Membaca Data
def load_all_sheets():
    try:
        xls = pd.ExcelFile('data.xlsx')
        sheet_dict = {}
        for sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, skiprows=3)
            df = df.dropna(how='all')
            df = df.fillna("-")
            
            # Bersihkan nama lajur asal Excel
            df.columns = [str(c).strip() for c in df.columns]
            
            daerah_col = None
            status_col = None
            gps_col = None
            link_gambar_col = None
            gambar_perkara_col = None
            lokasi_col = None

            for c in df.columns:
                c_low = c.lower()
                if c_low == 'daerah':
                    daerah_col = c
                elif 'status' in c_low and 'kefungsian' in c_low:
                    status_col = c
                elif 'gps' in c_low or 'kedudukan' in c_low:
                    gps_col = c
                elif c_low == 'gambar': 
                    link_gambar_col = c
                elif 'gambar' in c_low and 'perkara' in c_low: 
                    gambar_perkara_col = c
                elif 'lokasi' in c_low:
                    lokasi_col = c

            if not daerah_col:
                for c in df.columns:
                    if 'daerah' in c.lower():
                        daerah_col = c
                        break

            new_df = pd.DataFrame()
            if daerah_col:
                new_df['Daerah'] = df[daerah_col].astype(str).str.strip().str.title()
                new_df = new_df[(new_df['Daerah'] != "-") & (new_df['Daerah'] != "") & (~new_df['Daerah'].str.lower().isin(["nan", "none"]))]
            else:
                continue
                
            valid_indices = new_df.index
            
            if gambar_perkara_col:
                new_df['Gambar / Perkara'] = df.loc[valid_indices, gambar_perkara_col].astype(str)
            else:
                new_df['Gambar / Perkara'] = "-"
                
            if lokasi_col:
                new_df['Lokasi'] = df.loc[valid_indices, lokasi_col].astype(str)
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

            # Simpan senarai semua pautan imej yang telah diproses dalam bentuk list []
            if link_gambar_col:
                new_df['Senarai_Imej'] = df.loc[valid_indices, link_gambar_col].apply(proses_multilink_drive)
            else:
                new_df['Senarai_Imej'] = [[] for _ in range(len(new_df))]

            if not new_df.empty:
                sheet_dict[sheet] = new_df
                
        return sheet_dict
    except Exception as e:
        st.error(f"Ralat teknikal pembacaan Excel: {e}")
        return None

all_data = load_all_sheets()

if all_data and len(all_data) > 0:
    # 3. SIDEBAR: PILIHAN KAWASAN
    st.sidebar.header("🔍 Tapisan Utama")
    
    senarai_kawasan = list(all_data.keys())
    kawasan_terpilih = st.sidebar.selectbox("📂 Pilih Kawasan (Sheet):", options=senarai_kawasan)
    df_kawasan = all_data[kawasan_terpilih]
    
    senarai_daerah = sorted(df_kawasan["Daerah"].unique().tolist())
    pilihan_daerah = ["✨ SEMUA DAERAH"] + senarai_daerah
    daerah_terpilih = st.sidebar.selectbox("📍 Pilih Daerah:", options=pilihan_daerah)
    
    if daerah_terpilih == "✨ SEMUA DAERAH":
        df_filtered = df_kawasan.copy()
    else:
        df_filtered = df_kawasan[df_kawasan["Daerah"] == daerah_terpilih].copy()

    # 4. Ringkasan Eksekutif
    total_aset = len(df_filtered)
    aset_baik = df_filtered['Status_Bersih'].str.contains('Baik', case=False, na=False).sum()
    
    m1, m2, m3 = st.columns(3)
    with m1: st.metric(label="📊 Jumlah Aset", value=total_aset)
    with m2: st.metric(label="📂 Kawasan Aktif", value=kawasan_terpilih)
    with m3: st.metric(label="✅ Kondisi Baik", value=aset_baik)
        
    st.markdown("---")

    # 5. Visualisasi
    col1, col2 = st.columns([1, 1])
    with col1:
        st.markdown(f"### 📈 Statistik Daerah: {kawasan_terpilih}")
        kiraan_daerah = df_filtered["Daerah"].value_counts().reset_index()
        kiraan_daerah.columns = ['Daerah', 'Jumlah']
        if not kiraan_daerah.empty and total_aset > 0:
            fig_bar = px.bar(kiraan_daerah, x='Daerah', y='Jumlah', text='Jumlah', color='Daerah', color_discrete_sequence=px.colors.sequential.Greens_r, template='plotly_white')
            fig_bar.update_traces(textposition='outside')
            st.plotly_chart(fig_bar, use_container_width=True)
        else:
            st.info("Tiada statistik grafik dipaparkan.")

    with col2:
        st.markdown("### 📊 Status Kondisi Aset")
        kiraan_status = df_filtered["Status_Bersih"].value_counts().reset_index()
        kiraan_status.columns = ['Status Kondisi', 'Jumlah']
        if not kiraan_status.empty and total_aset > 0:
            fig_status = px.bar(kiraan_status, x='Jumlah', y='Status Kondisi', orientation='h', text='Jumlah', color='Status Kondisi', color_discrete_sequence=px.colors.sequential.YlGnBu_r, template='plotly_white')
            fig_status.update_layout(yaxis={'categoryorder':'total ascending'}, showlegend=False)
            st.plotly_chart(fig_status, use_container_width=True)
        else:
            st.info("Tiada statistik status dipaparkan.")

    # 6. Jadual Perincian
    st.markdown(f"### 📋 Senarai Aset: {kawasan_terpilih}")
    
    tab1, tab2 = st.tabs(["🗂️ Jadual Pangkalan Data", "🏢 Ringkasan Jenis"])
    
    with tab1:
        # KITA KELUARKAN LAJUR GAMBAR DARI JADUAL UTAMA SUPAYA JADUAL KEMBALI KEMAS & TIDAK MEMANJANG
        lajur_paparan = ['Gambar / Perkara', 'Lokasi', 'Daerah', 'Status_Bersih', 'Jenis_Bangunan', 'Pautan_Peta']
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
                    "Pautan_Peta": st.column_config.LinkColumn("🗺️ Peta", display_text="Buka Peta 📍")
                },
                use_container_width=True, hide_index=True
            )
        else:
            st.warning("Tiada rekod data dijumpai untuk paparan jadual ini.")
        
    with tab2:
        if total_aset > 0:
            st.table(df_filtered["Jenis_Bangunan"].value_counts().reset_index().rename(columns={'index':'Jenis', 'Jenis_Bangunan':'Jumlah'}))
        else:
            st.info("Tiada ringkasan kategori.")

    # --- 🌟 CARA 2: BAHAGIAN GALERI ALBUM FOTO PREMIUM DI BAWAH JADUAL ---
    st.markdown("---")
    st.markdown("### 📸 Galeri Album Foto Aset Pilihan")
    st.markdown("*Sila pilih nama bangunan di bawah untuk melihat koleksi penuh gambar dalam saiz besar dan jelas.*")
    
    if total_aset > 0:
        senarai_nama_aset = sorted(df_filtered['Gambar / Perkara'].unique().tolist())
        aset_dipilih = st.selectbox("🎯 Pilih Bangunan / Perkara Kaji:", options=senarai_nama_aset)
        
        # Ambil baris data bagi aset yang dipilih sahaja
        df_aset_tunggal = df_filtered[df_filtered['Gambar / Perkara'] == aset_dipilih]
        
        if not df_aset_tunggal.empty:
            senarai_imej = df_aset_tunggal.iloc[0]['Senarai_Imej']
            
            if senarai_imej and len(senarai_imej) > 0:
                num_images = len(senarai_imej)
                # Maksimum susunan 3 gambar sebaris untuk susunan grid interaktif yang premium
                cols = st.columns(min(num_images, 3))
                
                for idx, img_url in enumerate(senarai_imej):
                    col_idx = idx % 3
                    with cols[col_idx]:
                        st.image(img_url, caption=f"Pandangan {idx+1}: {aset_dipilih}", use_container_width=True)
            else:
                st.info("ℹ️ Tiada fail imej dimasukkan untuk aset ini di dalam Excel atau pautan Google Drive tidak sah.")
    else:
        st.info("Tiada aset tersedia untuk paparan galeri.")

else:
    st.error("Ralat Utama: Gagal membaca lembaran data (sheets) daripada Excel. Pastikan ia telah ditutup daripada perisian Microsoft Excel.")