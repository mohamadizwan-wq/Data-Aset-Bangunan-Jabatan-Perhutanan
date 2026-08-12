import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import re
from datetime import datetime, timedelta, timezone

# ==========================================
# 1. KONFIGURASI HALAMAN & FAVICON (LOGO TAB BROWSER)
# ==========================================
current_dir = os.getcwd()

# Cari fail logo dalam folder Github untuk dijadikan ikon tab browser
senarai_logo_favicon = glob.glob(os.path.join(current_dir, "*[L|l][O|o][G|g][O|o]*.*"))
favicon_path = "https://upload.wikimedia.org/wikipedia/commons/thumb/c/cd/Jata_Negara_Malaysia.png/200px-Jata_Negara_Malaysia.png"

for fail in senarai_logo_favicon:
    if fail.lower().endswith(('.png', '.jpg', '.jpeg', '.ico')):
        favicon_path = fail
        break

st.set_page_config(
    page_title="Dashboard Aset JPNS", 
    page_icon=favicon_path, 
    layout="wide"
)

# ==========================================
# CSS PREMIUM: SOROK IKON SEMAK DI ATAS KANAN
# ==========================================
st.markdown(
    """
    <style>
    .stActionButton, div[data-testid="stHeaderActionElements"] {
        display: none !important;
    }
    header {
        background-color: rgba(0,0,0,0) !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ==========================================
# 2. REKABENTUK HEADER & PENGESAN LOGO AUTOMATIK
# ==========================================
col1, col2 = st.columns([1, 6])
with col1:
    if favicon_path and os.path.exists(favicon_path):
        st.image(favicon_path, width=110)
    else:
        st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/cd/Jata_Negara_Malaysia.png/200px-Jata_Negara_Malaysia.png", width=100)

with col2:
    st.title("Sistem Pengurusan Aset Bangunan")
    st.markdown("**JABATAN PERHUTANAN NEGERI SEMBILAN (JPNS)**")

st.markdown("<hr style='border:1px solid #4CAF50'>", unsafe_allow_html=True)


# ==========================================
# 3. FUNGSI BACA DATA KHAS JPNS (LOGIK BLOCK_ID)
# ==========================================
@st.cache_data(ttl=5)
def load_data():
    df_list = []
    
    def process_df(df, nama_kategori):
        df.columns = df.columns.astype(str).str.strip()
        
        perkara_col = next((col for col in df.columns if 'perkara' in col.lower()), None)
        if perkara_col:
            df.rename(columns={perkara_col: 'Perkara'}, inplace=True)
        else:
            return pd.DataFrame()

        df = df[~df['Perkara'].astype(str).str.contains('Perkara', case=False, na=False)]
        df['Perkara'] = df['Perkara'].replace(r'^\s*$', pd.NA, regex=True)
        
        df['Block_ID'] = df['Perkara'].notna().cumsum()
        df = df[df['Block_ID'] > 0]
        
        kumpul_data = []
        for block_id, group in df.groupby('Block_ID'):
            baris_utama = group.iloc[0].copy()
            
            for col in group.columns:
                if col not in ['Block_ID', 'Perkara']:
                    values = group[col].dropna().astype(str).str.strip()
                    values = [v for v in values if v != '' and v.lower() != 'nan']
                    if values:
                        baris_utama[col] = " | ".join(list(dict.fromkeys(values)))
            
            semua_teks = ' '.join(group.fillna('').astype(str).values.flatten())
            urls = re.findall(r'(https?://\S+)', semua_teks)
            urls = [u.strip('",\'') for u in urls]
            
            drive_links = [u for u in urls if 'drive.google' in u]
            maps_links = [u for u in urls if 'maps' in u or 'googleusercontent' in u]
            
            baris_utama['Pautan Gambar'] = ", ".join(list(dict.fromkeys(drive_links))) if drive_links else None
            baris_utama['Pautan Maps'] = maps_links[0] if maps_links else None
            
            kumpul_data.append(baris_utama)
            
        df_bersih = pd.DataFrame(kumpul_data)
        df_bersih['Kategori_Aset'] = nama_kategori
        return df_bersih

    # 1. Baca Excel Dulu
    fail_excel_lain = glob.glob(os.path.join(current_dir, "*.xlsx"))
    for fail_excel in fail_excel_lain:
        if '~$' in fail_excel: continue
        try:
            xls = pd.ExcelFile(fail_excel, engine='openpyxl')
            for sheet in xls.sheet_names:
                temp_df = pd.read_excel(fail_excel, sheet_name=sheet, header=None, dtype=str)
                header_idx = -1
                for i, row in temp_df.iterrows():
                    if 'perkara' in ' '.join(row.fillna('').astype(str)).lower():
                        header_idx = i
                        break
                if header_idx != -1:
                    df = pd.read_excel(fail_excel, sheet_name=sheet, skiprows=header_idx)
                    df_bersih = process_df(df, sheet)
                    if not df_bersih.empty:
                        df_list.append(df_bersih)
        except:
            pass

    # 2. Baca CSV jika tiada Excel
    if not df_list:
        fail_csv_lain = glob.glob(os.path.join(current_dir, "*.csv"))
        for fail_csv in fail_csv_lain:
            for bahasa in ['utf-8', 'cp1252', 'latin1']:
                try:
                    temp_df = pd.read_csv(fail_csv, header=None, encoding=bahasa, dtype=str, on_bad_lines='skip')
                    header_idx = -1
                    for i, row in temp_df.iterrows():
                        if 'perkara' in ' '.join(row.fillna('').astype(str)).lower():
                            header_idx = i
                            break
                    if header_idx != -1:
                        df = pd.read_csv(fail_csv, skiprows=header_idx, encoding=bahasa, on_bad_lines='skip')
                        nama_kat = os.path.basename(fail_csv).replace("data.xlsx - ", "").replace(".csv", "")
                        df_bersih = process_df(df, nama_kat)
                        if not df_bersih.empty:
                            df_list.append(df_bersih)
                        break 
                except:
                    pass

    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return pd.DataFrame()

df = load_data()


# ==========================================
# 4. PEMBINAAN ELEMEN DASHBOARD
# ==========================================
if not df.empty:
    
    # --- BAHAGIAN SIDEBAR (LOGIK CASCADING DROPDOWN) ---
    st.sidebar.header("🏛️ Kategori & Pentadbiran")
    
    daerah_col = next((col for col in df.columns if 'pentadbiran' in col.lower()), None)
    daerah_sivil_col = next((col for col in df.columns if 'sivil' in col.lower()), None)
    
    if not daerah_col:
        daerah_col = next((col for col in df.columns if 'daerah' in col.lower()), None)
    
    if daerah_col:
        senarai_daerah = ["Semua"] + list(df[daerah_col].dropna().unique())
        pilihan_daerah = st.sidebar.selectbox("Pilih Daerah Pentadbiran:", senarai_daerah)
    else:
        pilihan_daerah = "Semua"

    if pilihan_daerah != "Semua" and daerah_col:
        df_untuk_kategori = df[df[daerah_col] == pilihan_daerah]
    else:
        df_untuk_kategori = df

    senarai_kategori = ["Semua"] + list(df_untuk_kategori['Kategori_Aset'].dropna().unique())
    pilihan_kategori = st.sidebar.selectbox("Pilih Kategori Aset:", senarai_kategori)
    
    df_tapis = df.copy()
    if pilihan_daerah != "Semua" and daerah_col:
        df_tapis = df_tapis[df_tapis[daerah_col] == pilihan_daerah]
    if pilihan_kategori != "Semua":
        df_tapis = df_tapis[df_tapis['Kategori_Aset'] == pilihan_kategori]
        
    # --- BAHAGIAN RINGKASAN METRIK (KPI) ---
    st.subheader("📊 Ringkasan Prestasi Aset Semasa")
    
    status_col = next((col for col in df_tapis.columns if 'status' in col.lower() or 'kefungsian' in col.lower()), None)
    
    # LOGIK PENAPISAN STATUS YANG DAH DIBAIKI KETAT & TEPAT
    kriteria_rosak_regex = r'Rosak|Selenggara|Proses|Penyelenggaraan|Penambahbaikan'
    
    jumlah_aset_kpi = len(df_tapis)
    if status_col:
        s_series = df_tapis[status_col].astype(str)
        mask_rosak = s_series.str.contains(kriteria_rosak_regex, case=False, na=False)
        mask_baik = (s_series.str.contains(r'Baik|Aktif', case=False, na=False)) & (~mask_rosak)
        
        aset_baik_kpi = len(df_tapis[mask_baik])
        aset_rosak_kpi = len(df_tapis[mask_rosak])
    else:
        aset_baik_kpi = 0
        aset_rosak_kpi = 0

    m1, m2, m3 = st.columns(3)
    m1.metric("Jumlah Keseluruhan Aset", f"{jumlah_aset_kpi} Unit")
    m2.metric("🟢 Aset Berkeadaan Baik", f"{aset_baik_kpi} Unit")
    m3.metric("🔴 Aset Rosak/Selenggara", f"{aset_rosak_kpi} Unit")
    
    st.markdown("### 🔍 Tapisan Pantas Kondisi Aset")
    pilihan_status = st.radio(
        "Pilih untuk mengecilkan skop senarai jadual dan paparan gambar di bawah:",
        ["Papar Semua Aset", "🟢 Aset Berkeadaan Baik", "🔴 Aset Rosak/Selenggara"],
        horizontal=True
    )
    
    if pilihan_status == "🟢 Aset Berkeadaan Baik" and status_col:
        df_tapis = df_tapis[mask_baik]
    elif pilihan_status == "🔴 Aset Rosak/Selenggara" and status_col:
        df_tapis = df_tapis[mask_rosak]

    st.markdown("---")
    
    # --- BAHAGIAN VISUALISASI CARTA ---
    st.subheader("📈 Analisis Visual Data Aset")
    
    c1, c2 = st.columns(2)
    with c1:
        st.write("**Taburan Aset Mengikut Daerah Pentadbiran (Perhutanan)**")
        if daerah_col and not df_tapis[daerah_col].dropna().empty:
            df_daerah_chart = df_tapis[daerah_col].value_counts().reset_index()
            df_daerah_chart.columns = ['Daerah Pentadbiran', 'Bilangan']
            fig_bar1 = px.bar(df_daerah_chart, x='Daerah Pentadbiran', y='Bilangan', color='Daerah Pentadbiran', text_auto=True)
            st.plotly_chart(fig_bar1, use_container_width=True)
        else:
            st.info("Data Daerah Pentadbiran tidak ditemui.")
            
    with c2:
        st.write("**Taburan Aset Mengikut Daerah Sivil (Negeri)**")
        if daerah_sivil_col and not df_tapis[daerah_sivil_col].dropna().empty:
            df_sivil_chart = df_tapis[daerah_sivil_col].value_counts().reset_index()
            df_sivil_chart.columns = ['Daerah Sivil', 'Bilangan']
            fig_bar2 = px.bar(df_sivil_chart, x='Daerah Sivil', y='Bilangan', color='Daerah Sivil', text_auto=True, color_discrete_sequence=px.colors.qualitative.Dark2)
            st.plotly_chart(fig_bar2, use_container_width=True)
        else:
            st.info("Lajur 'Daerah Sivil' tidak ditemui dalam fail data Excel.")

    st.write("**Pecahan Keseluruhan Mengikut Kategori Aset**")
    if not df_tapis['Kategori_Aset'].dropna().empty:
        df_kat = df_tapis['Kategori_Aset'].value_counts().reset_index()
        df_kat.columns = ['Kategori', 'Bilangan']
        fig_pie = px.pie(df_kat, names='Kategori', values='Bilangan', hole=0.3, color_discrete_sequence=px.colors.qualitative.Pastel)
        st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")

    # --- BAHAGIAN JADUAL TERPERINCI ---
    st.subheader("📋 Senarai Terperinci Aset Bangunan")
    kolum_wujud = []
    for k in ['Kategori_Aset', 'Perkara', 'Lokasi', 'Daerah Pentadbiran', 'Daerah Sivil', 'Status', 'Pautan Maps']:
        for col in df_tapis.columns:
            if k.lower() in col.lower() and col not in kolum_wujud:
                kolum_wujud.append(col)
                
    if kolum_wujud:
        st.dataframe(
            df_tapis[kolum_wujud], 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Pautan Maps": st.column_config.LinkColumn("🌎 Google Maps (Klik Sini)")
            }
        )
    else:
        st.dataframe(df_tapis, use_container_width=True, hide_index=True)

    st.markdown("---")

    # --- DROPDOWN ASET & GALERI GAMBAR PREMIUM ---
    st.subheader("🖼️ Galeri Gambar Semakan Aset")
    df_gambar = df_tapis.dropna(subset=['Pautan Gambar', 'Perkara'])
    
    if not df_gambar.empty:
        senarai_aset_gambar = df_gambar['Perkara'].unique()
        pilihan_aset = st.selectbox("Sila pilih aset untuk meneliti struktur bangunan:", senarai_aset_gambar)
        
        if pilihan_aset:
            links_str = df_gambar[df_gambar['Perkara'] == pilihan_aset]['Pautan Gambar'].iloc[0]
            links_list = [link.strip() for link in str(links_str).split(",") if 'http' in link]
            
            if links_list:
                st.markdown("### 📸 Visualisasi Struktur Bangunan (Klik gambar untuk besarkan)")
                for i in range(0, len(links_list), 2):
                    pasangan_links = links_list[i:i+2]
                    img_cols = st.columns(2)
                    for idx, link_terpilih in enumerate(pasangan_links):
                        posisi_asal = i + idx
                        with img_cols[idx]:
                            file_id = None
                            if "/d/" in link_terpilih:
                                file_id = link_terpilih.split("/d/")[1].split("/")[0]
                            elif "id=" in link_terpilih:
                                file_id = link_terpilih.split("id=")[1].split("&")[0]
                                
                            if file_id:
                                bypass_view_url = f"https://lh3.googleusercontent.com/d/{file_id}"
                                st.markdown(
                                    f'''
                                    <div style="text-align: left; margin-bottom: 15px; color: #333;">
                                        <div style="margin-bottom: 5px;"><b>📸 Pandangan Gambar {posisi_asal+1}</b></div>
                                        <a href="{bypass_view_url}" target="_blank">
                                            <img src="{bypass_view_url}" style="width:100%; max-height:400px; object-fit:cover; border-radius:10px; border:1px solid #ddd; cursor:pointer;">
                                        </a>
                                    </div>
                                    ''',
                                    unsafe_allow_html=True
                                )
                            else:
                                st.markdown(f"🔗 [Pautan Manual Gambar {posisi_asal+1}]({link_terpilih})")
            else:
                st.info("Tiada imej yang sah ditemui bagi pautan yang didaftarkan.")
    else:
        st.info("Tiada data pautan gambar Google Drive untuk aset di bawah penapisan semasa.")

else:
    st.error("Sistem gagal muat naik data. Pastikan fail data.xlsx/CSV diletakkan dalam direktori yang betul.")


# ==============================================================================
# 6. LOGIK AUTOMATIK TARIKH KEMASKINI & FOOTER PRESTIGE (IZWAN RADZI)
# ==========================================
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown("<hr style='border:0.5px solid #e0e0e0'>", unsafe_allow_html=True)

fail_sumber = glob.glob(os.path.join(current_dir, "*.xlsx")) + glob.glob(os.path.join(current_dir, "*.csv"))
tarikh_kemaskini = "Tidak Diketahui"

if fail_sumber:
    fail_terkini = max(fail_sumber, key=os.path.getmtime)
    timestamp = os.path.getmtime(fail_terkini)
    
    waktu_lokal = datetime.fromtimestamp(timestamp, tz=timezone(timedelta(hours=8)))
    tarikh_kemaskini = waktu_lokal.strftime("%d/%m/%Y, %I:%M %p")

f1, f2 = st.columns(2)
with f1:
    st.markdown(f"⏳ **Kemaskini Data Terakhir:** `{tarikh_kemaskini} (Waktu Malaysia)`")
with f2:
    st.markdown(
        f'''
        <div style="text-align: right; color: #555; font-size: 14px;">
            💻 <b>Developer:</b> Izwan Radzi | 🌲 <b>Sistem:</b> Dashboard Bangunan JPNS v1.0
        </div>
        ''', 
        unsafe_allow_html=True
    )