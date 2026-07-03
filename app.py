import streamlit as st
import pandas as pd
import plotly.express as px
import os
import glob
import re

# Konfigurasi Halaman Dashboard (Wide Layout)
st.set_page_config(page_title="Dashboard Aset JPNS", page_icon="🌲", layout="wide")

current_dir = os.getcwd()

# --- REKABENTUK HEADER & PENGESAN LOGO AUTOMATIK ---
col1, col2 = st.columns([1, 6])
with col1:
    senarai_logo = glob.glob(os.path.join(current_dir, "*[L|l][O|o][G|g][O|o]*.*"))
    logo_path = None
    for fail in senarai_logo:
        if fail.lower().endswith(('.png', '.jpg', '.jpeg')):
            logo_path = fail
            break
            
    if logo_path:
        st.image(logo_path, width=110)
    else:
        st.image("https://upload.wikimedia.org/wikipedia/commons/thumb/c/cd/Jata_Negara_Malaysia.png/200px-Jata_Negara_Malaysia.png", width=100)

with col2:
    st.title("Sistem Pengurusan Aset Bangunan")
    st.markdown("**JABATAN PERHUTANAN NEGERI SEMBILAN (JPNS)**")

st.markdown("<hr style='border:1px solid #4CAF50'>", unsafe_allow_html=True)

# --- FUNGSI BACA DATA KHAS JPNS (Logik Block_ID) ---
@st.cache_data
def load_data():
    df_list = []
    
    def process_df(df, nama_kategori):
        df.columns = df.columns.astype(str).str.strip()
        
        for col in df.columns:
            if 'perkara' in col.lower():
                df.rename(columns={col: 'Perkara'}, inplace=True)
                break
                
        if 'Perkara' not in df.columns:
            return pd.DataFrame()

        df = df[~df['Perkara'].astype(str).str.contains('Perkara', case=False, na=False)]
        df['Perkara'] = df['Perkara'].replace(r'^\s*$', pd.NA, regex=True)
        
        df['Block_ID'] = df['Perkara'].notna().cumsum()
        df = df[df['Block_ID'] > 0]
        
        kumpul_data = []
        for block_id, group in df.groupby('Block_ID'):
            baris_utama = group.iloc[0].copy()
            
            semua_teks = ' '.join(group.fillna('').astype(str).values.flatten())
            urls = re.findall(r'(https?://\S+)', semua_teks)
            urls = [u.strip('",\'') for u in urls]
            
            drive_links = [u for u in urls if 'drive.google' in u]
            maps_links = [u for u in urls if 'maps' in u or 'googleusercontent' in u]
            
            baris_utama['Pautan Gambar'] = ", ".join(list(set(drive_links))) if drive_links else None
            baris_utama['Pautan Maps'] = maps_links[0] if maps_links else None
            
            kumpul_data.append(baris_utama)
            
        df_bersih = pd.DataFrame(kumpul_data)
        df_bersih['Kategori_Aset'] = nama_kategori
        return df_bersih

    # Baca CSV
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

    # Baca Excel
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

    if df_list:
        return pd.concat(df_list, ignore_index=True)
    return pd.DataFrame()

df = load_data()

if not df.empty:
    # --- BAHAGIAN SIDEBAR ---
    st.sidebar.header("🏛️ Kategori & Pentadbiran")
    
    daerah_col = None
    for col in df.columns:
        if 'pentadbiran' in col.lower() or 'daerah' in col.lower():
            daerah_col = col
            if 'pentadbiran' in col.lower():
                break 
                
    if daerah_col:
        senarai_daerah = ["Semua"] + list(df[daerah_col].dropna().unique())
        pilihan_daerah = st.sidebar.selectbox("Pilih Daerah Pentadbiran:", senarai_daerah)
    else:
        pilihan_daerah = "Semua"

    senarai_kategori = ["Semua"] + list(df['Kategori_Aset'].dropna().unique())
    pilihan_kategori = st.sidebar.selectbox("Pilih Kategori Aset:", senarai_kategori)
    
    df_tapis = df.copy()
    if pilihan_daerah != "Semua" and daerah_col:
        df_tapis = df_tapis[df_tapis[daerah_col] == pilihan_daerah]
    if pilihan_kategori != "Semua":
        df_tapis = df_tapis[df_tapis['Kategori_Aset'] == pilihan_kategori]
        
    # --- BAHAGIAN RINGKASAN METRIK (KPI) ---
    st.subheader("📊 Ringkasan Prestasi Aset Semasa")
    m1, m2, m3 = st.columns(3)
    jumlah_aset = len(df_tapis)
    
    status_col = next((col for col in df_tapis.columns if 'status' in col.lower() or 'kefungsian' in col.lower()), None)
    if status_col:
        aset_rosak = len(df_tapis[df_tapis[status_col].astype(str).str.contains('Rosak|Perlu', case=False, na=False)])
        aset_baik = len(df_tapis[df_tapis[status_col].astype(str).str.contains('Baik|Guna|Aktif', case=False, na=False)])
    else:
        aset_rosak = 0
        aset_baik = 0
    
    m1.metric("Jumlah Keseluruhan Aset", f"{jumlah_aset} Unit")
    m2.metric("🟢 Aset Keadaan Baik", f"{aset_baik} Unit")
    m3.metric("🔴 Aset Rosak/Senggara", f"{aset_rosak} Unit")
    
    st.markdown("---")
    
    # --- BAHAGIAN CARTA (VISUALISASI EXPRES) ---
    c1, c2 = st.columns(2)
    with c1:
        st.write("**Taburan Aset Mengikut Daerah Pentadbiran**")
        if daerah_col and not df_tapis[daerah_col].dropna().empty:
            df_daerah_chart = df_tapis[daerah_col].value_counts().reset_index()
            df_daerah_chart.columns = ['Daerah', 'Bilangan']
            fig_bar = px.bar(df_daerah_chart, x='Daerah', y='Bilangan', color='Daerah', text_auto=True)
            st.plotly_chart(fig_bar, use_container_width=True)
            
    with c2:
        st.write("**Pecahan Mengikut Kategori**")
        if not df_tapis['Kategori_Aset'].dropna().empty:
            df_kat = df_tapis['Kategori_Aset'].value_counts().reset_index()
            df_kat.columns = ['Kategori', 'Bilangan']
            fig_pie = px.pie(df_kat, names='Kategori', values='Bilangan', hole=0.3, color_discrete_sequence=px.colors.qualitative.Pastel)
            st.plotly_chart(fig_pie, use_container_width=True)

    st.markdown("---")

    # --- BAHAGIAN JADUAL TERPERINCI (LINKCOLUMN FOR MAPS) ---
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

    # --- DROPDOWN ASET & PREMANENT LOOK GALERI GAMBAR (2 LAJUR BERKEMBAR) ---
    st.subheader("🖼️ Galeri Gambar Semakan Aset")
    
    df_gambar = df_tapis.dropna(subset=['Pautan Gambar', 'Perkara'])
    if not df_gambar.empty:
        pilihan_aset = st.selectbox("Sila pilih aset untuk meneliti struktur bangunan:", df_gambar['Perkara'].unique())
        
        if pilihan_aset:
            links_str = df_gambar[df_gambar['Perkara'] == pilihan_aset]['Pautan Gambar'].iloc[0]
            links_list = [link.strip() for link in str(links_str).split(",") if 'http' in link]
            
            if links_list:
                # Memproses susunan gambar ke dalam grid 2-lajur berkembar secara bersih
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
                                
                                # HTML PREMIUM LOOK: Klik terus buka imej penuh/bersih di tab baharu tanpa iframe
                                st.markdown(
                                    f'<div style="text-align: left; margin-bottom: 6px; color: #333;"><b>📸 Pandangan Gambar {posisi_asal+1}</b></div>'
                                    f'<a href="{bypass_view_url}" target="_blank">'
                                    f'<img src="{bypass_view_url}" style="width:100%; height:320px; object-fit:cover; border-radius:8px; border:1px solid #e0e0e0; transition: transform .2s;">'
                                    f'</a>',
                                    unsafe_allow_html=True
                                )
                            else:
                                st.markdown(f"🔗 [Pautan Manual Gambar {posisi_asal+1}]({link_terpilih})")
            else:
                st.info("Tiada imej yang sah ditemui bagi pautan yang didaftarkan.")
    else:
        st.info("Tiada data pautan gambar Google Drive untuk aset di bawah penapisan semasa.")

else:
    st.error("Sistem gagal memuat naik data. Sila pastikan fail data.xlsx atau CSV diletakkan dalam direktori yang betul.")