import streamlit as st
import pandas as pd
import os

# Sayfa Yapılandırması
st.set_page_config(
    page_title="PLANLAMA - Arama Uygulaması",
    page_icon="📍",
    layout="centered"
)

EXCEL_FILE = "veri.xlsx"

@st.cache_data
def load_data(path):
    """Excel verilerini güvenli şekilde yükler."""
    if not os.path.exists(path):
        return None, None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names
        mahalle_sheet = next((s for s in sheets if "MAHALLE" in s.upper()), None)
        madde_sheet = next((s for s in sheets if "MADDE" in s.upper()), None)

        df_mahalle = pd.read_excel(path, sheet_name=mahalle_sheet) if mahalle_sheet else pd.DataFrame()
        df_madde = pd.read_excel(path, sheet_name=madde_sheet) if madde_sheet else pd.DataFrame()
        return df_mahalle, df_madde
    except Exception as e:
        return None, None

df_mahalle, df_madde = load_data(EXCEL_FILE)

if df_mahalle is None or df_madde is None or df_mahalle.empty or df_madde.empty:
    st.error(f"'{EXCEL_FILE}' dosyası bulunamadı veya 'MAHALLE' / 'MADDE' sayfaları eksik!")
    st.stop()

def tr_upper(text):
    if not isinstance(text, str):
        text = str(text) if pd.notna(text) else ""
    return text.replace("i", "İ").replace("ı", "I").upper()

# ARAYÜZ
st.markdown("<h2 style='text-align: center; color: #1C3B68;'>MAHALLE ve MADDE ARAMA</h2>", unsafe_allow_html=True)
st.markdown("---")

mode = st.radio(
    "Arama Modunu Seçin:",
    ["📍 Mahalle ve Kolluk Birimleri", "⚖️ Madde ve Suç Tanımları"],
    horizontal=True
)

search_text = st.text_input("🔍 Aramak istediğiniz metni girin...", "").strip()
search_text_upper = tr_upper(search_text)

st.markdown("---")

# --- MAHALLE MODU (Mobil Uyumlu Kart Görünümü) ---
if "Mahalle" in mode:
    bulunan_sayi = 0
    for _, row in df_mahalle.iterrows():
        val_mahalle = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
        val_kolluk = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])
        val_ilce = "" if len(row) <= 3 or pd.isna(row.iloc[3]) else str(row.iloc[3])

        mahalle_val_upper = tr_upper(val_mahalle)

        if not search_text_upper or search_text_upper in mahalle_val_upper:
            bulunan_sayi += 1
            # Her sonucu şık bir kutu (card) içinde gösteriyoruz
            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #f8fafc; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 10px;">
                        <span style="color: #2563eb; font-weight: bold; font-size: 16px;">📍 {val_mahalle}</span><br>
                        <b style="color: #475569;">İlçe:</b> {val_ilce}<br>
                        <b style="color: #475569;">Kolluk Birimi:</b> {val_kolluk}
                    </div>
                    """,
                    unsafe_allow_html=True
                )
    if bulunan_sayi == 0:
        st.info("Aranan kriterlere uygun sonuç bulunamadı.")

# --- MADDE MODU (Mobil Uyumlu Kart Görünümü) ---
else:
    bulunan_sayi = 0
    for _, row in df_madde.iterrows():
        full_row_text = " ".join([tr_upper(x) for x in row.values])

        if not search_text_upper or search_text_upper in full_row_text:
            bulunan_sayi += 1
            val_madde = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            val_suc = "" if len(row) <= 1 or pd.isna(row.iloc[1]) else str(row.iloc[1])
            val_kanun = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])

            with st.container():
                st.markdown(
                    f"""
                    <div style="background-color: #f8fafc; padding: 12px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 10px;">
                        <span style="color: #16a34a; font-weight: bold; font-size: 16px;">⚖️ Madde {val_madde}</span> <span style="color: #64748b; font-size: 12px;">({val_kanun})</span><br>
                        <b style="color: #475569;">Suç Tanımı:</b> {val_suc}
                    </div>
                    """,
                    unsafe_allow_html=True
                )
    if bulunan_sayi == 0:
        st.info("Aranan kriterlere uygun sonuç bulunamadı.")