import streamlit as st
import pandas as pd
import os

# Sayfa Yapılandırması
st.set_page_config(
    page_title="PLANLAMA - Arama Uygulaması",
    page_icon="📍",
    layout="wide"
)

EXCEL_FILE = "veri.xlsx"

@st.cache_data
def load_data(path):
    """Excel verilerini güvenli şekilde önbelleğe alarak yükler."""
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
        st.error(f"Excel okunurken hata oluştu: {e}")
        return None, None

df_mahalle, df_madde = load_data(EXCEL_FILE)

if df_mahalle is None or df_madde is None:
    st.error(f"'{EXCEL_FILE}' dosyası bulunamadı veya 'MAHALLE' / 'MADDE' sayfaları eksik!")
    st.stop()

# Türkçe karakter uyumlu büyük harfe çevirme
def tr_upper(text):
    if not isinstance(text, str):
        text = str(text) if pd.notna(text) else ""
    return text.replace("i", "İ").replace("ı", "I").upper()

# --- ARAYÜZ ---
st.markdown("<h2 style='text-align: center; color: #1C3B68;'>MAHALLE-KOLLUK ve MADDE-SUÇ ARAMA</h2>", unsafe_allow_html=True)
st.markdown("---")

# Mod Seçimi (Buton mantığı yerine şık bir radio buton)
mode = st.radio(
    "Arama Modunu Seçin:",
    ["📍 Mahalle ve Kolluk Birimleri", "⚖️ Madde ve Suç Tanımları"],
    horizontal=True
)

# Arama Kutusu
search_text = st.text_input("🔍 Aramak istediğiniz metni girin...", "").strip()
search_text_upper = tr_upper(search_text)

# --- MAHALLE MODU ---
if "Mahalle" in mode:
    filtered_rows = []
    for _, row in df_mahalle.iterrows():
        val_mahalle = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
        val_kolluk = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])
        val_ilce = "" if len(row) <= 3 or pd.isna(row.iloc[3]) else str(row.iloc[3])

        mahalle_val_upper = tr_upper(val_mahalle)

        if not search_text_upper or search_text_upper in mahalle_val_upper:
            filtered_rows.append({
                "MAHALLE": val_mahalle,
                "BAĞLI BULUNDUĞU KOLLUK BİRİMİ": val_kolluk,
                "İLÇE": val_ilce
            })
    
    df_display = pd.DataFrame(filtered_rows)
    st.dataframe(df_display, use_container_width=True, hide_index=True)

# --- MADDE MODU ---
else:
    filtered_rows = []
    for _, row in df_madde.iterrows():
        full_row_text = " ".join([tr_upper(x) for x in row.values])

        if not search_text_upper or search_text_upper in full_row_text:
            val_madde = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            val_suc = "" if len(row) <= 1 or pd.isna(row.iloc[1]) else str(row.iloc[1])
            val_kanun = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])

            filtered_rows.append({
                "MADDE NO": val_madde,
                "SUÇ TANIMI": val_suc,
                "KANUN": val_kanun
            })
    
    df_display = pd.DataFrame(filtered_rows)
    st.dataframe(df_display, use_container_width=True, hide_index=True)