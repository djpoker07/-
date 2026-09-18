import streamlit as st
import pandas as pd
import os

def check_password():
  """Şifre kontrol fonksiyonu"""

  def password_entered():
    if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
      st.session_state["password_correct"] = True
      del st.session_state["password"]  # Şifreyi hafıradan sil
    else:
      st.session_state["password_correct"] = False

  # Daha önce giriş yapılmadıysa şifre iste
  if "password_correct" not in st.session_state:
    st.text_input(
        "Lütfen Şifreyi Girin:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    return False
  elif not st.session_state["password_correct"]:
    st.text_input(
        "Lütfen Şifreyi Girin:",
        type="password",
        on_change=password_entered,
        key="password",
    )
    st.error("😕 Şifre yanlış")
    return False
  else:
    return True


# Şifre doğru değilse uygulamanın geri kalanını çalıştırma
if not check_password():
  st.stop()

# ==========================================
# BURADAN SONRASI SİZİN ASIL UYGULAMA KODLARINIZ
# ==========================================
st.title("Planlama Uygulaması")
st.write("Hoş geldiniz, şifre başarıyla girildi!")

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
st.markdown("<h2 style='text-align: center;'>MAHALLE ve MADDE ARAMA</h2>", unsafe_allow_html=True)
st.markdown("---")

mode = st.radio(
    "Arama Modunu Seçin:",
    ["📍 Mahalle ve Kolluk Birimleri", "⚖️ Madde ve Suç Tanımları"],
    horizontal=True
)

search_text = st.text_input("🔍 Aramak istediğiniz metni girin...", "").strip()
search_text_upper = tr_upper(search_text)

st.markdown("---")

# --- MAHALLE MODU (Tema Uyumlu Kart Görünümü) ---
if "Mahalle" in mode:
    bulunan_sayi = 0
    for _, row in df_mahalle.iterrows():
        val_mahalle = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
        val_kolluk = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])
        val_ilce = "" if len(row) <= 3 or pd.isna(row.iloc[3]) else str(row.iloc[3])

        mahalle_val_upper = tr_upper(val_mahalle)

        if not search_text_upper or search_text_upper in mahalle_val_upper:
            bulunan_sayi += 1
            # st.container(border=True) otomatik olarak açık/koyu moda uyum sağlar
            with st.container(border=True):
                st.markdown(f"**📍 {val_mahalle}**")
                st.markdown(f"**İlçe:** {val_ilce}")
                st.markdown(f"**Kolluk Birimi:** {val_kolluk}")
                
    if bulunan_sayi == 0:
        st.info("Aranan kriterlere uygun sonuç bulunamadı.")

# --- MADDE MODU (Tema Uyumlu Kart Görünümü) ---
else:
    bulunan_sayi = 0
    for _, row in df_madde.iterrows():
        full_row_text = " ".join([tr_upper(x) for x in row.values])

        if not search_text_upper or search_text_upper in full_row_text:
            bulunan_sayi += 1
            val_madde = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
            val_suc = "" if len(row) <= 1 or pd.isna(row.iloc[1]) else str(row.iloc[1])
            val_kanun = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])

            with st.container(border=True):
                st.markdown(f"**⚖️ Madde {val_madde}** *({val_kanun})*")
                st.markdown(f"**Suç Tanımı:** {val_suc}")
                
    if bulunan_sayi == 0:
        st.info("Aranan kriterlere uygun sonuç bulunamadı.")