import os
import sys
import ctypes
import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont, QIcon, QPalette, QColor
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
    QStyleFactory
)


def resource_path(relative_path):
    """Hem Normal/Debug hem PyInstaller .exe modunda dosya yollarını bulur."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class ModernSearchApp(QMainWindow):

    def __init__(self, excel_path):
        super().__init__()
        self.excel_path = excel_path
        self.setWindowTitle("PLANLAMA")
        self.resize(1100, 700)

        # Aktif Arama Modu ("mahalle" veya "madde")
        self.current_mode = "mahalle"

        # Excel Verilerini Yükle
        if not self.load_data():
            sys.exit()

        self.init_ui()
        self.update_mode_ui()

    def load_data(self):
        """Excel verilerini güvenli şekilde yükler."""
        if not os.path.exists(self.excel_path):
            QMessageBox.critical(
                self, "Hata", f"'{self.excel_path}' dosyası bulunamadı!"
            )
            return False
        try:
            xls = pd.ExcelFile(self.excel_path)
            sheets = xls.sheet_names
            mahalle_sheet = next(
                (s for s in sheets if "MAHALLE" in s.upper()), None
            )
            madde_sheet = next(
                (s for s in sheets if "MADDE" in s.upper()), None
            )

            if not mahalle_sheet or not madde_sheet:
                QMessageBox.critical(
                    self, "Sayfa Hatası", "Excel'de gerekli sayfalar bulunamadı."
                )
                return False

            self.df_mahalle = pd.read_excel(
                self.excel_path, sheet_name=mahalle_sheet
            )
            self.df_madde = pd.read_excel(
                self.excel_path, sheet_name=madde_sheet
            )
            return True
        except Exception as e:
            QMessageBox.critical(
                self, "Yükleme Hatası", f"Excel okunurken hata oluştu:\n{e}"
            )
            return False

    def init_ui(self):
        icon_file = resource_path("app_icon.png")
        if os.path.exists(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # En Gerideki Arka Plan & Temel Bilesen Stilleri
        self.setStyleSheet(
            """
            QMainWindow {
                background-color: #D3E2F2;
            }
            QWidget {
                font-family: 'Segoe UI', sans-serif;
                color: #1F2937;
            }
            /* --- DİKEY KAYDIRMA ÇUBUĞU (VERTICAL SCROLLBAR) --- */
            QScrollBar:vertical {
                border: none;
                background-color: #E8F1FA;
                width: 12px;
                margin: 0px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical {
                background-color: #8BB3E1;
                min-height: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:vertical:hover {
                background-color: #2563EB; /* Üzerine gelindiğinde koyu mavi */
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px; /* Üst/Alt ok butonlarını gizler */
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }

            /* --- YATAY KAYDIRMA ÇUBUĞU (HORIZONTAL SCROLLBAR) --- */
            QScrollBar:horizontal {
                border: none;
                background-color: #E8F1FA;
                height: 12px;
                margin: 0px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal {
                background-color: #8BB3E1;
                min-width: 20px;
                border-radius: 6px;
            }
            QScrollBar::handle:horizontal:hover {
                background-color: #2563EB; /* Üzerine gelindiğinde koyu mavi */
            }
            QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
                width: 0px; /* Sağ/Sol ok butonlarını gizler */
            }
            QScrollBar::add-page:horizontal, QScrollBar::sub-page:horizontal {
                background: none;
            }
        """
        )

        # --- 1. ÜST BAŞLIK ALANI ---
        header_frame = QFrame()
        header_frame.setStyleSheet(
            """
            QFrame {
                background-color: rgba(255, 255, 255, 0.85);
                border-radius: 12px;
                border: 1px solid #B0C4DE;
            }
        """
        )
        header_layout = QHBoxLayout(header_frame)

        lbl_title = QLabel("MAHALLE-KOLLUK ve MADDE-SUÇ ARAMA")
        lbl_title.setFont(QFont("Segoe UI", 18, QFont.Weight.Bold))
        lbl_title.setStyleSheet("color: #1C3B68; border: none;")
        lbl_title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        header_layout.addWidget(lbl_title)

        main_layout.addWidget(header_frame)

        # --- 2. KONTROL & ARAMA PANELİ ---
        control_layout = QHBoxLayout()

        self.btn_mahalle = QPushButton("📍 Mahalle ve Kolluk Birimleri")
        self.btn_mahalle.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_mahalle.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.btn_mahalle.clicked.connect(lambda: self.switch_mode("mahalle"))

        self.btn_madde = QPushButton("⚖️ Madde ve Suç Tanımları")
        self.btn_madde.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_madde.setFont(QFont("Segoe UI", 10, QFont.Weight.Bold))
        self.btn_madde.clicked.connect(lambda: self.switch_mode("madde"))

        control_layout.addWidget(self.btn_mahalle)
        control_layout.addWidget(self.btn_madde)

        self.entry_search = QLineEdit()
        self.entry_search.setPlaceholderText("Aramak istediğiniz metni girin... 🔍")
        self.entry_search.setMaximumWidth(320)
        self.entry_search.setStyleSheet(
            """
            QLineEdit {
                background-color: #FFFFFF;
                border: 2px solid #8BB3E1;
                border-radius: 18px;
                padding: 6px 12px;
                font-size: 13px;
                color: #2C3E50;
            }
            QLineEdit:focus {
                border: 2px solid #2563EB;
            }
        """
        )
        self.entry_search.textChanged.connect(self.filter_data)
        control_layout.addWidget(self.entry_search)

        self.btn_clear = QPushButton("✖ Temizle")
        self.btn_clear.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_clear.setStyleSheet(
            """
            QPushButton {
                background-color: #EF4444;
                color: white;
                border-radius: 18px;
                padding: 8px 16px;
                font-weight: bold;
                border: none;
            }
            QPushButton:hover {
                background-color: #DC2626;
            }
        """
        )
        self.btn_clear.clicked.connect(self.clear_search)
        control_layout.addWidget(self.btn_clear)

        main_layout.addLayout(control_layout)

        # --- 3. TABLO ALANI ---
        self.table = QTableWidget()
        self.table.setEditTriggers(
            QAbstractItemView.EditTrigger.NoEditTriggers
        )
        self.table.setSelectionBehavior(
            QAbstractItemView.SelectionBehavior.SelectRows
        )
        self.table.setSelectionMode(
            QAbstractItemView.SelectionMode.SingleSelection
        )
        self.table.setAlternatingRowColors(True)

        self.table.setStyleSheet(
            """
            QTableWidget {
                background-color: #FFFFFF;
                alternate-background-color: #E8F1FA;
                gridline-color: #D0E1F9;
                border: 1px solid #B0C4DE;
                border-radius: 8px;
                font-size: 13px;
                color: #1F2937;
            }
            QHeaderView::section {
                background-color: #A3C2E0;
                color: #1C3B68;
                font-weight: bold;
                padding: 10px;
                border: none;
            }
            QTableWidget::item {
                color: #1F2937;
            }
            QTableWidget::item:selected {
                background-color: #3B82F6;
                color: #FFFFFF;
            }
        """
        )

        main_layout.addWidget(self.table)
        self.entry_search.setFocus()

    def switch_mode(self, mode):
        """Seçenek butonuna tıklandığında modu değiştirir."""
        if self.current_mode != mode:
            self.current_mode = mode
            self.update_mode_ui()

    def update_mode_ui(self):
        """Aktif butona göre arayüzü ve tablo başlıklarını günceller."""
        self.entry_search.blockSignals(True)
        self.entry_search.clear()
        self.entry_search.blockSignals(False)

        # Hatalı Hex kodları geçerli RGB/Hex renkleri ile güncellendi
        active_style = """
            QPushButton {
                background-color: #FEF08A;
                color: #1E293B;
                border-radius: 18px;
                padding: 8px 18px;
                border: 2px solid #CA8A04;
                font-weight: bold;
            }
        """
        inactive_style = """
            QPushButton {
                background-color: #FFFFFF;
                color: #1C3B68;
                border-radius: 18px;
                padding: 8px 18px;
                border: 1px solid #B0C4DE;
            }
            QPushButton:hover {
                background-color: #E8F1FA;
            }
        """

        if self.current_mode == "mahalle":
            self.btn_mahalle.setStyleSheet(active_style)
            self.btn_madde.setStyleSheet(inactive_style)
            headers = ["MAHALLE", "BAĞLI BULUNDUĞU KOLLUK BİRİMİ", "İLÇE"]
        else:
            self.btn_mahalle.setStyleSheet(inactive_style)
            self.btn_madde.setStyleSheet(active_style)
            headers = ["MADDE NO", "SUÇ TANIMI", "KANUN"]

        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)

        header = self.table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeMode.Interactive)

        if self.current_mode == "mahalle":
            self.table.setColumnWidth(0, 220)
            self.table.setColumnWidth(1, 500)
            self.table.setColumnWidth(2, 200)
        else:
            self.table.setColumnWidth(0, 150)
            self.table.setColumnWidth(1, 550)
            self.table.setColumnWidth(2, 220)

        self.table.setHorizontalScrollBarPolicy(
            Qt.ScrollBarPolicy.ScrollBarAsNeeded
        )
        self.filter_data()

    def clear_search(self):
        """Arama kutusunu temizler ve odağı arama kutusuna verir."""
        self.entry_search.clear()
        self.entry_search.setFocus()

    @staticmethod
    def tr_upper(text):
        """Türkçe karakter uyumlu büyütme işlemi."""
        if not isinstance(text, str):
            text = str(text) if pd.notna(text) else ""
        return text.replace("i", "İ").replace("ı", "I").upper()

    def filter_data(self):
        """Metne göre tablo verilerini optimize şekilde filtreler."""
        self.table.setUpdatesEnabled(False)
        self.table.setRowCount(0)

        search_text = self.tr_upper(self.entry_search.text().strip())

        if self.current_mode == "mahalle":
            df = self.df_mahalle
            for _, row in df.iterrows():
                val_mahalle = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                val_kolluk = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])
                val_ilce = "" if len(row) <= 3 or pd.isna(row.iloc[3]) else str(row.iloc[3])

                mahalle_val_upper = self.tr_upper(val_mahalle)

                if not search_text or search_text in mahalle_val_upper:
                    row_idx = self.table.rowCount()
                    self.table.insertRow(row_idx)
                    self.table.setItem(row_idx, 0, QTableWidgetItem(val_mahalle))
                    self.table.setItem(row_idx, 1, QTableWidgetItem(val_kolluk))
                    self.table.setItem(row_idx, 2, QTableWidgetItem(val_ilce))
        else:
            df = self.df_madde
            for _, row in df.iterrows():
                full_row_text = " ".join([self.tr_upper(x) for x in row.values])

                if not search_text or search_text in full_row_text:
                    row_idx = self.table.rowCount()
                    self.table.insertRow(row_idx)

                    val_madde = "" if pd.isna(row.iloc[0]) else str(row.iloc[0])
                    val_suc = "" if len(row) <= 1 or pd.isna(row.iloc[1]) else str(row.iloc[1])
                    val_kanun = "" if len(row) <= 2 or pd.isna(row.iloc[2]) else str(row.iloc[2])

                    self.table.setItem(row_idx, 0, QTableWidgetItem(val_madde))
                    self.table.setItem(row_idx, 1, QTableWidgetItem(val_suc))
                    self.table.setItem(row_idx, 2, QTableWidgetItem(val_kanun))

        self.table.setUpdatesEnabled(True)


if __name__ == "__main__":
    try:
        myappid = "mycompany.myproduct.subproduct.1"
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
    except Exception:
        pass

    app = QApplication(sys.argv)
    
    # --- SISTEM TEMASI DEĞİŞİMİNİ ENGELLEYEN AYARLAR ---
    app.setStyle(QStyleFactory.create("Fusion"))
    
    # Varsayılan paleti açık moda sabitleme
    light_palette = QPalette()
    light_palette.setColor(QPalette.ColorRole.Window, QColor("#D3E2F2"))
    light_palette.setColor(QPalette.ColorRole.WindowText, QColor("#1F2937"))
    light_palette.setColor(QPalette.ColorRole.Base, QColor("#FFFFFF"))
    light_palette.setColor(QPalette.ColorRole.AlternateBase, QColor("#E8F1FA"))
    light_palette.setColor(QPalette.ColorRole.ToolTipBase, QColor("#FFFFFF"))
    light_palette.setColor(QPalette.ColorRole.ToolTipText, QColor("#1F2937"))
    light_palette.setColor(QPalette.ColorRole.Text, QColor("#1F2937"))
    light_palette.setColor(QPalette.ColorRole.Button, QColor("#FFFFFF"))
    light_palette.setColor(QPalette.ColorRole.ButtonText, QColor("#1C3B68"))
    light_palette.setColor(QPalette.ColorRole.BrightText, QColor("#13E663"))
    light_palette.setColor(QPalette.ColorRole.Highlight, QColor("#3B82F6"))
    light_palette.setColor(QPalette.ColorRole.HighlightedText, QColor("#FFFFFF"))
    app.setPalette(light_palette)
    # --------------------------------------------------

    EXCEL_FILE = resource_path("veri.xlsx")
    window = ModernSearchApp(EXCEL_FILE)
    window.show()
    sys.exit(app.exec())

    EXCEL_FILE = "veri.xlsx"
    