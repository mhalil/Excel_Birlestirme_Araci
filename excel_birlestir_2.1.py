import sys
from pathlib import Path
import pandas as pd
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QPushButton, QLabel, QFileDialog, QGroupBox, QLineEdit, QCheckBox,
                               QSpinBox, QTextEdit, QMessageBox, QFormLayout)
from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QFont, QIcon

class ExcelMergerWorker(QThread):
    """
    Excel dosyalarını birleştiren arka plan işlemi.
    UI'ın donmasını engeller.
    """
    log_msg = Signal(str)
    finished_msg = Signal(bool, str)

    def __init__(self, folder_path, params):
        super().__init__()
        self.folder_path = folder_path
        self.params = params
        self.is_running = True

    def run(self):
        folder = Path(self.folder_path)
        excel_files = []
        for ext in ["*.xls", "*.xlsx", "*.xlsm"]:
            excel_files.extend(folder.glob(ext))
        
        if not excel_files:
            self.finished_msg.emit(False, "Seçilen klasörde Excel dosyası bulunamadı.")
            return

        merged_df = pd.DataFrame()
        
        try:
            for file_path in excel_files:
                if not self.is_running:
                    self.finished_msg.emit(False, "İşlem kullanıcı tarafından iptal edildi.")
                    return
                
                self.log_msg.emit(f"İşleniyor: {file_path.name}")
                df = self.process_file(file_path)
                
                if df is not None and not df.empty:
                    # pd.concat is more efficient than append
                    merged_df = pd.concat([merged_df, df], ignore_index=True)
            
            if merged_df.empty:
                self.finished_msg.emit(False, "Birleştirilecek veri bulunamadı.\n(Dosyalar boş veya belirlediğiniz format kurallarına uymuyor olabilir.)")
                return

            # Dosyayı kaydet
            save_path = self.params.get("save_path")
            if save_path:
                self.log_msg.emit("Dosya kaydediliyor, lütfen bekleyin...")
                
                # Excel motorunu uzantıya göre seç
                engine = 'xlwt' if save_path.endswith('.xls') else 'openpyxl'
                
                merged_df.to_excel(save_path, index=False, engine=engine)
                self.finished_msg.emit(True, f"İşlem Başarılı!\nDosya kaydedildi:\n{save_path}")
            else:
                self.finished_msg.emit(False, "Kayıt yeri seçilmediği için işlem tamamlanamadı.")
                
        except Exception as e:
            self.finished_msg.emit(False, f"Beklenmeyen bir hata oluştu:\n{str(e)}")

    def process_file(self, file_path):
        p = self.params
        use_sheet = p.get("use_sheet_name")
        sheet_name = p.get("sheet_name")
        
        use_header = p.get("use_header_row")
        header_row = p.get("header_row") if use_header else 1
        
        use_first_data = p.get("use_first_data_row")
        first_data_row = p.get("first_data_row") if use_first_data else (header_row + 1)
        
        use_cols_filter = p.get("use_cols_filter")
        usecols = p.get("usecols") if use_cols_filter else None
        
        use_text_cols = p.get("use_text_cols")
        text_cols_raw = p.get("text_cols")
        
        use_pattern = p.get("use_pattern")
        keep_rows = p.get("keep_rows")
        skip_rows = p.get("skip_rows")

        # Okuma parametrelerini ayarla
        read_kwargs = {}
        # .xls dosyaları için calamine motoru (yeni pandas sürümlerinde xlrd desteği kalktı)
        if str(file_path).lower().endswith('.xls') and not str(file_path).lower().endswith('.xlsm'):
            read_kwargs["engine"] = "calamine"
        if use_sheet and sheet_name:
            if sheet_name.isdigit():
                read_kwargs["sheet_name"] = int(sheet_name)
            else:
                read_kwargs["sheet_name"] = sheet_name
        else:
            read_kwargs["sheet_name"] = 0 # Varsayılan: İlk sayfa

        if use_text_cols and text_cols_raw:
            dtypes_dict = {}
            for col in text_cols_raw.split(","):
                col = col.strip()
                if col:
                    dtypes_dict[int(col) if col.isdigit() else col] = str
            if dtypes_dict:
                read_kwargs["dtype"] = dtypes_dict

        try:
            # 1. Başlık satırını bul ve oku
            header_list = None
            if header_row > 1:
                # Başlık satırı 1. satır değilse, sadece başlık satırını oku
                # dtype kullanmıyoruz çünkü sadece başlık okuyoruz
                df_header = pd.read_excel(
                    file_path, header=None, skiprows=header_row - 1, nrows=1, 
                    usecols=usecols, sheet_name=read_kwargs.get("sheet_name", 0)
                )
                header_list = df_header.iloc[0].astype(str).tolist()
            else:
                # Başlık satırı 1. satırsa
                df_header = pd.read_excel(
                    file_path, header=0, nrows=0, usecols=usecols, 
                    sheet_name=read_kwargs.get("sheet_name", 0)
                )
                header_list = df_header.columns.tolist()

            # 2. Verileri okumaya başla
            # İlk veri satırına kadar olan her şeyi atla
            skip_for_data = list(range(0, first_data_row - 1))
            df = pd.read_excel(
                file_path, header=None, names=header_list, 
                skiprows=skip_for_data, usecols=usecols, **read_kwargs
            )
            
            # Platform bağımsız dosya adı ekle
            df["Dosya ADI"] = file_path.name
            
            # 3. Döngüsel satır silme örüntüsü (Ataşman cetveli gibi durumlar için)
            if use_pattern and keep_rows > 0 and skip_rows > 0:
                total_rows = len(df)
                indices_to_drop = []
                current_idx = keep_rows
                
                while current_idx < total_rows:
                    # skip_rows kadar satırı listeye ekle
                    end_idx = min(current_idx + skip_rows, total_rows)
                    indices_to_drop.extend(range(current_idx, end_idx))
                    current_idx += (keep_rows + skip_rows)
                    
                df.drop(indices_to_drop, axis=0, inplace=True)
                
            return df
            
        except ValueError as ve:
             self.log_msg.emit(f"UYARI: {file_path.name} okunamadı (Sütun/Sayfa hatası). Hata: {str(ve)}")
             return None
        except Exception as e:
            self.log_msg.emit(f"UYARI: {file_path.name} okunamadı. Hata: {str(e)}")
            return None

    def stop(self):
        self.is_running = False

class ExcelMergerApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.folder_path = None
        self.worker = None
        self.init_ui()
        self.apply_stylesheet()

    def init_ui(self):
        self.setWindowTitle(".:: Excel Dosyalarını Birleştir ::.")
        self.setMinimumSize(650, 600)
        
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        main_layout = QVBoxLayout(main_widget)
        main_layout.setSpacing(15)
        main_layout.setContentsMargins(20, 20, 20, 20)

        # --- Klasör Seçim Bölümü ---
        folder_group = QGroupBox("1. Kaynak Klasör Seçimi")
        folder_layout = QHBoxLayout(folder_group)
        
        self.btn_select_folder = QPushButton("Klasör Seç...")
        self.btn_select_folder.setCursor(Qt.PointingHandCursor)
        self.btn_select_folder.clicked.connect(self.select_folder)
        
        self.lbl_folder_info = QLabel("Henüz klasör seçilmedi.")
        self.lbl_folder_info.setWordWrap(True)
        self.lbl_folder_info.setStyleSheet("color: #555; font-style: italic;")
        
        folder_layout.addWidget(self.btn_select_folder)
        folder_layout.addWidget(self.lbl_folder_info, 1) # strech=1
        main_layout.addWidget(folder_group)

        # --- Parametreler Bölümü ---
        params_group = QGroupBox("2. Birleştirme Parametreleri (Okuma Kuralları)")
        params_layout = QVBoxLayout(params_group)
        
        left_form = QFormLayout()
        left_form.setSpacing(10)
        
        self.chk_sheet = QCheckBox("Özel Sayfa Adı Belirt:")
        self.chk_sheet.toggled.connect(self.toggle_sheet_input)
        self.entry_sheet = QLineEdit()
        self.entry_sheet.setPlaceholderText("Örn: Sayfa1 veya 0")
        self.entry_sheet.setEnabled(False)
        left_form.addRow(self.chk_sheet, self.entry_sheet)

        self.chk_header = QCheckBox("Özel Başlık Satırı No:")
        self.chk_header.setChecked(False)
        self.chk_header.toggled.connect(lambda c: self.spin_header.setEnabled(c))
        self.spin_header = QSpinBox()
        self.spin_header.setRange(1, 1000)
        self.spin_header.setValue(4)
        self.spin_header.setEnabled(False)
        left_form.addRow(self.chk_header, self.spin_header)

        self.chk_first_data = QCheckBox("Özel İlk Veri Satırı No:")
        self.chk_first_data.setChecked(False)
        self.chk_first_data.toggled.connect(lambda c: self.spin_first_data.setEnabled(c))
        self.spin_first_data = QSpinBox()
        self.spin_first_data.setRange(1, 1000)
        self.spin_first_data.setValue(5)
        self.spin_first_data.setEnabled(False)
        left_form.addRow(self.chk_first_data, self.spin_first_data)

        self.chk_cols = QCheckBox("Özel Sütunları Kopyala:")
        self.chk_cols.setChecked(False)
        self.chk_cols.toggled.connect(lambda c: self.entry_cols.setEnabled(c))
        self.entry_cols = QLineEdit()
        self.entry_cols.setText("B:G")
        self.entry_cols.setPlaceholderText("Örn: B:G, A,C,D")
        self.entry_cols.setToolTip("Seçili değilse tüm sütunlar alınır.")
        self.entry_cols.setEnabled(False)
        left_form.addRow(self.chk_cols, self.entry_cols)

        self.chk_text_cols = QCheckBox("Metin Olarak Okunacak Sütunlar:")
        self.chk_text_cols.toggled.connect(lambda c: self.entry_text_cols.setEnabled(c))
        self.entry_text_cols = QLineEdit()
        self.entry_text_cols.setPlaceholderText("Örn: TCKN, IBAN veya 0, 1")
        self.entry_text_cols.setToolTip("Bilimsel gösterimi (1.5E+16 vb.) veya 0 silinmesini engellemek için sütun isimlerini veya indekslerini virgülle girin.")
        self.entry_text_cols.setEnabled(False)
        left_form.addRow(self.chk_text_cols, self.entry_text_cols)
        
        params_layout.addLayout(left_form)
        main_layout.addWidget(params_group)

        # --- Tekrarlayan Satır Silme Bölümü ---
        pattern_group = QGroupBox("3. Özel Satır Silme (Örüntü - Ataşman vb. için)")
        pattern_layout = QFormLayout(pattern_group)
        pattern_layout.setSpacing(10)
        
        self.chk_pattern = QCheckBox("Tekrarlayan Alt Satırları Sil")
        self.chk_pattern.setToolTip("Ataşman cetvellerinde olan ara toplam/imza satırlarını döngüsel olarak silmek için.")
        self.chk_pattern.setChecked(False)
        self.chk_pattern.toggled.connect(self.toggle_pattern_inputs)
        pattern_layout.addRow(self.chk_pattern)
        
        self.spin_keep = QSpinBox()
        self.spin_keep.setRange(1, 5000)
        self.spin_keep.setValue(7)
        self.spin_keep.setEnabled(False)
        pattern_layout.addRow("Veri Satırı Sayısı:", self.spin_keep)
        
        self.spin_skip = QSpinBox()
        self.spin_skip.setRange(1, 5000)
        self.spin_skip.setValue(5)
        self.spin_skip.setEnabled(False)
        pattern_layout.addRow("Silinecek Satır Sayısı:", self.spin_skip)

        main_layout.addWidget(pattern_group)

        # --- Log Ekranı ---
        self.text_log = QTextEdit()
        self.text_log.setReadOnly(True)
        self.text_log.setPlaceholderText("İşlem detayları burada görüntülenecek...")
        self.text_log.setStyleSheet("background-color: #f8f9fa; border: 1px solid #ced4da; border-radius: 4px; padding: 5px;")
        main_layout.addWidget(QLabel("4. İşlem Kayıtları:"))
        main_layout.addWidget(self.text_log)

        # --- Alt Butonlar ---
        btn_layout = QHBoxLayout()
        
        self.btn_help = QPushButton("Yardım")
        self.btn_help.setCursor(Qt.PointingHandCursor)
        self.btn_help.clicked.connect(self.show_help)
        
        self.btn_about = QPushButton("Hakkında")
        self.btn_about.setCursor(Qt.PointingHandCursor)
        self.btn_about.clicked.connect(self.show_about)
        
        self.btn_merge = QPushButton("Dosyaları Birleştir ve Kaydet")
        self.btn_merge.setCursor(Qt.PointingHandCursor)
        self.btn_merge.setFixedHeight(40)
        self.btn_merge.setEnabled(False) # Klasör seçilene kadar pasif
        self.btn_merge.clicked.connect(self.start_merge)
        
        btn_layout.addWidget(self.btn_help)
        btn_layout.addWidget(self.btn_about)
        btn_layout.addStretch()
        btn_layout.addWidget(self.btn_merge)
        
        main_layout.addLayout(btn_layout)

    def apply_stylesheet(self):
        # Modern ve temiz bir arayüz stili
        self.setStyleSheet("""
            QMainWindow {
                background-color: #ffffff;
            }
            QGroupBox {
                font-weight: bold;
                border: 1px solid #dee2e6;
                border-radius: 6px;
                margin-top: 10px;
                padding-top: 15px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                padding: 0 5px;
                color: #2b8a3e;
            }
            QPushButton {
                background-color: #e9ecef;
                border: 1px solid #ced4da;
                border-radius: 4px;
                padding: 6px 15px;
                color: #212529;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #dee2e6;
            }
            QPushButton:disabled {
                background-color: #f8f9fa;
                color: #adb5bd;
            }
            QLineEdit, QSpinBox {
                padding: 5px;
                border: 1px solid #ced4da;
                border-radius: 4px;
                min-width: 100px;
            }
            QLineEdit:focus, QSpinBox:focus {
                border: 1px solid #4dabf7;
            }
        """)
        # Kaydet butonu için özel stil
        self.btn_merge.setStyleSheet("""
            QPushButton {
                background-color: #2b8a3e;
                color: white;
                border: none;
                font-size: 13px;
                padding: 0 20px;
            }
            QPushButton:hover { background-color: #2f9e44; }
            QPushButton:disabled { background-color: #8ce99a; color: #ebfbee; }
        """)

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Excel Dosyalarının Bulunduğu Klasörü Seç")
        if folder:
            self.folder_path = folder
            
            # Klasördeki excel dosyalarını say
            excel_count = len(list(Path(folder).glob("*.xls*")))
            
            if excel_count > 0:
                self.lbl_folder_info.setText(f"<b>Seçilen Klasör:</b> {folder}<br><i>{excel_count} adet Excel dosyası bulundu.</i>")
                self.lbl_folder_info.setStyleSheet("color: #000;")
                self.btn_merge.setEnabled(True)
                self.log(f"Klasör seçildi. {excel_count} dosya işlenmeyi bekliyor.")
            else:
                self.lbl_folder_info.setText("Seçilen klasörde hiç Excel dosyası (.xls, .xlsx) bulunamadı.")
                self.lbl_folder_info.setStyleSheet("color: #c92a2a;")
                self.btn_merge.setEnabled(False)
                self.folder_path = None

    def toggle_sheet_input(self, checked):
        self.entry_sheet.setEnabled(checked)

    def toggle_pattern_inputs(self, checked):
        self.spin_keep.setEnabled(checked)
        self.spin_skip.setEnabled(checked)

    def log(self, message):
        self.text_log.append(message)
        # Scroll to bottom
        scrollbar = self.text_log.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def show_help(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Yardım ve Parametre Açıklamaları")
        msg.setIcon(QMessageBox.Information)
        msg.setText("<b>Excel Birleştirme Aracı Kullanım Rehberi</b><br>Tüm onay kutuları (checkbox) isteğe bağlıdır. İşaretsiz bırakılırlarsa Excel dosyası standart bir tablo gibi baştan sona okunur.")
        
        help_text = (
            "<b>1. Kaynak Klasör Seçimi:</b> Birleştirmek istediğiniz Excel dosyalarının (.xls, .xlsx) bulunduğu klasörü seçin.<br><br>"
            
            "<b>2. Özel Sayfa Adı Belirt:</b> Dosyadaki tüm sayfalardan ziyade sadece belirli bir sayfadan (Örn: 'Sayfa1' veya ilk sayfa için '0') verileri çekmek isterseniz işaretleyin.<br><br>"
            
            "<b>3. Özel Başlık Satırı No:</b> Sütun başlıklarının (ad, soyad, tutar vb.) bulunduğu satır numarasıdır. İşaretsiz ise 1. satır (en üst) başlık kabul edilir.<br><br>"
            
            "<b>4. Özel İlk Veri Satırı No:</b> Asıl verilerin (rakamlar, kişi bilgileri vb.) başladığı ilk satır numarasıdır. İşaretsiz ise başlık satırının hemen altından başladığı kabul edilir.<br><br>"
            
            "<b>5. Özel Sütunları Kopyala:</b> Sadece belirli sütunları (Örn: 'B:G' veya 'A, C, F') birleştirmek isterseniz kullanın. İşaretsiz ise tüm sütunlar alınır.<br><br>"
            
            "<b>6. Metin Olarak Okunacak Sütunlar:</b> TCKN, IBAN veya Telefon gibi sayısal ama büyük boyutlu veriler Excel'de '1,50E+16' şeklinde bilimsel formata dönüşebilir veya başındaki '0' silinebilir. Bunu önlemek için buraya o sütunların adlarını (Örn: 'TCKN, IBAN') veya indekslerini (Örn: '0, 1') yazabilirsiniz.<br><br>"
            
            "<b>7. Tekrarlayan Alt Satırları Sil (Örüntü):</b> Ataşman cetveli gibi standart olmayan excel şablonlarında periyodik olarak tekrar eden satırları (imza alanları, ara toplamlar vb.) döngüsel olarak silmek içindir.<br>"
            "&nbsp;&nbsp;&nbsp;• <i>Veri Satırı Sayısı:</i> Döngüde ardışık olarak tutulacak asıl veri satır adedidir.<br>"
            "&nbsp;&nbsp;&nbsp;• <i>Silinecek Satır Sayısı:</i> Veri bloğundan sonra atlanacak (silinecek) gereksiz satır adedidir."
        )
        msg.setInformativeText(help_text)
        msg.exec()

    def show_about(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Hakkında")
        msg.setIcon(QMessageBox.Information)
        msg.setText("<b>Excel Birleştirme Aracı v2.1</b>")
        msg.setInformativeText(
            "Bu program, seçili dizindeki Excel dosyalarını belirlediğiniz kurallara göre "
            "tek bir dosyada birleştirmek için geliştirilmiştir.<br><br>"
            "<b>Vibe Coder:</b> Mustafa Halil GÖRENTAŞ<br>"
            "<li>Kaynak Kod: <a href=\"https://github.com/mhalil/Excel_Birlestir_GUI\">github.com/mhalil/Excel_Birlestir_GUI</a></li>"
            "<p><b>Teknik Bilgiler:</b></p>"
            "<li>Platform: Google Antigravity ve OpenCode</li>"
            "<li>Metodoloji: Vibe Coding</li>"
            "<li>Programlama Dili: Python (3.12.4)</li>"
            "<li>Kullanılan Kütüphaneler: Pandas, Openpyxl, xlwt, xlrd, PySide6</li>"
            "<br>" 
            "GPL Lisansı Altında Dağıtılmaktadır. | 2026"
        )
        msg.exec()

    def start_merge(self):
        if not self.folder_path:
            QMessageBox.warning(self, "Uyarı", "Lütfen önce bir klasör seçin.")
            return

        # Kayıt yeri sor
        save_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Birleştirilmiş Dosyayı Kaydet", 
            "Birlestirilmis_Tablo.xlsx", 
            "Excel Dosyası (*.xlsx);;Eski Excel Dosyası (*.xls)"
        )
        
        if not save_path:
            self.log("Kayıt işlemi iptal edildi.")
            return

        # Parametreleri topla
        usecols_text = self.entry_cols.text().strip()
        params = {
            "use_sheet_name": self.chk_sheet.isChecked(),
            "sheet_name": self.entry_sheet.text().strip(),
            "use_header_row": self.chk_header.isChecked(),
            "header_row": self.spin_header.value(),
            "use_first_data_row": self.chk_first_data.isChecked(),
            "first_data_row": self.spin_first_data.value(),
            "use_cols_filter": self.chk_cols.isChecked(),
            "usecols": usecols_text if usecols_text else None,
            "use_text_cols": self.chk_text_cols.isChecked(),
            "text_cols": self.entry_text_cols.text().strip(),
            "use_pattern": self.chk_pattern.isChecked(),
            "keep_rows": self.spin_keep.value(),
            "skip_rows": self.spin_skip.value(),
            "save_path": save_path
        }

        # Arayüzü kilitle
        self.btn_merge.setEnabled(False)
        self.btn_select_folder.setEnabled(False)
        self.text_log.clear()
        self.log("BİRLEŞTİRME İŞLEMİ BAŞLATILDI...")
        self.log("-" * 40)

        # Worker thread başlat
        self.worker = ExcelMergerWorker(self.folder_path, params)
        self.worker.log_msg.connect(self.log)
        self.worker.finished_msg.connect(self.on_merge_finished)
        self.worker.start()

    def on_merge_finished(self, success, message):
        self.log("-" * 40)
        self.log(message)
        
        if success:
            QMessageBox.information(self, "İşlem Tamamlandı", message)
        else:
            QMessageBox.warning(self, "İşlem Durduruldu / Hata", message)
            
        # Arayüzü aç
        self.btn_merge.setEnabled(True)
        self.btn_select_folder.setEnabled(True)
        self.worker = None

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
        event.accept()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    
    # Font ayarları
    font = QFont("Segoe UI", 10)
    app.setFont(font)
    
    window = ExcelMergerApp()
    window.show()
    sys.exit(app.exec())
