import sys
import os
import json
import xml.etree.ElementTree as ET
import csv
from datetime import datetime
from collections import defaultdict

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QListWidget, QLineEdit, QFileDialog, QMessageBox,
    QLabel, QGroupBox, QStatusBar, QAbstractItemView, QDialog,
    QTableWidget, QTableWidgetItem, QHeaderView, QComboBox, QTabWidget
)
from PyQt6.QtCore import Qt, QDir

# --- FUNGSI GLOBAL & KONSTANTA ---
PROFILE_DIR = "profiles"
PROFILE_EXTENSION = ".json"

def get_default_profile_config():
    """Mengembalikan dictionary berisi konfigurasi profil default yang lengkap, termasuk untuk CSV->XML."""
    return {
        "settings": {
            "item_loop_path": ".//multiShipmentOrderLineItem",
            # --- BARU: Pengaturan untuk CSV ke XML ---
            "root_element_name": "multiShipmentOrder",
            "xml_grouping_key": "Order Number"
        },
        "columns": {
            "Item Number": {"type": "attribute", "path": "number", "source": "item"},
            "Order Date": {"type": "xpath", "path": ".//creationDateTime", "source": "root"},
            "Order Number": {"type": "xpath", "path": ".//uniqueCreatorIdentification", "source": "root"},
            "Kode Barang": {"type": "xpath_indexed", "path": ".//additionalTradeItemIdentification/additionalTradeItemIdentificationValue", "index": 0, "source": "item"},
            "Item Description": {"type": "xpath_indexed", "path": ".//additionalTradeItemIdentification/additionalTradeItemIdentificationValue", "index": 1, "source": "item"},
            "Quantity": {"type": "xpath_indexed", "path": ".//additionalTradeItemIdentification/additionalTradeItemIdentificationValue", "index": 8, "source": "item"},
            "Unit Price": {"type": "xpath", "path": ".//netPrice/amount/monetaryAmount", "source": "item"},
            "Total Price": {"type": "xpath", "path": ".//netAmount/amount/monetaryAmount", "source": "item"},
            "Qty Barang": {"type": "calculated", "formula": "{Total Price}/{Unit Price}", "source": "item"},
            "Ship to": {"type": "xpath", "path": ".//shipToLogistics/shipTo/additionalPartyIdentification/additionalPartyIdentificationValue", "source": "root"},
            "GTIN": {"type": "xpath", "path": ".//gtin", "source": "item"},
            "Document Status": {"type": "attribute", "path": "documentStatus", "source": "root"},
            "Seller": {"type": "xpath", "path": ".//seller/additionalPartyIdentification/additionalPartyIdentificationValue", "source": "root"},
            "CTRI": {"type": "xpath", "path": ".//additionalPartyIdentification[additionalPartyIdentificationType='FOR_INTERNAL_USE_11']/additionalPartyIdentificationValue", "source": "root"}
        }
    }

# --- Jendela Dialog Pengaturan (DENGAN PERUBAHAN YANG DIMINTA) ---
class MappingDialog(QDialog):
    def __init__(self, active_profile_path, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Pengaturan Mapping & Profil Dinamis")
        self.resize(850, 700) # Ganti setGeometry dengan resize
        self.active_profile_path = active_profile_path
        self.config = self.load_config(self.active_profile_path)
        self.detected_xpaths = []

        # --- PERUBAHAN 2a: Ekstrak path yang sudah ada dari konfigurasi JSON ---
        self.known_json_paths = []
        columns_data = self.config.get("columns", {})
        known_paths_set = set()
        for details in columns_data.values():
            if "path" in details:
                known_paths_set.add(details["path"])
        self.known_json_paths = sorted(list(known_paths_set))
        # --- AKHIR PERUBAHAN 2a ---

        self.initUI()
        self.populate_table_from_config()

        # --- PERUBAHAN 1: Posisikan dialog di tengah parent ---
        if parent:
            parent_rect = parent.frameGeometry()
            dialog_rect = self.frameGeometry()
            dialog_rect.moveCenter(parent_rect.center())
            self.move(dialog_rect.topLeft())
        # --- AKHIR PERUBAHAN 1 ---

    def load_config(self, profile_path):
        try:
            with open(profile_path, 'r', encoding='utf-8') as f: return json.load(f)
        except (FileNotFoundError, json.JSONDecodeError): return get_default_profile_config()
    def initUI(self):
        layout = QVBoxLayout(self); profile_group = QGroupBox("Manajemen Profil"); profile_layout = QHBoxLayout(); btn_save_as = QPushButton("Simpan Profil Sebagai..."); btn_save_as.clicked.connect(self.save_profile_as); btn_delete = QPushButton("Hapus Profil Ini"); btn_delete.clicked.connect(self.delete_current_profile); profile_layout.addWidget(btn_save_as); profile_layout.addWidget(btn_delete); profile_group.setLayout(profile_layout); layout.addWidget(profile_group)
        detection_group = QGroupBox("Deteksi Otomatis"); detection_layout = QHBoxLayout(); btn_detect = QPushButton("Pilih XML & Tambah Kolom Otomatis"); btn_detect.clicked.connect(self.run_detection_and_add); detection_layout.addWidget(btn_detect); detection_group.setLayout(detection_layout); layout.addWidget(detection_group)
        mapping_group = QGroupBox("Editor Mapping Kolom"); mapping_layout = QVBoxLayout(); table_actions_layout = QHBoxLayout(); btn_add_row = QPushButton("➕ Tambah Baris Manual"); btn_add_row.clicked.connect(self.add_manual_row); btn_remove_row = QPushButton("➖ Hapus Baris Terpilih"); btn_remove_row.clicked.connect(self.remove_selected_rows); table_actions_layout.addWidget(btn_add_row); table_actions_layout.addWidget(btn_remove_row); table_actions_layout.addStretch(); mapping_layout.addLayout(table_actions_layout); self.table = QTableWidget(); self.table.setColumnCount(4); self.table.setHorizontalHeaderLabels(["Nama Kolom CSV", "Tipe", "Path / Formula", "Sumber"]); self.table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch); mapping_layout.addWidget(self.table); loop_path_layout = QHBoxLayout(); loop_path_layout.addWidget(QLabel("Path untuk Perulangan Item:")); self.txt_item_loop_path = QLineEdit(self.config.get("settings", {}).get("item_loop_path", "")); loop_path_layout.addWidget(self.txt_item_loop_path); mapping_layout.addLayout(loop_path_layout); mapping_group.setLayout(mapping_layout); layout.addWidget(mapping_group)
        btn_layout = QHBoxLayout(); btn_save = QPushButton("Simpan Perubahan pada Profil Ini"); btn_save.clicked.connect(self.save_and_accept); btn_cancel = QPushButton("Batal"); btn_cancel.clicked.connect(self.reject); btn_layout.addStretch(); btn_layout.addWidget(btn_save); btn_layout.addWidget(btn_cancel); layout.addLayout(btn_layout)
    def run_detection_and_add(self):
        xml_file, _ = QFileDialog.getOpenFileName(self, "Pilih File XML Sampel", "", "XML Files (*.xml)");
        if not xml_file: return
        try:
            tree = ET.parse(xml_file); root = tree.getroot(); detected_paths = set(); self._recursive_detect(root, ".", detected_paths);
            # Gabungkan path yang baru dideteksi dengan yang sudah ada di JSON
            self.detected_xpaths = sorted(list(detected_paths))
            mapped_paths = set();
            for i in range(self.table.rowCount()):
                path_widget = self.table.cellWidget(i, 2)
                if path_widget: mapped_paths.add(path_widget.currentText())
            new_paths = detected_paths - mapped_paths
            if not new_paths: QMessageBox.information(self, "Informasi", "Tidak ada kolom baru yang ditemukan."); return
            for path in sorted(list(new_paths)): self.add_manual_row(default_details={"col_name": self._path_to_col_name(path), "type": "attribute" if "/@" in path else "xpath", "path": path, "source": "item"})
            QMessageBox.information(self, "Sukses", f"{len(new_paths)} kolom baru telah ditambahkan.")
        except Exception as e: QMessageBox.critical(self, "Error", f"Gagal mem-parsing file XML: {e}")
    def _recursive_detect(self, element, current_path, paths):
        for attr_name in element.attrib: paths.add(f"{current_path}/@{attr_name}")
        if element.text and element.text.strip(): paths.add(current_path)
        for child in element: self._recursive_detect(child, current_path + "/" + child.tag.split('}')[-1], paths)
    def _path_to_col_name(self, path):
        name = path.split('/')[-1].replace('@', ''); return name.replace('_', ' ').title()
    def add_manual_row(self, default_details=None):
        if not isinstance(default_details, dict): default_details = {"col_name": "Kolom Baru", "type": "xpath", "path": "", "source": "item"}
        row_position = self.table.rowCount(); self.table.insertRow(row_position)
        self.table.setItem(row_position, 0, QTableWidgetItem(default_details.get("col_name", "Error")))
        combo_type = QComboBox(); combo_type.addItems(["xpath", "attribute", "xpath_indexed", "calculated", "static_value"]); combo_type.setCurrentText(default_details.get("type", "xpath")); self.table.setCellWidget(row_position, 1, combo_type)
        combo_path = QComboBox(); combo_path.setEditable(True)
        # --- PERUBAHAN 2b: Gabungkan path dari JSON dan deteksi otomatis untuk mengisi dropdown ---
        all_available_paths = sorted(list(set(self.known_json_paths + self.detected_xpaths)))
        combo_path.addItems(all_available_paths)
        # --- AKHIR PERUBAHAN 2b ---
        combo_path.setCurrentText(default_details.get("path", "")); self.table.setCellWidget(row_position, 2, combo_path)
        combo_source = QComboBox(); combo_source.addItems(["root", "item"]); combo_source.setCurrentText(default_details.get("source", "item")); self.table.setCellWidget(row_position, 3, combo_source)
    def remove_selected_rows(self):
        selected_indexes = self.table.selectedIndexes()
        if not selected_indexes: QMessageBox.warning(self, "Tidak Ada Pilihan", "Pilih baris untuk dihapus."); return
        rows_to_remove = sorted(list(set(index.row() for index in selected_indexes)), reverse=True)
        for row in rows_to_remove: self.table.removeRow(row)
    def populate_table_from_config(self):
        self.table.setRowCount(0); columns_data = self.config.get("columns", {})
        for col_name, details in columns_data.items(): self.add_manual_row(default_details={"col_name": col_name, "type": details.get("type", "xpath"), "path": details.get("path", details.get("formula", "")), "source": details.get("source", "item")})
    def save_and_accept(self):
        current_config = {"settings": {}, "columns": {}};
        old_settings = self.config.get("settings", {})
        current_config["settings"]["item_loop_path"] = self.txt_item_loop_path.text()
        current_config["settings"]["root_element_name"] = old_settings.get("root_element_name", "multiShipmentOrder")
        current_config["settings"]["xml_grouping_key"] = old_settings.get("xml_grouping_key", "Order Number")
        for i in range(self.table.rowCount()):
            col_name = self.table.item(i, 0).text(); col_type = self.table.cellWidget(i, 1).currentText(); path_val = self.table.cellWidget(i, 2).currentText(); source = self.table.cellWidget(i, 3).currentText()
            details = {"type": col_type, "source": source};
            if col_type == "calculated": details["formula"] = path_val
            else: details["path"] = path_val
            if col_type == "xpath_indexed": details["index"] = 0
            current_config["columns"][col_name] = details
        try:
            os.makedirs(os.path.dirname(self.active_profile_path), exist_ok=True)
            with open(self.active_profile_path, 'w', encoding='utf-8') as f: json.dump(current_config, f, indent=4)
            self.accept()
        except Exception as e: QMessageBox.critical(self, "Error", f"Gagal menyimpan file profil:\n{e}")
    def save_profile_as(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Simpan Profil Baru", os.path.join(os.getcwd(), PROFILE_DIR), f"Profil (*{PROFILE_EXTENSION})")
        if file_path:
            if not file_path.endswith(PROFILE_EXTENSION): file_path += PROFILE_EXTENSION
            self.active_profile_path = file_path; self.save_and_accept()
    def delete_current_profile(self):
        profile_name = os.path.basename(self.active_profile_path)
        if profile_name == f"default{PROFILE_EXTENSION}": QMessageBox.warning(self, "Aksi Dilarang", "Profil 'default' tidak dapat dihapus."); return
        reply = QMessageBox.question(self, "Konfirmasi Hapus", f"Yakin ingin menghapus profil '{profile_name}'?", QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No, QMessageBox.StandardButton.No)
        if reply == QMessageBox.StandardButton.Yes:
            try:
                os.remove(self.active_profile_path); QMessageBox.information(self, "Sukses", f"Profil '{profile_name}' telah dihapus."); self.accept()
            except Exception as e: QMessageBox.critical(self, "Error", f"Gagal menghapus profil: {e}")


# --- Kelas Utama Aplikasi dengan UI Bertab (TIDAK ADA PERUBAHAN DI SINI) ---
class XMLConverterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Konverter Data XML <-> CSV v0.6")
        self.setGeometry(100, 100, 700, 500)
        
        self.active_profile_path = ""
        os.makedirs(PROFILE_DIR, exist_ok=True)
        
        self.xml_files = []
        self.xml_to_csv_output_file = ""

        self.csv_file_input = ""
        self.csv_to_xml_output_dir = ""

        self.initUI()
        self.load_profiles()

    def initUI(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        profile_layout = QHBoxLayout(); profile_layout.addWidget(QLabel("Profil Aktif:")); self.combo_profiles = QComboBox(); self.combo_profiles.currentTextChanged.connect(self.on_profile_changed); profile_layout.addWidget(self.combo_profiles, 1); self.btn_settings = QPushButton("Pengaturan Profil"); self.btn_settings.clicked.connect(self.open_settings); profile_layout.addWidget(self.btn_settings)
        main_layout.addLayout(profile_layout)

        tab_widget = QTabWidget()
        self.xml_to_csv_tab = self.create_xml_to_csv_tab()
        self.csv_to_xml_tab = self.create_csv_to_xml_tab()
        
        tab_widget.addTab(self.xml_to_csv_tab, "XML  ➔  CSV")
        tab_widget.addTab(self.csv_to_xml_tab, "CSV  ➔  XML")
        
        main_layout.addWidget(tab_widget)
        self.statusBar = QStatusBar(); self.setStatusBar(self.statusBar)

    def create_xml_to_csv_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab)
        input_groupbox = QGroupBox("1. Pilih File-file XML Input"); input_layout = QVBoxLayout(); button_layout = QHBoxLayout(); self.btn_select_xmls = QPushButton("Pilih File XML"); self.btn_select_xmls.clicked.connect(self.select_xml_files); self.btn_remove_selected = QPushButton("Hapus File Terpilih"); self.btn_remove_selected.clicked.connect(self.remove_selected_files); button_layout.addWidget(self.btn_select_xmls); button_layout.addWidget(self.btn_remove_selected); self.list_widget_files = QListWidget(); self.list_widget_files.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection); input_layout.addLayout(button_layout); input_layout.addWidget(QLabel("File yang dipilih:")); input_layout.addWidget(self.list_widget_files); input_groupbox.setLayout(input_layout)
        output_groupbox = QGroupBox("2. Tentukan File CSV Output"); output_layout = QHBoxLayout(); self.txt_xml_to_csv_output = QLineEdit(); self.txt_xml_to_csv_output.setReadOnly(True); btn_select_output = QPushButton("..."); btn_select_output.setFixedWidth(40); btn_select_output.clicked.connect(self.select_xml_to_csv_output); output_layout.addWidget(self.txt_xml_to_csv_output); output_layout.addWidget(btn_select_output); output_groupbox.setLayout(output_layout)
        convert_groupbox = QGroupBox("3. Jalankan Konversi"); convert_layout = QVBoxLayout(); btn_convert = QPushButton("Konversi ke CSV"); btn_convert.setStyleSheet("font-size: 16px; padding: 10px; background-color: #4CAF50; color: white;"); btn_convert.clicked.connect(self.run_xml_to_csv_conversion); convert_layout.addWidget(btn_convert); convert_groupbox.setLayout(convert_layout)
        layout.addWidget(input_groupbox); layout.addWidget(output_groupbox); layout.addWidget(convert_groupbox); layout.addStretch()
        return tab

    def create_csv_to_xml_tab(self):
        tab = QWidget(); layout = QVBoxLayout(tab)
        input_groupbox = QGroupBox("1. Pilih File CSV Input"); input_layout = QHBoxLayout(); self.txt_csv_input = QLineEdit(); self.txt_csv_input.setReadOnly(True); btn_select_csv = QPushButton("..."); btn_select_csv.setFixedWidth(40); btn_select_csv.clicked.connect(self.select_csv_input); input_layout.addWidget(self.txt_csv_input); input_layout.addWidget(btn_select_csv); input_groupbox.setLayout(input_layout)
        output_groupbox = QGroupBox("2. Tentukan Folder Output untuk File XML"); output_layout = QHBoxLayout(); self.txt_csv_to_xml_output = QLineEdit(); self.txt_csv_to_xml_output.setReadOnly(True); btn_select_dir = QPushButton("..."); btn_select_dir.setFixedWidth(40); btn_select_dir.clicked.connect(self.select_csv_to_xml_output_dir); output_layout.addWidget(self.txt_csv_to_xml_output); output_layout.addWidget(btn_select_dir); output_groupbox.setLayout(output_layout)
        convert_groupbox = QGroupBox("3. Jalankan Konversi"); convert_layout = QVBoxLayout(); btn_convert = QPushButton("Konversi ke XML"); btn_convert.setStyleSheet("font-size: 16px; padding: 10px; background-color: #008CBA; color: white;"); btn_convert.clicked.connect(self.run_csv_to_xml_conversion); convert_layout.addWidget(btn_convert); convert_groupbox.setLayout(convert_layout)
        layout.addWidget(input_groupbox); layout.addWidget(output_groupbox); layout.addWidget(convert_groupbox); layout.addStretch()
        return tab
    
    def select_xml_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Pilih File XML", "", "XML Files (*.xml)");
        if files:
            for f in files:
                if f not in self.xml_files: self.xml_files.append(f)
            self.list_widget_files.clear(); self.list_widget_files.addItems([os.path.basename(f) for f in self.xml_files])
    def remove_selected_files(self):
        selected_items = self.list_widget_files.selectedItems()
        if not selected_items: return
        files_to_remove = {item.text() for item in selected_items}; self.xml_files = [p for p in self.xml_files if os.path.basename(p) not in files_to_remove]; self.list_widget_files.clear(); self.list_widget_files.addItems([os.path.basename(f) for f in self.xml_files])
    def select_xml_to_csv_output(self):
        output_file, _ = QFileDialog.getSaveFileName(self, "Simpan File CSV", "", "CSV Files (*.csv)")
        if output_file: self.xml_to_csv_output_file = output_file; self.txt_xml_to_csv_output.setText(self.xml_to_csv_output_file)
    def run_xml_to_csv_conversion(self):
        if not self.xml_files: QMessageBox.warning(self, "Input Tidak Lengkap", "Pilih file XML."); return
        if not self.xml_to_csv_output_file: QMessageBox.warning(self, "Input Tidak Lengkap", "Tentukan file CSV output."); return
        if not self.active_profile_path or not os.path.exists(self.active_profile_path): QMessageBox.warning(self, "Profil Tidak Valid", "Profil yang dipilih tidak ada."); return
        self.statusBar.showMessage("Memproses XML ke CSV..."); QApplication.processEvents()
        try:
            with open(self.active_profile_path, 'r', encoding='utf-8') as f: config = json.load(f)
            self.xml_to_csv_logic(self.xml_files, self.xml_to_csv_output_file, config)
            self.statusBar.showMessage("Konversi XML ke CSV berhasil!"); QMessageBox.information(self, "Sukses", f"Data berhasil dikonversi ke:\n{self.xml_to_csv_output_file}")
        except Exception as e: self.statusBar.showMessage(f"Kesalahan: {e}"); QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat konversi:\n{e}")

    def select_csv_input(self):
        file, _ = QFileDialog.getOpenFileName(self, "Pilih File CSV", "", "CSV Files (*.csv)");
        if file: self.csv_file_input = file; self.txt_csv_input.setText(file)
    def select_csv_to_xml_output_dir(self):
        dir = QFileDialog.getExistingDirectory(self, "Pilih Folder Output")
        if dir: self.csv_to_xml_output_dir = dir; self.txt_csv_to_xml_output.setText(dir)
    def run_csv_to_xml_conversion(self):
        if not self.csv_file_input: QMessageBox.warning(self, "Input Tidak Lengkap", "Pilih file CSV."); return
        if not self.csv_to_xml_output_dir: QMessageBox.warning(self, "Input Tidak Lengkap", "Tentukan folder output."); return
        if not self.active_profile_path or not os.path.exists(self.active_profile_path): QMessageBox.warning(self, "Profil Tidak Valid", "Profil yang dipilih tidak ada."); return
        self.statusBar.showMessage("Memproses CSV ke XML..."); QApplication.processEvents()
        try:
            with open(self.active_profile_path, 'r', encoding='utf-8') as f: config = json.load(f)
            count = self.csv_to_xml_logic(self.csv_file_input, self.csv_to_xml_output_dir, config)
            self.statusBar.showMessage("Konversi CSV ke XML berhasil!"); QMessageBox.information(self, "Sukses", f"{count} file XML berhasil dibuat di folder:\n{self.csv_to_xml_output_dir}")
        except Exception as e: self.statusBar.showMessage(f"Kesalahan: {e}"); QMessageBox.critical(self, "Error", f"Terjadi kesalahan saat konversi:\n{e}")
    
    def load_profiles(self):
        self.combo_profiles.blockSignals(True); self.combo_profiles.clear()
        profile_files = QDir(PROFILE_DIR).entryList([f"*{PROFILE_EXTENSION}"], QDir.Filter.Files, QDir.SortFlag.Name)
        if not profile_files:
            default_path = os.path.join(PROFILE_DIR, f"default{PROFILE_EXTENSION}")
            with open(default_path, 'w', encoding='utf-8') as f: json.dump(get_default_profile_config(), f, indent=4)
            self.combo_profiles.addItem(os.path.basename(default_path), default_path)
        else:
            for file_name in profile_files: self.combo_profiles.addItem(file_name, os.path.join(PROFILE_DIR, file_name))
        self.combo_profiles.blockSignals(False)
        if self.combo_profiles.count() > 0: self.on_profile_changed(self.combo_profiles.currentText())
    def on_profile_changed(self, text):
        if not text: return
        self.active_profile_path = self.combo_profiles.currentData(); self.statusBar.showMessage(f"Profil '{text}' telah aktif.")
    def open_settings(self):
        if not self.active_profile_path: QMessageBox.warning(self, "Tidak Ada Profil", "Pilih profil terlebih dahulu."); return
        dialog = MappingDialog(self.active_profile_path, self)
        if dialog.exec() == QDialog.DialogCode.Accepted: self.load_profiles()

    def xml_to_csv_logic(self, xml_files, output_csv_file, config):
        columns_config = config["columns"]; settings_config = config["settings"]; csv_header = list(columns_config.keys()); rows = []; item_loop_path = settings_config.get("item_loop_path", ".")
        for xml_file in xml_files:
            tree = ET.parse(xml_file); root = tree.getroot()
            root_data = {col: self.extract_value(root, details) for col, details in columns_config.items() if details.get("source") == "root"}
            for item in root.findall(item_loop_path):
                row_data = {}
                for col_name, details in columns_config.items():
                    if details.get("type") != "calculated" and details.get("source") == "item": row_data[col_name] = self.extract_value(item, details)
                for col_name, details in columns_config.items():
                    if details.get("type") == "calculated": row_data[col_name] = self.extract_value(None, details, row_data)
                row_data.update(root_data); rows.append([row_data.get(h, "") for h in csv_header])
        with open(output_csv_file, "w", newline='', encoding='utf-8') as f:
            writer = csv.writer(f); writer.writerow([f"Conversion Date: {datetime.now().strftime('%Y-%m-%d')}"]); writer.writerow(csv_header); writer.writerows(rows)
    def extract_value(self, base_element, details, current_row_data=None):
        col_type = details.get("type"); path = details.get("path", "");
        if base_element is None and col_type != "calculated": return ""
        if col_type == "xpath": el = base_element.find(path); return el.text.strip() if el is not None and el.text is not None else ""
        elif col_type == "attribute": return base_element.get(path, "")
        elif col_type == "xpath_indexed":
            elements = base_element.findall(path); index = details.get("index", 0)
            if len(elements) > index and elements[index].text is not None: return elements[index].text.strip()
            return ""
        elif col_type == "calculated":
            formula = details.get("formula", "");
            if current_row_data is None: return "Error: Data for calc not found"
            for key, val in current_row_data.items():
                if f"{{{key}}}" in formula and val:
                    try: formula = formula.replace(f"{{{key}}}", str(float(val)))
                    except (ValueError, TypeError): return "Error: Invalid number"
            try: return f"{eval(formula):.2f}"
            except: return "Error: Formula failed"
        return ""
    
    def csv_to_xml_logic(self, csv_file_path, output_dir, config):
        settings = config.get("settings", {})
        columns = config.get("columns", {})
        group_key = settings.get("xml_grouping_key")
        root_name = settings.get("root_element_name", "root")
        item_loop_path = settings.get("item_loop_path", ".//item")
        item_tag = item_loop_path.split('/')[-1]

        if not group_key: raise ValueError("`xml_grouping_key` tidak diatur di profil.")
        
        grouped_data = defaultdict(list)
        with open(csv_file_path, mode='r', encoding='utf-8') as f:
            first_line = f.readline()
            if not first_line.startswith("Conversion Date"):
                f.seek(0)

            reader = csv.DictReader(f)
            if group_key not in reader.fieldnames:
                raise ValueError(f"Kolom '{group_key}' untuk pengelompokan tidak ditemukan di file CSV.")

            for row in reader:
                key_value = row.get(group_key)
                if key_value:
                    grouped_data[key_value].append(row)

        file_count = 0
        for group_id, items in grouped_data.items():
            root_el = ET.Element(root_name)
            
            first_item_data = items[0]
            for col_name, details in columns.items():
                if details.get("source") == "root" and details.get("type") != "calculated":
                    self._insert_value(root_el, details, first_item_data.get(col_name))

            items_parent_el = root_el 
            
            for item_data in items:
                item_el = ET.SubElement(items_parent_el, item_tag)
                for col_name, details in columns.items():
                    if details.get("source") == "item" and details.get("type") != "calculated":
                        self._insert_value(item_el, details, item_data.get(col_name))

            tree = ET.ElementTree(root_el)
            ET.indent(tree, space="  ", level=0) 
            safe_group_id = "".join(c for c in group_id if c.isalnum() or c in (' ', '.', '_')).rstrip()
            output_filename = os.path.join(output_dir, f"{safe_group_id}.xml")
            tree.write(output_filename, encoding='utf-8', xml_declaration=True)
            file_count += 1
        
        return file_count

    def _insert_value(self, base_element, details, value):
        if value is None or str(value).strip() == '':
            return

        col_type = details.get("type")
        path_str = details.get("path", "")
        
        if col_type == "attribute":
            base_element.set(path_str, str(value))
        elif col_type in ["xpath", "xpath_indexed"]:
            current_el = base_element
            path_parts = path_str.strip("./").split("/")
            
            for i, part in enumerate(path_parts):
                tag_name = part
                predicate = {}
                if '[' in part and part.endswith(']'):
                    tag_name = part.split('[', 1)[0]
                    pred_str = part[len(tag_name)+1:-1]
                    if '=' in pred_str:
                        attr_name, attr_val = pred_str.split('=', 1)
                        attr_name = attr_name.strip().lstrip('@')
                        attr_val = attr_val.strip("'\"")
                        predicate = {attr_name: attr_val}

                found_el = None
                for child in current_el.findall(tag_name):
                    match = True
                    if predicate:
                        for p_key, p_val in predicate.items():
                            if child.get(p_key) != p_val:
                                match = False
                                break
                    if match:
                        found_el = child
                        break

                if found_el is None:
                    found_el = ET.SubElement(current_el, tag_name)
                    for p_key, p_val in predicate.items():
                        found_el.set(p_key, p_val)
                
                current_el = found_el
            current_el.text = str(value)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    main_window = XMLConverterApp()
    main_window.show()
    sys.exit(app.exec())