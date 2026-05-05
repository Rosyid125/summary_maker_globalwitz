# AGENTS.md - System Context for Future AI Agents

Dokumen ini adalah konteks teknis menyeluruh untuk agent berikutnya agar bisa langsung bekerja tanpa mengulang eksplorasi dari nol.

## 1) Tujuan Sistem

Aplikasi ini adalah desktop app Python (Tkinter) untuk:
- membaca data impor dari file Excel,
- melakukan normalisasi + agregasi data dengan logika yang meniru versi JavaScript lama,
- menghasilkan file Excel summary multi-sheet dengan formatting tabel yang kompleks.

Nama aplikasi: **Excel Summary Maker - GlobalWitz X Volza**.

## 2) Arsitektur Tingkat Tinggi

### Entry point
- `main.py` membuat window Tkinter, setup logger, lalu mount `MainWindow`.

### Layer utama
- **GUI layer**: `src/gui/main_window.py`
- **Core processing (aktif di runtime)**:
  - `src/core/js_excel_reader.py`
  - `src/core/js_processor.py`
  - `src/core/js_output_formatter.py`
- **Core processing (legacy/alternatif, saat ini bukan jalur utama)**:
  - `src/core/excel_reader.py`
  - `src/core/data_aggregator.py` (dipakai oleh `js_processor` untuk fungsi agregasi inti)
  - `src/core/output_formatter.py` (formatter openpyxl klasik, bukan writer utama saat proses GUI)
- **Utilities**:
  - `src/utils/helpers.py`
  - `src/utils/constants.py`
  - `src/utils/logger.py`
  - `src/utils/settings.py` - Settings manager untuk default column mappings

### Mode operasi
- **Development mode**: jalankan `python main.py`.
- **Frozen/built mode** (PyInstaller): path input/output disesuaikan memakai `sys.frozen`.

## 3) Struktur Folder dan Peran

- `main.py`: bootstrap aplikasi.
- `src/`: source code.
- `original_excel/`: folder input default.
- `processed_excel/`: folder output default.
- `logs/`: log harian.
- `build.bat`: script build PyInstaller (Windows).
- `README.md`: dokumentasi pengguna.
- `DISTRIBUTION.md`: panduan distribusi executable.

Catatan `.gitignore`:
- `original_excel/`, `processed_excel/`, `logs/`, `__pycache__/` di-ignore.

## 4) Alur Data End-to-End (Jalur yang Dipakai GUI)

1. User pilih file + sheet di tab GUI.
2. GUI mengumpulkan `column_mapping`, `date_format`, `number_format`, `target_year`, `incoterm`, mode, dll.
3. `JSStyleExcelReader.read_and_preprocess_data(...)` membaca sheet via pandas.
4. Reader mengubah setiap baris menjadi schema internal JavaScript-style.
5. `JSStyleProcessor.process_data_like_javascript(...)`:
   - optional swap importer<->supplier jika `supplier_as_sheet == "ya"`,
   - split data valid importer vs blank importer,
   - group per importer (atau per supplier setelah swap),
   - untuk tiap group memanggil `DataAggregator.perform_aggregation(...)`,
   - menyusun blok tabel output row-based.
6. `js_output_formatter.OutputFormatter.write_output_to_file(...)` menulis workbook dengan `xlsxwriter` (merge cells, warna quarter, section totals).
7. GUI menampilkan sukses + path output.

## 5) Kontrak Data Internal (Penting)

### 5.1 Raw row schema dari `js_excel_reader`
Setiap row hasil preprocessing punya keys berikut:
- `month` (str, contoh: `Jan`, `Feb`, ..., `Des`, atau `-`)
- `hsCode`
- `itemDesc`
- `gsm`
- `item`
- `addOn`
- `importer`
- `supplier`
- `originCountry`
- `incoterms`
- `usdQtyUnit` (float)
- `qty` (float)

Nilai kosong distandarkan banyaknya ke `"-"` (string), bukan `None`.

### 5.2 Output agregator
`DataAggregator.perform_aggregation(...)` mengembalikan:
- `summaryLvl1`: agregasi bulanan per kombinasi `(month, hsCode, item, gsm, addOn)`
  - fields: `month`, `hsCode`, `item`, `gsm`, `addOn`, `avgPrice`, `totalQty`
- `summaryLvl2`: rekap lintas bulan per kombinasi `(hsCode, item, gsm, addOn)`
  - fields: `hsCode`, `item`, `gsm`, `addOn`, `avgOfSummaryPrice`, `totalOfSummaryQty`

### 5.3 Workbook intermediate (sebelum ditulis)
`js_processor` membentuk list `workbook_data_for_excel_js`, tiap elemen berupa:
- `name`: nama sheet,
- `allRowsForSheetContent`: matriks row/col,
- `supplierGroupsMeta`: metadata block per supplier/importer,
- `totalColumns`: jumlah kolom total.

`totalColumns = 5 + (12 * 2) + 3 = 32`.

## 6) Detail GUI dan State

File: `src/gui/main_window.py`

### Tab
- File Selection
- Configuration
- Column Mapping
- Processing

### Settings Dialog (bukan tab, terpisah di pojok kanan atas)
- Diakses via tombol "⚙ Settings" di header window
- Berisi Default Mapping Set untuk konfigurasi kolom default
- Auto-apply option untuk otomatis mapping saat load file
- Dapat export/import konfigurasi ke file JSON

### State penting
- File/sheet: `current_file_path`, `selected_sheet`
- Parse options: `date_format`, `number_format`
- Output controls: `target_year`, `incoterm`, `incoterm_mode`, `output_filename`
- Mode bisnis: `supplier_as_sheet` (`ya`/`tidak`)
- Mapping: dict `column_mappings` untuk field wajib/opsional.

### Validasi proses
Saat klik Start:
- file harus ada,
- sheet harus dipilih,
- output filename harus ada,
- minimal 3 kolom sudah ter-mapping.

### Threading
- Proses jalan di thread background (`threading.Thread(..., daemon=True)`).
- Update UI via `root.after(...)`.

## 7) Parsing Logic

### Date parsing
`JSStyleExcelReader` mendukung:
- DD/MM/YYYY,
- MM/DD/YYYY,
- DD-MONTH-YYYY (EN + ID month names),
- ISO (`YYYY-MM-DD`),
- Excel serial date.

Mode `auto` mencoba berurutan:
1) DD-MONTH-YYYY, 2) DD/MM/YYYY, 3) MM/DD/YYYY.

### Number parsing
`parse_number(value, number_format)`:
- format `AMERICAN`: koma ribuan, titik desimal,
- selain itu dianggap `EUROPEAN`: titik ribuan, koma desimal.

Catatan: GUI mengirim `"american"` / `"european"` lowercase. Di `js_excel_reader`, pemeriksaan memakai uppercase (`"AMERICAN"`), sehingga selain uppercase akan jatuh ke jalur European. Ini behavior aktual yang perlu diperhatikan saat debugging angka.

## 8) Aggregation Logic (Business Core)

`DataAggregator.perform_aggregation(...)` melakukan:
- skip row jika `month` atau `hsCode` kosong / `-`,
- grouping key: `month-hsCode-item-gsm-addOn`,
- `avgPrice` dihitung dari nilai `usdQtyUnit` yang > 0 saja,
- `totalQty` menjumlah `qty` numeric,
- rekap level 2 menyatukan lintas bulan dan ambil rata-rata dari average bulanan (>0).

Fungsi bantu kunci: `average_greater_than_zero(...)` di `helpers.py`.

## 9) Output Excel Logic

Writer utama: `src/core/js_output_formatter.py` (xlsxwriter).

### Ciri format output
- Header period merge 1 baris (`<year> PERIODE`) jika `period_year` tersedia.
- Block per group (supplier/importer) dengan:
  - 5 kolom identitas,
  - 12 bulan x (PRICE,QTY),
  - RECAP (AVG PRICE, INCOTERM, TOTAL QTY).
- Warna quarter:
  - Q1 `#FFC000`
  - Q2 `#00B050`
  - Q3 `#FFFF00`
  - Q4 `#00B0F0`
- Section tambahan:
  - `TOTAL PER ITEM`
  - `TOTAL ALL SUPPLIER/IMPORTER PER MO`
  - `TOTAL ALL SUPPLIER/IMPORTER PER QUARTAL`

### Incoterm mode
- `manual`: pakai `global_incoterm` untuk semua kombinasi.
- `from_column`: ambil incoterm dari row pertama yang match kombinasi, lalu potong 3 huruf uppercase pertama.

### Sheet name safety
- sanitasi invalid chars,
- max 31 chars,
- deduplikasi case-insensitive (`_1`, `_2`, dst) jika bentrok.

## 10) Pathing dan Runtime Environment

`src/utils/constants.py`:
- `get_app_data_dir()`:
  - frozen: folder executable,
  - dev: root repo.
- `get_safe_output_dir()` selalu menunjuk ke `<app_dir>/processed_excel` dan memastikan folder ada.

`DEFAULT_INPUT_FOLDER` = `<app_dir>/original_excel`.
`DEFAULT_OUTPUT_FOLDER` = `<app_dir>/processed_excel`.

## 11) Logging dan Observability

`setup_logger()` membuat:
- console handler,
- file handler `logs/excel_summary_maker_YYYYMMDD.log`.

GUI juga menampilkan log stream realtime di tab Processing (`log_text`).

## 12) Dependensi

`requirements.txt`:
- `numpy>=1.24.0,<2.0.0`
- `pandas>=2.1.0`
- `openpyxl==3.1.2`
- `xlsxwriter==3.1.9`
- `python-dateutil==2.8.2`
- `tkinter-tooltip==2.1.0`
- `pyinstaller>=5.0`

## 13) Build/Distribusi

`build.bat`:
- cek python + pyinstaller,
- install requirements,
- bersihkan build lama,
- jalankan PyInstaller `--onedir --windowed` dengan data folders,
- hasil utama: `dist/ExcelSummaryMaker/ExcelSummaryMaker.exe`.

`DISTRIBUTION.md` menjelaskan paket distribusi folder-based (portable).

## 14) Temuan Penting / Potensi Masalah

1. **Dua formatter dengan nama class sama**
   - `src/core/output_formatter.py` (openpyxl) dan `src/core/js_output_formatter.py` (xlsxwriter) sama-sama bernama `OutputFormatter`.
   - Jalur aktif `js_processor` memakai formatter xlsxwriter.

2. **Duplikasi method di `js_output_formatter.py`**
   - `extract_incoterm_from_value` dan `get_incoterm_for_combination` terdefinisi dua kali (awal dan akhir file).
   - Python akan memakai definisi terakhir.

3. **Mismatch case pada `number_format`**
   - GUI nilai lowercase (`american/european`), parser cek uppercase (`AMERICAN`).
   - Efek: jalur American bisa tidak aktif jika value tidak dinormalisasi.

4. **UI minor issue**
   - `auto_map_btn.pack(...)` dipanggil dua kali pada setup mapping tab.

5. **Redundant operation**
   - `show_sheet_info()` memanggil `self.info_text.delete(...)` dua kali berurutan.

6. **Kode legacy masih diinisialisasi**
   - `MainWindow` membuat instance `ExcelReader`, `DataAggregator`, `OutputFormatter` klasik, namun alur proses utama memakai JS-style classes.

7. **Kontradiksi kecil dokumen build**
   - Beberapa teks di `build.bat` menyebut output ke Documents, sementara constants menunjuk `processed_excel` di app dir.

## 15) Panduan Kerja untuk Agent Selanjutnya

Jika ingin modifikasi fitur, ikuti prinsip ini:

- Untuk behavior proses utama, fokus di:
  - `src/gui/main_window.py`
  - `src/core/js_excel_reader.py`
  - `src/core/js_processor.py`
  - `src/core/js_output_formatter.py`
  - `src/core/data_aggregator.py`
- Jaga kompatibilitas schema row (`hsCode`, `addOn`, `usdQtyUnit`, dst) karena saling terikat ketat.
- Hati-hati mengubah urutan/posisi kolom output; formatter mengandalkan indeks kolom statis dan merge ranges.
- Saat ubah parsing number/date, verifikasi dengan data nyata yang punya format campuran.
- Jika membersihkan tech debt, pertimbangkan refactor bertahap agar tidak merusak kesesuaian output terhadap versi JS lama.

## 16) Checklist Verifikasi Cepat Setelah Perubahan

1. Jalankan `python main.py`.
2. Load file contoh di `original_excel/`.
3. Cek mapping otomatis dan manual.
4. Proses mode:
   - `incoterm_mode=manual`
   - `incoterm_mode=from_column`
   - `supplier_as_sheet=tidak`
   - `supplier_as_sheet=ya`
5. Buka output di `processed_excel/` dan cek:
   - jumlah sheet,
   - merge header,
   - quarter colors,
   - section total per month/quarter,
   - nilai numerik tetap numerik di Excel (bukan text).
6. Cek log file harian di `logs/` untuk warning/error.

---

Dokumen ini dibuat berdasarkan codebase saat ini dan dimaksudkan sebagai baseline konteks teknis penuh untuk AI agent berikutnya.
