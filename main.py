"""
main.py
Aplikasi Utama Autofill POP SIPGN — SPPG Kalisapu (Mode Cepat & Aman)
Membaca data Excel, memvalidasi porsi, dan menghasilkan kode otomatisasi instan untuk browser.
"""

import sys
import os
import json
import datetime
import logging
from typing import List, Dict, Any, Optional

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd

from excel_reader import ExcelReader
from validator import validate_dataset


class TextHandler(logging.Handler):
    """Handler logging khusus untuk mengarahkan pesan log ke ScrolledText Tkinter."""
    def __init__(self, text_widget: scrolledtext.ScrolledText):
        super().__init__()
        self.text_widget = text_widget

    def emit(self, record):
        msg = self.format(record)
        def append():
            try:
                self.text_widget.configure(state='normal')
                tag = record.levelname
                self.text_widget.insert(tk.END, msg + '\n', tag)
                self.text_widget.see(tk.END)
                self.text_widget.configure(state='disabled')
            except Exception:
                pass
        self.text_widget.after(0, append)


class PopSipgnApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("POP SIPGN Autofill — SPPG Kalisapu (Langkah 2: Porsi per KPM)")
        self.root.geometry("1020x720")
        self.root.minsize(920, 620)

        # Inisialisasi folder logs dan reports
        os.makedirs("logs", exist_ok=True)
        os.makedirs("reports", exist_ok=True)

        # Setup Logging
        self.log_filename = os.path.join(
            "logs", f"autofill_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        )
        self.logger = logging.getLogger("pop_sipgn")
        self.logger.setLevel(logging.INFO)
        file_handler = logging.FileHandler(self.log_filename, encoding='utf-8')
        file_formatter = logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%H:%M:%S")
        file_handler.setFormatter(file_formatter)
        self.logger.addHandler(file_handler)

        # State Aplikasi
        self.excel_reader: Optional[ExcelReader] = None
        self.current_file_path: str = ""
        self.raw_columns: List[str] = []
        self.valid_excel_records: List[Dict[str, Any]] = []
        self.invalid_excel_records: List[Dict[str, Any]] = []

        # Bangun Tampilan GUI
        self._build_ui()
        self.logger.info("Aplikasi Autofill POP SIPGN SPPG Kalisapu (Mode Praktis) dimulai.")

    def _build_ui(self):
        # Konfigurasi Tema & Gaya
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TLabel', font=('Segoe UI', 9))
        style.configure('Header.TLabel', font=('Segoe UI', 12, 'bold'), foreground='#1E3A8A')
        style.configure('SubHeader.TLabel', font=('Segoe UI', 9, 'italic'), foreground='#4B5563')
        style.configure('Card.TLabelframe', relief='groove', padding=8)
        style.configure('Card.TLabelframe.Label', font=('Segoe UI', 9, 'bold'), foreground='#1E40AF')
        style.configure('Primary.TButton', font=('Segoe UI', 9, 'bold'), foreground='#FFFFFF', background='#2563EB')
        style.configure('Action.TButton', font=('Segoe UI', 9, 'bold'))

        # Header Banner
        header_frame = ttk.Frame(self.root, padding="10 8 10 5")
        header_frame.pack(fill=tk.X)
        ttk.Label(header_frame, text="AUTOFILL POP SIPGN — SPPG KALISAPU", style='Header.TLabel').pack(anchor=tk.W)
        ttk.Label(
            header_frame,
            text="Pengisian Cepat & Aman Formulir 'Jumlah Porsi per KPM' (Langkah 2 dari 4) Berbasis Excel",
            style='SubHeader.TLabel'
        ).pack(anchor=tk.W)

        # Main Paned Window
        main_paned = ttk.PanedWindow(self.root, orient=tk.VERTICAL)
        main_paned.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        top_container = ttk.Frame(main_paned)
        main_paned.add(top_container, weight=4)

        bottom_container = ttk.Frame(main_paned)
        main_paned.add(bottom_container, weight=2)

        # -------------------------------------------------------------
        # 1. BAGIAN SUMBER DATA & PEMETAAN EXCEL
        # -------------------------------------------------------------
        ctrl_frame = ttk.LabelFrame(top_container, text="1. Sumber Data Excel & Pemetaan Kolom", style='Card.TLabelframe')
        ctrl_frame.pack(fill=tk.X, pady=(0, 6))

        # Baris 1: Pemilihan File
        f_row = ttk.Frame(ctrl_frame)
        f_row.pack(fill=tk.X, pady=3)
        ttk.Label(f_row, text="File Excel (.xlsx):", width=16).pack(side=tk.LEFT)
        self.lbl_excel_path = ttk.Entry(f_row)
        self.lbl_excel_path.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        ttk.Button(f_row, text="Pilih File...", command=self._choose_excel_file).pack(side=tk.LEFT)

        # Baris 2: Pemilihan Sheet & Pemetaan Kolom
        map_row = ttk.Frame(ctrl_frame)
        map_row.pack(fill=tk.X, pady=4)

        ttk.Label(map_row, text="Pilih Sheet:").pack(side=tk.LEFT)
        self.cmb_sheet = ttk.Combobox(map_row, state="readonly", width=18)
        self.cmb_sheet.pack(side=tk.LEFT, padx=(3, 10))
        self.cmb_sheet.bind("<<ComboboxSelected>>", self._on_sheet_selected)

        ttk.Label(map_row, text="Kolom Sekolah:").pack(side=tk.LEFT)
        self.cmb_col_school = ttk.Combobox(map_row, state="readonly", width=18)
        self.cmb_col_school.pack(side=tk.LEFT, padx=(3, 10))

        ttk.Label(map_row, text="Porsi Besar:").pack(side=tk.LEFT)
        self.cmb_col_besar = ttk.Combobox(map_row, state="readonly", width=14)
        self.cmb_col_besar.pack(side=tk.LEFT, padx=(3, 10))

        ttk.Label(map_row, text="Porsi Kecil:").pack(side=tk.LEFT)
        self.cmb_col_kecil = ttk.Combobox(map_row, state="readonly", width=14)
        self.cmb_col_kecil.pack(side=tk.LEFT, padx=(3, 10))

        btn_load = ttk.Button(map_row, text="📥 Muat Data", command=self._load_and_validate_excel, style='Action.TButton')
        btn_load.pack(side=tk.LEFT, padx=6)

        # -------------------------------------------------------------
        # 2. BAGIAN AKSI CEPAT & PRATINJAU DATA
        # -------------------------------------------------------------
        action_frame = ttk.LabelFrame(top_container, text="2. Eksekusi Pengisian & Pratinjau Data", style='Card.TLabelframe')
        action_frame.pack(fill=tk.BOTH, expand=True)

        # Tombol Aksi Utama di Atas Tabel
        action_btn_row = ttk.Frame(action_frame)
        action_btn_row.pack(fill=tk.X, pady=(2, 6))

        self.btn_copy_code = ttk.Button(
            action_btn_row,
            text="📋 SALIN KODE AUTOFILL (PASTE DI CONSOLE BROWSER)",
            command=self._copy_quick_script,
            style='Action.TButton'
        )
        self.btn_copy_code.pack(side=tk.LEFT, padx=3)

        ttk.Button(
            action_btn_row,
            text="💾 Ekspor Data (CSV)",
            command=self._export_report
        ).pack(side=tk.LEFT, padx=5)

        # Petunjuk Singkat
        hint_lbl = ttk.Label(
            action_btn_row,
            text="👉 Cara Pakai: Buka POP SIPGN Langkah 2 ➔ Tekan F12 (Console) ➔ Tekan Ctrl+V lalu Enter.",
            foreground='#15803D',
            font=('Segoe UI', 9, 'italic')
        )
        hint_lbl.pack(side=tk.LEFT, padx=10)

        # Tabel Pratinjau Data
        columns = ("row", "school", "besar", "kecil", "status")
        self.tree_excel = ttk.Treeview(action_frame, columns=columns, show="headings", height=8)
        self.tree_excel.heading("row", text="Baris")
        self.tree_excel.heading("school", text="Nama Sekolah / KPM")
        self.tree_excel.heading("besar", text="Porsi Besar")
        self.tree_excel.heading("kecil", text="Porsi Kecil")
        self.tree_excel.heading("status", text="Status Validasi")

        self.tree_excel.column("row", width=60, anchor=tk.CENTER)
        self.tree_excel.column("school", width=420)
        self.tree_excel.column("besar", width=100, anchor=tk.CENTER)
        self.tree_excel.column("kecil", width=100, anchor=tk.CENTER)
        self.tree_excel.column("status", width=220)

        tree_scroll = ttk.Scrollbar(action_frame, orient=tk.VERTICAL, command=self.tree_excel.yview)
        self.tree_excel.configure(yscrollcommand=tree_scroll.set)
        self.tree_excel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # -------------------------------------------------------------
        # 3. BAGIAN BAWAH: RINGKASAN & LOG AKTIVITAS
        # -------------------------------------------------------------
        self._build_bottom_section(bottom_container)

    def _build_bottom_section(self, parent: ttk.Frame):
        summary_frame = ttk.LabelFrame(parent, text="Ringkasan Data", style='Card.TLabelframe')
        summary_frame.pack(fill=tk.X, pady=(0, 4))

        sum_row = ttk.Frame(summary_frame)
        sum_row.pack(fill=tk.X, padx=5, pady=2)

        self.lbl_sum_total = ttk.Label(sum_row, text="Total Sekolah Valid: 0", font=('Segoe UI', 9, 'bold'), foreground='#047857')
        self.lbl_sum_total.pack(side=tk.LEFT, padx=12)

        self.lbl_sum_besar = ttk.Label(sum_row, text="Total Porsi Besar: 0", font=('Segoe UI', 9, 'bold'), foreground='#1E40AF')
        self.lbl_sum_besar.pack(side=tk.LEFT, padx=12)

        self.lbl_sum_kecil = ttk.Label(sum_row, text="Total Porsi Kecil: 0", font=('Segoe UI', 9, 'bold'), foreground='#7C3AED')
        self.lbl_sum_kecil.pack(side=tk.LEFT, padx=12)

        self.lbl_sum_invalid = ttk.Label(sum_row, text="Baris Dilewati: 0", font=('Segoe UI', 9, 'bold'), foreground='#DC2626')
        self.lbl_sum_invalid.pack(side=tk.LEFT, padx=12)

        log_frame = ttk.LabelFrame(parent, text="Log Aktivitas & Audit Trail", style='Card.TLabelframe')
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=6, font=('Consolas', 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)
        self.log_text.tag_config('INFO', foreground='#111827')
        self.log_text.tag_config('WARNING', foreground='#B45309')
        self.log_text.tag_config('ERROR', foreground='#B91C1C')
        self.log_text.configure(state='disabled')

        text_handler = TextHandler(self.log_text)
        text_handler.setFormatter(logging.Formatter("[%(asctime)s] [%(levelname)s] %(message)s", "%H:%M:%S"))
        self.logger.addHandler(text_handler)

    # -------------------------------------------------------------
    # LOGIKA PENGOLAHAN EXCEL
    # -------------------------------------------------------------
    def _choose_excel_file(self):
        file_path = filedialog.askopenfilename(
            title="Pilih File Excel (.xlsx)",
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if not file_path:
            return

        self.current_file_path = file_path
        self.lbl_excel_path.delete(0, tk.END)
        self.lbl_excel_path.insert(0, file_path)

        try:
            self.excel_reader = ExcelReader(file_path)
            sheets = self.excel_reader.get_sheet_names()
            self.cmb_sheet['values'] = sheets
            if sheets:
                oct_sheet = next((s for s in sheets if 'oktober' in s.lower()), sheets[0])
                self.cmb_sheet.set(oct_sheet)
                self._on_sheet_selected()
            self.logger.info(f"File Excel dimuat: {os.path.basename(file_path)}. Sheets: {sheets}")
        except Exception as e:
            messagebox.showerror("Kesalahan", f"Gagal membaca file Excel:\n{str(e)}")
            self.logger.error(f"Gagal membaca file Excel: {str(e)}")

    def _on_sheet_selected(self, event=None):
        sheet_name = self.cmb_sheet.get()
        if not self.excel_reader or not sheet_name:
            return

        try:
            cols = self.excel_reader.get_columns(sheet_name)
            self.raw_columns = cols
            self.cmb_col_school['values'] = cols
            self.cmb_col_besar['values'] = cols
            self.cmb_col_kecil['values'] = cols

            s_col, b_col, k_col = ExcelReader.auto_detect_columns(cols)
            if s_col:
                self.cmb_col_school.set(s_col)
            if b_col:
                self.cmb_col_besar.set(b_col)
            if k_col:
                self.cmb_col_kecil.set(k_col)

            self.logger.info(f"Sheet '{sheet_name}' dipilih. Kolom terdeteksi: {cols}")
        except Exception as e:
            messagebox.showerror("Kesalahan", f"Gagal membaca sheet {sheet_name}:\n{str(e)}")
            self.logger.error(f"Gagal membaca sheet: {str(e)}")

    def _load_and_validate_excel(self):
        if not self.excel_reader:
            messagebox.showwarning("Peringatan", "Silakan pilih file Excel terlebih dahulu.")
            return

        sheet_name = self.cmb_sheet.get()
        s_col = self.cmb_col_school.get()
        b_col = self.cmb_col_besar.get()
        k_col = self.cmb_col_kecil.get()

        if not all([sheet_name, s_col, b_col, k_col]):
            messagebox.showwarning("Peringatan", "Pastikan sheet dan seluruh pemetaan kolom sudah ditentukan.")
            return

        try:
            df = self.excel_reader.load_sheet_data(sheet_name)
            valid_recs, invalid_recs = validate_dataset(df, s_col, b_col, k_col)
            self.valid_excel_records = valid_recs
            self.invalid_excel_records = invalid_recs

            for item in self.tree_excel.get_children():
                self.tree_excel.delete(item)

            tot_besar = sum(item['porsi_besar'] for item in valid_recs)
            tot_kecil = sum(item['porsi_kecil'] for item in valid_recs)

            for item in valid_recs:
                self.tree_excel.insert("", tk.END, values=(
                    item['row_index'],
                    item['nama_sekolah'],
                    item['porsi_besar'],
                    item['porsi_kecil'],
                    "Valid (Siap Diisi)"
                ))

            for item in invalid_recs:
                self.tree_excel.insert("", tk.END, values=(
                    item['row_index'],
                    item['nama_sekolah_raw'],
                    item['porsi_besar_raw'],
                    item['porsi_kecil_raw'],
                    f"Dilewati: {item['alasan']}"
                ))

            self.lbl_sum_total.config(text=f"Total Sekolah Valid: {len(valid_recs)}")
            self.lbl_sum_besar.config(text=f"Total Porsi Besar: {tot_besar}")
            self.lbl_sum_kecil.config(text=f"Total Porsi Kecil: {tot_kecil}")
            self.lbl_sum_invalid.config(text=f"Baris Dilewati: {len(invalid_recs)}")

            self.logger.info(
                f"Validasi data berhasil: {len(valid_recs)} sekolah valid "
                f"(Total Besar={tot_besar}, Kecil={tot_kecil}). Dilewati={len(invalid_recs)} baris."
            )
            messagebox.showinfo("Sukses", f"Berhasil memuat {len(valid_recs)} data sekolah valid dari sheet '{sheet_name}'.")

        except Exception as e:
            messagebox.showerror("Kesalahan Validasi", f"Gagal memvalidasi data Excel:\n{str(e)}")
            self.logger.error(f"Gagal memvalidasi data Excel: {str(e)}")

    # -------------------------------------------------------------
    # LOGIKA GENERATOR KODE AUTOFILL CONSOLE
    # -------------------------------------------------------------
    def _copy_quick_script(self):
        """
        Menyalin kode JavaScript autofill ke clipboard.
        Pengguna cukup menempelkan (Ctrl+V) kode ini di Console DevTools browser yang sudah login.
        """
        if not self.valid_excel_records:
            messagebox.showwarning("Peringatan", "Silakan muat dan validasi data Excel terlebih dahulu.")
            return

        items = []
        for r in self.valid_excel_records:
            items.append({
                "nama": r['nama_sekolah'],
                "besar": r['porsi_besar'],
                "kecil": r['porsi_kecil']
            })

        items_json = json.dumps(items, ensure_ascii=False)

        # Skrip JavaScript yang sepenuhnya mandiri, aman, dan memicu event React
        js_code = (
            "(function autofillPopSipgn() {\n"
            "    console.log('%c=== MEMULAI AUTOFILL POP SIPGN ===', 'color: #2563eb; font-weight: bold; font-size: 14px;');\n"
            f"    const items = {items_json};\n\n"
            "    function norm(s) {\n"
            "        if (!s) return '';\n"
            "        let t = String(s).trim().toLowerCase();\n"
            "        t = t.replace(/[\\.\\,\\-\\-\\(\\)\\[\\]_:]/g, ' ');\n"
            "        t = t.replace(/[^\\w\\s]/g, '');\n"
            "        t = t.replace(/\\s+/g, ' ').trim();\n"
            "        const tokens = t.split(' ');\n"
            "        const out = [];\n"
            "        for (let i = 0; i < tokens.length; i++) {\n"
            "            const tok = tokens[i];\n"
            "            if (tok === 'sdn') out.push('sd', 'negeri');\n"
            "            else if (tok === 'smpn') out.push('smp', 'negeri');\n"
            "            else if (tok === 'sman') out.push('sma', 'negeri');\n"
            "            else if (tok === 'tkn') out.push('tk', 'negeri');\n"
            "            else if (tok === 'kec') out.push('kecamatan');\n"
            "            else if (tok === 'kel') out.push('kelurahan');\n"
            "            else out.push(tok);\n"
            "        }\n"
            "        return out.join(' ');\n"
            "    }\n\n"
            "    function setVal(el, val) {\n"
            "        if (!el) return false;\n"
            "        el.focus();\n"
            "        const strVal = String(val);\n"
            "        const lastVal = el.value;\n"
            "        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')?.set;\n"
            "        if (setter) { setter.call(el, strVal); } else { el.value = strVal; }\n"
            "        const tracker = el._valueTracker;\n"
            "        if (tracker) tracker.setValue(lastVal);\n"
            "        el.dispatchEvent(new Event('input', { bubbles: true }));\n"
            "        el.dispatchEvent(new Event('change', { bubbles: true }));\n"
            "        el.dispatchEvent(new Event('blur', { bubbles: true }));\n"
            "        return el.value === strVal;\n"
            "    }\n\n"
            "    const potentialCards = document.querySelectorAll('div, section, article');\n"
            "    const cards = [];\n"
            "    potentialCards.forEach(card => {\n"
            "        const inps = card.querySelectorAll('input');\n"
            "        if (inps.length === 2) {\n"
            "            const cardText = card.innerText || '';\n"
            "            if (cardText.includes('Porsi') || cardText.includes('Besar') || cardText.includes('Kecil')) {\n"
            "                const nested = card.querySelectorAll('div');\n"
            "                let isLeaf = true;\n"
            "                for (let nc of nested) {\n"
            "                    if (nc !== card && nc.querySelectorAll('input').length === 2) { isLeaf = false; break; }\n"
            "                }\n"
            "                if (isLeaf && !cards.includes(card)) cards.push(card);\n"
            "            }\n"
            "        }\n"
            "    });\n\n"
            "    console.log('Ditemukan ' + cards.length + ' kartu sekolah di halaman.');\n"
            "    let successCount = 0;\n"
            "    const notFound = [];\n\n"
            "    const cardData = cards.map(c => {\n"
            "        let name = '';\n"
            "        const headings = c.querySelectorAll('h1, h2, h3, h4, h5, h6, strong, b, .font-bold, .font-semibold, .font-medium');\n"
            "        for (let h of headings) {\n"
            "            const t = (h.innerText || '').trim();\n"
            "            if (t && !t.toLowerCase().includes('porsi') && !t.toLowerCase().includes('kpm')) { name = t; break; }\n"
            "        }\n"
            "        if (!name) {\n"
            "            const lines = (c.innerText || '').split('\\n').map(l => l.trim()).filter(l => l.length > 0);\n"
            "            for (let l of lines) {\n"
            "                if (!l.toLowerCase().includes('porsi') && !l.toLowerCase().includes('besar') && !l.toLowerCase().includes('kecil')) { name = l; break; }\n"
            "            }\n"
            "        }\n"
            "        const inps = Array.from(c.querySelectorAll('input'));\n"
            "        return { card: c, name: name, normName: norm(name), inputs: inps };\n"
            "    });\n\n"
            "    items.forEach(item => {\n"
            "        const normTarget = norm(item.nama);\n"
            "        const matched = cardData.find(c => c.normName === normTarget);\n"
            "        if (matched && matched.inputs.length >= 2) {\n"
            "            setVal(matched.inputs[0], item.besar);\n"
            "            setVal(matched.inputs[1], item.kecil);\n"
            "            matched.card.style.border = '2px solid #10b981';\n"
            "            matched.card.style.backgroundColor = '#f0fdf4';\n"
            "            console.log('✅ ' + item.nama + ' -> Besar: ' + item.besar + ', Kecil: ' + item.kecil);\n"
            "            successCount++;\n"
            "        } else {\n"
            "            notFound.push(item.nama);\n"
            "            console.warn('⚠️ Tidak ditemukan di halaman: ' + item.nama);\n"
            "        }\n"
            "    });\n\n"
            "    let alertMsg = '🎉 AUTOFILL POP SIPGN SELESAI!\\n\\n' +\n"
            "        '✅ Berhasil diisi: ' + successCount + ' sekolah\\n';\n"
            "    if (notFound.length > 0) {\n"
            "        alertMsg += '⚠️ Tidak ditemukan: ' + notFound.length + ' sekolah (cek console DevTools)\\n';\n"
            "    }\n"
            "    alertMsg += '\\nCatatan: Silakan tinjau kembali kartu di layar, lalu klik tombol \"Lanjut\" secara manual.';\n"
            "    alert(alertMsg);\n"
            "})();"
        )

        self.root.clipboard_clear()
        self.root.clipboard_append(js_code)
        self.root.update()

        self.logger.info(f"Kode autofill ({len(items)} sekolah) berhasil disalin ke clipboard.")
        messagebox.showinfo(
            "Kode Berhasil Disalin! 📋",
            f"Kode autofill untuk {len(items)} sekolah berhasil disalin ke Clipboard!\n\n"
            f"LANGKAH DI BROWSER ANDA:\n"
            f"1. Buka halaman POP SIPGN Langkah 2 yang sudah aktif.\n"
            f"2. Tekan F12 pada keyboard (atau Klik Kanan ➔ Inspect), lalu buka tab 'Console'.\n"
            f"3. Tekan Ctrl + V (Paste), lalu tekan Enter.\n\n"
            f"⚡ Seluruh kartu sekolah akan langsung terisi dan diberi tanda hijau dalam 1 detik!"
        )

    def _export_report(self):
        if not self.valid_excel_records:
            messagebox.showwarning("Peringatan", "Belum ada data untuk diekspor.")
            return

        report_filename = os.path.join(
            "reports", f"data_autofill_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        )
        try:
            export_data = []
            for r in self.valid_excel_records:
                export_data.append({
                    'Baris_Excel': r['row_index'],
                    'Nama_Sekolah': r['nama_sekolah'],
                    'Porsi_Besar': r['porsi_besar'],
                    'Porsi_Kecil': r['porsi_kecil'],
                    'Status': 'Valid'
                })

            rep_df = pd.DataFrame(export_data)
            rep_df.to_csv(report_filename, index=False, encoding='utf-8-sig')
            self.logger.info(f"Data berhasil diekspor ke: {report_filename}")
            messagebox.showinfo("Data Disimpan", f"Data berhasil disimpan ke file:\n{report_filename}")
        except Exception as e:
            messagebox.showerror("Gagal Ekspor", f"Gagal menyimpan file:\n{str(e)}")
            self.logger.error(f"Gagal ekspor: {str(e)}")


def main():
    root = tk.Tk()
    app = PopSipgnApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
