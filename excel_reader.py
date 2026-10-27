"""
excel_reader.py
Modul untuk membaca dan menginspeksi file Excel (.xlsx) tanpa mengubah file sumber.
"""

from typing import List, Dict, Any, Optional, Tuple
import os
import openpyxl
import pandas as pd


class ExcelReader:
    def __init__(self, file_path: str):
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File tidak ditemukan: {file_path}")
        if not file_path.lower().endswith(('.xlsx', '.xlsm')):
            raise ValueError("File harus berekstensi .xlsx atau .xlsm")
        
        self.file_path = file_path
        self._sheet_names: Optional[List[str]] = None

    def get_sheet_names(self) -> List[str]:
        """
        Membaca daftar nama sheet yang tersedia di file Excel tanpa memuat seluruh data ke memori.
        """
        if self._sheet_names is None:
            # Gunakan openpyxl read_only=True untuk kecepatan dan keamanan memori
            wb = openpyxl.load_workbook(self.file_path, read_only=True, data_only=True)
            self._sheet_names = wb.sheetnames
            wb.close()
        return self._sheet_names

    def load_sheet_data(self, sheet_name: str) -> pd.DataFrame:
        """
        Membaca sheet tertentu ke dalam pandas DataFrame tanpa mengubah file asli.
        """
        sheet_names = self.get_sheet_names()
        if sheet_name not in sheet_names:
            raise ValueError(f"Sheet '{sheet_name}' tidak ditemukan. Sheet yang tersedia: {sheet_names}")

        # dtype=object untuk mempertahankan nilai 0 dan mencegah konversi tak diinginkan
        df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl')
        return df

    def get_columns(self, sheet_name: str) -> List[str]:
        """
        Mengambil nama kolom yang ada pada sheet yang dipilih.
        """
        df = self.load_sheet_data(sheet_name)
        return [str(c).strip() for c in df.columns]

    def get_preview(self, sheet_name: str, max_rows: int = 5) -> List[Dict[str, Any]]:
        """
        Mengambil pratinjau data (beberapa baris awal) untuk ditampilkan ke pengguna.
        """
        df = self.load_sheet_data(sheet_name)
        preview_df = df.head(max_rows)
        # Ganti NaN dengan representasi kosong untuk tampilan GUI
        records = []
        for _, row in preview_df.iterrows():
            item = {}
            for col in preview_df.columns:
                val = row[col]
                item[str(col)] = "" if pd.isna(val) else str(val)
            records.append(item)
        return records

    @staticmethod
    def auto_detect_columns(columns: List[str]) -> Tuple[Optional[str], Optional[str], Optional[str]]:
        """
        Mendeteksi kemungkinan kolom yang sesuai secara cerdas untuk kemudahan pengguna:
        - Kolom nama sekolah/KPM
        - Kolom porsi besar
        - Kolom porsi kecil
        """
        school_col = None
        besar_col = None
        kecil_col = None

        for col in columns:
            c_lower = col.lower()
            if not school_col and any(k in c_lower for k in ['sekolah', 'kpm', 'sasaran', 'penerima', 'nama']):
                school_col = col
            elif not besar_col and ('besar' in c_lower or 'dewasa' in c_lower):
                besar_col = col
            elif not kecil_col and ('kecil' in c_lower or 'anak' in c_lower or 'balita' in c_lower):
                kecil_col = col

        return school_col, besar_col, kecil_col
