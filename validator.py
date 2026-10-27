"""
validator.py
Modul untuk validasi data dari file Excel dan input formulir POP SIPGN.
"""

from typing import Any, Tuple, Optional, List, Dict
import pandas as pd


def validate_portion(value: Any, field_name: str = "Porsi") -> Tuple[bool, Optional[int], Optional[str]]:
    """
    Memvalidasi jumlah porsi sebagai bilangan bulat nonnegatif (>= 0).
    Mempertahankan nilai 0 sebagai 0, bukan data kosong.
    
    Returns:
        (is_valid, parsed_integer_or_None, error_message_or_None)
    """
    if value is None or pd.isna(value):
        return False, None, f"{field_name} kosong atau tidak terisi."
    
    # Jika bertipe string
    if isinstance(value, str):
        cleaned = value.strip()
        if not cleaned:
            return False, None, f"{field_name} berupa teks kosong."
        try:
            # Coba parse float dulu jika format "10.0"
            f_val = float(cleaned)
            if not f_val.is_integer():
                return False, None, f"{field_name} harus bilangan bulat, didapat desimal: {cleaned}"
            i_val = int(f_val)
        except ValueError:
            return False, None, f"{field_name} bukan angka valid: '{cleaned}'"
    elif isinstance(value, (int,)):
        i_val = value
    elif isinstance(value, (float,)):
        if not value.is_integer():
            return False, None, f"{field_name} harus bilangan bulat, didapat desimal: {value}"
        i_val = int(value)
    else:
        try:
            i_val = int(value)
        except (ValueError, TypeError):
            return False, None, f"{field_name} tipe data tidak valid: {type(value).__name__}"

    if i_val < 0:
        return False, None, f"{field_name} tidak boleh bernilai negatif: {i_val}"

    return True, i_val, None


def validate_dataset(
    df: pd.DataFrame,
    school_col: str,
    porsi_besar_col: str,
    porsi_kecil_col: str
) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
    """
    Memvalidasi seluruh baris data dari DataFrame Excel.
    
    Returns:
        (valid_records, invalid_records)
        Setiap item valid: {
            'row_index': int,
            'nama_sekolah': str,
            'porsi_besar': int,
            'porsi_kecil': int
        }
        Setiap item invalid: {
            'row_index': int,
            'nama_sekolah_raw': Any,
            'porsi_besar_raw': Any,
            'porsi_kecil_raw': Any,
            'alasan': str
        }
    """
    valid_records: List[Dict[str, Any]] = []
    invalid_records: List[Dict[str, Any]] = []

    for idx, row in df.iterrows():
        excel_row_num = idx + 2  # Asumsi header baris 1 (1-indexed)

        raw_school = row.get(school_col)
        raw_besar = row.get(porsi_besar_col)
        raw_kecil = row.get(porsi_kecil_col)

        # Cek jika seluruh baris kosong
        if pd.isna(raw_school) and pd.isna(raw_besar) and pd.isna(raw_kecil):
            continue  # Abaikan baris kosong dengan aman

        # Cek nama sekolah
        if pd.isna(raw_school) or str(raw_school).strip() == "":
            invalid_records.append({
                'row_index': excel_row_num,
                'nama_sekolah_raw': raw_school,
                'porsi_besar_raw': raw_besar,
                'porsi_kecil_raw': raw_kecil,
                'alasan': "Nama sekolah/KPM kosong."
            })
            continue

        school_name = str(raw_school).strip()

        # Validasi porsi besar
        ok_besar, val_besar, err_besar = validate_portion(raw_besar, "Porsi Besar")
        if not ok_besar:
            invalid_records.append({
                'row_index': excel_row_num,
                'nama_sekolah_raw': school_name,
                'porsi_besar_raw': raw_besar,
                'porsi_kecil_raw': raw_kecil,
                'alasan': err_besar
            })
            continue

        # Validasi porsi kecil
        ok_kecil, val_kecil, err_kecil = validate_portion(raw_kecil, "Porsi Kecil")
        if not ok_kecil:
            invalid_records.append({
                'row_index': excel_row_num,
                'nama_sekolah_raw': school_name,
                'porsi_besar_raw': raw_besar,
                'porsi_kecil_raw': raw_kecil,
                'alasan': err_kecil
            })
            continue

        valid_records.append({
            'row_index': excel_row_num,
            'nama_sekolah': school_name,
            'porsi_besar': val_besar,
            'porsi_kecil': val_kecil
        })

    return valid_records, invalid_records
