"""
test_core.py
Skrip pengujian unit untuk modul pembaca Excel dan validator data porsi.
"""

from excel_reader import ExcelReader
from validator import validate_dataset, validate_portion

def test_portions():
    print("--- 1. Testing Portion Validator ---")
    assert validate_portion(0)[0] is True and validate_portion(0)[1] == 0, "Zero should be preserved!"
    assert validate_portion("0")[0] is True and validate_portion("0")[1] == 0, "String '0' should be preserved!"
    assert validate_portion(45)[0] is True and validate_portion(45)[1] == 45
    assert validate_portion(" 60 ")[0] is True and validate_portion(" 60 ")[1] == 60
    assert validate_portion(15.0)[0] is True and validate_portion(15.0)[1] == 15
    assert validate_portion(-5)[0] is False, "Negative should fail"
    assert validate_portion(12.5)[0] is False, "Float should fail"
    assert validate_portion("abc")[0] is False, "Text should fail"
    assert validate_portion(None)[0] is False, "None should fail"
    print("Portion tests passed successfully!")

def test_excel_reader():
    print("\n--- 2. Testing Excel Reader and Validator ---")
    reader = ExcelReader("contoh_data_sppg.xlsx")
    sheets = reader.get_sheet_names()
    print("Sheets in file:", sheets)
    assert "Oktober" in sheets

    cols = reader.get_columns("Oktober")
    print("Columns in Oktober:", cols)

    s_col, b_col, k_col = ExcelReader.auto_detect_columns(cols)
    print(f"Auto-detected: School='{s_col}', Besar='{b_col}', Kecil='{k_col}'")

    df = reader.load_sheet_data("Oktober")
    valid_rows, invalid_rows = validate_dataset(df, s_col, b_col, k_col)
    print(f"Validated rows: {len(valid_rows)} valid, {len(invalid_rows)} invalid")
    assert len(valid_rows) > 0, "Should have valid rows!"

if __name__ == "__main__":
    test_portions()
    test_excel_reader()
    print("\nALL CORE TESTS COMPLETED SUCCESSFULLY!")
