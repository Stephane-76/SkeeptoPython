"""Smoke test: workbook + formula. Requires a built engine and `pip install -e .` in a venv."""

from pathlib import Path

from skeepto import SpreadSheet

_EXCEL_DIR = Path(__file__).resolve().parent.parent / "excel"


def test_sum_formula():
    with SpreadSheet() as wSs:
        wSs.new_workbook("demo")
        assert wSs.value("A1", "10")
        assert wSs.value("A2", "20")
        assert wSs.value("B1", "=SUM(A1:A2)")
        wSs.recalculate_all()
        assert wSs.get_value("B1") == "30"


def test_sheet_and_formula():
    with SpreadSheet() as wSs:
        wSs.new_workbook("demo")
        assert wSs.add_sheet("Data", "")
        assert wSs.value("A1", "7", "Data")
        assert wSs.value("B1", "=A1*2", "Data")
        wSs.recalculate_all()
        assert wSs.get_value("B1", "Data") == "14"
        wFormula = wSs.get_formula("B1", "Data")
        assert "A1" in wFormula
        wList = wSs.sheets_list()
        assert "Data" in wList


def test_xlsx_roundtrip():
    _EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    wPath = str(_EXCEL_DIR / "demo.xlsx")
    with SpreadSheet() as wSs:
        wSs.new_workbook("demo")
        assert wSs.value("A1", "10")
        assert wSs.value("A2", "20")
        assert wSs.value("B1", "=SUM(A1:A2)")
        wSs.recalculate_all()
        assert wSs.save_xlsx(wPath)
    with SpreadSheet() as wSs:
        assert wSs.open_xlsx(wPath)
        wSs.recalculate_all()
        assert wSs.get_value("A1") == "10"
        assert wSs.get_value("A2") == "20"
        assert wSs.get_value("B1") == "30"
        assert "SUM" in wSs.get_formula("B1").upper()


def test_undo_actif_default_off():
    with SpreadSheet() as wSs:
        assert wSs.is_undo_actif() is False
        wSs.set_undo_actif(True)
        assert wSs.is_undo_actif() is True
        wSs.set_undo_actif(False)
        assert wSs.is_undo_actif() is False


if __name__ == "__main__":
    test_sum_formula()
    test_sheet_and_formula()
    test_xlsx_roundtrip()
    test_undo_actif_default_off()
    print("ok")
