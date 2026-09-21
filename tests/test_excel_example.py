"""Richer example: formats, CSS, functions, and a bar chart, then write an .xlsx.

Requires a built engine and `pip install -e .` in a venv.

    python tests/test_excel_example.py
"""

from pathlib import Path

from skeepto import SpreadSheet

_EXCEL_DIR = Path(__file__).resolve().parent.parent / "excel"
_XLSX_PATH = _EXCEL_DIR / "exemple.xlsx"

_SHEET_FORMATS = "Formats"
_SHEET_FUNCTIONS = "Fonctions"

_HEADER_CSS = (
    'background-color:#1a365d;color:white;font:"Arial",serif 11pt;'
    "text-align:center;font-weight:bold;"
)
_TITLE_CSS = (
    'background-color:#1a365d;color:white;font:"Arial",serif 18pt;'
    "text-align:center;font-weight:bold;"
)
_LABEL_CSS = 'font:"Arial",serif 10pt;font-weight:bold;text-align:left;'
_NUMBER_CSS = 'font:"Arial",serif 10pt;text-align:right;'


def _build_formats_sheet(sSs):
    sSs.value("A1", "Skeepto — number formats and cell styles", _SHEET_FORMATS)
    sSs.merge("A1:D1", _SHEET_FORMATS)
    sSs.format("A1:D1", _TITLE_CSS, _SHEET_FORMATS)
    sSs.size_row(1, 1, 28, _SHEET_FORMATS)

    sSs.value("A2", "Description", _SHEET_FORMATS)
    sSs.value("B2", "Value", _SHEET_FORMATS)
    sSs.value("C2", "Format string", _SHEET_FORMATS)
    sSs.value("D2", "TEXT()", _SHEET_FORMATS)
    sSs.format("A2:D2", _HEADER_CSS, _SHEET_FORMATS)

    wRows = [
        (3, "Integer", "1234", "0"),
        (4, "Thousands", "1234567.8", "#,##0"),
        (5, "Two decimals", "1234.5", "#,##0.00"),
        (6, "Percent", "0.125", "0.00%"),
        (7, "Scientific", "1234567", "0.00E+00"),
        (8, "Currency", "1999.5", "$#,##0.00"),
        (9, "Date", "=DATE(2026,9,21)", "mm/dd/yyyy"),
        (10, "Date long", "=DATE(2026,9,21)", "yyyy-mm-dd"),
    ]
    for wRow, wLabel, wValue, wFmt in wRows:
        sSs.value(f"A{wRow}", wLabel, _SHEET_FORMATS)
        sSs.value(f"B{wRow}", wValue, _SHEET_FORMATS)
        sSs.value(f"C{wRow}", wFmt, _SHEET_FORMATS)
        # Engine TEXT() with US masks. Frozen after recalc: French Excel would
        # otherwise reinterpret "," / "dd" / "yyyy" as its own format tokens.
        sSs.value(f"D{wRow}", f'=TEXT(B{wRow},"{wFmt}")', _SHEET_FORMATS)
        sSs.apply_format_string(f"B{wRow}", wFmt, _SHEET_FORMATS)
        sSs.format(f"A{wRow}", _LABEL_CSS, _SHEET_FORMATS)
        sSs.format(f"B{wRow}", _NUMBER_CSS, _SHEET_FORMATS)

    sSs.value("A12", "CSS styles", _SHEET_FORMATS)
    sSs.format("A12", 'font:"Arial",serif 12pt;font-weight:bold;color:#1a365d;', _SHEET_FORMATS)

    # Named CSS colors (yellow, navy) are fragile on xlsx export — use hex.
    sSs.value("A13", "Bold navy", _SHEET_FORMATS)
    sSs.format("A13", 'font:"Arial",serif 11pt;font-weight:bold;color:#000080;', _SHEET_FORMATS)

    sSs.value("B13", "Yellow fill", _SHEET_FORMATS)
    sSs.format("B13", "background-color:#FFFF00;color:#000000;text-align:center;", _SHEET_FORMATS)

    sSs.value("C13", "Red on light blue", _SHEET_FORMATS)
    sSs.format("C13", "background-color:#ADD8E6;color:#FF0000;text-align:center;", _SHEET_FORMATS)

    sSs.value("A14", "Centered wrap", _SHEET_FORMATS)
    sSs.format(
        "A14:B14",
        "text-align:center;vertical-align:middle;text-wrap:wrap;background-color:#e2e8f0;",
        _SHEET_FORMATS,
    )
    sSs.merge("A14:B14", _SHEET_FORMATS)

    sSs.border("A2:D10", 1, "solid black 1px;", _SHEET_FORMATS)
    sSs.size_col(1, 1, 140, _SHEET_FORMATS)
    sSs.size_col(2, 2, 130, _SHEET_FORMATS)
    sSs.size_col(3, 3, 140, _SHEET_FORMATS)
    sSs.size_col(4, 4, 160, _SHEET_FORMATS)


def _build_functions_sheet(sSs):
    sSs.value("A1", "Skeepto — spreadsheet functions", _SHEET_FUNCTIONS)
    sSs.merge("A1:D1", _SHEET_FUNCTIONS)
    sSs.format("A1:D1", _TITLE_CSS, _SHEET_FUNCTIONS)
    sSs.size_row(1, 1, 28, _SHEET_FUNCTIONS)

    sSs.value("A2", "Product", _SHEET_FUNCTIONS)
    sSs.value("B2", "Qty", _SHEET_FUNCTIONS)
    sSs.value("C2", "Price", _SHEET_FUNCTIONS)
    sSs.value("D2", "Line total", _SHEET_FUNCTIONS)
    sSs.format("A2:D2", _HEADER_CSS, _SHEET_FUNCTIONS)

    wLines = [
        (3, "Apples", "10", "2"),
        (4, "Pears", "5", "4"),
        (5, "Plums", "8", "3"),
    ]
    for wRow, wName, wQty, wPrice in wLines:
        sSs.value(f"A{wRow}", wName, _SHEET_FUNCTIONS)
        sSs.value(f"B{wRow}", wQty, _SHEET_FUNCTIONS)
        sSs.value(f"C{wRow}", wPrice, _SHEET_FUNCTIONS)
        sSs.value(f"D{wRow}", f"=B{wRow}*C{wRow}", _SHEET_FUNCTIONS)
        sSs.apply_format_string(f"C{wRow}", "0.00", _SHEET_FUNCTIONS)
        sSs.apply_format_string(f"D{wRow}", "#,##0.00", _SHEET_FUNCTIONS)

    sSs.value("A6", "Totals / stats", _SHEET_FUNCTIONS)
    sSs.format("A6", _LABEL_CSS, _SHEET_FUNCTIONS)

    sSs.value("A7", "SUM", _SHEET_FUNCTIONS)
    sSs.value("B7", "=SUM(D3:D5)", _SHEET_FUNCTIONS)

    sSs.value("A8", "AVERAGE qty", _SHEET_FUNCTIONS)
    sSs.value("B8", "=AVERAGE(B3:B5)", _SHEET_FUNCTIONS)

    sSs.value("A9", "MIN / MAX price", _SHEET_FUNCTIONS)
    sSs.value("B9", "=MIN(C3:C5)", _SHEET_FUNCTIONS)
    sSs.value("C9", "=MAX(C3:C5)", _SHEET_FUNCTIONS)

    sSs.value("A10", "COUNT", _SHEET_FUNCTIONS)
    sSs.value("B10", "=COUNT(B3:B5)", _SHEET_FUNCTIONS)

    sSs.value("A11", "IF total>50", _SHEET_FUNCTIONS)
    sSs.value("B11", '=IF(B7>50,"OK","LOW")', _SHEET_FUNCTIONS)

    sSs.value("A12", "ROUND / ABS", _SHEET_FUNCTIONS)
    sSs.value("B12", "=ROUND(AVERAGE(C3:C5),2)", _SHEET_FUNCTIONS)
    sSs.value("C12", "=ABS(-3)", _SHEET_FUNCTIONS)

    sSs.value("A13", "POWER / SQRT", _SHEET_FUNCTIONS)
    sSs.value("B13", "=POWER(2,10)", _SHEET_FUNCTIONS)
    sSs.value("C13", "=SQRT(16)", _SHEET_FUNCTIONS)

    sSs.value("A14", "IFERROR", _SHEET_FUNCTIONS)
    sSs.value("B14", '=IFERROR(1/0,"err")', _SHEET_FUNCTIONS)

    sSs.value("A16", "Text", _SHEET_FUNCTIONS)
    sSs.format("A16", _LABEL_CSS, _SHEET_FUNCTIONS)
    sSs.value("B16", "Skeepto", _SHEET_FUNCTIONS)
    sSs.value("A17", "CONCATENATE", _SHEET_FUNCTIONS)
    sSs.value("B17", '=CONCATENATE(B16," ","Engine")', _SHEET_FUNCTIONS)
    sSs.value("A18", "LEFT / LEN", _SHEET_FUNCTIONS)
    sSs.value("B18", "=LEFT(B16,3)", _SHEET_FUNCTIONS)
    sSs.value("C18", "=LEN(B16)", _SHEET_FUNCTIONS)
    sSs.value("A19", "UPPER / LOWER", _SHEET_FUNCTIONS)
    sSs.value("B19", "=UPPER(B16)", _SHEET_FUNCTIONS)
    sSs.value("C19", "=LOWER(B16)", _SHEET_FUNCTIONS)

    sSs.value("A21", "Dates", _SHEET_FUNCTIONS)
    sSs.format("A21", _LABEL_CSS, _SHEET_FUNCTIONS)
    sSs.value("A22", "DATE", _SHEET_FUNCTIONS)
    sSs.value("B22", "=DATE(2026,9,21)", _SHEET_FUNCTIONS)
    sSs.apply_format_string("B22", "yyyy-mm-dd", _SHEET_FUNCTIONS)
    sSs.value("A23", "YEAR / MONTH / DAY", _SHEET_FUNCTIONS)
    sSs.value("B23", "=YEAR(B22)", _SHEET_FUNCTIONS)
    sSs.value("C23", "=MONTH(B22)", _SHEET_FUNCTIONS)
    sSs.value("D23", "=DAY(B22)", _SHEET_FUNCTIONS)

    sSs.border("A2:D5", 1, "solid black 1px;", _SHEET_FUNCTIONS)
    sSs.size_col(1, 1, 160, _SHEET_FUNCTIONS)
    sSs.size_col(2, 4, 110, _SHEET_FUNCTIONS)


def _ensure_chart_classes(sSs):
    sSs.register_class_attribute("SkCellClassLineChart", "LineChart", "Javascript")
    sSs.add_property("Title", "string", "Title", 0, "", "")
    sSs.add_property("chartData", "string", "Labels range", 1, "", "range")
    sSs.add_property("DataRange", "string", "Values range", 2, "", "range")
    sSs.add_property("seriesLabels", "string", "Series labels range", 3, "", "range")
    sSs.add_property("chartType", "string", "Chart type", 4, "bar", "enum:line,bar,area")
    sSs.add_property("barDirection", "string", "Bar direction", 5, "vertical", "enum:vertical,horizontal")


def _add_sales_chart(sSs):
    _ensure_chart_classes(sSs)
    assert sSs.insert_floating_object(
        "SalesChart",
        "SkCellClassLineChart",
        _SHEET_FUNCTIONS,
        _SHEET_FUNCTIONS,
        8.0,
        8.0,
        460.0,
        280.0,
        1.0,
        f"{_SHEET_FUNCTIONS}!F2",
    )
    sSs.floating_object_attribute("SalesChart", "Title", "Line totals", _SHEET_FUNCTIONS)
    sSs.floating_object_attribute(
        "SalesChart", "chartData", "=DATARANGE(A3:A5)", _SHEET_FUNCTIONS
    )
    sSs.floating_object_attribute(
        "SalesChart", "DataRange", "=DATARANGE(D3:D5)", _SHEET_FUNCTIONS
    )
    sSs.floating_object_attribute("SalesChart", "chartType", "bar", _SHEET_FUNCTIONS)


def _freeze_text_column(sSs):
    """Replace TEXT() formulas with their computed strings so Excel (FR) does not
    re-parse US format masks (mm/dd/yyyy → 09/dd/yyyy, #,##0.00 → 1234,50.00).
    """
    for wRow in range(3, 11):
        wShown = sSs.get_value(f"D{wRow}", _SHEET_FORMATS)
        sSs.value(f"D{wRow}", wShown, _SHEET_FORMATS)


def _assert_example(sSs):
    assert sSs.get_value("B7", _SHEET_FUNCTIONS) == "64"
    assert sSs.get_value("B10", _SHEET_FUNCTIONS) == "3"
    assert sSs.get_value("B11", _SHEET_FUNCTIONS) == "OK"
    assert sSs.get_value("C12", _SHEET_FUNCTIONS) == "3"
    assert sSs.get_value("B13", _SHEET_FUNCTIONS) == "1024"
    assert sSs.get_value("C13", _SHEET_FUNCTIONS) == "4"
    assert sSs.get_value("B14", _SHEET_FUNCTIONS) == "err"
    assert sSs.get_value("B18", _SHEET_FUNCTIONS) == "Ske"
    assert sSs.get_value("C18", _SHEET_FUNCTIONS) == "7"
    assert sSs.get_value("B19", _SHEET_FUNCTIONS) == "SKEEPTO"
    assert sSs.get_value("B23", _SHEET_FUNCTIONS) == "2026"
    assert sSs.get_value("C23", _SHEET_FUNCTIONS) == "9"
    assert sSs.get_value("D23", _SHEET_FUNCTIONS) == "21"
    assert "DATE" in sSs.get_formula("B9", _SHEET_FORMATS).upper()
    assert sSs.get_value("D3", _SHEET_FORMATS) == "1234"
    assert sSs.get_value("D5", _SHEET_FORMATS) == "1,234.50"
    assert sSs.get_value("D6", _SHEET_FORMATS) == "12.50%"
    assert sSs.get_value("D9", _SHEET_FORMATS) == "09/21/2026"
    assert sSs.get_value("D10", _SHEET_FORMATS) == "2026-09-21"
    wChartJson = sSs.json_floating_objects_for_sheet(_SHEET_FUNCTIONS, _SHEET_FUNCTIONS)
    assert "SalesChart" in wChartJson
    assert "SkCellClassLineChart" in wChartJson


def test_excel_example():
    _EXCEL_DIR.mkdir(parents=True, exist_ok=True)
    wPath = str(_XLSX_PATH)
    with SpreadSheet() as wSs:
        wSs.new_workbook("exemple")
        wSs.set_lang("us")
        assert wSs.add_sheet(_SHEET_FORMATS, "")
        assert wSs.add_sheet(_SHEET_FUNCTIONS, "")
        wSs.delete_sheet("Sheet1")
        _build_formats_sheet(wSs)
        _build_functions_sheet(wSs)
        _add_sales_chart(wSs)
        wSs.recalculate_all()
        _freeze_text_column(wSs)
        _assert_example(wSs)
        assert wSs.save_xlsx(wPath)
    assert _XLSX_PATH.is_file()
    print(wPath)


if __name__ == "__main__":
    test_excel_example()
    print("ok")
