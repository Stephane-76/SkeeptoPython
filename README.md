# Skeepto Python

A **real spreadsheet engine** — a competitor to Excel and Google Sheets.
It **reads and writes `.xlsx`**.

Native Python bindings for [skeepto-engine](https://github.com/Stephane-76/SkeeptoEngine).
This repo is a **client** of the C++ core (same role as the WASM `UISpreadSheet`
module): formulas, recalc, and the grid run in the engine, not in Python.

Excel `.xlsx` import/export is the C++ `SkExcel` pipeline (`open_xlsx` /
`save_xlsx`). Do not layer openpyxl or pandas as a second engine.

## Layout

Keep the checkouts as siblings:

```text
Projects/
├── skeepto-engine/     # C++ core (you compile this)
└── skeepto-python/     # this package (links the .a / .lib)
```

Override the engine path with `SK_ENGINE_DIR` if they are not siblings.

## Do not pollute the system Python

Always use a **virtualenv**. Never `sudo pip` and never install into Homebrew
or the macOS system interpreter.

```bash
# 1. Build the engine (once, in skeepto-engine)
cd ../skeepto-engine
cmake -B build-unix
cmake --build build-unix --parallel

# 2. Isolated venv (in skeepto-python)
cd ../skeepto-python
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate

pip install -e ".[test]"
pytest
```

`pip install -e .` compiles `_core` into `.venv/` and fetches **pybind11 into
the build directory**. Nothing is written into skeepto-engine or the system
Python.

## Usage

```python
from skeepto import SpreadSheet

with SpreadSheet() as ss:
    ss.new_workbook("demo")
    ss.value("A1", "10")
    ss.value("A2", "20")
    ss.value("B1", "=SUM(A1:A2)")
    ss.recalculate_all()
    print(ss.get_value("B1"))   # 30

    ss.save_xlsx("demo.xlsx")

with SpreadSheet() as ss:
    ss.open_xlsx("demo.xlsx")
    ss.recalculate_all()
    print(ss.get_value("B1"))
```

`tApplication` is a process-wide singleton in the engine: one Python process,
one engine. Several workbooks on one `SpreadSheet` are fine. Prefer
`with SpreadSheet() as ss:` so the UI object is destroyed before process exit.

A **DEBUG** engine (`_DEBUGLeak`) can print leftover `tClass` counts when
`tApplication` is destroyed. The Python module turns that **exit dump off**
so pytest stays readable. To see it:

```bash
SK_DEBUG_LEAK=1 pytest
```

Or call `skeepto.debug_memory()` yourself. A **RELEASE** engine has no
counter and no dump.

## What this wraps

`SpreadSheet` is `tUISpreadSheet` — the same class bound to JavaScript as
`UISpreadSheet` in `SkReactSpreadSheet`. Method names are snake_case
(`new_workbook`, `get_value`, `json_view`, …) instead of PascalCase.

Internals (`tCell*`, `tSheet*`, `tClass`) are not exposed. `.xlsx` goes through
`SkExcel` (`open_xlsx` / `save_xlsx`).

API reference (default sheet, undo, Excel, …): [doc.md](doc.md).

## License

MIT.
