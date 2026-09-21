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

**Unix / macOS**

```bash
# 1. Engine (skeepto-engine)
cd ../skeepto-engine
cmake -B build-unix
cmake --build build-unix --parallel
# Release (no leak counters):
#   cmake -B build-unix -DCMAKE_BUILD_TYPE=Release
#   cmake --build build-unix --parallel

# 2. Python (skeepto-python)
cd ../skeepto-python
python3 -m venv .venv
source .venv/bin/activate
pip install -e ".[test]"
pytest
```

**Windows** (Visual Studio + x64, Release — matches `windows/lib/Release`)

```bat
cd ..\skeepto-engine
cmake -A x64 -B build-windows -DSK_PLATFORM=windows -DSK_CONFIGS=Release
cmake --build build-windows --config Release --parallel

cd ..\skeepto-python
python -m venv .venv
.venv\Scripts\activate
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

API reference (default sheet, undo, Excel, …): [doc.md](doc.md).

## License

[MIT](LICENSE) © 2026 Stéphane Allez.
