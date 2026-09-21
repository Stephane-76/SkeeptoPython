# Skeepto Python API

`SpreadSheet` is the native client of **skeepto-engine** (`tUISpreadSheet`),
the same surface as the WASM `UISpreadSheet` module. Method names are
`snake_case`. C++ internals (`tCell*`, `tSheet*`, `tClass`) are not exposed.

```python
from skeepto import SpreadSheet

with SpreadSheet() as ss:
    ss.new_workbook("demo")
    ss.value("A1", "10")
    ss.value("A2", "20")
    ss.value("B1", "=SUM(A1:A2)")
    ss.recalculate_all()
    print(ss.get_value("B1"))  # 30
```

Prefer `with SpreadSheet() as ss:`: one Python process is one engine
(`tApplication` singleton). Destroy the object before process exit.

---

## Default sheet (`sheet=""`)

Almost every cell-level method takes an optional `sheet=""` argument.

**If `sheet` is empty (`""`, the default), the call uses the active
sheet** (the workbook’s current sheet). It does not look up a sheet
named empty.

In the engine, `SetSheet("")` is a no-op and the operation continues on
`ActiveSheet()`.

```python
ss.value("A1", "10")                 # active sheet
ss.value("A1", "10", "")             # same: active sheet
ss.value("A1", "10", "Data")         # named sheet Data
ss.set_active_sheet("Data")
ss.get_value("A1")                   # reads Data, now active
```

`new_workbook()` typically creates `Sheet1` and makes it active.
`get_active_sheet()` / `set_active_sheet(name)` read or change that
current sheet. `add_sheet(name, left="")` adds a sheet (to the left of
`left` when `left` is set).

If `sheet` names a sheet that **does not exist**, the call fails (`False`
or an empty string, depending on the method).

---

## Lifetime

| Method | Role |
|--------|------|
| `SpreadSheet()` | Creates the native UI. Undo is **off** by default. |
| `close()` | Destroys the object (same as `__exit__`). |
| `set_lang(lang)` | Engine locale (`"us"`, `"fr"`, …). Excel / `.sker` formulas stay US-style (comma between arguments). |

---

## Workbook

| Method | Role |
|--------|------|
| `new_workbook(uri)` | New workbook (free-form URI, e.g. `"demo"`). |
| `add_workbook(uri)` | Adds a workbook. |
| `set_active_workbook(uri)` / `get_active_workbook()` | Current workbook. |
| `delete_workbook(uri)` / `rename_workbook(uri, uri_to)` | Delete / rename. |
| `write_json(uri)` / `read_json(json)` | Engine JSON (not xlsx). |
| `json_workbook(uri)` / `json_workbooks()` | Metadata. |

### Excel (C++ `SkExcel` pipeline, not openpyxl)

| Method | Role |
|--------|------|
| `open_xlsx(path, materialize_indirect=True, recalculate=False)` | Import `.xlsx`. |
| `save_xlsx(path)` | Export `.xlsx`. |

After `open_xlsx`, call `recalculate_all()` if you need computed values.

---

## Sheets

| Method | Role |
|--------|------|
| `add_sheet(name, left="")` | New sheet. |
| `set_active_sheet(name)` / `get_active_sheet()` | Default sheet for `sheet=""`. |
| `rename_sheet(name, new_name)` | Rename. |
| `swap_sheet(name1, name2, insert_after=False)` | Tab order. |
| `delete_sheet(name)` | Delete. |
| `sheets_list()` | List (JSON string). |

---

## Cells and formulas

Refs look like `"A1"`, `"B2:D10"`, `"A1;C3"` (multi-selection). `value`
accepts a constant or a formula (`"=SUM(A1:A2)"`).

| Method | Role |
|--------|------|
| `value(ref, value, sheet="")` | Write a value or a formula. |
| `value_int` / `value_double` / `value_string` | Typed writes. |
| `get_value(ref, sheet="")` | Displayed value (after recalc). |
| `get_input_value(ref, sheet="")` | Entered value. |
| `get_formula(ref, sheet="", user=False)` | Formula. |
| `get_calculable_scalar(ref, sheet="")` | Calculable scalar. |
| `ensure_cell(ref, sheet="")` | Ensure the cell exists. |
| `fill_series(source_ref, dest_ref, sheet="")` | Fill series. |
| `raz(ref, keep_format=False, sheet="")` | Clear content. |
| `raz_format(ref, sheet="")` | Clear format. |

Formula compile errors:

`compile_error()`, `compile_error_line()`, `compile_error_column()`,
`compile_error_with_detail()`.

Recalc:

| Method | Role |
|--------|------|
| `recalculate_all()` | Blocking full-workbook recalc. |
| `begin_recalculate_all_cooperative()` + `step_recalculate_all_cooperative(max_ms)` | Time-sliced recalc. |
| `recalculate_all_cooperative_progress()` / `is_recalculate_all_cooperative_active()` | Progress. |
| `set_cooperative_calculate_enabled(enabled)` | Enable cooperative mode. |

---

## Undo / redo

By default **`is_undo_actif()` is `False`**: commands still apply but are
**not** pushed on the stack (Python batch / import). To record history:

```python
ss.set_undo_actif(True)
ss.value("A1", "1")
ss.value("A1", "2")
ss.undo()   # A1 is 1 again
ss.redo()
```

`set_extra_undo(json)` / `get_extra_undo()`: UI payload attached to the
stack (selection, and so on).

---

## Format

Cell format is engine **CSS** (`background-color:`, `font:`, `color:`,
…). Named CSS colors (`yellow`) and hex (`#FFFF00`) are accepted.

| Method | Role |
|--------|------|
| `format(ref, value, sheet="")` | Apply CSS. |
| `get_format(ref, sheet="")` | Current CSS. |
| `apply_format_string(ref, value, sheet="")` | Excel number mask (`"#,##0.00"`). |
| `border(ref, border, value, sheet="")` | Borders (`border` = side code, `value` = CSS). |
| `precision(ref, inc, sheet="")` | Decimal places +/-. |
| `merge(ref, sheet="")` | Merge. |
| `return_merged(row, col, sheet="")` / `return_range_merged(ref, sheet="")` | Merge info. |
| `conditional_format(...)` / `delete_conditional_format(type, ref, sheet="")` | Conditional format. |
| `json_conditional_format(sheet="")` / `json_format_string()` | JSON dump. |

---

## Rows, columns, outline

Indices are **1-based** (row 1 is the first row, like Excel).

| Method | Role |
|--------|------|
| `insert_row(begin, end, sheet="")` / `delete_row(begin, end, sheet="")` | Insert / delete rows. |
| `insert_col(begin, end, sheet="")` / `delete_col(begin, end, sheet="")` | Insert / delete columns. |
| `insert_row_by_rect(ref, sheet="")` / `delete_row_by_rect` / `*_col_by_rect` | Via an A1 rectangle. |
| `size_row(begin, end, size, sheet="")` / `size_col(...)` | Height / width. |
| `get_size_row(index, sheet="")` / `get_size_col` | Read size. |
| `open_close_tree_row(row, sheet="")` / `open_close_tree_col` | Outline. |
| `change_tree_row(right, row, size, sheet="")` / `change_tree_col` | Outline level. |
| `split_view(cde, position, sheet="")` | Split / freeze. |
| `split_freeze_col(sheet="")` / `split_freeze_row(sheet="")` | Freeze position. |

---

## Copy / paste / move

| Method | Role |
|--------|------|
| `copy(ref, sheet="")` / `cut(ref, sheet="")` | Engine clipboard. |
| `paste(ref, sheet="")` | Paste. |
| `move(source_ref, dest_ref, sheet="")` | Move. |

---

## Names, tables, find

| Method | Role |
|--------|------|
| `insert_named_range(name, ref, sheet="")` | Named range. |
| `update_named_range(old_name, new_name, new_ref, sheet="")` | Update. |
| `delete_named_range(name, sheet="")` | Delete. |
| `insert_formula_named(name, formula, sheet="")` / `delete_formula_named` | Named formula. |
| `json_range_named()` / `json_formula_named()` / `json_range_data()` | JSON dump. |
| `undo_add_range_data(name, ref, json_data="", sheet="")` | Table / RangeData. |
| `undo_apply_range_data(name, json_data, sheet="")` | Table filter / sort. |
| `json_find_cell(search, match_case=False, match_entire_cell=False, sheet="")` | Find. |
| `json_find_unique_value(ref, sheet="")` | Unique values. |

`is_rename_allowed()` is `False` in multi-user mode (name rebase in
formulas). `set_multi_user_active(active)` forwards that state.

---

## Floating objects (charts, …)

| Method | Role |
|--------|------|
| `insert_floating_object(name, class_name, target_sheet_name, sheet="", ...)` | Insert. |
| `floating_object_layout(...)` / `floating_object_attribute(...)` | Position / props. |
| `floating_object_bring_to_front(name, sheet="")` | Z-order. |
| `delete_floating_object(name, sheet="")` | Delete. |
| `json_floating_objects(sheet="")` | JSON list. |

---

## JSON view (grid)

Useful for a custom renderer. `sheet=""` is the active sheet.

| Method | Role |
|--------|------|
| `json_view(row, col, view_height, view_width, diff_y, diff_x, sheet="", css=False)` | Viewport. |
| `json_bottom_right(sheet="")` / `json_pixel_bottom_right(sheet="")` | Bottom-right corner. |
| `sum_pixel_height(row_start, row_end, sheet="")` / `sum_pixel_width` | Pixel sums. |
| `max_row()` / `max_col()` | Bounds of the active workbook. |
| `json_print_parameters(sheet="")` / `set_json_print_parameters(json, sheet="")` | Print. |

---

## Utilities

| Method | Role |
|--------|------|
| `base10_to_alpha(value)` / `alpha_to_base10(value)` | `1 → "A"`, `"A" → 1`. |
| `parse_cell(ref)` / `parse_range(ref)` | Parse A1. |
| `cell_ref(ref_anchor, ref, sheet="")` | Relative reference. |
| `qualify_refs_for_sheet(sheet, text)` | Qualify refs for a sheet. |
| `call(function_name, json_params)` | JSON dispatch (same map as WASM `_Call`). |

Cell classes: `register_class_attribute`, `add_property`,
`cell_class(ref, class_name)`, `value_attribute`, `json_cell_class`.

Synthetic load (tests): `pressure(dynamic_row, dynamic_col, sheet="")`.

---

## Multi-sheet example

```python
from skeepto import SpreadSheet

with SpreadSheet() as ss:
    ss.new_workbook("demo")
    ss.add_sheet("Data")
    ss.set_active_sheet("Data")

    ss.value("A1", "7")                    # Data (active)
    ss.value("B1", "=A1*2")                # Data
    ss.value("A1", "hello", "Sheet1")      # other sheet, explicit

    ss.recalculate_all()
    assert ss.get_value("B1") == "14"            # Data
    assert ss.get_value("A1", "Sheet1") == "hello"
```
