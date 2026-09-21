//=============================================================================
// SkCore.cpp — pybind11 wrapper around tUISpreadSheet (same API as WASM).
//=============================================================================
#include <pybind11/pybind11.h>
#include <cstdlib>
#include <filesystem>
#include <stdexcept>

#include <SkClass.hpp>
#include <SkExcel2SpreadSheet.hpp>
#include <SkSpreadSheet.hpp>
#include <SkSpreadSheet2Excel.hpp>
#include <SkUISpreadSheetApi.hpp>

namespace py = pybind11;

using SkRoot::DebugMemory;
using SkRoot::ReportLeakAtExit;
using SkRoot::tBool;
using SkRoot::tByte;
using SkRoot::tDouble;
using SkRoot::tIndex;
using SkRoot::tInt;
using SkRoot::tSize;
using SkRoot::tString;
using SkSpreadSheet::tSpreadSheetContainer;
using SkSpreadSheet::tUISpreadSheet;

namespace {

class tPySpreadSheet {
private:
    tUISpreadSheet* m_UI;

    tUISpreadSheet& Ui() {
        if (m_UI == nullptr) {
            throw std::runtime_error("SpreadSheet is closed");
        }
        return *m_UI;
    }

public:
    tPySpreadSheet() : m_UI(new tUISpreadSheet()) {
        // WASM sets Client(true) to PostMessage undo to the JS host. Python is
        // in-process: that path can dereference ActiveWorkBook() and crash.
        m_UI->Client(false);
        // tApi default is off; keep it explicit so a later UI ctor change cannot
        // turn Python undo recording back on.
        m_UI->IsUndoActif(false);
        // Drop the JS OnCellChange hook (dangling this after Close / DoneSpreadSheet).
        tSpreadSheetContainer::SetOnCellChange(nullptr);
    }
    ~tPySpreadSheet() { Close(); }

    tPySpreadSheet(const tPySpreadSheet&) = delete;
    tPySpreadSheet& operator=(const tPySpreadSheet&) = delete;

    void Close() {
        if (m_UI == nullptr) {
            return;
        }
        tSpreadSheetContainer::SetOnCellChange(nullptr);
        delete m_UI;
        m_UI = nullptr;
    }

    tString GetActiveWorkBook() { return Ui()._GetActiveWorkBook(); }
    tBool SetActiveWorkBook(tString sUri) { return Ui()._SetActiveWorkBook(sUri); }
    tBool NewWorkBook(tString sUri) { return Ui()._NewWorkBook(sUri); }
    tBool AddWorkBook(tString sUri) { return Ui()._AddWorkBook(sUri); }
    tBool DeleteWorkBook(tString sUri) { return Ui()._DeleteWorkBook(sUri); }
    tBool RenameWorkBook(tString sUri, tString sUriTo) {
        return Ui()._RenameWorkBook(sUri, sUriTo);
    }
    tString JsonWorkBook(tString sUri) { return Ui()._JsonWorkBook(sUri); }
    tString JsonWorkBooks() { return Ui()._JsonWorkBooks(); }
    tString WriteJson(tString sUri) { return Ui()._WriteJson(sUri); }
    tBool ReadJson(tString sJson) { return Ui()._ReadJson(sJson); }

    tBool RecalculateAll() { return Ui()._RecalculateAll(); }
    void BeginRecalculateAllCooperative() { Ui()._BeginRecalculateAllCooperative(); }
    tBool StepRecalculateAllCooperative(tInt sMaxMs) {
        return Ui()._StepRecalculateAllCooperative(sMaxMs);
    }
    tInt RecalculateAllCooperativeProgress() {
        return Ui()._RecalculateAllCooperativeProgress();
    }
    tBool IsRecalculateAllCooperativeActive() {
        return Ui()._IsRecalculateAllCooperativeActive();
    }
    void SetCooperativeCalculateEnabled(tBool sEnabled) {
        Ui()._SetCooperativeCalculateEnabled(sEnabled);
    }
    tBool CooperativeCalculateEnabled() { return Ui()._CooperativeCalculateEnabled(); }

    void Undo() { Ui()._Undo(); }
    void Redo() { Ui()._Redo(); }
    tBool IsUndoActif() { return Ui()._IsUndoActif(); }
    void SetIsUndoActif(tBool sIsUndoActif) { Ui()._SetIsUndoActif(sIsUndoActif); }
    void SetUserInterface(tString sJson) { Ui()._SetUserInterface(sJson); }
    tString GetUserInterface() { return Ui()._GetUserInterface(); }
    void SetExtraUndo(tString sJson) { Ui()._SetExtraUndo(sJson); }
    tString GetExtraUndo() { return Ui()._GetExtraUndo(); }

    tBool SetActiveSheet(tString sName) { return Ui()._SetActiveSheet(sName); }
    tString GetActiveSheet() { return Ui()._GetActiveSheet(); }
    tBool AddSheet(tString sName, tString sLeft) { return Ui()._AddSheet(sName, sLeft); }
    tBool RenameSheet(tString sName, tString sNewName) {
        return Ui()._RenameSheet(sName, sNewName);
    }
    tBool SwapSheet(tString sName1, tString sName2, tBool sInsertAfter) {
        return Ui()._SwapSheet(sName1, sName2, sInsertAfter);
    }
    tBool DeleteSheet(tString sName) { return Ui()._DeleteSheet(sName); }
    tString SheetsList() { return Ui()._SheetsList(); }

    tBool Value(tString sRef, tString sValue, tString sSheet) {
        return Ui()._Value(sRef, sValue, sSheet);
    }
    tBool FillSeries(tString sSourceRef, tString sDestRef, tString sSheet) {
        return Ui()._FillSeries(sSourceRef, sDestRef, sSheet);
    }
    tBool ValueAttribute(tString sRef, tString sName, tString sValue, tString sSheet) {
        return Ui()._ValueAttribute(sRef, sName, sValue, sSheet);
    }
    tBool CellClassAttributes(tString sRef, tString sJson, tString sSheet) {
        return Ui()._CellClassAttributes(sRef, sJson, sSheet);
    }
    tBool ValueClassCalculable(tString sRef, tString sValue, tString sSheet) {
        return Ui()._ValueClassCalculable(sRef, sValue, sSheet);
    }
    tBool ValueString(tString sRef, tString sValue, tString sSheet) {
        return Ui()._ValueString(sRef, sValue, sSheet);
    }
    tBool ValueInt(tString sRef, tInt sValue, tString sSheet) {
        return Ui()._ValueInt(sRef, sValue, sSheet);
    }
    tBool ValueDouble(tString sRef, tDouble sValue, tString sSheet) {
        return Ui()._ValueDouble(sRef, sValue, sSheet);
    }
    tString GetValue(tString sRef, tString sSheet) { return Ui()._GetValue(sRef, sSheet); }
    tString GetCalculableScalar(tString sRef, tString sSheet) {
        return Ui()._GetCalculableScalar(sRef, sSheet);
    }
    tString GetValueAttribute(tString sRef, tString sName, tString sSheet) {
        return Ui()._GetValueAttribute(sRef, sName, sSheet);
    }
    tString GetFormulaAttribute(tString sRef, tString sName, tString sSheet) {
        return Ui()._GetFormulaAttribute(sRef, sName, sSheet);
    }
    tString GetInputValue(tString sRef, tString sSheet) {
        return Ui()._GetInputValue(sRef, sSheet);
    }
    tString GetFormula(tString sRef, tString sSheet, tBool sUser) {
        return Ui()._GetFormula(sRef, sSheet, sUser);
    }
    tString CellRef(tString sRefAnchor, tString sRef, tString sSheet) {
        return Ui()._CellRef(sRefAnchor, sRef, sSheet);
    }
    tString CompileError() { return Ui()._Error(); }
    tInt CompileErrorLine() { return Ui()._ErrorLine(); }
    tInt CompileErrorColumn() { return Ui()._ErrorColumn(); }
    tString CompileErrorWithDetail() { return Ui()._ErrorWithDetail(); }

    tBool Raz(tString sRef, tBool sKeepFormat, tString sSheet) {
        return Ui()._Raz(sRef, sKeepFormat, sSheet);
    }
    tBool RazFormat(tString sRef, tString sSheet) { return Ui()._RazFormat(sRef, sSheet); }

    tBool SizeRow(tInt sBegin, tInt sEnd, tDouble sSize, tString sSheet) {
        return Ui()._SizeRow(sBegin, sEnd, sSize, sSheet);
    }
    tBool SizeCol(tInt sBegin, tInt sEnd, tDouble sSize, tString sSheet) {
        return Ui()._SizeCol(sBegin, sEnd, sSize, sSheet);
    }
    tDouble GetSizeRow(tInt sIndex, tString sSheet) { return Ui()._GetSizeRow(sIndex, sSheet); }
    tDouble GetSizeCol(tInt sIndex, tString sSheet) { return Ui()._GetSizeCol(sIndex, sSheet); }

    tBool InsertRow(tInt sBegin, tInt sEnd, tString sSheet) {
        return Ui()._InsertRow(sBegin, sEnd, sSheet);
    }
    tBool InsertRowByRect(tString sRef, tString sSheet) {
        return Ui()._InsertRowByRect(sRef, sSheet);
    }
    tBool InsertRowByRectWithLabel(tString sRef, tString sLabelRef, tString sLabelValue,
                                   tString sSheet) {
        return Ui()._InsertRowByRectWithLabel(sRef, sLabelRef, sLabelValue, sSheet);
    }
    tBool DeleteRow(tInt sBegin, tInt sEnd, tString sSheet) {
        return Ui()._DeleteRow(sBegin, sEnd, sSheet);
    }
    tBool DeleteRowByRect(tString sRef, tString sSheet) {
        return Ui()._DeleteRowByRect(sRef, sSheet);
    }
    tBool InsertCol(tInt sBegin, tInt sEnd, tString sSheet) {
        return Ui()._InsertCol(sBegin, sEnd, sSheet);
    }
    tBool InsertColByRect(tString sRef, tString sSheet) {
        return Ui()._InsertColByRect(sRef, sSheet);
    }
    tBool DeleteCol(tInt sBegin, tInt sEnd, tString sSheet) {
        return Ui()._DeleteCol(sBegin, sEnd, sSheet);
    }
    tBool DeleteColByRect(tString sRef, tString sSheet) {
        return Ui()._DeleteColByRect(sRef, sSheet);
    }

    tBool Copy(tString sRef, tString sSheet) { return Ui()._Copy(sRef, sSheet); }
    tBool Cut(tString sRef, tString sSheet) { return Ui()._Cut(sRef, sSheet); }
    tBool Paste(tString sRef, tString sSheet) { return Ui()._Paste(sRef, sSheet); }
    tBool Move(tString sSourceRef, tString sDestRef, tString sSheet) {
        return Ui()._Move(sSourceRef, sDestRef, sSheet);
    }

    tBool Format(tString sRef, tString sValue, tString sSheet) {
        return Ui()._Format(sRef, sValue, sSheet);
    }
    tBool Precision(tString sRef, tBool sInc, tString sSheet) {
        return Ui()._Precision(sRef, sInc, sSheet);
    }
    tBool ConditionalFormat(tString sType, tString sRef, tString sP1, tString sP2, tString sP3,
                            tString sP4, tString sP5, tString sP6, tString sP7, tString sP8,
                            tString sP9, tString sP10, tString sSheet) {
        return Ui()._ConditionalFormat(sType, sRef, sP1, sP2, sP3, sP4, sP5, sP6, sP7, sP8, sP9,
                                      sP10, sSheet);
    }
    tBool DeleteConditionalFormat(tString sType, tString sRef, tString sSheet) {
        return Ui()._DeleteConditionalFormat(sType, sRef, sSheet);
    }
    tBool Border(tString sRef, tInt sBorder, tString sValue, tString sSheet) {
        return Ui()._Border(sRef, sBorder, sValue, sSheet);
    }
    tString GetFormat(tString sRef, tString sSheet) { return Ui()._GetFormat(sRef, sSheet); }
    tString DebugFormat() { return Ui()._DEBUGSKFormat(); }
    tInt FormatCount() { return Ui()._FormatCount(); }
    tBool ApplyFormatString(tString sRef, tString sValue, tString sSheet) {
        // CSS parser: format-string:"#,##0.00" — unquoted #, starts a color token.
        tString wValue = sValue;
        const tBool wQuoted =
            wValue.size() >= 2 &&
            ((wValue.front() == '"' && wValue.back() == '"') ||
             (wValue.front() == '\'' && wValue.back() == '\''));
        if (!wQuoted) {
            wValue = "\"" + sValue + "\"";
        }
        return Ui()._ApplyFormatString(sRef, wValue, sSheet);
    }
    tString DefaultFormatString(tString sFormatString) {
        return Ui()._DefaultFormatString(sFormatString);
    }
    tString FormatValueWithCellFormat(tString sRef, tString sValue, tString sSheet) {
        return Ui()._FormatValueWithCellFormat(sRef, sValue, sSheet);
    }
    tString JsonFormatString() { return Ui()._JsonFormatString(); }
    tString JsonConditionalFormat(tString sSheet) { return Ui()._JsonConditionalFormat(sSheet); }

    tBool Merge(tString sRef, tString sSheet) { return Ui()._Merge(sRef, sSheet); }
    tString ReturnMerged(tInt sRow, tInt sCol, tString sSheet) {
        return Ui()._ReturnMerged(sRow, sCol, sSheet);
    }
    tString ReturnRangeMerged(tString sRef, tString sSheet) {
        return Ui()._ReturnRangeMerged(sRef, sSheet);
    }
    tString ReturnRangeMergedFusion(tString sRef, tString sSheet) {
        return Ui()._ReturnRangeMergedFusion(sRef, sSheet);
    }

    tBool InsertNamedRange(tString sName, tString sRef, tString sSheet) {
        return Ui()._InsertNamedRange(sName, sRef, sSheet);
    }
    tBool DeleteNamedRange(tString sName, tString sSheet) {
        return Ui()._DeleteNamedRange(sName, sSheet);
    }
    tBool UpdateNamedRange(tString sOldName, tString sNewName, tString sNewRef, tString sSheet) {
        return Ui()._UpdateNamedRange(sOldName, sNewName, sNewRef, sSheet);
    }
    tBool IsRenameAllowed() { return Ui()._IsRenameAllowed(); }
    void SetMultiUserActive(tBool sActive) { Ui()._SetMultiUserActive(sActive); }
    tBool InsertFormulaNamed(tString sName, tString sFormula, tString sSheet) {
        return Ui()._InsertFormulaNamed(sName, sFormula, sSheet);
    }
    tBool DeleteFormulaNamed(tString sName, tString sSheet) {
        return Ui()._DeleteFormulaNamed(sName, sSheet);
    }

    tBool InsertFloatingObject(tString sName, tString sClassName, tString sTargetSheetName,
                               tString sSheet, tDouble sDiffX, tDouble sDiffY, tDouble sWidth,
                               tDouble sHeight, tDouble sOpacity, tString sAnchorCellRef) {
        return Ui()._InsertFloatingObject(sName, sClassName, sTargetSheetName, sSheet, sDiffX,
                                         sDiffY, sWidth, sHeight, sOpacity, sAnchorCellRef);
    }
    tBool DeleteFloatingObject(tString sName, tString sSheet) {
        return Ui()._DeleteFloatingObject(sName, sSheet);
    }
    tBool FloatingObjectLayout(tString sName, tDouble sDiffX, tDouble sDiffY, tDouble sWidth,
                               tDouble sHeight, tDouble sOpacity, tString sAnchorCellRef,
                               tString sSheet) {
        return Ui()._FloatingObjectLayout(sName, sDiffX, sDiffY, sWidth, sHeight, sOpacity,
                                         sAnchorCellRef, sSheet);
    }
    tBool FloatingObjectBringToFront(tString sName, tString sSheet) {
        return Ui()._FloatingObjectBringToFront(sName, sSheet);
    }
    tBool FloatingObjectAttribute(tString sName, tString sAttribute, tString sValue,
                                  tString sSheet) {
        return Ui()._FloatingObjectAttribute(sName, sAttribute, sValue, sSheet);
    }
    tBool FloatingObjectAttributes(tString sName, tString sJson, tString sSheet) {
        return Ui()._FloatingObjectAttributes(sName, sJson, sSheet);
    }

    tString JsonView(tInt sRow, tInt sCol, tDouble sViewHeight, tDouble sViewWidth, tDouble sDiffY,
                     tDouble sDiffX, tString sSheet, tBool sCss) {
        return Ui()._JsonView(sRow, sCol, sViewHeight, sViewWidth, sDiffY, sDiffX, sSheet, sCss);
    }
    tString JsonFloatingObjectsForSheet(tString sTargetSheetName, tString sSheet) {
        return Ui()._JsonFloatingObjectsForSheet(sTargetSheetName, sSheet);
    }
    tString JsonFloatingObjects(tString sSheet) { return Ui()._JsonFloatingObjects(sSheet); }
    tString JsonRightJustify(tIndex sCol, tDouble sViewWidth, tString sSheet) {
        return Ui()._JsonRightJustify(sCol, sViewWidth, sSheet);
    }
    tString JsonBottomJustify(tIndex sRow, tDouble sViewHeight, tString sSheet) {
        return Ui()._JsonBottomJustify(sRow, sViewHeight, sSheet);
    }
    tString JsonColByPixel(tIndex sColStart, tDouble sPixel, tString sSheet) {
        return Ui()._JsonColByPixel(sColStart, sPixel, sSheet);
    }
    tString JsonRowByPixel(tIndex sRowStart, tDouble sPixel, tString sSheet) {
        return Ui()._JsonRowByPixel(sRowStart, sPixel, sSheet);
    }
    tString JsonBottomRight(tString sSheet) { return Ui()._JsonBottomRight(sSheet); }
    tString JsonPixelBottomRight(tString sSheet) { return Ui()._JsonPixelBottomRight(sSheet); }
    tString JsonRangeNamed() { return Ui()._JsonRangeNamed(); }
    tString JsonFormulaNamed() { return Ui()._JsonFormulaNamed(); }
    tString JsonPrintParameters(tString sSheet) { return Ui()._JsonPrintParameters(sSheet); }
    tBool SetJsonPrintParameters(tString sJson, tString sSheet) {
        return Ui()._SetJsonPrintParameters(sJson, sSheet);
    }
    tString JsonRangeData() { return Ui()._JsonRangeData(); }
    tString JsonFindUniqueValue(tString sRef, tString sSheet) {
        return Ui()._JsonFindUniqueValue(sRef, sSheet);
    }
    tString JsonFindCell(tString sSearch, tBool sMatchCase, tBool sMatchEntireCell, tString sSheet) {
        return Ui()._JsonFindCell(sSearch, sMatchCase, sMatchEntireCell, sSheet);
    }
    tBool UndoApplyRangeData(tString sName, tString sJsonData, tString sSheet) {
        return Ui()._UndoApplyRangeData(sName, sJsonData, sSheet);
    }
    tBool UndoAddRangeData(tString sName, tString sRef, tString sJsonData, tString sSheet) {
        return Ui()._UndoAddRangeData(sName, sRef, sJsonData, sSheet);
    }

    tBool OpenCloseTreeRow(tIndex sRow, tString sSheet) {
        return Ui()._OpenCloseTreeRow(sRow, sSheet);
    }
    tBool OpenCloseTreeCol(tIndex sCol, tString sSheet) {
        return Ui()._OpenCloseTreeCol(sCol, sSheet);
    }
    tBool ChangeTreeRow(tBool sRight, tIndex sRow, tIndex sSize, tString sSheet) {
        return Ui()._ChangeTreeRow(sRight, sRow, sSize, sSheet);
    }
    tBool ChangeTreeCol(tBool sRight, tIndex sCol, tIndex sSize, tString sSheet) {
        return Ui()._ChangeTreeCol(sRight, sCol, sSize, sSheet);
    }
    tBool SplitView(tInt sCde, tIndex sPosition, tString sSheet) {
        return Ui()._SplitView(static_cast<tByte>(sCde), sPosition, sSheet);
    }
    tIndex SplitFreezeCol(tString sSheet) { return Ui()._SplitFreezeCol(sSheet); }
    tIndex SplitFreezeRow(tString sSheet) { return Ui()._SplitFreezeRow(sSheet); }

    tBool RegisterClassAttribute(tString sClassName, tString sLabel, tString sFamily) {
        return Ui()._RegisterClassAttribute(sClassName, sLabel, sFamily);
    }
    tBool AddProperty(tString sName, tString sType, tString sLabel, tSize sOrder,
                      tString sDefaultValue, tString sKind) {
        return Ui()._AddProperty(sName, sType, sLabel, sOrder, sDefaultValue, sKind);
    }
    tBool CellClass(tString sRef, tString sClassName) { return Ui()._CellClass(sRef, sClassName); }
    tString JsonCellClass() { return Ui()._JsonCellClass(); }
    tString JsonCellClassByName(tString sClassName) { return Ui()._JsonCellClassByName(sClassName); }
    tBool ApplyUnit(tString sRef, tString sFamily, tString sUnit, tString sSheet) {
        return Ui()._ApplyUnit(sRef, sFamily, sUnit, sSheet);
    }
    tString MoveCell(tInt sRow, tInt sCol, tInt sKey, tInt sMeta, tInt sTop, tInt sLeft,
                     tInt sBottom, tInt sRight, tString sSheet) {
        return Ui()._MoveCell(sRow, sCol, static_cast<tByte>(sKey), static_cast<tByte>(sMeta), sTop,
                             sLeft, sBottom, sRight, sSheet);
    }
    tString MoveToCell(tInt sRow, tInt sCol, tInt sDirection, tString sSheet) {
        return Ui()._MoveToCell(sRow, sCol, sDirection, sSheet);
    }
    tDouble SumPixelHeight(tIndex sRowStart, tIndex sRowEnd, tString sSheet) {
        return Ui()._SumPixelHeight(sRowStart, sRowEnd, sSheet);
    }
    tDouble SumPixelWidth(tIndex sColStart, tIndex sColEnd, tString sSheet) {
        return Ui()._SumPixelWidth(sColStart, sColEnd, sSheet);
    }

    tString Base10toAlpha(tInt sValue) { return Ui()._Base10toAlpha(sValue); }
    tInt AlphaToBase10(tString sValue) { return Ui()._AlphaToBase10(sValue); }
    tString ParseCell(tString sRef) { return Ui()._ParseCell(sRef); }
    tString ParseRange(tString sRef) { return Ui()._ParseRange(sRef); }
    tString QualifyRefsForSheet(tString sSheet, tString sText) {
        return Ui()._QualifyRefsForSheet(sSheet, sText);
    }
    tString StripTargetSheetFromRefs(tString sSheet, tString sText) {
        return Ui()._StripTargetSheetFromRefs(sSheet, sText);
    }
    tString CollectFormulaRefs(tString sText, tString sSheet, tIndex sRow, tIndex sCol) {
        return Ui()._CollectFormulaRefs(sText, sSheet, sRow, sCol);
    }
    tIndex MaxCol() { return Ui()._MaxCol(); }
    tIndex MaxRow() { return Ui()._MaxRow(); }
    tBool EnsureCell(tString sRef, tString sSheet) { return Ui()._EnsureCell(sRef, sSheet); }
    void SetLang(tString sLang) {
        // Locale(tString) also clears the number-formatter pool (Lang() alone does not).
        SkRoot::tApplication::Instance()->Locale(sLang);
    }
    tBool AddFunction(tString sName, tString sLabel, tString sFamily, tInt sNbArg) {
        return Ui()._AddFunction(sName, sLabel, sFamily, sNbArg);
    }
    tString Call(tString sFunctionName, tString sJsonParams) {
        return Ui()._Call(sFunctionName, sJsonParams);
    }
    tBool PostMessage(tString sMessage) { return Ui().PostMessage(sMessage); }
    tBool GetMessage(tString sMessage) { return Ui().GetMessage(sMessage); }

    void Pressure(tInt sDynamicRow, tInt sDynamicCol, tString sSheet) {
        Ui()._Pressure(sDynamicRow, sDynamicCol, sSheet);
    }
    void BeginPressureCooperative(tInt sDynamicRow, tInt sDynamicCol, tString sSheet) {
        Ui()._BeginPressureCooperative(sDynamicRow, sDynamicCol, sSheet);
    }
    tBool StepPressureCooperative(tInt sMaxRows) { return Ui()._StepPressureCooperative(sMaxRows); }
    tInt PressureCooperativeProgress() { return Ui()._PressureCooperativeProgress(); }
    tBool IsPressureCooperativeActive() { return Ui()._IsPressureCooperativeActive(); }
    void EndPressureCooperative() { Ui()._EndPressureCooperative(); }

    tBool OpenXlsx(tString sPath, tBool sMaterializeIndirect, tBool sRecalculate) {
        if (sPath.empty() || !std::filesystem::exists(sPath)) {
            throw std::runtime_error("open_xlsx: file not found: " + sPath);
        }
        try {
            SkExcel::tExcel2SpreadSheet::RegisterJavascriptCellClasses(Ui());
            Ui()._NewWorkBook(sPath);
            Ui()._DeleteSheet("Sheet1");
            SkExcel::tExcel2SpreadSheet wImporter;
            wImporter.SetMaterializeIndirectFrenchL1C1CellRefs(sMaterializeIndirect);
            wImporter.SetRecalculateAtImport(sRecalculate);
            wImporter.SetWriteSkerOnImport(false);
            if (!wImporter.ImportXlsxToApi(sPath, Ui())) {
                throw std::runtime_error("open_xlsx: ImportXlsxToApi failed: " + sPath);
            }
            return true;
        } catch (const std::runtime_error&) {
            throw;
        } catch (const std::exception& sError) {
            throw std::runtime_error(tString("open_xlsx: ") + sError.what());
        } catch (...) {
            throw std::runtime_error("open_xlsx: unknown exception: " + sPath);
        }
    }

    tBool SaveXlsx(tString sPath) {
        if (sPath.empty()) {
            throw std::runtime_error("save_xlsx: empty path");
        }
        try {
            SkExcel::tExcel2SpreadSheet::RegisterJavascriptCellClasses(Ui());
            if (!SkExcel::ExportApiToXlsx(Ui(), sPath)) {
                throw std::runtime_error("save_xlsx: ExportApiToXlsx failed: " + sPath);
            }
            return true;
        } catch (const std::runtime_error&) {
            throw;
        } catch (const std::exception& sError) {
            throw std::runtime_error(tString("save_xlsx: ") + sError.what());
        } catch (...) {
            throw std::runtime_error("save_xlsx: unknown exception: " + sPath);
        }
    }
};

} // namespace

PYBIND11_MODULE(_core, sModule) {
    sModule.doc() = "Native bindings for UISpreadSheet (same surface as the WASM module)";

    const char* wLeak = std::getenv("SK_DEBUG_LEAK");
    ReportLeakAtExit(wLeak != nullptr && wLeak[0] == '1' && wLeak[1] == '\0');
    sModule.def("debug_memory", &DebugMemory);

    py::class_<tPySpreadSheet>(sModule, "SpreadSheet")
        .def(py::init<>())
        .def("close", &tPySpreadSheet::Close)
        .def("__enter__", [](tPySpreadSheet& sSelf) -> tPySpreadSheet& { return sSelf; })
        .def("__exit__", [](tPySpreadSheet& sSelf, py::object, py::object, py::object) {
            sSelf.Close();
            return false;
        })
        .def("get_active_workbook", &tPySpreadSheet::GetActiveWorkBook)
        .def("set_active_workbook", &tPySpreadSheet::SetActiveWorkBook, py::arg("uri"))
        .def("new_workbook", &tPySpreadSheet::NewWorkBook, py::arg("uri"))
        .def("add_workbook", &tPySpreadSheet::AddWorkBook, py::arg("uri"))
        .def("delete_workbook", &tPySpreadSheet::DeleteWorkBook, py::arg("uri"))
        .def("rename_workbook", &tPySpreadSheet::RenameWorkBook, py::arg("uri"), py::arg("uri_to"))
        .def("json_workbook", &tPySpreadSheet::JsonWorkBook, py::arg("uri"))
        .def("json_workbooks", &tPySpreadSheet::JsonWorkBooks)
        .def("write_json", &tPySpreadSheet::WriteJson, py::arg("uri"))
        .def("read_json", &tPySpreadSheet::ReadJson, py::arg("json"))
        .def("recalculate_all", &tPySpreadSheet::RecalculateAll,
             py::call_guard<py::gil_scoped_release>())
        .def("begin_recalculate_all_cooperative", &tPySpreadSheet::BeginRecalculateAllCooperative)
        .def("step_recalculate_all_cooperative", &tPySpreadSheet::StepRecalculateAllCooperative,
             py::arg("max_ms"), py::call_guard<py::gil_scoped_release>())
        .def("recalculate_all_cooperative_progress",
             &tPySpreadSheet::RecalculateAllCooperativeProgress)
        .def("is_recalculate_all_cooperative_active",
             &tPySpreadSheet::IsRecalculateAllCooperativeActive)
        .def("set_cooperative_calculate_enabled", &tPySpreadSheet::SetCooperativeCalculateEnabled,
             py::arg("enabled"))
        .def("cooperative_calculate_enabled", &tPySpreadSheet::CooperativeCalculateEnabled)
        .def("undo", &tPySpreadSheet::Undo)
        .def("redo", &tPySpreadSheet::Redo)
        .def("is_undo_actif", &tPySpreadSheet::IsUndoActif)
        .def("set_undo_actif", &tPySpreadSheet::SetIsUndoActif, py::arg("active"))
        .def("set_user_interface", &tPySpreadSheet::SetUserInterface, py::arg("json"))
        .def("get_user_interface", &tPySpreadSheet::GetUserInterface)
        .def("set_extra_undo", &tPySpreadSheet::SetExtraUndo, py::arg("json"))
        .def("get_extra_undo", &tPySpreadSheet::GetExtraUndo)
        .def("set_active_sheet", &tPySpreadSheet::SetActiveSheet, py::arg("name"))
        .def("get_active_sheet", &tPySpreadSheet::GetActiveSheet)
        .def("add_sheet", &tPySpreadSheet::AddSheet, py::arg("name"), py::arg("left") = "")
        .def("rename_sheet", &tPySpreadSheet::RenameSheet, py::arg("name"), py::arg("new_name"))
        .def("swap_sheet", &tPySpreadSheet::SwapSheet, py::arg("name1"), py::arg("name2"),
             py::arg("insert_after") = false)
        .def("delete_sheet", &tPySpreadSheet::DeleteSheet, py::arg("name"))
        .def("sheets_list", &tPySpreadSheet::SheetsList)
        .def("value", &tPySpreadSheet::Value, py::arg("ref"), py::arg("value"),
             py::arg("sheet") = "")
        .def("fill_series", &tPySpreadSheet::FillSeries, py::arg("source_ref"), py::arg("dest_ref"),
             py::arg("sheet") = "")
        .def("value_attribute", &tPySpreadSheet::ValueAttribute, py::arg("ref"), py::arg("name"),
             py::arg("value"), py::arg("sheet") = "")
        .def("cell_class_attributes", &tPySpreadSheet::CellClassAttributes, py::arg("ref"),
             py::arg("attributes_json"), py::arg("sheet") = "")
        .def("value_class_calculable", &tPySpreadSheet::ValueClassCalculable, py::arg("ref"),
             py::arg("value"), py::arg("sheet") = "")
        .def("value_string", &tPySpreadSheet::ValueString, py::arg("ref"), py::arg("value"),
             py::arg("sheet") = "")
        .def("value_int", &tPySpreadSheet::ValueInt, py::arg("ref"), py::arg("value"),
             py::arg("sheet") = "")
        .def("value_double", &tPySpreadSheet::ValueDouble, py::arg("ref"), py::arg("value"),
             py::arg("sheet") = "")
        .def("get_value", &tPySpreadSheet::GetValue, py::arg("ref"), py::arg("sheet") = "")
        .def("get_calculable_scalar", &tPySpreadSheet::GetCalculableScalar, py::arg("ref"),
             py::arg("sheet") = "")
        .def("get_value_attribute", &tPySpreadSheet::GetValueAttribute, py::arg("ref"),
             py::arg("name"), py::arg("sheet") = "")
        .def("get_formula_attribute", &tPySpreadSheet::GetFormulaAttribute, py::arg("ref"),
             py::arg("name"), py::arg("sheet") = "")
        .def("get_input_value", &tPySpreadSheet::GetInputValue, py::arg("ref"),
             py::arg("sheet") = "")
        .def("get_formula", &tPySpreadSheet::GetFormula, py::arg("ref"), py::arg("sheet") = "",
             py::arg("user") = false)
        .def("cell_ref", &tPySpreadSheet::CellRef, py::arg("ref_anchor"), py::arg("ref"),
             py::arg("sheet") = "")
        .def("compile_error", &tPySpreadSheet::CompileError)
        .def("compile_error_line", &tPySpreadSheet::CompileErrorLine)
        .def("compile_error_column", &tPySpreadSheet::CompileErrorColumn)
        .def("compile_error_with_detail", &tPySpreadSheet::CompileErrorWithDetail)
        .def("raz", &tPySpreadSheet::Raz, py::arg("ref"), py::arg("keep_format") = false,
             py::arg("sheet") = "")
        .def("raz_format", &tPySpreadSheet::RazFormat, py::arg("ref"), py::arg("sheet") = "")
        .def("size_row", &tPySpreadSheet::SizeRow, py::arg("begin"), py::arg("end"), py::arg("size"),
             py::arg("sheet") = "")
        .def("size_col", &tPySpreadSheet::SizeCol, py::arg("begin"), py::arg("end"), py::arg("size"),
             py::arg("sheet") = "")
        .def("get_size_row", &tPySpreadSheet::GetSizeRow, py::arg("index"), py::arg("sheet") = "")
        .def("get_size_col", &tPySpreadSheet::GetSizeCol, py::arg("index"), py::arg("sheet") = "")
        .def("insert_row", &tPySpreadSheet::InsertRow, py::arg("begin"), py::arg("end"),
             py::arg("sheet") = "")
        .def("insert_row_by_rect", &tPySpreadSheet::InsertRowByRect, py::arg("ref"),
             py::arg("sheet") = "")
        .def("insert_row_by_rect_with_label", &tPySpreadSheet::InsertRowByRectWithLabel,
             py::arg("ref"), py::arg("label_ref"), py::arg("label_value"), py::arg("sheet") = "")
        .def("delete_row", &tPySpreadSheet::DeleteRow, py::arg("begin"), py::arg("end"),
             py::arg("sheet") = "")
        .def("delete_row_by_rect", &tPySpreadSheet::DeleteRowByRect, py::arg("ref"),
             py::arg("sheet") = "")
        .def("insert_col", &tPySpreadSheet::InsertCol, py::arg("begin"), py::arg("end"),
             py::arg("sheet") = "")
        .def("insert_col_by_rect", &tPySpreadSheet::InsertColByRect, py::arg("ref"),
             py::arg("sheet") = "")
        .def("delete_col", &tPySpreadSheet::DeleteCol, py::arg("begin"), py::arg("end"),
             py::arg("sheet") = "")
        .def("delete_col_by_rect", &tPySpreadSheet::DeleteColByRect, py::arg("ref"),
             py::arg("sheet") = "")
        .def("copy", &tPySpreadSheet::Copy, py::arg("ref"), py::arg("sheet") = "")
        .def("cut", &tPySpreadSheet::Cut, py::arg("ref"), py::arg("sheet") = "")
        .def("paste", &tPySpreadSheet::Paste, py::arg("ref"), py::arg("sheet") = "")
        .def("move", &tPySpreadSheet::Move, py::arg("source_ref"), py::arg("dest_ref"),
             py::arg("sheet") = "")
        .def("format", &tPySpreadSheet::Format, py::arg("ref"), py::arg("value"),
             py::arg("sheet") = "")
        .def("precision", &tPySpreadSheet::Precision, py::arg("ref"), py::arg("inc"),
             py::arg("sheet") = "")
        .def("conditional_format", &tPySpreadSheet::ConditionalFormat, py::arg("type"),
             py::arg("ref"), py::arg("param1") = "", py::arg("param2") = "", py::arg("param3") = "",
             py::arg("param4") = "", py::arg("param5") = "", py::arg("param6") = "",
             py::arg("param7") = "", py::arg("param8") = "", py::arg("param9") = "",
             py::arg("param10") = "", py::arg("sheet") = "")
        .def("delete_conditional_format", &tPySpreadSheet::DeleteConditionalFormat, py::arg("type"),
             py::arg("ref"), py::arg("sheet") = "")
        .def("border", &tPySpreadSheet::Border, py::arg("ref"), py::arg("border"), py::arg("value"),
             py::arg("sheet") = "")
        .def("get_format", &tPySpreadSheet::GetFormat, py::arg("ref"), py::arg("sheet") = "")
        .def("debug_format", &tPySpreadSheet::DebugFormat)
        .def("format_count", &tPySpreadSheet::FormatCount)
        .def("apply_format_string", &tPySpreadSheet::ApplyFormatString, py::arg("ref"),
             py::arg("value"), py::arg("sheet") = "")
        .def("default_format_string", &tPySpreadSheet::DefaultFormatString, py::arg("format_string"))
        .def("format_value_with_cell_format", &tPySpreadSheet::FormatValueWithCellFormat,
             py::arg("ref"), py::arg("value"), py::arg("sheet") = "")
        .def("json_format_string", &tPySpreadSheet::JsonFormatString)
        .def("json_conditional_format", &tPySpreadSheet::JsonConditionalFormat,
             py::arg("sheet") = "")
        .def("merge", &tPySpreadSheet::Merge, py::arg("ref"), py::arg("sheet") = "")
        .def("return_merged", &tPySpreadSheet::ReturnMerged, py::arg("row"), py::arg("col"),
             py::arg("sheet") = "")
        .def("return_range_merged", &tPySpreadSheet::ReturnRangeMerged, py::arg("ref"),
             py::arg("sheet") = "")
        .def("return_range_merged_fusion", &tPySpreadSheet::ReturnRangeMergedFusion, py::arg("ref"),
             py::arg("sheet") = "")
        .def("insert_named_range", &tPySpreadSheet::InsertNamedRange, py::arg("name"),
             py::arg("ref"), py::arg("sheet") = "")
        .def("delete_named_range", &tPySpreadSheet::DeleteNamedRange, py::arg("name"),
             py::arg("sheet") = "")
        .def("update_named_range", &tPySpreadSheet::UpdateNamedRange, py::arg("old_name"),
             py::arg("new_name"), py::arg("new_ref"), py::arg("sheet") = "")
        .def("is_rename_allowed", &tPySpreadSheet::IsRenameAllowed)
        .def("set_multi_user_active", &tPySpreadSheet::SetMultiUserActive, py::arg("active"))
        .def("insert_formula_named", &tPySpreadSheet::InsertFormulaNamed, py::arg("name"),
             py::arg("formula"), py::arg("sheet") = "")
        .def("delete_formula_named", &tPySpreadSheet::DeleteFormulaNamed, py::arg("name"),
             py::arg("sheet") = "")
        .def("insert_floating_object", &tPySpreadSheet::InsertFloatingObject, py::arg("name"),
             py::arg("class_name"), py::arg("target_sheet_name"), py::arg("sheet") = "",
             py::arg("diff_x") = 0.0, py::arg("diff_y") = 0.0, py::arg("width") = 0.0,
             py::arg("height") = 0.0, py::arg("opacity") = 0.0, py::arg("anchor_cell_ref") = "")
        .def("delete_floating_object", &tPySpreadSheet::DeleteFloatingObject, py::arg("name"),
             py::arg("sheet") = "")
        .def("floating_object_layout", &tPySpreadSheet::FloatingObjectLayout, py::arg("name"),
             py::arg("diff_x"), py::arg("diff_y"), py::arg("width"), py::arg("height"),
             py::arg("opacity"), py::arg("anchor_cell_ref") = "", py::arg("sheet") = "")
        .def("floating_object_bring_to_front", &tPySpreadSheet::FloatingObjectBringToFront,
             py::arg("name"), py::arg("sheet") = "")
        .def("floating_object_attribute", &tPySpreadSheet::FloatingObjectAttribute, py::arg("name"),
             py::arg("attribute"), py::arg("value"), py::arg("sheet") = "")
        .def("floating_object_attributes", &tPySpreadSheet::FloatingObjectAttributes,
             py::arg("name"), py::arg("attributes_json"), py::arg("sheet") = "")
        .def("json_view", &tPySpreadSheet::JsonView, py::arg("row"), py::arg("col"),
             py::arg("view_height"), py::arg("view_width"), py::arg("diff_y"), py::arg("diff_x"),
             py::arg("sheet") = "", py::arg("css") = false)
        .def("json_floating_objects_for_sheet", &tPySpreadSheet::JsonFloatingObjectsForSheet,
             py::arg("target_sheet_name"), py::arg("sheet") = "")
        .def("json_floating_objects", &tPySpreadSheet::JsonFloatingObjects, py::arg("sheet") = "")
        .def("json_right_justify", &tPySpreadSheet::JsonRightJustify, py::arg("col"),
             py::arg("view_width"), py::arg("sheet") = "")
        .def("json_bottom_justify", &tPySpreadSheet::JsonBottomJustify, py::arg("row"),
             py::arg("view_height"), py::arg("sheet") = "")
        .def("json_col_by_pixel", &tPySpreadSheet::JsonColByPixel, py::arg("col_start"),
             py::arg("pixel"), py::arg("sheet") = "")
        .def("json_row_by_pixel", &tPySpreadSheet::JsonRowByPixel, py::arg("row_start"),
             py::arg("pixel"), py::arg("sheet") = "")
        .def("json_bottom_right", &tPySpreadSheet::JsonBottomRight, py::arg("sheet") = "")
        .def("json_pixel_bottom_right", &tPySpreadSheet::JsonPixelBottomRight,
             py::arg("sheet") = "")
        .def("json_range_named", &tPySpreadSheet::JsonRangeNamed)
        .def("json_formula_named", &tPySpreadSheet::JsonFormulaNamed)
        .def("json_print_parameters", &tPySpreadSheet::JsonPrintParameters, py::arg("sheet") = "")
        .def("set_json_print_parameters", &tPySpreadSheet::SetJsonPrintParameters, py::arg("json"),
             py::arg("sheet") = "")
        .def("json_range_data", &tPySpreadSheet::JsonRangeData)
        .def("json_find_unique_value", &tPySpreadSheet::JsonFindUniqueValue, py::arg("ref"),
             py::arg("sheet") = "")
        .def("json_find_cell", &tPySpreadSheet::JsonFindCell, py::arg("search"),
             py::arg("match_case") = false, py::arg("match_entire_cell") = false,
             py::arg("sheet") = "")
        .def("undo_apply_range_data", &tPySpreadSheet::UndoApplyRangeData, py::arg("name"),
             py::arg("json_data"), py::arg("sheet") = "")
        .def("undo_add_range_data", &tPySpreadSheet::UndoAddRangeData, py::arg("name"),
             py::arg("ref"), py::arg("json_data") = "", py::arg("sheet") = "")
        .def("open_close_tree_row", &tPySpreadSheet::OpenCloseTreeRow, py::arg("row"),
             py::arg("sheet") = "")
        .def("open_close_tree_col", &tPySpreadSheet::OpenCloseTreeCol, py::arg("col"),
             py::arg("sheet") = "")
        .def("change_tree_row", &tPySpreadSheet::ChangeTreeRow, py::arg("right"), py::arg("row"),
             py::arg("size"), py::arg("sheet") = "")
        .def("change_tree_col", &tPySpreadSheet::ChangeTreeCol, py::arg("right"), py::arg("col"),
             py::arg("size"), py::arg("sheet") = "")
        .def("split_view", &tPySpreadSheet::SplitView, py::arg("cde"), py::arg("position"),
             py::arg("sheet") = "")
        .def("split_freeze_col", &tPySpreadSheet::SplitFreezeCol, py::arg("sheet") = "")
        .def("split_freeze_row", &tPySpreadSheet::SplitFreezeRow, py::arg("sheet") = "")
        .def("register_class_attribute", &tPySpreadSheet::RegisterClassAttribute,
             py::arg("class_name"), py::arg("label"), py::arg("family"))
        .def("add_property", &tPySpreadSheet::AddProperty, py::arg("name"), py::arg("type"),
             py::arg("label"), py::arg("order"), py::arg("default_value"), py::arg("kind") = "")
        .def("cell_class", &tPySpreadSheet::CellClass, py::arg("ref"), py::arg("class_name"))
        .def("json_cell_class", &tPySpreadSheet::JsonCellClass)
        .def("json_cell_class_by_name", &tPySpreadSheet::JsonCellClassByName, py::arg("class_name"))
        .def("apply_unit", &tPySpreadSheet::ApplyUnit, py::arg("ref"), py::arg("family"),
             py::arg("unit"), py::arg("sheet") = "")
        .def("move_cell", &tPySpreadSheet::MoveCell, py::arg("row"), py::arg("col"), py::arg("key"),
             py::arg("meta"), py::arg("top"), py::arg("left"), py::arg("bottom"), py::arg("right"),
             py::arg("sheet") = "")
        .def("move_to_cell", &tPySpreadSheet::MoveToCell, py::arg("row"), py::arg("col"),
             py::arg("direction"), py::arg("sheet") = "")
        .def("sum_pixel_height", &tPySpreadSheet::SumPixelHeight, py::arg("row_start"),
             py::arg("row_end"), py::arg("sheet") = "")
        .def("sum_pixel_width", &tPySpreadSheet::SumPixelWidth, py::arg("col_start"),
             py::arg("col_end"), py::arg("sheet") = "")
        .def("base10_to_alpha", &tPySpreadSheet::Base10toAlpha, py::arg("value"))
        .def("alpha_to_base10", &tPySpreadSheet::AlphaToBase10, py::arg("value"))
        .def("parse_cell", &tPySpreadSheet::ParseCell, py::arg("ref"))
        .def("parse_range", &tPySpreadSheet::ParseRange, py::arg("ref"))
        .def("qualify_refs_for_sheet", &tPySpreadSheet::QualifyRefsForSheet, py::arg("sheet"),
             py::arg("text"))
        .def("strip_target_sheet_from_refs", &tPySpreadSheet::StripTargetSheetFromRefs,
             py::arg("sheet"), py::arg("text"))
        .def("collect_formula_refs", &tPySpreadSheet::CollectFormulaRefs, py::arg("text"),
             py::arg("sheet"), py::arg("row"), py::arg("col"))
        .def("max_col", &tPySpreadSheet::MaxCol)
        .def("max_row", &tPySpreadSheet::MaxRow)
        .def("ensure_cell", &tPySpreadSheet::EnsureCell, py::arg("ref"), py::arg("sheet") = "")
        .def("set_lang", &tPySpreadSheet::SetLang, py::arg("lang"))
        .def("add_function", &tPySpreadSheet::AddFunction, py::arg("name"), py::arg("label"),
             py::arg("family"), py::arg("nb_arg"))
        .def("call", &tPySpreadSheet::Call, py::arg("function_name"), py::arg("json_params"))
        .def("post_message", &tPySpreadSheet::PostMessage, py::arg("message"))
        .def("get_message", &tPySpreadSheet::GetMessage, py::arg("message"))
        .def("pressure", &tPySpreadSheet::Pressure, py::arg("dynamic_row"), py::arg("dynamic_col"),
             py::arg("sheet") = "", py::call_guard<py::gil_scoped_release>())
        .def("begin_pressure_cooperative", &tPySpreadSheet::BeginPressureCooperative,
             py::arg("dynamic_row"), py::arg("dynamic_col"), py::arg("sheet") = "")
        .def("step_pressure_cooperative", &tPySpreadSheet::StepPressureCooperative,
             py::arg("max_rows"), py::call_guard<py::gil_scoped_release>())
        .def("pressure_cooperative_progress", &tPySpreadSheet::PressureCooperativeProgress)
        .def("is_pressure_cooperative_active", &tPySpreadSheet::IsPressureCooperativeActive)
        .def("end_pressure_cooperative", &tPySpreadSheet::EndPressureCooperative,
             py::call_guard<py::gil_scoped_release>())
        .def("open_xlsx", &tPySpreadSheet::OpenXlsx, py::arg("path"),
             py::arg("materialize_indirect") = true, py::arg("recalculate") = false,
             py::call_guard<py::gil_scoped_release>())
        .def("save_xlsx", &tPySpreadSheet::SaveXlsx, py::arg("path"),
             py::call_guard<py::gil_scoped_release>());
}
