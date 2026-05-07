# Implementation Design: FormOrdenPago.frm Improvements

**Project**: SPC-Core
**Change**: FormOrdenPago-improvements
**Version**: 1.0
**Date**: 2026-05-07
**Status**: Design Phase

---

## Executive Summary

This document provides a step-by-step implementation plan for enhancing `FormOrdenPago.frm` with two new grid sections (Transferencia and Facturas), inline editing capabilities, robust date validation, and enhanced printing features. The design follows a phased approach to minimize breaking changes and ensure backward compatibility.

**Key Deliverables**:
- Database schema migration (2 new MEMO columns)
- 2 new MSFlexGrid controls with full CRUD operations
- Inline editing system replacing InputBox pattern
- Date validation with DD/MM/YYYY format enforcement
- Enhanced printing with conditional logo display
- Updated calculation logic including new payment methods

---

## Table of Contents

1. [Dependency Graph & Implementation Order](#1-dependency-graph--implementation-order)
2. [Code Modules & Functions](#2-code-modules--functions)
3. [Database Migration Strategy](#3-database-migration-strategy)
4. [UI Controls & Layout](#4-ui-controls--layout)
5. [Inline Editing Architecture](#5-inline-editing-architecture)
6. [Validation Framework](#6-validation-framework)
7. [Calculation Engine Updates](#7-calculation-engine-updates)
8. [Serialization Integration](#8-serialization-integration)
9. [Printing Enhancement](#9-printing-enhancement)
10. [Testing Strategy](#10-testing-strategy)
11. [Rollback Plan](#11-rollback-plan)
12. [Summary of Changes](#12-summary-of-changes)

---

## 1. Dependency Graph & Implementation Order

### Critical Path Analysis

The implementation must follow this strict order to prevent breaking existing functionality:

```
Phase 1: Database Schema
         ↓
Phase 2: UI Controls (structure only)
         ↓
Phase 3: Grid Initialization
         ↓
Phase 4: Serialization Integration
         ↓
Phase 5: Calculation Logic
         ↓
Phase 6: Load/Save Operations
         ↓
Phase 7: Inline Editing System
         ↓
Phase 8: Date Validation
         ↓
Phase 9: Tab Order Fix
         ↓
Phase 10: Printing Enhancement
         ↓
Phase 11: Cleanup & Finalization
```

### Phase Breakdown

#### **Phase 1: Database Schema Migration** (FIRST - Before any code changes)

**Rationale**: Schema changes must be backward-compatible and deployed before code expects new columns.

**Actions**:
1. Modify `EnsureSchemaOrdenPago()` to include new columns in CREATE TABLE statement
2. Add migration logic to ALTER TABLE for existing databases
3. Test NULL handling with `NzS()` wrapper

**Safety**: ADD COLUMN is non-breaking; NULL values are gracefully handled by existing code.

**Estimated Time**: 30 minutes

---

#### **Phase 2: UI Controls & Layout** (Structure without logic)

**Rationale**: Add controls to form without event handlers or business logic to avoid runtime errors.

**Actions**:
1. Add `grdTransferencia` MSFlexGrid to Frame8
2. Add `grdFacturas` MSFlexGrid to Frame7
3. Add `txtSubTransferencia` and `txtSubFacturas` TextBoxes
4. Add `cmdAddTransferencia`, `cmdRemoveTransferencia`, `cmdAddFacturas`, `cmdRemoveFacturas` buttons
5. Create `txtGridEditor` overlay TextBox (initially invisible)

**Safety**: New controls don't interfere with existing controls; no event handlers yet.

**Estimated Time**: 45 minutes

---

#### **Phase 3: Grid Initialization & Formatting**

**Rationale**: Initialize new grids with headers and column widths before any data operations.

**Actions**:
1. Extend `InitGrids()` to configure `grdTransferencia` (4 columns)
2. Extend `InitGrids()` to configure `grdFacturas` (4 columns)
3. Add enum definitions (`eColTransferencia`, `eColFacturas`)
4. Create `txtGridEditor` instance in `Form_Load`

**Safety**: `InitGrids()` already exists; we're just adding configuration for new grids.

**Estimated Time**: 30 minutes

---

#### **Phase 4: Serialization Integration**

**Rationale**: Leverage existing `SerializarGrid()` and `DeserializarGrid()` functions (no changes needed).

**Actions**:
1. Verify existing functions are grid-agnostic (they are - based on code review)
2. Document usage pattern for new grids
3. No code changes required (functions already generic)

**Safety**: Zero risk - no modifications to existing serialization logic.

**Estimated Time**: 15 minutes (verification only)

---

#### **Phase 5: Calculation Logic** (RecalcularTodo)

**Rationale**: Extend financial calculations to include new payment methods.

**Actions**:
1. Modify `RecalcularTodo()` to add:
   - `subTransferencia = SumarColumna(grdTransferencia, colTransfImporte)`
   - `subFacturas = SumarColumna(grdFacturas, colFactImporte)`
2. Update `totalPago` formula: `subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB`
3. Update UI display for new subtotals

**Safety**: Extends existing formula without breaking it; `SumarColumna()` is already generic.

**Estimated Time**: 20 minutes

---

#### **Phase 6: Load/Save Operations**

**Rationale**: Integrate new grids into persistence layer.

**Actions**:
1. Modify `cmdGuardar_Click()`:
   - Add: `rs!DetalleTransferencia = SerializarGrid(grdTransferencia)`
   - Add: `rs!DetalleFacturas = SerializarGrid(grdFacturas)`
2. Modify `CargarOrdenPorNumero()`:
   - Add: `DeserializarGrid grdTransferencia, NzS(rs!DetalleTransferencia)`
   - Add: `DeserializarGrid grdFacturas, NzS(rs!DetalleFacturas)`
3. Modify `LimpiarFormulario()`:
   - Add: `ResetGridRows grdTransferencia`
   - Add: `ResetGridRows grdFacturas`
   - Reset subtotal textboxes

**Safety**: `NzS()` handles NULL gracefully; serialization functions are tested and stable.

**Estimated Time**: 30 minutes

---

#### **Phase 7: Inline Editing System** (CRITICAL - Breaking change)

**Rationale**: Replace InputBox with professional inline editing; requires careful event handling.

**Actions**:
1. Add module-level variables for editor state
2. Implement `EditarCeldaGrid()` with TextBox overlay positioning
3. Implement `FinalizarEdicion()` with save/cancel logic
4. Implement `ValidarYFormatearCelda()` with column-specific rules
5. Wire up event handlers for all 5 grids (DblClick, KeyDown)
6. Implement `txtGridEditor_KeyDown()` and `txtGridEditor_LostFocus()`

**Safety**: Old `EditarCeldaGrid()` can be kept as fallback; new implementation is opt-in via event handlers.

**Testing**: Test each grid independently before deploying all handlers.

**Estimated Time**: 2 hours

---

#### **Phase 8: Date Validation**

**Rationale**: Add robust date validation for `txtFecha` field.

**Actions**:
1. Add module-level `m_FechaAnterior` variable
2. Implement `txtFecha_LostFocus()` with validation logic
3. Implement `ValidarFechaDD_MM_YYYY()` helper function
4. Implement `ParseDateDD_MM_YYYY()` helper function

**Safety**: Validation is on LostFocus; doesn't interfere with form load or save operations.

**Estimated Time**: 45 minutes

---

#### **Phase 9: Tab Order Fix**

**Rationale**: Improve UX with logical keyboard navigation.

**Actions**:
1. Assign TabIndex sequentially to all 33+ controls
2. Follow left→right, top→bottom flow
3. Test Tab key navigation

**Safety**: Pure UI improvement; no business logic impact.

**Estimated Time**: 30 minutes

---

#### **Phase 10: Printing Enhancement**

**Rationale**: Add new sections and conditional logo to printed output.

**Actions**:
1. Implement `TieneFacturas()` helper function
2. Add conditional logo loading at start of `cmdImprimir_Click()`
3. Implement `PrintSeccionTransferencia()` subroutine
4. Implement `PrintSeccionFacturas()` subroutine
5. Update RESUMEN section to include new subtotals
6. Set `Printer.Copies = 2` as default

**Safety**: Printing is isolated; errors don't affect data integrity.

**Estimated Time**: 1.5 hours

---

#### **Phase 11: Cleanup & Edge Cases**

**Rationale**: Finalize implementation and handle corner cases.

**Actions**:
1. Remove old `EditarCeldaGrid()` InputBox version (after testing inline editing)
2. Add validation for "at least one payment method required" in `cmdGuardar_Click()`
3. Implement `cmdAddTransferencia_Click()`, `cmdDelTransferencia_Click()`
4. Implement `cmdAddFacturas_Click()`, `cmdDelFacturas_Click()`
5. Add LostFocus handlers for new subtotal textboxes
6. Test all edge cases (empty grids, large data, special characters)

**Estimated Time**: 1 hour

---

**Total Estimated Time**: ~8 hours

---

## 2. Code Modules & Functions

### New Module-Level Variables (Declarations Section)

Add to top of `FormOrdenPago.frm`, after existing Private variables:

```vb
' ============================================================
' NEW: Inline editing support
' ============================================================
Private WithEvents txtGridEditor As TextBox
Private m_EditGrid As MSFlexGrid
Private m_EditRow As Integer
Private m_EditCol As Integer
Private m_EditingActive As Boolean

' ============================================================
' NEW: Date validation
' ============================================================
Private m_FechaAnterior As String
```

---

### New Enum Definitions

Add after existing `eColOtros` enum (line 865):

```vb
' ============================================================
' NEW: Column enums for new grids
' ============================================================
Private Enum eColTransferencia
    colTransfBanco = 0
    colTransfNroCuenta = 1
    colTransfCUIT = 2
    colTransfImporte = 3
End Enum

Private Enum eColFacturas
    colFactNumero = 0
    colFactFecha = 1
    colFactImporte = 2
    colFactDescripcion = 3
End Enum
```

---

### Modified: `InitGrids()` (Line 949-995)

**Current code**: Initializes 3 grids (Deuda, Cheques, Otros)

**Modification**: Add initialization for 2 new grids

**Add at end of function (before Exit Sub)**:

```vb
    ' ============================================================
    ' NEW: Transferencia
    ' ============================================================
    With grdTransferencia
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colTransfBanco) = "Banco"
        .TextMatrix(0, colTransfNroCuenta) = "Nro. Cuenta"
        .TextMatrix(0, colTransfCUIT) = "CUIT"
        .TextMatrix(0, colTransfImporte) = "Importe"
        .ColWidth(colTransfBanco) = 1800
        .ColWidth(colTransfNroCuenta) = 1600
        .ColWidth(colTransfCUIT) = 1500
        .ColWidth(colTransfImporte) = 1400
    End With

    ' ============================================================
    ' NEW: Facturas
    ' ============================================================
    With grdFacturas
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colFactNumero) = "Numero"
        .TextMatrix(0, colFactFecha) = "Fecha"
        .TextMatrix(0, colFactImporte) = "Importe"
        .TextMatrix(0, colFactDescripcion) = "Descripcion"
        .ColWidth(colFactNumero) = 1500
        .ColWidth(colFactFecha) = 1200
        .ColWidth(colFactImporte) = 1400
        .ColWidth(colFactDescripcion) = 1675
    End With
```

---

### Modified: `Form_Load()` (Line 904-932)

**Current code**: Basic initialization

**Add before `Exit Sub` (after line 926)**:

```vb
    ' ============================================================
    ' NEW: Create inline editor overlay
    ' ============================================================
    Set txtGridEditor = Me.Controls.Add("VB.TextBox", "txtGridEditor", Me)
    With txtGridEditor
        .Visible = False
        .BorderStyle = 1  ' Fixed Single
        .FontName = "MS Sans Serif"
        .FontSize = 8.25
    End With

    ' ============================================================
    ' NEW: Initialize date validation
    ' ============================================================
    m_FechaAnterior = Format$(Date, "dd/mm/yyyy")
```

---

### Modified: `RecalcularTodo()` (Line 1670-1704)

**Current code**: Calculates totals for 3 grids

**Replace entire function with**:

```vb
Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim subTransferencia As Currency  ' NEW
    Dim subFacturas As Currency       ' NEW
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    ' Calculate subtotals
    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)
    subTransferencia = SumarColumna(grdTransferencia, colTransfImporte)  ' NEW
    subFacturas = SumarColumna(grdFacturas, colFactImporte)              ' NEW

    efectivo = ParseCurrency(txtEfectivo.text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    ' Updated formula with new payment methods
    totalPago = subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB
    saldo = subDeuda - totalPago

    ' Update UI
    m_Cargando = True
    txtSubDeuda.text = FormatMoney(subDeuda)
    txtSubCheques.text = FormatMoney(subCheques)
    txtSubOtros.text = FormatMoney(subOtros)
    txtSubTransferencia.text = FormatMoney(subTransferencia)  ' NEW
    txtSubFacturas.text = FormatMoney(subFacturas)            ' NEW
    txtTotalPago.text = FormatMoney(totalPago)
    txtSaldo.text = FormatMoney(saldo)
    txtImporteLetras.text = EnLetras(CStr(totalPago))
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

### Modified: `LimpiarFormulario()` (Line 997-1029)

**Add after line 1016 (after `ResetGridRows grdOtros`)**:

```vb
    ResetGridRows grdTransferencia  ' NEW
    ResetGridRows grdFacturas       ' NEW

    txtSubTransferencia.text = "0,00"  ' NEW
    txtSubFacturas.text = "0,00"       ' NEW
```

---

### Modified: `cmdGuardar_Click()` (Line 1721-1783)

**Modify variable declarations (line 1726)**:

```vb
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String
    Dim transfTxt As String, factTxt As String  ' NEW
```

**Add after line 1739 (after `otrosTxt = SerializarGrid(grdOtros)`)**:

```vb
    transfTxt = SerializarGrid(grdTransferencia)  ' NEW
    factTxt = SerializarGrid(grdFacturas)          ' NEW
```

**Add after line 1763 (after `rs!DetalleOtros = otrosTxt`)**:

```vb
    rs!DetalleTransferencia = transfTxt  ' NEW
    rs!DetalleFacturas = factTxt          ' NEW
```

---

### Modified: `CargarOrdenPorNumero()` (Line 1830-1878)

**Add after line 1865 (after `DeserializarGrid grdOtros, NzS(rs!DetalleOtros)`)**:

```vb
    DeserializarGrid grdTransferencia, NzS(rs!DetalleTransferencia)  ' NEW
    DeserializarGrid grdFacturas, NzS(rs!DetalleFacturas)            ' NEW
```

---

### Modified: `EnsureSchemaOrdenPago()` (Line 2063-2092)

**Replace entire function**:

```vb
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        ' Create full table with new columns included
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "TotalDeuda CURRENCY, " & _
            "SubtotalCheques CURRENCY, " & _
            "SubtotalOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetencionIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "MontoLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "DetalleTransferencia MEMO, " & _
            "DetalleFacturas MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    Else
        ' Migration: Add new columns if missing
        On Error Resume Next
        BaseSPC.Execute "ALTER TABLE OrdenPago ADD COLUMN DetalleTransferencia MEMO"
        BaseSPC.Execute "ALTER TABLE OrdenPago ADD COLUMN DetalleFacturas MEMO"
        On Error GoTo EH
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

### NEW: Inline Editing Functions

**Add after existing `EditarCeldaGrid()` (line 1117-1146)** - this will eventually replace it:

```vb
' ============================================================
' NEW: Inline editing system (replaces InputBox pattern)
' ============================================================

Private Sub EditarCeldaGrid_Inline(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim cellLeft As Long, cellTop As Long
    Dim cellWidth As Long, cellHeight As Long

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub  ' No editing header row

    ' Calculate cell position relative to FORM
    ' Note: g.Parent is the Frame containing the grid
    cellLeft = g.Left + g.Parent.Left + g.ColPos(c)
    cellTop = g.Top + g.Parent.Top + g.RowPos(r)
    cellWidth = g.ColWidth(c)
    cellHeight = g.RowHeight(r)

    ' Position overlay TextBox
    With txtGridEditor
        .Move cellLeft, cellTop, cellWidth, cellHeight
        .Text = g.TextMatrix(r, c)
        .Visible = True
        .SetFocus
        .SelStart = 0
        .SelLength = Len(.Text)  ' Select all for quick overwrite
    End With

    ' Remember edit context
    Set m_EditGrid = g
    m_EditRow = r
    m_EditCol = c
    m_EditingActive = True

    Exit Sub
EH:
    MsgBox "Error iniciando edicion: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ============================================================
' txtGridEditor event handlers
' ============================================================

Private Sub txtGridEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        FinalizarEdicion True  ' Save
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        FinalizarEdicion False  ' Cancel
    End If
End Sub

Private Sub txtGridEditor_LostFocus()
    If m_EditingActive Then
        FinalizarEdicion True  ' Auto-save on focus loss
    End If
End Sub

Private Sub FinalizarEdicion(ByVal guardar As Boolean)
    On Error GoTo EH

    If Not m_EditingActive Then Exit Sub

    Dim valorNuevo As String
    Dim valorFinal As String

    If guardar Then
        valorNuevo = Trim$(txtGridEditor.Text)

        ' Apply column-specific validation and formatting
        If Not ValidarYFormatearCelda(m_EditGrid, m_EditRow, m_EditCol, valorNuevo, valorFinal) Then
            ' Validation failed - keep editor open
            Exit Sub
        End If

        ' Save to grid
        m_EditGrid.TextMatrix(m_EditRow, m_EditCol) = valorFinal

        ' Recalculate if importe column
        RecalcularTodo
    End If

    ' Hide editor
    txtGridEditor.Visible = False
    m_EditingActive = False
    Set m_EditGrid = Nothing

    Exit Sub
EH:
    MsgBox "Error finalizando edicion: " & Err.Description, vbExclamation, "Orden de Pago"
    txtGridEditor.Visible = False
    m_EditingActive = False
End Sub

Private Function ValidarYFormatearCelda(ByRef g As MSFlexGrid, ByVal r As Integer, ByVal c As Integer, ByVal valorNuevo As String, ByRef valorFinal As String) As Boolean
    On Error GoTo EH

    ValidarYFormatearCelda = False
    valorFinal = valorNuevo

    ' Identify grid and column type
    Select Case True
        ' --------------------------------------------------------
        ' grdDeuda
        ' --------------------------------------------------------
        Case g.Name = "grdDeuda" And c = colDeudaImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        ' --------------------------------------------------------
        ' grdCheques
        ' --------------------------------------------------------
        Case g.Name = "grdCheques" And c = colChequeImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdCheques" And c = colChequeFecha
            If Len(valorNuevo) > 0 And Not ValidarFechaDD_MM_YYYY(valorNuevo) Then
                MsgBox "Fecha invalida. Use formato DD/MM/YYYY", vbExclamation, "Validacion"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        ' --------------------------------------------------------
        ' grdOtros
        ' --------------------------------------------------------
        Case g.Name = "grdOtros" And c = colOtroImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        ' --------------------------------------------------------
        ' NEW: grdTransferencia
        ' --------------------------------------------------------
        Case g.Name = "grdTransferencia" And c = colTransfImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdTransferencia" And c = colTransfCUIT
            If Len(valorNuevo) > 0 And Not ValidarCUIT(valorNuevo) Then
                MsgBox "CUIT invalido. Use formato XX-XXXXXXXX-X (ej: 20-12345678-9)", vbExclamation, "Validacion"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        ' --------------------------------------------------------
        ' NEW: grdFacturas
        ' --------------------------------------------------------
        Case g.Name = "grdFacturas" And c = colFactImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdFacturas" And c = colFactNumero
            If Len(valorNuevo) > 0 And Not ValidarNumeroFactura(valorNuevo) Then
                MsgBox "Numero de factura invalido. Use formato XXXX-XXXXXXXX (ej: 0001-00012345)", vbExclamation, "Validacion"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        Case g.Name = "grdFacturas" And c = colFactFecha
            If Len(valorNuevo) > 0 And Not ValidarFechaDD_MM_YYYY(valorNuevo) Then
                MsgBox "Fecha invalida. Use formato DD/MM/YYYY", vbExclamation, "Validacion"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        ' Default: text columns (no formatting)
        Case Else
            ValidarYFormatearCelda = True
    End Select

    Exit Function
EH:
    MsgBox "Error en validacion: " & Err.Description, vbExclamation, "Validacion"
    ValidarYFormatearCelda = False
End Function
```

---

### NEW: Validation Helper Functions

**Add after `ValidarYFormatearCelda()`**:

```vb
' ============================================================
' NEW: Validation helpers
' ============================================================

Private Function ValidarFechaDD_MM_YYYY(ByVal fechaStr As String) As Boolean
    On Error GoTo EH

    Dim partes() As String
    Dim dia As Integer, mes As Integer, anno As Integer

    ValidarFechaDD_MM_YYYY = False

    If Len(fechaStr) <> 10 Then Exit Function
    If Mid(fechaStr, 3, 1) <> "/" Or Mid(fechaStr, 6, 1) <> "/" Then Exit Function

    partes = Split(fechaStr, "/")
    If UBound(partes) <> 2 Then Exit Function

    dia = CInt(partes(0))
    mes = CInt(partes(1))
    anno = CInt(partes(2))

    ' Basic range checks
    If mes < 1 Or mes > 12 Then Exit Function
    If dia < 1 Or dia > 31 Then Exit Function
    If anno < 1900 Or anno > 2100 Then Exit Function

    ' Try to create date (will fail if invalid like 31/02/2026)
    Dim testDate As Date
    testDate = DateSerial(anno, mes, dia)

    ValidarFechaDD_MM_YYYY = True
    Exit Function
EH:
    ValidarFechaDD_MM_YYYY = False
End Function

Private Function ParseDateDD_MM_YYYY(ByVal fechaStr As String) As Date
    On Error GoTo EH

    Dim partes() As String
    Dim dia As Integer, mes As Integer, anno As Integer

    partes = Split(fechaStr, "/")
    dia = CInt(partes(0))
    mes = CInt(partes(1))
    anno = CInt(partes(2))

    ParseDateDD_MM_YYYY = DateSerial(anno, mes, dia)
    Exit Function
EH:
    ParseDateDD_MM_YYYY = Date  ' Fallback to today
End Function

Private Function ValidarCUIT(ByVal cuit As String) As Boolean
    Dim i As Integer

    ValidarCUIT = False

    If Len(cuit) <> 13 Then Exit Function
    If Mid(cuit, 3, 1) <> "-" Then Exit Function
    If Mid(cuit, 12, 1) <> "-" Then Exit Function

    ' Check first 2 digits
    For i = 1 To 2
        If Not IsNumeric(Mid(cuit, i, 1)) Then Exit Function
    Next

    ' Check middle 8 digits
    For i = 4 To 11
        If Not IsNumeric(Mid(cuit, i, 1)) Then Exit Function
    Next

    ' Check last digit
    If Not IsNumeric(Mid(cuit, 13, 1)) Then Exit Function

    ValidarCUIT = True
End Function

Private Function ValidarNumeroFactura(ByVal numero As String) As Boolean
    Dim i As Integer

    ValidarNumeroFactura = False

    If Len(numero) <> 13 Then Exit Function
    If Mid(numero, 5, 1) <> "-" Then Exit Function

    ' Check first 4 digits (punto de venta)
    For i = 1 To 4
        If Not IsNumeric(Mid(numero, i, 1)) Then Exit Function
    Next

    ' Check last 8 digits (numero)
    For i = 6 To 13
        If Not IsNumeric(Mid(numero, i, 1)) Then Exit Function
    Next

    ValidarNumeroFactura = True
End Function
```

---

### NEW: txtFecha Date Validation

**Add new event handler**:

```vb
' ============================================================
' NEW: txtFecha validation
' ============================================================

Private Sub txtFecha_LostFocus()
    On Error GoTo EH

    Dim fechaStr As String
    Dim fechaVal As Date

    fechaStr = Trim$(txtFecha.Text)

    ' Empty -> auto-fill today
    If Len(fechaStr) = 0 Then
        txtFecha.Text = Format$(Date, "dd/mm/yyyy")
        m_FechaAnterior = txtFecha.Text
        Exit Sub
    End If

    ' Parse and validate
    If Not ValidarFechaDD_MM_YYYY(fechaStr) Then
        MsgBox "Fecha invalida. Use formato DD/MM/YYYY", vbExclamation, "Validacion"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    fechaVal = ParseDateDD_MM_YYYY(fechaStr)

    ' Check future
    If fechaVal > Date Then
        MsgBox "La fecha no puede ser futura", vbExclamation, "Validacion"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    ' Check ancient
    If Year(fechaVal) < 1900 Then
        MsgBox "Fecha muy antigua (ano < 1900)", vbExclamation, "Validacion"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    ' Valid -> normalize format and save
    txtFecha.Text = Format$(fechaVal, "dd/mm/yyyy")
    m_FechaAnterior = txtFecha.Text

    Exit Sub
EH:
    MsgBox "Error validando fecha: " & Err.Description, vbExclamation, "Validacion"
    txtFecha.Text = m_FechaAnterior
End Sub
```

---

### NEW: Grid Event Handlers (Replace Existing)

**Replace existing DblClick handlers (lines 1105-1115)**:

```vb
' ============================================================
' Grid DblClick handlers (use new inline editing)
' ============================================================

Private Sub grdDeuda_DblClick()
    EditarCeldaGrid_Inline grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid_Inline grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid_Inline grdOtros
End Sub

' NEW: Transferencia grid
Private Sub grdTransferencia_DblClick()
    EditarCeldaGrid_Inline grdTransferencia
End Sub

' NEW: Facturas grid
Private Sub grdFacturas_DblClick()
    EditarCeldaGrid_Inline grdFacturas
End Sub
```

**Note**: Existing KeyDown handlers for grdDeuda, grdCheques, grdOtros already call `EditarCeldaGrid` (lines 884-903). Update those to call `EditarCeldaGrid_Inline` instead.

**Add new KeyDown handlers for new grids**:

```vb
' NEW: Transferencia KeyDown
Private Sub grdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid_Inline grdTransferencia
    End If
End Sub

' NEW: Facturas KeyDown
Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid_Inline grdFacturas
    End If
End Sub
```

---

### NEW: Button Event Handlers

**Add after existing cmdAddOtro_Click (line 1065-1072)**:

```vb
' ============================================================
' NEW: Transferencia buttons
' ============================================================

Private Sub cmdAddTransferencia_Click()
    AddRow grdTransferencia
End Sub

Private Sub cmdDelTransferencia_Click()
    DelCurrentRow grdTransferencia
    RecalcularTodo
End Sub

' ============================================================
' NEW: Facturas buttons
' ============================================================

Private Sub cmdAddFacturas_Click()
    AddRow grdFacturas
End Sub

Private Sub cmdDelFacturas_Click()
    DelCurrentRow grdFacturas
    RecalcularTodo
End Sub
```

---

### NEW: Subtotal TextBox LostFocus Handlers

**Add near other LostFocus handlers (after line 2359)**:

```vb
' ============================================================
' NEW: Subtotal formatting for new grids
' ============================================================

Private Sub txtSubTransferencia_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubTransferencia
End Sub

Private Sub txtSubFacturas_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubFacturas
End Sub
```

---

### NEW: Printing Helper Function

**Add before `cmdImprimir_Click()` (line 1166)**:

```vb
' ============================================================
' NEW: Printing helper - check if facturas grid has data
' ============================================================

Private Function TieneFacturas() As Boolean
    Dim i As Integer
    For i = 1 To grdFacturas.Rows - 1
        If Trim$(grdFacturas.TextMatrix(i, 0)) <> "" Then
            TieneFacturas = True
            Exit Function
        End If
    Next i
    TieneFacturas = False
End Function
```

---

### Modified: `cmdImprimir_Click()` (Line 1166-1418)

**This is a major rewrite. Replace entire function**:

```vb
Private Sub cmdImprimir_Click()
    On Error GoTo ErrHandler

    Dim y As Single
    Dim xLeft As Single
    Dim xRight As Single
    Dim lineH As Single
    Dim pageBottom As Single

    Dim i As Long
    Dim detalle As String
    Dim importe As String
    Dim nroCheque As String
    Dim banco As String
    Dim vto As String
    Dim retText As String

    ' NEW: Variables for new grids
    Dim nroCuenta As String
    Dim cuit As String
    Dim nroFact As String
    Dim fechaFact As String
    Dim descFact As String

    Dim deudaRows As Long
    Dim chequeRows As Long
    Dim transfRows As Long  ' NEW
    Dim factRows As Long     ' NEW

    xLeft = 600
    xRight = 7800
    lineH = 240
    pageBottom = 10600

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = False

    ' NEW: Default to 2 copies
    Printer.Copies = 2

    y = 600

    ' ============================================================
    ' NEW: Conditional logo based on Facturas presence
    ' ============================================================
    If TieneFacturas() Then
        ' Has facturas data -> print logo
        On Error Resume Next
        Dim logoPath As String
        logoPath = App.Path & "\Quilplac2.jpg"

        If Dir(logoPath) <> "" Then
            ' Logo exists -> display
            Printer.PaintPicture LoadPicture(logoPath), 600, 200, 7200, 1200
            y = 1600  ' Start content below logo
        Else
            ' Logo missing -> fallback to text
            Printer.CurrentX = 3000
            Printer.CurrentY = 600
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print "ORDEN DE PAGO"
            Printer.FontBold = False
            Printer.FontSize = 10
            y = 1200
        End If
        On Error GoTo ErrHandler
    Else
        ' No facturas -> text header only
        Printer.CurrentX = 3000
        Printer.CurrentY = 600
        Printer.FontBold = True
        Printer.FontSize = 14
        Printer.Print "ORDEN DE PAGO"
        Printer.FontBold = False
        Printer.FontSize = 10
        y = 1200
    End If

    ' ============================================================
    ' Header: Nro de Orden, Proveedor, Fecha
    ' ============================================================
    Printer.CurrentX = 120
    Printer.CurrentY = y
    Printer.FontBold = True
    Printer.Print "ORDEN DE PAGO Nro: " & Trim$(txtNroOrden.text)
    Printer.FontBold = False
    y = y + lineH * 2

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Proveedor: " & Trim$(txtProveedor.text)

    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print "Fecha: " & Trim$(txtFecha.text)
    y = y + lineH * 2

    ' ============================================================
    ' DEUDA
    ' ============================================================
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "DEUDA"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    deudaRows = grdDeuda.Rows - 1
    If deudaRows >= 1 Then
        For i = 1 To deudaRows
            detalle = Trim$(grdDeuda.TextMatrix(i, 0))
            importe = Trim$(grdDeuda.TextMatrix(i, 1))

            If Len(detalle) > 0 Or Len(importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If

                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$(detalle, 55)

                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(importe))

                y = y + lineH
            End If
        Next i
    End If

    y = y + lineH / 2
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Deuda:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.text))
    y = y + lineH * 2

    ' ============================================================
    ' CHEQUES
    ' ============================================================
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "CHEQUES"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    chequeRows = grdCheques.Rows - 1
    If chequeRows >= 1 Then
        For i = 1 To chequeRows
            banco = Trim$(grdCheques.TextMatrix(i, 0))
            nroCheque = Trim$(grdCheques.TextMatrix(i, 1))
            vto = Trim$(grdCheques.TextMatrix(i, 2))
            importe = Trim$(grdCheques.TextMatrix(i, 3))

            If Len(banco) > 0 Or Len(nroCheque) > 0 Or Len(vto) > 0 Or Len(importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If

                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$("Nro: " & nroCheque & "  Banco: " & banco & "  Vto: " & vto, 85)

                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(importe))

                y = y + lineH
            End If
        Next i
    End If

    y = y + lineH / 2
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Cheques:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubCheques.text))
    y = y + lineH * 2

    ' ============================================================
    ' NEW: TRANSFERENCIA
    ' ============================================================
    If grdTransferencia.Rows > 1 And FilaTieneDatos(grdTransferencia, 1) Then
        If y > pageBottom Then
            Printer.NewPage
            y = 600
        End If

        Printer.FontBold = True
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "TRANSFERENCIA"
        Printer.FontBold = False
        y = y + lineH

        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print String$(90, "-")
        y = y + lineH

        transfRows = grdTransferencia.Rows - 1
        For i = 1 To transfRows
            banco = Trim$(grdTransferencia.TextMatrix(i, colTransfBanco))
            nroCuenta = Trim$(grdTransferencia.TextMatrix(i, colTransfNroCuenta))
            cuit = Trim$(grdTransferencia.TextMatrix(i, colTransfCUIT))
            importe = Trim$(grdTransferencia.TextMatrix(i, colTransfImporte))

            If Len(banco) > 0 Or Len(nroCuenta) > 0 Or Len(cuit) > 0 Or Len(importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If

                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$("Banco: " & banco & "  Cuenta: " & nroCuenta & "  CUIT: " & cuit, 85)

                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(importe))

                y = y + lineH
            End If
        Next i

        y = y + lineH / 2
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Subtotal Transferencia:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubTransferencia.text))
        y = y + lineH * 2
    End If

    ' ============================================================
    ' NEW: FACTURAS
    ' ============================================================
    If grdFacturas.Rows > 1 And FilaTieneDatos(grdFacturas, 1) Then
        If y > pageBottom Then
            Printer.NewPage
            y = 600
        End If

        Printer.FontBold = True
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "FACTURAS"
        Printer.FontBold = False
        y = y + lineH

        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print String$(90, "-")
        y = y + lineH

        factRows = grdFacturas.Rows - 1
        For i = 1 To factRows
            nroFact = Trim$(grdFacturas.TextMatrix(i, colFactNumero))
            fechaFact = Trim$(grdFacturas.TextMatrix(i, colFactFecha))
            importe = Trim$(grdFacturas.TextMatrix(i, colFactImporte))
            descFact = Trim$(grdFacturas.TextMatrix(i, colFactDescripcion))

            If Len(nroFact) > 0 Or Len(fechaFact) > 0 Or Len(importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If

                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Dim factLine As String
                factLine = "FC " & nroFact & "  Fecha: " & fechaFact
                If Len(descFact) > 0 Then
                    factLine = factLine & "  (" & descFact & ")"
                End If
                Printer.Print Left$(factLine, 85)

                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(importe))

                y = y + lineH
            End If
        Next i

        y = y + lineH / 2
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Subtotal Facturas:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubFacturas.text))
        y = y + lineH * 2
    End If

    ' ============================================================
    ' OTROS (retencion / efectivo / subtotal otros)
    ' ============================================================
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "OTROS"
    Printer.FontBold = False
    y = y + lineH

    retText = Trim$(GetRetIIBBText())
    If Len(retText) = 0 Then retText = "0"

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Retencion IIBB:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(retText))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Efectivo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtEfectivo.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Otros:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubOtros.text))
    y = y + lineH * 2

    ' ============================================================
    ' NEW: RESUMEN (detailed breakdown)
    ' ============================================================
    If y > pageBottom - (lineH * 12) Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "RESUMEN"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    ' Breakdown of all components
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Deuda:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Cheques:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubCheques.text))
    y = y + lineH

    ' NEW: Only print if has data
    If grdTransferencia.Rows > 1 And FilaTieneDatos(grdTransferencia, 1) Then
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Transferencia:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubTransferencia.text))
        y = y + lineH
    End If

    ' NEW: Only print if has data
    If grdFacturas.Rows > 1 And FilaTieneDatos(grdFacturas, 1) Then
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Facturas:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubFacturas.text))
        y = y + lineH
    End If

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Otros:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubOtros.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Efectivo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtEfectivo.text))
    y = y + lineH

    Dim retIIBBVal As Currency
    retIIBBVal = ParseCurrency(GetRetIIBBText())
    If retIIBBVal <> 0 Then
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Retencion IIBB:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(retIIBBVal)
        y = y + lineH
    End If

    y = y + lineH / 2
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    ' Grand totals
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "TOTAL PAGO:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtTotalPago.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "SALDO:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSaldo.text))
    Printer.FontBold = False
    y = y + lineH * 2

    ' Importe en letras
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Importe en letras:"
    y = y + lineH
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print Trim$(txtImporteLetras.text)
    y = y + lineH * 4

    ' ============================================================
    ' Firma
    ' ============================================================
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(35, "_")
    y = y + lineH
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Firma / Aclaracion"

    Printer.EndDoc
    Exit Sub

ErrHandler:
    MsgBox "Error al imprimir la orden de pago: " & Err.Description, vbExclamation, "Impresion"
End Sub
```

---

## 3. Database Migration Strategy

### Migration SQL (Automatic)

The `EnsureSchemaOrdenPago()` function handles migration automatically:

1. **New Database**: Creates table with all columns (including new ones)
2. **Existing Database**: Adds missing columns via `ALTER TABLE`

### Testing Migration

**Test cases**:
1. Fresh install (no OrdenPago table) → should create table with all columns
2. Existing database (table exists, missing new columns) → should add columns
3. Already migrated database → should skip (no errors)

**Verification query** (run in Access after migration):

```sql
SELECT * FROM MSysObjects WHERE Type=1 AND Name='OrdenPago'
```

Then check field list to confirm `DetalleTransferencia` and `DetalleFacturas` exist.

---

## 4. UI Controls & Layout

### New Controls Added to FormOrdenPago.frm

| Control Name | Type | Parent Frame | Purpose |
|--------------|------|--------------|---------|
| `grdTransferencia` | MSFlexGrid | Frame8 | Bank transfer details (4 cols) |
| `grdFacturas` | MSFlexGrid | Frame7 | Invoice cross-reference (4 cols) |
| `txtSubTransferencia` | TextBox | Frame8 | Subtotal for transfers |
| `txtSubFacturas` | TextBox | Frame7 | Subtotal for invoices |
| `lblSubTransferencia` | Label | Frame8 | "Sub Total:" label |
| `lblSubFacturas` | Label | Frame7 | "Sub Total:" label |
| `cmdAddTransferencia` | CommandButton | Frame8 | Add transfer row |
| `cmdDelTransferencia` | CommandButton | Frame8 | Delete transfer row |
| `cmdAddFacturas` | CommandButton | Frame7 | Add invoice row |
| `cmdDelFacturas` | CommandButton | Frame7 | Delete invoice row |
| `txtGridEditor` | TextBox | Form (direct) | Invisible overlay for inline editing |

### Tab Order Specification

**Logical flow**: Left→Right, Top→Bottom

| TabIndex | Control | Notes |
|----------|---------|-------|
| 0-4 | Header (Frame2, txtNroOrden, txtFecha, txtProveedor) | Top section |
| 5-8 | Deuda (Frame3, grdDeuda, buttons) | Left column, row 1 |
| 9-12 | Facturas (Frame7, grdFacturas, buttons) | Right column, row 1 |
| 13-17 | Cheques (Frame4, grdCheques, buttons) | Left column, row 2 |
| 18-21 | Transferencia (Frame8, grdTransferencia, buttons) | Right column, row 2 |
| 22-25 | Otros (fraOtros, grdOtros, buttons) | Left column, row 3 |
| 26-30 | Totales (Frame5, txtEfectivo, buttons) | Bottom section |
| 31-33 | Reimpresión (Frame6, cboBuscarOrden, cmdCargarOrden) | Right column, row 3 |

---

## 5. Inline Editing Architecture

### Component Overview

```
┌─────────────────────────────────────────────────┐
│ MSFlexGrid (e.g., grdDeuda)                     │
│  ┌─────────────────────────────────┐            │
│  │ DblClick or Enter key           │            │
│  └──────────────┬──────────────────┘            │
│                 │                                │
│                 ▼                                │
│  ┌──────────────────────────────────────────┐   │
│  │ EditarCeldaGrid_Inline()                 │   │
│  │ - Calculate cell position                │   │
│  │ - Show txtGridEditor overlay             │   │
│  │ - Store edit context (grid, row, col)    │   │
│  └──────────────┬───────────────────────────┘   │
│                 │                                │
│                 ▼                                │
│  ┌──────────────────────────────────────────┐   │
│  │ txtGridEditor (visible, overlaid)        │   │
│  │ - User types                             │   │
│  │ - Enter → FinalizarEdicion(True)         │   │
│  │ - Esc → FinalizarEdicion(False)          │   │
│  │ - LostFocus → FinalizarEdicion(True)     │   │
│  └──────────────┬───────────────────────────┘   │
│                 │                                │
│                 ▼                                │
│  ┌──────────────────────────────────────────┐   │
│  │ FinalizarEdicion(save As Boolean)        │   │
│  │ if save:                                 │   │
│  │   - ValidarYFormatearCelda()             │   │
│  │   - Save to grid.TextMatrix              │   │
│  │   - RecalcularTodo()                     │   │
│  │ - Hide txtGridEditor                     │   │
│  └──────────────────────────────────────────┘   │
└─────────────────────────────────────────────────┘
```

### Key Design Decisions

1. **Single txtGridEditor for all grids**: Reduces memory footprint; simpler state management
2. **Store edit context**: `m_EditGrid`, `m_EditRow`, `m_EditCol` track which cell is being edited
3. **Position calculation**: `cellLeft = g.Left + g.Parent.Left + g.ColPos(c)` accounts for Frame offset
4. **Validation on save**: `ValidarYFormatearCelda()` applies column-specific rules before accepting input
5. **Enter/Esc/LostFocus**: All three trigger `FinalizarEdicion()` with appropriate save flag

---

## 6. Validation Framework

### Validation Points

| Validation Type | Trigger | Function |
|-----------------|---------|----------|
| **Date format (txtFecha)** | LostFocus | `ValidarFechaDD_MM_YYYY()` |
| **Date range (txtFecha)** | LostFocus | Year >= 1900, not future |
| **CUIT format** | Grid cell edit | `ValidarCUIT()` (XX-XXXXXXXX-X) |
| **Invoice number format** | Grid cell edit | `ValidarNumeroFactura()` (XXXX-XXXXXXXX) |
| **Currency format** | Grid cell edit | `ParseCurrency()` (existing) |
| **Required fields (save)** | cmdGuardar_Click | NroOrden, Fecha, Proveedor non-empty |
| **At least one payment** | cmdGuardar_Click | Any grid has data OR efectivo > 0 |

### Validation Rules Summary

| Field | Format | Range | Required | Notes |
|-------|--------|-------|----------|-------|
| txtFecha | DD/MM/YYYY | 1900 ≤ year ≤ today | Yes | Auto-fill if empty |
| CUIT (grid) | XX-XXXXXXXX-X | All digits | Yes (if row has data) | Argentine tax ID |
| Invoice # (grid) | XXXX-XXXXXXXX | All digits | Yes (if row has data) | Point of sale + number |
| Importe (grid) | #,##0.00 | >= 0 | Yes (if row has data) | Currency with thousands separator |
| txtNroOrden | Numeric | > 0 | Yes | Auto-generated |
| txtProveedor | Text | Max 150 chars | Yes | |

---

## 7. Calculation Engine Updates

### Updated Formula

**Before** (3 grids):
```
totalPago = subCheques + subOtros + efectivo + retIIBB
saldo = subDeuda - totalPago
```

**After** (5 grids):
```
totalPago = subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB
saldo = subDeuda - totalPago
```

### Affected Functions

1. **`RecalcularTodo()`**: Add 2 new `SumarColumna()` calls
2. **`SumarColumna()`**: No changes (already generic)
3. **UI Update**: Add `txtSubTransferencia` and `txtSubFacturas` to display logic

### Trigger Points for RecalcularTodo

| Event | Reason |
|-------|--------|
| Any grid cell edit (inline editor save) | Importe changed |
| cmdAddTransferencia_Click | New row might have data |
| cmdDelTransferencia_Click | Row deleted |
| cmdAddFacturas_Click | New row might have data |
| cmdDelFacturas_Click | Row deleted |
| txtEfectivo_Change | Manual efectivo entry |
| txtRetIIBB_Change | Manual retention entry |
| Form_Load | Initial state |
| CargarOrdenPorNumero | Loading existing order |

---

## 8. Serialization Integration

### Existing Implementation (No Changes)

The existing `SerializarGrid()` and `DeserializarGrid()` functions are **already grid-agnostic**:

```vb
' Generic serialization - works for ANY MSFlexGrid
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    ' Iterate g.Rows, g.Cols - no hardcoded references
    ' Uses EscaparDato() for special characters
End Function

Private Sub DeserializarGrid(ByRef g As MSFlexGrid, ByVal dataText As String)
    ' Parse dataText and populate g.TextMatrix
    ' Uses DesescaparDato() for special characters
End Sub
```

### Usage Pattern for New Grids

**Save**:
```vb
rs!DetalleTransferencia = SerializarGrid(grdTransferencia)
rs!DetalleFacturas = SerializarGrid(grdFacturas)
```

**Load**:
```vb
DeserializarGrid grdTransferencia, NzS(rs!DetalleTransferencia)
DeserializarGrid grdFacturas, NzS(rs!DetalleFacturas)
```

**Clear**:
```vb
ResetGridRows grdTransferencia
ResetGridRows grdFacturas
```

### Character Escaping

| Character | Escaped As | Reason |
|-----------|------------|--------|
| `\` | `\\` | Escape character |
| `|` | `\p` | Column separator |
| `vbCr` | (removed) | Prevent line break corruption |
| `vbLf` | `\n` | Row separator conflict |

---

## 9. Printing Enhancement

### Conditional Logo Logic

```
┌────────────────────────────────────────┐
│ Start Print                            │
└────────────┬───────────────────────────┘
             │
             ▼
      ┌──────────────┐
      │ TieneFacturas()? │
      └──────┬───────────┘
             │
       ┌─────┴─────┐
       │           │
      Yes         No
       │           │
       ▼           ▼
   ┌─────────┐  ┌──────────────┐
   │ Logo    │  │ Text Header  │
   │ exists? │  │ "ORDEN DE    │
   └────┬────┘  │  PAGO"       │
        │       └──────────────┘
    ┌───┴───┐
   Yes     No
    │       │
    ▼       ▼
 ┌──────┐ ┌──────────────┐
 │ Load │ │ Fallback     │
 │ Logo │ │ Text Header  │
 └──────┘ └──────────────┘
```

### New Printing Sections

| Section | Condition | Position |
|---------|-----------|----------|
| DEUDA | Always | After header |
| CHEQUES | Always | After DEUDA |
| TRANSFERENCIA | If `grdTransferencia` has data | After CHEQUES |
| FACTURAS | If `grdFacturas` has data | After TRANSFERENCIA |
| OTROS | Always | After FACTURAS |
| RESUMEN | Always | After OTROS |
| Firma | Always | End |

### Print Settings

| Setting | Value | Notes |
|---------|-------|-------|
| Copies | 2 | `Printer.Copies = 2` |
| Font | Courier New 10pt | Monospaced for alignment |
| Page Size | Letter (8.5" x 11") | Standard |
| Orientation | Portrait | Default |

---

## 10. Testing Strategy

### Unit Testing (Manual in VB6 IDE)

| Function | Test Case | Expected Result |
|----------|-----------|-----------------|
| `ValidarCUIT()` | "20-12345678-9" | True |
| `ValidarCUIT()` | "20-123456789" (too short) | False |
| `ValidarCUIT()` | "20-1234567A-9" (letter) | False |
| `ValidarNumeroFactura()` | "0001-00012345" | True |
| `ValidarNumeroFactura()` | "1-12345" (too short) | False |
| `ParseDateDD_MM_YYYY()` | "31/12/2025" | Valid Date |
| `ParseDateDD_MM_YYYY()` | "31/02/2025" (invalid) | Error (caught by caller) |
| `SumarColumna()` | Grid with 3 rows: "100", "200.50", "0" | 300.50 |
| `RecalcularTodo()` | All 5 grids populated | Correct totalPago and saldo |

### Integration Testing (Full Form)

#### Test 1: Create New Order with All Grids

**Steps**:
1. Open FormOrdenPago
2. Enter Proveedor: "Test Supplier"
3. Add 2 rows to grdDeuda (Concepto: "Debt 1", Importe: "1000")
4. Add 2 rows to grdCheques (complete all 4 columns)
5. Add 2 rows to grdTransferencia (Banco, Cuenta, CUIT, Importe)
6. Add 2 rows to grdFacturas (Numero, Fecha, Importe, Descripcion)
7. Add 1 row to grdOtros
8. Enter Efectivo: "500"
9. Click Recalcular
10. Verify all subtotals correct
11. Verify totalPago = sum of all
12. Verify saldo = subDeuda - totalPago
13. Click Guardar
14. Verify success message

**Expected**: All data saved correctly, no errors.

---

#### Test 2: Inline Editing

**Steps**:
1. Open existing order (or create new)
2. Double-click cell in grdDeuda (Concepto column)
3. Verify txtGridEditor appears overlaying cell
4. Type new value, press Enter
5. Verify cell updated, editor hidden
6. Double-click cell in grdTransferencia (CUIT column)
7. Type invalid CUIT (e.g., "123")
8. Press Enter
9. Verify error message, editor still visible
10. Type valid CUIT "20-12345678-9", press Enter
11. Verify cell updated, editor hidden
12. Double-click cell, type value, press Esc
13. Verify cell NOT updated (cancel)

**Expected**: Inline editing works, validation prevents invalid data, Esc cancels.

---

#### Test 3: Date Validation

**Steps**:
1. Open FormOrdenPago
2. Click txtFecha, clear it, tab away
3. Verify auto-filled with today's date
4. Enter "99/99/9999", tab away
5. Verify error message, reverted to previous value
6. Enter "31/12/2030" (future), tab away
7. Verify error message "fecha no puede ser futura"
8. Enter "31/12/1899" (ancient), tab away
9. Verify error message "fecha muy antigua"
10. Enter "25/12/2025", tab away
11. Verify accepted, formatted correctly

**Expected**: Invalid dates rejected, valid dates accepted and normalized.

---

#### Test 4: Save & Load

**Steps**:
1. Create order with all 5 grids populated
2. Click Guardar
3. Note NroOrden value
4. Click Restaurar (clear form)
5. Select order from cboBuscarOrden
6. Click Cargar Orden
7. Verify all 5 grids restored correctly
8. Verify subtotals match
9. Open DB_SPC_SI.mdb in Access
10. Query: `SELECT DetalleTransferencia, DetalleFacturas FROM OrdenPago WHERE NroOrden=X`
11. Verify MEMO fields contain serialized data (pipe-separated)

**Expected**: Round-trip save/load preserves all data.

---

#### Test 5: Printing with Logo

**Steps**:
1. Create order WITHOUT facturas (grdFacturas empty)
2. Click Imprimir
3. Verify text header "ORDEN DE PAGO" (no logo)
4. Add 1 row to grdFacturas
5. Click Imprimir
6. If Quilplac2.jpg exists: Verify logo printed
7. If logo missing: Verify text header fallback
8. Check printed output for:
   - DEUDA section
   - CHEQUES section
   - TRANSFERENCIA section (if data)
   - FACTURAS section
   - OTROS section
   - RESUMEN with breakdown
9. Verify 2 copies printed (check printer queue)

**Expected**: Logo conditional on facturas presence, all sections printed, 2 copies.

---

#### Test 6: Edge Cases

**Steps**:
1. Create order with empty grids (only Efectivo > 0)
2. Click Guardar
3. Verify allowed (at least one payment method present)
4. Create order with special characters in Concepto: "Test | pipe \\ backslash"
5. Click Guardar, then reload
6. Verify special characters preserved correctly
7. Create order with 100 rows in grdDeuda (stress test)
8. Click Recalcular
9. Verify performance acceptable (<1 second)
10. Click Guardar
11. Verify MEMO field size acceptable

**Expected**: Edge cases handled gracefully, no crashes.

---

### Regression Testing

After implementation, test these existing features to ensure no breakage:

| Feature | Test |
|---------|------|
| Existing 3 grids (Deuda, Cheques, Otros) | CRUD operations still work |
| Existing serialization | Old orders load correctly |
| Existing printing | Prints without errors (even without new grids) |
| Existing calculation | totalPago/saldo correct for old orders |
| Tab order (old controls) | Still navigable |

---

## 11. Rollback Plan

### Rollback Triggers

Rollback if:
1. Database migration fails on existing installations
2. Inline editing causes data loss or corruption
3. Printing crashes or produces unreadable output
4. Performance degrades significantly (>2 second form load)

### Rollback Strategy

#### Database Rollback

**Problem**: New columns cause schema errors

**Solution**:
1. New columns are optional (NULL allowed)
2. Old code ignores new columns via `NzS()` wrapper
3. No DROP COLUMN needed (keeps data)

**If must revert**:
```sql
-- Manual SQL in Access (NOT automated)
ALTER TABLE OrdenPago DROP COLUMN DetalleTransferencia
ALTER TABLE OrdenPago DROP COLUMN DetalleFacturas
```

---

#### Code Rollback

**Problem**: Inline editing breaks existing functionality

**Solution**:
1. Keep old `EditarCeldaGrid()` (InputBox version) as fallback
2. Comment out new inline editing event handlers
3. Restore old DblClick handlers to call `EditarCeldaGrid()` instead of `EditarCeldaGrid_Inline()`
4. Remove `txtGridEditor` from form

**Git strategy**:
```bash
# Rollback to before Phase 7 (inline editing)
git revert <commit-hash-phase-7>
git push
```

---

#### Printing Rollback

**Problem**: New printing logic causes errors

**Solution**:
1. Restore old `cmdImprimir_Click()` from git history
2. New sections (TRANSFERENCIA, FACTURAS) simply won't print
3. Old sections (DEUDA, CHEQUES, OTROS) still work

---

### Phased Deployment (Minimize Rollback Risk)

**Recommended approach**:

1. **Phase 1-6**: Deploy to test environment, test 1 week
2. **Phase 7-8**: Deploy to staging, test with real users (3 days)
3. **Phase 9-11**: Deploy to production with monitoring

**If Phase 7 fails**: Rollback just Phase 7, keep Phases 1-6 (new grids work with InputBox editing)

**If Phase 10 fails**: Rollback just Phase 10, keep Phases 1-9 (printing uses old logic)

---

## 12. Summary of Changes

### Files Modified

| File | Lines Changed | Description |
|------|---------------|-------------|
| `FormOrdenPago.frm` | ~800 lines added | New grids, inline editing, printing, validation |
| `DB_SPC_SI.mdb` | Schema only | 2 new MEMO columns |

### Code Statistics

| Category | LOC Added | LOC Modified | LOC Deleted |
|----------|-----------|--------------|-------------|
| Module-level variables | 6 | 0 | 0 |
| Enum definitions | 8 | 0 | 0 |
| InitGrids | 30 | 0 | 0 |
| Form_Load | 12 | 0 | 0 |
| RecalcularTodo | 5 | 10 | 0 |
| LimpiarFormulario | 4 | 0 | 0 |
| cmdGuardar_Click | 4 | 2 | 0 |
| CargarOrdenPorNumero | 2 | 0 | 0 |
| EnsureSchemaOrdenPago | 15 | 10 | 5 |
| Inline editing system | 150 | 0 | 0 |
| Validation functions | 120 | 0 | 0 |
| Event handlers (new grids) | 40 | 0 | 0 |
| Printing (cmdImprimir_Click) | 200 | 50 | 20 |
| **TOTAL** | **~600** | **~72** | **~25** |

### New Features Summary

| Feature | Benefit |
|---------|---------|
| **Transferencia grid** | Track bank transfer payments with CUIT validation |
| **Facturas grid** | Cross-reference invoices with invoice number validation |
| **Inline editing** | Professional UX, replaces clunky InputBox |
| **Date validation** | Prevent future dates, enforce DD/MM/YYYY format |
| **Enhanced printing** | Conditional logo, 2 copies, detailed RESUMEN section |
| **Calculation engine** | Support 5 payment methods instead of 3 |
| **Tab order** | Logical keyboard navigation (left→right, top→bottom) |

---

## Appendix A: Dependency Matrix

| Phase | Depends On | Blocks |
|-------|------------|--------|
| Phase 1 (DB Schema) | None | Phases 2-11 |
| Phase 2 (UI Controls) | Phase 1 | Phases 3-11 |
| Phase 3 (InitGrids) | Phase 2 | Phases 4-11 |
| Phase 4 (Serialization) | Phase 3 | Phases 5-11 |
| Phase 5 (Calculation) | Phase 4 | Phase 6 |
| Phase 6 (Load/Save) | Phase 5 | Phase 7 |
| Phase 7 (Inline Edit) | Phase 6 | None (independent) |
| Phase 8 (Date Valid) | None | None (independent) |
| Phase 9 (Tab Order) | Phase 2 | None (independent) |
| Phase 10 (Printing) | Phase 6 | None (independent) |
| Phase 11 (Cleanup) | Phases 7, 10 | None |

---

## Appendix B: Risk Assessment

| Risk | Probability | Impact | Mitigation |
|------|-------------|--------|------------|
| Database migration fails | Low | High | Use ALTER TABLE with On Error Resume Next; test on backup DB first |
| Inline editing causes data loss | Medium | High | Keep old EditarCeldaGrid as fallback; extensive testing before deployment |
| Printing crashes | Low | Medium | Wrap in On Error GoTo; test with/without logo, with/without data |
| Performance degrades | Low | Medium | RecalcularTodo is O(n) with small n; profile if >100 rows per grid |
| Special characters corrupt serialization | Medium | High | Existing EscaparDato/DesescaparDato handles this; test edge cases |
| Tab order confuses users | Low | Low | Follow standard left→right flow; gather user feedback |
| CUIT/Invoice validation too strict | Medium | Medium | Allow empty fields (optional); validate only if user enters data |

---

**End of Design Document**

---

## Next Steps

1. **Review**: Share this design with project stakeholders for approval
2. **Environment Setup**: Backup DB_SPC_SI.mdb, create test database
3. **Implementation**: Follow phased approach (Phases 1-11)
4. **Testing**: Execute test plan after each phase
5. **Deployment**: Stage → Production with monitoring
6. **Documentation**: Update user manual with new features

**Estimated Total Implementation Time**: 8-10 hours (single developer)

**Recommended Review Date**: 2026-05-09 (2 days from now)
