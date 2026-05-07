# Technical Specification: FormOrdenPago.frm Improvements

**Project**: SPC-Core
**Change**: FormOrdenPago-improvements
**Version**: 1.0
**Date**: 2026-05-07

---

## Table of Contents

1. [Overview](#overview)
2. [Database Schema](#database-schema)
3. [Grid Specifications](#grid-specifications)
4. [UI Layout & Tab Order](#ui-layout--tab-order)
5. [Inline Editing Specification](#inline-editing-specification)
6. [Date Validation Specification](#date-validation-specification)
7. [Calculation Specification](#calculation-specification)
8. [Serialization Specification](#serialization-specification)
9. [Printing Specification](#printing-specification)
10. [Serialization & Deserialization Code Patterns](#serialization--deserialization-code-patterns)
11. [Validation & Error Handling](#validation--error-handling)

---

## 1. Overview

This specification details improvements to `FormOrdenPago.frm`, adding two new grid sections (Transferencia and Facturas) to support comprehensive payment order tracking per client requirements (Pedido Pato.pdf).

**Key Changes**:
- Add `grdTransferencia` grid with 4 columns (Banco, Numero Cuenta, CUIT, Importe)
- Add `grdFacturas` grid with 4 columns (Numero, Fecha, Importe, Descripcion)
- Extend database schema with 2 new MEMO fields
- Update calculation logic to include new grids
- Update printing to include new sections with conditional logo display
- Implement inline editing for all grids (replace InputBox pattern)
- Add robust date validation

**Requirements Source**: Pedido Pato.pdf - Section 3: "Módulo de Orden de Pago"
- "Otros (ej: cruce de facturas tipo 'FC 8383 $12.x65')" → grdFacturas
- "Pago con Cheque (monto, datos a mano)" → existing grdCheques
- Additional payment methods → grdTransferencia

---

## 2. Database Schema

### 2.1 Existing Table: OrdenPago

Current schema (from EnsureSchemaOrdenPago, line 2069-2086):

```sql
CREATE TABLE OrdenPago (
    NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY,
    Fecha DATETIME,
    Proveedor TEXT(150),
    SubDeuda CURRENCY,
    SubCheques CURRENCY,
    SubOtros CURRENCY,
    Efectivo CURRENCY,
    RetIIBB CURRENCY,
    TotalPago CURRENCY,
    Saldo CURRENCY,
    ImporteLetras MEMO,
    DetalleDeuda MEMO,
    DetalleCheques MEMO,
    DetalleOtros MEMO,
    FechaAlta DATETIME
)
```

### 2.2 Schema Migration

Add 2 new columns to support new grids:

```sql
ALTER TABLE OrdenPago
ADD COLUMN DetalleTransferencia MEMO;

ALTER TABLE OrdenPago
ADD COLUMN DetalleFacturas MEMO;
```

**Implementation Notes**:
- Execute in `EnsureSchemaOrdenPago` after table creation check
- Use error handling to skip if columns already exist
- MEMO fields support unlimited text for serialized grid data
- NULL handling: Use `NzS()` wrapper when reading (existing pattern)

**Migration Code Pattern**:
```vb
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        ' Create full table (existing code + new columns)
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
        ' Migration: add columns if missing
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

### 2.3 Backwards Compatibility

- Old records: DetalleTransferencia, DetalleFacturas will be NULL
- Reading: Use `NzS(rs!DetalleTransferencia, "")` → empty string → no grid rows
- Writing: Always save both fields (even if empty string)
- No data loss: Existing DetalleDeuda, DetalleCheques, DetalleOtros unchanged

---

## 3. Grid Specifications

### 3.1 grdDeuda (Existing - Document Current State)

**Purpose**: List debt items (concepts and amounts)

| Column | Name | Type | Max Length | Required | Validation |
|--------|------|------|------------|----------|------------|
| 0 | Concepto | TEXT | 100 chars | Yes | Non-empty before save |
| 1 | Importe | CURRENCY | - | Yes | >= 0, parsed via ParseCurrency |

**UI Properties**:
- Rows: 2 (1 fixed header + 1 data row minimum)
- Cols: 2
- FixedRows: 1
- FixedCols: 0
- ColWidth(0): 3000 twips (~4.17 cm)
- ColWidth(1): 1500 twips (~2.08 cm)
- Font: MS Sans Serif 8.25pt (default form font)
- Height: 1575 twips
- Width: 5775 twips

**Header Labels**:
- Row 0, Col 0: "Concepto"
- Row 0, Col 1: "Importe"

**Validation Rules**:
- Before save: At least one row must have Concepto non-empty OR Importe > 0
- Importe column: Must parse to valid currency (handled by ParseCurrency, line 2112-2189)

**Serialization Format**: Pipe-separated columns, CRLF-separated rows (SEP_COL = "|", SEP_ROW = vbCrLf)

---

### 3.2 grdCheques (Existing - Document Current State)

**Purpose**: List check payments

| Column | Name | Type | Max Length | Required | Validation |
|--------|------|------|------------|----------|------------|
| 0 | Banco | TEXT | 50 chars | Yes | Non-empty |
| 1 | Numero | TEXT | 20 chars | Yes | Non-empty |
| 2 | Fecha | DATE | 10 chars (DD/MM/YYYY) | Yes | Valid date, not future |
| 3 | Importe | CURRENCY | - | Yes | >= 0 |

**UI Properties**:
- Rows: 2
- Cols: 4
- FixedRows: 1
- FixedCols: 0
- ColWidth(0): 2200 twips
- ColWidth(1): 1600 twips
- ColWidth(2): 1300 twips
- ColWidth(3): 1400 twips
- Font: MS Sans Serif 8.25pt
- Height: 1575 twips
- Width: 5775 twips

**Header Labels**:
- Row 0, Col 0: "Banco"
- Row 0, Col 1: "Numero"
- Row 0, Col 2: "Fecha"
- Row 0, Col 3: "Importe"

**Validation Rules**:
- All fields required if row has any data
- Fecha: DD/MM/YYYY format, not future date
- Importe: >= 0

**Serialization Format**: Same pipe/CRLF pattern

---

### 3.3 grdOtros (Existing - Document Current State)

**Purpose**: Other payment concepts

| Column | Name | Type | Max Length | Required | Validation |
|--------|------|------|------------|----------|------------|
| 0 | Concepto | TEXT | 100 chars | Yes | Non-empty |
| 1 | Importe | CURRENCY | - | Yes | >= 0 |

**UI Properties**:
- Rows: 2
- Cols: 2
- FixedRows: 1
- FixedCols: 0
- ColWidth(0): 3000 twips
- ColWidth(1): 1500 twips
- Font: MS Sans Serif 8.25pt
- Height: 1455 twips
- Width: 5775 twips

**Header Labels**:
- Row 0, Col 0: "Concepto"
- Row 0, Col 1: "Importe"

**Validation Rules**: Same as grdDeuda

**Serialization Format**: Same pipe/CRLF pattern

---

### 3.4 grdTransferencia (NEW)

**Purpose**: Bank transfer payment details

| Column | Name | Type | Max Length | Required | Validation |
|--------|------|------|------------|----------|------------|
| 0 | Banco | TEXT | 50 chars | Yes | Non-empty |
| 1 | Numero Cuenta | TEXT | 20 chars | Yes | Non-empty |
| 2 | CUIT | TEXT | 13 chars | Yes | Format XX-XXXXXXXX-X (11 digits + 2 hyphens) |
| 3 | Importe | CURRENCY | - | Yes | >= 0 |

**UI Properties**:
- Rows: 2
- Cols: 4
- FixedRows: 1
- FixedCols: 0
- ColWidth(0): 1800 twips (~2.5 cm, narrower than Cheques for CUIT space)
- ColWidth(1): 1600 twips
- ColWidth(2): 1500 twips (CUIT needs ~13 chars width)
- ColWidth(3): 1400 twips
- Font: MS Sans Serif 8.25pt
- Height: 1455 twips
- Width: 5775 twips
- Positioned in Frame8 (existing, lines 79-190)

**Header Labels**:
- Row 0, Col 0: "Banco"
- Row 0, Col 1: "Nro. Cuenta"
- Row 0, Col 2: "CUIT"
- Row 0, Col 3: "Importe"

**Validation Rules**:
- Banco: Required, non-empty, max 50 chars
- Numero Cuenta: Required, non-empty, max 20 chars
- CUIT: Required, format validation XX-XXXXXXXX-X (where X = digit)
  - Regex: `^\d{2}-\d{8}-\d$`
  - VB6 implementation: Check length = 13, positions 3 and 12 are "-", rest are digits
- Importe: Required, >= 0, parsed via ParseCurrency

**CUIT Validation Function**:
```vb
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
```

**Serialization Format**: Same pipe/CRLF pattern

**Enum Definition**:
```vb
Private Enum eColTransferencia
    colTransfBanco = 0
    colTransfNroCuenta = 1
    colTransfCUIT = 2
    colTransfImporte = 3
End Enum
```

**InitGrids Addition**:
```vb
' Transferencia
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
```

---

### 3.5 grdFacturas (NEW)

**Purpose**: Invoice cross-reference for payment ("cruce de facturas tipo FC 8383 $12.x65")

| Column | Name | Type | Max Length | Required | Validation |
|--------|------|------|------------|----------|------------|
| 0 | Numero | TEXT | 15 chars | Yes | Format XXXX-XXXXXXXX (point of sale - number) |
| 1 | Fecha | DATE | 10 chars (DD/MM/YYYY) | Yes | Valid date, not future |
| 2 | Importe | CURRENCY | - | Yes | >= 0 |
| 3 | Descripcion | TEXT | 100 chars | No | Optional notes |

**UI Properties**:
- Rows: 2
- Cols: 4
- FixedRows: 1
- FixedCols: 0
- ColWidth(0): 1500 twips (invoice number format ~12 chars)
- ColWidth(1): 1200 twips (date 10 chars)
- ColWidth(2): 1400 twips (importe)
- ColWidth(3): 1675 twips (remaining space for descripcion)
- Font: MS Sans Serif 8.25pt
- Height: 1455 twips
- Width: 5775 twips
- Positioned in Frame7 (existing, lines 191-302)

**Header Labels**:
- Row 0, Col 0: "Numero"
- Row 0, Col 1: "Fecha"
- Row 0, Col 2: "Importe"
- Row 0, Col 3: "Descripcion"

**Validation Rules**:
- Numero: Required, format XXXX-XXXXXXXX (4 digits, hyphen, 8 digits)
  - Regex: `^\d{4}-\d{8}$`
  - VB6 implementation: Length = 13, position 5 is "-", rest digits
- Fecha: Required, DD/MM/YYYY format, not future date (use ParseDateDD_MM_YYYY)
- Importe: Required, >= 0
- Descripcion: Optional, max 100 chars

**Invoice Number Validation Function**:
```vb
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

**Serialization Format**: Same pipe/CRLF pattern

**Enum Definition**:
```vb
Private Enum eColFacturas
    colFactNumero = 0
    colFactFecha = 1
    colFactImporte = 2
    colFactDescripcion = 3
End Enum
```

**InitGrids Addition**:
```vb
' Facturas
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

## 4. UI Layout & Tab Order

### 4.1 Form Dimensions

- FormOrdenPago.Width: 16710 twips (~23.2 cm)
- FormOrdenPago.Height: 11200 twips (~15.6 cm)
- FormOrdenPago.Top: 0 (set in Form_Load, line 909)

### 4.2 Frame Organization (Left to Right, Top to Bottom)

| Frame | Caption | Top | Left | Width | Height | Contains |
|-------|---------|-----|------|-------|--------|----------|
| Frame2 | (Header) | 120 | 600 | 15255 | 855 | txtNroOrden, txtFecha, txtProveedor |
| Frame3 | "Detalle de Deuda" | 1080 | 600 | 7455 | 2415 | grdDeuda, buttons |
| Frame4 | "Detalle de Cheques" | 3600 | 600 | 7455 | 2535 | grdCheques, buttons |
| fraOtros | "Detalle de Otros" | 6240 | 600 | 7455 | 2415 | grdOtros, buttons |
| Frame7 | "Detalle de Facturas" | 1080 | 8400 | 7455 | 2415 | grdFacturas, buttons |
| Frame8 | "Detalle de Transferencia" | 3720 | 8400 | 7455 | 2415 | grdTransferencia, buttons |
| Frame6 | "Reimpresión" | 6240 | 8400 | 7455 | 2415 | cboBuscarOrden, cmdCargarOrden |
| Frame5 | "Totales" | 8760 | 600 | 15375 | 1335 | Totals, buttons |

**Visual Flow**: 2x3 grid layout
```
+---------------------+---------------------+
| Header (Frame2)                           |
+---------------------+---------------------+
| Deuda (Frame3)      | Facturas (Frame7)   |
+---------------------+---------------------+
| Cheques (Frame4)    | Transferencia (F8)  |
+---------------------+---------------------+
| Otros (fraOtros)    | Reimpresión (F6)    |
+---------------------+---------------------+
| Totales (Frame5)                          |
+-------------------------------------------+
```

### 4.3 Tab Order Specification

**Goal**: Left-to-right, top-to-bottom navigation matching user mental model

| TabIndex | Control | Type | Frame |
|----------|---------|------|-------|
| 0 | Frame1 | Frame | Container |
| 1 | Frame2 | Frame | Header container |
| 2 | txtNroOrden | TextBox | Frame2 |
| 3 | txtFecha | TextBox | Frame2 |
| 4 | txtProveedor | TextBox | Frame2 |
| 5 | Frame3 | Frame | Deuda container |
| 6 | grdDeuda | MSFlexGrid | Frame3 |
| 7 | cmdAddDeuda | CommandButton | Frame3 |
| 8 | cmdDelDeuda | CommandButton | Frame3 |
| 9 | Frame7 | Frame | Facturas container |
| 10 | grdFacturas | MSFlexGrid | Frame7 |
| 11 | cmdAddFacturas | CommandButton | Frame7 |
| 12 | cmdDelFacturas | CommandButton | Frame7 |
| 13 | Frame4 | Frame | Cheques container |
| 14 | grdCheques | MSFlexGrid | Frame4 |
| 15 | cmdAddCheque | CommandButton | Frame4 |
| 16 | cmdDelCheque | CommandButton | Frame4 |
| 17 | Frame8 | Frame | Transferencia container |
| 18 | grdTransferencia | MSFlexGrid | Frame8 |
| 19 | cmdAddTransferencia | CommandButton | Frame8 |
| 20 | cmdDelTransferencia | CommandButton | Frame8 |
| 21 | fraOtros | Frame | Otros container |
| 22 | grdOtros | MSFlexGrid | fraOtros |
| 23 | cmdAddOtros | CommandButton | fraOtros |
| 24 | cmdDelOtros | CommandButton | fraOtros |
| 25 | txtEfectivo | TextBox | Frame5 |
| 26 | txtRetIIBB | TextBox | Frame5 (if exists) |
| 27 | cmdRecalcular | CommandButton | Frame5 |
| 28 | cmdRestaurarAuto | CommandButton | Frame5 |
| 29 | cmdGuardar | CommandButton | Frame5 |
| 30 | cmdImprimir | CommandButton | Frame5 |
| 31 | Frame6 | Frame | Reimpresión container |
| 32 | cboBuscarOrden | ComboBox | Frame6 |
| 33 | cmdCargarOrden | CommandButton | Frame6 |

**Implementation**: Set TabIndex property in VB6 IDE Form Designer for each control

---

## 5. Inline Editing Specification

### 5.1 Overview

Replace current InputBox editing (lines 1117-1146) with inline TextBox overlay for professional UX.

**Current Pattern (InputBox)**:
```vb
v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
```

**New Pattern (Inline Overlay)**:
- Double-click or Enter key → show TextBox overlaying grid cell
- TextBox pre-filled with current cell value
- Enter key → save and hide
- Esc key → cancel and hide
- LostFocus → save and hide

### 5.2 Module-Level Variables

Declare in form's Declarations section:

```vb
' Inline editing overlay
Private WithEvents txtGridEditor As TextBox
Private m_EditGrid As MSFlexGrid
Private m_EditRow As Integer
Private m_EditCol As Integer
Private m_EditingActive As Boolean
```

### 5.3 TextBox Creation (Form_Load)

Add to Form_Load after InitGrids:

```vb
' Create invisible inline editor
Set txtGridEditor = Me.Controls.Add("VB.TextBox", "txtGridEditor", Me)
txtGridEditor.Visible = False
txtGridEditor.BorderStyle = 1  ' Fixed Single
txtGridEditor.FontName = "MS Sans Serif"
txtGridEditor.FontSize = 8.25
```

### 5.4 Positioning Formula

When user double-clicks or presses Enter on grid cell:

```vb
Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim cellLeft As Long, cellTop As Long
    Dim cellWidth As Long, cellHeight As Long

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub  ' No editing header row

    ' Calculate cell position relative to FORM
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
    MsgBox "Error iniciando edición: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

### 5.5 Event Handlers

**txtGridEditor_KeyDown**:
```vb
Private Sub txtGridEditor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        FinalizarEdicion True  ' Save
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        FinalizarEdicion False  ' Cancel
    End If
End Sub
```

**txtGridEditor_LostFocus**:
```vb
Private Sub txtGridEditor_LostFocus()
    If m_EditingActive Then
        FinalizarEdicion True  ' Auto-save on focus loss
    End If
End Sub
```

**FinalizarEdicion** (shared save/cancel logic):
```vb
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
    MsgBox "Error finalizando edición: " & Err.Description, vbExclamation, "Orden de Pago"
    txtGridEditor.Visible = False
    m_EditingActive = False
End Sub
```

### 5.6 Validation and Formatting

**ValidarYFormatearCelda** (column-specific rules):
```vb
Private Function ValidarYFormatearCelda(ByRef g As MSFlexGrid, ByVal r As Integer, ByVal c As Integer, ByVal valorNuevo As String, ByRef valorFinal As String) As Boolean
    On Error GoTo EH

    ValidarYFormatearCelda = False
    valorFinal = valorNuevo

    ' Identify grid and column type
    Select Case True
        ' grdDeuda
        Case g.Name = "grdDeuda" And c = colDeudaImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        ' grdCheques
        Case g.Name = "grdCheques" And c = colChequeImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdCheques" And c = colChequeFecha
            If Not ValidarFechaDD_MM_YYYY(valorNuevo) Then
                MsgBox "Fecha inválida. Use formato DD/MM/YYYY", vbExclamation, "Validación"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        ' grdOtros
        Case g.Name = "grdOtros" And c = colOtroImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        ' grdTransferencia (NEW)
        Case g.Name = "grdTransferencia" And c = colTransfImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdTransferencia" And c = colTransfCUIT
            If Len(valorNuevo) > 0 And Not ValidarCUIT(valorNuevo) Then
                MsgBox "CUIT inválido. Use formato XX-XXXXXXXX-X (ej: 20-12345678-9)", vbExclamation, "Validación"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        ' grdFacturas (NEW)
        Case g.Name = "grdFacturas" And c = colFactImporte
            valorFinal = FormatMoney(ParseCurrency(valorNuevo))
            ValidarYFormatearCelda = True

        Case g.Name = "grdFacturas" And c = colFactNumero
            If Len(valorNuevo) > 0 And Not ValidarNumeroFactura(valorNuevo) Then
                MsgBox "Número de factura inválido. Use formato XXXX-XXXXXXXX (ej: 0001-00012345)", vbExclamation, "Validación"
                txtGridEditor.SelStart = 0
                txtGridEditor.SelLength = Len(txtGridEditor.Text)
                Exit Function
            End If
            ValidarYFormatearCelda = True

        Case g.Name = "grdFacturas" And c = colFactFecha
            If Len(valorNuevo) > 0 And Not ValidarFechaDD_MM_YYYY(valorNuevo) Then
                MsgBox "Fecha inválida. Use formato DD/MM/YYYY", vbExclamation, "Validación"
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
    MsgBox "Error en validación: " & Err.Description, vbExclamation, "Validación"
    ValidarYFormatearCelda = False
End Function
```

### 5.7 Grid Event Hookup

Replace existing DblClick subs (lines 1105-1115) and add KeyDown handlers:

**grdDeuda**:
```vb
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdDeuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdDeuda
    End If
End Sub
```

**grdCheques**:
```vb
Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdCheques
    End If
End Sub
```

**grdOtros**:
```vb
Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub grdOtros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdOtros
    End If
End Sub
```

**grdTransferencia** (NEW):
```vb
Private Sub grdTransferencia_DblClick()
    EditarCeldaGrid grdTransferencia
End Sub

Private Sub grdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdTransferencia
    End If
End Sub
```

**grdFacturas** (NEW):
```vb
Private Sub grdFacturas_DblClick()
    EditarCeldaGrid grdFacturas
End Sub

Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdFacturas
    End If
End Sub
```

---

## 6. Date Validation Specification

### 6.1 txtFecha Validation

**Trigger**: txtFecha_LostFocus

**Rules**:
1. If empty: Set to Today() (tolerate empty → auto-fill)
2. If invalid format: Show error, restore previous value or Today()
3. If future date: Show error, restore previous valid date
4. If year < 1900: Show error, restore previous valid date

**Implementation**:
```vb
Private m_FechaAnterior As String  ' Module-level variable

Private Sub Form_Load()
    ' ... existing code ...
    m_FechaAnterior = Format$(Date, "dd/mm/yyyy")
End Sub

Private Sub txtFecha_LostFocus()
    On Error GoTo EH

    Dim fechaStr As String
    Dim fechaVal As Date

    fechaStr = Trim$(txtFecha.Text)

    ' Empty → auto-fill today
    If Len(fechaStr) = 0 Then
        txtFecha.Text = Format$(Date, "dd/mm/yyyy")
        m_FechaAnterior = txtFecha.Text
        Exit Sub
    End If

    ' Parse and validate
    If Not ValidarFechaDD_MM_YYYY(fechaStr) Then
        MsgBox "Fecha inválida. Use formato DD/MM/YYYY", vbExclamation, "Validación"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    fechaVal = ParseDateDD_MM_YYYY(fechaStr)

    ' Check future
    If fechaVal > Date Then
        MsgBox "La fecha no puede ser futura", vbExclamation, "Validación"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    ' Check ancient
    If Year(fechaVal) < 1900 Then
        MsgBox "Fecha muy antigua (año < 1900)", vbExclamation, "Validación"
        txtFecha.Text = m_FechaAnterior
        txtFecha.SetFocus
        Exit Sub
    End If

    ' Valid → normalize format and save
    txtFecha.Text = Format$(fechaVal, "dd/mm/yyyy")
    m_FechaAnterior = txtFecha.Text

    Exit Sub
EH:
    MsgBox "Error validando fecha: " & Err.Description, vbExclamation, "Validación"
    txtFecha.Text = m_FechaAnterior
End Sub
```

### 6.2 Helper Functions

**ValidarFechaDD_MM_YYYY** (returns True if valid format and date):
```vb
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
```

**ParseDateDD_MM_YYYY** (returns Date object from DD/MM/YYYY string):
```vb
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
```

---

## 7. Calculation Specification

### 7.1 RecalcularTodo Formula (Updated)

**Current formula** (lines 1670-1704):
```vb
subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
subCheques = SumarColumna(grdCheques, colChequeImporte)
subOtros = SumarColumna(grdOtros, colOtroImporte)
efectivo = ParseCurrency(txtEfectivo.text)
retIIBB = ParseCurrency(GetRetIIBBText())
totalPago = subCheques + subOtros + efectivo + retIIBB
saldo = subDeuda - totalPago
```

**NEW formula** (add new grids):
```vb
subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
subCheques = SumarColumna(grdCheques, colChequeImporte)
subOtros = SumarColumna(grdOtros, colOtroImporte)
subTransferencia = SumarColumna(grdTransferencia, colTransfImporte)  ' NEW
subFacturas = SumarColumna(grdFacturas, colFactImporte)  ' NEW
efectivo = ParseCurrency(txtEfectivo.text)
retIIBB = ParseCurrency(GetRetIIBBText())

totalPago = subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB
saldo = subDeuda - totalPago
```

### 7.2 Updated RecalcularTodo Implementation

```vb
Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim subTransferencia As Currency  ' NEW
    Dim subFacturas As Currency  ' NEW
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)
    subTransferencia = SumarColumna(grdTransferencia, colTransfImporte)  ' NEW
    subFacturas = SumarColumna(grdFacturas, colFactImporte)  ' NEW

    efectivo = ParseCurrency(txtEfectivo.text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    totalPago = subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB
    saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.text = FormatMoney(subDeuda)
    txtSubCheques.text = FormatMoney(subCheques)
    txtSubOtros.text = FormatMoney(subOtros)
    txtSubTransferencia.text = FormatMoney(subTransferencia)  ' NEW
    txtSubFacturas.text = FormatMoney(subFacturas)  ' NEW
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

### 7.3 UI Controls for New Subtotals

**Add to Frame7 (Facturas)** (already exist in .frm, lines 266-301):
- txtSubFacturas (TextBox, ReadOnly, shows subtotal)
- lblSubFacturas (Label, "Sub Total:")

**Add to Frame8 (Transferencia)** (already exist in .frm, lines 154-189):
- txtSubTransferencia (TextBox, ReadOnly, shows subtotal)
- lblSubTransferencia (Label, "Sub Total:")

**LostFocus handlers** (format on exit):
```vb
Private Sub txtSubTransferencia_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubTransferencia
End Sub

Private Sub txtSubFacturas_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubFacturas
End Sub
```

### 7.4 Trigger Points

RecalcularTodo must be called after:
1. Any grid cell edit (inline editor save)
2. Any row add/delete
3. txtEfectivo.Change
4. txtRetIIBB.Change
5. Form load (initial state)
6. Cargar orden from database

---

## 8. Serialization Specification

### 8.1 Format

**Existing pattern** (lines 1978-2058):
- Column separator: `SEP_COL = "|"`
- Row separator: `SEP_ROW = vbCrLf`
- Escape function: `EscaparDato()` replaces:
  - `\` → `\\`
  - `|` → `\p`
  - `vbCr` → (removed)
  - `vbLf` → `\n`
- Unescape function: `DesescaparDato()` reverses

**Example serialized grid** (3 rows, 2 columns):
```
Concepto 1|1234,56
Concepto 2|5678,90
Concepto con \p pipe|999,00
```

### 8.2 SerializarGrid (Existing - No Changes)

```vb
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    Dim r As Integer, c As Integer
    Dim linea As String
    Dim salida As String

    salida = ""
    For r = 1 To g.Rows - 1
        If FilaTieneDatos(g, r) Then
            linea = ""
            For c = 0 To g.Cols - 1
                If c > 0 Then linea = linea & SEP_COL
                linea = linea & EscaparDato(g.TextMatrix(r, c))
            Next c
            If Len(salida) > 0 Then salida = salida & SEP_ROW
            salida = salida & linea
        End If
    Next r

    SerializarGrid = salida
End Function
```

### 8.3 DeserializarGrid (Existing - No Changes)

```vb
Private Sub DeserializarGrid(ByRef g As MSFlexGrid, ByVal dataText As String)
    On Error GoTo EH

    Dim rowsArr() As String
    Dim colsArr() As String
    Dim r As Long, c As Long
    Dim targetRow As Integer

    ResetGridRows g

    If Trim$(dataText) = "" Then Exit Sub

    rowsArr = Split(dataText, SEP_ROW)
    g.Rows = UBound(rowsArr) + 2

    targetRow = 1
    For r = LBound(rowsArr) To UBound(rowsArr)
        colsArr = Split(rowsArr(r), SEP_COL)
        For c = 0 To g.Cols - 1
            If c <= UBound(colsArr) Then
                g.TextMatrix(targetRow, c) = DesescaparDato(colsArr(c))
            Else
                g.TextMatrix(targetRow, c) = ""
            End If
        Next c
        targetRow = targetRow + 1
    Next r

    Exit Sub
EH:
    MsgBox "Error deserializando grilla: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

## 9. Printing Specification

### 9.1 Page Setup

- Page size: Letter (8.5" x 11" = 12240 x 15840 twips)
- Orientation: Portrait
- Margins: ~1" all sides (implicit via positioning)
- Font family: Courier New (monospaced for alignment)
- Font sizes:
  - Headers: 10pt bold
  - Section titles: 10pt bold
  - Data: 10pt regular
  - Invoice number in header: 10pt bold
- Default copies: 2 (per Pedido Pato.pdf requirement)

### 9.2 Logo Display Logic (NEW)

**Conditional logo based on grdFacturas data**:

```vb
' Add at start of cmdImprimir_Click (after Printer setup):
Printer.Copies = 2  ' Default to 2 copies

' After setting fonts, BEFORE printing header:
If grdFacturas.Rows > 1 And FilaTieneDatos(grdFacturas, 1) Then
    ' Has facturas data → print logo
    On Error Resume Next
    Dim logoPath As String
    logoPath = App.Path & "\Quilplac2.jpg"

    If Dir(logoPath) <> "" Then
        ' Logo exists → display
        Printer.PaintPicture LoadPicture(logoPath), 600, 200, 7200, 1200
        y = 1600  ' Start content below logo
    Else
        ' Logo missing → fallback to text
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
    ' No facturas → text header only
    Printer.CurrentX = 3000
    Printer.CurrentY = 600
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print "ORDEN DE PAGO"
    Printer.FontBold = False
    Printer.FontSize = 10
    y = 1200
End If
```

**Logo file**: `Quilplac2.jpg` (must be in `App.Path`, same directory as exe)

**Logo dimensions**: 7200x1200 twips (~10cm x 1.67cm)

**Logo position**: x=600, y=200 (top-left with small margin)

### 9.3 Section Order

1. DEUDA
2. CHEQUES
3. TRANSFERENCIA (NEW)
4. FACTURAS (NEW)
5. OTROS (includes grdOtros details + RetIIBB + Efectivo)
6. RESUMEN (totals breakdown)

### 9.4 NEW Section: TRANSFERENCIA

Insert after CHEQUES section:

```vb
' Transferencia
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

    Dim transfRows As Long
    transfRows = grdTransferencia.Rows - 1
    For i = 1 To transfRows
        Dim banco As String, nroCuenta As String, cuit As String
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
```

### 9.5 NEW Section: FACTURAS

Insert after TRANSFERENCIA section:

```vb
' Facturas
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

    Dim factRows As Long
    factRows = grdFacturas.Rows - 1
    For i = 1 To factRows
        Dim nroFact As String, fechaFact As String, descFact As String
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
```

### 9.6 Updated RESUMEN Section

Replace existing TOTALES section with detailed breakdown:

```vb
' RESUMEN
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

If grdTransferencia.Rows > 1 And FilaTieneDatos(grdTransferencia, 1) Then
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Transferencia:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubTransferencia.text))
    y = y + lineH
End If

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

' Firma (existing code continues...)
```

---

## 10. Serialization & Deserialization Code Patterns

### 10.1 cmdGuardar_Click (Updated)

Add serialization for new grids:

```vb
Dim deudaTxt As String, chequesTxt As String, otrosTxt As String
Dim transfTxt As String, factTxt As String  ' NEW

deudaTxt = SerializarGrid(grdDeuda)
chequesTxt = SerializarGrid(grdCheques)
otrosTxt = SerializarGrid(grdOtros)
transfTxt = SerializarGrid(grdTransferencia)  ' NEW
factTxt = SerializarGrid(grdFacturas)  ' NEW

' ... recordset operations ...

rs!DetalleDeuda = deudaTxt
rs!DetalleCheques = chequesTxt
rs!DetalleOtros = otrosTxt
rs!DetalleTransferencia = transfTxt  ' NEW
rs!DetalleFacturas = factTxt  ' NEW
```

### 10.2 CargarOrdenPorNumero (Updated)

Add deserialization for new grids:

```vb
DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
DeserializarGrid grdOtros, NzS(rs!DetalleOtros)
DeserializarGrid grdTransferencia, NzS(rs!DetalleTransferencia)  ' NEW
DeserializarGrid grdFacturas, NzS(rs!DetalleFacturas)  ' NEW
```

### 10.3 LimpiarFormulario (Updated)

Reset new grids:

```vb
ResetGridRows grdDeuda
ResetGridRows grdCheques
ResetGridRows grdOtros
ResetGridRows grdTransferencia  ' NEW
ResetGridRows grdFacturas  ' NEW

txtSubTransferencia.text = "0,00"  ' NEW
txtSubFacturas.text = "0,00"  ' NEW
```

---

## 11. Validation & Error Handling

### 11.1 Required Fields Check (cmdGuardar_Click)

Add validation before saving:

```vb
' Validate required fields
If Trim$(txtNroOrden.text) = "" Or Val(txtNroOrden.text) <= 0 Then
    MsgBox "Número de orden inválido", vbExclamation, "Validación"
    txtNroOrden.SetFocus
    Exit Sub
End If

If Trim$(txtFecha.text) = "" Then
    MsgBox "Fecha requerida", vbExclamation, "Validación"
    txtFecha.SetFocus
    Exit Sub
End If

If Trim$(txtProveedor.text) = "" Then
    MsgBox "Proveedor requerido", vbExclamation, "Validación"
    txtProveedor.SetFocus
    Exit Sub
End If

' At least one payment method or debt required
Dim tieneDeuda As Boolean, tienePago As Boolean
tieneDeuda = (grdDeuda.Rows > 1 And FilaTieneDatos(grdDeuda, 1))
tienePago = (grdCheques.Rows > 1 And FilaTieneDatos(grdCheques, 1)) Or _
            (grdOtros.Rows > 1 And FilaTieneDatos(grdOtros, 1)) Or _
            (grdTransferencia.Rows > 1 And FilaTieneDatos(grdTransferencia, 1)) Or _
            (grdFacturas.Rows > 1 And FilaTieneDatos(grdFacturas, 1)) Or _
            (ParseCurrency(txtEfectivo.text) > 0) Or _
            (ParseCurrency(GetRetIIBBText()) > 0)

If Not tieneDeuda And Not tienePago Then
    MsgBox "Debe ingresar al menos un concepto de deuda o pago", vbExclamation, "Validación"
    Exit Sub
End If
```

---

## Appendix A: Summary of Changes

| Category | Changes |
|----------|---------|
| Database Schema | +2 MEMO columns (DetalleTransferencia, DetalleFacturas) |
| Grids | +2 grids (grdTransferencia with 4 cols, grdFacturas with 4 cols) |
| UI Controls | +4 TextBoxes (txtSubTransferencia, txtSubFacturas + labels), +8 buttons |
| Enums | +2 enums (eColTransferencia, eColFacturas) |
| Inline Editing | NEW: Replace InputBox with TextBox overlay for all grids |
| Date Validation | NEW: txtFecha_LostFocus with DD/MM/YYYY strict validation |
| Calculation | Update RecalcularTodo to include 2 new grids in totalPago |
| Serialization | Reuse existing functions, add calls for 2 new grids |
| Printing | +2 sections (TRANSFERENCIA, FACTURAS), conditional logo, detailed RESUMEN, 2 copies default |
| Validation | +Required field checks, +CUIT format (XX-XXXXXXXX-X), +Invoice format (XXXX-XXXXXXXX) |

---

**End of Specification**
