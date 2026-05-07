# SDD orchestration for FormOrdenPago.frm change

**Session ID:** ses_2159648e9ffeR1ogXMx10Q2yrT
**Created:** 2/5/2026, 17:38:09
**Updated:** 5/5/2026, 19:25:00

---

## User

Follow the SDD orchestrator workflow to continue the active change.

WORKFLOW:
1. Check which artifacts already exist for the active change (proposal, specs, design, tasks)
2. Determine the next phase needed based on the dependency graph:
   proposal → [specs ∥ design] → tasks → apply → verify → archive
3. Launch the appropriate sub-agent(s) for the next phase
4. Present the result and ask the user to proceed

CONTEXT:
- Working directory: /c/TrabajosActivos/SPC-Core
- Current project: SPC-Core
- Change name: FormOrdenPago.frm
- Artifact store mode: engram

ENGRAM NOTE:
To check which artifacts exist, search: mem_search(query: "sdd/FormOrdenPago.frm/", project: "{project}") to list all artifacts for this change.
Sub-agents handle persistence automatically with topic_key "sdd/FormOrdenPago.frm/{type}".

Read the orchestrator instructions to coordinate this workflow. Do NOT execute phase work inline — delegate to sub-agents.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 11.3s)

_Thinking:_



**Tool: engram_mem_search**

**Tool: engram_mem_search**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 73.1s)

_Thinking:_



**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 8.7s)

_Thinking:_



Done — I followed the SDD orchestration flow and delegated the next phase.

## 1) Artifact check (active change: `FormOrdenPago.frm`)
I checked engram for `sdd/FormOrdenPago.frm/` artifacts and found **none** existing initially (no proposal/spec/design/tasks yet).

## 2) Next phase by dependency graph
Given no artifacts existed, the correct next phase was: **proposal**.

## 3) Sub-agent launched
I launched a sub-agent to execute **SDD Proposal** for this change (with engram persistence).

## 4) Result from sub-agent
**Status:** completed

**Created artifacts in engram:**
- `sdd/FormOrdenPago.frm/proposal` ✅
- `sdd/FormOrdenPago.frm/state` ✅ (DAG state updated; proposal marked complete)

**Recommended next phases:**
- `sdd-spec` (next required)
- `sdd-design` (can run in parallel with spec or immediately after)
- then `sdd-tasks`

**Noted risks:**
- Business-rule ambiguity for payment orders
- Access schema/report coupling risk
- Potential regression risk in menu/global-state integrations

---

Shall I proceed now with the next phase and launch **`sdd-spec`** for `FormOrdenPago.frm`?

---

## User

please anwers in spanish

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 4.4s)

¡Perfecto! A partir de ahora te respondo en español. 👍

Ya quedó creada la **propuesta** de `FormOrdenPago.frm` en Engram.

¿Querés que avance con la siguiente fase y lance **`sdd-spec`** ahora?

---

## User

sdd-spec para generar la propuesta, tomar lo indicado en "Pedido Pato.pdf" y [Image 1] 

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 92.8s)

_Thinking:_

**Delegating tasks effectively**

I see I need to delegate some tasks here. Maybe I should pass along the user-provided artifact references. I’ll ask where the PDF is located since the subagent can help find it—since the orchestrator can’t read files. I think I can delegate in general to run the SDD-spec skill, which can include extracting image details and text from the prompt. I’ll just focus on doing the task!

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 9.2s)

Listo ✅  
Ya ejecuté **`sdd-spec`** para el cambio `FormOrdenPago.frm`, tomando como base:

- la propuesta previa en Engram,
- **`Pedido Pato.pdf`**,
- y la imagen de referencia del comprobante **“PAGO A PROVEEDORES”**.

### Resultado
**Estado:** `completed`

**Artifacts guardados en Engram:**
- `sdd/FormOrdenPago.frm/spec`
- `sdd/FormOrdenPago.frm/state`
- adicional: `sdd/FormOrdenPago.frm/saldo-rule` (decisión sobre regla de saldo/redondeo)

### Qué quedó especificado
- Requisitos y escenarios (éxito/error) para:
  - Deuda
  - Cheques
  - Otros
  - Efectivo
  - Retención IIBB
  - Impresión del comprobante
- Reglas de cálculo/validación explícitas:
  - sumatorias y subtotales,
  - consistencia **Total Deuda vs Total Pago**,
  - cálculo de **Saldo**,
  - tolerancia de redondeo `<= 0,01`,
  - manejo de saldo a favor/en contra.

### Riesgos marcados
- Convención contable de “saldo a favor” (perspectiva empresa/proveedor).
- Definir si diferencias > 0,01 bloquean o permiten guardar con confirmación.
- Ajuste fino del layout de impresión para que replique el formato real.

¿Avanzo con la siguiente fase y lanzo **`sdd-design`** ahora?

---

## User

Si, avancemos, pero tené en cuenta lo solicitado en "Pedido Pato.pdf", no valida información contra la bbdd, solo se rellenan campos. Calculá, Totales, Subtotales pero que el usuario pueda modificar. Es todo escrito a mano, no trae nada de la bbdd, ni valida, por ahora.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 86.9s)

_Thinking:_



**Tool: engram_mem_save_prompt**

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 8.5s)

Perfecto, avanzamos con tus condiciones ✅

Ya quedó ejecutada la fase **`sdd-design`** para `FormOrdenPago.frm`, incorporando exactamente lo que pediste:

- sin traer datos de BBDD,
- sin validar contra BBDD,
- carga 100% manual de campos,
- cálculo automático de subtotales/totales/saldo,
- **pero editable manualmente por el usuario**.

### Resultado de la fase
**Estado:** `completed`

**Artifacts guardados en Engram:**
- `sdd/FormOrdenPago.frm/design`
- `sdd/FormOrdenPago.frm/state`
- decisión adicional: `decision/formordenpago-manual-only`

### Qué definió el diseño
- Arquitectura **manual-only** (sin DAO/ADO/SQL).
- Modelo de estado por bloques (Deuda, Cheques, Otros, Efectivo, Retención, Total, Saldo).
- Reglas de “auto-calculado vs manual” con banderas de override.
- Acciones UX sugeridas:
  - **Recalcular**
  - **Restaurar Auto**
  - indicador visual **AUTO/MANUAL**
- Estrategia de impresión para replicar el formato “PAGO A PROVEEDORES”.

### Siguiente fase recomendada
👉 **`sdd-tasks`** para bajar esto a checklist concreta de implementación.

¿La lanzo ahora?

---

## User

si, tené en cuenta en utilizar la codificación ANSI que corresponda para que no pinche o ponga caracteres extraños. Podemos crear una simple tabla en el access para guardar los datos cargados por si se necesita re-imprimir la orden. Agregué en la tabla "UltimosNumeros" los 2 campos correspondientes, para tomar el ultimo numero guardado e iniciar el modulo con el siguiente. Si podés creá la tabla OrdenPago en la base de datos access, si no decime que campos tipos y tamaños y lo hago yo.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 88.3s)

_Thinking:_



**Tool: engram_mem_save_prompt**

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 8.7s)

_Thinking:_



Excelente, ya quedó ejecutada la fase **`sdd-tasks`** ✅

### Resultado
- Artifact generado: `sdd/FormOrdenPago.frm/tasks`
- Estado DAG actualizado: `sdd/FormOrdenPago.frm/state` → **tasks completado**
- Próximo paso: **`sdd-apply`**

### Lo que ya contemplan las tareas
- Compatibilidad **VB6 ANSI + CRLF**
- Carga manual con cálculos editables
- Persistencia mínima en Access para **reimpresión**
- Numeración correlativa desde `UltimosNumeros` (**último + 1**)
- Impresión del comprobante
- Plan B con esquema manual si no se puede crear tabla automáticamente

---

Si te parece, ahora lanzo **`sdd-apply`** para implementar todo, incluyendo:

1) intentar crear tabla `OrdenPago` en Access,  
2) si no se puede por entorno/permisos, devolverte esquema exacto (campos/tipos/tamaños) para que la crees vos.

¿Avanzo?

---

## User

si

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 227.9s)

_Thinking:_



**Tool: engram_mem_save_prompt**

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 10.0s)

_Thinking:_

**Confirming user preferences**

I need to ask the user for their choice and summarize what I’ve gathered so far. It seems like they might have intended to say "yes," but now I'm not sure since something is blocked. As the orchestrator, I can possibly launch another subagent to handle option two, which might be more backend-oriented. I should clearly ask the user what they prefer. Let’s make sure to keep things straightforward and helpful!

Perfecto. Avancé con `sdd-apply`, pero el sub-agente quedó **bloqueado parcialmente** por una limitación real del entorno (sin VB6 IDE para diseñar/formatear de forma segura el `.frm` completo).

### Lo que sí quedó confirmado/realizado
- Se validó Access vía ACE OLEDB.
- Se confirmó `UltimosNumeros` con columnas:
  - `IdSucursal`
  - `IDTabla`
  - `UltimoNumero`
- Se detectaron filas para `tOrdenPago` (sucursales 2 y 5).
- Se pudo **crear tabla `OrdenPago`** + índices en la base (según reporte del sub-agente).

### Bloqueo actual
Para implementar bien el `FormOrdenPago.frm` completo (controles/eventos/layout) sin romper compatibilidad VB6, es muy riesgoso hacerlo “a mano” sin IDE.

---

Te propongo elegir uno de estos caminos para seguir **ya**:

1) **Implementación mínima completa sin diseñador complejo**  
   (funciona, pero visualmente básica en el `.frm`).

2) **Backend primero (recomendado)**  
   Dejar listo ahora: numeración, guardado, reimpresión y lógica de cálculos en módulos;  
   luego ajustar layout visual en VB6 IDE (vos o yo en otra pasada).

Si querés, recomiendo **opción 2** para no arriesgar corrupción del formulario.  
Decime “2” y la ejecuto inmediatamente.

---

## User

te puedo dar acceso al IDE lo tengo abierto en el proyecto, como hago?

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 10.3s)

_Thinking:_

**Explaining orchestration limits**

I need to respond regarding orchestration, and it looks like I won’t need any tools for that. I might want to explain that I can’t provide a remote control function for the user. Instead, I can guide them through the manual steps they’d need to take. I can also generate patches for the user, so they have a clear way to tackle their tasks. I hope this makes sense and helps!

¡Genial que lo tengas abierto! 🙌  
Yo no puedo “tomar control” directo del IDE ni conectarme remoto a tu pantalla desde acá.

Pero sí podemos trabajar **en modo asistido** muy bien:

1) Vos ejecutás en VB6 IDE lo que te voy indicando (paso a paso corto).
2) Yo preparo el contenido exacto (código/eventos/propiedades).
3) Vos pegás/ajustás y me pasás cualquier error textual.
4) Iteramos hasta que compile y funcione.

Si te parece, arrancamos ahora con un flujo práctico:

- **Paso A:** confirmar nombre del formulario (`FormOrdenPago.frm`) y si ya tiene controles creados.
- **Paso B:** te doy lista exacta de controles (nombre/tipo) para armar rápido el layout.
- **Paso C:** te paso código de eventos (cálculo, override manual, guardar, reimprimir, numeración con `UltimosNumeros`).
- **Paso D:** validación manual en IDE.

Si querés, empiezo ya con el **Paso B** (lista de controles mínima recomendada).

---

## User

Ya está el FormOrdenPago.frm creado, lo unico que tiene es un Frame para contener los controles, dime que controles es necesario Crear, Labels, Textbox, Combos Command Buttons etc.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 18.2s)

_Thinking:_

**Considering task delegation**

I need to focus on obeying the orchestrator and think about whether I should delegate all real work, including analysis. The user has asked for a list of controls. I could provide a concise answer based on known information, even without reading the code. It's more of a design-related question, but I can certainly keep my response brief and in Spanish as requested. I’ll aim for clarity!

Perfecto. Te paso una **estructura mínima y suficiente** para que funcione como pediste (manual, calculado pero editable, reimpresión).

## 1) Encabezado
Dentro del Frame principal (`fraOrdenPago`):

- **Label + TextBox**
  - `lblNroOrden` / `txtNroOrden` (Locked=True)
  - `lblFecha` / `txtFecha` (o `DTPicker` si lo usás)
  - `lblProveedor` / `txtProveedor`

## 2) Bloque DEUDA
Usá grilla para que sea más simple de mantener:

- **MSFlexGrid** `grdDeuda`  
  Columnas: `Detalle | Importe`
- **CommandButton**
  - `cmdAddDeuda` (Agregar fila)
  - `cmdDelDeuda` (Eliminar fila)
- **Label + TextBox**
  - `lblSubDeuda` / `txtSubDeuda` (editable manualmente)

## 3) Bloque CHEQUES
- **MSFlexGrid** `grdCheques`  
  Columnas: `NroCheque | Banco | Vto | Importe`
- **CommandButton**
  - `cmdAddCheque`
  - `cmdDelCheque`
- **Label + TextBox**
  - `lblSubCheques` / `txtSubCheques` (editable)

## 4) Bloque OTROS
- **MSFlexGrid** `grdOtros`  
  Columnas: `Referencia | Importe`
- **CommandButton**
  - `cmdAddOtro`
  - `cmdDelOtro`
- **Label + TextBox**
  - `lblSubOtros` / `txtSubOtros` (editable)

## 5) Totales finales
- **Label + TextBox**
  - `lblEfectivo` / `txtEfectivo`
  - `lblRetIIBB` / `txtRetIIBB`
  - `lblTotalPago` / `txtTotalPago` (editable)
  - `lblSaldo` / `txtSaldo` (editable)
  - `lblImporteLetras` / `txtImporteLetras` (Locked=True o editable si querés)
- **CheckBox** (opcional pero recomendable para override)
  - `chkManualSubDeuda`
  - `chkManualSubCheques`
  - `chkManualSubOtros`
  - `chkManualTotalPago`
  - `chkManualSaldo`

## 6) Acciones
- **CommandButton**
  - `cmdRecalcular`
  - `cmdRestaurarAuto`
  - `cmdGuardar`
  - `cmdReimprimir`
  - `cmdImprimir`
  - `cmdNuevo`
  - `cmdCerrar`

## 7) Persistencia/reimpresión
- **ComboBox** `cboBuscarOrden` (para elegir Nro Orden guardado)
- **CommandButton** `cmdCargarOrden`

---

### Recomendación práctica VB6
- Todo `.frm/.bas` en **ANSI + CRLF**.
- `txtNroOrden` bloqueado (solo lectura).
- Montos con alineación derecha (`Alignment = 1 - Right Justify`).
- Si no querés usar muchos `CheckBox`, podés usar una sola convención:  
  cuando el usuario edita un subtotal/total, lo marcás como manual internamente.

---

Si querés, en el próximo paso te doy:
1) **propiedades exactas** sugeridas (tab order, formatos), y  
2) el **esqueleto de eventos** (`Change`, `Click`) para que recalculen todo.

---

## User

Bueno ya creé el formulario con los controles que me indicaste con esos nombres excepto los checkbox, por el momento eso no creo que sea necesario implementarlo

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 7.2s)

_Thinking:_



---

## User

Bueno ya creé el formulario con los controles que me indicaste con esos nombres excepto los checkbox, por el momento eso no creo que sea necesario implementarlo, te paso el código: Private Sub MSFlexGrid1_Click()

End Sub


Private Sub Label1_Click()

End Sub


Private Sub Form_Load()

End Sub podés pasarme el código en un archivo plano, copio y pego

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 146.7s)

_Thinking:_



**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 111.9s)

Perfecto. Te paso el bloque listo para copiar y pegar en `FormOrdenPago.frm` (módulo de código).

**Prerequisitos mínimos:**  
- Referencia DAO 3.6 activa.  
- `BaseSPC As Database` global ya abierta en tu proyecto.  
- Si usás otro `IdSucursal`, ajustá `GetIdSucursalActual()`.

```vb
Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Sub Form_Load()
    On Error GoTo EH

    m_Cargando = True

    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.Text = ""
    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.Text = "0,00"
    txtRetIIBB.Text = "0,00"

    txtSubDeuda.Text = "0,00"
    txtSubCheques.Text = "0,00"
    txtSubOtros.Text = "0,00"
    txtTotalPago.Text = "0,00"
    txtSaldo.Text = "0,00"
    txtImporteLetras.Text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros

    If nuevoNumero Then
        txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    g.TextMatrix(r, c) = Trim$(v)

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' Compatibilidad con evento legado mencionado por usuario
Private Sub MSFlexGrid1_Click()
    ' TODO: Alias legado. Si existe una grilla historica llamada MSFlexGrid1,
    ' redireccionar aca la logica necesaria.
End Sub

' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
    RecalcularTodo
End Sub

Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    ' No bloquea. Recalculo bajo demanda por boton.
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    ' No bloquea. Recalculo bajo demanda por boton.
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)

    efectivo = ParseCurrency(txtEfectivo.Text)
    retIIBB = ParseCurrency(txtRetIIBB.Text)

    totalPago = subCheques + subOtros + efectivo + retIIBB
    saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.Text = FormatMoney(subDeuda)
    txtSubCheques.Text = FormatMoney(subCheques)
    txtSubOtros.Text = FormatMoney(subOtros)
    txtTotalPago.Text = FormatMoney(totalPago)
    txtSaldo.Text = FormatMoney(saldo)
    txtImporteLetras.Text = NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String

    nroOrden = CLng(Val(NzS(txtNroOrden.Text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!NroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.Text)
    rs!Proveedor = NzS(txtProveedor.Text)

    rs!SubDeuda = ParseCurrency(txtSubDeuda.Text)
    rs!SubCheques = ParseCurrency(txtSubCheques.Text)
    rs!SubOtros = ParseCurrency(txtSubOtros.Text)
    rs!Efectivo = ParseCurrency(txtEfectivo.Text)
    rs!RetIIBB = ParseCurrency(txtRetIIBB.Text)
    rs!TotalPago = ParseCurrency(txtTotalPago.Text)
    rs!Saldo = ParseCurrency(txtSaldo.Text)
    rs!ImporteLetras = NzS(txtImporteLetras.Text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim texto As String

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        texto = CStr(NzL(rs!NroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!NroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.Text = CStr(NzL(rs!NroOrden))
    txtFecha.Text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.Text = NzS(rs!Proveedor)

    txtSubDeuda.Text = FormatMoney(NzC(rs!SubDeuda))
    txtSubCheques.Text = FormatMoney(NzC(rs!SubCheques))
    txtSubOtros.Text = FormatMoney(NzC(rs!SubOtros))
    txtEfectivo.Text = FormatMoney(NzC(rs!Efectivo))
    txtRetIIBB.Text = FormatMoney(NzC(rs!RetIIBB))
    txtTotalPago.Text = FormatMoney(NzC(rs!TotalPago))
    txtSaldo.Text = FormatMoney(NzC(rs!Saldo))
    txtImporteLetras.Text = NzS(rs!ImporteLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub

' ------------------------------------------------------------
' Impresion (stub funcional)
' ------------------------------------------------------------
Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Printer.Print "==========================================="
    Printer.Print "ORDEN DE PAGO Nro: " & NzS(txtNroOrden.Text)
    Printer.Print "Fecha: " & NzS(txtFecha.Text)
    Printer.Print "Proveedor: " & NzS(txtProveedor.Text)
    Printer.Print "-------------------------------------------"
    Printer.Print "SubTotal Deuda:   " & NzS(txtSubDeuda.Text)
    Printer.Print "SubTotal Cheques: " & NzS(txtSubCheques.Text)
    Printer.Print "SubTotal Otros:   " & NzS(txtSubOtros.Text)
    Printer.Print "Efectivo:         " & NzS(txtEfectivo.Text)
    Printer.Print "Ret IIBB:         " & NzS(txtRetIIBB.Text)
    Printer.Print "TOTAL PAGO:       " & NzS(txtTotalPago.Text)
    Printer.Print "SALDO:            " & NzS(txtSaldo.Text)
    Printer.Print "Importe en letras: " & NzS(txtImporteLetras.Text)
    Printer.Print "==========================================="
    Printer.EndDoc

    ' TODO: Reemplazar por layout final (Crystal / VSReport / formato legal)
    MsgBox "Impresion enviada (stub).", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    idSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    idSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    ' TODO: Conectar a variable global real del sistema si existe (ej: codSucursal)
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    Dim r As Integer, c As Integer
    Dim line As String
    Dim out As String

    out = ""
    For r = 1 To g.Rows - 1
        If FilaTieneDatos(g, r) Then
            line = ""
            For c = 0 To g.Cols - 1
                If c > 0 Then line = line & SEP_COL
                line = line & EscaparDato(g.TextMatrix(r, c))
            Next c
            If Len(out) > 0 Then out = out & SEP_ROW
            out = out & line
        End If
    Next r

    SerializarGrid = out
End Function

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

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ParseCurrency = 0
        Exit Function
    End If

    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    ParseCurrency = CCur(Val(t))
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    FormatMoney = Format$(v, "0.00")
    FormatMoney = Replace(FormatMoney, ".", ",")
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal importe As Currency) As String
    ' TODO: Reemplazar por funcion completa de importe a letras
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(importe)
End Function
```

Si te da error al compilar, pasame **el texto exacto del error y la línea** y te lo ajusto al vuelo.

---

## User

ahí compiló pero hay errores, mañana seguimos persistí todo y seguimos desde acá.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 18.2s)

_Thinking:_

**Following session protocols**

I need to make sure I comply with the session close protocol by creating a mandatory session summary. Responding concisely in Spanish is also important. I should call the memory tools to save what we've discussed and maybe conclude with a session end. I want to ensure I'm following the right steps to wrap things up effectively!

**Tool: engram_mem_session_summary**

**Tool: engram_mem_session_end**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 13.0s)

_Thinking:_

**Addressing summary failure**

I need to let the user know that the summary attempt failed due to a lock issue. It seems like I should also consider trying again with a different session ID or perhaps just retrying with the same database since the lock might be transient. There’s a chance that the next attempt will work, so I'm definitely considering a retry to resolve this issue for the user.

**Tool: engram_mem_session_summary**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 9.6s)

Perfecto, quedó todo guardado para retomar mañana ✅

Persistí el resumen de sesión con:
- lo ya hecho en SDD (proposal/spec/design/tasks + apply en curso),
- tus preferencias (ANSI + CRLF, carga manual, persistencia mínima),
- y lo pendiente (corregir errores exactos que te aparecen).

Mañana seguimos desde acá.  
Cuando vuelvas, pasame **el primer error exacto (mensaje + línea)** y lo resolvemos rápido.

---

## User

gracias, hasta mañana.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 3.8s)

¡De nada! 🙌  
Hasta mañana, seguimos desde ahí.

---

## User

Hola buenas tardes, fijate las capturas C:/users/pablo/downloads/ordenpago-01.png 02 03 04 y 05, fue tirando esos errores. Te paso el código: Private Sub MSFlexGrid1_Click()

End Sub


Private Sub Label1_Click()

End Sub


Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Sub Form_Load()
    On Error GoTo EH

    m_Cargando = True

    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.text = ""
    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.text = "0,00"
    txtRetIIBB.text = "0,00"

    txtSubDeuda.text = "0,00"
    txtSubCheques.text = "0,00"
    txtSubOtros.text = "0,00"
    txtTotalPago.text = "0,00"
    txtSaldo.text = "0,00"
    txtImporteLetras.text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros

    If nuevoNumero Then
        txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    g.TextMatrix(r, c) = Trim$(v)

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub


' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
    RecalcularTodo
End Sub

Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    ' No bloquea. Recalculo bajo demanda por boton.
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    ' No bloquea. Recalculo bajo demanda por boton.
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim Saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)

    efectivo = ParseCurrency(txtEfectivo.text)
    retIIBB = ParseCurrency(txtRetIIBB.text)

    totalPago = subCheques + subOtros + efectivo + retIIBB
    Saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.text = FormatMoney(subDeuda)
    txtSubCheques.text = FormatMoney(subCheques)
    txtSubOtros.text = FormatMoney(subOtros)
    txtTotalPago.text = FormatMoney(totalPago)
    txtSaldo.text = FormatMoney(Saldo)
    txtImporteLetras.text = NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String

    nroOrden = CLng(Val(NzS(txtNroOrden.text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!nroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.text)
    rs!Proveedor = NzS(txtProveedor.text)

    rs!subDeuda = ParseCurrency(txtSubDeuda.text)
    rs!subCheques = ParseCurrency(txtSubCheques.text)
    rs!subOtros = ParseCurrency(txtSubOtros.text)
    rs!efectivo = ParseCurrency(txtEfectivo.text)
    rs!retIIBB = ParseCurrency(txtRetIIBB.text)
    rs!totalPago = ParseCurrency(txtTotalPago.text)
    rs!Saldo = ParseCurrency(txtSaldo.text)
    rs!ImporteLetras = NzS(txtImporteLetras.text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim Texto As String

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        Texto = CStr(NzL(rs!nroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem Texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!nroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.text = CStr(NzL(rs!nroOrden))
    txtFecha.text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.text = NzS(rs!Proveedor)

    txtSubDeuda.text = FormatMoney(NzC(rs!subDeuda))
    txtSubCheques.text = FormatMoney(NzC(rs!subCheques))
    txtSubOtros.text = FormatMoney(NzC(rs!subOtros))
    txtEfectivo.text = FormatMoney(NzC(rs!efectivo))
    txtRetIIBB.text = FormatMoney(NzC(rs!retIIBB))
    txtTotalPago.text = FormatMoney(NzC(rs!totalPago))
    txtSaldo.text = FormatMoney(NzC(rs!Saldo))
    txtImporteLetras.text = NzS(rs!ImporteLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub

' ------------------------------------------------------------
' Impresion (stub funcional)
' ------------------------------------------------------------
Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Printer.Print "==========================================="
    Printer.Print "ORDEN DE PAGO Nro: " & NzS(txtNroOrden.text)
    Printer.Print "Fecha: " & NzS(txtFecha.text)
    Printer.Print "Proveedor: " & NzS(txtProveedor.text)
    Printer.Print "-------------------------------------------"
    Printer.Print "SubTotal Deuda:   " & NzS(txtSubDeuda.text)
    Printer.Print "SubTotal Cheques: " & NzS(txtSubCheques.text)
    Printer.Print "SubTotal Otros:   " & NzS(txtSubOtros.text)
    Printer.Print "Efectivo:         " & NzS(txtEfectivo.text)
    Printer.Print "Ret IIBB:         " & NzS(txtRetIIBB.text)
    Printer.Print "TOTAL PAGO:       " & NzS(txtTotalPago.text)
    Printer.Print "SALDO:            " & NzS(txtSaldo.text)
    Printer.Print "Importe en letras: " & NzS(txtImporteLetras.text)
    Printer.Print "==========================================="
    Printer.EndDoc

    ' TODO: Reemplazar por layout final (Crystal / VSReport / formato legal)
    MsgBox "Impresion enviada (stub).", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim IdSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    IdSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(IdSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = IdSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim IdSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    IdSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(IdSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = IdSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    ' TODO: Conectar a variable global real del sistema si existe (ej: codSucursal)
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    Dim r As Integer, c As Integer
    Dim line As String
    Dim out As String

    out = ""
    For r = 1 To g.Rows - 1
        If FilaTieneDatos(g, r) Then
            line = ""
            For c = 0 To g.Cols - 1
                If c > 0 Then line = line & SEP_COL
                line = line & EscaparDato(g.TextMatrix(r, c))
            Next c
            If Len(out) > 0 Then out = out & SEP_ROW
            out = out & line
        End If
    Next r

    SerializarGrid = out
End Function

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

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ParseCurrency = 0
        Exit Function
    End If

    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    ParseCurrency = CCur(Val(t))
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    FormatMoney = Format$(v, "0.00")
    FormatMoney = Replace(FormatMoney, ".", ",")
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal Importe As Currency) As String
    ' TODO: Reemplazar por funcion completa de importe a letras
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(Importe)
End Function


Private Sub txtProveedor_Change()

End Sub 

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 249.5s)

_Thinking:_

**Delegating code inspection**

I realize I need to delegate tasks appropriately since the orchestrator shouldn't handle code or image analysis directly. I think I need to launch a sub-agent to inspect the code, especially given the potential errors from the screenshots. It seems I should prompt the general agent to produce a corrected code patch and provide clear steps for the user. I’ll also explore if a subagent can read the screenshots using available tools. Let's get started on that!

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 118.6s)

¡Hola, buenas tardes! 👋  
Gracias por pasar todo tan completo. Ya revisé y te dejo la corrección lista.

Los errores que te está tirando vienen de 3 puntos típicos:

1) `Option Explicit` no debe quedar después de Subs (tiene que estar arriba del todo).  
2) `BaseSPC` a veces llega `Nothing` y explota en `OpenRecordset`.  
3) En varios lugares se usa `txtRetIIBB`, pero si ese control no existe en el diseño del form, rompe.

---

### Qué hacer ahora (rápido)

1. En el módulo de `FormOrdenPago.frm`, **borrá todo** lo que tenés ahora (incluidos esos Subs vacíos `MSFlexGrid1_Click`, `Label1_Click` si no los usás).  
2. Pegá este bloque completo tal cual:

```vb
Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Sub Form_Load()
    On Error GoTo EH

    m_Cargando = True

    EnsureDatabaseReady
    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.Text = ""
    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.Text = "0,00"
    SetRetIIBBText "0,00"

    txtSubDeuda.Text = "0,00"
    txtSubCheques.Text = "0,00"
    txtSubOtros.Text = "0,00"
    txtTotalPago.Text = "0,00"
    txtSaldo.Text = "0,00"
    txtImporteLetras.Text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros

    If nuevoNumero Then
        txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    g.TextMatrix(r, c) = Trim$(v)

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
    RecalcularTodo
End Sub

Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)

    efectivo = ParseCurrency(txtEfectivo.Text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    totalPago = subCheques + subOtros + efectivo + retIIBB
    saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.Text = FormatMoney(subDeuda)
    txtSubCheques.Text = FormatMoney(subCheques)
    txtSubOtros.Text = FormatMoney(subOtros)
    txtTotalPago.Text = FormatMoney(totalPago)
    txtSaldo.Text = FormatMoney(saldo)
    txtImporteLetras.Text = NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String

    EnsureDatabaseReady

    nroOrden = CLng(Val(NzS(txtNroOrden.Text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!nroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.Text)
    rs!Proveedor = NzS(txtProveedor.Text)

    rs!subDeuda = ParseCurrency(txtSubDeuda.Text)
    rs!subCheques = ParseCurrency(txtSubCheques.Text)
    rs!subOtros = ParseCurrency(txtSubOtros.Text)
    rs!efectivo = ParseCurrency(txtEfectivo.Text)
    rs!retIIBB = ParseCurrency(GetRetIIBBText())
    rs!totalPago = ParseCurrency(txtTotalPago.Text)
    rs!saldo = ParseCurrency(txtSaldo.Text)
    rs!ImporteLetras = NzS(txtImporteLetras.Text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim texto As String

    EnsureDatabaseReady

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        texto = CStr(NzL(rs!nroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!nroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & CStr(Err.Number) & " - " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.Text = CStr(NzL(rs!nroOrden))
    txtFecha.Text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.Text = NzS(rs!Proveedor)

    txtSubDeuda.Text = FormatMoney(NzC(rs!subDeuda))
    txtSubCheques.Text = FormatMoney(NzC(rs!subCheques))
    txtSubOtros.Text = FormatMoney(NzC(rs!subOtros))
    txtEfectivo.Text = FormatMoney(NzC(rs!efectivo))
    SetRetIIBBText FormatMoney(NzC(rs!retIIBB))
    txtTotalPago.Text = FormatMoney(NzC(rs!totalPago))
    txtSaldo.Text = FormatMoney(NzC(rs!saldo))
    txtImporteLetras.Text = NzS(rs!ImporteLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub

' ------------------------------------------------------------
' Impresion (stub funcional)
' ------------------------------------------------------------
Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Printer.Print "==========================================="
    Printer.Print "ORDEN DE PAGO Nro: " & NzS(txtNroOrden.Text)
    Printer.Print "Fecha: " & NzS(txtFecha.Text)
    Printer.Print "Proveedor: " & NzS(txtProveedor.Text)
    Printer.Print "-------------------------------------------"
    Printer.Print "SubTotal Deuda:   " & NzS(txtSubDeuda.Text)
    Printer.Print "SubTotal Cheques: " & NzS(txtSubCheques.Text)
    Printer.Print "SubTotal Otros:   " & NzS(txtSubOtros.Text)
    Printer.Print "Efectivo:         " & NzS(txtEfectivo.Text)
    Printer.Print "Ret IIBB:         " & NzS(GetRetIIBBText())
    Printer.Print "TOTAL PAGO:       " & NzS(txtTotalPago.Text)
    Printer.Print "SALDO:            " & NzS(txtSaldo.Text)
    Printer.Print "Importe en letras: " & NzS(txtImporteLetras.Text)
    Printer.Print "==========================================="
    Printer.EndDoc

    MsgBox "Impresion enviada (stub).", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
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

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ParseCurrency = 0
        Exit Function
    End If

    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    ParseCurrency = CCur(Val(t))
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    FormatMoney = Format$(v, "0.00")
    FormatMoney = Replace(FormatMoney, ".", ",")
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal importe As Currency) As String
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(importe)
End Function

' ------------------------------------------------------------
' Helpers de compatibilidad (control opcional txtRetIIBB)
' ------------------------------------------------------------
Private Sub EnsureDatabaseReady()
    On Error GoTo EH

    If BaseSPC Is Nothing Then
        Set BaseSPC = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    End If
    Exit Sub
EH:
    MsgBox "No se pudo abrir la base de datos: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

Private Function TieneControl(ByVal nombreControl As String) As Boolean
    On Error GoTo EH
    Dim ctl As Control
    Set ctl = Me.Controls(nombreControl)
    TieneControl = Not (ctl Is Nothing)
    Exit Function
EH:
    TieneControl = False
End Function

Private Function GetRetIIBBText() As String
    On Error GoTo EH
    If TieneControl("txtRetIIBB") Then
        GetRetIIBBText = CStr(Me.Controls("txtRetIIBB").Text)
    Else
        GetRetIIBBText = "0,00"
    End If
    Exit Function
EH:
    GetRetIIBBText = "0,00"
End Function

Private Sub SetRetIIBBText(ByVal valor As String)
    On Error Resume Next
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").Text = valor
    End If
End Sub
```

---

Si después de pegar esto te aparece otro error, pasame **solo**:
- mensaje exacto,
- número de línea resaltada,
- y en qué acción ocurre (abrir form / guardar / cargar / imprimir).  
Y te doy el fix puntual enseguida.

---

## User

Anduvo bien, guardó, tuve que cambiar algunos nombres de campos que estaban mal.  Necesito pasar de campo a campo sea con enter y / o tab. Y que la información en las grillas, en vez de dobleclick puede ser con enter?. El Boton restaurar cmdRestaurarAuto_Click debería limpiar la pantalla para cargar una nueva. Y Necesito Un command para imprimir el pago con el formato similar al que pasé originalmente con la foto

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 72.7s)

_Thinking:_

**Delegating code enhancements**

I need to delegate the real work since the user is asking for enhancements. I should probably launch a sub-agent to provide instructions for the code changes. I don't think any tools are reading the file; it looks like the user opened a "temporal.md" that may not be relevant. Plus, I should respond in Spanish. I'll need to ask about controls for entering navigation, maybe using KeyPreview, and focus on generating incremental code modifications.

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 37.0s)

¡Excelente avance! 🙌  
Perfecto, hacemos esos 4 ajustes.

Te paso **solo cambios puntuales** para pegar (adaptá nombres que te queden distinto, porque vos cambiaste algunos).

---

### 1) Enter como Tab en el formulario

**Propiedad en diseño:**
- `FormOrdenPago.KeyPreview = True`

**Código:**
```vb
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As Control
        Set c = Me.ActiveControl

        If Not c Is Nothing Then
            Select Case TypeName(c)
                Case "MSFlexGrid", "MshFlexGrid", "VSFlexGrid"
                    ' En grilla se maneja aparte
                Case Else
                    KeyCode = 0
                    SendKeys "{TAB}"
            End Select
        End If
    End If
End Sub
```

---

### 2) En grillas, editar con Enter (en vez de doble click)

```vb
Private Sub grdDeuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdDeuda
    End If
End Sub

Private Sub grdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdCheques
    End If
End Sub

Private Sub grdOtros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdOtros
    End If
End Sub
```

> Como ya tenés `EditarCeldaGrid`, esto te reutiliza lo mismo del doble click.

---

### 3) `cmdRestaurarAuto_Click` debe limpiar pantalla y dejar nuevo número

**Reemplazá tu evento actual por:**
```vb
Private Sub cmdRestaurarAuto_Click()
    On Error GoTo EH

    LimpiarFormulario True   ' True => limpia + pide siguiente numero
    txtProveedor.SetFocus
    Exit Sub

EH:
    MsgBox "Error al restaurar: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

### 4) Nuevo botón para impresión “parecida a la foto”

Agregá un botón (por ejemplo `cmdImprimirFormato`) y pegá este evento.  
(usa nombres de tus controles reales; abajo puse los más comunes de tu código).

```vb
Private Sub cmdImprimirFormato_Click()
    On Error GoTo EH

    Dim y As Single
    Dim i As Integer
    Dim imp As String

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 9

    y = 400

    Printer.CurrentX = 1200: Printer.CurrentY = y
    Printer.FontBold = True
    Printer.Print "PAGO A PROVEEDORES"
    Printer.FontBold = False

    y = y + 350
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Proveedor: " & txtProveedor.Text
    Printer.CurrentX = 6500: Printer.CurrentY = y: Printer.Print "Fecha: " & txtFecha.Text

    y = y + 350
    Printer.Line (300, y)-(10500, y)
    y = y + 150

    ' DEUDA
    Printer.FontBold = True
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "DEUDA"
    Printer.FontBold = False
    y = y + 220
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Detalle"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (300, y)-(5000, y)
    y = y + 120

    For i = 1 To grdDeuda.Rows - 1
        If Trim$(grdDeuda.TextMatrix(i, 0)) <> "" Then
            Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print Left$(grdDeuda.TextMatrix(i, 0), 30)
            imp = grdDeuda.TextMatrix(i, 1) ' ajusta indice si cambia
            Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print imp
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Total Deuda:"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print txtSubDeuda.Text

    ' CHEQUES
    y = y + 350
    Printer.FontBold = True
    Printer.CurrentX = 5600: Printer.CurrentY = y - 350: Printer.Print "CHEQUES"
    Printer.FontBold = False
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Nro"
    Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print "Banco"
    Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print "Vto"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (5600, y)-(10500, y)
    y = y + 120

    For i = 1 To grdCheques.Rows - 1
        If Trim$(grdCheques.TextMatrix(i, 0)) <> "" Or Trim$(grdCheques.TextMatrix(i, 1)) <> "" Then
            Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 1) ' nro
            Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 0) ' banco
            Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 2) ' vto
            Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 3) ' importe
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Cheques:"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print txtSubCheques.Text

    ' OTROS + TOTALES
    y = y + 300
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Otros: " & txtSubOtros.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Efectivo: " & txtEfectivo.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Retencion IIBB: " & GetRetIIBBText()
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Total Pago: " & txtTotalPago.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Saldo: " & txtSaldo.Text

    y = y + 350
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Recibi la suma de: " & txtImporteLetras.Text

    y = y + 600
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Firma: ____________________________"

    Printer.EndDoc
    MsgBox "Impresion enviada.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir formato: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

Si querés, en el próximo mensaje te paso una **versión 2 del formato de impresión** todavía más parecida al papel (con “cajas”/líneas y mejor alineación), ya ajustada a los **índices reales** de tus grillas.

---

## User

mirá la captura ordenpago-06.png necesito que cuando ponga un importe en el textbox txtefectivo, lo sume al pago y lo reste del saldo. No lo hace

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 26.8s)

_Thinking:_

**Delegating Coding Tasks**

I need to make sure that I don't do the analysis inline since the orchestrator has to handle that. It seems like answering simple questions is fine, but for coding and debugging, I should delegate to a subagent. The user is asking for a specific fix, so I should launch a task to get the patch snippet for them. It’s about getting things done efficiently for the user!

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 10.2s)

Perfecto, ahí está el problema: el `Change` de `txtEfectivo` no estaba recalculando.

Pegá **esto** reemplazando tus eventos:

```vb
Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

```vb
Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

Opcional (recomendado): formatear `txtEfectivo` al salir del campo:

```vb
Private Sub txtEfectivo_LostFocus()
    Dim v As Double

    If m_Cargando Then Exit Sub

    On Error GoTo EH
    If Trim$(txtEfectivo.Text) = "" Then
        v = 0
    Else
        v = CDbl(txtEfectivo.Text)
    End If

    m_Cargando = True
    txtEfectivo.Text = Format$(v, "0.00")
    m_Cargando = False

    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = True
    txtEfectivo.Text = "0.00"
    m_Cargando = False
    RecalcularTodo
End Sub
```

Prueba rápida:
1) cargá deuda,  
2) escribí importe en `txtEfectivo`,  
3) verificá que `txtTotalPago` sube y `txtSaldo` baja en el momento.  

Si querés, después te lo adapto para aceptar coma decimal argentina (`0,00`) sin problemas.

---

## User

Casi quedó perfecto, solo necesitaría que todos los números que representen importes, se formateen a formato europeo, ya sea en grilla, textbox e impresión ejemplo: 150.236,36  Te paso todo el código como quedó con mis modificaciones.  Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             Dim c As Control
             Set c = Me.ActiveControl
                                                                                                                                       
             If Not c Is Nothing Then
                 Select Case TypeName(c)
                     Case "MSFlexGrid", "MshFlexGrid", "VSFlexGrid"
                         ' En grilla se maneja aparte
                     Case Else
                         KeyCode = 0
                         Sendkeys "{TAB}"
                 End Select
             End If
         End If
End Sub

Private Sub grdDeuda_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdDeuda
         End If
     End Sub
                                                                                                                                       
     Private Sub grdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdCheques
         End If
     End Sub
                                                                                                                                       
     Private Sub grdOtros_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdOtros
         End If
     End Sub
Private Sub Form_Load()
    On Error GoTo EH

    m_Cargando = True

    EnsureDatabaseReady
    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.text = ""
    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.text = "0,00"
    SetRetIIBBText "0,00"

    txtSubDeuda.text = "0,00"
    txtSubCheques.text = "0,00"
    txtSubOtros.text = "0,00"
    txtTotalPago.text = "0,00"
    txtSaldo.text = "0,00"
    txtImporteLetras.text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros

    If nuevoNumero Then
        txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    g.TextMatrix(r, c) = Trim$(v)

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
         On Error GoTo EH
                                                                                                                                       
         LimpiarFormulario True   ' True => limpia + pide siguiente numero
         txtProveedor.SetFocus
         Exit Sub
                                                                                                                                       
EH:
         MsgBox "Error al restaurar: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Dim y As Single
    Dim i As Integer
    Dim impi As String

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 9

    y = 400

    Printer.CurrentX = 1200: Printer.CurrentY = y
    Printer.FontBold = True
    Printer.Print "PAGO A PROVEEDORES"
    Printer.FontBold = False

    y = y + 350
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Proveedor: " & txtProveedor.text
    Printer.CurrentX = 6500: Printer.CurrentY = y: Printer.Print "Fecha: " & txtFecha.text

    y = y + 350
    Printer.Line (300, y)-(10500, y)
    y = y + 150

    ' DEUDA
    Printer.FontBold = True
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "DEUDA"
    Printer.FontBold = False
    y = y + 220
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Detalle"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (300, y)-(5000, y)
    y = y + 120

    For i = 1 To grdDeuda.Rows - 1
        If Trim$(grdDeuda.TextMatrix(i, 0)) <> "" Then
            Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print Left$(grdDeuda.TextMatrix(i, 0), 30)
            impi = grdDeuda.TextMatrix(i, 1) ' ajusta indice si cambia
            Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print impi
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Total Deuda:"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print txtSubDeuda.text

    ' CHEQUES
    y = y + 350
    Printer.FontBold = True
    Printer.CurrentX = 5600: Printer.CurrentY = y - 350: Printer.Print "CHEQUES"
    Printer.FontBold = False
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Nro"
    Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print "Banco"
    Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print "Vto"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (5600, y)-(10500, y)
    y = y + 120

    For i = 1 To grdCheques.Rows - 1
        If Trim$(grdCheques.TextMatrix(i, 0)) <> "" Or Trim$(grdCheques.TextMatrix(i, 1)) <> "" Then
            Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 1) ' nro
            Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 0) ' banco
            Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 2) ' vto
            Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 3) ' importe
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Cheques:"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print txtSubCheques.text

    ' OTROS + TOTALES
    y = y + 300
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Otros: " & txtSubOtros.text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Efectivo: " & txtEfectivo.text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Retencion IIBB: " & GetRetIIBBText()
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Total Pago: " & txtTotalPago.text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Saldo: " & txtSaldo.text

    y = y + 350
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Recibi la suma de: " & txtImporteLetras.text

    y = y + 800
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Firma: ____________________________"

    Printer.EndDoc
    MsgBox "Impresion enviada.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir formato: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
Private Function EnLetras(numero As String) As String
    
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        'If Mid(numero, paso, 1) = "." Then
        If Mid(numero, paso, 1) = "," Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
   
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
   
    flag = "N"
    
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
               
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
               
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                   
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
               
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
               
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
               
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
               
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
               
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
               
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
               
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                End If
            End If
            
            If paso = 7 Then
                'MsgBox (Mid(entero, 1, 1))
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
       
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci & "/100"
            Else
                EnLetras = expresion & "con " & deci & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
       
End Function
Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)

    efectivo = ParseCurrency(txtEfectivo.text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    totalPago = subCheques + subOtros + efectivo + retIIBB
    saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.text = FormatMoney(subDeuda)
    txtSubCheques.text = FormatMoney(subCheques)
    txtSubOtros.text = FormatMoney(subOtros)
    txtTotalPago.text = FormatMoney(totalPago)
    txtSaldo.text = FormatMoney(saldo)
    txtImporteLetras.text = EnLetras(CStr(totalPago)) 'NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String

    EnsureDatabaseReady

    nroOrden = CLng(Val(NzS(txtNroOrden.text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!nroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.text)
    rs!Proveedor = NzS(txtProveedor.text)

    rs!TotalDeuda = ParseCurrency(txtSubDeuda.text)
    rs!SubtotalCheques = ParseCurrency(txtSubCheques.text)
    rs!SubtotalOtros = ParseCurrency(txtSubOtros.text)
    rs!efectivo = ParseCurrency(txtEfectivo.text)
    rs!RetencionIIBB = ParseCurrency(GetRetIIBBText())
    rs!totalPago = ParseCurrency(txtTotalPago.text)
    rs!saldo = ParseCurrency(txtSaldo.text)
    rs!MontoLetras = NzS(txtImporteLetras.text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim texto As String

    EnsureDatabaseReady

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        texto = CStr(NzL(rs!nroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!nroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & CStr(Err.Number) & " - " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.text = CStr(NzL(rs!nroOrden))
    txtFecha.text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.text = NzS(rs!Proveedor)

    txtSubDeuda.text = FormatMoney(NzC(rs!TotalDeuda))
    txtSubCheques.text = FormatMoney(NzC(rs!SubtotalCheques))
    txtSubOtros.text = FormatMoney(NzC(rs!SubtotalOtros))
    txtEfectivo.text = FormatMoney(NzC(rs!efectivo))
    SetRetIIBBText FormatMoney(NzC(rs!RetencionIIBB))
    txtTotalPago.text = FormatMoney(NzC(rs!totalPago))
    txtSaldo.text = FormatMoney(NzC(rs!saldo))
    txtImporteLetras.text = NzS(rs!MontoLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub



' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
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

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ParseCurrency = 0
        Exit Function
    End If

    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    ParseCurrency = CCur(Val(t))
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    FormatMoney = Format$(v, "0.00")
    FormatMoney = Replace(FormatMoney, ".", ",")
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal importe As Currency) As String
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(importe)
End Function

' ------------------------------------------------------------
' Helpers de compatibilidad (control opcional txtRetIIBB)
' ------------------------------------------------------------
Private Sub EnsureDatabaseReady()
    On Error GoTo EH

    If BaseSPC Is Nothing Then
        Set BaseSPC = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    End If
    Exit Sub
EH:
    MsgBox "No se pudo abrir la base de datos: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

Private Function TieneControl(ByVal nombreControl As String) As Boolean
    On Error GoTo EH
    Dim ctl As Control
    Set ctl = Me.Controls(nombreControl)
    TieneControl = Not (ctl Is Nothing)
    Exit Function
EH:
    TieneControl = False
End Function

Private Function GetRetIIBBText() As String
    On Error GoTo EH
    If TieneControl("txtRetIIBB") Then
        GetRetIIBBText = CStr(Me.Controls("txtRetIIBB").text)
    Else
        GetRetIIBBText = "0,00"
    End If
    Exit Function
EH:
    GetRetIIBBText = "0,00"
End Function

Private Sub SetRetIIBBText(ByVal valor As String)
    On Error Resume Next
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").text = valor
    End If
End Sub

Private Sub txtEfectivo_LostFocus()
    Dim v As Double

    If m_Cargando Then Exit Sub

    On Error GoTo EH
    If Trim$(txtEfectivo.text) = "" Then
        v = 0
    Else
        v = CDbl(txtEfectivo.text)
    End If

    m_Cargando = True
    txtEfectivo.text = Format$(v, "0.00")
    m_Cargando = False

    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = True
    txtEfectivo.text = "0.00"
    m_Cargando = False
    RecalcularTodo
End Sub 

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 62.7s)

_Thinking:_

**Delegating code guidance**

I need to provide an answer in Spanish for the user. They’ve asked for numeric formatting in the European style everywhere, and I think we can do targeted replacements without needing a tool. However, the orchestrator says I must delegate the real work. So, I’ll pass this task to a subagent for guidance on the code patch. I want to handle this quickly to make sure I’m helping the user effectively!

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 37.6s)

¡Quedó muy bien tu avance! 👏  
Para cerrar lo de formato europeo en **todo** (textbox + grilla + impresión), pegá estos bloques.

---

### 1) Reemplazá tus utilitarios `ParseCurrency` y `FormatMoney` por estos

```vb
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    Dim pDot As Long, pCom As Long, decSep As String
    Dim intPart As String, decPart As String
    Dim sign As Double
    
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseCurrency = 0
        Exit Function
    End If
    
    t = Replace$(t, " ", "")
    t = Replace$(t, "$", "")
    
    sign = 1
    If Left$(t, 1) = "-" Then
        sign = -1
        t = Mid$(t, 2)
    ElseIf Left$(t, 1) = "+" Then
        t = Mid$(t, 2)
    End If
    
    pDot = InStrRev(t, ".")
    pCom = InStrRev(t, ",")
    
    If pDot > 0 And pCom > 0 Then
        If pDot > pCom Then
            decSep = "."
        Else
            decSep = ","
        End If
    ElseIf pDot > 0 Then
        If Len(t) - pDot = 3 Then
            decSep = ""    ' miles
        Else
            decSep = "."   ' decimal
        End If
    ElseIf pCom > 0 Then
        If Len(t) - pCom = 3 Then
            decSep = ""    ' miles
        Else
            decSep = ","   ' decimal
        End If
    Else
        decSep = ""
    End If
    
    If decSep = "" Then
        t = Replace$(t, ".", "")
        t = Replace$(t, ",", "")
        If Len(t) = 0 Then t = "0"
        ParseCurrency = CCur(sign * Val(t))
        Exit Function
    End If
    
    If decSep = "." Then
        intPart = Left$(t, InStrRev(t, ".") - 1)
        decPart = Mid$(t, InStrRev(t, ".") + 1)
    Else
        intPart = Left$(t, InStrRev(t, ",") - 1)
        decPart = Mid$(t, InStrRev(t, ",") + 1)
    End If
    
    intPart = Replace$(intPart, ".", "")
    intPart = Replace$(intPart, ",", "")
    If Len(intPart) = 0 Then intPart = "0"
    
    If Len(decPart) = 0 Then
        decPart = "00"
    ElseIf Len(decPart) = 1 Then
        decPart = decPart & "0"
    ElseIf Len(decPart) > 2 Then
        decPart = Left$(decPart, 2)
    End If
    
    ParseCurrency = CCur(sign * Val(intPart & "." & decPart))
End Function

Private Function GroupThousands(ByVal digits As String) As String
    Dim i As Long, out As String, cnt As Long
    
    out = ""
    cnt = 0
    
    For i = Len(digits) To 1 Step -1
        out = Mid$(digits, i, 1) & out
        cnt = cnt + 1
        If cnt = 3 And i > 1 Then
            out = "." & out
            cnt = 0
        End If
    Next i
    
    GroupThousands = out
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    Dim isNeg As Boolean
    Dim absV As Currency
    Dim cents As Long
    Dim intPart As Long
    Dim decPart As Long
    Dim intTxt As String
    
    isNeg = (v < 0)
    absV = Abs(v)
    
    cents = CLng((absV * 100) + 0.5)
    intPart = cents \ 100
    decPart = cents Mod 100
    
    intTxt = GroupThousands(CStr(intPart))
    
    If isNeg Then
        FormatMoney = "-" & intTxt & "," & Right$("0" & CStr(decPart), 2)
    Else
        FormatMoney = intTxt & "," & Right$("0" & CStr(decPart), 2)
    End If
End Function

Private Sub FormatImporteTextBox(ByRef txt As TextBox)
    txt.Text = FormatMoney(ParseCurrency(txt.Text))
End Sub
```

---

### 2) Corregí `txtRetIIBB_Change` (en tu código quedó sin recalcular)

```vb
Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

---

### 3) LostFocus para todos los importes

```vb
Private Sub txtEfectivo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtEfectivo
    RecalcularTodo
End Sub

Private Sub txtSubDeuda_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubDeuda
End Sub

Private Sub txtSubCheques_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubCheques
End Sub

Private Sub txtSubOtros_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubOtros
End Sub

Private Sub txtTotalPago_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtTotalPago
End Sub

Private Sub txtSaldo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSaldo
End Sub

Private Sub txtRetIIBB_LostFocus()
    If m_Cargando Then Exit Sub
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").Text = FormatMoney(ParseCurrency(Me.Controls("txtRetIIBB").Text))
    End If
    RecalcularTodo
End Sub
```

---

### 4) Formatear importes en grilla al editar celda

En tu `EditarCeldaGrid`, reemplazá la parte final por esto:

```vb
v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
v = Trim$(v)

If g.Name = "grdDeuda" And c = colDeudaImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
ElseIf g.Name = "grdCheques" And c = colChequeImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
ElseIf g.Name = "grdOtros" And c = colOtroImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
Else
    g.TextMatrix(r, c) = v
End If
```

---

### 5) Impresión: siempre formatear con función

Donde imprimís importes, usá patrón:

```vb
Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.Text))
```

Ejemplos concretos para reemplazar en `cmdImprimir_Click`:

```vb
Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.Text))
...
Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print FormatMoney(ParseCurrency(txtSubCheques.Text))
...
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Efectivo: " & FormatMoney(ParseCurrency(txtEfectivo.Text))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Retencion IIBB: " & FormatMoney(ParseCurrency(GetRetIIBBText()))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Total Pago: " & FormatMoney(ParseCurrency(txtTotalPago.Text))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Saldo: " & FormatMoney(ParseCurrency(txtSaldo.Text))
```

Y para grillas:
```vb
impi = FormatMoney(ParseCurrency(grdDeuda.TextMatrix(i, 1)))
...
Printer.Print FormatMoney(ParseCurrency(grdCheques.TextMatrix(i, 3)))
```

---

Si querés, en el próximo paso te paso **tu `cmdImprimir_Click` ya reescrito completo** con estos formatos aplicados línea por línea (copiar/pegar directo).

---

## User

si, por favor pasame el cmdImprimir_Click completo gracias.

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 48.9s)

_Thinking:_

**Delegating task generation**

I need to provide some code, but it seems tools aren't necessary right now. The orchestrator rule suggests I shouldn’t do inline work, so I must delegate all the real work. It looks like I need to use the task tool to generate a complete routine, possibly based on the current code. I’ll focus on that next!

**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 35.4s)

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
    
    Dim deudaRows As Long
    Dim chequeRows As Long
    
    xLeft = 600
    xRight = 7800
    lineH = 240
    pageBottom = 10600
    
    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = False
    
    y = 600
    
    ' Encabezado
    Printer.FontBold = True
    Printer.Print Tab(2); "ORDEN DE PAGO"
    Printer.FontBold = False
    y = y + lineH * 2
    
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Proveedor: " & Trim$(txtProveedor.Text)
    
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print "Fecha: " & Trim$(txtFecha.Text)
    y = y + lineH * 2
    
    ' Deuda
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
    Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.Text))
    y = y + lineH * 2
    
    ' Cheques
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
    Printer.Print FormatMoney(ParseCurrency(txtSubCheques.Text))
    y = y + lineH * 2
    
    ' Otros (retencion / efectivo / subtotal otros)
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
    Printer.Print FormatMoney(ParseCurrency(txtEfectivo.Text))
    y = y + lineH
    
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Otros:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubOtros.Text))
    y = y + lineH * 2
    
    ' Totales
    If y > pageBottom - (lineH * 6) Then
        Printer.NewPage
        y = 600
    End If
    
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "TOTALES"
    Printer.FontBold = False
    y = y + lineH
    
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Total Pago:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtTotalPago.Text))
    y = y + lineH
    
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Saldo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSaldo.Text))
    y = y + lineH * 2
    
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Importe en letras:"
    y = y + lineH
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print Trim$(txtImporteLetras.Text)
    y = y + lineH * 4
    
    ' Firma
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

Si querés, después te paso una variante de este mismo `cmdImprimir_Click` con marco/cajitas más parecido al diseño del comprobante de la foto.

---

## User

ahora te digo creo que lo unico que faltaría sería poner el Nro de Orden de Pago. No lo veo en la impresión

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 25.9s)

_Thinking:_



**Tool: task**

---

## Assistant (Sdd-Orchestrator · GPT-5.3 Codex · 7.9s)

Tal cual, buena observación 👍  
Agregá este bloque en `cmdImprimir_Click` dentro del encabezado:

```vb
' Nro de Orden de Pago visible en encabezado
Printer.CurrentX = 120
Printer.CurrentY = 520
Printer.FontBold = True
Printer.Print "ORDEN DE PAGO Nro: " & Trim$(txtNroOrden.Text)
Printer.FontBold = False
```

Si lo querés en la misma línea que fecha:

```vb
Printer.CurrentX = 120
Printer.CurrentY = 520
Printer.Print "Fecha: " & txtFecha.Text & "    OP Nro: " & Trim$(txtNroOrden.Text)
```

Con eso ya te queda visible en la impresión.

---

