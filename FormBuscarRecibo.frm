VERSION 5.00
Begin VB.Form FormBuscarRecibo 
   Caption         =   "Buscar e-Recibo"
   ClientHeight    =   3090
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5970
   LinkTopic       =   "Form2"
   ScaleHeight     =   3090
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Buscar e-Recibo"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
      Begin VB.PictureBox PictureQP 
         Height          =   495
         Left            =   5160
         ScaleHeight     =   435
         ScaleWidth      =   1035
         TabIndex        =   4
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimirRecibo 
         Caption         =   "Generar e-Recibo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtNroPago 
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.ComboBox cmbSucursales 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   840
         Width           =   2175
      End
   End
End
Attribute VB_Name = "FormBuscarRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit

' --- VARIABLES GLOBALES DEL FORMULARIO ---
Private BaseSPC As Database
Private Const RUTA_DB As String = "c:\TrabajosActivos\SPC-Core\DB_SPC_SI.mdb"

' =============================================================================
' 2. BOTÓN BUSCAR E IMPRIMIR
' =============================================================================
Private Sub cmdImprimirRecibo_Click()

    Call ImprimirReciboE

'    On Error GoTo ErrorHandler
'
'    ' --- Validaciones ---
'    If cmbSucursales.ListIndex = -1 Then
'        MsgBox "Seleccione una sucursal.", vbExclamation
'        Exit Sub
'    End If
'    If Trim(txtNroPago.text) = "" Or Not IsNumeric(txtNroPago.text) Then
'        MsgBox "Ingrese un nÓmero de pago válido.", vbExclamation
'        Exit Sub
'    End If
'
'    ' --- Obtener ID Sucursal y Nro Pago ---
'    Dim IdSucursal As Long
'    Dim NroPago As Long
'    Dim strSucursal As String
'
'    ' Extraer ID del combo (ej: "1 - Quilmes" -> obtiene 1)
'    strSucursal = cmbSucursales.text
'    If InStr(strSucursal, " - ") > 0 Then
'        IdSucursal = Val(Left(strSucursal, InStr(strSucursal, " - ") - 1))
'    Else
'        IdSucursal = Val(strSucursal)
'    End If
'
'    NroPago = Val(txtNroPago.text)
'    ' --- Buscar en Base de Datos ---
'    'Dim rsC As Recordset ' Cabecera
'    'Dim rsD As Recordset ' Detalle
'
'    Dim rsC  ' Cabecera
'    Dim rsD  ' Detalle
'    Dim sqlC As String
'    Dim sqlD As String
'
'    ' 1. Buscar Cabecera
'    sqlC = "SELECT * FROM RecibosC WHERE IdSucursal = " & IdSucursal & " AND NroPago = " & NroPago
'    Set rsC = BaseSPC.OpenRecordset(sqlC, dbOpenSnapshot)
'
'    If rsC.EOF Then
'        MsgBox "No se encontró el Recibo Nº " & NroPago & " en la Sucursal " & IdSucursal, vbInformation
'        rsC.Close
'        Exit Sub
'    End If
'
'    ' 2. Buscar Detalle
'    sqlD = "SELECT * FROM RecibosD WHERE IdSucursal = " & IdSucursal & " AND NroPago = " & NroPago
'    Set rsD = BaseSPC.OpenRecordset(sqlD, dbOpenSnapshot)
'
'    ' --- LLAMAR A IMPRESIÓN ---
'    Call ImprimirReciboDesdeBD(rsC, rsD)
'
'    ' Limpieza
'    rsC.Close
'    rsD.Close
'
'    Exit Sub
'ErrorHandler:
'    MsgBox "Error en el proceso de búsqueda: " & Err.Description, vbCritical


End Sub


' =============================================================================
' 1. CARGA INICIAL Y CONEXIÓN
' =============================================================================
Private Sub Form_Load()
    On Error GoTo ErrorHandler
    
    ' Conectar a la base de datos
    Set BaseSPC = OpenDatabase(RUTA_DB)
    Dim rsSuc
    
    ' Cargar Sucursales en el Combo
    Set rsSuc = BaseSPC.OpenRecordset("SELECT IdSucursal, NombreSucursal FROM Sucursales ORDER BY IdSucursal", dbOpenSnapshot)
    
    cmbSucursales.Clear
    Do While Not rsSuc.EOF
        ' Formato: 1 - Quilmes
        cmbSucursales.AddItem rsSuc!IdSucursal & " - " & rsSuc!NombreSucursal
        rsSuc.MoveNext
    Loop
    rsSuc.Close
    
    Exit Sub
ErrorHandler:
    MsgBox "Error al cargar sucursales: " & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cerrar la base al salir para no dejar conexiones abiertas
    If Not BaseSPC Is Nothing Then BaseSPC.Close
End Sub


' =============================================================================
' 3. RUTINA DE IMPRESIÓN (DISEÑO EXACTO A LA FOTO)
' =============================================================================
'Private Sub ImprimirReciboDesdeBD(rsC As Recordset, rsD As Recordset)
Private Sub ImprimirReciboDesdeBD(rsC, rsD)
    On Error GoTo CapturaErrores
    
    Dim tClientes
    Dim tDomicilios
    
    ' Variables de Datos
    Dim NroReciboStr As String, IdSucStr As String
    Dim vFecha As Date
    Dim vIdCliente As Long
    Dim vTotalRecibo As Double
    Dim vConcepto As String
    Dim vClienteCUIT As String
    Dim I As Integer, Largo As Integer
    
    ' Variables para Acumulados (Sumas)
    Dim vTotalEfectivo As Double
    Dim vTotalCheques As Double
    Dim vTotalTransf As Double
    Dim vTotalRetenc As Double
    Dim FormaPagoStr As String
    
    ' --- A. LEER DATOS ---
    vFecha = rsC!FechaPago
    vIdCliente = rsC!IdCliente
    vTotalRecibo = Val("" & rsC!TotalAbonado)
    vConcepto = "" & rsC!Corresponde
    NroReciboStr = CStr(rsC!NroPago)
    IdSucStr = CStr(rsC!IdSucursal)
    
    ' --- B. CALCULAR TOTALES DESDE EL DETALLE ---
    vTotalEfectivo = 0: vTotalCheques = 0: vTotalTransf = 0: vTotalRetenc = 0
    
    If Not rsD.EOF Then
        rsD.MoveFirst
        Do While Not rsD.EOF
            FormaPagoStr = UCase("" & rsD!FormaPago)
            
            ' Lógica de suma según el texto en la base de datos
            If InStr(FormaPagoStr, "EFECTIVO") > 0 Then
                vTotalEfectivo = vTotalEfectivo + Val(rsD!ImportePago)
            ElseIf InStr(FormaPagoStr, "CHEQUE") > 0 Then
                vTotalCheques = vTotalCheques + Val(rsD!ImportePago)
            ElseIf InStr(FormaPagoStr, "TRANSF") > 0 Then
                vTotalTransf = vTotalTransf + Val(rsD!ImportePago)
            ElseIf InStr(FormaPagoStr, "RETEN") > 0 Or InStr(FormaPagoStr, "PERCEP") > 0 Then
                vTotalRetenc = vTotalRetenc + Val(rsD!ImportePago)
            Else
                ' Por defecto a cheques/varios si no se reconoce
                vTotalCheques = vTotalCheques + Val(rsD!ImportePago)
            End If
            rsD.MoveNext
        Loop
    End If
    
    ' --- C. ABRIR TABLAS MAESTRAS ---
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    Set tDomicilios = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
    tClientes.Index = "PrimaryKey"
    tDomicilios.Index = "PrimaryKey"

    ' --- D. DIBUJAR EN IMPRESORA ---
    ' Configurar Impresora PDF si existe
    For I = 0 To Printers.Count - 1
        If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
    Next
    
    Printer.ScaleMode = 6 ' Milímetros (Facilita el diseño)
    Printer.ScaleHeight = 297
    Printer.ScaleWidth = 210
    Printer.FontName = "Arial"
        
        ' 1. LOGO
        On Error Resume Next
        PictureQP.PaintPicture LoadPicture(App.Path & "\Quilplac.JPG"), 10, 5, 40, 15
        On Error GoTo CapturaErrores
        
    ' 2. CABECERA DERECHA
    Printer.CurrentX = 10: Printer.CurrentY = 2: Printer.FontSize = 10: Printer.FontBold = False
    Printer.Print "ORIGINAL"
        
    Printer.CurrentX = 90: Printer.CurrentY = 15: Printer.FontSize = 14: Printer.FontBold = True
    Printer.Print "RECIBO OFICIAL"
        
        ' Formateo de Ceros (0001-00001234)
        Largo = 8 - Len(NroReciboStr): For I = 1 To Largo: NroReciboStr = "0" & NroReciboStr: Next I
        Largo = 4 - Len(IdSucStr): For I = 1 To Largo: IdSucStr = "0" & IdSucStr: Next I
        
    Printer.FontSize = 12: Printer.FontBold = True
    Printer.CurrentX = 150: Printer.CurrentY = 10
    Printer.Print IdSucStr & "-" & NroReciboStr
        
    Printer.FontSize = 10: Printer.FontBold = True
    Printer.CurrentX = 150: Printer.CurrentY = 16
    Printer.Print "Fecha: " & Format(vFecha, "DD/MM/YYYY")
        
    Printer.FontSize = 8: Printer.FontBold = False
    Printer.CurrentX = 150: Printer.CurrentY = 22: Printer.Print "C.U.I.T NÂ­ 30-70843254-3"
    Printer.CurrentX = 150: Printer.CurrentY = 26: Printer.Print "Ing.Brutos NÂ­ 30-70843254-3"
    Printer.CurrentX = 150: Printer.CurrentY = 30: Printer.Print "Inicio de Actividades: 11-06-2003"
    Printer.CurrentX = 150: Printer.CurrentY = 34: Printer.Print "I.V.A. Responsable Inscripto"
        
    ' 3. DATOS QUILPLAC
    Printer.CurrentX = 10: Printer.CurrentY = 25: Printer.FontSize = 9: Printer.FontBold = True
    Printer.Print "QUILPLAC S.A."
    Printer.FontBold = False
    Printer.CurrentX = 10: Printer.CurrentY = 30: Printer.Print "AndrÂ¨s Baranda 520 - CP (1878) - Quilmes"
    Printer.CurrentX = 10: Printer.CurrentY = 34: Printer.Print "Pcia. Buenos Aires"
    Printer.CurrentX = 10: Printer.CurrentY = 38: Printer.Print "Tel. 4257-5875"
        
    ' LÂ­nea separadora doble
    Printer.DrawWidth = 2
    Printer.Line (10, 42)-(200, 42)
    Printer.Line (10, 43)-(200, 43)
    
    ' 4. CAJA CLIENTE
    Printer.Line (10, 45)-(200, 65), , B
        
    tClientes.Seek "=", vIdCliente
    If Not tClientes.NoMatch Then
        Printer.CurrentX = 12: Printer.CurrentY = 47: Printer.FontSize = 9: Printer.FontBold = True
        Printer.Print "SeÂ±or(es): ";
        Printer.FontBold = False: Printer.Print tClientes!RazonSocial
            
            tDomicilios.Seek "=", vIdCliente
            If Not tDomicilios.NoMatch Then
                Printer.CurrentX = 12: Printer.CurrentY = 53: Printer.FontBold = True
                Printer.Print "Domicilio: ";
                Printer.FontBold = False: Printer.Print tDomicilios!Domicilio
                
                Printer.CurrentX = 12: Printer.CurrentY = 59: Printer.FontBold = True
                Printer.Print "Localidad: ";
                Printer.FontBold = False: Printer.Print tDomicilios!localidad
            End If
            
            Printer.CurrentX = 12: Printer.CurrentY = 65: Printer.FontBold = True
            Printer.Print "I.V.A: ";
            Printer.FontBold = False: Printer.Print "" & tClientes!condicionIva
            
            ' Columna Derecha Cliente
            Printer.CurrentX = 130: Printer.CurrentY = 47: Printer.FontBold = True
            Printer.Print "C.U.I.T NÂ§: ";
            Printer.FontBold = False: Printer.Print "" & tClientes!CUIT
            
            Printer.CurrentX = 130: Printer.CurrentY = 53: Printer.FontBold = True
            Printer.Print "TelÂ¡fono: ";
            Printer.FontBold = False: Printer.Print "" & tClientes!Tel
        End If
        
        ' 5. WEB Y DETALLE
        Printer.Line (10, 68)-(200, 73), , B
        Printer.CurrentX = 80: Printer.CurrentY = 69: Printer.FontSize = 9: Printer.FontBold = True
        Printer.Print "*** www.quilplac.com ***"
        
        Printer.Line (10, 75)-(200, 80), , B
        Printer.CurrentX = 85: Printer.CurrentY = 76
        Printer.Print "DETALLE DEL RECIBO"
        
        ' Recuadro del Cuerpo Principal
        Printer.Line (10, 80)-(200, 200), , B
        
        ' 6. FORMAS DE PAGO (Posiciones fijas según foto)
        Dim YPos As Double
        YPos = 90
        Printer.FontSize = 10: Printer.FontBold = True
        
        ' Efectivo
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.Print "* Efectivo:"
        Printer.CurrentX = 80: Printer.CurrentY = YPos: Printer.FontBold = False
        Printer.Print "$ " & Format(vTotalEfectivo, "#,##0.00")
        
        ' Transferencia
        YPos = YPos + 8
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.FontBold = True: Printer.Print "* Transferencia:"
        Printer.CurrentX = 80: Printer.CurrentY = YPos: Printer.FontBold = False
        Printer.Print "$ " & Format(vTotalTransf, "#,##0.00")
        
        ' Cheques
        YPos = YPos + 8
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.FontBold = True: Printer.Print "* Cheques Varios:"
        Printer.CurrentX = 80: Printer.CurrentY = YPos: Printer.FontBold = False
        Printer.Print "$ " & Format(vTotalCheques, "#,##0.00")
        
        ' Retenciones
        YPos = YPos + 8
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.FontBold = True: Printer.Print "* Retenciones:"
        Printer.CurrentX = 80: Printer.CurrentY = YPos: Printer.FontBold = False
        Printer.Print "$ " & Format(vTotalRetenc, "#,##0.00")
        
        ' 7. GRILLA INFERIOR (Facturas)
        YPos = 130
        Printer.FontSize = 9: Printer.FontBold = True
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.Print "Fecha"
        Printer.CurrentX = 65: Printer.CurrentY = YPos: Printer.Print "Nro. Factura"
        Printer.CurrentX = 110: Printer.CurrentY = YPos: Printer.Print "Importe"
        
        Printer.Line (30, YPos + 4)-(150, YPos + 4)
        
        YPos = 136
        Printer.FontBold = False
        Printer.CurrentX = 30: Printer.CurrentY = YPos: Printer.Print Format(vFecha, "DD/MM/YYYY")
        Printer.CurrentX = 65: Printer.CurrentY = YPos: Printer.Print "VARIAS"
        Printer.CurrentX = 110: Printer.CurrentY = YPos: Printer.Print "0,00"
        
        Printer.CurrentX = 15: Printer.CurrentY = 160: Printer.FontBold = True
        Printer.Print StrConv(vConcepto, vbUpperCase)
        
        ' 8. PIE DE PÁGINA (Totales)
        Printer.Line (10, 200)-(200, 200) ' Linea superior pie
        
        ' Etiqueta Recibimos Pesos (Fondo negro)
        Printer.Line (10, 202)-(50, 208), vbBlack, BF
        Printer.CurrentX = 12: Printer.CurrentY = 203: Printer.FontSize = 9: Printer.FontBold = True: Printer.ForeColor = vbWhite
        Printer.Print "RECIBIMOS PESOS:"
        
        ' Texto en Letras
        Printer.ForeColor = vbBlack: Printer.FontSize = 10: Printer.FontBold = False
        Dim vLetras As String
        vLetras = EnLetras(CStr(Format(vTotalRecibo, "Fixed")))
        Printer.CurrentX = 12: Printer.CurrentY = 210
        Printer.Print StrConv(vLetras, vbUpperCase)
        
        ' CAJA TOTAL NEGRA
        Dim BoxTop As Double, BoxLeft As Double
        BoxTop = 215: BoxLeft = 140
        Printer.Line (BoxLeft, BoxTop)-(200, BoxTop + 10), vbBlack, BF
        
        Printer.CurrentX = BoxLeft + 2: Printer.CurrentY = BoxTop + 2
        Printer.ForeColor = vbWhite: Printer.FontSize = 12: Printer.FontBold = True
        Printer.Print "TOTAL:"
        
    Printer.CurrentX = BoxLeft + 30: Printer.CurrentY = BoxTop + 2
    Printer.Print "$ " & Format(vTotalRecibo, "#,##0.00")
    
    Printer.EndDoc
    
    tClientes.Close
    tDomicilios.Close
    Exit Sub
    
CapturaErrores:
    MsgBox "Error imprimiendo: " & Err.Description, vbCritical
End Sub

Private Sub ImprimirReciboE()
    'On Error GoTo CapturaErrores
   
    Dim NroRecibo As String
    Dim IdSuc As String
    Dim sqlD As String
    Dim Largo As Integer, LargoSuc As Integer
    Dim TotalFac As Double, SubTotalFac As Double
    Dim vEfete As Double, vCheques As Double, vRetenciones As Double, vTransf As Double
    Dim CUIT As String, Texto As String
    Dim I As Integer, J As Integer
    Dim strSucursal As String, IdSucursal As Long, NroPago As Long
    
    ' --- 1. PEDIDO MANUAL DE DATOS (Parte HÍbrida) ---
    Dim vCantFac As Integer
    Dim vFF(6) As Date
    Dim vNFac(6) As String
    Dim vImpF(6) As Double
    Dim vConcepto As String
    
    vCantFac = CInt(InputBox("Ingrese Cantidad de Facturas a Detallar (Máx 6)", "DETALLE MANUAL"))
    
    If vCantFac > 6 Then
        MsgBox "Corrija la Cantidad. Máximo 6.", vbExclamation
        vCantFac = InputBox("Ingrese Cantidad (Máx 6)", "DETALLE MANUAL")
    End If
       
    For J = 1 To vCantFac
        vFF(J) = InputBox("Fecha Factura #" & J, "Dato Manual")
        vNFac(J) = InputBox("Nro Factura #" & J, "Dato Manual")
        vImpF(J) = Val(InputBox("Importe Factura #" & J, "Dato Manual"))
    Next J
       
    vConcepto = InputBox("Ingrese El Concepto General", "Dato Manual")
       
    ' --- 2. CONEXIÓN A BASE DE DATOS ---
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
    Set tRecibosC = BaseSPC.OpenRecordset("RecibosC", dbOpenTable)
    Set tRecibosD = BaseSPC.OpenRecordset("RecibosD", dbOpenTable) ' Para Formas de Pago
    
    tClientes.Index = "PrimaryKey"
    tDomiciliosClientes.Index = "PrimaryKey"
    tRecibosC.Index = "PrimaryKey"
    
    tRecibosC.Seek "=", IdSucursal, NroPago
    
    If Not tRecibosC.NoMatch Then
        vFechaPago = tRecibosC!FechaPago
    End If

    ' --- 3. DATOS CABECERA DESDE PANTALLA ---
    strSucursal = cmbSucursales.text ' OJO: Verifica si tu combo se llama cmbSucursal o Combo1
    If InStr(strSucursal, " - ") > 0 Then
        IdSucursal = Val(Left(strSucursal, InStr(strSucursal, " - ") - 1))
    Else
        IdSucursal = Val(strSucursal)
    End If
    
    NroPago = Val(txtNroPago.text) ' Verifica si es Text1 o txtNroPago

    ' Buscar Cabecera (Opcional si ya la tienes, pero util para validar)
    ' tRecibosC.Seek "=", IdSucursal, NroPago ...

    ' --- 4. IMPRESIÓN ---
    ' Configurar Impresora
    For I = 0 To Printers.Count - 1
        If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
    Next
    
    Printer.ScaleHeight = 297
    Printer.ScaleWidth = 210
    
    ' Logo
   ' On Error Resume Next
    Printer.PaintPicture LoadPicture(App.Path & "\Quilplac.JPG"), 10, 5, 40, 15
   ' On Error GoTo CapturaErrores
 
    ' Encabezado
    Printer.DrawWidth = 10
    Printer.Line (10, 7)-(200, 7)
        
        Printer.CurrentX = 85: Printer.CurrentY = 14: Printer.Font = "Arial": Printer.FontSize = 12: Printer.FontBold = True
        Printer.Print "RECIBO OFICIAL"
        Printer.CurrentX = 15: Printer.CurrentY = 2: Printer.FontSize = 12: Printer.FontBold = False
        Printer.Print "ORIGINAL"
        
        ' Número Recibo
        Printer.FontSize = 12: Printer.CurrentY = 9: Printer.CurrentX = 150
        NroRecibo = CStr(NroPago)
        Largo = 8 - Len(NroRecibo)
        For I = 1 To Largo: NroRecibo = "0" & NroRecibo: Next I
        
        IdSuc = CStr(IdSucursal)
        LargoSuc = 4 - Len(IdSuc)
        For J = 1 To LargoSuc: IdSuc = "0" & IdSuc: Next J
        
        Printer.Print IdSuc & "-" & NroRecibo
        
        ' Fechas y Datos Fiscales
        Printer.CurrentX = 150: Printer.CurrentY = Printer.CurrentY + 2: Printer.FontSize = 12
        Printer.Print "Fecha: " & Format(vFechaPago, "DD/MM/YYYY")
        
        Printer.CurrentX = 150: Printer.CurrentY = Printer.CurrentY + 2: Printer.FontSize = 9: Printer.FontBold = False
        Printer.Print "C.U.I.T NÂ§ 30-70843254-3"
        Printer.CurrentX = 150: Printer.Print "Ing.Brutos NÂ§ 30-70843254-3"
        Printer.CurrentX = 150: Printer.Print "Inicio de Actividades: 11-06-2003"
        Printer.CurrentX = 150: Printer.Print "I.V.A. Responsable Inscripto"
        
        Printer.Line (10, 42)-(200, 42)
        
        ' Datos Quilplac
        Printer.CurrentX = 12: Printer.CurrentY = 20: Printer.Font = "Arial": Printer.FontSize = 10: Printer.FontBold = True
        Printer.Print "QUILPLAC S.A."
        Printer.CurrentX = 12: Printer.Print "AndrÂ¨s Baranda 520 - CP (1878) - Quilmes"
        Printer.CurrentX = 12: Printer.Print "Pcia. Buenos Aires"
        Printer.CurrentX = 12: Printer.Print "Tel. 4257-5875"

        ' --- RECUADRO CLIENTE ---
        Printer.DrawWidth = 10
        Printer.Line (10, 47)-(200, 47)
        Printer.Line (10, 47)-(10, 75)
        Printer.Line (200, 47)-(200, 75)
        Printer.Line (10, 75)-(200, 75)
            
        ' Buscar Cliente en BD
        tClientes.Seek "=", TextCodigoCliente.text
        If Not tClientes.NoMatch Then
            Printer.CurrentX = 15: Printer.CurrentY = 48: Printer.FontSize = 10: Printer.FontBold = True
            Printer.Print "SeÂ±or(es): "
            Printer.CurrentX = 35: Printer.CurrentY = 48: Printer.FontBold = False
            Printer.Print tClientes!RazonSocial
            
            Printer.CurrentX = 130: Printer.CurrentY = 48: Printer.FontBold = True
            Printer.Print "C.U.I.T NÂ§:"
            Printer.CurrentX = 150: Printer.CurrentY = 48: Printer.FontBold = False
            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
            Printer.Print CUIT
                
            tDomiciliosClientes.Seek "=", tClientes!IdCliente
            If Not tDomiciliosClientes.NoMatch Then
                Printer.CurrentX = 15: Printer.CurrentY = 55: Printer.FontBold = True
                Printer.Print "Domicilio: "
                Printer.CurrentX = 35: Printer.CurrentY = 55: Printer.FontBold = False
                Printer.Print tDomiciliosClientes!Domicilio
                
                Printer.CurrentX = 15: Printer.CurrentY = 62: Printer.FontBold = True
                Printer.Print "Localidad: "
                Printer.CurrentX = 35: Printer.CurrentY = 62: Printer.FontBold = False
                Printer.Print tDomiciliosClientes!localidad
            End If
            
            Printer.CurrentX = 130: Printer.CurrentY = 62: Printer.FontBold = True
            Printer.Print "TelÂ¡fono: "
            Printer.CurrentX = 150: Printer.CurrentY = 62: Printer.FontBold = False
            Printer.Print tClientes!Tel
            
            Printer.CurrentX = 15: Printer.CurrentY = 69: Printer.FontBold = True
            Printer.Print "I.V.A: "
            Printer.CurrentX = 35: Printer.CurrentY = 69: Printer.FontBold = False
            ' Asegurate de tener la funcion BuscarCondicionIva, sino usa tClientes!condicionIva directo
            Printer.Print tClientes!condicionIva
        End If ' <--- ESTE FALTABA: CIERRE DEL IF CLIENTE
        
        ' Separadores
        Printer.Line (10, 78)-(200, 78)
        Printer.Line (10, 85)-(200, 85)
        Printer.CurrentX = 83: Printer.CurrentY = 80: Printer.FontSize = 10: Printer.FontBold = True
        Printer.Print "*** www.quilplac.com ***"

        ' --- CUERPO DETALLE ---
        Printer.Line (10, 90)-(200, 90)
        Printer.Line (10, 240)-(200, 240)
        Printer.Line (10, 90)-(10, 240)
        Printer.Line (200, 90)-(200, 240)
        Printer.Line (10, 97)-(200, 97)
        
        Printer.CurrentX = 86: Printer.CurrentY = 92: Printer.FontSize = 10
        Printer.Print "DETALLE DEL RECIBO"
        
        ' --- A. IMPRIMIR FORMAS DE PAGO (Desde Base de Datos RecibosD) ---
        sqlD = "SELECT * FROM RecibosD WHERE IdSucursal = " & IdSucursal & " AND NroPago = " & NroPago
        Set rsD = BaseSPC.OpenRecordset(sqlD, dbOpenDynaset)
        
        Dim PosY As Integer
        PosY = 110
        
        If Not rsD.EOF Then
            rsD.MoveFirst
            Do While Not rsD.EOF ' Agregamos bucle por si hay mas de un pago
                Select Case rsD!FormaPago
                    Case "Efectivo"
                        Printer.CurrentX = 32: Printer.CurrentY = PosY
                        vEfete = Val(rsD!ImportePago)
                        Printer.Print "* Efectivo: " & Chr(9) & Chr(9) & Format(vEfete, "Currency")
                        PosY = PosY + 5
                        
                    Case "Transferencia"
                        Printer.CurrentX = 32: Printer.CurrentY = PosY
                        vTransf = Val(rsD!ImportePago)
                        Printer.Print "* Transferencia: " & Chr(9) & Format(vTransf, "Currency")
                        PosY = PosY + 5
                        
                    Case "Cheque"
                        Printer.CurrentX = 32: Printer.CurrentY = PosY
                        vCheques = Val(rsD!ImportePago)
                        Printer.Print "* Cheques Varios: " & Chr(9) & Format(vCheques, "Currency")
                        PosY = PosY + 5
                        
                    Case "Retencion"
                        Printer.CurrentX = 32: Printer.CurrentY = PosY
                        vRetenciones = Val(rsD!ImportePago)
                        Printer.Print "* Retenciones: " & Chr(9) & Format(vRetenciones, "Currency")
                        PosY = PosY + 5
                End Select
                rsD.MoveNext
            Loop
        End If
        ' Cerramos el tema BD Formas de Pago, ahora vamos a lo manual
                                            
        ' --- B. IMPRIMIR FACTURAS (Desde InputBox Manual - Arrays vFF, vNFac) ---
        Printer.FontSize = 8
        Printer.CurrentX = 32: Printer.CurrentY = 150: Printer.FontUnderline = True: Printer.Print "Fecha"
        Printer.CurrentX = 60: Printer.CurrentY = 150: Printer.FontUnderline = True: Printer.Print "Nro. Factura"
        Printer.CurrentX = 90: Printer.CurrentY = 150: Printer.FontUnderline = True: Printer.Print "Importe"
        Printer.FontUnderline = False
        
        For J = 1 To vCantFac
            ' Usamos Offset para las lineas (5mm entre cada una)
            Dim YLine As Integer
            YLine = 155 + ((J - 1) * 5)
            
            Printer.CurrentX = 32: Printer.CurrentY = YLine
            Printer.Print Format(vFF(J), "DD/MM/YYYY")
            
            Printer.CurrentX = 60: Printer.CurrentY = YLine
            Printer.Print vNFac(J)
            
            Printer.CurrentX = 90: Printer.CurrentY = YLine
            Printer.Print Format(vImpF(J), "##,##0.00")
            
            ' Sumamos al Total Final
            TotalFac = TotalFac + vImpF(J)
        Next J
                
        ' Concepto Manual
        Printer.CurrentX = 20: Printer.CurrentY = 185: Printer.Font = "Arial": Printer.FontSize = 10
        Printer.Print StrConv(vConcepto, vbUpperCase)
        
        ' --- PIE DE PAGINA Y TOTALES ---
        Printer.Line (130, 240)-(130, 262)
        Printer.Line (200, 240)-(200, 262)
        Printer.Line (130, 240)-(130, 262) ' Repetida por seguridad grafica
        
        ' Subtotal (Suma de pagos o facturas, según prefieras. Aquí uso Facturas manuales)
        vSubTotal = TotalFac
        Printer.CurrentX = 135: Printer.CurrentY = 245: Printer.FontName = "Arial": Printer.FontSize = 10
        Printer.CurrentX = 165: Printer.CurrentY = 245
        Printer.Print Format(vSubTotal, "Currency")
        
        ' Total Final Fondo Negro
        Printer.Line (130, 262)-(200, 270), vbBlack, BF
        Printer.CurrentX = 135: Printer.CurrentY = 264: Printer.Font = "Arial": Printer.FontSize = 12: Printer.ForeColor = vbWhite
        Printer.Print "TOTAL: "
        Printer.CurrentX = 165: Printer.CurrentY = 264
        Printer.Print Format(TotalFac, "Currency")
        
        ' Letras y Pie
        Printer.FontSize = 10: Printer.CurrentX = 15: Printer.CurrentY = 245: Printer.ForeColor = vbBlack
        Printer.Line (10, 245)-(48, 250), vbBlack, BF
        Printer.CurrentX = 12: Printer.CurrentY = 245: Printer.ForeColor = vbWhite
        Printer.Print "RECIBIMOS PESOS: "
        
        Printer.ForeColor = vbBlack
        Dim vImporteEnLetras As String
        vImporteEnLetras = EnLetras(CStr(Format(TotalFac, "Fixed")))
        
        Printer.CurrentX = 12: Printer.CurrentY = 253
        If Len(vImporteEnLetras) <= 50 Then
            Printer.Print StrConv(vImporteEnLetras, vbUpperCase)
        Else
            ' Corte simple
             Printer.Print StrConv(Left(vImporteEnLetras, 50), vbUpperCase)
             Printer.CurrentX = 12: Printer.CurrentY = 258
             Printer.Print StrConv(Mid(vImporteEnLetras, 51), vbUpperCase)
        End If

        Printer.EndDoc

Exit Sub

CapturaErrores:
    MsgBox "Error: " & Err.Description
End Sub
Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
End Function
' =============================================================================
' 4. FUNCIÓN AUXILIAR: NUMEROS A LETRAS (Simple)
' =============================================================================
Public Function EnLetras(numero As String) As String
    
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



