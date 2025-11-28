VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormLibroIvaVentas 
   Caption         =   "Libro IVA Ventas"
   ClientHeight    =   8145
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14580
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   14580
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   6960
      Width           =   14295
      Begin VB.CommandButton cmdSIAP 
         Caption         =   "S.I.A.&P"
         Height          =   510
         Left            =   6480
         TabIndex        =   16
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton CmdExportarTXT 
         Caption         =   "&Exportar TXT"
         Height          =   510
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   510
         Left            =   12480
         TabIndex        =   7
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton BotonCancelar 
         Caption         =   "&Cancelar"
         Height          =   510
         Left            =   8520
         TabIndex        =   5
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Height          =   510
         Left            =   10560
         TabIndex        =   6
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton BotonBuscar 
         Caption         =   "&Buscar"
         Height          =   510
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1470
      End
      Begin VB.CommandButton BotonExportar 
         Caption         =   "E&xportar al Excel"
         Height          =   510
         Left            =   4560
         TabIndex        =   4
         Top             =   360
         Width           =   1470
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5895
      Left            =   120
      TabIndex        =   11
      Top             =   1080
      Width           =   14295
      Begin VB.TextBox txtImporteTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   11880
         TabIndex        =   14
         Top             =   5280
         Width           =   1575
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   4935
         Left            =   240
         TabIndex        =   12
         Top             =   240
         Width           =   13785
         _ExtentX        =   24315
         _ExtentY        =   8705
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Importe Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   10320
         TabIndex        =   15
         Top             =   5280
         Width           =   1455
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   14295
      Begin VB.TextBox TextFechaDesde 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4920
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TextFechaHasta 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3480
         TabIndex        =   10
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   8040
         TabIndex        =   9
         Top             =   480
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FormLibroIvaVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstCliente As DAO.Recordset
 Dim rstFacturaD As DAO.Recordset


Private Sub GenerarArchivoAlicuotas()

  'Genero el archivo para SIAP de Alicuotas
    Dim NombreArchivo, linea As String
    Dim NroFac As String
    Dim Tt As String
    Dim TotalFac() As String
    Dim totalIva() As String
    Dim TtFac, TtIVA As String
    Dim rstSIAP, vSQL
    Dim Desde, Hasta
    
    'Variable para usar WSH
    Dim Wscript As Object
      
    'Creamos la referencia para usar Windows Scripting Host
    Set Wscript = CreateObject("WScript.Shell")
    
                
    NombreArchivo = Wscript.SpecialFolders("Desktop") & "\VentasCITIAlicuota_" & Format(TextFechaDesde.text, "yyyy") & Format(TextFechaDesde.text, "mm") & Format(TextFechaDesde.text, "dd")
    NombreArchivo = NombreArchivo & "_" & Format(TextFechaHasta.text, "yyyy") & Format(TextFechaHasta.text, "mm") & Format(TextFechaHasta.text, "dd") & ".txt"
        
    'MsgBox (NombreArchivo)
    Desde = "#" & Format(TextFechaDesde.text, "mm/dd/yyyy") & "#"
    Hasta = "#" & Format(TextFechaHasta.text, "mm/dd/yyyy") & "#"
    
    vSQL = "SELECT * FROM FacturaC WHERE FechaFactura >=" & Desde & " AND FechaFactura <=" & Hasta & " Order By TipoFactura, FechaFactura, NroFactura"
            
    'MsgBox (vSQL)
    
    Set rstSIAP = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    If Not rstSIAP.EOF Then
        rstSIAP.MoveFirst
     Else
        MsgBox ("No Hay registros en el archivo")
        Exit Sub
    End If
    
    Open NombreArchivo For Output As #1
        
    While Not rstSIAP.EOF
        'Punto 01 Tipo de Comprobante
             If rstSIAP!tipofactura = "A" Then linea = linea & "001"
             If rstSIAP!tipofactura = "B" Then linea = linea & "006"
        
        'Punto 02 Punto de venta
             linea = linea & "00003"
             
        'Punto 03 Numero de comprobante
             NroFac = rstSIAP!NroFactura
             
             For I = 1 To (20 - Len(NroFac))
                 NroFac = "0" & NroFac
             Next I
             
             linea = linea & NroFac

        'Punto 04 Importe Neto Gravado
             Tt = Format(CStr(rstSIAP!SubTotalFactura), "#0.00")
             TotalFac = Split(Tt, ",", -1)
             TtFac = TotalFac(0)
             For I = 1 To (13 - Len(TotalFac(0)))
                 TtFac = "0" & TtFac
             Next I
             
             TtFac = TtFac & TotalFac(1)
             
             linea = linea & TtFac
        
        'Punto 05 Alicuota de IVA
            linea = linea & "0005"
        
        'Punto 06 Impuesto liquidado
            Tt = Format(CStr(rstSIAP!totalIva), "#0.00")
            totalIva = Split(Tt, ",", -1)
            TtIVA = totalIva(0)
            For I = 1 To (13 - Len(totalIva(0)))
                TtIVA = "0" & TtIVA
            Next I
            
            TtIVA = TtIVA & totalIva(1)
            
            linea = linea & TtIVA

        Print #1, linea
        linea = ""

        rstSIAP.MoveNext
    Wend
    
    rstSIAP.Close
    
 'Ahora las NOTA DE CREDITO
    
    vSQL = "SELECT * FROM NotaCreditoC WHERE FechaNotaCredito >=" & Desde & " AND FechaNotaCredito <=" & Hasta & " Order By TipoNotaCredito, FechaNotaCredito, NroNotaCredito"
            
    'MsgBox (vSQL)
    
    Set rstSIAP = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    If Not rstSIAP.EOF Then
        rstSIAP.MoveFirst
     Else
        MsgBox ("No Hay registros en el archivo")
        Exit Sub
    End If
    
    'Open NombreArchivo For Output As #1
        
    While Not rstSIAP.EOF
        'Punto 01 Tipo de Comprobante
             If rstSIAP!TipoNotaCredito = "A" Then linea = linea & "003"
             If rstSIAP!TipoNotaCredito = "B" Then linea = linea & "008"
        
        'Punto 02 Punto de venta
             linea = linea & "00003"
             
        'Punto 03 Numero de comprobante
             NroFac = rstSIAP!NroNotaCredito
             
             For I = 1 To (20 - Len(NroFac))
                 NroFac = "0" & NroFac
             Next I
             
             linea = linea & NroFac

        'Punto 04 Importe Neto Gravado
             Tt = Format(CStr(rstSIAP!SubTotalNotaCredito), "#0.00")
             TotalFac = Split(Tt, ",", -1)
             TtFac = TotalFac(0)
             For I = 1 To (13 - Len(TotalFac(0)))
                 TtFac = "0" & TtFac
             Next I
             
             TtFac = TtFac & TotalFac(1)
             
             linea = linea & TtFac
        
        'Punto 05 Alicuota de IVA
            linea = linea & "0005"
        
        'Punto 06 Impuesto liquidado
            Tt = Format(CStr(rstSIAP!totalIva), "#0.00")
            totalIva = Split(Tt, ",", -1)
            TtIVA = totalIva(0)
            For I = 1 To (13 - Len(totalIva(0)))
                TtIVA = "0" & TtIVA
            Next I
            
            TtIVA = TtIVA & totalIva(1)
            
            linea = linea & TtIVA

        Print #1, linea
        linea = ""

        rstSIAP.MoveNext
    Wend
    
    Close #1
    
    MsgBox ("Archivo " & NombreArchivo & " Generado con Exito")

End Sub

Private Function DatosClientes(Cliente As Long) As String
    
    Dim docCli, Nombre As String
    
    Set tClientes = db.OpenRecordset("Clientes", dbOpenTable)
    tClientes.Index = "PrimaryKey"
    
    tClientes.Seek "=", Cliente
    
    If Not tClientes.NoMatch Then
    
        If Len(tClientes!CUIT) = 11 Then
            DatosClientes = 80
            If Cliente = 9999 Then DatosClientes = 96
         Else
            DatosClientes = 96
        End If
        
        If Cliente = 9999 Then
            docCli = "11111111111"
            For I = 1 To (20 - Len(docCli))
                docCli = "0" & docCli
            Next I
            DatosClientes = DatosClientes & docCli
         Else
            If tClientes!CUIT <> "" Then
                docCli = tClientes!CUIT
                For I = 1 To (20 - Len(docCli))
                    docCli = "0" & docCli
                Next I
                DatosClientes = DatosClientes & docCli
            End If
        End If
    
        Nombre = tClientes!RazonSocial
        
        If Len(Nombre) <= 30 Then
            For I = 1 To (30 - Len(Nombre))
                Nombre = Nombre & " "
            Next I
        Else
            Nombre = Left(tClientes!RazonSocial, 30)
        End If
        
        DatosClientes = DatosClientes & Nombre
    
    End If

End Function

Private Sub FechasIniciales()
    
    Mes = Format(Date, "MM")
          
    Select Case Mes
        Case 1
            TextFechaDesde.text = "01/" & "01/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "01/" & Format(Date, "YYYY")
        
        Case 2
            If ((Format(Date, "YYYY") Mod 4) = 0) Then
                TextFechaDesde.text = "01/" & "02/" & Format(Date, "YYYY")
                TextFechaHasta.text = "29/" & "02/" & Format(Date, "YYYY")
            Else
                TextFechaDesde.text = "01/" & "02/" & Format(Date, "YYYY")
                TextFechaHasta.text = "28/" & "02/" & Format(Date, "YYYY")
            End If
        
        Case 3
            TextFechaDesde.text = "01/" & "03/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "03/" & Format(Date, "YYYY")
        
        Case 4
            TextFechaDesde.text = "01/" & "04/" & Format(Date, "YYYY")
            TextFechaHasta.text = "30/" & "04/" & Format(Date, "YYYY")
        
        Case 5
            TextFechaDesde.text = "01/" & "05/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "05/" & Format(Date, "YYYY")
        
        Case 6
            TextFechaDesde.text = "01/" & "06/" & Format(Date, "YYYY")
            TextFechaHasta.text = "30/" & "06/" & Format(Date, "YYYY")
        
        Case 7
            TextFechaDesde.text = "01/" & "07/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "07/" & Format(Date, "YYYY")
        
        Case 8
            TextFechaDesde.text = "01/" & "08/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "08/" & Format(Date, "YYYY")
        
        Case 9
            TextFechaDesde.text = "01/" & "09/" & Format(Date, "YYYY")
            TextFechaHasta.text = "30/" & "09/" & Format(Date, "YYYY")
        
        Case 10
            TextFechaDesde.text = "01/" & "10/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "10/" & Format(Date, "YYYY")
        
        Case 11
            TextFechaDesde.text = "01/" & "11/" & Format(Date, "YYYY")
            TextFechaHasta.text = "30/" & "11/" & Format(Date, "YYYY")
        
        Case 12
            TextFechaDesde.text = "01/" & "12/" & Format(Date, "YYYY")
            TextFechaHasta.text = "31/" & "12/" & Format(Date, "YYYY")
    End Select

End Sub

Private Sub titulos()

    MSFlexGrid2.Cols = 13
    MSFlexGrid2.Row = 0
    
    MSFlexGrid2.Col = 0
    MSFlexGrid2.text = "Suc"
    MSFlexGrid2.ColWidth(0) = 700
    MSFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 1
    MSFlexGrid2.text = "Cbnte"
    MSFlexGrid2.ColWidth(1) = 700
    MSFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.text = "Nº Cbnte"
    MSFlexGrid2.ColWidth(2) = 1000
    MSFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 3
    MSFlexGrid2.text = "Tipo"
    MSFlexGrid2.ColWidth(3) = 700
    MSFlexGrid2.ColAlignment(3) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 4
    MSFlexGrid2.text = "Fecha"
    MSFlexGrid2.ColWidth(4) = 1000
    MSFlexGrid2.ColAlignment(4) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 5
    MSFlexGrid2.text = "Cliente"
    MSFlexGrid2.ColWidth(5) = 3700
    MSFlexGrid2.ColAlignment(5) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 6
    MSFlexGrid2.text = "CUIT"
    MSFlexGrid2.ColWidth(6) = 1300
    MSFlexGrid2.ColAlignment(6) = flexAlignCenterCenter
        
    MSFlexGrid2.Col = 7
    MSFlexGrid2.text = "Neto"
    MSFlexGrid2.ColWidth(7) = 1300
    MSFlexGrid2.ColAlignment(7) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 8
    MSFlexGrid2.text = "Percep IIBB"
    MSFlexGrid2.ColWidth(8) = 1300
    MSFlexGrid2.ColAlignment(8) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 9
    MSFlexGrid2.text = "IVA"
    MSFlexGrid2.ColWidth(9) = 1300
    MSFlexGrid2.ColAlignment(9) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 10
    MSFlexGrid2.text = "Total"
    MSFlexGrid2.ColWidth(10) = 1300
    MSFlexGrid2.ColAlignment(10) = flexAlignCenterCenter
   
    MSFlexGrid2.Col = 11
    MSFlexGrid2.text = "Localidad"
    MSFlexGrid2.ColWidth(11) = 1700
    MSFlexGrid2.ColAlignment(11) = flexAlignCenterCenter
   
    MSFlexGrid2.Col = 12
    MSFlexGrid2.text = "Provincia"
    MSFlexGrid2.ColWidth(12) = 1800
    MSFlexGrid2.ColAlignment(12) = flexAlignCenterCenter
End Sub

    
Private Sub BotonBuscar_Click()
    Call buscodatos
End Sub

Private Sub BotonCancelar_Click()
     Call blanqueototal
End Sub

Private Sub BotonExportar_Click()

   If Exportar_Excel(App.Path & "\Libro IVA Ventas.xls", MSFlexGrid2) Then
        vVal = Shell(App.Path & "\ConverFormat.exe", 1)
        MsgBox " Datos exportados en " & App.Path, vbInformation
   End If
    
End Sub
Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    ' -- Bucle para Exportar los datos
    With MSFlexGrid2
        For Fila = 1 To .Rows - 1
            'If linea11 = 1 Then
            '    For Columna = 0 To .Cols - 3
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            
            'If linea22 = 1 Then
                
            '    For Columna = 0 To .Cols - 2
            '
            '        .ColWidth(3) = 0
            '        .Col = 3
            '        .Visible = False
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            'For Columna = 0 To .Cols - 2
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            Next
        
        Next
    End With
    'o_Libro.Close True, sOutputPath
    
    o_Libro.Close True, sOutputPath
    Set o_Libro = o_Excel.Workbooks.Open(sOutputPath)
    o_Excel.Visible = True
    
    'Call blanqueototal
    
    ' -- Cerrar Excel
    'o_Excel.Quit
    ' -- Terminar instancias
    'Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'Exportar_Excel = True
Exit Function

  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
    
End Function
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub
Private Sub blanqueototal()

    TextFechaDesde.text = ""
    TextFechaHasta.text = ""
    MSFlexGrid2.Clear
    TextFechaDesde.SetFocus
    txtImporteTotal.text = 0
    Call titulos
    
End Sub

Private Sub BotonImprimir_Click()

    Dim Nombre As String
    Dim direccion As String
    
    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    Nombre = "    QUILPLAC S.A."
    direccion = "     Av. Andres Baranda Nº520 Quilmes"
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pie de pagina
        .TextEncabezado1 = Chr(9) & "LIBRO IVA VENTAS"
            
                    'nombre = Chr(9) & direccion
'                    Pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.Text, 2)
                    
                     '& Chr(10) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Desarrollado por SPC Consulting"
                    'Pie = "Desarrollado por SPC Software Integral"
        
        .TextEncabezado2 = Chr(9) & Nombre & Chr(10) & Chr(9) & direccion & Chr(10) & Chr(9) & "    Desde el " & Format(TextFechaDesde.text, "DD/MM/YYYY") & " al " & Format(TextFechaHasta.text, "DD/MM/YYYY")
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Pie
               
        'Colores de la fuentes
        .ColorPiePagina = QBColor(4)
        'txtPiePagina.ForeColor
        .ColorEncabezado1 = QBColor(1)
        'txtEncabezado1.ForeColor
        .ColorEncabezado2 = QBColor(0)
        'txtEncabezado2.ForeColor
        
        'Tamaños de las fuentes
        .SizeEncabezado1 = 12
        .SizeEncabezado2 = 10
        .SizePiePagina = 11
        .AjustarColumnas = True
      
        .Orientacion = Horizontal
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid MSFlexGrid2
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)
    
    

End Sub

Private Sub BotonSalir_Click()

    Unload FormLibroIvaVentas
    
End Sub

Private Sub CmdExportarTXT_Click()

    FormTxtPercepciones.Show

End Sub

Private Sub cmdSIAP_Click()
  
  'Genero el archivo para SIAP
    Dim NombreArchivo, linea As String
    Dim NroFac As String
    Dim Tt As String
    Dim TotalFac() As String
    Dim TotalIIBB() As String
    Dim TtFac, TtIIBB As String
    Dim rstSIAP, vSQL
    Dim Desde, Hasta
    
    'Variable para usar WSH
        Dim Wscript As Object
      
    'Creamos la referencia para usar Windows Scripting Host
        Set Wscript = CreateObject("WScript.Shell")
    
                
        NombreArchivo = Wscript.SpecialFolders("Desktop") & "\VentasCITI_" & Format(TextFechaDesde.text, "yyyy") & Format(TextFechaDesde.text, "mm") & Format(TextFechaDesde.text, "dd")
        NombreArchivo = NombreArchivo & "_" & Format(TextFechaHasta.text, "yyyy") & Format(TextFechaHasta.text, "mm") & Format(TextFechaHasta.text, "dd") & ".txt"
        
    'MsgBox (NombreArchivo)
        Desde = "#" & Format(TextFechaDesde.text, "mm/dd/yyyy") & "#"
        Hasta = "#" & Format(TextFechaHasta.text, "mm/dd/yyyy") & "#"
    
        vSQL = "SELECT * FROM FacturaC WHERE FechaFactura >=" & Desde & " AND FechaFactura <=" & Hasta & " Order By TipoFactura, FechaFactura, NroFactura"
            
        'MsgBox (vSQL)
    
    Set rstSIAP = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    If Not rstSIAP.EOF Then
        rstSIAP.MoveFirst
     Else
        MsgBox ("No Hay registros en el archivo")
        Exit Sub
    End If
    
    Open NombreArchivo For Output As #1
        
    While Not rstSIAP.EOF
        'Punto 01 Fecha de comprobante
            linea = Format(rstSIAP!FechaFactura, "yyyymmdd")
        'Punto 02 Tipo de comprobante
            If rstSIAP!tipofactura = "A" Then linea = linea & "001"
            If rstSIAP!tipofactura = "B" Then linea = linea & "006"
        'Punto 03 Punto de venta
            linea = linea & "00003"
        'Punto 04 y 05 Numero de comprobante y Numero de comprobante hasta
            NroFac = rstSIAP!NroFactura
            
            For I = 1 To (20 - Len(NroFac))
                NroFac = "0" & NroFac
            Next I
            
            linea = linea & NroFac & NroFac
        
        'Punto 06 Codigo Doc Comprador 07 Nro ID comprador y 08 Apellido comprador
            linea = linea & DatosClientes(rstSIAP!CodCliente)
        
        'Punto 09 Importe total de la operacion
            Tt = Format(CStr(rstSIAP!TotalFactura), "#0.00")
            TotalFac = Split(Tt, ",", -1)
            TtFac = TotalFac(0)
            For I = 1 To (13 - Len(TotalFac(0)))
                TtFac = "0" & TtFac
            Next I
            
            TtFac = TtFac & TotalFac(1)
            
            linea = linea & TtFac
        
        'Punto 10
            linea = linea & "000000000000000"
        'Punto 11
            linea = linea & "000000000000000"
        'Punto 12
            linea = linea & "000000000000000"
        'Punto 13
            linea = linea & "000000000000000"
        'Punto 14 Percepcion IIBB
           ' linea = linea & "000000000000000"

            Tt = Format(CStr(rstSIAP!ImportePercepIIBB), "#0.00")
            TotalIIBB = Split(Tt, ",", -1)
            TtIIBB = TotalIIBB(0)
            For I = 1 To (13 - Len(TotalIIBB(0)))
                TtIIBB = "0" & TtIIBB
            Next I
            
            TtIIBB = TtIIBB & TotalIIBB(1)
            
            linea = linea & TtIIBB

        'Punto 15
            linea = linea & "000000000000000"
        'Punto 16
            linea = linea & "000000000000000"

            linea = linea & "PES"
            linea = linea & "0001000000"
            linea = linea & "1"
            linea = linea & "0"
        'Punto 21
            linea = linea & "000000000000000"
            linea = linea & Format(rstSIAP!FechaFactura, "yyyymmdd")

        Print #1, linea

        rstSIAP.MoveNext
    Wend
    
    rstSIAP.Close
    
'Ahora las Nota de Crédito
    
    vSQL = "SELECT * FROM NotaCreditoC WHERE FechaNotaCredito >=" & Desde & " AND FechaNotaCredito <=" & Hasta & " Order By TipoNotaCredito, FechaNotaCredito, NroNotaCredito"
    
    'MsgBox (vSQL)
    
    Set rstSIAP = db.OpenRecordset(vSQL, dbOpenDynaset)
    
    If Not rstSIAP.EOF Then
        rstSIAP.MoveFirst
     Else
        MsgBox ("No Hay registros en el archivo")
        Exit Sub
    End If
    
    'Open NombreArchivo For Output As #1
        
    While Not rstSIAP.EOF
        'Punto 01 Fecha de comprobante
            linea = Format(rstSIAP!FechaNotaCredito, "yyyymmdd")
      
        'Punto 02 Tipo de comprobante
            If rstSIAP!TipoNotaCredito = "A" Then linea = linea & "003"
            If rstSIAP!TipoNotaCredito = "B" Then linea = linea & "008"
      
        'Punto 03 Punto de venta
            linea = linea & "00003"
        
        'Punto 04 y 05 Numero de comprobante y Numero de comprobante hasta
            NroFac = rstSIAP!NroNotaCredito
            
            For I = 1 To (20 - Len(NroFac))
                NroFac = "0" & NroFac
            Next I
            
            linea = linea & NroFac & NroFac
        
        'Puntos 06 Codigo documento comprador. 07 Nro ID comprador y 08 Apellido Comprador
            linea = linea & DatosClientes(rstSIAP!CodCliente)

        'Punto 09 Importe total operacion
            Tt = Format(CStr(rstSIAP!TotalNotaCredito), "#0.00")
            TotalFac = Split(Tt, ",", -1)
            TtFac = TotalFac(0)
            For I = 1 To (13 - Len(TotalFac(0)))
                TtFac = "0" & TtFac
            Next I
            
            TtFac = TtFac & TotalFac(1)
            
            linea = linea & TtFac
        
        'Punto 10
            linea = linea & "000000000000000"
        'Punto 11
            linea = linea & "000000000000000"
        'Punto 12
            linea = linea & "000000000000000"
        'Punto 13
            linea = linea & "000000000000000"
        'Punto 14 Percepcion IIBB
           ' linea = linea & "000000000000000"

            Tt = Format(CStr(rstSIAP!ImportePercepIIBB), "#0.00")
            TotalIIBB = Split(Tt, ",", -1)
            TtIIBB = TotalIIBB(0)
            For I = 1 To (13 - Len(TotalIIBB(0)))
                TtIIBB = "0" & TtIIBB
            Next I
            
            TtIIBB = TtIIBB & TotalIIBB(1)
            
            linea = linea & TtIIBB

        'Punto 15
            linea = linea & "000000000000000"
        'Punto 16
            linea = linea & "000000000000000"

            linea = linea & "PES"
            linea = linea & "0001000000"
            linea = linea & "1"
            linea = linea & "0"
        'Punto 21
            linea = linea & "000000000000000"
            linea = linea & Format(rstSIAP!FechaNotaCredito, "yyyymmdd")

        Print #1, linea

        rstSIAP.MoveNext
    Wend
    
    rstSIAP.Close
    
    Close #1
    
    tClientes.Close
    
    MsgBox ("Archivo " & NombreArchivo & " Generado con Exito")
    
    Call GenerarArchivoAlicuotas

End Sub

Private Sub Form_Load()

    Dim Mes
    
    FormLibroIvaVentas.Height = 8715
    FormLibroIvaVentas.Width = 14820
    FormLibroIvaVentas.Top = 1000
    FormLibroIvaVentas.Left = 1000
    
    Call FechasIniciales
    Call titulos

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub buscodatos()

Dim neto As Double
Dim iva As Double
Dim percepcion As Double
Dim total As Double
Dim LiqTotal As Double
Dim rstNC

'***************Busco en PagoProvret
    
'On Error GoTo Error_Handler
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
          
    Desde = "#" & Format$(TextFechaDesde.text, "mm/dd/yyyy") & "#"
    'Desde = "#" & TextFechaDesde.Text & "#"
    Hasta = "#" & Format$(TextFechaHasta.text, "mm/dd/yyyy") & "#"
    'Hasta = "#" & TextFechaHasta.Text & "#"
    
    
    eseqele = "SELECT * FROM FacturaC WHERE FechaFactura >=" & Desde & " AND FechaFactura <=" & Hasta & " Order By TipoFactura, FechaFactura, NroFactura"
    eseqeleNC = "SELECT * FROM NotaCreditoC WHERE FechaNotaCredito >=" & Desde & " AND FechaNotaCredito <=" & Hasta & " Order By TipoNotaCredito, FechaNotaCredito, NroNotaCredito"
    
    'MsgBox (eseqele)
    
    Set rst = db.OpenRecordset(eseqele, dbOpenDynaset)
    Set rstNC = db.OpenRecordset(eseqeleNC, dbOpenDynaset)
   
    
    
    MSFlexGrid2.Rows = 2
    MSFlexGrid2.Clear
    MSFlexGrid2.Visible = True
    
    LiqTotal = 0
    
    total = 0
    
    Call titulos
    
    
    If Not rst.EOF Then rst.MoveFirst

    linea2 = 1
   'Do While Not rst.NoMatch
   While Not rst.EOF
         MSFlexGrid2.AddItem ""
         MSFlexGrid2.Row = linea2
         MSFlexGrid2.Col = 1
         MSFlexGrid2.text = "Factura"
         MSFlexGrid2.Col = 2
         MSFlexGrid2.text = rst.Fields!NroFactura
         MSFlexGrid2.Col = 0
         MSFlexGrid2.text = "0003"
         
         MSFlexGrid2.Col = 3
         MSFlexGrid2.text = rst.Fields!tipofactura
         MSFlexGrid2.Col = 4
         MSFlexGrid2.text = Format(rst.Fields!FechaFactura, "dd-MMM-yyyy")
         
         '**** Busco datos Cliente
         
       
         CodigoClie = rst.Fields!CodCliente
    
                        
         rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
  
         MSFlexGrid2.Col = 5
         MSFlexGrid2.text = rstCliente.Fields!RazonSocial
         MSFlexGrid2.Col = 6
         If rstCliente.Fields!CUIT <> "" Then MSFlexGrid2.text = rstCliente.Fields!CUIT
         MSFlexGrid2.Col = 11
         MSFlexGrid2.text = rstCliente.Fields!localidad
         MSFlexGrid2.Col = 12
         MSFlexGrid2.text = rstCliente.Fields!Prov
         
         '********************************************
         
         MSFlexGrid2.Col = 7
         MSFlexGrid2.text = Format(rst.Fields!SubTotalFactura, "#00.00")
         neto = rst.Fields!SubTotalFactura
         MSFlexGrid2.Col = 8
         MSFlexGrid2.text = Format(rst.Fields!ImportePercepIIBB, "#00.00")
         percepcion = rst.Fields!ImportePercepIIBB
         MSFlexGrid2.Col = 9
         'iva = (neto * 21) / 100
         iva = Format(rst.Fields!totalIva, "#00.00")
         MSFlexGrid2.text = Format(iva, "#00.00")
         'total = neto + percepcion + iva
         total = Format(rst.Fields!TotalFactura, "#00.00")
         MSFlexGrid2.Col = 10
         MSFlexGrid2.text = Format(total, "#00.00")
         
         LiqTotal = LiqTotal + total
         
         linea2 = linea2 + 1
         
         '**** agregado el 01/09/2016
            total = 0
         '***************************
         
         rst.MoveNext
        
   'Loop
   Wend

If Not rstNC.EOF Then rstNC.MoveFirst
   total = 0

   While Not rstNC.EOF
         MSFlexGrid2.AddItem ""
         MSFlexGrid2.Row = linea2
         MSFlexGrid2.Col = 1
         MSFlexGrid2.text = "NC"
         MSFlexGrid2.Col = 2
         MSFlexGrid2.text = rstNC.Fields!NroNotaCredito
         MSFlexGrid2.Col = 0
         MSFlexGrid2.text = "0003"
         
         MSFlexGrid2.Col = 3
         MSFlexGrid2.text = rstNC.Fields!TipoNotaCredito
         MSFlexGrid2.Col = 4
         MSFlexGrid2.text = Format(rstNC.Fields!FechaNotaCredito, "dd-MMM-yyyy")
         
         '**** Busco datos Cliente
         
       
         CodigoClie = rstNC.Fields!CodCliente
    
                        
         rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
  
         MSFlexGrid2.Col = 5
         MSFlexGrid2.text = rstCliente.Fields!RazonSocial
         MSFlexGrid2.Col = 6
         If rstCliente!CUIT <> "" Then MSFlexGrid2.text = rstCliente.Fields!CUIT
         MSFlexGrid2.Col = 11
         MSFlexGrid2.text = rstCliente.Fields!localidad
         MSFlexGrid2.Col = 12
         MSFlexGrid2.text = rstCliente.Fields!Prov
         
         '********************************************
         
         MSFlexGrid2.Col = 7
         MSFlexGrid2.text = (Format(rstNC.Fields!SubTotalNotaCredito, "#00.00")) * -1
         neto = rstNC.Fields!SubTotalNotaCredito
         MSFlexGrid2.Col = 8
         MSFlexGrid2.text = (Format(rstNC.Fields!ImportePercepIIBB, "#00.00")) * -1
         percepcion = rstNC.Fields!ImportePercepIIBB
         MSFlexGrid2.Col = 9
         'iva = (neto * 21) / 100
         iva = (Format(rstNC.Fields!totalIva, "#00.00")) * -1
         MSFlexGrid2.text = (Format(iva, "#00.00"))
         'total = neto + percepcion + iva
         total = (Format(rstNC.Fields!TotalNotaCredito, "#00.00")) * -1
         MSFlexGrid2.Col = 10
         MSFlexGrid2.text = (Format(total, "#00.00"))
         
         LiqTotal = LiqTotal + total
         
         linea2 = linea2 + 1
         
         '**** agregado el 01/09/2016
            total = 0
         '***************************
         
         rstNC.MoveNext
   'Loop
   Wend

        txtImporteTotal.text = FormatCurrency(LiqTotal, 2)
        
Error_Handler:
    
    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If
    
   'TxtTotalRetencion.Text = Format(totalrete, "#0.00")
   ' TxtTOTAL.Text = Format(totalpa, "#0.00")
    
    Exit Sub
   
End Sub

Private Sub TextFechaDesde_GotFocus()

    TextFechaDesde.SelLength = Len(TextFechaDesde.text)

End Sub

Private Sub TextFechaDesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TextFechaDesde_LostFocus()
    
    If Not IsDate(TextFechaDesde.text) Then
        MsgBox "Formato de Fecha Incorrecto", vbCritical, "ERROR !"
        TextFechaDesde.text = Format(Date, "DD/MM/YYYY")
    End If

End Sub


Private Sub TextFechaHasta_GotFocus()

    TextFechaHasta.SelLength = Len(TextFechaHasta.text)

End Sub

Private Sub TextFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TextFechaHasta_LostFocus()
    
    If Not IsDate(TextFechaHasta.text) Then
        MsgBox "Formato de Fecha Incorrecto", vbCritical, "ERROR !"
        TextFechaHasta.text = Format(Date, "DD/MM/YYYY")
    End If

End Sub


