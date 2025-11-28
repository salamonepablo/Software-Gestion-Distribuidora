VERSION 5.00
Begin VB.Form FormImprimirND 
   Caption         =   "Imprimir"
   ClientHeight    =   3660
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   8760
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicQR 
      Height          =   375
      Left            =   2040
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   10
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PictureQP 
      Height          =   375
      Left            =   480
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   8
      Top             =   3000
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   7935
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6120
         TabIndex        =   5
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BotonAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7935
      Begin VB.TextBox TextTipoFactura 
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
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Imprimir ND E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdGenerarPDF 
         Caption         =   "Generar PDF"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton cmdFacturaEl 
         Caption         =   "Generar ND E"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TextNumeroFactura 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         TabIndex        =   1
         Text            =   "15094"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroRemito 
         Alignment       =   1  'Right Justify
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
         Left            =   2160
         TabIndex        =   3
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CheckBox CheckImprimirRemito 
         Caption         =   "Imprimir Remito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CheckBox CheckImprimirFactura 
         Caption         =   "Imprimir ND"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   0
         Top             =   360
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FormImprimirND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cl As New arisBarcode

Private Function CalcularBarCode() As String
    
    Dim TipoC, FechaVC As String
    
    If tNotaDebitoC!TipoDebito = "A" Then TipoC = "02"
    If tNotaDebitoC!TipoDebito = "B" Then TipoC = "07"
    
    'FechaVC = Year(tFacturaC!FechaVC) & Month(tFacturaC!FechaVC) & Day(tFacturaC!FechaVC)
    FechaVC = Year(tNotaDebitoC!FechaVC) & Format(Month(tNotaDebitoC!FechaVC), "00") & Format(Day(tNotaDebitoC!FechaVC), "00")
    
  '  MsgBox (FechaVC)

    CalcularBarCode = "30708432543" & TipoC & "0003" & tNotaDebitoC!CAE & FechaVC & CalculoDigitoVerificador("30708432543")

End Function

Public Function CalculoDigitoVerificador(CUIT As String) As String

    Dim Texto As Variant
    Dim SumaImp, SumaPar, SumaTotal  As Long
    
    SumaImp = 0
    SumaPar = 0
    SumaTotal = 0
    
    Texto = CUIT
    
    For I = 1 To 11
        Select Case I
            Case 1
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 2
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 3
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 4
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 5
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 6
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 7
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 8
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 9
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 10
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 11
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case Else
        End Select
    Next I

    SumaImp = SumaImp * 3
    SumaTotal = SumaImp + SumaPar
    
    For J = 0 To 9
        
        If (SumaTotal + J) Mod (10) = 0 Then
            CalculoDigitoVerificador = CStr(J)
            Exit For
        End If
    Next J

   ' MsgBox (CalculoDigitoVerificador)
    
End Function


Private Sub CrearBarCode(Texto As String)

    'PicBC.FontName = Me.FontName
    'PicBC.FontSize = Me.FontSize
    'PicBC.Cls
    
    'cl.Code128 PicBC, 0.5, Texto, True
    'SavePicture PicBC.Picture, App.Path & "\BarCode.jpg"

End Sub

Private Sub GenerarFEB()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tNotaDebitoD = BaseSPC.OpenRecordset("NotaDebitoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaDebitoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaDebitoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Debito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Nota de Débito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub ImprimirFactura()
    
    Dim PU, TL, ImpIva, ImpIIBB, SubTotalFac, TotalFac, Cant As Variant
    
    X = -4
    Y = -4
    renglon = 0
    vNroRemito = "0003- " & TextNumeroRemito.text
    
    With Printer
        
        'On Error GoTo CapturaErrores
            .Copies = 2
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Imprimir Fecha
            .CurrentX = X + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(FormFactura.TextFechaFactura.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = X + 37
            .CurrentY = Y + 54
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print FormFactura.TextApellidoNombre.text
            
        'Imprimir Direccion
            .CurrentX = X + 37
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print FormFactura.TextDireccion.text
            
        'Imprimir Localidad
            .CurrentX = X + 37
            .CurrentY = Y + 65
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print FormFactura.TextLocalidad.text
            
        'Imprimir CUIT
            .CurrentX = X + 125
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print FormFactura.TextCuit.text
            
        'Imprimir Marca Responsable Inscripto
            .CurrentX = X + 57
            .CurrentY = Y + 70
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
            .CurrentX = X + 70
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca CtaCte
            .CurrentX = X + 100
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Nro Remito
            .CurrentX = X + 138
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 9
            .FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
            sqlfc = "SELECT * FROM FacturaC WHERE TipoFactura='" & FormFactura.TextTipoFactura.text & "' AND NroFactura=" & FormFactura.TextNumeroFactura.text & " ORDER By NroFactura"
            vsqlFD = "SELECT * FROM FacturaD WHERE TipoFactura='" & FormFactura.TextTipoFactura.text & "' AND NroFactura=" & FormFactura.TextNumeroFactura.text & " ORDER By NroFactura"
            
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
                        
            Set FacC = BaseSPC.OpenRecordset(sqlfc, dbOpenDynaset)
            Set FacD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            FacC.MoveFirst
            FacD.MoveFirst
                
                    While Not FacD.EOF
                        'Imprimo el detalle
                            .CurrentX = X + 22
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Cant = CDbl(FacD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print FacD!cantidad
                            
                        'Detalle
                            .CurrentX = X + 40
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'Printer.Print FacD!IdCodProd & Chr(9) & Descripcion(FacD!IdCodProd)
                            Printer.Print Descripcion(FacD!IdCodProd)
                        
                        'Precio
                            .CurrentX = X + 122
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(FacD!precioUnitario) - (CDbl(FacD!precioUnitario) * CDbl(FacD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU
                    
                        
                        'Importe
                            .CurrentX = X + 142
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            TL = Format(FacD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                        
                         renglon = renglon + 5
                            
                        FacD.MoveNext
                    Wend
           
            'Importe SubTotal
                .CurrentX = X + 142
                .CurrentY = Y + 176
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                SubTotalFac = Format(CDbl(FacC!SubTotalFactura), "Standard")
                Hasta = CInt(14 - Len(SubTotalFac))
                For I = 0 To Hasta
                    SubTotalFac = " " & SubTotalFac
                Next I

                Printer.Print SubTotalFac

                
            'Alicuota IVA
                .CurrentX = X + 132
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print "21"
            
            'Importe IVA
                .CurrentX = X + 142
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                ImpIva = Format(CDbl(FacC!totalIva), "Standard")
                Hasta = CInt(14 - Len(ImpIva))
                For I = 0 To Hasta
                    ImpIva = " " & ImpIva
                Next I
                
                Printer.Print ImpIva
            
            If FacC!ImportePercepIIBB > 0 Then
                'Alicuota IIBB
                    .CurrentX = X + 122
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print "Per.IIBB"
                
                'Importe IIBB
                    .CurrentX = X + 142
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    ImpIIBB = Format(CDbl(FacC!ImportePercepIIBB), "Standard")
                    Hasta = CInt(14 - Len(ImpIIBB))
                    For I = 0 To Hasta
                        ImpIIBB = " " & ImpIIBB
                    Next I
                    
                    Printer.Print ImpIIBB
            End If
            
            'Importe Total
                .CurrentX = X + 142
                .CurrentY = Y + 194
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                TotalFac = Format(CDbl(FacC!TotalFactura), "Standard")
                Hasta = CInt(14 - Len(TotalFac))
                For I = 0 To Hasta
                    TotalFac = " " & TotalFac
                Next I
                
                Printer.Print TotalFac
            
        .EndDoc
        
    End With
    
    FacC.Close
    FacD.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If

End Sub

Private Sub ImprimirFacturaB()
    
    Dim PU, TL, ImpIva, ImpIIBB, SubTotalFac, TotalFac, Cant As Variant
    
    X = -4
    Y = -4
    renglon = 0
    vNroRemito = "0003- " & TextNumeroRemito.text
    
    With Printer
        
        'On Error GoTo CapturaErrores
            .Copies = 2
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Imprimir Fecha
            .CurrentX = X + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(FormFactura.TextFechaFactura.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = X + 37
            .CurrentY = Y + 54
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print FormFactura.TextApellidoNombre.text
            
        'Imprimir Direccion
            .CurrentX = X + 37
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print FormFactura.TextDireccion.text
            
        'Imprimir Localidad
            .CurrentX = X + 37
            .CurrentY = Y + 65
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print FormFactura.TextLocalidad.text
            
        'Imprimir CUIT
            .CurrentX = X + 125
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            'Printer.Print FormFactura.TextCuit.Text
            
        'Imprimir Marca Responsable Inscripto
            .CurrentX = X + 57
            .CurrentY = Y + 70
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
            .CurrentX = X + 70
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca CtaCte
            .CurrentX = X + 100
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Nro Remito
            .CurrentX = X + 138
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 9
            .FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
            sqlfc = "SELECT * FROM FacturaC WHERE TipoFactura='" & FormFactura.TextTipoFactura.text & "' AND NroFactura=" & FormFactura.TextNumeroFactura.text & " ORDER By NroFactura"
            vsqlFD = "SELECT * FROM FacturaD WHERE TipoFactura='" & FormFactura.TextTipoFactura.text & "' AND NroFactura=" & FormFactura.TextNumeroFactura.text & " ORDER By NroFactura"
            
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
                        
            Set FacC = BaseSPC.OpenRecordset(sqlfc, dbOpenDynaset)
            Set FacD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            FacC.MoveFirst
            FacD.MoveFirst
                
                    While Not FacD.EOF
                        'Imprimo el detalle
                            .CurrentX = X + 22
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Cant = CDbl(FacD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print FacD!cantidad
                            
                        'Detalle
                            .CurrentX = X + 40
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'Printer.Print FacD!IdCodProd & Chr(9) & Descripcion(FacD!IdCodProd)
                            Printer.Print Descripcion(FacD!IdCodProd)
                        
                        'Precio
                            .CurrentX = X + 122
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = (CDbl(FacD!precioUnitario) - (CDbl(FacD!precioUnitario) * CDbl(FacD!PorcentajeDescuento) / 100)) * 1.21
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU
                    
                        
                        'Importe
                            .CurrentX = X + 142
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'TL = Format(FacD!totalLinea, "Standard")
                            TL = PU * Cant
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                        
                         renglon = renglon + 5
                            
                        FacD.MoveNext
                    Wend
           
            'Importe SubTotal
           '     .CurrentX = X + 142
           '     .CurrentY = Y + 176
           '     .Font = "Courier New"
           '     .FontSize = 8
           '     .FontBold = False
           '     SubTotalFac = Format(CDbl(FacC!SubTotalFactura), "Standard")
           '     Hasta = CInt(14 - Len(SubTotalFac))
           '     For I = 0 To Hasta
           '         SubTotalFac = " " & SubTotalFac
           '     Next I'

           '     Printer.Print SubTotalFac

                
            'Alicuota IVA
           '     .CurrentX = X + 132
           '     .CurrentY = Y + 182
           '     .Font = "Courier New"
           '     .FontSize = 8
           '     .FontBold = False
           '     Printer.Print "21"
            
            'Importe IVA
            '    .CurrentX = X + 142
            '    .CurrentY = Y + 182
            '    .Font = "Courier New"
            '    .FontSize = 8
            '    .FontBold = False
             '   ImpIva = Format(CDbl(FacC!TotalIVA), "Standard")
            '    Hasta = CInt(14 - Len(ImpIva))
            '    For I = 0 To Hasta
            '        ImpIva = " " & ImpIva
            '    Next I
            '
            '    Printer.Print ImpIva
            
           ' If FacC!ImportePercepIIBB > 0 Then
                'Alicuota IIBB
           '         .CurrentX = X + 122
           '         .CurrentY = Y + 187
           '         .Font = "Courier New"
           '         .FontSize = 8
           '         .FontBold = False
           '         Printer.Print "Per.IIBB"
                
                'Importe IIBB
           '         .CurrentX = X + 142
           '         .CurrentY = Y + 187
           '         .Font = "Courier New"
           '         .FontSize = 8
           '         .FontBold = False
           '         ImpIIBB = Format(CDbl(FacC!ImportePercepIIBB), "Standard")
           '         Hasta = CInt(14 - Len(ImpIIBB))
           '         For I = 0 To Hasta
           '             ImpIIBB = " " & ImpIIBB
           '         Next I
           '
            '        Printer.Print ImpIIBB
           ' End If
            
            'Importe Total
                .CurrentX = X + 142
                .CurrentY = Y + 194
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                TotalFac = Format(CDbl(FacC!TotalFactura), "Standard")
                Hasta = CInt(14 - Len(TotalFac))
                For I = 0 To Hasta
                    TotalFac = " " & TotalFac
                Next I
                
                Printer.Print TotalFac
            
        .EndDoc
        
    End With
    
    FacC.Close
    FacD.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If

End Sub

Private Sub ImprimirFE()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Débito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        'For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                        '    If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        'Next
                    'Set Printer = Printers(5)
                    
                    For I = 0 To Printers.Count - 1
                        'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                        If Printers(I).DeviceName = "HP Laser 103 107 108" Then Set Printer = Printers(I)
                    Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Factura Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub ImprimirFEB()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tNotaDebitoD = BaseSPC.OpenRecordset("NotaDebitoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaDebitoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaDebitoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Debito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Set Printer = Printers(6)
                    
                    For I = 0 To Printers.Count - 1
                        'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                        If Printers(I).DeviceName = "HP Laser 103 107 108" Then Set Printer = Printers(I)
                    Next
                    
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 15
                        .CurrentY = 245
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 15
                        .CurrentY = 252
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 12
                        .CurrentY = 263
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        .FontName = "Interleaved 2of5"
                        .FontSize = 20
                        Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Nota de Débito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:


End Sub
Private Sub ImprimirRemito()

    Dim RemC
    Dim RemD
        
    'On Error GoTo CapturaErrores

    X = -4
    Y = -4
          renglon = 0
    vNroRemito = "0003- " & TextNumeroRemito.text
    
    vSQLRc = "SELECT * FROM RemitoC WHERE NroRemito='" & TextNumeroRemito.text & "' ORDER By NroRemito"
    vSQLRd = "SELECT * FROM RemitoD WHERE NroRemito='" & TextNumeroRemito.text & "' ORDER By NroRemito"
    vSQLRdir = "SELECT * FROM RemitoD WHERE NroRemito='" & TextNumeroRemito.text & "' ORDER By NroRemito"
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    
    Set RemC = BaseSPC.OpenRecordset(vSQLRc, dbOpenDynaset)
    Set RemD = BaseSPC.OpenRecordset(vSQLRd, dbOpenDynaset)
      
        
    'With p
        'Seteo escala a mm
            Printer.Copies = 3
            Printer.ScaleMode = 6
        
        'Imprimir Fecha
            Printer.CurrentX = X + 130
            Printer.CurrentY = Y + 32
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print Format(FormFactura.TextFechaFactura.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
           Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 57
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = True
            Printer.Print FormFactura.TextApellidoNombre.text
            
        'Imprimir Direccion
            Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 64
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormFactura.TextDireccion.text
            
        'Imprimir Localidad
            Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 70
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormFactura.TextLocalidad.text
            
        'Imprimir CUIT
            Printer.CurrentX = X + 125
            Printer.CurrentY = Y + 70
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormFactura.TextCuit.text
            
        'Imprimir Marca Responsable Inscripto
            Printer.CurrentX = X + 115
            Printer.CurrentY = Y + 76
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
         '   Printer.CurrentX = X + 70
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Marca CtaCte
         '   Printer.CurrentX = X + 100
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Nro Remito
            Printer.CurrentX = X + 138
            Printer.CurrentY = Y + 80
            Printer.Font = "Courier New"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
       '     sqlFC = "SELECT * FROM FacturaC WHERE TipoFactura='" & TextTipoFactura.Text & "' AND NroFactura=" & TextNumeroFactura.Text & " ORDER By NroFactura"
       '     vsqlFD = "SELECT * FROM FacturaD WHERE TipoFactura='" & TextTipoFactura.Text & "' AND NroFactura=" & TextNumeroFactura.Text & " ORDER By NroFactura"
            
          '  Set RemC = BaseSPC.OpenRecordset(vsqlFC, dbOpenDynaset)
          '  Set RemD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            RemC.MoveFirst
            RemD.MoveFirst
                
                    While Not RemD.EOF
                        'Imprimo el detalle
                            Printer.CurrentX = X + 30
                            Printer.CurrentY = Y + 96 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            Printer.Print RemD!cantidad
                            
                        'Detalle
                            Printer.CurrentX = X + 50
                            Printer.CurrentY = Y + 96 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            'Printer.Print RemD!IdCodProd & Chr(9) & Descripcion(RemD!IdCodProd)
                            Printer.Print Chr(9) & Descripcion(RemD!IdCodProd)
                        
                         renglon = renglon + 5
                            
                        RemD.MoveNext
                    Wend
        Printer.EndDoc
        
'    End With
    
    RemC.Close
    RemD.Close
    
    BaseSPC.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If
End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function

Private Sub BotonAceptar_Click()

    If FormImprimir.CheckImprimirFactura.Value = 1 Then
        If FormFactura.TextTipoFactura = "A" Then Call ImprimirFactura
        If FormFactura.TextTipoFactura = "B" Then Call ImprimirFacturaB
    End If
        
    If FormImprimir.CheckImprimirRemito.Value = 1 Then
        Call ImprimirRemito
    End If

End Sub


Private Sub BotonSalir_Click()
    
    
    Call FormNotasdeDebito.blanqueototal
        
    FormNotasdeDebito.TextCodigoCliente.SetFocus
    FormNotasdeDebito.MSFlexGrid1.Visible = False
                  
    Unload FormImprimirND

End Sub

Private Sub cmdFacturaEl_Click()
        
        Dim vVal As Double
        vVal = Shell(App.Path & "\NotaDebito.exe " & TextNumeroFactura.text & " " & TextTipoFactura.text, 1)
        'vVal = Shell(App.Path & "\FacturacionElectronicaTest.exe " & TextNumeroFactura.Text & " " & TextTipoFactura.Text, 1)

End Sub

Private Sub Command1_Click()

End Sub


Private Sub cmdGenerarPDF_Click()

    If TextTipoFactura.text = "A" Then
        Call GenerarFE
        'MsgBox ("Genera Duplicado")
        'Call GenerarFED
    End If

    If TextTipoFactura.text = "B" Then
        Call GenerarFEB
        'MsgBox ("Genera Duplicado")
        'Call GenerarFEBD
    End If

End Sub


Private Sub cmdImprimir_Click()

    If TextTipoFactura.text = "A" Then
        Call ImprimirFE
        'Call ImprimirFED
    End If
    
    If TextTipoFactura.text = "B" Then
        Call ImprimirFEB
        'Call ImprimirFEBD
    End If
    

End Sub


Private Sub Form_Load()

    TextNumeroFactura.text = vNroNDImp
    TextNumeroRemito.text = vNroRemNDImp
    TextTipoFactura.text = vTipoNDImp

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call FormNotasdeDebito.blanqueototal
    
End Sub


Private Sub TextNumeroFactura_GotFocus()
     
     TextNumeroFactura.SelLength = Len(TextNumeroFactura.text)

End Sub
Private Sub TextNumeroRemito_GotFocus()
     
     TextNumeroRemito.SelLength = Len(TextNumeroRemito.text)

End Sub

Private Sub GenerarFE()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Débito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0003-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0003-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Factura Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
End Function

Private Function BarCodeIL2of5(Cadena As String) As String
    
    Dim I As Long
    
    BarCodeIL2of5 = Chr(40)
    
    For I = 1 To Len(Cadena) Step 2
        If Val(Mid(Cadena, I, 2)) < 50 Then
          BarCodeIL2of5 = BarCodeIL2of5 & Chr(Val(Mid(Cadena, I, 2)) + 48)
        Else
          BarCodeIL2of5 = BarCodeIL2of5 & Chr(Val(Mid(Cadena, I, 2)) + 142)
        End If
    Next I
    
    BarCodeIL2of5 = BarCodeIL2of5 & Chr(41)


End Function

