VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormComprobanteIIBBFacturas 
   Caption         =   "Generar Comprobante de Retencion IIBB"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6750
   ScaleWidth      =   13050
   Begin VB.Frame Frame1 
      Caption         =   "Pago a Proveedores / Comprobante de Retención IIBB"
      Height          =   6375
      Left            =   0
      TabIndex        =   17
      Top             =   120
      Width           =   12975
      Begin VB.TextBox TxtTotLinea 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10200
         TabIndex        =   36
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox Textnumfac 
         Height          =   285
         Left            =   10680
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.ComboBox cmbCodProv 
         Height          =   315
         Index           =   1
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   4815
      End
      Begin VB.ComboBox cmbCodProv 
         Height          =   315
         Index           =   0
         Left            =   4440
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox TxtTotalRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   5040
         Width           =   1695
      End
      Begin VB.TextBox TxtCodProv 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   10
         Top             =   150
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.TextBox TxtPayDate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Text            =   "12/06/2016"
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3120
         TabIndex        =   18
         Top             =   4920
         Width           =   9615
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            Enabled         =   0   'False
            Height          =   735
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdNew 
            Caption         =   "&Nuevo"
            Height          =   735
            Left            =   6840
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmndExcel 
            Caption         =   "&Excel"
            Height          =   735
            Left            =   5520
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Imprimir"
            Height          =   735
            Left            =   4200
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "&Guardar"
            Height          =   735
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "&Salir"
            Height          =   735
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtCUIT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   3
         Top             =   750
         Width           =   1215
      End
      Begin VB.TextBox TxtTOTAL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   5640
         Width           =   1695
      End
      Begin VB.TextBox TxtTotFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3240
         TabIndex        =   4
         Top             =   4320
         Width           =   1335
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   0
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TxtBaseI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   5
         Top             =   4320
         Width           =   1335
      End
      Begin VB.TextBox TxtAlic 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6360
         TabIndex        =   6
         Top             =   4320
         Width           =   615
      End
      Begin VB.TextBox TxtImpRet 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7200
         TabIndex        =   7
         Top             =   4320
         Width           =   1215
      End
      Begin VB.TextBox TxtPayNr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtIva 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8640
         TabIndex        =   8
         Top             =   4320
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2175
         Left            =   840
         TabIndex        =   33
         Top             =   1680
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Línea:"
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
         Left            =   10200
         TabIndex        =   37
         Top             =   4080
         Width           =   1065
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Total Retencion:"
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
         Left            =   960
         TabIndex        =   30
         Top             =   4800
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Proveedor"
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
         Left            =   4080
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago:"
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
         Left            =   2520
         TabIndex        =   28
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nro. Pago:"
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
         Left            =   1320
         TabIndex        =   27
         Top             =   480
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Facturas a Pagar"
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
         Left            =   240
         TabIndex        =   26
         Top             =   1200
         Width           =   2400
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CUIT:"
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
         Left            =   9960
         TabIndex        =   25
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre / Razón Social"
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
         Left            =   4080
         TabIndex        =   24
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total a Pagar:"
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
         Left            =   960
         TabIndex        =   23
         Top             =   5400
         Width           =   1230
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total Factura"
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
         Left            =   3120
         TabIndex        =   22
         Top             =   4080
         Width           =   1155
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Base Imponible"
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
         Left            =   4680
         TabIndex        =   21
         Top             =   4080
         Width           =   1305
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "% - Imp Ret IIBB"
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
         Left            =   6240
         TabIndex        =   20
         Top             =   4080
         Width           =   1410
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Importe IVA"
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
         Left            =   8520
         TabIndex        =   19
         Top             =   4080
         Width           =   1005
      End
   End
End
Attribute VB_Name = "FormComprobanteIIBBFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim rstFacturaC As DAO.Recordset
  Public totalesRetenciones As Double
    Public totalesFacturas As Double

Private Sub CleanDatos()

    TxtNroFac.text = ""
    TxtFF.text = ""
    TxtTotFac.text = ""
    TxtBaseI.text = ""
    TxtImpRet.text = ""
    'TxtAlic.Text = ""
    TxtIVA.text = ""
    TxtTotLinea.text = ""
    
    
    
    TxtNroFac.SetFocus
    
End Sub

Private Sub CleanDatos2()

    TxtPayNr.text = ""
  '  TxtPayDate.Text = ""
  '  CmbCodProv.Text = ""
    TxtProvName.text = ""
    TxtCUIT.text = ""
    TxtNroFac.text = ""
    TxtFF.text = ""
    TxtTotFac.text = ""
    TxtBaseI.text = ""
    TxtImpRet.text = ""
    TxtAlic.text = ""
    TxtIVA.text = ""
    TxtTotLinea.text = ""
    TxtTotalRetencion.text = ""
    txtTotal.text = ""
    FG1.Clear
    
    TxtPayNr.SetFocus

End Sub

Private Sub GrabarPago(NroPago As String, CodProv As Integer, NFac As String, FF As Date, TotFac As Double, ImpRet As Double, iva As Double, totalLinea As Double)
    
End Sub


Private Sub LlenaGrilla()
    
    item = item + 1
    FG1.Row = Fila + 1
    FG1.Col = Columna
    FG1.text = item
    FG1.Col = 1
    
    FG1.text = Format(TxtPayNr.text, "")
    FG1.Col = 2
    
    FG1.TextMatrix(FG1.Row, FG1.Col) = TxtNroFac.text
    FG1.Col = 3
    
    FG1.text = TxtFF.text
    FG1.Col = 4
    
    FG1.text = TxtTotFac.text
    FG1.Col = 5
    
    FG1.text = TxtBaseI.text
    FG1.Col = 6
    
    FG1.text = TxtAlic.text
    FG1.Col = 7
    
    FG1.text = TxtImpRet.text
    FG1.Col = 8
    
    FG1.text = TxtIVA.text
    FG1.Col = 9

    FG1.text = TxtTotLinea.text
    
    txtTotal.text = Format(total, "#,###,###,#0.00")
    TxtTotalRetencion.text = Format(TOTALRETENCION, "#,###,###,#0.00")
    
    Columna = 0
    
    Fila = Fila + 1
    
    FG1.Rows = FG1.Rows + 1
    
    Call CleanDatos
    
End Sub


Private Sub Agregar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub





Private Sub busco()

   Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
   Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
    
   CodiProv = Val(cmbCodProv(0).text)
      
    rst.FindFirst "CodProv= " + Str(CodiProv)
    
    TxtProvName.text = rst.Fields!NombreProv
    TxtCUIT.text = rst.Fields!CUIT
   
    
    Set tPadron = Padron.OpenRecordset("Padron", dbOpenTable)
    
    tPadron.Index = "CUIT"
    
    With tPadron
        tPadron.Seek "=", TxtCUIT.text
        If .NoMatch = False Then
            Alicuota = !AlicuotaRetencion
        End If
    End With

End Sub

Private Sub CmbCodProvViejo_LostFocus()

    'Call busco
      
   ' Provs.Index = "Primary"
    
   
    
    'With Provs
      ' Provs.Seek "=", CmbCodProv.Text
      
     '   If .NoMatch = False Then
      '      TxtProvName.Text = rst.Fields!NombreProv
      '      TxtCUIT.Text = rst.Fields!Cuit
      '  End If
    'End With
    
   ' Set TPadron = Padron.OpenRecordset("Padron", dbOpenTable)
    
   ' TPadron.Index = "CUIT"
    
   ' With TPadron
   '     TPadron.Seek "=", TxtCUIT.Text
   '     If .NoMatch = False Then
   '         Alicuota = !AlicuotaRetencion
   '     End If
   ' End With

End Sub

Private Sub cmbCodProv_Click(Index As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex

End Sub

Private Sub CmbCodProv_KeyPress(Index As Integer, KeyAscii As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub

Private Sub CmbCodProv_LostFocus(Index As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex
    
    
 If cmbCodProv(0).text <> "" Then
        
     With Provs
        .Index = "Primary"
        Provs.Seek "=", cmbCodProv(0).text
       
        If .NoMatch = False Then
            ' TxtProvName.Text = rst.Fields!NombreProv
             TxtCUIT.text = !CUIT
             TxtCodProv.text = !CodProv
        End If
     End With
     
    Set tPadron = Padron.OpenRecordset("Padron", dbOpenTable)
     
    tPadron.Index = "CUIT"
     
     With tPadron
         tPadron.Seek "=", TxtCUIT.text
         If .NoMatch = False Then
             Alicuota = !AlicuotaRetencion
             TxtAlic.text = Format(!AlicuotaRetencion, "#,###,###,#0.00")
             TxtAlic.Enabled = True
             TxtImpRet.Enabled = True
         End If
     End With
 
  Else
        A = MsgBox("Debe Ingresar un Proveedor", vbOKOnly, "ERROR")
 End If

End Sub

Private Sub CmdAgregar_Click()
    
    ' Calculo Total Retencion
        'If TxtImpRet.Text = "" Then TxtImpRet.Text = 0
        totalrete = TxtImpRet.text
        'MsgBox (totalrete)
        TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#,###,###,#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
        total = Format((total + totalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
    
    Call LlenaGrilla

End Sub

Private Sub CmdBuscar_Click()
    Call buscofacturas
End Sub
Private Sub buscofacturas()

On Error GoTo Error_Handler

    ruta = App.Path & "\Padron.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaCProv", dbOpenDynaset)
    
'    Set BaseSPC = OpenDatabase(App.Path & "\Padron.mdb")
    
    CodigoClie = Val(TxtCodProv.text)

    rstFacturaC.FindFirst "CodProv = " + Str(CodigoClie)
'     rstFacturaC.FindFirst "CodProv = " + str(CodigoClie),  order by NroFactura
'    rstFacturaD.FindFirst "NroFactura= " + Str(numDoc)
    
    
    FG1.Rows = 2
    FG1.Clear
    FG1.Visible = True
    
    Call SeteoGrilla


    linea2 = 1
    Do While Not rstFacturaC.NoMatch
        FG1.AddItem " "
        FG1.Row = linea2
        FG1.Col = 0
        FG1.text = rstFacturaC.Fields!NroFactura
        FG1.Col = 1
        FG1.text = rstFacturaC.Fields!FechaFactura
        FG1.Col = 2
        FG1.text = rstFacturaC.Fields!TipoFactura
        FG1.Col = 3
        FG1.text = rstFacturaC.Fields!SubTotalFactura
        FG1.Col = 4
        FG1.text = rstFacturaC.Fields!totalIva
        FG1.Col = 5
        FG1.text = rstFacturaC.Fields!percepcion
        FG1.Col = 6
        FG1.text = rstFacturaC.Fields!TotalFactura
        FG1.Col = 8
        FG1.text = 1
        
        
        linea2 = linea2 + 1
        rstFacturaC.FindNext "CodProv = " + Str(CodigoClie)
    Loop
    
Error_Handler:

    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If

    Exit Sub

CmdBuscar.Enabled = False
End Sub


Private Sub CmdExit_Click()
    Unload Me
    
    total = 0
    TOTALRETENCION = 0
    
End Sub

Private Sub CmdNew_Click()

    Call CleanDatos2
    Call SeteoGrilla
    
    Set tUltNums = Padron.OpenRecordset("UltNums", dbOpenTable)
    
    With tUltNums
        .Index = "PrimaryKey"
        .Seek "=", "PAGO"
        If Not .NoMatch Then
            TxtPayNr.text = !UltNum + 1
        End If
        .Close
    End With
    
        

End Sub

Private Sub CmdPrint_Click()

With Printer
    
    On Error GoTo CapturaErrores
    
    'Seteo de Tamaño de Papel
        '.PaperSize = 9
        .ScaleHeight = 297
        .ScaleWidth = 210
    
    'Titulo del comprobante
        Printer.Line (10, 7)-(200, 7)
        .CurrentX = 60
        .CurrentY = Printer.CurrentY + 2
        .Font = "Arial"
        .FontSize = 12
        .FontBold = True
        '.FontUnderline = True
        Printer.Print "Certificado de Retención de Ingresos Brutos"
        .CurrentX = 80
        
        If TxtPayNr.text = "" Then
            MsgBox "Debe Ingresar un Nro de Pago Antes de Imprimir", vbCritical, "ERROR !"
            TxtPayNr.SetFocus
            Exit Sub
        End If
        
        Printer.Print "Nro: " + TxtPayNr.text
        .CurrentY = Printer.CurrentY + 1
        Printer.Line (10, 20)-(200, 20)
                
    'Datos del Agente de Retención
        .CurrentX = 10
        .CurrentY = .CurrentY + 4
        .Font = "Arial"
        .FontSize = 10
        .FontBold = True
        .FontUnderline = False
        Printer.Print "Referencias del Agente de Retención"
        
        .CurrentY = .CurrentY + 2
        
            'Empresa
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Empresa:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print "Quilplac S.A."
        
            'Domicilio
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Domicilio:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print "Andrés Baranda 520"
        
            'Localidad
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Localidad:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print "Quilmes" + Chr(9) + "CP: B1878DKL"
        
            'CUIT
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "C.U.I.T:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print "30-70843254-3"
            
            'IIBB
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Nº IIBB" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print "30-70843254-3"
        
    'Datos del Proveedor
        Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
        CodiProv = Val(cmbCodProv(0).text)
        sql_prov = "SELECT * FROM Proveedores WHERE CodProv = " & CodiProv & "  Order By CodProv"
        Set Prov = db.OpenRecordset(sql_prov, dbOpenDynaset)
    
        If Not Prov.EOF Then
            ProvCod = Prov!CodProv
            ProvName = Prov!NombreProv
            ProvCuit = Prov!CUIT
            ProvDireccion = Prov!direccion
            ProvLocalidad = Prov!localidad
            ProvProvincia = Prov!provincia
            ProvCP = Prov!CP
        End If
    
        Prov.Close
        
        .CurrentX = 10
        .CurrentY = .CurrentY + 4
        .Font = "Arial"
        .FontSize = 10
        .FontBold = True
        .FontUnderline = False
        Printer.Print "Referencias del Proveedor"
        
        .CurrentY = .CurrentY + 2
        
            'Proveedor
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Proveedor:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print Str(ProvCod) + Chr(9) + ProvName
        
            'Domicilio
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Domicilio:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print ProvDireccion
        
            'Localidad
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Localidad:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print ProvLocalidad + Chr(9) + "CP: " + ProvCP
        
            'Provincia
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "Provincia:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print ProvProvincia
            
            'CUIT
                .CurrentX = 15
                .Font = "Arial"
                .FontSize = 8
                .FontBold = False
                .FontUnderline = False
                Printer.Print "C.U.I.T:" + Chr(9);
                .FontBold = False
                .Font = "MS Sans Serif"
                .FontSize = 8
                Printer.Print ProvCuit
    
     'Nota de constancia
        .CurrentY = .CurrentY + 4
        .CurrentX = 20
        .FontSize = 7
        .FontBold = True
        Printer.Print "Dejamos constancia de haber efectuado en la fecha la retención de Ingresos Brutos sobre el/los documento/s que detallamos a continuación, "
        .CurrentX = 20
        Printer.Print "por los importes que en cada caso se indican. ";
        
        Printer.Line (10, (.CurrentY + 6))-(200, (.CurrentY + 6))
     
     
     'Detalle de la Retención Efectuada
        .CurrentY = .CurrentY + 6
        .CurrentX = 15
        .Font = "Arial"
        .FontSize = 8
        .FontBold = False
        .FontUnderline = False
        Printer.Print "Detalle de la Retención Efectuada"
       
            'Encabezado del Detalle
                .CurrentY = .CurrentY + 2
                .CurrentX = 21
                .Font = "Arial"
                .FontBold = True
                .FontSize = 7
                Printer.Print "Nro Factura";
                Printer.CurrentX = 45
                Printer.Print "Fecha Doc.";
                Printer.CurrentX = 75
                Printer.Print "Imp. Comprob";
                Printer.CurrentX = 103
                Printer.Print "Imp. Suj. a Ret";
                Printer.CurrentX = 130
                Printer.Print "Imp. IVA";
                Printer.CurrentX = 155
                Printer.Print "Imp. Retenido";
                Printer.CurrentX = 185
                Printer.Print "% Ret.";
                Printer.Line (20, (.CurrentY + 4))-(195, (.CurrentY + 4))
                
            'Lineas del Detalle
                Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
                NroPago = Val(TxtPayNr.text)
                sql_lp = "SELECT * FROM LineasPagoRet WHERE NroPago = " & NroPago & "  Order By NroPago, Item"
                Set LineasP = db.OpenRecordset(sql_lp, dbOpenDynaset)
                
                With LineasP
                    .MoveFirst
                
                Printer.CurrentY = Printer.CurrentY + 2
                
                 While Not .EOF
                    Printer.CurrentX = 21
                    Printer.FontBold = False
                    Printer.Print !NroFactProv;
                    Printer.CurrentX = 46
                    Printer.Print Format(!FechaFactProv, "DD/MM/YYYY") + Chr(9) + Chr(9);
                    Printer.CurrentX = 77
                    Printer.Print Format(!TotalFac, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 105
                    Printer.Print Format(!BaseImponible, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 130
                    Printer.Print Format(!importeIva, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 160
                    Printer.Print Format(!ImporteRet, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 186
                    Printer.Print Format(!Alicuota, "Standard") + Chr(9) + Chr(9);
                    
                    'TOTAL = TOTAL + (!BaseImponible + !ImporteIva + !ImporteRet)
                    total = total + (!BaseImponible + !importeIva)
                    TOTALRETENCION = TOTALRETENCION + !ImporteRet
                    
                    .MoveNext
                    Printer.CurrentY = Printer.CurrentY + 3
                 Wend
                
                Printer.Line (20, (Printer.CurrentY + 4))-(195, (Printer.CurrentY + 4))
                
               ' Printer.CurrentY = Printer.CurrentY + 4
               ' Printer.CurrentX = 77
               ' Printer.FontBold = True
               ' Printer.Print "Total Retenido: " + Format(TOTALRETENCION, "Standard")
               ' Printer.CurrentY = Printer.CurrentY + 4
                'Printer.CurrentX = 77
               ' Printer.Print "Total General: " + Format(TOTAL, "Standard")
                
                Printer.CurrentY = Printer.CurrentY + 4
                Printer.CurrentX = 143
                Printer.FontBold = True
                Printer.Print "Total Retenido: " + Format(TOTALRETENCION, "Standard");
                Printer.CurrentX = 60
                Printer.Print "Total General: " + Format(total, "Standard")
                
                
                Printer.CurrentY = Printer.CurrentY + 10
                Printer.CurrentX = 40
                
                Printer.FontBold = True
                
                
                Printer.Print "Lugar y Fecha: " + Chr(9) + "Quilmes, " + Format(Date, "dd") + " de " + Format(Date, "mmmm") + " de " + Format(Date, "yyyy")
               ' Printer.Print "Lugar y Fecha: " + Chr(9) + "Quilmes, " + Format(Date, "dd-mmm-yyyy")
                
                If Printer.CurrentY < 220 Then
                    Printer.CurrentY = 220
                    Printer.CurrentX = 160
                    
                    Printer.Line (120, Printer.CurrentY)-(170, Printer.CurrentY)
                    Printer.CurrentY = Printer.CurrentY + 2
                    Printer.CurrentX = 135
                    Printer.Print "QUILPLAC S.A."
                Else
                    Printer.CurrentY = Printer.CurrentY + 5
                    Printer.CurrentY = Printer.CurrentY + 2
                    Printer.CurrentX = 160
                    Printer.Line (120, Printer.CurrentY)-(170, Printer.CurrentY)
                    Printer.CurrentX = 135
                    Printer.Print "QUILPLAC S.A."
                End If
                
                End With
                
    .EndDoc
    
    TOTALRETENCION = 0
    total = 0
    
End With
    
CapturaErrores:
    'If Err = 321 Then
    'End If

End Sub

Private Sub CmdSave_Click()

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rstPagoRet = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
    Set db2 = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rstLineaRet = db2.OpenRecordset("LineasPagoRet", dbOpenDynaset)
    Set tUltNums = db.OpenRecordset("UltNums", dbOpenTable)
    
   
        rstPagoRet.AddNew
        rstPagoRet.Fields!NroPago = TxtPayNr.text
        rstPagoRet.Fields!FechaPago = TxtPayDate.text
        
        
        rstPagoRet.Fields!CodProv = cmbCodProv(0).text
        rstPagoRet.Fields!NombreProv = TxtProvName.text
        rstPagoRet.Fields!CUIT = TxtCUIT.text
        rstPagoRet.Fields!TotalReten = TxtTotalRetencion.text
        rstPagoRet.Fields!TotalPago = txtTotal.text
        rstPagoRet.Update
        
        FG1.Col = 0
        FG1.Row = 1
        Filas = FG1.Rows
        linea = 1
        
        'Do While linea < filas
        Do While (FG1.TextMatrix(FG1.Row, 0) <> "")
           FG1.Row = linea
              rstLineaRet.AddNew
              
              FG1.Col = 0
              rstLineaRet.Fields!item = Val(FG1.TextMatrix(FG1.Row, FG1.Col))
              'Val (FG1.Text)
              
              FG1.Col = 1
              rstLineaRet.Fields!NroPago = Val(FG1.TextMatrix(FG1.Row, FG1.Col))
              'Val(FG1.Text)
                      
              FG1.Col = 2
              rstLineaRet.Fields!NroFactProv = FG1.TextMatrix(FG1.Row, FG1.Col)
              'FG1.Text
              
              FG1.Col = 3
              'MsgBox (FG1.TextMatrix(FG1.Row, FG1.Col))
              rstLineaRet.Fields!FechaFactProv = FG1.TextMatrix(FG1.Row, FG1.Col)
              
              FG1.Col = 4
              rstLineaRet.Fields!TotalFac = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              FG1.Col = 5
              rstLineaRet.Fields!BaseImponible = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              FG1.Col = 6
              rstLineaRet.Fields!Alicuota = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              FG1.Col = 7
              rstLineaRet.Fields!ImporteRet = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              FG1.Col = 8
              rstLineaRet.Fields!importeIva = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              FG1.Col = 9
              rstLineaRet.Fields!TotalLineaFactura = FG1.TextMatrix(FG1.Row, FG1.Col)
              'Val(FG1.Text)
              
              rstLineaRet.Update
           'End If
              
            linea = linea + 1
            FG1.Row = linea
        Loop
        
        MsgBox "Pago Grabado Con Exito", vbInformation, "QUILPLAC SA"
        
        With tUltNums
            .Index = "PrimaryKey"
            .Seek "=", "PAGO"
            If Not .NoMatch Then
                .Edit
                    !UltNum = CInt(TxtPayNr.text)
                .Update
            End If
            .Close
        End With

        
        Call SeteoGrilla
        linea = 0
        item = 0
    
   

End Sub

Private Sub FG1_Click()

    Dim importefac As Double
    Dim totalFactu As Double
    Dim importeIva As Double
    Dim totalIva As Double
    Dim tRetenciones As Double
    Dim tTotales As Double
  


    If FG1.Col = 7 Then
    
        FG1.Col = 4
        importeIva = FG1.text

        FG1.Col = 6
        importetfac = FG1.text
        
        tRetenciones = TxtImpRet.text
        tTotles = TxtTotLinea.text
        
        FG1.Col = 8
        si = FG1.text
        If si = 1 Then
            totalIva = Val(TxtIVA.text) + importeIva
            totalFactu = Val(TxtTotFac.text) + importetfac
            
            FG1.Col = 7
            FG1.CellPictureAlignment = flexAlignCenterCenter
            Set FG1.CellPicture = LoadPicture(App.Path & "\Iconos\APPROV16.BMP")
            FG1.Col = 8
            FG1.text = 0
        Else
            totalIva = Val(TxtIVA.text) - importeIva
            totalFactu = Val(TxtTotFac.text) - importetfac
            totalesRetenciones = Val(TxtImpRet.text) - tRetenciones
            totalesFacturas = Val(TxtTotLinea.text) - tTotales
            FG1.Col = 7
            FG1.CellPictureAlignment = flexAlignCenterCenter
            Set FG1.CellPicture = LoadPicture(App.Path & "\Iconos\VACIO.BMP")
            FG1.Col = 8
            FG1.text = 1
        End If
        
    End If
    
    TxtTotFac.text = totalFactu
    TxtIVA.text = totalIva
   

End Sub

Private Sub Form_Load()
    
    FormComprobanteIIBBFacturas.Height = 7185
    FormComprobanteIIBBFacturas.Width = 13290
    FormComprobanteIIBBFacturas.Top = 300
    FormComprobanteIIBBFacturas.Left = 300
    
    Set Padron = OpenDatabase(App.Path & "\Padron.mdb")
    Set Provs = Padron.OpenRecordset("Proveedores")
    Set tUltNums = Padron.OpenRecordset("UltNums", dbOpenTable)
    
    
    With Provs
        .MoveFirst
        While Not .EOF
           cmbCodProv(0).AddItem (!CodProv)
           cmbCodProv(1).AddItem (!NombreProv)
           .MoveNext
        Wend
    End With
    
    With tUltNums
        .Index = "PrimaryKey"
        .Seek "=", "PAGO"
        If Not .NoMatch Then
            TxtPayNr.text = !UltNum + 1
        End If
        .Close
    End With
    
        
    
    Fila = 0
    Columna = 0
    item = 0
        
    Call SeteoGrilla
   
End Sub

Sub SeteoGrilla()
    
    'FG1.AutoSizeMode = klexAutoSizeColWidth
    FG1.Row = 0
    FG1.Col = 0
        
    FG1.Col = 0
    FG1.ColWidth(0) = 1500
    FG1.text = "Nº Fact. Proveedor"
    FG1.ColAlignment(0) = flexAlignCenterCenter
    
    FG1.Col = 1
    FG1.ColWidth(1) = 1000
    FG1.text = "Fecha FP"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1000
    FG1.text = "Tipo FP"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1400
    FG1.text = "Sub Total"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 1400
    FG1.text = "IVA"
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    FG1.Col = 5
    FG1.ColWidth(5) = 1400
    FG1.text = "Percepcion"
    FG1.ColAlignment(5) = flexAlignCenterCenter
    
    FG1.Col = 6
    FG1.ColWidth(6) = 1400
    FG1.text = "Importe Total"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 1400
    FG1.text = "Seleccion"
    FG1.ColAlignment(7) = flexAlignCenterCenter
    
    
    FG1.Col = 8
    FG1.ColWidth(8) = 0
    FG1.text = ""
    FG1.ColAlignment(8) = flexAlignCenterCenter
        
    Columna = 0
    Fila = 0
    linea = 0
    
End Sub

Private Sub TxtBaseI_Change()
        Dim totalrete As Double
            
    'Calculo el importe total
        BaseI = TxtBaseI.text
        TotFac = (BaseI * 1.21)
        TxtTotFac.text = Format(TotFac, "#,###,###,#0.00")
    
        Alicuota = Format(TxtAlic.text, "#,###,###,#0.00")
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
        End If
        
    'Calculo Importe Iva
        iva = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        
    ' Calculo Total Retencion
        totalrete = TxtImpRet.text
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
   '     TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
        
        
        '******************************************************

    ' Calculo Total Retencion
        'If TxtImpRet.Text = "" Then TxtImpRet.Text = 0
        totalrete = TxtImpRet.text
        'MsgBox (totalrete)
        TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#,###,###,#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
        total = Format((total + totalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
   
        txtTotal.text = Format(total, "#,###,###,#0.00")
        TxtTotalRetencion.text = Format(TOTALRETENCION, "#,###,###,#0.00")
    
End Sub

Private Sub TxtBaseI_GotFocus()

    TxtBaseI.SelLength = Len(TxtBaseI.text)

End Sub

Private Sub TxtBaseI_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtBaseI_LostFocus()
    
    Dim totalrete As Double
            
    'Calculo el importe total
        BaseI = TxtBaseI.text
        TotFac = (BaseI * 1.21)
        TxtTotFac.text = Format(TotFac, "#,###,###,#0.00")
    
        Alicuota = Format(TxtAlic.text, "#,###,###,#0.00")
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
        End If
        
    'Calculo Importe Iva
        iva = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        
    ' Calculo Total Retencion
        totalrete = TxtImpRet.text
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
   '     TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
        
        
        '******************************************************

    ' Calculo Total Retencion
        'If TxtImpRet.Text = "" Then TxtImpRet.Text = 0
        totalrete = TxtImpRet.text
        'MsgBox (totalrete)
        TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#,###,###,#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
        total = Format((total + totalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
   
        txtTotal.text = Format(total, "#,###,###,#0.00")
        TxtTotalRetencion.text = Format(TOTALRETENCION, "#,###,###,#0.00")
      

End Sub


Private Sub TxtFF_GotFocus()

    TxtFF.SelLength = Len(TxtFF.text)

End Sub

Private Sub TxtFF_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtFF_LostFocus()

    If Not IsDate(TxtFF.text) Then
    
        MsgBox "Formato o Fecha Incorrecta", vbCritical, "ERROR !"
        TxtFF.text = Format(Date, "DD/MM/YYYY")
    
    End If
    

End Sub


Private Sub TxtCodProv_Change()
    If TxtCodProv.text <> "" Then
        CmdBuscar.Enabled = True
    End If
End Sub

Private Sub TxtImpRet_KeyPress(KeyAscii As Integer)

    If KeyAscii = 46 Then
        KeyAscii = 44
    End If

    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtImpRet_LostFocus()

Dim totalrete As Double
            
    'Calculo el importe total
        BaseI = TxtBaseI.text
        TotFac = (BaseI * 1.21)
        TxtTotFac.text = Format(TotFac, "#,###,###,#0.00")
    
        Alicuota = Format(TxtAlic.text, "#,###,###,#0.00")
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
        End If
        
    'Calculo Importe Iva
        iva = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        
    ' Calculo Total Retencion
        totalrete = TxtImpRet.text
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
   '     TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False


End Sub

Private Sub TxtIva_GotFocus()

    TxtIVA.SelLength = Len(TxtIVA.text)
    
End Sub


Private Sub TxtIva_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtIva_LostFocus()
    
    Dim totalrete As Double
            
    'Calculo el importe total
        'BaseI = (TxtIva.Text / 0.21)
        If (iva = "") Or (iva = 0) Then iva = Format(TxtIVA.text, "#,###,###,#0.00")
        BaseI = (iva / 0.21)
        
        TxtBaseI.text = Format(BaseI, "#,###,###,#0.00")
        TotFac = (BaseI * 1.21)
        TxtTotFac.text = Format(TotFac, "#,###,###,#0.00")
    
        Alicuota = Format(TxtAlic.text, "#,###,###,#0.00")
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
        End If
        
    'Calculo Importe Iva
        iva = TxtIVA.text
        'IVA = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        
    ' Calculo Total Retencion
        totalrete = TxtImpRet.text
    
    'Calculo Total Linea
        totalLinea = BaseI + iva
   '     TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
        TxtTotLinea.text = Format(totalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False

    FormComprobanteIIBBFacturas.Refresh

End Sub

Private Sub TxtNroFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtPayDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
           KeyAscii = 0
           Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub TxtPayDate_LostFocus()

    If Not IsDate(TxtPayDate.text) Then
        MsgBox "Formato de Fecha Incorrecto", vbCritical, "ERROR !"
        TxtPayDate.text = Format(Date, "DD/MM/YYYY")
    End If
End Sub

Private Sub TxtPayNr_GotFocus()

    TxtPayNr.SelLength = Len(TxtPayNr.text)

End Sub

Private Sub TxtPayNr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub TxtPayNr_LostFocus()

On Error GoTo CapturaErrores

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rstPagoRet = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
    If (TxtPayNr.text <> "") Then
        numeroPago = Val(TxtPayNr.text)
        rstPagoRet.FindFirst "NroPago= " + Str(numeroPago)
        
        If rstPagoRet.Fields!NroPago = Val(TxtPayNr.text) Then
            respuesta = MsgBox("Numero de PagoExistente", vbCritical, " ")
            TxtPayNr.text = ""
            TxtPayNr.SetFocus
        End If
    Else
        'respuesta = MsgBox("Debe Ingresar un Nro de Pago", vbCritical, " ")
    End If

CapturaErrores:
    If Err = 321 Then
    End If
End Sub



Private Sub TxtTotFac_Change()

    Dim totalrete As Double
            
    'Formateo a moneda el importe ingresado de la factura
        TotFac = TxtTotFac.text
        TxtTotFac.text = Format(TxtTotFac.text, "#,###,###,#0.00")
    
    'Calculo la Base Imponible
        BaseI = TotFac / 1.21
        TxtBaseI.text = Format((TxtTotFac.text / 1.21), "#,###,###,#0.00")
        'TxtBaseI.Enabled = False
    
    'Busco Alicuota
       ' Set TPadron = Padron.OpenRecordset("Padron", dbOpenTable)
        
       ' TPadron.Index = "CUIT"
        
       ' With TPadron
       '     .Seek "=", TxtCUIT.Text
       '     If Not .NoMatch() Then
       '         Alicuota = !AlicuotaRetencion
        '        TxtAlic.Text = Format(!AlicuotaRetencion, "#,###,###,#0.00")
        '        TxtAlic.Enabled = False
        '    End If
        'End With
        
       ' TPadron.Close
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         '   TxtImpRet.Enabled = False
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
          '  TxtImpRet.Enabled = False
        End If
        
    'Calculo Importe Iva
        iva = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        'TxtIva.Enabled = False
        
' anulado el 21/09/2014
        
    ' Calculo Total Retencion
        'If TxtImpRet.Text = "" Then TxtImpRet.Text = 0
'        totalrete = TxtImpRet.Text
        'MsgBox (totalrete)
 '       TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#,###,###,#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
 '       TotalLinea = BaseI + IVA
 '       TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
 '       TxtTotLinea.Text = Format(TotalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
     totalesRetenciones = Val(TxtImpRet.text) + tRetenciones
            totalesFacturas = Val(TxtTotLinea.text) + tTotales
    
    TxtTotalRetencion.text = totalesRetenciones
    txtTotal.text = totalesFacturas

End Sub

Private Sub TxtTotFac_GotFocus()

    TxtTotFac.SelLength = Len(TxtTotFac.text)

End Sub

Private Sub TxtTotFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


Private Sub TxtTotFac_LostFocus()
    
    Dim totalrete As Double
            
    'Formateo a moneda el importe ingresado de la factura
        TotFac = TxtTotFac.text
        TxtTotFac.text = Format(TxtTotFac.text, "#,###,###,#0.00")
    
    'Calculo la Base Imponible
        BaseI = TotFac / 1.21
        TxtBaseI.text = Format((TxtTotFac.text / 1.21), "#,###,###,#0.00")
        'TxtBaseI.Enabled = False
    
    'Busco Alicuota
       ' Set TPadron = Padron.OpenRecordset("Padron", dbOpenTable)
        
       ' TPadron.Index = "CUIT"
        
       ' With TPadron
       '     .Seek "=", TxtCUIT.Text
       '     If Not .NoMatch() Then
       '         Alicuota = !AlicuotaRetencion
        '        TxtAlic.Text = Format(!AlicuotaRetencion, "#,###,###,#0.00")
        '        TxtAlic.Enabled = False
        '    End If
        'End With
        
       ' TPadron.Close
        
    'Calculo Importe Retención
        If BaseI > 50 Then
            ImpRet = BaseI * (Alicuota / 100)
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
         '   TxtImpRet.Enabled = False
         Else
            ImpRet = 0
            TxtImpRet.text = Format(ImpRet, "#,###,###,#0.00")
          '  TxtImpRet.Enabled = False
        End If
        
    'Calculo Importe Iva
        iva = (BaseI * 0.21)
        TxtIVA.text = Format(iva, "#,###,###,#0.00")
        'TxtIva.Enabled = False
        
' anulado el 21/09/2014
        
    ' Calculo Total Retencion
        'If TxtImpRet.Text = "" Then TxtImpRet.Text = 0
'        totalrete = TxtImpRet.Text
        'MsgBox (totalrete)
 '       TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#,###,###,#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
 '       TotalLinea = BaseI + IVA
 '       TOTAL = Format((TOTAL + TotalLinea), "#,###,###,#0.00")
 '       TxtTotLinea.Text = Format(TotalLinea, "#,###,###,#0.00")
        'TxtTotLinea.Enabled = False
        
        

End Sub

Private Sub TxtTotFac_Validate(Cancel As Boolean)

    If TxtTotFac.text = "" Then TxtTotFac.text = 0

End Sub



Private Sub TxtTotLinea_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 46 Then
        KeyAscii = 44
    End If
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If

End Sub


