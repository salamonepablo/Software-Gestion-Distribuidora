VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormComprobanteIIBB 
   Caption         =   "Generar Comprobante de Retencion IIBB"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   13185
   Begin VB.Frame Frame1 
      Caption         =   "Pago a Proveedores / Comprobante de Retención IIBB"
      Height          =   5535
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   12975
      Begin VB.TextBox TxtTotalRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   15
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox TxtProvName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         Top             =   750
         Width           =   4215
      End
      Begin VB.TextBox TxtPayDate 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3240
         TabIndex        =   22
         Top             =   3960
         Width           =   9615
         Begin VB.CommandButton CmdNew 
            Caption         =   "&Nuevo"
            Height          =   735
            Left            =   4560
            Picture         =   "FormComprobanteIIBB.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmndExcel 
            Caption         =   "&Excel"
            Height          =   735
            Left            =   3240
            Picture         =   "FormComprobanteIIBB.frx":00FA
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Imprimir"
            Height          =   735
            Left            =   1920
            Picture         =   "FormComprobanteIIBB.frx":053C
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdSave 
            Caption         =   "&Guardar"
            Height          =   735
            Left            =   600
            Picture         =   "FormComprobanteIIBB.frx":0636
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "&Salir"
            Height          =   735
            Left            =   8280
            Picture         =   "FormComprobanteIIBB.frx":0730
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtCUIT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   4
         Top             =   750
         Width           =   1215
      End
      Begin VB.ComboBox CmbCodProv 
         Height          =   315
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtTOTAL 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1200
         TabIndex        =   16
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox TxtNroFac 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxtFF 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3000
         TabIndex        =   6
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxtTotFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4560
         TabIndex        =   7
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   2760
         Width           =   735
      End
      Begin VB.TextBox TxtBaseI 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6120
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxtAlic 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7680
         TabIndex        =   9
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox TxtImpRet 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8520
         TabIndex        =   10
         Top             =   1560
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
         Left            =   9960
         TabIndex        =   11
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TxtTotLinea 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9960
         TabIndex        =   12
         Top             =   2040
         Width           =   1335
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   1215
         Left            =   960
         TabIndex        =   14
         Top             =   2640
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
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
         Left            =   1080
         TabIndex        =   37
         Top             =   3840
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
         Left            =   3960
         TabIndex        =   36
         Top             =   480
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
         TabIndex        =   35
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
         TabIndex        =   34
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
         Left            =   360
         TabIndex        =   33
         Top             =   2280
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
         TabIndex        =   32
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
         Left            =   5520
         TabIndex        =   31
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
         Left            =   1080
         TabIndex        =   30
         Top             =   4440
         Width           =   1230
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro Factura"
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
         TabIndex        =   29
         Top             =   1320
         Width           =   1020
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura"
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
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   1245
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
         Left            =   4440
         TabIndex        =   27
         Top             =   1320
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
         Left            =   6000
         TabIndex        =   26
         Top             =   1320
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
         Left            =   7560
         TabIndex        =   25
         Top             =   1320
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
         Left            =   9840
         TabIndex        =   24
         Top             =   1320
         Width           =   1005
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
         Left            =   8760
         TabIndex        =   23
         Top             =   2040
         Width           =   1065
      End
   End
End
Attribute VB_Name = "FormComprobanteIIBB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CleanDatos()

    TxtNroFac.Text = ""
    TxtFF.Text = ""
    TxtTotFac.Text = ""
    TxtBaseI.Text = ""
    TxtImpRet.Text = ""
    TxtAlic.Text = ""
    TxtIva.Text = ""
    TxtTotLinea.Text = ""
    
    
    
    TxtNroFac.SetFocus
    
End Sub

Private Sub CleanDatos2()

    TxtPayNr.Text = ""
  '  TxtPayDate.Text = ""
  '  CmbCodProv.Text = ""
    TxtProvName.Text = ""
    TxtCUIT.Text = ""
    TxtNroFac.Text = ""
    TxtFF.Text = ""
    TxtTotFac.Text = ""
    TxtBaseI.Text = ""
    TxtImpRet.Text = ""
    TxtAlic.Text = ""
    TxtIva.Text = ""
    TxtTotLinea.Text = ""
    TxtTotalRetencion.Text = ""
    TxtTOTAL.Text = ""
    FG1.Clear
    
    TxtPayNr.SetFocus

End Sub

Private Sub GrabarPago(NroPago As String, CodProv As Integer, NFac As String, FF As Date, TotFac As Double, ImpRet As Double, IVA As Double, TotalLinea As Double)
    
End Sub


Private Sub LlenaGrilla()
    
    Item = Item + 1
    FG1.Row = Fila + 1
    FG1.Col = Columna
    FG1.Text = Item
    FG1.Col = 1
    
    FG1.Text = Format(TxtPayNr.Text, "")
    FG1.Col = 2
    
    FG1.TextMatrix(FG1.Row, FG1.Col) = TxtNroFac.Text
    FG1.Col = 3
    
    FG1.Text = TxtFF.Text
    FG1.Col = 4
    
    FG1.Text = TxtTotFac.Text
    FG1.Col = 5
    
    FG1.Text = TxtBaseI.Text
    FG1.Col = 6
    
    FG1.Text = TxtAlic.Text
    FG1.Col = 7
    
    FG1.Text = TxtImpRet.Text
    FG1.Col = 8
    
    FG1.Text = TxtIva.Text
    FG1.Col = 9

    FG1.Text = TxtTotLinea.Text
    
    TxtTOTAL.Text = Format(TOTAL, "#0.00")
    TxtTotalRetencion.Text = Format(TOTALRETENCION, "#0.00")
    
    Columna = 0
    
    Fila = Fila + 1
    
    FG1.Rows = FG1.Rows + 1
    
    Call CleanDatos
    
End Sub


Private Sub Agregar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If

End Sub




Private Sub CmbCodProv_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
    'Call busco
 End If

End Sub

Private Sub busco()

   Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
   Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
    
   CodiProv = Val(CmbCodProv.Text)
      
    rst.FindFirst "CodProv= " + Str(CodiProv)
    
    TxtProvName.Text = rst.Fields!NombreProv
    TxtCUIT.Text = rst.Fields!Cuit
   
    
    Set TPadron = Padron.OpenRecordset("Padron", dbOpenTable)
    
    TPadron.Index = "CUIT"
    
    With TPadron
        TPadron.Seek "=", TxtCUIT.Text
        If .NoMatch = False Then
            Alicuota = !AlicuotaRetencion
        End If
    End With

End Sub

Private Sub CmbCodProv_LostFocus()

    Call busco
      
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

Private Sub CmdAgregar_Click()
    
    Call LlenaGrilla

End Sub

Private Sub CmdExit_Click()
    Unload Me
    
    TOTAL = 0
    TOTALRETENCION = 0
    
End Sub

Private Sub CmdNew_Click()

    Call CleanDatos2
    Call SeteoGrilla

End Sub

Private Sub CmdPrint_Click()

With Printer
    
    On Error GoTo CapturaErrores
    
    TotalFacturado = 0
    TotalRetenido = 0
    
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
        
        If TxtPayNr.Text = "" Then
            MsgBox "Debe Ingresar un Nro de Pago Antes de Imprimir", vbCritical, "ERROR !"
            TxtPayNr.SetFocus
            Exit Sub
        End If
        
        Printer.Print "Nro: " + TxtPayNr.Text
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
        Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
        CodiProv = Val(CmbCodProv.Text)
        sql_prov = "SELECT * FROM Proveedores WHERE CodProv = " & CodiProv & "  Order By CodProv"
        Set prov = db.OpenRecordset(sql_prov, dbOpenDynaset)
    
        If Not prov.EOF Then
            ProvCod = prov!CodProv
            ProvName = prov!NombreProv
            ProvCuit = prov!Cuit
            ProvDireccion = prov!Direccion
            ProvLocalidad = prov!Localidad
            ProvProvincia = prov!Provincia
            ProvCP = prov!Cp
        End If
    
        prov.Close
        
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
        .FontBold = True
        .FontUnderline = False
        Printer.Print "Detalle de la Retención Efectuada";
        Printer.CurrentX = 145
        Printer.Print "Fecha de Retención: ";
        Printer.CurrentX = 175
        Printer.FontBold = False
        Printer.Print Format(Date, "DD/MM/YYYY")
        
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
                Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
                NroPago = Val(TxtPayNr.Text)
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
                    Printer.Print Format(!ImporteIva, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 160
                    Printer.Print Format(!ImporteRet, "Standard") + Chr(9) + Chr(9);
                    Printer.CurrentX = 186
                    Printer.Print Format(!Alicuota, "Standard") + Chr(9) + Chr(9);
                    
                    TotalFacturado = TotalFacturado + (!BaseImponible + !ImporteIva + !ImporteRet)
                    TotalRetenido = TotalRetenido + !ImporteRet
                    
                    .MoveNext
                    Printer.CurrentY = Printer.CurrentY + 3
                 Wend
                
                Printer.Line (20, (Printer.CurrentY + 4))-(195, (Printer.CurrentY + 4))
                
                Printer.CurrentY = Printer.CurrentY + 4
                Printer.CurrentX = 143
                Printer.FontBold = True
                Printer.Print "Total Retenido: " + Format(TotalRetenido, "Standard");
                Printer.CurrentX = 60
                Printer.Print "Total General: " + Format(TotalFacturado, "Standard")
                
                Printer.CurrentY = Printer.CurrentY + 10
                Printer.CurrentX = 40
                
                Printer.FontBold = True
                
                Printer.FontBold = True
                Printer.Print "Lugar y Fecha: ";
                Printer.FontBold = False
                Printer.Print "  Quilmes, " + Format(Date, "dd-mmm-yyyy")

                If Printer.CurrentY < 220 Then
                    Printer.CurrentY = 220
                    Printer.CurrentX = 160
                    
                    Printer.Line (120, Printer.CurrentY)-(170, Printer.CurrentY)
                    Printer.CurrentY = Printer.CurrentY + 2
                    Printer.CurrentX = 135
                    Printer.FontBold = True
                    Printer.Print "QUILPLAC S.A."
                Else
                    Printer.CurrentY = Printer.CurrentY + 5
                    Printer.CurrentY = Printer.CurrentY + 2
                    Printer.CurrentX = 160
                    Printer.Line (120, Printer.CurrentY)-(170, Printer.CurrentY)
                    Printer.CurrentX = 135
                    Printer.FontBold = True
                    Printer.Print "QUILPLAC S.A."
                End If
                
                End With
                                
                TotalFacturado = 0
                TotalRetenido = 0
                
    .EndDoc
    
End With
    
CapturaErrores:
    'If Err = 321 Then
    'End If

End Sub

Private Sub CmdSave_Click()

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rstPagoRet = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
    Set db2 = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rstLineaRet = db2.OpenRecordset("LineasPagoRet", dbOpenDynaset)
    
   
        rstPagoRet.AddNew
        rstPagoRet.Fields!NroPago = TxtPayNr.Text
        rstPagoRet.Fields!FechaPago = TxtPayDate.Text
        rstPagoRet.Fields!CodProv = CmbCodProv.Text
        rstPagoRet.Fields!NombreProv = TxtProvName.Text
        rstPagoRet.Fields!Cuit = TxtCUIT.Text
        rstPagoRet.Fields!TotalReten = TxtTotalRetencion.Text
        rstPagoRet.Fields!TotalPago = TxtTOTAL.Text
        rstPagoRet.Update
        
        FG1.Col = 0
        FG1.Row = 1
        filas = FG1.Rows
        linea = 1
        
        'Do While linea < filas
        Do While (FG1.TextMatrix(FG1.Row, 0) <> "")
           FG1.Row = linea
              rstLineaRet.AddNew
              
              FG1.Col = 0
              rstLineaRet.Fields!Item = Val(FG1.TextMatrix(FG1.Row, FG1.Col))
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
              rstLineaRet.Fields!ImporteIva = FG1.TextMatrix(FG1.Row, FG1.Col)
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
        
        Call SeteoGrilla
        linea = 0
        Item = 0
    
   

End Sub

Private Sub Form_Load()
    
    FormComprobanteIIBB.Height = 6195
    FormComprobanteIIBB.Width = 13425
    
    Set Padron = OpenDatabase("C:\QuilplacVB\Padron.mdb")
    
    Set Provs = Padron.OpenRecordset("Proveedores")
    
            
     With Provs
        .MoveFirst
        While Not .EOF
           CmbCodProv.AddItem (!CodProv)
           .MoveNext
        Wend
    End With
    
    TxtPayDate.Text = Format(Date, "DD/MM/YYYY")
    
    Fila = 0
    Columna = 0
    Item = 0
    
        
    Call SeteoGrilla
   
End Sub

Sub SeteoGrilla()
    
    'FG1.AutoSizeMode = klexAutoSizeColWidth
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 700
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.Text = "Item"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 700
    FG1.Text = "Nº Pago"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1500
    FG1.Text = "Nº Fact. Proveedor"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1000
    FG1.Text = "Fecha FP"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 1400
    FG1.Text = "Importe Total"
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    FG1.Col = 5
    FG1.ColWidth(5) = 1400
    FG1.Text = "Base Imp."
    FG1.ColAlignment(5) = flexAlignCenterCenter
    
    FG1.Col = 6
    FG1.ColWidth(6) = 700
    FG1.Text = "Alicuota"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 1400
    FG1.Text = "Importe Ret."
    FG1.ColAlignment(7) = flexAlignCenterCenter
    
    FG1.Col = 8
    FG1.ColWidth(8) = 1400
    FG1.Text = "Importe IVA"
    FG1.ColAlignment(8) = flexAlignCenterCenter
    
    FG1.Col = 9
    FG1.ColWidth(9) = 1400
    FG1.Text = "Total Linea"
    FG1.ColAlignment(9) = flexAlignCenterCenter
    
    Columna = 0
    Fila = 0
    linea = 0
    
End Sub

Private Sub TxtFF_GotFocus()

    TxtFF.SelLength = Len(TxtFF.Text)

End Sub

Private Sub TxtFF_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If

End Sub


Private Sub TxtFF_LostFocus()

    If Not IsDate(TxtFF.Text) Then
    
        MsgBox "Formato o Fecha Incorrecta", vbCritical, "ERROR !"
        TxtFF.Text = Format(Date, "DD/MM/YYYY")
    
    End If
    

End Sub


Private Sub TxtNroFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If

End Sub


Private Sub TxtPayDate_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
           KeyAscii = 0
           SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtPayDate_LostFocus()

    If Not IsDate(TxtPayDate.Text) Then
        MsgBox "Formato de Fecha Incorrecto", vbCritical, "ERROR !"
        TxtPayDate.Text = Format(Date, "DD/MM/YYYY")
    End If
End Sub

Private Sub TxtPayNr_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TxtPayNr_LostFocus()

On Error GoTo CapturaErrores

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rstPagoRet = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
    If (TxtPayNr.Text <> "") Then
        NumeroPago = Val(TxtPayNr.Text)
        rstPagoRet.FindFirst "NroPago= " + Str(NumeroPago)
        
        If rstPagoRet.Fields!NroPago = Val(TxtPayNr.Text) Then
            respuesta = MsgBox("Numero de PagoExistente", vbCritical, " ")
            TxtPayNr.Text = ""
            TxtPayNr.SetFocus
        End If
    Else
        'respuesta = MsgBox("Debe Ingresar un Nro de Pago", vbCritical, " ")
    End If

CapturaErrores:
    If Err = 321 Then
    End If
End Sub

Private Sub TxtTotFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       SendKeys "{TAB}"
    End If

End Sub


Private Sub TxtTotFac_LostFocus()
    
    Dim totalrete As Double
            
    'Formateo a moneda el importe ingresado de la factura
        TotFac = TxtTotFac.Text
        TxtTotFac.Text = Format(TxtTotFac.Text, "#0.00")
    
    'Calculo la Base Imponible
        BaseI = TotFac / 1.21
        TxtBaseI.Text = Format((TxtTotFac.Text / 1.21), "#0.00")
        TxtBaseI.Enabled = False
    
    'Busco Alicuota
        Set TPadron = Padron.OpenRecordset("Padron", dbOpenTable)
        
        TPadron.Index = "CUIT"
        
        With TPadron
            .Seek "=", TxtCUIT.Text
            If Not .NoMatch() Then
                Alicuota = !AlicuotaRetencion
                TxtAlic.Text = Format(!AlicuotaRetencion, "#0.00")
                TxtAlic.Enabled = False
            End If
        End With
        
        TPadron.Close
        
    'Calculo Importe Retención
        ImpRet = BaseI * (Alicuota / 100)
        TxtImpRet.Text = Format(ImpRet, "#0.00")
        TxtImpRet.Enabled = False
        
    'Calculo Importe Iva
        IVA = (BaseI * 0.21)
        TxtIva.Text = Format(IVA, "#0.00")
        TxtIva.Enabled = False
        
    ' Calculo Total Retencion
        totalrete = TxtImpRet.Text
        'MsgBox (totalrete)
        TOTALRETENCION = TOTALRETENCION + totalrete
   '    totalrete = Format(TxtImpRet.Text, "#0.00")
        'MsgBox (TOTALRETENCION)
                
    
    'Calculo Total Linea
        TotalLinea = BaseI + ImpRet + IVA
        TOTAL = Format((TOTAL + TotalLinea), "#0.00")
        TxtTotLinea.Text = Format(TotalLinea, "#0.00")
        TxtTotLinea.Enabled = False
        
        

End Sub

Private Sub TxtTotFac_Validate(Cancel As Boolean)

    If TxtTotFac.Text = "" Then TxtTotFac.Text = 0

End Sub



