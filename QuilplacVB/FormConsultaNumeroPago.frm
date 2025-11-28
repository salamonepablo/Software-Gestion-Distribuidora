VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormConsultaNumeroPago 
   Caption         =   "Consulta por Numero de Pago"
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
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   12975
      Begin VB.TextBox TxtCodProv 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4080
         TabIndex        =   19
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtPayNr 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   720
         Width           =   855
      End
      Begin VB.TextBox TxtTOTAL 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   8
         Top             =   4680
         Width           =   1695
      End
      Begin VB.TextBox TxtCUIT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   10080
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3240
         TabIndex        =   5
         Top             =   3960
         Width           =   9615
         Begin VB.CommandButton CmdPrintQry 
            Caption         =   "&Imprimir"
            Height          =   735
            Left            =   1800
            Picture         =   "FormConsultaNumeroPago.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "&Salir"
            Height          =   735
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            Height          =   735
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtPayDate 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtProvName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   3
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox TxtTotalRetencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   4080
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2295
         Left            =   960
         TabIndex        =   9
         Top             =   1560
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   10
         FixedCols       =   0
         Enabled         =   0   'False
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
         TabIndex        =   17
         Top             =   4440
         Width           =   1230
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
         TabIndex        =   16
         Top             =   480
         Width           =   1995
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
         TabIndex        =   15
         Top             =   480
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalle de Facturas Pagas"
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
         TabIndex        =   14
         Top             =   1320
         Width           =   2265
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
         TabIndex        =   13
         Top             =   480
         Width           =   930
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
         TabIndex        =   12
         Top             =   480
         Width           =   1095
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
         TabIndex        =   11
         Top             =   480
         Width           =   1335
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
         TabIndex        =   10
         Top             =   3840
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FormConsultaNumeroPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()
    Call busco
End Sub

Private Sub CmdExit_Click()
    Unload FormConsultaNumeroPago
End Sub

Private Sub CmdPrint_Click()

End Sub

Private Sub CmdPrintQry_Click()

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
        Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
        CodiProv = Val(TxtCodProv.Text)
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
                Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
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
Private Sub Form_Load()

    FormConsultaNumeroPago.Height = 6195
    FormConsultaNumeroPago.Width = 13425
    
    Call SeteoGrilla
    
End Sub
Private Sub CleanDatos2()

    TxtPayNr.Text = ""
    TxtPayDate.Text = ""
  '  CmbCodProv.Text = ""
    TxtProvName.Text = ""
    TxtCUIT.Text = ""
    TxtTotalRetencion.Text = ""
    TxtTOTAL.Text = ""
    FG1.Clear
    
    Call SeteoGrilla
    TxtPayNr.SetFocus

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
    
End Sub

Private Sub TxtPayNr_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Call busco
     End If
End Sub

Private Sub busco()

'***************Busco en PagoProvret


    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rst = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
    NumeroPago = Val(TxtPayNr.Text)
      
    rst.FindFirst "NroPago= " + Str(NumeroPago)
    If rst.Fields!NroPago <> Val(TxtPayNr.Text) Then
       mensaje = MsgBox("Numero de Pago Inexistente", vbCritical, "Final de la busqueda")
       TxtPayNr.Text = ""
       TxtPayNr.SetFocus
       Call CleanDatos2
    Else
       TxtPayNr.Text = rst.Fields!NroPago
       TxtPayDate.Text = rst.Fields!FechaPago
       TxtCodProv.Text = rst.Fields!CodProv
       TxtProvName.Text = rst.Fields!NombreProv
       TxtCUIT.Text = rst.Fields!Cuit
       TxtTotalRetencion.Text = Format(rst.Fields!TotalReten, "#0.00")
       TxtTOTAL.Text = Format(rst.Fields!TotalPago, "#0.00")
    
    
'***************Busco en LineasPagoRet

    Set db2 = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rst2 = db2.OpenRecordset("LineasPagoret", dbOpenDynaset)
    
   FG1.Rows = 2
   FG1.Clear
   FG1.Visible = True
    
    
    Call SeteoGrilla
    
    NumeroPago = Val(TxtPayNr.Text)
      
    rst2.FindFirst "NroPago= " + Str(NumeroPago)
    If rst2.Fields!NroPago <> Val(TxtPayNr.Text) Then
       mensaje = MsgBox("Numero de Pago Inexistente", vbCritical, "Final de la busqueda")
       TxtPayNr.Text = ""
       TxtPayNr.SetFocus
    End If
    linea2 = 1
    Do While Not rst2.NoMatch
       FG1.AddItem " "
       FG1.Row = linea2
       FG1.Col = 0
       FG1.Text = rst2.Fields!Item
       FG1.Col = 1
       FG1.Text = rst2.Fields!NroPago
       FG1.Col = 2
       FG1.Text = rst2.Fields!NroFactProv
       FG1.Col = 3
       FG1.Text = rst2.Fields!FechaFactProv
       FG1.Col = 4
       FG1.Text = Format(rst2.Fields!TotalFac, "#0.00")
       FG1.Col = 5
       FG1.Text = Format(rst2.Fields!BaseImponible, "#0.00")
       FG1.Col = 6
       FG1.Text = rst2.Fields!Alicuota
       FG1.Col = 7
       FG1.Text = Format(rst2.Fields!ImporteRet, "#0.00")
       FG1.Col = 8
       FG1.Text = Format(rst2.Fields!ImporteIva, "#0.00")
       FG1.Col = 9
       FG1.Text = Format(rst2.Fields!TotalLineaFactura, "#0.00")
       linea2 = linea2 + 1
       
       rst2.FindNext "NroPago= " + Str(NumeroPago)
       
    Loop
 End If

End Sub
