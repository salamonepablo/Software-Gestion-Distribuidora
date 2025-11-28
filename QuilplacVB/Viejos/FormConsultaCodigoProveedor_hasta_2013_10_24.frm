VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsultaCodigoProveedor 
   Caption         =   "Consulta por Codigo de Proveedor"
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
      TabIndex        =   8
      Top             =   0
      Width           =   12975
      Begin VB.ComboBox CmbCodProv 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtTotalRetencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox TxtProvName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4320
         TabIndex        =   1
         Top             =   720
         Width           =   4215
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3240
         TabIndex        =   9
         Top             =   3960
         Width           =   9615
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            Height          =   735
            Left            =   600
            Picture         =   "FormConsultaCodigoProveedor.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdExit 
            Caption         =   "&Salir"
            Height          =   735
            Left            =   8400
            Picture         =   "FormConsultaCodigoProveedor.frx":00FA
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtCUIT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   8760
         TabIndex        =   2
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox TxtTOTAL 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   4680
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2295
         Left            =   960
         TabIndex        =   3
         Top             =   1560
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         Enabled         =   0   'False
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
         TabIndex        =   15
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
         Left            =   2640
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Detalle Pagos"
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
         TabIndex        =   13
         Top             =   1320
         Width           =   1200
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
         Left            =   8640
         TabIndex        =   12
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
         Left            =   4200
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   4440
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FormConsultaCodigoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbCodProv_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call busco
     End If
End Sub

Private Sub CmdBuscar_Click()
    Call busco
End Sub

Private Sub CmdExit_Click()
    Unload FormConsultaCodigoProveedor
End Sub

Private Sub Form_Load()

    FormConsultaCodigoProveedor.Height = 6195
    FormConsultaCodigoProveedor.Width = 13425
    
      
    Set Padron = OpenDatabase("C:\QuilplacVB\Padron.mdb")
    
    Set Provs = Padron.OpenRecordset("Proveedores")
    
            
     With Provs
        .MoveFirst
        While Not .EOF
           CmbCodProv.AddItem (!CodProv)
           .MoveNext
        Wend
    End With
    
    Call SeteoGrilla
    
End Sub
Private Sub CleanDatos2()

  
  '  CmbCodProv.Text = ""
    TxtProvName.Text = ""
    TxtCUIT.Text = ""
    TxtTotalRetencion.Text = ""
    TxtTOTAL.Text = ""
    FG1.Clear
    
    Call SeteoGrilla
  

End Sub

Sub SeteoGrilla()
    
    'FG1.AutoSizeMode = klexAutoSizeColWidth
    FG1.Row = 0
    FG1.Col = 0
    
    
    FG1.Col = 0
    FG1.ColWidth(0) = 700
    FG1.Text = "Nº Pago"
    FG1.ColAlignment(0) = flexAlignCenterCenter
    
    FG1.Col = 1
    FG1.ColWidth(1) = 1500
    FG1.Text = "Fecha Pago"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1500
    FG1.Text = "Total Retenido"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1500
    FG1.Text = "Total Pagado"
    FG1.ColAlignment(3) = flexAlignCenterCenter
         
       
End Sub


Private Sub busco()

    
'***************Busco en PagoProvret


    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rst = db.OpenRecordset("PagoProvRet", dbOpenDynaset)
    
  '  CodiProv = Val(CmbCodProv.Text)
      
  '  rst.FindFirst "CodProv= " + Str(CodiProv)
  '  If rst.Fields!CodProv <> Val(CmbCodProv.Text) Then
    
    
    CodiProv = Val(CmbCodProv.Text)
      
    rst.FindFirst "CodProv= " + Str(CodiProv)
    If rst.Fields!CodProv <> Val(CmbCodProv.Text) Then
       mensaje = MsgBox("No Existen Pagos", vbCritical, "Final de la busqueda")
       
        CmbCodProv.SetFocus
        Call CleanDatos2
    Else
      TxtProvName.Text = rst.Fields!NombreProv
      TxtCUIT.Text = rst.Fields!Cuit
      FG1.Rows = 2
      FG1.Clear
      FG1.Visible = True
       
       Call SeteoGrilla
       
       linea2 = 1
       Do While Not rst.NoMatch
          FG1.AddItem " "
          FG1.Row = linea2
          FG1.Col = 0
          FG1.Text = rst.Fields!NroPago
          FG1.Col = 1
          FG1.Text = rst.Fields!FechaPago
          FG1.Col = 2
          FG1.Text = Format(rst.Fields!TotalReten, "#0.00")
          totalrete = totalrete + rst.Fields!TotalReten
          FG1.Col = 3
          FG1.Text = Format(rst.Fields!TotalPago, "#0.00")
          totalpa = totalpa + rst.Fields!TotalPago
          linea2 = linea2 + 1
          rst.FindNext "CodProv= " + Str(CodiProv)
          
       Loop
    End If
    
    TxtTotalRetencion.Text = Format(totalrete, "#0.00")
    TxtTOTAL.Text = Format(totalpa, "#0.00")
    
End Sub



