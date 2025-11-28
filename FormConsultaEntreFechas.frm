VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormConsultaEntreFechas 
   Caption         =   "Consulta entre Fechas"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   13575
   Begin VB.Frame Frame1 
      Height          =   6195
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   13425
      Begin VB.TextBox TxtTOTAL 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   4680
         Width           =   1695
      End
      Begin VB.Frame Frame2 
         Height          =   1095
         Left            =   3240
         TabIndex        =   4
         Top             =   3960
         Width           =   9615
         Begin VB.CommandButton CmdExit 
            Caption         =   "&Salir"
            Height          =   735
            Left            =   8400
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            Height          =   735
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.TextBox TxtTotalRetencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         Top             =   4080
         Width           =   1695
      End
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4320
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   7320
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2775
         Left            =   960
         TabIndex        =   8
         Top             =   1080
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   4
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
         TabIndex        =   13
         Top             =   4440
         Width           =   1230
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
         Left            =   480
         TabIndex        =   12
         Top             =   840
         Width           =   1200
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
         TabIndex        =   11
         Top             =   3840
         Width           =   1440
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta"
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
         Left            =   7320
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FormConsultaEntreFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()
     Call busco
End Sub

Private Sub CmdExit_Click()
    Unload FormConsultaEntreFechas
End Sub


Private Sub Form_Load()
    
    FormConsultaEntreFechas.Height = 6195
    FormConsultaEntreFechas.Width = 13815
    
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
    
On Error GoTo Error_Handler

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    
       
    Desde = "#" & Format$(TxtFechaDesde.Text, "mm/dd/yyyy") & "#"
    Hasta = "#" & Format$(TxtFechaHasta.Text, "mm/dd/yyyy") & "#"
    
    eseqele = "SELECT * FROM PagoProvRet WHERE FechaPago >=" & Desde & " AND FechaPago <=" & Hasta & " Order By NroPago, FechaPago"
    
    Set rst = db.OpenRecordset(eseqele, dbOpenDynaset)
   
      FG1.Rows = 2
      FG1.Clear
      FG1.Visible = True
       
       Call SeteoGrilla
       
       rst.MoveFirst

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
         ' rst.FindNext
         rst.MoveNext
         

       Loop
    
    
   

Error_Handler:
    
    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If
    
    TxtTotalRetencion.Text = Format(totalrete, "#0.00")
    TxtTOTAL.Text = Format(totalpa, "#0.00")
    
    Exit Sub
    
    

End Sub

