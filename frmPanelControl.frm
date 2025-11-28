VERSION 5.00
Begin VB.Form FormPanelControl 
   Caption         =   "PANEL DE CONTROL"
   ClientHeight    =   8670
   ClientLeft      =   6945
   ClientTop       =   2655
   ClientWidth     =   7425
   LinkTopic       =   "Form1"
   ScaleHeight     =   8670
   ScaleWidth      =   7425
   Begin VB.Frame Frame1 
      Height          =   8295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdDepositos 
         Caption         =   "Depositos"
         Height          =   615
         Left            =   360
         TabIndex        =   8
         Top             =   3480
         Width           =   2055
      End
      Begin VB.PictureBox Picture1 
         Height          =   1815
         Left            =   120
         Picture         =   "frmPanelControl.frx":0000
         ScaleHeight     =   1755
         ScaleWidth      =   5355
         TabIndex        =   7
         Top             =   360
         Width           =   5415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Prueba"
         Height          =   1335
         Left            =   0
         TabIndex        =   6
         Top             =   4800
         Width           =   3855
      End
      Begin VB.CommandButton cmdEmpleados 
         Caption         =   "Empleados"
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   2400
         Width           =   2895
      End
      Begin VB.CommandButton btnProductos 
         Caption         =   "&Productos"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3360
         TabIndex        =   3
         Top             =   4155
         Width           =   3855
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   6360
         Width           =   3855
      End
      Begin VB.CommandButton btnFacturas 
         Caption         =   "&Facturas"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3360
         TabIndex        =   2
         Top             =   3360
         Width           =   3855
      End
      Begin VB.CommandButton btnClientes 
         Caption         =   "&Clientes"
         BeginProperty Font 
            Name            =   "Berlin Sans FB"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3360
         TabIndex        =   1
         Top             =   2400
         Width           =   3855
      End
   End
End
Attribute VB_Name = "FormPanelControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnClientes_Click()

    FormClientes.Show

End Sub

Private Sub btnFacturas_Click()

    FormFactura.Show

End Sub

Private Sub btnProductos_Click()

    FormProductos.Show

End Sub

Private Sub btnSalir_Click()

    Unload Me
    

End Sub

Private Sub cmdDepositos_Click()
Depositos.Show

End Sub

Private Sub cmdEmpleados_Click()
Empleados.Show

End Sub

Private Sub Command1_Click()
Prueba.Show


End Sub

Private Sub Form_Load()

    FormPanelControl.Width = 7590
    FormPanelControl.Height = 8970
    FormPanelControl.Left = 5000
    FormPanelControl.Top = 1500
    
End Sub

