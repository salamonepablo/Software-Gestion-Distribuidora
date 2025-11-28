VERSION 5.00
Begin VB.Form FormBusquedaFactura 
   Caption         =   "Busqueda de Factura"
   ClientHeight    =   2565
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5535
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonBuscar 
         Caption         =   "&Buscar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   480
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox TextTipoFactura 
         Alignment       =   2  'Center
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
         Left            =   4080
         TabIndex        =   3
         Top             =   480
         Width           =   375
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
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo"
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
         TabIndex        =   5
         Top             =   480
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         Left            =   600
         TabIndex        =   4
         Top             =   480
         Width           =   915
      End
   End
End
Attribute VB_Name = "FormBusquedaFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BotonBuscar_Click()
    
    FormModificacionFactura.Show
    
End Sub

Private Sub BotonSalir_Click()

    Unload FormBusquedaFactura

End Sub

Private Sub Form_Load()

    FormBusquedaFactura.Height = 3135
    FormBusquedaFactura.Width = 6030
    FormBusquedaFactura.Top = 1000
    FormBusquedaFactura.Left = 1000

End Sub

Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        TextTipoFactura.SetFocus
    End If
    
End Sub

Private Sub TextTipoFactura_Change()

    If TextTipoFactura.Text <> "" Then
        BotonBuscar.Enabled = True
    End If

End Sub

Private Sub TextTipoFactura_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then
        BotonBuscar.Enabled = True
        BotonBuscar.SetFocus
    End If

End Sub
