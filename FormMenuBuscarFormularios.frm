VERSION 5.00
Begin VB.Form FormBuscarFormularios 
   Caption         =   "Buscar Ver e Imprimir"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   5475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton cmdBuscarNotaDebitoInt 
         Caption         =   "Buscar Nota de Débito &Interna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   7
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarNotaDebito 
         Caption         =   "Buscar Nota de &Débito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   6
         Top             =   2640
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   5
         Top             =   3480
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarNotaCredito 
         Caption         =   "Buscar &Nota de Crédito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   4
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarRemito 
         Caption         =   "Buscar &Remito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   3
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarPresupuesto 
         Caption         =   "Buscar &Presupuesto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscarFactura 
         Caption         =   "Buscar &Factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   1
         Top             =   720
         Width           =   1695
      End
   End
End
Attribute VB_Name = "FormBuscarFormularios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscarFactura_Click()

    FormBuscarFactura.Show

End Sub

Private Sub Command6_Click()

End Sub

Private Sub cmdBuscarFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub

Private Sub cmdBuscarNotaCredito_Click()

    FormBuscarNotaCredito.Show

End Sub

Private Sub cmdBuscarPago_Click()

    FormVerPagoFacturas.Show

End Sub

Private Sub cmdBuscarNotaCredito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub

Private Sub cmdBuscarNotaDebito_Click()

    FormBuscarNotaDebito.Show

End Sub

Private Sub cmdBuscarNotaDebitoInt_Click()

    FormBuscarNDInterna.Show

End Sub

Private Sub cmdBuscarPresupuesto_Click()

    FormBuscarPresupuesto.Show

End Sub

Private Sub cmdBuscarPresupuesto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub


Private Sub cmdBuscarRemito_Click()

    FormBuscarRemito.Show

End Sub

Private Sub cmdBuscarRemito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub


Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub cmdSalir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
        End
    End If

End Sub

Private Sub Form_Load()

    FormBuscarFormularios.Top = 1000
    FormBuscarFormularios.Left = 6000
    
End Sub


