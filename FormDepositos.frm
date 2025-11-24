VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5280
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14430
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   14430
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14055
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   13575
         Begin VB.CommandButton btnPrimero 
            Caption         =   "|<"
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAtras 
            Caption         =   "<<"
            Height          =   615
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAdelante 
            Caption         =   ">>"
            Height          =   615
            Left            =   2640
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnUltimo 
            Caption         =   ">|"
            Height          =   615
            Left            =   3840
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnGrabar 
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   6240
            TabIndex        =   12
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnModificar 
            Caption         =   "&Modificar"
            Height          =   615
            Left            =   7440
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   8640
            TabIndex        =   10
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnEliminar 
            Caption         =   "&Eliminar"
            Height          =   615
            Left            =   9840
            TabIndex        =   9
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnCtaCte 
            Caption         =   "&Cta Cte"
            Height          =   615
            Left            =   11040
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   12240
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "&Limpiar"
            Height          =   615
            Left            =   5040
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1080
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Depósito:"
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
         TabIndex        =   2
         Top             =   480
         Width           =   810
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnBuscar_Click()

    If vFlagBuscar = 0 Then
        vFlagBuscar = 1
        txtIDCliente.Enabled = True
        txtIDCliente.Text = ""
        txtIDCliente.SetFocus
     Else
        
        If txtIDCliente.Text <> "" Then
            Campo = "IDCliente= "
            Valor = txtIDCliente.Text
         Else
            If txtRazonSocial.Text <> "" Then
                Campo = "RazonSocial Like "
                Valor = "'" + txtRazonSocial.Text + "*'"
             Else
                If txtNombreFantasia.Text <> "" Then
                    Campo = "NombreFantasia Like "
                    Valor = "'" + txtNombreFantasia.Text + "*'"
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Clientes WHERE " & Campo & Valor & " ORDER BY IDCliente"
        
        'MsgBox (vsql)
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL)
    
        If Not tClientes.NoMatch Then
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
        Else
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
        End If
        
        vFlagBuscar = 0
        
    End If


End Sub

Private Sub btnLimpiar_Click()

End Sub
