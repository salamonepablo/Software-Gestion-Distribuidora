VERSION 5.00
Begin VB.Form FormRecibo 
   Caption         =   "Datos Complemetarios del Recibo"
   ClientHeight    =   5250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10755
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   10755
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5055
      Left            =   120
      TabIndex        =   20
      Top             =   0
      Width           =   10455
      Begin VB.CommandButton cmdAceptarRecibo 
         Caption         =   "&Aceptar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         TabIndex        =   19
         Top             =   4200
         Width           =   9615
      End
      Begin VB.Frame Frame3 
         Caption         =   "En Concepto de..."
         Height          =   3735
         Left            =   5760
         TabIndex        =   22
         Top             =   360
         Width           =   4335
         Begin VB.TextBox txtConcepto 
            Height          =   2775
            Left            =   360
            MaxLength       =   92
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   18
            Top             =   600
            Width           =   3615
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Facturas"
         Height          =   3735
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   5175
         Begin VB.TextBox txtImpF6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   17
            Top             =   3120
            Width           =   1335
         End
         Begin VB.TextBox txtFF6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   15
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox txtF6 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   16
            Top             =   3120
            Width           =   1455
         End
         Begin VB.TextBox txtImpF5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   14
            Top             =   2640
            Width           =   1335
         End
         Begin VB.TextBox txtFF5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   12
            Top             =   2640
            Width           =   975
         End
         Begin VB.TextBox txtF5 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   13
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox txtImpF4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   11
            Top             =   2160
            Width           =   1335
         End
         Begin VB.TextBox txtFF4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   9
            Top             =   2160
            Width           =   975
         End
         Begin VB.TextBox txtF4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   10
            Top             =   2160
            Width           =   1455
         End
         Begin VB.TextBox txtImpF3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   8
            Top             =   1680
            Width           =   1335
         End
         Begin VB.TextBox txtFF3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   6
            Top             =   1680
            Width           =   975
         End
         Begin VB.TextBox txtF3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   7
            Top             =   1680
            Width           =   1455
         End
         Begin VB.TextBox txtImpF2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   5
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtFF2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   3
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtF2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   4
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtImpF1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   3360
            TabIndex        =   2
            Top             =   720
            Width           =   1335
         End
         Begin VB.TextBox txtF1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   1560
            TabIndex        =   1
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtFF1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   360
            TabIndex        =   0
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe"
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
            Left            =   3720
            TabIndex        =   25
            Top             =   360
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Factura Nº"
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
            Left            =   1800
            TabIndex        =   24
            Top             =   360
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   23
            Top             =   360
            Width           =   540
         End
      End
   End
End
Attribute VB_Name = "FormRecibo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public vFF1
Public vFF2
Public vFF3
Public vFF4
Public vFF5
Public vFF6
Public vNFac1
Public vNFac2
Public vNFac3
Public vNFac4
Public vNFac5
Public vNFac6
Public vImpF1
Public vImpF2
Public vImpF3
Public vImpF4
Public vImpF5
Public vImpF6
Public vConcepto

Private Sub cmdAceptarRecibo_Click()

    FormPagoFacturas.Show
    
End Sub

Private Sub Form_Load()
 
 vFF1 = ""
 vFF2 = ""
 vFF3 = ""
 vFF4 = ""
 vFF5 = ""
 vFF6 = ""
 vNFac1 = ""
 vNFac2 = ""
 vNFac3 = ""
 vNFac4 = ""
 vNFac5 = ""
 vNFac6 = ""
 vImpF1 = ""
 vImpF2 = ""
 vImpF3 = ""
 vImpF4 = ""
 vImpF5 = ""
 vImpF6 = ""
 vConcepto = ""
 
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtF1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtF2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtF3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtF4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtF5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub txtF6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFF6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF3_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF4_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF5_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtImpF6_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


