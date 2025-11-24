VERSION 5.00
Begin VB.Form FormLogin 
   Caption         =   "Ingreso al Sistema "
   ClientHeight    =   4395
   ClientLeft      =   1005
   ClientTop       =   3000
   ClientWidth     =   5550
   Icon            =   "FormLogin.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4395
   ScaleWidth      =   5550
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   2280
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   2280
      TabIndex        =   4
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton CmdIngresar 
      Caption         =   "Ingresar"
      Height          =   1095
      Left            =   480
      MouseIcon       =   "FormLogin.frx":08CA
      Picture         =   "FormLogin.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   1095
      Left            =   3000
      Picture         =   "FormLogin.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label Label0 
      Caption         =   "Bienvenido !!!!!!"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label LabelClave 
      Caption         =   "Clave de Usuario"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   1575
   End
   Begin VB.Label LabelUsuario 
      AutoSize        =   -1  'True
      Caption         =   "Nombre de Usuario"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   1845
   End
End
Attribute VB_Name = "FormLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdIngresar_Click()
CmdIngresar.Picture = LoadPicture("C:\temp\ico\ok.ico")

If Text1.Text = "" Then MsgBox "Ingrese un nombre de usuario", vbInformation, "nombre vacio": Text1.SetFocus: Exit Sub
If Text2.Text = "" Then MsgBox "Ingrese una Clave de usuario", vbInformation, "clave vacio": Text2.SetFocus: Exit Sub

With RsUsuarios
    .Requery
    .Find "IDUsuario='" & Trim(Text1.Text) & "'"
    If .EOF Then
      MsgBox "Usuario Inexistente", vbInformation, "No se encuentra Usuario"
    Else
      If !Clave = Trim(Text2.Text) Then
      'FormPanelControl.Show
      MenuPrincipal.Show
    Else
      MsgBox "Clave Incorrecta", vbCritical, "Clave Incorrecta"
    Exit Sub
    End If
    End If
End With
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Usuarios
End Sub

