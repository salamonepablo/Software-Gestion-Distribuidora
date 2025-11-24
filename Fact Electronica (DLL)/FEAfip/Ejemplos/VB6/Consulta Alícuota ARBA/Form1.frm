VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   4785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10410
   LinkTopic       =   "Form1"
   ScaleHeight     =   4785
   ScaleWidth      =   10410
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1080
      TabIndex        =   2
      Text            =   "20123456789"
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar Alicuota"
      Height          =   375
      Left            =   3600
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label lbGrupoRetencion 
      Height          =   255
      Left            =   2280
      TabIndex        =   10
      Top             =   3360
      Width           =   7575
   End
   Begin VB.Label Label7 
      Caption         =   "Grupo Retencion:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label lbGrupoPercepcion 
      Height          =   255
      Left            =   2280
      TabIndex        =   8
      Top             =   2640
      Width           =   7575
   End
   Begin VB.Label Label3 
      Caption         =   "Grupo Percepcion:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   1335
   End
   Begin VB.Label lbRetencion 
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   7575
   End
   Begin VB.Label Label4 
      Caption         =   "Alicuota Retencion:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label lbPercepcion 
      Height          =   255
      Left            =   2280
      TabIndex        =   4
      Top             =   1320
      Width           =   7575
   End
   Begin VB.Label Label2 
      Caption         =   "Alicuota Percepcion:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "CUIT:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim lwsPadron As wsPadronARBA
    Set lwsPadron = New wsPadronARBA
    lwsPadron.User = "20939802593"
    lwsPadron.Password = "123456"
    lwsPadron.ModoProduccion = False ' Debe dar de alta el cuit en el entorno de test de ARBA http://www.test.arba.gov.ar/
    If lwsPadron.ConsultaAlicuota("20160701", "20160731", CDbl(Replace(Text1.Text, "-", ""))) Then
        lbPercepcion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaPercepcion)
        lbRetencion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaRetencion)
        lbGrupoPercepcion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.GrupoPercepcion)
        lbGrupoRetencion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.GrupoRetencion)
    Else
        MsgBox (lwsPadron.ErrorDesc)
    End If

End Sub

