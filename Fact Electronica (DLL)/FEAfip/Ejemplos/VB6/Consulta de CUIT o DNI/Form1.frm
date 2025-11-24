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
      Left            =   1920
      TabIndex        =   2
      Text            =   "20939802593"
      Top             =   600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar CUIT/DNI"
      Height          =   495
      Left            =   4560
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   4080
      Width           =   7815
   End
   Begin VB.Label Label5 
      Caption         =   "Observaciones:"
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Label lbDomicilio 
      Height          =   255
      Left            =   2040
      TabIndex        =   10
      Top             =   3360
      Width           =   7815
   End
   Begin VB.Label Label7 
      Caption         =   "Domicilio Fiscal:"
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label lbEstado 
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   2640
      Width           =   7815
   End
   Begin VB.Label Label3 
      Caption         =   "Estado:"
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   2640
      Width           =   615
   End
   Begin VB.Label lbTipo 
      Height          =   255
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   7815
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo de Persona:"
      Height          =   255
      Left            =   480
      TabIndex        =   5
      Top             =   1920
      Width           =   1335
   End
   Begin VB.Label lbNombre 
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   1320
      Width           =   7815
   End
   Begin VB.Label Label2 
      Caption         =   "Nombre:"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "CUIT/DNI:"
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Dim lwsPadron As wsPadron
    Dim Contribuyente As Contribuyente
    Dim lDomicilio As Domicilio
    
    Set lwsPadron = New wsPadron

    Set Contribuyente = New Contribuyente
    
    lwsPadron.CUIT = 20939802593#
    lwsPadron.ModoProduccion = False ' Para poder consultar un CUIT real debe habilitar el modo producción.
    If lwsPadron.login("certificado.crt", "clave.key") Then
        If lwsPadron.consultar(Text1.Text, Contribuyente) Then
            lbNombre.Caption = Contribuyente.nombre
            lbTipo.Caption = Contribuyente.tipoPersona
            lbEstado.Caption = Contribuyente.estadoClave
            Set lDomicilio = Contribuyente.domicilioFiscal
            If lDomicilio.direccion <> "" Then
                lbDomicilio.Caption = lDomicilio.direccion + ", " + lDomicilio.localidad + ", " + lDomicilio.provincia
            End If
        ' Control de constancia de inscripción. Si hay observaciones el contribuyente tiene errores de constancia.
            Label6.Caption = Contribuyente.observaciones
        Else
            MsgBox (lwsPadron.ErrorDesc)
        End If
    Else
        MsgBox (lwsPadron.ErrorDesc)
    End If

End Sub
