VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   5205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5205
   ScaleWidth      =   6315
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Consultar Comprobante"
      Height          =   735
      Left            =   3840
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos"
      Height          =   3495
      Left            =   360
      TabIndex        =   1
      Top             =   1440
      Width           =   5535
      Begin VB.Label lbNumero 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   20
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label10 
         Caption         =   "Número"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lbIdentificador 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   17
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label lbTributos 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   16
         Top             =   2760
         Width           =   2895
      End
      Begin VB.Label lbIVA 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   15
         Top             =   2400
         Width           =   2895
      End
      Begin VB.Label lbNeto 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   14
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label lbTotal 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   1680
         Width           =   2895
      End
      Begin VB.Label lbFecha 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   12
         Top             =   1320
         Width           =   2895
      End
      Begin VB.Label lbVencimiento 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   11
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label lbCAE 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   600
         Width           =   2895
      End
      Begin VB.Label Label7 
         Caption         =   "Tributos"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "Iva"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Neto"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Total"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Vencimiento"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Identificador"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "CAE"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Label Label9 
      Caption         =   $"Form1.frx":0000
      Height          =   855
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        ' Los nombres de los parametros de las funciones se obtienen en FEAFIP.pdf
        
        'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
        Dim wsfev1 As FEAFIPLib.wsfev1 ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto
        Dim nro As Double
        Dim CAE As String
        Dim Vencimiento As String
        PtoVta = 30
        TipoComp = 1 ' Factura A(Ver excel referencias codigos AFIP www.bitingenieria.com.ar/#!/soporte)
        FechaComp = Format(Now(), "yyyymmdd")
         
        Set wsfev1 = New FEAFIPLib.wsfev1
        wsfev1.CUIT = 20939802593#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            ' Se obtiene el ultimo numero de comprobante a modo de ejemplo.
            If wsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
                ' Las variables CAE y Vencimiento recuperan estos valores.
                'Para obtener el resto de los valores informados a la AFIP se explora wsfev1.CmpConsultarCbte
                If wsfev1.CmpConsultar(TipoComp, PtoVta, nro, CAE, Vencimiento) Then
                   Dim cbte As FEAFIPLib.Comprobante
                   Set cbte = wsfev1.CmpConsultarCbte
                   lbNumero.Caption = cbte.Cbtedesde
                   lbCAE.Caption = cbte.CodAutorizacion
                   lbFecha.Caption = cbte.CbteFch
                   lbIdentificador.Caption = cbte.DocNro
                   lbIVA.Caption = cbte.ImpIVA
                   lbNeto.Caption = cbte.ImpNeto
                   lbTotal.Caption = cbte.Imptotal
                   lbTributos.Caption = cbte.ImpTrib
                   lbVencimiento.Caption = cbte.FchVto
                Else
                  MsgBox (wsfev1.ErrorDesc)
                End If
            Else
                MsgBox (wsfev1.ErrorDesc)
            End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If

End Sub

