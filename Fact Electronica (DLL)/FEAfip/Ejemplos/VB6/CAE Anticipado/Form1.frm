VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Consultar"
      Height          =   735
      Left            =   4080
      TabIndex        =   4
      Top             =   720
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Informar"
      Enabled         =   0   'False
      Height          =   1215
      Left            =   2040
      TabIndex        =   3
      Top             =   4440
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "CAE"
      Height          =   1455
      Left            =   1200
      TabIndex        =   1
      Top             =   2040
      Width           =   4935
      Begin VB.Label lbCAE 
         Height          =   375
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Solicitar CAE Anticipado"
      Height          =   735
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
        Dim wsfev1 As FEAFIPLib.wsfev1 ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto
Private Sub Command1_Click()
        ' Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
        ' y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
        ' la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.
        
        Dim nro As Double
        CAE$ = ""
        Periodo = 201403
        Orden = 2
        FechavigDesde$ = ""
        FechaVigHasta$ = ""
        FechaTope$ = ""
        FechaProceso$ = ""
         
        Set wsfev1 = New FEAFIPLib.wsfev1
        wsfev1.CUIT = 20939802593#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
                                
                wsfev1.Reset
                If Not wsfev1.CAEASolicitar(Periodo, Orden, CAE$, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
                    MsgBox wsfev1.ErrorDesc
                Else
                    Command2.Enabled = True
                    lbCAE.Caption = CAE
                    Command2.Enabled = True
                End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If

End Sub

Private Sub Command2_Click()
        Dim nro As Double
        CAE$ = ""
        Vencimiento$ = ""
        Resultado$ = ""
        Reproceso$ = ""
        nro = 0
        PtoVta = 10
        FechaComp = Format(Now(), "yyyymmdd")
        TipoComp = 1 ' Factura A(Ver excel referencias codigos AFIP documentacion.rar o ir a http://www.bitingenieria.com.ar/codigos.html)
        
        Set wsfev1 = New FEAFIPLib.wsfev1
                
        wsfev1.CUIT = 20939802593#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
                MsgBox (wsfev1.ErrorDesc)
            Else
                nro = nro + 1
                wsfev1.AgregaFactura 1, 80, 30707219072#, nro, nro, FechaComp, 121, 0, 100, 0, "", "", "", "PES", 1
                wsfev1.AgregaIVA 5, 100, 21 ' Ver Excel de referencias de codigos AFIP
                If wsfev1.CAEAInformar(PtoVta, TipoComp, lbCAE.Caption) Then
                    wsfev1.AutorizarRespuesta 0, CAE, Vencimiento, Resultado, Reproceso
                    If Resultado = "A" Then
                        MsgBox "Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento: " + CAE + " " + Vencimiento
                    Else
                        MsgBox wsfev1.AutorizarRespuestaObs(0)
                    End If
                Else
                    MsgBox wsfev1.ErrorDesc
                End If
            End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If

End Sub

Private Sub Command3_Click()
        Set wsfev1 = New FEAFIPLib.wsfev1
        
        wsfev1.CUIT = 20939802593#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
                Periodo = 201403
                Orden = 2
                FechavigDesde$ = ""
                FechaVigHasta$ = ""
                FechaTope$ = ""
                FechaProceso$ = ""
                                
                wsfev1.Reset
                If Not wsfev1.CAEAConsultar(Periodo, Orden, CAE$, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
                    MsgBox wsfev1.ErrorDesc
                Else
                    Command2.Enabled = True
                    lbCAE.Caption = CAE
                    Command2.Enabled = True
                End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If
End Sub
