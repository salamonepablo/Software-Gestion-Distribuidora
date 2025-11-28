VERSION 5.00
Begin VB.Form FormPrimerRegCC 
   Caption         =   "GENERA PRIMER REGISTRO CC"
   ClientHeight    =   4200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6120
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   6120
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Iniciar Proceso"
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   1080
      Width           =   2775
   End
End
Attribute VB_Name = "FormPrimerRegCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

 '** Genera Primer Registro de Cuentas Corrientes al Inicio de Instalación del Programa"
    
    Set MiBase = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    Set tCtaCte = MiBase.OpenRecordset("CtaCte", dbOpenTable)
    
    
    vSQL = "SELECT * FROM MovimientosCtaCte Order By IdCliente"
    Set tMovCC = MiBase.OpenRecordset(vSQL, dbOpenDynaset)

    tCtaCte.Index = "PrimaryKey"
    

    tMovCC.MoveFirst
    
    While Not tMovCC.EOF
        
        With tCtaCte
            
            .Seek "=", tMovCC!IDCliente
            
            If .NoMatch Then
                .AddNew
                    !IDCliente = tMovCC!IDCliente
                    !SaldoL1 = tMovCC!ImporteLinea1
                    !SaldoL2 = tMovCC!ImporteLinea2
                    !SaldoTotal = tMovCC!ImporteLinea1 + tMovCC!ImporteLinea2
                    !FechaActSaldo = Format(Date, "DD/MM/YYYY")
                .Update
            Else
                .Edit
                    !IDCliente = tMovCC!IDCliente
                    !SaldoL1 = !SaldoL1 + tMovCC!ImporteLinea1
                    !SaldoL2 = tMovCC!ImporteLinea2
                    !SaldoTotal = !SaldoL2 + tMovCC!ImporteLinea1 + tMovCC!ImporteLinea2
                    !FechaActSaldo = Format(Date, "DD/MM/YYYY")
                .Update
            End If
        
        End With
        tMovCC.MoveNext
    Wend

    tMovCC.Close
    tCtaCte.Close
    
    MiBase.Close
    
End Sub

