VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Actualizar Campo ""Corresponde"" de la Tabla PagosC"
   ClientHeight    =   3060
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdProceso 
      Caption         =   "Comenzar Proceso"
      Height          =   735
      Left            =   1440
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   195
      Left            =   3120
      TabIndex        =   2
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   720
      TabIndex        =   1
      Top             =   1440
      Width           =   480
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BaseSPC
Dim tPagos
Dim tMovimientosCC
Private Sub cmdProceso_Click()

    Dim Cont As Integer
    
    'Actualizar el campo Corresponde a L1
        Cont = 0
        
        vSQL = "SELECT * FROM MovimientosCtaCte WHERE TipoDoc Like '*Linea 1' Order By  NroDoc"
        
        Set tMovimientosCC = BaseSPC.OpenRecordset(vSQL)
        
        With tMovimientosCC
            .MoveFirst
        
            While Not .EOF
                Set tPagos = BaseSPC.OpenRecordset("PagoC", dbOpenTable)
                
                tPagos.Index = "PrimaryKey"
                             
                tPagos.Seek "=", tMovimientosCC!NroDoc
                
                If Not tPagos.NoMatch Then
                    tPagos.Edit
                        tPagos!Corresponde = "L1"
                        Cont = Cont + 1
                        Label2.Caption = (Cont) & " Línea 1"
                        Form1.Refresh
                    tPagos.Update
                End If
                .MoveNext
            Wend
        End With
        
    'Actualizar el campo Corresponde a L2
        Cont = 0
        Label3.Visible = True
        
        vSQL = "SELECT * FROM MovimientosCtaCte WHERE TipoDoc Like '*Linea 2' Order By  NroDoc"
        
        Set tMovimientosCC = BaseSPC.OpenRecordset(vSQL)
        
        With tMovimientosCC
            .MoveFirst
        
            While Not .EOF
                Set tPagos = BaseSPC.OpenRecordset("PagoC", dbOpenTable)
                
                tPagos.Index = "PrimaryKey"
                             
                tPagos.Seek "=", tMovimientosCC!NroDoc
                
                If Not tPagos.NoMatch Then
                    tPagos.Edit
                        tPagos!Corresponde = "L2"
                        Cont = Cont + 1
                        Label3.Caption = Cont & " Línea 2"
                        Form1.Refresh
                    tPagos.Update
                End If
                .MoveNext
            Wend
        End With

    MsgBox ("PROCESO FINALIZADO CON EXITO")
    
    End

End Sub

Private Sub Form_Load()

    
    Label1.FontName = "Arial"
    Label1.FontSize = 10
    Label1.FontBold = True
    Label1.Caption = "Registros Actualizados: "
    
    Label2.FontName = "Arial"
    Label2.FontSize = 10
    Label2.FontBold = True
    Label2.ForeColor = vbRed
    Label2.Caption = "00"
    
    Label3.FontName = "Arial"
    Label3.FontSize = 10
    Label3.FontBold = True
    Label3.ForeColor = vbRed
    Label3.Caption = "00"
    Label3.Visible = False
    
'Seteo la base
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")



    
    
End Sub
