VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Limpiar BBDD Listado de Clientes"
   ClientHeight    =   5190
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   615
         Left            =   5880
         TabIndex        =   3
         Top             =   3600
         Width           =   2175
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar BBDD Clientes"
         Height          =   615
         Left            =   3240
         TabIndex        =   1
         Top             =   3600
         Width           =   2175
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   375
         Left            =   2520
         TabIndex        =   5
         Top             =   2520
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   1440
         TabIndex        =   2
         Top             =   1560
         Width           =   1515
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   5400
      TabIndex        =   4
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BaseSPC

Public tClientes
Public tClientesAuxiliar
Public tDomicilios
Public tCtaCte
Public tFC
Public tMovCC
Public tNC
Public tPagoC
Public tPresupC
Public tRemitos

Public CodCli
Public CodCliDuplicado
Public CUIT

Public vSQL
Private Sub actCtaCte(Cliente As Long)

On Error GoTo caperr
    
    qCC = "SELECT * FROM CtaCte WHERE IdCLiente=" & Cliente & " Order By IdCliente"
    
    Set tCtaCte = BaseSPC.OpenRecordset(qCC, dbOpenDynaset)
    
    tCtaCte.MoveFirst
    
    While Not tCtaCte.EOF
    
        ' tCtaCte.Edit
            If tCtaCte!IdCliente <> CodCli Then tCtaCte.Delete
        ' tCtaCte.Update
        
        tCtaCte.MoveNext
    Wend
    
    tCtaCte.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub actDomicilios(Cliente As Long)
    
On Error GoTo caperr
    
    qDom = "SELECT * FROM DomiciliosClientes WHERE IDCliente=" & Cliente & " Order By IDCliente"
    'qdom = "SELECT * FROM DomiciliosClientes WHERE IdCliente=1 Order By IdCliente"
    
    'MsgBox (qDom)
    
    Set tDomicilios = BaseSPC.OpenRecordset(qDom, dbOpenDynaset)
    
    tDomicilios.MoveFirst
    
    While Not tDomicilios.EOF
    
        'tDomicilios.Edit
            If tDomicilios!IdCliente <> CodCli Then tDomicilios.Delete
        'tDomicilios.Update
        
        tDomicilios.MoveNext
    Wend
    
    tDomicilios.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub actMovimientosCC(Cliente As Long)

On Error GoTo caperr

    qMovCC = "SELECT * FROM MovimientosCtaCte WHERE IdCLiente=" & Cliente & " Order By IdCliente"
    
    Set tMovCC = BaseSPC.OpenRecordset(qMovCC, dbOpenDynaset)
    
    tMovCC.MoveFirst
    
    While Not tMovCC.EOF
    
        tMovCC.Edit
            If tMovCC!IdCliente <> CodCli Then tMovCC!IdCliente = CodCli
        tMovCC.Update
        
        tMovCC.MoveNext
    Wend
    
    tMovCC.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
        Case 3022
            tMovCC.Delete
            Resume Next
    End Select

End Sub


Private Sub ProcesarClientes(CUIT As String)

    On Error GoTo CapturaErrores
    
    vSQL = "SELECT * FROM Clientes WHERE CUIT ='" & CUIT & "' ORDER BY IDCliente"
'    MsgBox (vSQL)
    
    Set tClientesAuxiliar = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tClientesAuxiliar.MoveFirst
    tClientesAuxiliar.MoveNext
    
    While Not tClientesAuxiliar.EOF
            Call actDomicilios(tClientesAuxiliar!IdCliente)
            Call actMovimientosCC(tClientesAuxiliar!IdCliente)
            Call actCtaCte(tClientesAuxiliar!IdCliente)
            Call actFacturas(tClientesAuxiliar!IdCliente)
            Call actNC(tClientesAuxiliar!IdCliente)
            Call actPagos(tClientesAuxiliar!IdCliente)
            Call actPresupuestos(tClientesAuxiliar!IdCliente)
            Call actRemitos(tClientesAuxiliar!IdCliente)
            
            Call EliminoCliente(tClientesAuxiliar!IdCliente)
            
        tClientesAuxiliar.MoveNext
    Wend
    
CapturaErrores:

    Select Case Err
        Case 3021
            Resume Next
            Exit Sub
    End Select
End Sub
Private Sub actFacturas(Cliente As Long)

On Error GoTo caperr

    qFC = "SELECT * FROM FacturaC WHERE CodCliente=" & Cliente & " Order By CodCliente"
    
    Set tFC = BaseSPC.OpenRecordset(qFC, dbOpenDynaset)
    
    tFC.MoveFirst
    
    While Not tFC.EOF
    
        tFC.Edit
            If tFC!CodCliente <> CodCli Then tFC!CodCliente = CodCli
        tFC.Update
        
        tFC.MoveNext
    Wend
    
    tFC.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select


End Sub


Private Sub cmdLimpiar_Click()
        
    On Error GoTo CapturaErrores
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
                
        Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
        c = 0
        
        tClientes.MoveFirst
        
        CodCli = tClientes!IdCliente
        CUIT = tClientes!CUIT
        
        'If CUIT = "20046408218" Then MsgBox (CUIT)
        
        While Not tClientes.EOF()
           
           If Not IsNull(tClientes!CUIT) Then
               Call ProcesarClientes(tClientes!CUIT)
    
               tClientes.MoveNext
               CUIT = tClientes!CUIT
               c = c + 1
               Label3.Caption = "Registros Procesados: " & c
               Form1.Refresh
               
          '     If CUIT = "20046408218" Then MsgBox (CUIT)
               
               CodCli = tClientes!IdCliente
            Else
                tClientes.MoveNext
                Label3.Caption = c + 1
                Form1.Refresh
            End If
        Wend
        
        Call CompactBBDD
        
        MsgBox "Fin del Programa"
        
        Unload Me

CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub actNC(Cliente As Long)

On Error GoTo caperr

    qNC = "SELECT * FROM NotaCreditoC WHERE CodCliente=" & Cliente & " Order By CodCliente"
    
    Set tNC = BaseSPC.OpenRecordset(qNC, dbOpenDynaset)
    
    tNC.MoveFirst
    
    While Not tNC.EOF
    
        tNC.Edit
            If tNC!CodCliente <> CodCli Then tNC!CodCliente = CodCli
        tNC.Update
        
        tNC.MoveNext
    Wend
    
    tNC.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub
Private Sub actPagos(Cliente As Long)

On Error GoTo caperr

    qPC = "SELECT * FROM PagoC WHERE IdCliente=" & Cliente & " Order By IdCliente"
    
    Set tPagoC = BaseSPC.OpenRecordset(qPC, dbOpenDynaset)
    
    tPagoC.MoveFirst
    
    While Not tPagoC.EOF
    
        tPagoC.Edit
            If tPagoC!IdCliente <> CodCli Then tPagoC!IdCliente = CodCli
        tPagoC.Update
        
        tPagoC.MoveNext
    Wend
    
    tPagoC.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub
Private Sub actPresupuestos(Cliente As Long)

On Error GoTo caperr

qPresupC = "SELECT * FROM PresupuestoC WHERE CodCliente=" & Cliente & " Order By CodCliente"
    
    Set tPresupC = BaseSPC.OpenRecordset(qPresupC, dbOpenDynaset)
    
    tPresupC.MoveFirst
    
    While Not tPresupC.EOF
    
        tPresupC.Edit
            If tPresupC!CodCliente <> CodCli Then tPresupC!CodCliente = CodCli
        tPresupC.Update
        
        tPresupC.MoveNext
    Wend
    
    tPresupC.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub actRemitos(Cliente As Long)

On Error GoTo caperr

    qRemC = "SELECT * FROM RemitoC WHERE CodCliente=" & Cliente & " Order By CodCliente"
    
    Set tRemitos = BaseSPC.OpenRecordset(qRemC, dbOpenDynaset)
    
    tRemitos.MoveFirst
    
    While Not tRemitos.EOF
    
        tRemitos.Edit
            If tRemitos!CodCliente <> CodCli Then tRemitos!CodCliente = CodCli
        tRemitos.Update
        
        tRemitos.MoveNext
    Wend
    
    tRemitos.Close

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub
Private Sub EliminoCliente(Cliente As Long)

On Error GoTo caperr

    'qRemC = "SELECT * FROM RemitoC WHERE CodCliente=" & Cliente & " Order By CodCliente"
    
    Set tC = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    
    tC.Index = "PrimaryKey"
    tC.Seek "=", Cliente
    
    If Not tC.NoMatch Then tC.Delete
    
    tC.Close
    
    Label1.Caption = Cliente & " Eliminado por duplicado con " & CodCli

caperr:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub
Private Sub CompactBBDD()

    DatabaseName = "DB_SPC_SI.mdb"
    CommonDatabaseName = "DB_SPC_SI.mdb"
    
   ' Set db = OpenDatabase(App.Path & "\DB_SPC_SI.MDB")

    'Close DB
    BaseSPC.Close

    Screen.MousePointer = vbHourglass

'Compact DB
DBEngine.CompactDatabase App.Path & "\" & CommonDatabaseName, App.Path & "\" & CommonDatabaseName & "_"

'Delete old file and rename new one
Kill App.Path & "\" & DatabaseName
Name App.Path & "\" & DatabaseName & "_" As App.Path & "\" & DatabaseName

Screen.MousePointer = vbDefault

End Sub

