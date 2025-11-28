VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5640
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390
   LinkTopic       =   "Form1"
   ScaleHeight     =   5640
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3255
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   5741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   600
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    Dim tPagosC, tPagosD, tSucursales

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
        
    Set tSucursales = db.OpenRecordset("Sucursales", dbOpenTable)
    Set tPagosC = db.OpenRecordset("PagosC", dbOpenTable)
    
    
    tSucursales.Index = "PrimaryKey"
    tPagosC.Index = "PrimaryKey"
    'tPagosD.Index = "PrimaryKey"
    
    
    tPagosC.MoveFirst
    
    While Not tPagosC.EOF
        Select Case tPagos!IdSucursal
            Case 5
                vSQL = "SELECT * FROM tPagosD WHERE IdSucursal=" & tPagosC!IdSucursal & " AND tPagosD!NroPago=" & tPagosC!NroPago & " ORDER BY LineaPago"
                Set tPagosD = db.OpenRecordset(vSQL, dbOpenDynaset)
                 tPagosD.MoveFirst
                 While Not tPagosD.EOF
                    tPagosD.Edit
                        tPagosD!IdSucursal = 2
                    tPagosD.Update
                 Wend
                
        End Select
    
    Wend
    
    
    Do
        cmbSucursal.AddItem tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal
        tSucursales.MoveNext
    
    Loop Until tSucursales.EOF
    
    cmbSucursal.ListIndex = 0
    
    tSucursales.Close
    db.Close

End Sub


