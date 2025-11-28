VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormVerProductos 
   Caption         =   "Listado Productos"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11460
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11460
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   6840
      Width           =   10695
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   4680
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   8160
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonEliminar 
         Caption         =   "&Eliminar"
         Height          =   615
         Left            =   6360
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Height          =   630
         Index           =   0
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonMostrarProductos 
         Caption         =   "&Mostrar Productos"
         Height          =   630
         Index           =   1
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   1230
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   6615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   11668
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
   End
End
Attribute VB_Name = "FormVerProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(0) = flexAlignCenteright
    MSFlexGrid1.Text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 1200
    
        
    MSFlexGrid1.Col = 1
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(1) = flexAlignCenterright
    MSFlexGrid1.Text = "Descripcion"
    MSFlexGrid1.ColWidth(1) = 5000
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    MSFlexGrid1.Text = "UM"
    MSFlexGrid1.ColWidth(2) = 800
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    MSFlexGrid1.Text = "Precio Presupues."
    MSFlexGrid1.ColWidth(3) = 1800
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    MSFlexGrid1.Text = "Precio Factura"
    MSFlexGrid1.ColWidth(4) = 1800
 
 End Sub


Private Sub BotonAgregar_Click()

    FormProductos.Show

End Sub

Private Sub BotonEliminar_Click()

     On Error GoTo CapErr
     
     ruta = App.Path & "\DB_SPC_SI.mdb"
     Set db = DBEngine.OpenDatabase(ruta)
     Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
   
    tProductos.Index = "PrimaryKey"
    
    MSFlexGrid1.Col = 0
   ' MsgBox (MSFlexGrid1.Text)
            
    tProductos.Seek "=", MSFlexGrid1.Text
        
    M = MsgBox("¿Seguro desea eliminar el registro?", vbOKCancel, "INFO DEL SISTEMA")
        
    If M = 1 Then
        If Not tProductos.NoMatch Then
            tProductos.Delete
        End If
        Call muestrodatos
    End If

CapErr:
    Select Case Err
        Case 3021
          M = MsgBox("No se puede Eliminar, Producto con Movimientos", vbCritical, "INFO DEL SISTEMA")
          Exit Sub
    End Select

End Sub

Private Sub BotonGuardar_Click(Index As Integer)

    ruta = App.Path & "\DB_SPC_SI.mdb"
     Set db = DBEngine.OpenDatabase(ruta)
     Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
       MSFlexGrid1.Col = 0
       MSFlexGrid1.Row = 1
      Filas = MSFlexGrid1.Rows
      linea = 1
      Do While linea < Filas
          
           MSFlexGrid1.Row = linea
           MSFlexGrid1.Col = 0
          If MSFlexGrid1.Text <> " " Then
    
             MSFlexGrid1.Col = 0
             codigoprod = MSFlexGrid1.Text
        
             Dim busca1 As String, busca2 As String
             busca1 = RTrim(LTrim(codigoprod))
             busca2 = busca1 + "z"
        
             rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
             
             rstProductos.Edit
             MSFlexGrid1.Col = 0
             rstProductos.Fields!CodProd = MSFlexGrid1.Text
             
             MSFlexGrid1.Col = 1
             rstProductos.Fields!Descripcion = MSFlexGrid1.Text
                        
             MSFlexGrid1.Col = 2
             rstProductos.Fields!UnidadMedida = MSFlexGrid1.Text
                            
             MSFlexGrid1.Col = 3
             rstProductos.Fields!PrecioUnitarioPresupuesto = Format(MSFlexGrid1.Text, "#00.00")
                            
             MSFlexGrid1.Col = 4
             rstProductos.Fields!PrecioUnitarioFactura = Format(MSFlexGrid1.Text, "#00.00")
             
             rstProductos.Update
         
            
         End If
          linea = linea + 1
      Loop
      MSFlexGrid1.Clear
End Sub

Private Sub BotonMostrarProductos_Click(Index As Integer)

     Call muestrodatos
     
End Sub




Private Sub Command2_Click()

End Sub

Private Sub BotonSalir_Click()

    Unload Me

End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    FormVerProductos.Height = 8625
    FormVerProductos.Width = 11190
    FormVerProductos.Top = 1000
    FormVerProductos.Left = 1000
    
    Call muestrodatos

End Sub

Private Sub muestrodatos()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("select all * from Productos order by CodProd")
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
      
    If rstProductos.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Productos", vbCritical, "Final de la busqueda")
     
    End If
     
    linea2 = 1
    rstProductos.MoveFirst
    Do While Not rstProductos.EOF
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstProductos.Fields!CodProd
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstProductos.Fields!Descripcion
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstProductos.Fields!UnidadMedida
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#00.00")
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#00.00")
            linea2 = linea2 + 1
      
       rstProductos.MoveNext
       
    Loop
End Sub



Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii >= 32 And KeyAscii <= 127 Then
            MSFlexGrid1.Text = MSFlexGrid1.Text & Chr(KeyAscii)
       End If
       
    Select Case KeyAscii
    
       
       
       Case 13
      
            MSFlexGrid1.Col = 0
            codigoprodMA = UCase(MSFlexGrid1.Text)
       
       Case vbKeyBack
            
            If Len(MSFlexGrid1) >= 1 Then
               MSFlexGrid1 = Left$(MSFlexGrid1, Len(MSFlexGrid1) - 1)
            Else
                KeyAscii = 0
            End If
           
       End Select
       
        
       codigoprod = ""



End Sub
