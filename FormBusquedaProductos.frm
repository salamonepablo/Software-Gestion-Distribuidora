VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormBusquedaProducto 
   Caption         =   "Busqueda Producto"
   ClientHeight    =   5040
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   9735
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   5
      Top             =   3600
      Width           =   9375
      Begin VB.CommandButton BotonSalir 
         Caption         =   "Salir"
         Height          =   750
         Left            =   7920
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   9495
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   2175
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3836
         _Version        =   393216
         Rows            =   10
         Cols            =   4
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   9495
      Begin VB.TextBox TextCodigoProducto 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox TextNombreProducto 
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Top             =   240
         Width           =   4095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código Producro:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   " Nombre Producto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   3
         Top             =   240
         Width           =   1605
      End
   End
End
Attribute VB_Name = "FormBusquedaProducto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rstProductos As DAO.Recordset

Private Sub BotonSalir_Click()
   
    Unload FormBusquedaProducto

End Sub

Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSFlexGrid1.Text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 1200
    
        
    MSFlexGrid1.Col = 1
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
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
    MSFlexGrid1.Text = "Precio Unitario"
    MSFlexGrid1.ColWidth(3) = 1400
 
 End Sub

Private Sub Form_Load()

    FormBusquedaProducto.Height = 5610
    FormBusquedaProducto.Width = 9975

End Sub

Private Sub buscoPorDescripcion()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextNombreProducto.Text))
    busca2 = busca1 + "z"
    
    rstProductos.FindFirst "Descripcion >= '" & busca1 & "' and Descripcion <= '" & busca2 & "'"
    
    If rstProductos.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Productos", vbCritical, "Final de la busqueda")
       TextNombreProducto.Text = ""
       Call blanco
       TextNombreProducto.SetFocus
    End If
     
    linea2 = 1
    Do While Not rstProductos.NoMatch
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstProductos.Fields!CodProd
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstProductos.Fields!Descripcion
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstProductos.Fields!UnidadMedida
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#000.00")
            linea2 = linea2 + 1
      
       rstProductos.FindNext "Descripcion >= '" & busca1 & "' and Descripcion <= '" & busca2 & "'"
       
    Loop
 
End Sub

Private Sub buscoPorCodigo()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextCodigoProducto.Text))
    busca2 = busca1 + "z"
    
    rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
    
    If rstProductos.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Productos", vbCritical, "Final de la busqueda")
       TextCodigoProducto.Text = ""
       Call blanco
       TextCodigoProducto.SetFocus
    End If
     
    linea2 = 1
    Do While Not rstProductos.NoMatch
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstProductos.Fields!CodProd
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstProductos.Fields!Descripcion
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstProductos.Fields!UnidadMedida
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#000.00")
            linea2 = linea2 + 1
      
       rstProductos.FindNext "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
       
    Loop
 
End Sub

Private Sub MSFlexGrid1_DblClick()

    If FormFactura.FG1.Row = 0 Then
        FormFactura.FG1.Row = FormFactura.FG1.Row + 1
    End If
    
    FormFactura.FG1.Col = 0
    
    'If FormFactura.FG1.Text <> "" Then
        
        If FormFactura.FG1.Row >= 15 Then
             mensaje = MsgBox("No se puede incorporar mas productos a la Factuta", vbCritical, "Final Factura")
        Else
             'FormFactura.FG1.Row = FormFactura.FG1.Row + 1
        
             MSFlexGrid1.Col = 0
             FormFactura.FG1.Col = 0
             FormFactura.FG1.Text = MSFlexGrid1.Text
        
             MSFlexGrid1.Col = 1
             FormFactura.FG1.Col = 1
             FormFactura.FG1.Text = MSFlexGrid1.Text
         
             MSFlexGrid1.Col = 2
             FormFactura.FG1.Col = 2
             FormFactura.FG1.Text = MSFlexGrid1.Text
         
             MSFlexGrid1.Col = 3
             FormFactura.FG1.Col = 3
             FormFactura.FG1.Text = MSFlexGrid1.Text
        End If
        
    'End If
    
    Unload FormBusquedaProducto

End Sub

Private Sub TextCodigoProducto_GotFocus()
    TextCodigoProducto.SelLength = Len(TextCodigoProducto.Text)
End Sub

Private Sub TextCodigoProducto_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then
        Call buscoPorCodigo
    End If

End Sub

Private Sub TextNombreProducto_GotFocus()
    TextNombreProducto.SelLength = Len(TextNombreProducto.Text)
End Sub

Private Sub TextNombreProducto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call buscoPorDescripcion
    End If
    
End Sub

Private Sub blanco()

    TextNombreProducto.Text = ""
    TextCodigoProducto.Text = ""
    
End Sub

