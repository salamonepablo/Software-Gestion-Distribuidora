VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormVerRemito 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ver Remitos"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12105
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12105
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   11655
      Begin VB.TextBox TextNumeroRemito 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaRemito 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   23
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   4680
         TabIndex        =   22
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Remito"
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
         Left            =   480
         TabIndex        =   26
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Remito"
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
         Left            =   2640
         TabIndex        =   25
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Vendedor"
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
         Left            =   4680
         TabIndex        =   24
         Top             =   240
         Width           =   825
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextItemDomicilio 
         Height          =   285
         Left            =   4200
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextDireccion 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   15
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   5
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextProvincia 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Cliente:"
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
         TabIndex        =   14
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apellido Nombre:"
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
         Left            =   5160
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CUIT:"
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
         TabIndex        =   12
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dirección:"
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
         Left            =   5160
         TabIndex        =   11
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código Postal:"
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
         Left            =   4320
         TabIndex        =   10
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Localidad:"
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
         TabIndex        =   9
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia:"
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
         Left            =   7080
         TabIndex        =   8
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   8400
      Width           =   11655
      Begin VB.CommandButton BotonSalir 
         BackColor       =   &H00808000&
         Caption         =   "&Salir"
         Height          =   750
         Left            =   5880
         TabIndex        =   19
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonImprimir 
         BackColor       =   &H00808000&
         Caption         =   "&Imprimir"
         Height          =   750
         Left            =   4680
         TabIndex        =   18
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   5055
      Left            =   120
      TabIndex        =   16
      Top             =   3240
      Width           =   11655
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3615
         Left            =   1680
         TabIndex        =   27
         Top             =   720
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   16
         Cols            =   5
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
   End
End
Attribute VB_Name = "FormVerRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstRemitoC As DAO.Recordset
 Dim rstRemitoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim rstUltimosNumeros As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim vendedorCliente As String
 Dim nombreVendedor As Integer
 Dim LegajoEmpleado As String
 Dim modificaStock As Integer
 Dim Fila As Integer
 Dim fila2 As Integer
 Dim renglon As Integer
 Dim codCli


Private Sub BotonImprimir_Click()

    Call ImprimirRemito

End Sub
Private Sub ImprimirRemito()

    Dim RemC
    Dim RemD
        
    'On Error GoTo CapturaErrores

    X = -4
    Y = -4
          renglon = 0
    vNroRemito = "0002- " & TextNumeroRemito.Text
    
    vSQLRc = "SELECT * FROM RemitoC WHERE NroRemito='" & TextNumeroRemito.Text & "' ORDER By NroRemito"
    vSQLRd = "SELECT * FROM RemitoD WHERE NroRemito='" & TextNumeroRemito.Text & "' ORDER By NroRemito"
    vSQLRdir = "SELECT * FROM RemitoD WHERE NroRemito='" & TextNumeroRemito.Text & "' ORDER By NroRemito"
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    
    Set RemC = BaseSPC.OpenRecordset(vSQLRc, dbOpenDynaset)
    Set RemD = BaseSPC.OpenRecordset(vSQLRd, dbOpenDynaset)
      
        
    'With p
        'Seteo escala a mm
            Printer.Copies = 3
            Printer.ScaleMode = 6
        
        'Imprimir Fecha
            Printer.CurrentX = X + 130
            Printer.CurrentY = Y + 32
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print Format(FormVerRemito.TextFechaRemito.Text, "DD/MM/YYYY")
        
        'Imprimir Nombres
           Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 57
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = True
            Printer.Print FormVerRemito.TextApellidoNombre.Text
            
        'Imprimir Direccion
            Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 64
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormVerRemito.TextDireccion.Text
            
        'Imprimir Localidad
            Printer.CurrentX = X + 40
            Printer.CurrentY = Y + 70
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormVerRemito.TextLocalidad.Text
            
        'Imprimir CUIT
            Printer.CurrentX = X + 125
            Printer.CurrentY = Y + 70
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print FormVerRemito.TextCuit.Text
            
        'Imprimir Marca Responsable Inscripto
            Printer.CurrentX = X + 115
            Printer.CurrentY = Y + 76
            Printer.Font = "Courier New"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
         '   Printer.CurrentX = X + 70
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Marca CtaCte
         '   Printer.CurrentX = X + 100
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Nro Remito
            Printer.CurrentX = X + 138
            Printer.CurrentY = Y + 80
            Printer.Font = "Courier New"
            Printer.FontSize = 9
            Printer.FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
       '     sqlFC = "SELECT * FROM FacturaC WHERE TipoFactura='" & TextTipoFactura.Text & "' AND NroFactura=" & TextNumeroFactura.Text & " ORDER By NroFactura"
       '     vsqlFD = "SELECT * FROM FacturaD WHERE TipoFactura='" & TextTipoFactura.Text & "' AND NroFactura=" & TextNumeroFactura.Text & " ORDER By NroFactura"
            
          '  Set RemC = BaseSPC.OpenRecordset(vsqlFC, dbOpenDynaset)
          '  Set RemD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            RemC.MoveFirst
            RemD.MoveFirst
                
                    While Not RemD.EOF
                        'Imprimo el detalle
                            Printer.CurrentX = X + 30
                            Printer.CurrentY = Y + 96 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            Printer.Print RemD!cantidad
                            
                        'Detalle
                            Printer.CurrentX = X + 50
                            Printer.CurrentY = Y + 96 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            'Printer.Print RemD!IdCodProd & Chr(9) & Descripcion(RemD!IdCodProd)
                            Printer.Print Chr(9) & Descripcion(RemD!IdCodProd)
                        
                         renglon = renglon + 5
                            
                        RemD.MoveNext
                    Wend
        Printer.EndDoc
        
'    End With
    
    RemC.Close
    RemD.Close
    
    BaseSPC.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If
End Sub
Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function

Private Sub BotonSalir_Click()

    Unload FormVerRemito

End Sub



Private Sub Form_Load()

   FormVerRemito.Height = 10110
   FormVerRemito.Width = 12135
   FormVerRemito.Top = 600
   FormVerRemito.Left = 50
   
   numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
   'tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
   codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
    
    
    If Val(FormBuscarRemito.TextA) = 1 Then
        codCli = Val(FormBuscarRemito.TextCodigoCliente)
        numDoc = FormBuscarRemito.TextNumeroFactura
        TextNumeroRemito.Text = numDoc
         numDoc = TextNumeroRemito.Text
        Call SeteoGrilla
        Call buscofactura
    Else
        Call SeteoGrilla
        Call buscofactura
    End If
   
   numDoc = TextNumeroRemito.Text
End Sub
Private Sub buscofactura()


    Dim busca1 As String, busca2 As String
    Dim busca3 As String, busca4 As String
    Dim busca5 As String, busca6 As String
    Dim busca7 As String, busca8 As String
    Dim busca10 As String, busca11 As String

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstRemitoC = db.OpenRecordset("RemitoC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstRemitoD = db.OpenRecordset("RemitoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
      
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
   '************ Busco Vendedor
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
   
    busca3 = RTrim(LTrim(ComboVendedor.Text))
    busca4 = busca3 + "z"
    
    rstEmpleado.FindFirst "Nombre >= '" & busca3 & "' and Nombre <= '" & busca4 & "'"
    
'    If rstEmpleado.NoMatch Then
'       MSFlexGrid1.Visible = False
'       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
'
'    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
    
   
    rstCliente.FindFirst "IDCliente= " + Str(codCli)
   
    TextCodigoCliente.Text = rstCliente.Fields!IdCliente
    TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
    If rstCliente.Fields!CUIT <> "" Then TextCuit.Text = rstCliente.Fields!CUIT
    TextDireccion.Text = rstCliente.Fields!Domicilio
    TextLocalidad.Text = rstCliente.Fields!Localidad
    TextCodigoPostal.Text = rstCliente.Fields!CP
    TextProvincia.Text = rstCliente.Fields!Prov
   
  
    
    Call SeteoGrilla
 
'    rstRemitoC.FindFirst "NroRemito= " + Str(numDoc)
     
     numDoc = TextNumeroRemito.Text
     
     busca5 = RTrim(LTrim(numDoc))
     busca6 = busca5 + "z"
            
     rstRemitoC.FindFirst "NroRemito >= '" & busca5 & "' and NroRemito <= '" & busca6 & "'"
    
    TextNumeroRemito.Text = rstRemitoC.Fields!NroRemito
    TextFechaRemito.Text = rstRemitoC.Fields!FechaRemito
    
'    rstRemitoD.FindFirst "NroRemito= " + Str(numDoc)

    numDoc = TextNumeroRemito.Text
    
    busca7 = RTrim(LTrim(numDoc))
    busca8 = busca7 + "z"
            
    rstRemitoD.FindFirst "NroRemito >= '" & busca7 & "' and NroRemito <= '" & busca8 & "'"
    linea2 = 1
    Do While Not rstRemitoD.NoMatch
        
            FG1.AddItem " "
            FG1.Row = linea2
       
            FG1.Col = 0
            FG1.Text = rstRemitoD.Fields!IdCodProd
            
            FG1.Col = 0
            codigoprod = FG1.Text

            
            busca10 = RTrim(LTrim(codigoprod))
            busca11 = busca10 + "z"
       
            rstProductos.FindFirst "CodProd >= '" & busca10 & "' and CodProd <= '" & busca11 & "'"
            
            FG1.Col = 1
            FG1.Text = rstProductos.Fields!Descripcion
        
            FG1.Col = 2
            FG1.Text = rstRemitoD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.Text = rstRemitoD.Fields!cantidad
'            FG1.Col = 4
'            FG1.Text = rstRemitoD.Fields!item
           
       
'           rstRemitoD.FindNext "NroRemito= " + Str(numDoc)
            rstRemitoD.FindNext "NroRemito >= '" & busca7 & "' and NroRemito <= '" & busca8 & "'"
           linea2 = linea2 + 1
    Loop
    
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstRemitoC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.Text = rstEmpleado.Fields!nombre

    '****
    
   
    
End Sub





Sub SeteoGrilla()
    
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1000
    FG1.CellFontBold = True
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.Text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4700
    FG1.CellFontBold = True
    FG1.Text = "Descripción"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 700
    FG1.CellFontBold = True
    FG1.Text = "UM"
    FG1.ColAlignment(2) = flexAlignCenterCenter

    FG1.Col = 3
    FG1.ColWidth(3) = 900
    FG1.CellFontBold = True
    FG1.Text = "Cantidad"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 0
    FG1.CellFontBold = True
    FG1.Text = "Item"
End Sub


