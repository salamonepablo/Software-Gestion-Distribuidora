VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormConsignaciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Consignaciones"
   ClientHeight    =   9540
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9540
   ScaleWidth      =   12375
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureQP 
      Height          =   495
      Left            =   12000
      ScaleHeight     =   435
      ScaleWidth      =   315
      TabIndex        =   50
      Top             =   4560
      Width           =   375
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   43
      Top             =   3120
      Width           =   11655
      Begin VB.TextBox TextCantidad 
         Alignment       =   1  'Right Justify
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
         Left            =   7800
         TabIndex        =   5
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   10560
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextDescripcion 
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
         Left            =   3120
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.TextBox TextUnidadMedida 
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
         Left            =   7200
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox ComboArticulo 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad"
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
         Left            =   7800
         TabIndex        =   47
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Artículo"
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
         Left            =   2040
         TabIndex        =   46
         Top             =   360
         Width           =   675
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
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
         Left            =   4560
         TabIndex        =   45
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "UM"
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
         Left            =   7320
         TabIndex        =   44
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   34
      Top             =   2160
      Width           =   11655
      Begin VB.TextBox TextNumeroConsignacion 
         Alignment       =   1  'Right Justify
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
      Begin VB.TextBox TextFechaConsignacion 
         Alignment       =   2  'Center
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
         TabIndex        =   38
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   4680
         TabIndex        =   37
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TextSaldoCliente 
         Alignment       =   1  'Right Justify
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
         Left            =   7800
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox CheckModificaStock 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modifica Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9600
         TabIndex        =   35
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Consignación"
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
         Left            =   360
         TabIndex        =   42
         Top             =   240
         Width           =   1410
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Consignación"
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
         Left            =   2400
         TabIndex        =   41
         Top             =   240
         Width           =   1725
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
         TabIndex        =   40
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Saldo"
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
         Left            =   8160
         TabIndex        =   39
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   30
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton BotonDomicilio 
         BackColor       =   &H00808000&
         Caption         =   "&Domicilio Entrega"
         Enabled         =   0   'False
         Height          =   510
         Left            =   9960
         MaskColor       =   &H00808000&
         TabIndex        =   32
         Top             =   1320
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TextItemDomicilio 
         Height          =   285
         Left            =   4200
         TabIndex        =   31
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   21
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   11
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   10
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   23
      Top             =   8400
      Width           =   11655
      Begin VB.CommandButton cmdPDF 
         BackColor       =   &H00808000&
         Caption         =   "Generar &PDF"
         Enabled         =   0   'False
         Height          =   510
         Left            =   5280
         TabIndex        =   49
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00808000&
         Caption         =   "&Eliminar"
         Enabled         =   0   'False
         Height          =   510
         Left            =   6840
         TabIndex        =   48
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BotonGrabar 
         BackColor       =   &H00808000&
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   510
         Left            =   2040
         MaskColor       =   &H00808000&
         TabIndex        =   29
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BotonSalir 
         BackColor       =   &H00808000&
         Caption         =   "&Salir"
         Height          =   510
         Left            =   10080
         TabIndex        =   27
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BotonCancelar 
         BackColor       =   &H00808000&
         Caption         =   "&Cancelar"
         Height          =   510
         Left            =   8520
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00808000&
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   510
         Left            =   3720
         TabIndex        =   25
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton BotonNueva 
         BackColor       =   &H00808000&
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   510
         Left            =   360
         MaskColor       =   &H00808000&
         TabIndex        =   24
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   120
      TabIndex        =   22
      Top             =   4320
      Width           =   11655
      Begin VB.CommandButton BotonEliminarfila 
         Caption         =   "&Eliminar Fila"
         Height          =   495
         Left            =   9600
         TabIndex        =   33
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton BotonBuscarProducto 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   10800
         MaskColor       =   &H00808000&
         TabIndex        =   28
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3615
         Left            =   1320
         TabIndex        =   7
         Top             =   360
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
Attribute VB_Name = "FormConsignaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstConsignacionesC As DAO.Recordset
 Dim rstConsignacionesD As DAO.Recordset
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
 
Private Sub GenerarPDFDuplicado()

       'On Error GoTo CapturaErrores
        Dim NroConsigna As String
        'Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim Cant As Variant
        Dim Original, CantTotal As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tConsignasC = BaseSPC.OpenRecordset("ConsignacionesC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tConsignasC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tConsignasC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tConsignasC.Seek "=", TextNumeroConsignacion.text
            
           If Not tConsignasC.NoMatch Then
                
               ' If IsNull(tConsignasC!CAE) Then
               '     b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
               '     Exit Sub
               ' End If
                
                With Printer
                    
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 90
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "CONSIGNACION"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 8
                        Printer.Line (97, 16)-(116, 16)
                        Printer.Line (97, 16)-(97, 25)
                        Printer.Line (116, 16)-(116, 25)
                        Printer.Line (97, 25)-(116, 25)
                        
                        .CurrentX = 104
                        .CurrentY = 16
                        .FontSize = 18
                        .FontBold = True
                        Printer.Print "X"
                        'Printer.Print TextTipoFactura.Text
                        
                        .CurrentX = 98
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "No Válido Como Factura"
'                        .CurrentX = 94
 '                       .CurrentY = 24
 '                       .FontSize = 4
 '                       .FontBold = True
 '                        Printer.Print "Como Factura"
                         
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroConsigna = CStr(tConsignasC!NroConsignacion)
                        Largo = 8 - Len(tConsignasC!NroConsignacion)
                        For I = 1 To Largo
                            NroConsigna = "0" & NroConsigna
                        Next I
                        Printer.Print "Nº: 0004-" & NroConsigna
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tConsignasC!FechaC, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tConsignasC!IdCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                    Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                'Printer.Print tFacturaC!CondicionVenta
                                Printer.Print "En Consignación"
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                            '    Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                            '    NroRemito = CStr(tFacturaC!NroRemito)
                            '    LargoR = 8 - Len(tFacturaC!NroRemito)
                            '    For I = 1 To LargoR
                            '        NroRemito = "0" & NroRemito
                            '    Next I
                                
                             '   Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 10
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 95
                        .CurrentY = 92
                        .FontSize = 10
                        Printer.Print "DETALLE"
                       
                       'Imprimir Detalle de La Factura
                       
                        vSQL = "SELECT * FROM ConsignacionesD WHERE NroConsignacion=" & tConsignasC!NroConsignacion & " ORDER BY NroConsignacion, ItemConsignacion"
                        'MsgBox (vSQL)
                        
                        Set tConsignasD = BaseSPC.OpenRecordset(vSQL)
                        
                        tConsignasD.MoveFirst
                        'linea = .CurrentY + 3
                        
                        While Not tConsignasD.EOF
                            linea = .CurrentY + 3
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                          ' .FontBold = False
                            .FontSize = 10
                           Printer.Print Descripcion(tConsignasD!IdCodProd)
                            
                            .FontBold = True
                            .CurrentX = 18
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                            Cant = CDbl(tConsignasD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            CantTotal = CantTotal + Cant
                            tConsignasD.MoveNext
                        Wend
                        
                        'Importe Total
                            
                            Printer.Line (10, 250)-(200, 258), vbBlack, BF
                            
                            .CurrentX = 18
                            .CurrentY = 252
                            .Font = "Arial"
                            .FontSize = 12
                            .ForeColor = vbWhite
                            
                             Printer.Print "TOTAL DE BATERÍAS EN CONSIGNACIÓN: " & CantTotal & " ( " & StrConv(EnLetras(CStr(CantTotal)), vbUpperCase) & ")"
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Consignación Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub GenerarPDFOriginal()
                  
       'On Error GoTo CapturaErrores
        Dim NroConsigna As String
        'Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim Cant As Variant
        Dim Original, CantTotal As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tConsignasC = BaseSPC.OpenRecordset("ConsignacionesC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tConsignasC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tConsignasC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tConsignasC.Seek "=", TextNumeroConsignacion.text
            
           If Not tConsignasC.NoMatch Then
                
               ' If IsNull(tConsignasC!CAE) Then
               '     b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
               '     Exit Sub
               ' End If
                
                With Printer
                    
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 90
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "CONSIGNACION"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 8
                        Printer.Line (97, 16)-(116, 16)
                        Printer.Line (97, 16)-(97, 25)
                        Printer.Line (116, 16)-(116, 25)
                        Printer.Line (97, 25)-(116, 25)
                        
                        .CurrentX = 104
                        .CurrentY = 16
                        .FontSize = 18
                        .FontBold = True
                        Printer.Print "X"
                        'Printer.Print TextTipoFactura.Text
                        
                        .CurrentX = 98
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "No Válido Como Factura"
'                        .CurrentX = 94
 '                       .CurrentY = 24
 '                       .FontSize = 4
 '                       .FontBold = True
 '                        Printer.Print "Como Factura"
                         
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroConsigna = CStr(tConsignasC!NroConsignacion)
                        Largo = 8 - Len(tConsignasC!NroConsignacion)
                        For I = 1 To Largo
                            NroConsigna = "0" & NroConsigna
                        Next I
                        Printer.Print "Nº: 0004-" & NroConsigna
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tConsignasC!FechaC, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tConsignasC!IdCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                    Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                'Printer.Print tFacturaC!CondicionVenta
                                Printer.Print "En Consignación"
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                            '    Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                            '    NroRemito = CStr(tFacturaC!NroRemito)
                            '    LargoR = 8 - Len(tFacturaC!NroRemito)
                            '    For I = 1 To LargoR
                            '        NroRemito = "0" & NroRemito
                            '    Next I
                                
                             '   Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 10
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 95
                        .CurrentY = 92
                        .FontSize = 10
                        Printer.Print "DETALLE"
                       
                       'Imprimir Detalle de La Factura
                       
                        vSQL = "SELECT * FROM ConsignacionesD WHERE NroConsignacion=" & tConsignasC!NroConsignacion & " ORDER BY NroConsignacion, ItemConsignacion"
                        'MsgBox (vSQL)
                        
                        Set tConsignasD = BaseSPC.OpenRecordset(vSQL)
                        
                        tConsignasD.MoveFirst
                        'linea = .CurrentY + 3
                        
                        While Not tConsignasD.EOF
                            linea = .CurrentY + 3
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                          ' .FontBold = False
                            .FontSize = 10
                           Printer.Print Descripcion(tConsignasD!IdCodProd)
                            
                            .FontBold = True
                            .CurrentX = 18
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                            Cant = CDbl(tConsignasD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            CantTotal = CantTotal + Cant
                            tConsignasD.MoveNext
                        Wend
                        
                        'Importe Total
                            
                            Printer.Line (10, 250)-(200, 258), vbBlack, BF
                            
                            .CurrentX = 18
                            .CurrentY = 252
                            .Font = "Arial"
                            .FontSize = 12
                            .ForeColor = vbWhite
                            
                             Printer.Print "TOTAL DE BATERÍAS EN CONSIGNACIÓN: " & CantTotal & " ( " & StrConv(EnLetras(CStr(CantTotal)), vbUpperCase) & ")"
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Consignación Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub ImprimirConsignaDuplicado()

'On Error GoTo CapturaErrores
        Dim NroConsigna As String
        'Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim Cant As Variant
        Dim Original, CantTotal As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tConsignasC = BaseSPC.OpenRecordset("ConsignacionesC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tConsignasC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tConsignasC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tConsignasC.Seek "=", TextNumeroConsignacion.text
            
           If Not tConsignasC.NoMatch Then
                
               ' If IsNull(tConsignasC!CAE) Then
               '     b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
               '     Exit Sub
               ' End If
                
                With Printer
                    Set Printer = Printers(4)
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 90
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "CONSIGNACION"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 8
                        Printer.Line (97, 16)-(116, 16)
                        Printer.Line (97, 16)-(97, 25)
                        Printer.Line (116, 16)-(116, 25)
                        Printer.Line (97, 25)-(116, 25)
                        
                        .CurrentX = 104
                        .CurrentY = 16
                        .FontSize = 18
                        .FontBold = True
                        Printer.Print "X"
                        'Printer.Print TextTipoFactura.Text
                        
                        .CurrentX = 98
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "No Válido Como Factura"
'                        .CurrentX = 94
 '                       .CurrentY = 24
 '                       .FontSize = 4
 '                       .FontBold = True
 '                        Printer.Print "Como Factura"
                         
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroConsigna = CStr(tConsignasC!NroConsignacion)
                        Largo = 8 - Len(tConsignasC!NroConsignacion)
                        For I = 1 To Largo
                            NroConsigna = "0" & NroConsigna
                        Next I
                        Printer.Print "Nº: 0004-" & NroConsigna
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tConsignasC!FechaC, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tConsignasC!IdCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                    Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                'Printer.Print tFacturaC!CondicionVenta
                                Printer.Print "En Consignación"
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                            '    Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                            '    NroRemito = CStr(tFacturaC!NroRemito)
                            '    LargoR = 8 - Len(tFacturaC!NroRemito)
                            '    For I = 1 To LargoR
                            '        NroRemito = "0" & NroRemito
                            '    Next I
                                
                             '   Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 10
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 95
                        .CurrentY = 92
                        .FontSize = 10
                        Printer.Print "DETALLE"
                       
                       'Imprimir Detalle de La Factura
                       
                        vSQL = "SELECT * FROM ConsignacionesD WHERE NroConsignacion=" & tConsignasC!NroConsignacion & " ORDER BY NroConsignacion, ItemConsignacion"
                        'MsgBox (vSQL)
                        
                        Set tConsignasD = BaseSPC.OpenRecordset(vSQL)
                        
                        tConsignasD.MoveFirst
                        'linea = .CurrentY + 3
                        
                        While Not tConsignasD.EOF
                            linea = .CurrentY + 3
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                          ' .FontBold = False
                            .FontSize = 10
                           Printer.Print Descripcion(tConsignasD!IdCodProd)
                            
                            .FontBold = True
                            .CurrentX = 18
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                            Cant = CDbl(tConsignasD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            CantTotal = CantTotal + Cant
                            tConsignasD.MoveNext
                        Wend
                        
                        'Importe Total
                            
                            Printer.Line (10, 250)-(200, 258), vbBlack, BF
                            
                            .CurrentX = 18
                            .CurrentY = 252
                            .Font = "Arial"
                            .FontSize = 12
                            .ForeColor = vbWhite
                            
                             Printer.Print "TOTAL DE BATERÍAS EN CONSIGNACIÓN: " & CantTotal & " ( " & StrConv(EnLetras(CStr(CantTotal)), vbUpperCase) & ")"
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Consignación Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:


End Sub

Private Sub MuestraConsigna()
            
    TextNumeroConsignacion.text = tConsC!NroConsignacion
    TextFechaConsignacion.text = tConsC!FechaC
    'ComboVendedor.Text = BuscarNombreVendedor(tConsC!codVendedor)
    
    Call SeteoGrilla
    
    FG1.Col = 0
    FG1.Row = 1
        
    tConsD.MoveFirst
    
    While Not tConsD.EOF
        FG1.Col = 0
        FG1.text = tConsD!IdCodProd
        FG1.Col = 1
        FG1.text = Descripcion(tConsD!IdCodProd)
        FG1.Col = 2
        FG1.text = tConsD!UnidadMedida
        FG1.Col = 3
        FG1.text = tConsD!cantidad
        
        FG1.Row = FG1.Row + 1
        
        tConsD.MoveNext
    Wend
    
    cmdEliminar.Enabled = True
    cmdImprimir.Enabled = True
    cmdPDF.Enabled = True
    BotonNueva.Enabled = False
    
End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function

Private Function BuscarNombreVendedor(Leg As Integer) As String

    Set tVendedores = db.OpenRecordset("Empleados", dbOpenTable)

    tVendedores.Index = "PrimaryKey"
    
    tVendedores.Seek "=", Leg
    
    If Not tVendedores.NoMatch Then BuscarNombreVendedor = tVendedores!Nombre
    
    tVendedores.Close

End Function

Private Sub BotonAgregar_Click()
    If fila2 < renglon Then
        If Fila > 1 Then
         '   FG1.AddItem (ComboArticulo.Text)
            FG1.Row = Fila
        Else
            FG1.Row = 1
            'FG1.Col = 0
            'FG1.Text = ComboArticulo.Text
        End If
        
        FG1.Col = 0
        FG1.text = ComboArticulo.text
        FG1.Col = 1
        FG1.text = TextDescripcion.text
        FG1.Col = 2
        FG1.text = TextUnidadMedida.text
        FG1.Col = 3
        FG1.text = TextCantidad.text
                
        Fila = Fila + 1
        fila2 = fila2 + 1
            
    
        
        
'        ComboArticulo.Text = ""
        TextDescripcion.text = ""
        TextUnidadMedida.text = ""
        TextCantidad.text = ""
        
        
        ComboArticulo.SetFocus
    
  End If
        BotonGrabar.Enabled = True
        cmdImprimir.Enabled = True
End Sub

Private Sub BotonBuscarProducto_Click()

    FormBusquedaProductosRemito.Show

End Sub

Private Sub BotonCancelar_Click()

    Call blanqueototal
    
End Sub

Private Sub blanqueototal()

    TextCodigoCliente.text = ""
    TextApellidoNombre.text = ""
    TextCuit.text = ""
    TextDireccion.text = ""
    TextLocalidad.text = ""
    TextCodigoPostal.text = ""
    TextProvincia.text = ""
    TextNumeroConsignacion.text = ""
    ComboVendedor.text = ""
    TextSaldoCliente.text = ""
    FG1.Clear
    FG1.Enabled = False
    
    cmdEliminar.Enabled = False
    cmdImprimir.Enabled = False
    cmdPDF.Enabled = False
    BotonNueva.Enabled = True
     
    Call SeteoGrilla
    
'    MSFlexGrid1.Visible = False

End Sub

Private Sub BotonCancelar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonDomicilio_Click()

    FormDomiciliosRemito.Show
   
    
End Sub

Private Sub BotonEliminarfila_Click()

    If FG1.Row <= 0 Then
        MsgBox "Debe Seleccionar una fila"
    'ElseIf MSFlexGrid1.Row = 1 Then
    ' MSFlexGrid1.Clear
    Else
        FG1.RemoveItem (FG1.Row)
    End If
    
End Sub

Private Sub BotonGrabar_Click()

        Dim descuentoCantidad As Long
        Dim ultimo As Long
        Dim existeNumeroBD As Long
        Dim existeTipoBD As String
        Dim existeNumero As Long
        Dim existeTipo As String
        
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstConsignacionesC = db.OpenRecordset("ConsignacionesC", dbOpenDynaset)
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstConsignacionesD = db.OpenRecordset("ConsignacionesD", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
                
        existeNumero = TextNumeroConsignacion.text
        
        'buscamos el deposito para descontar el stock
        
        Set tDepositos = db.OpenRecordset("Depositos", dbOpenTable)
       
        On Error GoTo CapturaErrores
   
           tDepositos.Index = "IndXVendedor"
           
           tDepositos.MoveFirst
           tDepositos.Seek "=", LegajoEmpleado
           
           If Not tDepositos.NoMatch Then
            DepoOrigen = tDepositos!IDDEPOSITO
            'MsgBox (DepoOrigen)
           Else
            A = MsgBox("ERROR !!", vbCritical, "Vendedor sin Depósito Asociado")
           End If
              
           tDepositos.Close
           
            rstConsignacionesC.AddNew
            rstConsignacionesC.Fields!NroConsignacion = TextNumeroConsignacion.text
            rstConsignacionesC.Fields!FechaC = TextFechaConsignacion.text
            rstConsignacionesC.Fields!IdCliente = TextCodigoCliente.text
            rstConsignacionesC.Fields!codVendedor = LegajoEmpleado
            'rstConsignacionesC.Fields!ItemDomicilio = Val(TextItemDomicilio.Text)
            rstConsignacionesC.Update
            
            FG1.Col = 0
            FG1.Row = 1
            Filas = FG1.Rows
            linea = 1
            Do While linea < Filas
                  
                  FG1.Row = linea
                  FG1.Col = 0
                  If FG1.text <> "" Then
                        rstConsignacionesD.AddNew
                    
                        rstConsignacionesD.Fields!NroConsignacion = TextNumeroConsignacion.text
                        'rstConsignacionesD.Fields!TipoFactura = TextTipoFactura.Text
                    
                        FG1.Col = 0
                        rstConsignacionesD.Fields!IdCodProd = FG1.text
                    
                        FG1.Col = 2
                        rstConsignacionesD.Fields!UnidadMedida = FG1.text
                        
                        FG1.Col = 3
                        rstConsignacionesD.Fields!cantidad = Val(FG1.text)
                        descuentoCantidad = Val(FG1.text)
                        
                        FG1.Col = 4
                        rstConsignacionesD.Fields!ItemConsignacion = Val(FG1.text)
                        
                        
                        
                        
                        '*** Modifico Stock Producto
                       

                       'Call DesHagoStock(CodProd, descuentoCantidad)
                        
                        If modificaStock = 1 Then
                            FG1.Col = 0
                            codigoprod = FG1.text
                            
                            'Dim busca1 As String, busca2 As String
                            'busca1 = RTrim(LTrim(codigoprod))
                            'busca2 = busca1 + "z"
                       
                            'rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
                            
                            'rstProductos.Edit
                            'rstProductos.Fields!Stock = cantidadProducto - descuentoCantidad
                            'rstProductos.Update
                            
                            Call ActualizarStock(codigoprod, DepoOrigen, descuentoCantidad)
                        End If
                        
                       rstConsignacionesD.Update
                  End If
                  linea = linea + 1
            Loop
            
            '*** Actualizo Ultimo Numero Remito
            
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
        
            Dim busco As String
       
            'If TextTipoFactura.Text = "A" Then
                busco = "tConsignacionesC"
            'End If
            
            'If TextTipoFactura.Text = "B" Then
            '    busco = "tFacturaB"
            'End If
    
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
            rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
            ultimo = rstUltimosNumeros.Fields!UltimoNumero
            
            If ultimo < Val(TextNumeroConsignacion.text) Then
                rstUltimosNumeros.Edit
                'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
                     rstUltimosNumeros.Fields!UltimoNumero = TextNumeroConsignacion.text
                'End If
                rstUltimosNumeros.Update
            End If
        ' End If
        
        BotonGrabar.Enabled = False
        BotonNueva.Enabled = False
        BotonImprimir.Enabled = False
                        
        modificaStock = 0
        
        Call blanqueototal
        Call SeteoGrilla
        
        MSFlexGrid1.Visible = False
        
        TextCodigoCliente.SetFocus
        
CapturaErrores:
        
        Select Case Err
            Case 3021
                Resume Next
        End Select
        fila2 = 0
        Fila = 0
        
    
End Sub
Private Sub ActualizarStock(CodProd, IdDepoOrigen, Cant)

    'Sumo el Stock en Depósito Destino
        Set tS = db.OpenRecordset("Stock", dbOpenTable)
        
        tS.Index = "PrimaryKey"
        tS.MoveFirst
        
        'Resto el Stock en Depósito Origen
          tS.Seek "=", CodProd, IdDepoOrigen
            
        If Not tS.NoMatch Then
            tS.Edit
                tS!CodProd = CodProd
                tS!IDDEPOSITO = IdDepoOrigen
                tS!cantidad = tS.cantidad - FormatNumber(Cant, 2)
                tS!FechaUM = Format(Date, "DD/MM/YYYY")
            tS.Update
        End If
    
    'Sumo el Stock en Depósito Destino
    '    tS.Seek "=", CodProd, IdDepoDestino
              
        'Si tiene stock de este producto
     '       If Not tS.NoMatch Then
                'CantIni = tSotck!Stock
     '           tS.Edit
     '               tS.CodProd = CodProd
     '               tS.IdDeposito = IdDepoDestino
     '              tS.cantidad = tS!cantidad + FormatNumber(Cant, 2)
     '               tS.FechaUM = Format(Date, "DD/MM/YYYY")
     '           tS.Update
        'Si no tiene stock de este producto
     '        Else
     '           tS.AddNew
     '               tS!CodProd = CodProd
     '               tS!IdDeposito = IdDepoDestino
     '               tS!cantidad = FormatNumber(Cant, 2)
     '               tS!FechaUM = Format(Date, "DD/MM/YYYY")
     '           tS.Update
     '       End If
    
End Sub


Private Sub BotonGrabar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub




Private Sub BotonNueva_Click()

    Dim NumeroConsignacion As Long
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
 
    Dim busco As String
       
    busco = "tConsignacionesC"
        
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    NumeroConsignacion = rstUltimosNumeros.Fields!UltimoNumero
    
    'If rstUltimosNumeros.NoMatch Then
    '   FG1.Visible = False
    '   mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    'End If
    
    TextNumeroConsignacion.text = NumeroConsignacion + 1

    If TextCuit.text <> "" Then
       FG1.Enabled = True
    End If
    
    BotonNueva.Enabled = False
    
    FG1.Enabled = True
    FG1.Row = 1
    FG1.Col = 0
  
  ' *** AGREGADO POR PVS ***
    '  FG1.SetFocus
         TextNumeroConsignacion.SetFocus
    
End Sub



Private Sub BotonNueva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonSalir_Click()

    Unload FormConsignaciones

End Sub

Private Sub ImprimirConsigaOriginal()

       'On Error GoTo CapturaErrores
        Dim NroConsigna As String
        'Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim Cant As Variant
        Dim Original, CantTotal As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tConsignasC = BaseSPC.OpenRecordset("ConsignacionesC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tConsignasC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tConsignasC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tConsignasC.Seek "=", TextNumeroConsignacion.text
            
           If Not tConsignasC.NoMatch Then
                
               ' If IsNull(tConsignasC!CAE) Then
               '     b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
               '     Exit Sub
               ' End If
                
                With Printer
                    Set Printer = Printers(4)
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 90
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "CONSIGNACION"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 8
                        Printer.Line (97, 16)-(116, 16)
                        Printer.Line (97, 16)-(97, 25)
                        Printer.Line (116, 16)-(116, 25)
                        Printer.Line (97, 25)-(116, 25)
                        
                        .CurrentX = 104
                        .CurrentY = 16
                        .FontSize = 18
                        .FontBold = True
                        Printer.Print "X"
                        'Printer.Print TextTipoFactura.Text
                        
                        .CurrentX = 98
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "No Válido Como Factura"
'                        .CurrentX = 94
 '                       .CurrentY = 24
 '                       .FontSize = 4
 '                       .FontBold = True
 '                        Printer.Print "Como Factura"
                         
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroConsigna = CStr(tConsignasC!NroConsignacion)
                        Largo = 8 - Len(tConsignasC!NroConsignacion)
                        For I = 1 To Largo
                            NroConsigna = "0" & NroConsigna
                        Next I
                        Printer.Print "Nº: 0004-" & NroConsigna
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tConsignasC!FechaC, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tConsignasC!IdCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                    Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                'Printer.Print tFacturaC!CondicionVenta
                                Printer.Print "En Consignación"
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                            '    Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                            '    NroRemito = CStr(tFacturaC!NroRemito)
                            '    LargoR = 8 - Len(tFacturaC!NroRemito)
                            '    For I = 1 To LargoR
                            '        NroRemito = "0" & NroRemito
                            '    Next I
                                
                             '   Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 10
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 95
                        .CurrentY = 92
                        .FontSize = 10
                        Printer.Print "DETALLE"
                       
                       'Imprimir Detalle de La Factura
                       
                        vSQL = "SELECT * FROM ConsignacionesD WHERE NroConsignacion=" & tConsignasC!NroConsignacion & " ORDER BY NroConsignacion, ItemConsignacion"
                        'MsgBox (vSQL)
                        
                        Set tConsignasD = BaseSPC.OpenRecordset(vSQL)
                        
                        tConsignasD.MoveFirst
                        'linea = .CurrentY + 3
                        
                        While Not tConsignasD.EOF
                            linea = .CurrentY + 3
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                          ' .FontBold = False
                            .FontSize = 10
                           Printer.Print Descripcion(tConsignasD!IdCodProd)
                            
                            .FontBold = True
                            .CurrentX = 18
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                            Cant = CDbl(tConsignasD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            CantTotal = CantTotal + Cant
                            tConsignasD.MoveNext
                        Wend
                        
                        'Importe Total
                            
                            Printer.Line (10, 250)-(200, 258), vbBlack, BF
                            
                            .CurrentX = 18
                            .CurrentY = 252
                            .Font = "Arial"
                            .FontSize = 12
                            .ForeColor = vbWhite
                            
                             Printer.Print "TOTAL DE BATERÍAS EN CONSIGNACIÓN: " & CantTotal & " ( " & StrConv(EnLetras(CStr(CantTotal)), vbUpperCase) & ")"
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Consignación Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Public Function EnLetras(numero As String) As String
    
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "." Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
   
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
   
    flag = "N"
    
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
               
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
               
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                   
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
               
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
               
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
               
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
               
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
               
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
               
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
               
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                End If
            End If
            
            If paso = 7 Then
                MsgBox (Mid(entero, 1, 1))
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
       
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci & "/100"
            Else
                EnLetras = expresion & "con " & deci & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
       
End Function

Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
End Function

Private Sub BotonSalir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub CheckModificaStock_Click()

    If CheckModificaStock.Value = Unchecked Then
        modificaStock = 0
    End If
    
End Sub

Private Sub CmdEliminar_Click()

    Dim Rta As Integer

    Rta = MsgBox("¿Desea Eliminar La Consignación?", vbYesNoCancel, "CONFIRMAR ELIMINACION")
    
    If Rta = 6 Then
        
        tConsD.MoveFirst
        While Not tConsD.EOF
            tConsD.Delete
            tConsD.MoveNext
        Wend
        
        tConsC.Delete
       
        Call blanqueototal
        MSFlexGrid1.Visible = False
        
    Else
        If Rta = 7 Then Call blanqueototal
        If Rta = 2 Then TextCodigoCliente.SetFocus
    End If
        
End Sub

Private Sub cmdImprimir_Click()
    
    Call ImprimirConsigaOriginal
    Call ImprimirConsignaDuplicado

End Sub

Private Sub cmdImprimir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdPDF_Click()

    Call GenerarPDFOriginal
    MsgBox ("Genera Duplicado")
    Call GenerarPDFDuplicado

End Sub

Private Sub ComboVendedor_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(ComboVendedor.text))
    busca2 = busca1 + "z"
    
    rstEmpleado.FindFirst "Nombre >= '" & busca1 & "' and Nombre <= '" & busca2 & "'"
    
    If rstEmpleado.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       ComboVendedor.text = ""
       Call blanco
       ComboVendedor.SetFocus
    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
  
End Sub


Private Sub ComboVendedor_GotFocus()
    
    ComboVendedor.SelLength = Len(ComboVendedor.text)

End Sub

Private Sub ComboVendedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

 '**** PUSE ESTE IF PARA NO TENER QUE COMENTAR LINEA POR LINEA *** 'PVS' ****
 If PVS = 1 Then
    
    Dim precioUnitario As Double
    Dim cantidad As Integer
    Dim porcentaje As Double
    Dim total
    Dim totalLinea As Double
    Dim totalGrilla
    Dim subtotalPresuForm
    Dim porcentajePrecioUnitario As Double
    Dim descuentoPresup As Double
    Dim totalPresuForm As Double
    Dim iva As Double
    Dim impuesto As Double
    Dim percepcion As Double
    Dim columnaSeis As Integer
    Dim columnaSiete As Integer
    Dim bandera As Integer
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
     
    'Set db = DBEngine.OpenDatabase(ruta)
    'Set rstIva = db.OpenRecordset("Iva", dbOpenDynaset)

    'iva = rstIva.Fields!iva
    
    'If TextTipoFactura.Text = "A" Then
    '    Textiva.Text = Format(iva, "#,###,###,#0.00")
    'End If
        
    If KeyAscii >= 32 And KeyAscii <= 127 Then
        FG1.text = FG1.text & Chr(KeyAscii)
    End If

    Select Case KeyAscii
       Case 13
            
            FG1.Col = 0
            codigoprodMA = UCase(FG1.text)
                   
            '******* Busco Producto
            
           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
           tProductos.Index = "PrimaryKey"
           
           tProductos.MoveFirst
           tProductos.Seek "=", codigoprodMA
           
           If Not tProductos.NoMatch Then
                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(codigoprodMA))
                 busca2 = busca1 + "z"
                                     
                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
                 codigoProdTabla = rstProductos.Fields!CodProd
            
                cantidadProducto = rstProductos.Fields!Stock
                FG1.Col = 1
                FG1.text = rstProductos.Fields!Descripcion
                FG1.Col = 2
                FG1.text = rstProductos.Fields!UnidadMedida
                FG1.Col = 3
               
    
           Else
                 mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                 codigoprodMA = ""
                 FG1.Col = 1
                 FG1.text = ""
                 FG1.Col = 2
                 FG1.text = ""
                 FG1.Col = 3
                 FG1.text = ""
                 FG1.Col = 4
                 FG1.text = ""
                 FG1.Col = 0
                 FG1.text = ""
                 FG1.SetFocus
                 bandera = 1
           End If
              
           tProductos.Close
             
            
            
            '***********************
                   
            If FG1.Col = 3 And FG1.text <> "" Then
                FG1.Col = 0
                'If FG1.Row < 2 Then
                    FG1.Row = FG1.Row + 1
                    FG1.SetFocus
                    BotonGrabar.Enabled = True
                    BotonImprimir.Enabled = True
                'End If
            End If
     
             
       Case vbKeyBack
            
            If Len(FG1) >= 1 Then
               FG1 = Left$(FG1, Len(FG1) - 1)
            Else
                KeyAscii = 0
            End If
           
       End Select
       
        
       codigoprod = ""
  
  End If

End Sub
Private Sub muestrodatosproductos()

    cantidadProducto = rstProductos.Fields!Stock
    FG1.Col = 1
    FG1.text = rstProductos.Fields!Descripcion
    FG1.Col = 2
    FG1.text = rstProductos.Fields!UnidadMedida
    
           
End Sub


Private Sub Form_Load()

   FormConsignaciones.Height = 10110
   FormConsignaciones.Width = 12135
   FormConsignaciones.Top = 600
   FormConsignaciones.Left = 3000
   
    Fila = 1
    renglon = 16
    
    Call SeteoGrilla
    
    
    Call Cargo
    Call buscoarticulo
    
    TextFechaConsignacion.text = Format(Date, "dd/mm/yyyy")
    
    'bansera = 0
    modificaStock = 1
End Sub
Private Sub buscoarticulo()

    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    
   
    Do While Not rstProductos.EOF
        
       ComboArticulo.AddItem rstProductos!CodProd
       rstProductos.MoveNext
       
    Loop
    
    
End Sub

Private Sub ComboArticulo_Click()

    
    TextCantidad.text = ""
    
    
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
   
      
    Dim busca1 As String
    busca1 = RTrim(LTrim(ComboArticulo.text))
   
    
    rstProductos.FindFirst "CodProd >= '" & busca1 & "' "
    
   
    TextDescripcion.text = rstProductos.Fields!Descripcion
    TextUnidadMedida.text = rstProductos.Fields!UnidadMedida
    
      
   ' TextCantidad.SetFocus
    
    
    
End Sub
Private Sub ComboArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub
Private Sub ComboArticulo_LostFocus()
        
    On Error GoTo CE
   
    TextCantidad.text = ""
   
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    'If KeyAscii = 13 Then
      
           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
           tProductos.Index = "PrimaryKey"
           
           tProductos.MoveFirst
           tProductos.Seek "=", ComboArticulo.text
           
           If Not tProductos.NoMatch Then
                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(ComboArticulo.text))
                 busca2 = busca1 + "z"
                                     
                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
                             
                 TextDescripcion.text = rstProductos.Fields!Descripcion
                 
                
           Else
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                ComboArticulo.text = ""
                TextDescripcion.text = ""
                ComboArticulo.SetFocus
           End If
              
           tProductos.Close
          'TextCantidad.SetFocus
    'End If

CE:
Select Case Err

    Case 383
        Resume Next
End Select


End Sub

Sub SeteoGrilla()
    
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1000
    FG1.CellFontBold = True
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4700
    FG1.CellFontBold = True
    FG1.text = "Descripción"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 700
    FG1.CellFontBold = True
    FG1.text = "UM"
    FG1.ColAlignment(2) = flexAlignCenterCenter

    FG1.Col = 3
    FG1.ColWidth(3) = 900
    FG1.CellFontBold = True
    FG1.text = "Cantidad"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 0
    FG1.CellFontBold = True
    FG1.text = "Item"
    
    FG1.Row = 1
    item = 1
    linea = 1
    Do While FG1.Row <= 14
        FG1.Col = 4
        FG1.text = item
        item = (item + 1)
        FG1.Row = (Val(FG1.Row) + 1)
    Loop
        
    
    
End Sub

Private Sub Cargo()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    rstEmpleado.MoveFirst
    Do While Not rstEmpleado.EOF
        ComboVendedor.AddItem rstEmpleado!Nombre
        rstEmpleado.MoveNext
    Loop

End Sub

Private Sub busco()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextApellidoNombre.text))
    busca2 = busca1 + "z"
    
    rstCliente.FindFirst "Razonsocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
    
    If rstCliente.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       TextApellidoNombre.text = ""
       Call blanco
       TextApellidoNombre.SetFocus
    End If
     
    linea2 = 1
    Do While Not rstCliente.NoMatch
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.text = rstCliente.Fields!IdCliente
            MSFlexGrid1.Col = 1
            MSFlexGrid1.text = rstCliente.Fields!RazonSocial
            MSFlexGrid1.Col = 2
            If Not IsNull(rstCliente.Fields!CUIT) Then MSFlexGrid1.text = rstCliente.Fields!CUIT
            MSFlexGrid1.Col = 3
            MSFlexGrid1.text = rstCliente.Fields!Domicilio
            MSFlexGrid1.Col = 4
            MSFlexGrid1.text = rstCliente.Fields!localidad
            MSFlexGrid1.Col = 5
            MSFlexGrid1.text = rstCliente.Fields!CP
            MSFlexGrid1.Col = 6
            MSFlexGrid1.text = rstCliente.Fields!Prov
            MSFlexGrid1.Col = 7
            MSFlexGrid1.text = rstCliente.Fields!PorcentajeDescuento
            linea2 = linea2 + 1
      
       rstCliente.FindNext "RazonSocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
       
    Loop
    
    FG1.Enabled = True
    
End Sub

Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 900
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.text = "Apellido y Nombre"
    MSFlexGrid1.ColWidth(1) = 4700
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.text = "CUIT"
    MSFlexGrid1.ColWidth(2) = 1200
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.text = "Direccion"
    MSFlexGrid1.ColWidth(3) = 0
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.text = "Localidad"
    MSFlexGrid1.ColWidth(4) = 0
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.text = "CP"
    MSFlexGrid1.ColWidth(5) = 0
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.text = "Provincia"
    MSFlexGrid1.ColWidth(6) = 0
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.text = "Porcentaje Descuento"
    MSFlexGrid1.ColWidth(7) = 0

    
 End Sub

Private Sub MSFlexGrid1_Click()
   
    
    MSFlexGrid1.Col = 0
    TextCodigoCliente.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 1
    TextApellidoNombre.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 2
    TextCuit.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 3
    TextDireccion.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 4
    TextLocalidad.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 5
    TextCodigoPostal.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 6
    TextProvincia.text = MSFlexGrid1.text
        
    'MSFlexGrid1.Col = 7
    'descuentos = MSFlexGrid1.Text
    
    Call buscocuilyvendedor
    
    MSFlexGrid1.Visible = False
    
    FG1.Enabled = True

End Sub

Private Sub TextApellidoNombre_Change()
    Columna = 1
    Call FiltrarGrilla(MSFlexGrid1, TextApellidoNombre, CLng(Columna))
End Sub
Private Sub FiltrarGrilla(MSFlexGrid1 As Object, TBox As TextBox, Columna As Long)
    
    Dim A As Integer
    
    
    If (KeyRetroceso Or Len(TBox.text) = 0) Then
        'KeyRetroceso = False
        'Exit Sub
        TBox.text = ""
    End If
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    
    Call titulos
   
    A = Len(TBox.text)

    If A >= 4 Then
    
        vSQL = "SELECT * FROM Clientes WHERE RazonSocial Like '*" & TBox.text & "*' ORDER BY RazonSocial"
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        
        linea2 = 1
        
        Do While Not tClientes.EOF()
                MSFlexGrid1.AddItem " "
                MSFlexGrid1.Row = linea2
            
            
                MSFlexGrid1.Col = 0
                MSFlexGrid1.text = tClientes.Fields!IdCliente
                
                With Me.MSFlexGrid1

                    MSFlexGrid1.ColAlignment(1) = flexAlignLeftTop
                    MSFlexGrid1.Col = 0
                    MSFlexGrid1.text = tClientes.Fields!IdCliente
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.text = tClientes.Fields!RazonSocial
                    MSFlexGrid1.Col = 2
                    If tClientes.Fields!CUIT <> "" Then
                       MSFlexGrid1.text = tClientes.Fields!CUIT
                    End If
                    MSFlexGrid1.Col = 3
                    MSFlexGrid1.text = tClientes.Fields!Domicilio
                    MSFlexGrid1.Col = 4
                    MSFlexGrid1.text = tClientes.Fields!localidad
                    MSFlexGrid1.Col = 5
                    If tClientes.Fields!CP <> "" Then
                        MSFlexGrid1.text = tClientes.Fields!CP
                    End If
                    MSFlexGrid1.Col = 6
                    MSFlexGrid1.text = tClientes.Fields!Prov
                    MSFlexGrid1.Col = 7
                    MSFlexGrid1.text = tClientes.Fields!PorcentajeDescuento
                    
                End With
                linea2 = linea2 + 1
                tClientes.MoveNext
        Loop
    End If
MSFlexGrid1.Col = 4
'MSFlexGrid1.Sort = flexSortGenericAscending


End Sub

Private Sub TextApellidoNombre_GotFocus()
'    TextApellidoNombre.SelLength = Len(TextApellidoNombre.Text)
End Sub

Private Sub TextApellidoNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call busco
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub blanco()

    TextCodigoCliente.text = ""
    TextApellidoNombre.text = ""
    TextCuit.text = ""
    TextDireccion.text = ""
    TextLocalidad.text = ""
    TextCodigoPostal.text = ""
    TextProvincia.text = ""
    
End Sub


Private Sub TextCantidad_GotFocus()
    TextCantidad.SelLength = Len(TextCantidad.text)
End Sub

Private Sub TextCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub TextCantidad_LostFocus()

    If TextCantidad.text = "" Then
        A = MsgBox("NO PUEDE DEJAR LA CANTIDAD EN BLANCO", vbCritical, "E R R O R ! ! !")
        TextCantidad.SetFocus
    End If

End Sub

Private Sub TextCodigoCliente_GotFocus()
    
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.text)

End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
   
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    If KeyAscii = 13 Then
        If TextCodigoCliente.text = "" Then
            TextApellidoNombre.SetFocus
        Else
            CodigoClie = Val(TextCodigoCliente.text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.text = ""
                Call blanqueototal
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.text = rstCliente.Fields!IdCliente
                TextApellidoNombre.text = rstCliente.Fields!RazonSocial
                MSFlexGrid1.Visible = False
                TextCuit.text = rstCliente.Fields!CUIT
                TextDireccion.text = rstCliente.Fields!Domicilio
                TextLocalidad.text = rstCliente.Fields!localidad
                TextCodigoPostal.text = rstCliente.Fields!CP
                TextProvincia.text = rstCliente.Fields!Prov
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroConsignacion.text = ""
    End If
    
    If TextNumeroConsignacion <> "" Then
        FG1.Enabled = True
    Else
        FG1.Enabled = False
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub
Private Sub buscocuilyvendedor()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    '*** Busco CUIT
    
    CodigoClie = Val(TextCodigoCliente.text)
      
    rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
    
    TextCuit.text = rstCliente.Fields!CUIT
    codigovendedor = rstCliente!Vendedor
      
    
    '*** Busco Vendedor
    
    CodigoVend = codigovendedor
      
    rstEmpleado.FindFirst "Legajo >= '" & CodigoVend & "'"
    
    LegajoEmpleado = rstEmpleado.Fields!Legajo
    ComboVendedor.text = rstEmpleado.Fields!Nombre
    
    '*** Busco Saldo
    
   rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    
   TextSaldoCliente.text = Format(rstCtaCte.Fields!SaldoTotal, "#,###,###,#0.00")
   
   BotonDomicilio.Enabled = True
   
   BotonNueva.Enabled = True
   BotonNueva.SetFocus
   
    
End Sub

Private Sub TextCodigoCliente_LostFocus()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    'Busco si existen Consignaciones
        Set tConsC = db.OpenRecordset("ConsignacionesC", dbOpenTable)
        
        tConsC.Index = "SecondaryKey"
        tConsC.Seek "=", TextCodigoCliente.text
        
        If Not tConsC.NoMatch Then
            vSQL = "SELECT * FROM ConsignacionesD WHERE NroConsignacion =" & tConsC!NroConsignacion & " ORDER BY ItemConsignacion"
            'MsgBox (vSQL)
            Set tConsD = db.OpenRecordset(vSQL, dbOpenDynaset)
            Call MuestraConsigna
            Exit Sub
        End If
        
        
   

'    If KeyAscii = 13 Then
        If TextCodigoCliente.text = "" Then
            TextApellidoNombre.SetFocus
        Else
            CodigoClie = Val(TextCodigoCliente.text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.text = ""
                Call blanqueototal
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.text = rstCliente.Fields!IdCliente
                TextApellidoNombre.text = rstCliente.Fields!RazonSocial
                TextCuit.text = rstCliente.Fields!CUIT
                TextDireccion.text = rstCliente.Fields!Domicilio
                TextLocalidad.text = rstCliente.Fields!localidad
                TextCodigoPostal.text = rstCliente.Fields!CP
                TextProvincia.text = rstCliente.Fields!Prov
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroConsignacion.text = ""
 '   End If
    
    If TextNumeroConsignacion <> "" Then
        FG1.Enabled = True
    Else
        FG1.Enabled = False
    End If


End Sub

Private Sub TextCuit_Change()

    If TextCuit.text <> "" Then
        BotonNueva.Enabled = True
    End If
        
End Sub

Private Sub TextDescripcion_GotFocus()
    TextDescripcion.SelLength = Len(TextDescripcion.text)
End Sub

Private Sub TextDescripcion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextFechaConsignacion_GotFocus()
    TextFechaConsignacion.SelLength = Len(TextFechaConsignacion.text)
End Sub

Private Sub TextFechaConsignacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextItemDomicilio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextNumeroConsignacion_GotFocus()
    
    TextNumeroConsignacion.SelLength = Len(TextNumeroConsignacion.text)

End Sub

Private Sub TextNumeroConsignacion_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextUnidadMedida_GotFocus()
    TextUnidadMedida.SelLength = Len(TextUnidadMedida.text)
End Sub

Private Sub TextUnidadMedida_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextProvincia_Change()

    If TextProvincia.text <> "" Then
        ComboVendedor.SetFocus
    End If
End Sub


