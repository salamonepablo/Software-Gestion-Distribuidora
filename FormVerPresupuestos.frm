VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormVerPresupuestos 
   Caption         =   "Cosulta Presupuestos"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12090
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00808000&
         Caption         =   "&Imprimir"
         Height          =   750
         Left            =   3480
         MaskColor       =   &H00808000&
         TabIndex        =   36
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGrabar 
         BackColor       =   &H00808000&
         Caption         =   "&Anular"
         Height          =   750
         Left            =   4560
         MaskColor       =   &H00808000&
         TabIndex        =   34
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   5640
         TabIndex        =   32
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextProvincia 
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
         Left            =   8280
         TabIndex        =   16
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
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
         Left            =   5640
         TabIndex        =   15
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextLocalidad 
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
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCuit 
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
         Left            =   1920
         TabIndex        =   13
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
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
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
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
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextDireccion 
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
         Left            =   7080
         TabIndex        =   11
         Top             =   600
         Width           =   4335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
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
         TabIndex        =   23
         Top             =   960
         Width           =   870
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         TabIndex        =   22
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
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
         TabIndex        =   21
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         TabIndex        =   20
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         TabIndex        =   19
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         TabIndex        =   17
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   24
      Top             =   6240
      Width           =   11655
      Begin VB.TextBox TextTotalPresupuesto 
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
         Left            =   9840
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextDescuentos 
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
         Left            =   2400
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextSubtotalPresupuesto 
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
         Left            =   480
         TabIndex        =   25
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Anulado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4920
         TabIndex        =   35
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Presupuesto:"
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
         Left            =   9720
         TabIndex        =   30
         Top             =   240
         Width           =   1620
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Descuentos:"
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
         Left            =   2520
         TabIndex        =   29
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal Presupuesto:"
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
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   1890
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4455
      Left            =   120
      TabIndex        =   1
      Top             =   1680
      Width           =   11655
      Begin VB.TextBox TextNumeroPresupuesto 
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
         Left            =   120
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaPresupuesto 
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
         Left            =   2400
         TabIndex        =   4
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
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
         Height          =   315
         Left            =   4560
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.TextBox TextDescuentoCliente 
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
         Left            =   7800
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3255
         Left            =   480
         TabIndex        =   33
         Top             =   960
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Presupuesto"
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
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Presupuesto"
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
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
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
         Left            =   5160
         TabIndex        =   7
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Descuento"
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
         Left            =   7560
         TabIndex        =   6
         Top             =   240
         Width           =   930
      End
   End
End
Attribute VB_Name = "FormVerPresupuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstPresupuestoC As DAO.Recordset
 Dim rstPresupuestoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim descuentos As Double
 Dim LegajoEmpleado As Integer
 Dim numDoc As Long
 Dim tipoDoc As String
 Dim codCli As Integer


Private Sub BotonGrabar_Click()
        
        Dim descuentoCantidad As Integer
        Dim ultimo As Long
        Dim existeNumeroBD As Integer
        Dim existeTipoBD As String
        Dim existeNumero As Long
        Dim existeTipo As String
        
       
        
        ruta = App.Path & "\DB_SPC_SI.mdb"
        
         Set db = DBEngine.OpenDatabase(ruta)
        Set rstPresupuestoC = db.OpenRecordset("PresupuestoC", dbOpenDynaset)
   
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
        
         respuesta = MsgBox("Esta Seguro de Anular el Presupuesto?", vbYesNo, "Pago")
         If respuesta = vbYes Then
        
             existeNumero = TextNumeroPresupuesto.text
            
             'buscamos el deposito para aumentar el stock
             
             Set tDepositos = db.OpenRecordset("Depositos", dbOpenTable)
           '     On Error GoTo CapturaErrores
        
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
                
                FG1.Col = 5
                descuentoCantidad = Val(FG1.text)
        
                 FG1.Col = 0
                 FG1.Row = 1
                 Filas = FG1.Rows
                 linea = 1
                 Do While linea < Filas
                       
                       FG1.Row = linea
                       FG1.Col = 0
                       If FG1.text <> "" Then
                       
                            
                                 FG1.Col = 0
                                 codigoprod = FG1.text
                                 
                                 Call ActualizarStock(codigoprod, DepoOrigen, descuentoCantidad)
                            
                             
                       End If
                       linea = linea + 1
                 Loop
                 
                 '*** Grabo Linea 2 en Cuenta Corriente
                 
                 CodigoClie = Val(TextCodigoCliente.text)
           
                 rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
                 If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                     mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                     'TextCodigoCliente.Text = ""
                     'Call blanqueototal
                     'TextCodigoCliente.SetFocus
                 Else
                     rstCtaCte.Edit
                     saldo1 = Format(rstCtaCte.Fields!SaldoL1, "#,###,###,#0.00")
                     saldo2 = Format(rstCtaCte.Fields!SaldoL2, "#,###,###,#0.00")
                     saldoLi2 = Format(TextTotalPresupuesto.text, "#,###,###,#0.00")
                     rstCtaCte.Fields!SaldoL2 = saldo2 - saldoLi2
                     saldo2 = Format(rstCtaCte.Fields!SaldoL2, "#,###,###,#0.00")
                     FormMovimientosCuentaCorriente.TextSaldoLinea2 = Format(saldo2, "#,###,###,#0.00")
                     rstCtaCte.Fields!SaldoTotal = Format(CDbl(saldo2) + CDbl(saldo1), "#,###,###,#0.00")
                     FormMovimientosCuentaCorriente.TextSaldoTotal.text = Format(rstCtaCte.Fields!SaldoTotal, "#,###,###,#0.00")
                     rstCtaCte.Update
                 End If
             
                 
                 '*** Grabo Movimientos Cuente corriente
             
                 rstMovimientosCtaCte.AddNew
                 rstMovimientosCtaCte.Fields!NroDoc = 99 & TextNumeroPresupuesto.text
                 'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
                 rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaPresupuesto.text, "dd/mm/yyyy")
                 rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.text
                 rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Presupuesto Nº " & TextNumeroPresupuesto.text & ""
                 
                 
                 rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
                 rstMovimientosCtaCte.Fields!ImporteLinea2 = Format(TextTotalPresupuesto.text, "-#0.00")
                 rstMovimientosCtaCte.Update
                 
                  CodigoPa = Val(TextNumeroPresupuesto.text)
                  rstPresupuestoC.FindFirst "NroPresu= " + Str(CodigoPa)
                  rstPresupuestoC.Edit
                    rstPresupuestoC.Fields!Anulado = "si"
                  rstPresupuestoC.Update
             
             
             BotonGrabar.Value = False
             
        End If
CapturaErrores:
        Select Case Err
            Case 3021
                Resume Next
        End Select
        
 Unload FormVerPresupuestos
 
End Sub
Private Sub ActualizarStock(CodProd, IdDepoOrigen, Cant)

    'Sumo el Stock en Depósito Destino
        Set tS = db.OpenRecordset("Stock", dbOpenTable)
        
        tS.Index = "PrimaryKey"
        tS.MoveFirst
        
        'Sumo el Stock en Depósito Origen
          tS.Seek "=", CodProd, IdDepoOrigen
            
        If Not tS.NoMatch Then
            tS.Edit
                tS!CodProd = CodProd
                tS!IDDEPOSITO = IdDepoOrigen
                tS!cantidad = tS.cantidad + FormatNumber(Cant, 2)
                tS!FechaUM = Format(Date, "DD/MM/YYYY")
            tS.Update
        End If
    
    


End Sub

Private Sub BotonSalir_Click()

    Unload FormVerPresupuestos

End Sub

Private Sub cmdImprimir_Click()

    Dim PU, TL, Cant, TotalPres As Variant
    'PU = 0
    'TL = 0
    x = -4
    Y = -14
    renglon = 0
     
    With Printer
        'On Error GoTo CapturaErrores
        
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Cantidad de Impresiones
            .Copies = 3
            
        'Imprimir Codigo de Cliente
            .CurrentX = x + 25
            .CurrentY = Y + 25
            .Font = "Courier New"
            .FontSize = 16
            .FontBold = True
            .ForeColor = vbRed
            Printer.Print TextCodigoCliente.text
        
        'Imprimir Fecha
            .ForeColor = vbBlack
            .CurrentX = x + 120
            .CurrentY = Y + 38
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaPresupuesto.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = x + 37
            .CurrentY = Y + 49
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextApellidoNombre.text
            
        'Imprimir Direccion
            .CurrentX = x + 37
            .CurrentY = Y + 56
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            'Printer.Print TextDireccion.Text & Chr(9) & Chr(9) & Chr(9) & Chr(9) & TextLocalidad.Text
            Printer.Print TextDireccion.text
        
        'Imprimir Localidad
            .CurrentX = x + 120
            .CurrentY = Y + 56
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextLocalidad.text
            
            
        'Imprimir Detalle
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
            vSQLPC = "SELECT * FROM PresupuestoC WHERE NroPresu=" & TextNumeroPresupuesto.text & " ORDER By NroPresu"
            vSQLPD = "SELECT * FROM PresupuestoD WHERE NroPresu=" & TextNumeroPresupuesto.text & " ORDER By NroPresu"
            
            Set PresuC = BaseSPC.OpenRecordset(vSQLPC, dbOpenDynaset)
            Set PresuD = BaseSPC.OpenRecordset(vSQLPD, dbOpenDynaset)
            
           
            PresuC.MoveFirst
            PresuD.MoveFirst
                
                    While Not PresuD.EOF
                        'Imprimo el detalle
                            .CurrentX = x + 13
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            
                            Cant = CDbl(PresuD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print PresuD!cantidad
                            
                        'Detalle
                            .CurrentX = x + 30
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'Printer.Print PresuD!IdCodProd & Chr(9) & Descripcion(PresuD!IdCodProd)
                            Printer.Print BuscarDescProd(PresuD!CodProd)
                        
                        'Precio
                            .CurrentX = x + 115
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(PresuD!precioUnitario) - (CDbl(PresuD!precioUnitario) * CDbl(PresuD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU
                        
                        'Importe
                            .CurrentX = x + 130
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            TL = Format(PresuD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                        
                         renglon = renglon + 5
                            
                        PresuD.MoveNext
                    Wend
            
            'Importe Total
                .CurrentX = x + 130
                .CurrentY = Y + 197
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                
                TotalPres = Format(CDbl(PresuC!TotalPresu), "Standard")
                Hasta = CInt(14 - Len(TotalPres))
                For I = 0 To Hasta
                    TotalPres = " " & TotalPres
                Next I
                
                Printer.Print TotalPres
            
        .EndDoc
        
    End With
    
    PresuC.Close
    PresuD.Close
    BaseSPC.Close
    
    'Call blanqueototal
    'MSFlexGrid1.Visible = False
    'TextCodigoCliente.SetFocus
    'Unload FormPagoFacturasDesdeFactura
        
CapturaErrores:
    'If Err = 321 Then
    'End If

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
       'Call blanco
       
    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
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
    FG1.ColWidth(3) = 1100
    FG1.CellFontBold = True
    FG1.text = "Precio Unit."
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 720
    FG1.CellFontBold = True
    FG1.text = "% Desc."
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    FG1.Col = 5
    FG1.ColWidth(5) = 900
    FG1.CellFontBold = True
    FG1.text = "Cantidad"
    FG1.ColAlignment(5) = flexAlignCenterCenter
        
    FG1.Col = 6
    FG1.ColWidth(6) = 1100
    FG1.CellFontBold = True
    FG1.text = "Total Línea"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 0
    FG1.CellFontBold = True
    FG1.text = "Importe Descuento"
    
   
    
End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

    FormVerPresupuestos.Height = 8970
    FormVerPresupuestos.Width = 12135
    FormVerPresupuestos.Top = 1000
    FormVerPresupuestos.Left = 1000
    
    numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
    'tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
    codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
    
    
    If Val(FormBuscarPresupuesto.TextA) = 1 Then
        codCli = Val(FormBuscarPresupuesto.TextCodigoCliente)
        numDoc = Val(FormBuscarPresupuesto.TextNumeroFactura)
        Call SeteoGrilla
        Call buscofactura
    Else
        Call SeteoGrilla
        Call buscofactura
    End If
    
'    Call SeteoGrilla
'
'    Call buscofactura
    

End Sub

Private Sub buscofactura()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPresupuestoC = db.OpenRecordset("PresupuestoC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPresupuestoD = db.OpenRecordset("PresupuestoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
      
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
   '************ Busco Vendedor
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Dim busca3 As String, busca4 As String
    busca3 = RTrim(LTrim(ComboVendedor.text))
    busca4 = busca3 + "z"
    
    rstEmpleado.FindFirst "Nombre >= '" & busca3 & "' and Nombre <= '" & busca4 & "'"
    
'    If rstEmpleado.NoMatch Then
'       MSFlexGrid1.Visible = False
'       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
'
'    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
    
   
    rstCliente.FindFirst "IDCliente= " + Str(codCli)
   
    TextCodigoCliente.text = rstCliente.Fields!IdCliente
    TextApellidoNombre.text = rstCliente.Fields!RazonSocial
    If rstCliente.Fields!CUIT <> "" Then TextCuit.text = rstCliente.Fields!CUIT
    TextDireccion.text = rstCliente.Fields!Domicilio
    TextLocalidad.text = rstCliente.Fields!localidad
    TextCodigoPostal.text = rstCliente.Fields!CP
    TextProvincia.text = rstCliente.Fields!Prov
    TextDescuentoCliente.text = rstCliente.Fields!PorcentajeDescuento
  
    
    Call SeteoGrilla
 
    rstPresupuestoC.FindFirst "NroPresu= " + Str(numDoc)
    
    TextNumeroPresupuesto.text = rstPresupuestoC.Fields!NroPresu
    TextFechaPresupuesto.text = rstPresupuestoC.Fields!FechaPresu
    
    rstPresupuestoD.FindFirst "NroPresu= " + Str(numDoc)
    'rstPresupuestoD.FindFirst "NroPresu= " + Str(Val(numDoc))
    linea2 = 1
    Do While Not rstPresupuestoD.NoMatch
        
            FG1.AddItem " "
            FG1.Row = linea2
       
            FG1.Col = 0
            FG1.text = rstPresupuestoD.Fields!CodProd
            
            FG1.Col = 0
            codigoprod = FG1.text

            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"
       
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            FG1.Col = 1
            FG1.text = rstProductos.Fields!Descripcion
        
            FG1.Col = 2
            FG1.text = rstPresupuestoD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.text = Format(rstPresupuestoD.Fields!precioUnitario, "#,###,###,#0.00")
            FG1.Col = 4
            FG1.text = rstPresupuestoD.Fields!PorcentajeDescuento
            FG1.Col = 5
            FG1.text = rstPresupuestoD.Fields!cantidad
            FG1.Col = 6
            FG1.text = Format(rstPresupuestoD.Fields!totalLinea, "#,###,###,#0.00")
           
       
           rstPresupuestoD.FindNext "NroPresu= " + Str(numDoc)
           linea2 = linea2 + 1
    Loop
    
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstPresupuestoC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.text = rstEmpleado.Fields!Nombre

    '****
    
    TextSubtotalPresupuesto.text = Format(rstPresupuestoC.Fields!SubTotalPresu, "#,###,###,#0.00")
    
    If rstPresupuestoC.Fields!ImporteDesc <> "" Then
        TextDescuentos.text = Format(rstPresupuestoC.Fields!ImporteDesc, "#,###,###,#0.00")
    End If
    
    TextTotalPresupuesto.text = Format(rstPresupuestoC.Fields!TotalPresu, "#,###,###,#0.00")
    
    If rstPresupuestoC.Fields!Anulado = "si" Then
      'A = MsgBox("Pago Anulado", vbOKOnly, "INFO DEL SISTEMA")
       Anulado.Caption = "PRESUPUESTO ANULADO"
      BotonGrabar.Visible = False
   End If
    
End Sub
