VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormImprimeRemito 
   Caption         =   "Generar Remito"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   17
      Top             =   6720
      Width           =   7695
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4320
         TabIndex        =   18
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGrabar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   2400
         TabIndex        =   4
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame3 
      Height          =   3255
      Left            =   120
      TabIndex        =   16
      Top             =   1440
      Width           =   7665
      Begin VB.TextBox TextItemDomicilio 
         Height          =   285
         Left            =   4080
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   10
         Top             =   2760
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   1800
         Width           =   3135
      End
      Begin VB.TextBox TextApellidoNombre 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   5
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1440
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
         Left            =   120
         TabIndex        =   25
         Top             =   2880
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
         TabIndex        =   24
         Top             =   1920
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
         Left            =   120
         TabIndex        =   23
         Top             =   2400
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
         Left            =   120
         TabIndex        =   22
         Top             =   1560
         Width           =   870
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
         Left            =   120
         TabIndex        =   21
         Top             =   1080
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
         TabIndex        =   20
         Top             =   600
         Width           =   1290
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   15
      Top             =   4680
      Width           =   7695
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1215
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   7455
         _ExtentX        =   13150
         _ExtentY        =   2143
         _Version        =   393216
         Cols            =   6
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox cmbSucursales 
         Height          =   315
         Left            =   4800
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox TextNumeroFactura 
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
         Left            =   1320
         TabIndex        =   2
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TextFechaRemito 
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
         Left            =   4800
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroRemito 
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
         Left            =   1320
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal"
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
         Left            =   3480
         TabIndex        =   27
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         TabIndex        =   26
         Top             =   840
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   3480
         TabIndex        =   14
         Top             =   360
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   13
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormImprimeRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstUltimosNumeros As DAO.Recordset
Dim rstDomiciliosClientes As DAO.Recordset
Dim rstRemitoC As DAO.Recordset
Dim rstRemitoD As DAO.Recordset

Private Sub BotonGrabar_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"

    Set db = DBEngine.OpenDatabase(ruta)
    Set rstRemitoC = db.OpenRecordset("RemitoC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstRemitoD = db.OpenRecordset("RemitoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
    
    vNroRemImp = ""
    
    '******** Grabo Numero Remito en Factura
       
'    Set db1 = DBEngine.OpenDatabase(ruta)
'
'        Set rstRemC = db1.OpenRecordset("RemitoD", dbOpenTable)
'
'        rstRemC.Index = "PrimaryKey"
'
'        rstRemC.Seek "=", Str(TextNumeroFactura.Text)
        
'    rstFacturaC.Index = "PrimaryKey"
'
'    rstFacturaC.Seek "=", Str(TextNumeroFactura.Text)
'
'    If Not rstFacturaC.NoMatch Then
'        A = MsgBox("Factura Existente", vbCritical, "INFO DEL SISTEMA")
'
'
'    Else
'
'
'
'        rstFacturaC.Edit
'        rstFacturaC.Fields!NroRemito = TextNumeroRemito.Text
'        rstFacturaC.Update
'    End If
'

    NumFac = Val(TextNumeroFactura.text)
      
    rstFacturaC.FindFirst "NroFactura= " + Str(NumFac)
    If rstFacturaC.Fields!NroFactura <> Val(TextNumeroFactura.text) Then
        mensaje = MsgBox("Factura Inexistente", vbCritical, "Final de la busqueda")
        'TextCodigoCliente.Text = ""
        'Call blanqueototal
        'TextCodigoCliente.SetFocus
    Else
        rstFacturaC.Edit
        rstFacturaC.Fields!IdSucursal = IdSucursal
        rstFacturaC.Fields!NroRemito = TextNumeroRemito.text
        rstFacturaC.Update
    End If
    
    
    '*******
    
        '*** Busco Remito Existente
       
        Set db1 = DBEngine.OpenDatabase(ruta)
        
        Set rstRemC = db1.OpenRecordset("RemitoD", dbOpenTable)
        
        rstRemC.Index = "PrimaryKey"
        
        rstRemC.Seek "=", Str(TextNumeroFactura.text)

        If Not rstRemC.NoMatch Then
            A = MsgBox("Remito Existente", vbCritical, "INFO DEL SISTEMA")
           
            TextNumeroRemito.text = num
            TextNumeroRemito.SetFocus
        Else
        
        rstRemC.Close
        db1.Close
     
            IdSucursal = Left(cmbSucursales.text, 1)
            rstRemitoC.AddNew
                rstRemitoC.Fields!IdSucursal = CLng(IdSucursal)
                rstRemitoC.Fields!NroRemito = TextNumeroRemito.text
                rstRemitoC.Fields!FechaRemito = TextFechaRemito.text
                rstRemitoC.Fields!item = TextItemDomicilio.text
                rstRemitoC.Fields!CodCliente = TextCodigoCliente.text
                rstRemitoC.Fields!codVendedor = FormFactura.TextLegajoEmpleado.text
                rstRemitoC.Fields!NroFactura = Val(FormFactura.TextNumeroFactura.text)
                rstRemitoC.Fields!TipoFactura = FormFactura.TextTipoFactura.text
                
            rstRemitoC.Update
            
            FormFactura.FG1.Col = 0
            FormFactura.FG1.Row = 1
            Filas = FormFactura.FG1.Rows
            linea = 1
            Do While linea < Filas
                  
                  FormFactura.FG1.Row = linea
                  FormFactura.FG1.Col = 0
                  If FormFactura.FG1.text <> "" Then
                        rstRemitoD.AddNew
                    
                        rstRemitoD.Fields!IdSucursal = CInt(Left(cmbSucursales.text, 1))
                        rstRemitoD.Fields!NroRemito = TextNumeroRemito.text
                        
                    
                        FormFactura.FG1.Col = 0
                        rstRemitoD.Fields!IdCodProd = FormFactura.FG1.text
                    
                        FormFactura.FG1.Col = 2
                        rstRemitoD.Fields!UnidadMedida = FormFactura.FG1.text
                        
                        FormFactura.FG1.Col = 5
                        rstRemitoD.Fields!cantidad = Val(FormFactura.FG1.text)
                        
                        FormFactura.FG1.Col = 8
                        rstRemitoD.Fields!itemremito = Val(FormFactura.FG1.text)
                        
                        rstRemitoD.Update
                  End If
                  linea = linea + 1
            Loop
        
            '*************
              'Guardo en la variable global
                vNroRemImp = TextNumeroRemito.text
            '*************
            
            '*** Actualizo Ultimo Numero Remito
            
            Dim busco As String
       
            'If TextTipoFactura.Text = "A" Then
                busco = "tRemitoC"
            'End If
            
            'If TextTipoFactura.Text = "B" Then
            '    busco = "tFacturaB"
            'End If
    
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
            rstUltimosNumeros.Seek "=", busco, CLng(Left(cmbSucursales.text, 1))
            
            If Not rstUltimosNumeros.NoMatch Then
                ultimo = rstUltimosNumeros.Fields!UltimoNumero
             Else
            End If
            
            If ultimo < Val(TextNumeroRemito.text) Then
                rstUltimosNumeros.Edit
                'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
                     rstUltimosNumeros.Fields!UltimoNumero = TextNumeroRemito.text
                'End If
                rstUltimosNumeros.Update
            End If
        ' End If
            
        End If
        
         Unload FormImprimeRemito
        
        respuesta = MsgBox("Desea Realizar un Pago", vbYesNo, "Pago")
        If respuesta = vbYes Then
            'FormPagoFacturas.Show
            LlamaPagoFactura = True
            FormPagoFacturasDesdeFactura.Show
        Else
           ' If respuesta = vbNo Then Call FormFactura.blanqueototal
            respuesta = MsgBox("Desea Imprimir?", vbYesNo, "Remito")
             
            If respuesta = vbYes Then
                FormImprimir.Show
              Else
               Call FormFactura.SeteoGrilla
               FormFactura.BotonImprimir.Enabled = True
               FormFactura.BotonNueva.Enabled = True
               FormFactura.TextCodigoCliente.SetFocus
            End If

        End If
        

End Sub

Private Sub BotonSalir_Click()

    If FormFactura.TextCodigoCliente <> "" Then
        FormFactura.SeteoGrilla
        FormFactura.BotonImprimir.Enabled = True
        FormFactura.BotonNueva.Enabled = True
        FormFactura.TextCodigoCliente.SetFocus
    End If
    
    Unload FormImprimeRemito

End Sub

Private Sub cmbSucursales_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub Form_Load()

    Dim tSucursales
    
    FormImprimeRemito.Height = 8625
    FormImprimeRemito.Width = 8055
    FormImprimeRemito.Top = 1000
    FormImprimeRemito.Left = 12300

    'Call titulos
    
    Dim NumeroRemito As Long
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    'Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenTable)
    
    rstUltimosNumeros.Index = "PrimaryKey"
    
    '** Cargamos el combo de sucursales, modificacion agregada con los nuevos remitos 2025-06 ***************
        Set tSucursales = db.OpenRecordset("Sucursales", dbOpenTable)
        
        tSucursales.MoveFirst
        
        While Not tSucursales.EOF
            
            cmbSucursales.AddItem tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal
            tSucursales.MoveNext
        
        Wend
        
        cmbSucursales.ListIndex = 1
    '*****************************************************************************************************
    
    Dim busco As String
     
    busco = "tRemitoC"
    
    
  
    
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    rstUltimosNumeros.Seek "=", busco, CLng(Left(cmbSucursales.text, 1))
    
    If Not rstUltimosNumeros.NoMatch Then
        NumeroRemito = rstUltimosNumeros.Fields!UltimoNumero
    End If
    
    'If rstUltimosNumeros.NoMatch Then
    '   FG1.Visible = False
    '   mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    'End If
    
    TextNumeroRemito.text = NumeroRemito + 1

    TextFechaRemito.text = Format(Date, "dd/mm/yyyy")
    
    TextNumeroFactura.text = FormFactura.TextNumeroFactura.text
    TextCodigoCliente.text = FormFactura.TextCodigoCliente.text
    TextApellidoNombre.text = FormFactura.TextApellidoNombre.text
    
    If TextCodigoCliente.text <> "" Then
        
        MSHFlexGrid1.Col = 1
        TextDireccion.text = MSHFlexGrid1.text
        
        MSHFlexGrid1.Col = 2
        TextLocalidad.text = MSHFlexGrid1.text
        
        MSHFlexGrid1.Col = 3
        TextCodigoPostal = MSHFlexGrid1.text
        
        MSHFlexGrid1.Col = 4
        TextProvincia.text = MSHFlexGrid1.text
        
        MSHFlexGrid1.Col = 5
        TextItemDomicilio.text = MSHFlexGrid1.text
        
        BotonGrabar.Enabled = True
    
    End If


End Sub

Private Sub titulos()

    MSHFlexGrid1.Row = 0
    
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.text = "Item"
    MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSHFlexGrid1.ColWidth(0) = 0
    
        
    MSHFlexGrid1.Col = 1
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.text = "Domicilio"
    MSHFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
    MSHFlexGrid1.ColWidth(1) = 4000
    
    MSHFlexGrid1.Col = 2
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.text = "Localidad"
    MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    MSHFlexGrid1.ColWidth(2) = 3000
    
    MSHFlexGrid1.Col = 3
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.text = "Cod.Pos"
    MSHFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    MSHFlexGrid1.ColWidth(3) = 0
    
    MSHFlexGrid1.Col = 4
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.text = "Provincia"
    MSHFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    MSHFlexGrid1.ColWidth(4) = 0
    
    MSHFlexGrid1.Col = 5
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(5) = flexAlignCenterCenter
    MSHFlexGrid1.text = "Item"
    MSHFlexGrid1.ColWidth(5) = 0
 
 End Sub
 
 Private Sub buscodirecciones()
 
 MSHFlexGrid1.Clear
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstDomiciliosClientes = db.OpenRecordset("DomiciliosClientes", dbOpenDynaset)
    
    
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Visible = True
    
    Call titulos
    
    
    CodigoClie = Val(TextCodigoCliente.text)
    
    rstDomiciliosClientes.FindFirst "IDCliente= " + Str(CodigoClie)
    'facturacancelada = rstDomiciliosClientes.Fields!Cancelada
    codigoclientedetalle = rstDomiciliosClientes.Fields!IdCliente
    
    If rstDomiciliosClientes.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
            'MSHFlexGrid1.Visible = False
        
            'MSHFlexGrid1.AddItem " "
            'MSHFlexGrid1.Row = linea2
       
            'MSHFlexGrid1.Col = 0
            'MSHFlexGrid1.Text = 1
            'MSHFlexGrid1.Col = 1
            TextDireccion.text = FormFactura.TextDireccion.text
            'MSHFlexGrid1.Text =
            'MSHFlexGrid1.Col = 2
            TextLocalidad.text = FormFactura.TextLocalidad.text
            'MSHFlexGrid1.Text = FormFactura.TextLocalidad.Text
            'MSHFlexGrid1.Col = 3
            TextCodigoPostal.text = FormFactura.TextCodigoPostal.text
            'MSHFlexGrid1.Text = FormFactura.TextCodigoPostal.Text
            'MSHFlexGrid1.Col = 4
            TextProvincia.text = FormFactura.TextProvincia.text
            TextItemDomicilio.text = 0
            'MSHFlexGrid1.Text = FormFactura.TextProvincia.Text
            'MSHFlexGrid1.Col = 5
            'MSHFlexGrid1.Text = 0
            'facturacancelada = rstDomiciliosClientes.Fields!Cancelada
            'If facturacancelada = True Then
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "SI"
            'Else
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "NO"
            'End If
            'linea2 = linea2 + 1
                
            'rstDomiciliosClientes.FindNext "IDCliente= " + Str(CodigoClie)
        
        'mensaje = MsgBox("No Existen Domicilios", vbCritical, "Final de la busqueda")
        'TextCodigoCliente.Text = ""
        'Call blanco
        'TextCodigoCliente.SetFocus
        BotonGrabar.Enabled = True
        Exit Sub
    End If
    
    If codigoclientedetalle = CodigoClie Then
        MSHFlexGrid1.Rows = 2
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Visible = True
    
       
    Else
        MSHFlexGrid1.Visible = False
    End If
    Call titulos
    linea2 = 1
    Do While Not rstDomiciliosClientes.NoMatch
            MSHFlexGrid1.AddItem " "
            MSHFlexGrid1.Row = linea2
       
            MSHFlexGrid1.Col = 0
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!item
            MSHFlexGrid1.Col = 1
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!Domicilio
            MSHFlexGrid1.Col = 2
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!localidad
            MSHFlexGrid1.Col = 3
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!CP
            MSHFlexGrid1.Col = 4
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!Prov
            MSHFlexGrid1.Col = 5
            MSHFlexGrid1.text = rstDomiciliosClientes.Fields!item
            'facturacancelada = rstDomiciliosClientes.Fields!Cancelada
            'If facturacancelada = True Then
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "SI"
            'Else
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "NO"
            'End If
            linea2 = linea2 + 1
                
            rstDomiciliosClientes.FindNext "IDCliente= " + Str(CodigoClie)
    Loop
    
    

 
 End Sub

Private Sub MSHFlexGrid1_Click()

    MSHFlexGrid1.Col = 1
    TextDireccion.text = MSHFlexGrid1.text
    
    MSHFlexGrid1.Col = 2
    TextLocalidad.text = MSHFlexGrid1.text
    
    MSHFlexGrid1.Col = 3
    TextCodigoPostal = MSHFlexGrid1.text
    
    MSHFlexGrid1.Col = 4
    TextProvincia.text = MSHFlexGrid1.text
    
    MSHFlexGrid1.Col = 5
    TextItemDomicilio.text = MSHFlexGrid1.text
    
    
    BotonGrabar.Enabled = True

End Sub

Private Sub TextCodigoCliente_Change()

    Call buscodirecciones
    
End Sub





Private Sub TextFechaRemito_GotFocus()
    TextFechaRemito.SelLength = Len(TextFechaRemito.text)
End Sub

Private Sub TextFechaRemito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextNumeroRemito_GotFocus()
    TextNumeroRemito.SelLength = Len(TextNumeroRemito.text)
End Sub

Private Sub TextNumeroRemito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub
