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

Private Function ResolveDeliveryAddressItem() As Long
    Dim rawItem As String
    Dim numericItem As Double
    On Error GoTo UseDefault
    rawItem = Trim$(TextItemDomicilio.text)
    If Len(rawItem) = 0 Then GoTo UseDefault
    If Not IsNumeric(rawItem) Then GoTo UseDefault
    numericItem = CDbl(rawItem)
    If numericItem < 0 Or numericItem > 2147483647# Then GoTo UseDefault
    If numericItem <> Fix(numericItem) Then GoTo UseDefault
    ResolveDeliveryAddressItem = CLng(numericItem)
    Exit Function
UseDefault:
    SetDefaultDeliveryAddress
    LogDeliveryAddressFallback
    MsgBox "No se selecciono un domicilio adicional valido. Se usara el domicilio principal de la factura (identificador 0).", vbInformation, "Domicilio del remito"
    ResolveDeliveryAddressItem = 0
End Function

Private Sub SetDefaultDeliveryAddress()
    TextItemDomicilio.text = "0"
    TextDireccion.text = FormFactura.TextDireccion.text
    TextLocalidad.text = FormFactura.TextLocalidad.text
    TextCodigoPostal.text = FormFactura.TextCodigoPostal.text
    TextProvincia.text = FormFactura.TextProvincia.text
End Sub

Private Sub LogDeliveryAddressFallback()
    Dim logFile As Integer
    On Error GoTo LogUnavailable
    logFile = FreeFile
    Open App.Path & "\delivery-address-fallback.log" For Append As #logFile
    Print #logFile, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " client=" & TextCodigoCliente.text & " invoice=" & TextNumeroFactura.text & " delivery=" & TextNumeroRemito.text & " defaultItem=0"
    Close #logFile
    Exit Sub
LogUnavailable:
    On Error Resume Next
    If logFile > 0 Then Close #logFile
    MsgBox "Se aplico el domicilio principal, pero no se pudo escribir el registro de diagnostico.", vbExclamation
End Sub

Private Sub SelectDeliveryAddressRow()
    If MSHFlexGrid1.Row < 1 Then Exit Sub
    If Len(Trim$(MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5))) = 0 Then Exit Sub
    TextDireccion.text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 1)
    TextLocalidad.text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 2)
    TextCodigoPostal.text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 3)
    TextProvincia.text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 4)
    TextItemDomicilio.text = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Row, 5)
End Sub

Private Function TryReuseExistingDelivery() As Integer
    Dim checkDatabase As DAO.Database
    Dim existingDelivery As DAO.Recordset
    Dim branchNumber As Long
    Dim deliveryNumber As Long
    Dim sourceType As String
    Dim sourceNumber As Long
    Dim creditInput As String
    Dim creditNumber As Long
    Dim failureReason As String
    On Error GoTo Failed
    branchNumber = CLng(Val(cmbSucursales.text))
    deliveryNumber = CLng(TextNumeroRemito.text)
    If branchNumber <= 0 Or deliveryNumber <= 0 Then Err.Raise vbObjectError + 2250, , "Seleccione sucursal y numero de remito validos."
    Set checkDatabase = DBEngine.OpenDatabase(App.Path & "\DB_SPC_SI.mdb", False, True)
    Set existingDelivery = checkDatabase.OpenRecordset("SELECT * FROM RemitoC WHERE IdSucursal=" & branchNumber & " AND NroRemito=" & deliveryNumber, dbOpenSnapshot)
    If existingDelivery.EOF Then
        existingDelivery.Close
        checkDatabase.Close
        Exit Function
    End If
    TryReuseExistingDelivery = -1
    If Val(existingDelivery!CodCliente & "") <> CLng(TextCodigoCliente.text) Then Err.Raise vbObjectError + 2251, , "El remito pertenece a otro cliente."
    sourceType = Trim$(existingDelivery!TipoFactura & "")
    sourceNumber = Val(existingDelivery!NroFactura & "")
    existingDelivery.Close
    checkDatabase.Close
    If sourceType = FormFactura.TextTipoFactura.text And sourceNumber = CLng(FormFactura.TextNumeroFactura.text) Then
        MsgBox "El remito ya esta asociado a esta factura.", vbInformation
        TryReuseExistingDelivery = 1
        Exit Function
    End If
    creditInput = Trim$(InputBox("El remito esta asociado a la factura " & sourceType & " " & sourceNumber & ". Ingrese el numero de la NOTA DE CREDITO TOTAL de esa factura (misma letra). Cancelar conserva la asociacion actual.", "Reutilizar remito existente"))
    If Len(creditInput) = 0 Then Exit Function
    If Not IsNumeric(creditInput) Then Err.Raise vbObjectError + 2252, , "Numero de nota de credito no valido."
    If CDbl(creditInput) <> Fix(CDbl(creditInput)) Or CDbl(creditInput) <= 0 Then Err.Raise vbObjectError + 2252, , "Numero de nota de credito no valido."
    creditNumber = CLng(creditInput)
    If MsgBox("Confirma que la nota de credito " & sourceType & " " & creditNumber & " fue emitida para compensar totalmente la factura " & sourceType & " " & sourceNumber & "?" & vbCrLf & "La relacion historica no esta guardada en el sistema: verifique el comprobante. El remito " & branchNumber & "-" & deliveryNumber & " pasara a la factura " & FormFactura.TextTipoFactura.text & " " & FormFactura.TextNumeroFactura.text & ", conservando su detalle e historial.", vbYesNo Or vbExclamation Or vbDefaultButton2, "Confirmar reasociacion") <> vbYes Then Exit Function
    If ReassignDelivery(App.Path & "\DB_SPC_SI.mdb", branchNumber, deliveryNumber, FormFactura.TextTipoFactura.text, CLng(FormFactura.TextNumeroFactura.text), CLng(TextCodigoCliente.text), sourceType, creditNumber, failureReason) Then
        TryReuseExistingDelivery = 1
    Else
        MsgBox "No se pudo reasociar el remito: " & failureReason, vbExclamation
    End If
    Exit Function
Failed:
    TryReuseExistingDelivery = -1
    MsgBox "No se pudo verificar el remito: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not existingDelivery Is Nothing Then existingDelivery.Close
    If Not checkDatabase Is Nothing Then checkDatabase.Close
End Function

Private Sub BotonGrabar_Click()
    Dim addressItem As Long
    Dim reuseResult As Integer
    reuseResult = TryReuseExistingDelivery()
    If reuseResult < 0 Then Exit Sub
    If reuseResult = 1 Then
        vNroRemImp = TextNumeroRemito.text
        GoTo DeliverySaved
    End If
    IdSucursal = CLng(Val(cmbSucursales.text))
    addressItem = ResolveDeliveryAddressItem()

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
'        Set rstRemC = db1.OpenRecordset("RemitoC", dbOpenTable)
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
      
    rstFacturaC.FindFirst "NroFactura=" & CStr(NumFac) & " AND TipoFactura='" & FormFactura.TextTipoFactura.text & "'"
    If rstFacturaC.NoMatch Then
        mensaje = MsgBox("Factura Inexistente", vbCritical, "Final de la busqueda")
        Exit Sub
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
        
        Set rstRemC = db1.OpenRecordset("RemitoC", dbOpenTable)
        
        rstRemC.Index = "PrimaryKey"
        
        rstRemC.Seek "=", CLng(Val(cmbSucursales.text)), CLng(TextNumeroRemito.text)

        If Not rstRemC.NoMatch Then
            A = MsgBox("Remito Existente", vbCritical, "INFO DEL SISTEMA")
           
            TextNumeroRemito.text = num
            TextNumeroRemito.SetFocus
            rstRemC.Close
            db1.Close
            Exit Sub
        Else
        
        rstRemC.Close
        db1.Close
     
            IdSucursal = CLng(Val(cmbSucursales.text))
            rstRemitoC.AddNew
                rstRemitoC.Fields!IdSucursal = CLng(IdSucursal)
                rstRemitoC.Fields!NroRemito = TextNumeroRemito.text
                rstRemitoC.Fields!FechaRemito = TextFechaRemito.text
                rstRemitoC.Fields!item = addressItem
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
                    
                        rstRemitoD.Fields!IdSucursal = CLng(Val(cmbSucursales.text))
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
            rstUltimosNumeros.Seek "=", busco, CLng(Val(cmbSucursales.text))
            
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
        
DeliverySaved:
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
    rstUltimosNumeros.Seek "=", busco, CLng(Val(cmbSucursales.text))
    
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
        SelectDeliveryAddressRow
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
    Dim addressRows As DAO.Recordset
    Dim addressDatabase As DAO.Database
    Dim rowIndex As Long
    On Error GoTo AddressLoadFailed
    SetDefaultDeliveryAddress
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Clear
    titulos
    MSHFlexGrid1.Visible = False
    If Len(Trim$(TextCodigoCliente.text)) = 0 Then Exit Sub
    Set addressDatabase = DBEngine.OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set addressRows = addressDatabase.OpenRecordset("SELECT * FROM DomiciliosClientes WHERE IdCliente=" & CStr(Val(TextCodigoCliente.text)) & " ORDER BY item", dbOpenSnapshot)
    rowIndex = 1
    Do While Not addressRows.EOF
        MSHFlexGrid1.Rows = rowIndex + 1
        MSHFlexGrid1.TextMatrix(rowIndex, 0) = addressRows!item & ""
        MSHFlexGrid1.TextMatrix(rowIndex, 1) = addressRows!Domicilio & ""
        MSHFlexGrid1.TextMatrix(rowIndex, 2) = addressRows!localidad & ""
        MSHFlexGrid1.TextMatrix(rowIndex, 3) = addressRows!CP & ""
        MSHFlexGrid1.TextMatrix(rowIndex, 4) = addressRows!Prov & ""
        MSHFlexGrid1.TextMatrix(rowIndex, 5) = addressRows!item & ""
        rowIndex = rowIndex + 1
        addressRows.MoveNext
    Loop
    addressRows.Close
    addressDatabase.Close
    If rowIndex > 1 Then
        MSHFlexGrid1.Visible = True
        MSHFlexGrid1.Row = 1
        SelectDeliveryAddressRow
    End If
    BotonGrabar.Enabled = True
    Exit Sub
AddressLoadFailed:
    BotonGrabar.Enabled = False
    MsgBox "No se pudieron cargar los domicilios: " & Err.Description, vbExclamation
    On Error Resume Next
    If Not addressRows Is Nothing Then addressRows.Close
    If Not addressDatabase Is Nothing Then addressDatabase.Close
End Sub

Private Sub MSHFlexGrid1_Click()
    SelectDeliveryAddressRow
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
