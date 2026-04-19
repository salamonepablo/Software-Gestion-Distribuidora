Private Sub DibujarFormulario(ByVal Leyenda As String)
    Dim i As Integer
    Dim RowY As Single

    With Printer
        .ForeColor = vbBlack
        Printer.DrawWidth = 15

        ' ===== BORDE EXTERIOR =====
        Printer.Line (25, 10)-(200, 275), , B

        ' ===== SECCION ENCABEZADO =====
        Printer.DrawWidth = 2
        Printer.Line (25, 72)-(200, 72)

        'Recuadro Presupuesto (arriba derecha)
        Printer.Line (108, 12)-(198, 35), , B

        '"Documento no valido como Factura"
        .Font = "Arial"
        .FontSize = 9
        .FontBold = False
        .CurrentX = 120
        .CurrentY = 15
        Printer.Print "Documento no v" & Chr(225) & "lido como Factura"

        'Casilla [X]
        Printer.Line (112, 25)-(119, 32), , B
        .Font = "Arial"
        .FontSize = 11
        .FontBold = True
        .CurrentX = 114
        .CurrentY = 26
        Printer.Print "X"

        '"PRESUPUESTO N°"
        .Font = "Arial"
        .FontSize = 11
        .FontBold = True
        .CurrentX = 121
        .CurrentY = 26
        Printer.Print "PRESUPUESTO N" & Chr(186)

        'Leyenda de copia: ORIGINAL / DUPLICADO / TRIPLICADO
        .Font = "Arial"
        .FontSize = 14
        .FontBold = True
        .ForeColor = vbRed
        .CurrentX = 30
        .CurrentY = 15
        Printer.Print Leyenda
        .ForeColor = vbBlack

        '"Fecha:"
        .Font = "Arial"
        .FontSize = 10
        .FontBold = True
        .CurrentX = 120
        .CurrentY = 45
        Printer.Print "Fecha: ........................................"

        '"Señor/es:"
        .Font = "Arial"
        .FontSize = 10
        .FontBold = False
        .CurrentX = 27
        .CurrentY = 53
        Printer.Print "Se" & Chr(241) & "or/es: ..............................................................................................................."

        '"Domicilio:" y "Localidad:"
        .CurrentX = 27
        .CurrentY = 61
        Printer.Print "Domicilio: ......................................................  Localidad: .................................."

        'Linea punteada separadora
        .DrawStyle = 2  'Dot
        Printer.Line (27, 69)-(198, 69)
        .DrawStyle = 0  'Solid

        ' ===== TABLA DE DETALLE =====
        'Fila de encabezados
        .DrawWidth = 10
        Printer.Line (25, 72)-(200, 80), , B

        .Font = "Arial"
        .FontSize = 10
        .FontBold = True

        .CurrentX = 30
        .CurrentY = 74
        Printer.Print "CANTIDAD"

        .CurrentX = 82
        .CurrentY = 74
        Printer.Print "DETALLE"

        .CurrentX = 130
        .CurrentY = 74
        Printer.Print "PRECIO Unit."

        .CurrentX = 172
        .CurrentY = 74
        Printer.Print "IMPORTE"

        'Lineas verticales de columnas
        Printer.DrawWidth = 2
        Printer.Line (55, 72)-(55, 260)      'CANTIDAD | DETALLE
        Printer.Line (125, 72)-(125, 260)    'DETALLE  | PRECIO
        Printer.Line (163, 72)-(163, 260)    'PRECIO   | IMPORTE

        'Lineas horizontales de filas (22 filas de 8mm) dot
        Printer.DrawWidth = 1
        Printer.DrawStyle = 2 'dot
        For i = 1 To 22
            RowY = 80 + (i * 8)
            Printer.Line (25, RowY)-(200, RowY)
        Next i

        ' ===== FILA TOTAL =====
        Printer.DrawWidth = 10
        Printer.Line (25, 260)-(200, 275), , B
        Printer.DrawWidth = 2
        Printer.Line (163, 260)-(163, 275)  'separador monto

        .Font = "Arial"
        .FontSize = 10
        .FontBold = True
        .CurrentX = 125
        .CurrentY = 264
        Printer.Print "TOTAL $"

    End With
End Sub


Public Sub Imprimir()
    Dim PU As Variant, TL As Variant, Cant As Variant, TotalPres As Variant
    Dim renglon As Single
    Dim Copia As Integer
    Dim LeyendaCopia As String
    Dim Hasta As Integer, i As Integer

    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    vSQLPC = "SELECT * FROM PresupuestoC WHERE NroPresu=" & TextNumeroPresupuesto.text & " ORDER By NroPresu"
    vSQLPD = "SELECT * FROM PresupuestoD WHERE NroPresu=" & TextNumeroPresupuesto.text & " ORDER By NroPresu"

    With Printer
        .ScaleMode = 6  'mm

        For Copia = 1 To 3
            Select Case Copia
                Case 1: LeyendaCopia = "ORIGINAL"
                Case 2: LeyendaCopia = "DUPLICADO"
                Case 3: LeyendaCopia = "TRIPLICADO"
            End Select

            '=== DIBUJAR FORMULARIO ===
            DibujarFormulario LeyendaCopia

            '=== IMPRIMIR DATOS ===

            'Numero de Presupuesto (dentro del recuadro)
            .CurrentX = 165
            .CurrentY = 26
            .Font = "Courier New"
            .FontSize = 12
            .FontBold = True
            .ForeColor = vbBlue
            Printer.Print TextNumeroPresupuesto.text

            'Codigo de Cliente (arriba izquierda, en rojo)
            .CurrentX = 30
            .CurrentY = 25
            .Font = "Courier New"
            .FontSize = 14
            .FontBold = True
            .ForeColor = vbRed
            Printer.Print TextCodigoCliente.text
            .ForeColor = vbBlack

            'Fecha
            .CurrentX = 143
            .CurrentY = 44
            .Font = "Courier New"
            .FontSize = 11
            .FontBold = False
            Printer.Print Format(TextFechaPresupuesto.text, "DD/MM/YYYY")

            'Apellido y Nombre
            .CurrentX = 57
            .CurrentY = 52
            .Font = "Courier New"
            .FontSize = 11
            .FontBold = True
            Printer.Print TextApellidoNombre.text

            'Direccion
            .CurrentX = 57
            .CurrentY = 60
            .Font = "Courier New"
            .FontSize = 11
            .FontBold = False
            Printer.Print TextDireccion.text

            'Localidad
            .CurrentX = 110
            .CurrentY = 60
            .Font = "Courier New"
            .FontSize = 11
            .FontBold = False
            Printer.Print TextLocalidad.text

            '=== DETALLE ===
            Set PresuC = BaseSPC.OpenRecordset(vSQLPC, dbOpenDynaset)
            Set PresuD = BaseSPC.OpenRecordset(vSQLPD, dbOpenDynaset)

            PresuC.MoveFirst
            PresuD.MoveFirst
            renglon = 0

            While Not PresuD.EOF
                'Cantidad
                .CurrentX = 29
                .CurrentY = 82 + renglon
                .Font = "Courier New"
                .FontSize = 10
                .FontBold = False

                Cant = CDbl(PresuD!cantidad)
                Cant = Format(Cant, "Standard")
                Hasta = CInt(6 - Len(Cant))
                For i = 0 To Hasta
                    Cant = " " & Cant
                Next i
                Printer.Print Cant

                'Descripcion
                .CurrentX = 57
                .CurrentY = 82 + renglon
                .Font = "Courier New"
                .FontSize = 10
                .FontBold = False
                Printer.Print BuscarDescProd(PresuD!CodProd)

                'Precio Unitario (con descuento aplicado)
                .CurrentX = 129
                .CurrentY = 82 + renglon
                .Font = "Courier New"
                .FontSize = 10
                .FontBold = False
                PU = CDbl(PresuD!precioUnitario) - (CDbl(PresuD!precioUnitario) * CDbl(PresuD!PorcentajeDescuento) / 100)
                PU = Format(PU, "Standard")
                Hasta = CInt(10 - Len(PU))
                For i = 0 To Hasta
                    PU = " " & PU
                Next i
                Printer.Print PU

                'Importe
                .CurrentX = 165
                .CurrentY = 82 + renglon
                .Font = "Courier New"
                .FontSize = 10
                .FontBold = False
                TL = Format(PresuD!totalLinea, "Standard")
                Hasta = CInt(14 - Len(TL))
                For i = 0 To Hasta
                    TL = " " & TL
                Next i
                Printer.Print TL

                renglon = renglon + 8  '8mm por fila
                PresuD.MoveNext
            Wend

            'Total
            .CurrentX = 163
            .CurrentY = 263
            .Font = "Courier New"
            .FontSize = 11
            .FontBold = True
            TotalPres = Format(CDbl(PresuC!TotalPresu), "Standard")
            Hasta = CInt(14 - Len(TotalPres))
            For i = 0 To Hasta
                TotalPres = " " & TotalPres
            Next i
            Printer.Print TotalPres

            PresuC.Close
            PresuD.Close

            'Nueva pagina si no es la ultima copia
            If Copia < 3 Then
                Printer.NewPage
            End If
        Next Copia

        .EndDoc
    End With

    BaseSPC.Close

    Call blanqueototal
    MSFlexGrid1.Visible = False
    TextCodigoCliente.SetFocus
    Unload FormPagoFacturasDesdeFactura

End Sub