Attribute VB_Name = "Vars"
Public BaseSPC As Database
Public tClientes
Public tPaises
Public tProvincias
Public tLocalidades
Public tCondicionIVA
Public tUltimosNumeros
Public tVendedores
Public vFlagBuscar
Public tDepositos
Public tEmpleados



Public Sub LimpiarTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Text = ""
        End If
    Next
    
End Sub

Public Sub DisabledTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = False
        End If
        
        If TypeOf Control Is ComboBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = False
        End If
    Next
End Sub

Public Sub EnabledTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = True
        End If
        
        If TypeOf Control Is ComboBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = True
        End If
    Next
End Sub

