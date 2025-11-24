Imports FEAFIPLib

Public Class Form1


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim lwsPadron As wsPadron = New wsPadron

        Dim contribuyente As Contribuyente = New Contribuyente
        lwsPadron.CUIT = 20939802593
        ' Para poder consultar un CUIT real debe habilitar el modo producción. En modo de homologación la consulta no encuentra algunos CUITs
        lwsPadron.ModoProduccion = False
        If lwsPadron.login("certificado.crt", "clave.key") Then
            If lwsPadron.consultar(30610171601, contribuyente) Then
                Dim nombre As String = contribuyente.nombre
                Dim tipoPersona As String = contribuyente.tipoPersona
                Dim domicilioFiscal As Domicilio = contribuyente.domicilioFiscal
                Dim direccion As String = domicilioFiscal.direccion + ", " + domicilioFiscal.localidad + ", " + domicilioFiscal.provincia
                ' Solicito al cliente constancia porque no esta inscripto en ganancias
                Dim solicitarConstancia As Boolean = contribuyente.SolicitarConstanciaInscripcion
                Dim condicionIVA As String = contribuyente.condicionIVADesc
            Else
                MsgBox(lwsPadron.ErrorDesc)
            End If
        Else
            MsgBox(lwsPadron.ErrorDesc)
        End If
    End Sub
End Class
