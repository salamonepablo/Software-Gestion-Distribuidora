Attribute VB_Name = "Sentencias"
Option Explicit

Sub main()
With BASE
    .CursorLocation = adUseClient
    .Open " Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\DB_SPC_SI.mdb;Persist Security Info=False"
    FormLogin.Show
End With
End Sub

Sub USUARIOS()
With RsUsuarios
If .State = 1 Then .Close
.Open "select * from usuarios ", BASE, adOpenStatic, adLockOptimistic

End With
End Sub
