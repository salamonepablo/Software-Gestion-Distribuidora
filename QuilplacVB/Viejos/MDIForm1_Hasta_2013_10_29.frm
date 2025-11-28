VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "Pago a Proveedores"
   ClientHeight    =   3030
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu Provee 
      Caption         =   "&Proveedores"
      Begin VB.Menu ABM 
         Caption         =   "&ABM Proveedores"
      End
      Begin VB.Menu Consul 
         Caption         =   "&Consultas"
         Enabled         =   0   'False
      End
      Begin VB.Menu Listados 
         Caption         =   "&Listados"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu Rete 
      Caption         =   "&Retencion"
      Begin VB.Menu Compro 
         Caption         =   "Comprobante &Retencion"
      End
      Begin VB.Menu Consu 
         Caption         =   "&Consultas"
         Begin VB.Menu ConProveedor 
            Caption         =   "Consulta &Por Proveedor"
         End
         Begin VB.Menu connumpago 
            Caption         =   "Consulta por &Numero de Pago"
         End
         Begin VB.Menu Confecha 
            Caption         =   "Consulta entre &Fechas"
         End
      End
      Begin VB.Menu Elimino 
         Caption         =   "Eliminar Pago"
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ABM_Click()
    FormProveedores.Show
End Sub

Private Sub Compro_Click()
    FormComprobanteIIBB.Show
End Sub

Private Sub ConProveedor_Click()
    FormConsultaCodigoProveedor.Show
End Sub

Private Sub connumpago_Click()
    FormConsultaNumeroPago.Show
End Sub

Private Sub Elimino_Click()
    FormEliminarPago.Show
End Sub

