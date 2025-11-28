VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form FormPreviewProductos 
   Caption         =   "Form1"
   ClientHeight    =   8910
   ClientLeft      =   750
   ClientTop       =   525
   ClientWidth     =   9840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   9840
   Begin CRVIEWERLibCtl.CRViewer CRV 
      Height          =   8580
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12075
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   0   'False
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
   Begin VB.Label lblReporte 
      Caption         =   "Label1"
      Height          =   795
      Left            =   12600
      TabIndex        =   1
      Top             =   2505
      Width           =   840
   End
End
Attribute VB_Name = "FormPreviewProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub

Private Sub Form_Activate()
cambiarTamaño
mostrarReporte lblReporte.Caption
End Sub

Private Sub Form_Load()
cambiarTamaño
Screen.MousePointer = 0
End Sub
Sub mostrarReporte(reporte As String)
Dim RepApp As New CRAXDRT.Application
Dim repRep As CRAXDRT.Report
Me.WindowState = 2
Me.Caption = "Imprimir " & reporte
Set repRep = RepApp.OpenReport(App.Path & "\ListadoProductos.rpt")

CRV.ReportSource = repRep
CRV.EnableExportButton = True

CRV.ViewReport
End Sub
Private Sub Form_Resize()
cambiarTamaño
End Sub
Sub cambiarTamaño()
CRV.Left = 0
CRV.Top = 0
CRV.Height = FormPreviewProductos.Height - 120
CRV.Width = FormPreviewProductos.Width - 120
End Sub
