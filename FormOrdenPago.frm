VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormOrdenPago 
   Caption         =   "Orden de Pago"
   ClientHeight    =   10620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   ScaleHeight     =   10620
   ScaleWidth      =   16470
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   10455
      Left            =   120
      TabIndex        =   30
      Top             =   0
      Width           =   16215
      Begin VB.Frame Frame6 
         Caption         =   "Reimpresi?n"
         Height          =   2415
         Left            =   8400
         TabIndex        =   58
         Top             =   6240
         Width           =   7455
         Begin VB.ComboBox cboBuscarOrden 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   720
            Width           =   2415
         End
         Begin VB.CommandButton cmdCargarOrden 
            Caption         =   "Cargar Orden"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   600
            TabIndex        =   59
            Top             =   1200
            Width           =   1815
         End
         Begin VB.Label lblReImprimir 
            AutoSize        =   -1  'True
            Caption         =   "Orden de Pago"
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
            TabIndex        =   60
            Top             =   360
            Width           =   1290
         End
         Begin VB.Image Image1 
            Height          =   810
            Left            =   4080
            Picture         =   "FormOrdenPago.frx":0000
            Stretch         =   -1  'True
            Top             =   840
            Width           =   2775
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Detalle de Transferencia"
         Height          =   2415
         Left            =   8400
         TabIndex        =   54
         Top             =   3720
         Width           =   7455
         Begin VB.CommandButton cmdAddTransferencia 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   16
            Top             =   480
            Width           =   855
         End
         Begin VB.CommandButton cmdDelTransferencia 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   17
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   56
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   55
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtSubTransferencia 
            Height          =   285
            Left            =   1440
            TabIndex        =   18
            Top             =   1800
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdTransferencia 
            Height          =   1455
            Left            =   480
            TabIndex        =   15
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.Label lblSubTransferencia 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
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
            Left            =   480
            TabIndex        =   57
            Top             =   1800
            Width           =   900
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Detalle de Facturas"
         Height          =   2415
         Left            =   8400
         TabIndex        =   50
         Top             =   1080
         Width           =   7455
         Begin VB.CommandButton cmdAddFacturas 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.CommandButton cmdDelFacturas 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   9
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   52
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   51
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtSubFacturas 
            Height          =   285
            Left            =   1440
            TabIndex        =   10
            Top             =   1800
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdFacturas 
            Height          =   1455
            Left            =   480
            TabIndex        =   7
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.Label lblSubFacturas 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
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
            Left            =   480
            TabIndex        =   53
            Top             =   1800
            Width           =   900
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Totales"
         Height          =   1335
         Left            =   600
         TabIndex        =   43
         Top             =   8760
         Width           =   15375
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   12720
            TabIndex        =   49
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdGuardar 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9240
            TabIndex        =   27
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdRestaurarAuto 
            Caption         =   "R&estaurar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4800
            TabIndex        =   28
            Top             =   840
            Width           =   1455
         End
         Begin VB.CommandButton cmdRecalcular 
            Caption         =   "&Recalcular"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            TabIndex        =   48
            Top             =   840
            Width           =   1455
         End
         Begin VB.TextBox txtImporteLetras 
            Height          =   375
            Left            =   6960
            TabIndex        =   26
            Top             =   360
            Width           =   7935
         End
         Begin VB.TextBox txtSaldo 
            Height          =   375
            Left            =   4800
            TabIndex        =   25
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtTotalPago 
            Height          =   375
            Left            =   2880
            TabIndex        =   24
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtEfectivo 
            Height          =   375
            Left            =   960
            TabIndex        =   23
            Top             =   360
            Width           =   1695
         End
         Begin VB.Label lblImporteLetras 
            AutoSize        =   -1  'True
            Caption         =   "Importe en Letras"
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
            TabIndex        =   47
            Top             =   120
            Width           =   1500
         End
         Begin VB.Label lblSaldo 
            AutoSize        =   -1  'True
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
            Left            =   4920
            TabIndex        =   46
            Top             =   120
            Width           =   495
         End
         Begin VB.Label lblTotalPago 
            AutoSize        =   -1  'True
            Caption         =   "Total Pago"
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
            Left            =   3000
            TabIndex        =   45
            Top             =   120
            Width           =   945
         End
         Begin VB.Label lblEfectivo 
            AutoSize        =   -1  'True
            Caption         =   "Efectivo"
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
            Left            =   1080
            TabIndex        =   44
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame fraOtros 
         Caption         =   "Detalle de Otros"
         Height          =   2415
         Left            =   600
         TabIndex        =   39
         Top             =   6240
         Width           =   7455
         Begin VB.CommandButton cmdDelOtros 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   21
            Top             =   960
            Width           =   855
         End
         Begin VB.CommandButton cmdAddOtros 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   20
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSubOtros 
            Height          =   285
            Left            =   1440
            TabIndex        =   22
            Top             =   1800
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelOtro 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   41
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddOtro 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   13080
            TabIndex        =   40
            Top             =   600
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid grdOtros 
            Height          =   1455
            Left            =   480
            TabIndex        =   19
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2566
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.Label lblSubOtros 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
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
            Left            =   480
            TabIndex        =   42
            Top             =   1800
            Width           =   900
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle de Cheques"
         Height          =   2535
         Left            =   600
         TabIndex        =   37
         Top             =   3600
         Width           =   7455
         Begin VB.CommandButton cmdAddCheque 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   12
            Top             =   600
            Width           =   855
         End
         Begin VB.CommandButton cmdDelCheque 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   13
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtSubCheques 
            Height          =   285
            Left            =   1440
            TabIndex        =   14
            Top             =   2040
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdCheques 
            Height          =   1575
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2778
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
         End
         Begin VB.Label lblSubCheques 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
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
            Left            =   480
            TabIndex        =   38
            Top             =   2040
            Width           =   900
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Detalle de Deuda"
         Height          =   2415
         Left            =   600
         TabIndex        =   35
         Top             =   1080
         Width           =   7455
         Begin VB.TextBox txtSubDeuda 
            Height          =   285
            Left            =   1440
            TabIndex        =   6
            Top             =   1920
            Width           =   1695
         End
         Begin VB.CommandButton cmdDelDeuda 
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   5
            Top             =   1080
            Width           =   855
         End
         Begin VB.CommandButton cmdAddDeuda 
            Caption         =   "Agregar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   6360
            TabIndex        =   4
            Top             =   480
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid grdDeuda 
            Height          =   1575
            Left            =   480
            TabIndex        =   3
            Top             =   240
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2778
            _Version        =   393216
            FixedCols       =   0
         End
         Begin VB.Label lblSubDeuda 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total:"
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
            Left            =   480
            TabIndex        =   36
            Top             =   1920
            Width           =   900
         End
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   600
         TabIndex        =   31
         Top             =   120
         Width           =   15255
         Begin VB.TextBox txtProveedor 
            Height          =   375
            Left            =   3960
            TabIndex        =   2
            Top             =   360
            Width           =   11055
         End
         Begin VB.TextBox txtFecha 
            Height          =   375
            Left            =   2160
            TabIndex        =   1
            Top             =   360
            Width           =   1335
         End
         Begin VB.TextBox txtNroOrden 
            Height          =   375
            Left            =   360
            TabIndex        =   0
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label lblProveedor 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
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
            Left            =   4080
            TabIndex        =   34
            Top             =   120
            Width           =   885
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   33
            Top             =   120
            Width           =   540
         End
         Begin VB.Label lblNroOrden 
            AutoSize        =   -1  'True
            Caption         =   "Nro. Orden"
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
            Left            =   480
            TabIndex        =   32
            Top             =   120
            Width           =   945
         End
      End
   End
End
Attribute VB_Name = "FormOrdenPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Enum eColTransferencia
    colTransfBanco = 0
    colTransfCuenta = 1
    colTransfCUIT = 2
    colTransfImporte = 3
End Enum

Private Enum eColFacturas
    colFactNumero = 0
    colFactFecha = 1
    colFactImporte = 2
    colFactDescripcion = 3
End Enum

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             Dim c As Control
             Set c = Me.ActiveControl
                                                                                                                                       
             If Not c Is Nothing Then
                 Select Case TypeName(c)
                     Case "MSFlexGrid", "MshFlexGrid", "VSFlexGrid"
                         ' En grilla se maneja aparte
                     Case Else
                         KeyCode = 0
                         Sendkeys "{TAB}"
                 End Select
             End If
         End If
End Sub

Private Sub grdDeuda_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdDeuda
         End If
     End Sub
                                                                                                                                       
     Private Sub grdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdCheques
         End If
     End Sub
                                                                                                                                       
     Private Sub grdOtros_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdOtros
         End If
     End Sub

     Private Sub grdTransferencia_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdTransferencia
         End If
     End Sub

     Private Sub grdFacturas_KeyDown(KeyCode As Integer, Shift As Integer)
         If KeyCode = vbKeyReturn Then
             KeyCode = 0
             EditarCeldaGrid grdFacturas
         End If
     End Sub

Private Sub Form_Load()
    On Error GoTo EH

    FormOrdenPago.Width = 16710
    FormOrdenPago.Height = 11200
    FormOrdenPago.Top = 0
    
    
    
    m_Cargando = True

    EnsureDatabaseReady
    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    ' Transferencia
    With grdTransferencia
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colTransfBanco) = "Banco"
        .TextMatrix(0, colTransfCuenta) = "Nro. Cuenta"
        .TextMatrix(0, colTransfCUIT) = "CUIT"
        .TextMatrix(0, colTransfImporte) = "Importe"
        .ColWidth(colTransfBanco) = 2000
        .ColWidth(colTransfCuenta) = 1800
        .ColWidth(colTransfCUIT) = 1500
        .ColWidth(colTransfImporte) = 1400
    End With

    ' Facturas
    With grdFacturas
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colFactNumero) = "Numero"
        .TextMatrix(0, colFactFecha) = "Fecha"
        .TextMatrix(0, colFactImporte) = "Importe"
        .TextMatrix(0, colFactDescripcion) = "Descripcion"
        .ColWidth(colFactNumero) = 1600
        .ColWidth(colFactFecha) = 1200
        .ColWidth(colFactImporte) = 1400
        .ColWidth(colFactDescripcion) = 3000
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.text = ""
    txtFecha.text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.text = "0,00"
    SetRetIIBBText "0,00"

    txtSubDeuda.text = "0,00"
    txtSubCheques.text = "0,00"
    txtSubOtros.text = "0,00"
    txtSubTransferencia.text = "0,00"
    txtSubFacturas.text = "0,00"
    txtTotalPago.text = "0,00"
    txtSaldo.text = "0,00"
    txtImporteLetras.text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros
    ResetGridRows grdTransferencia
    ResetGridRows grdFacturas

    If nuevoNumero Then
        txtNroOrden.text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub cmdAddTransferencia_Click()
    AddRow grdTransferencia
End Sub

Private Sub cmdDelTransferencia_Click()
    DelCurrentRow grdTransferencia
    RecalcularTodo
End Sub

Private Sub cmdAddFacturas_Click()
    AddRow grdFacturas
End Sub

Private Sub cmdDelFacturas_Click()
    DelCurrentRow grdFacturas
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub grdTransferencia_DblClick()
    EditarCeldaGrid grdTransferencia
End Sub

Private Sub grdFacturas_DblClick()
    EditarCeldaGrid grdFacturas
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    v = Trim$(v)

    If g.Name = "grdDeuda" And c = colDeudaImporte Then
        g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
    ElseIf g.Name = "grdCheques" And c = colChequeImporte Then
        g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
    ElseIf g.Name = "grdOtros" And c = colOtroImporte Then
        g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
    Else
        g.TextMatrix(r, c) = v
    End If

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
         On Error GoTo EH
                                                                                                                                       
         LimpiarFormulario True   ' True => limpia + pide siguiente numero
         txtProveedor.SetFocus
         Exit Sub
                                                                                                                                       
EH:
         MsgBox "Error al restaurar: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo ErrHandler

    Dim y As Single
    Dim xLeft As Single
    Dim xRight As Single
    Dim lineH As Single
    Dim pageBottom As Single
    Dim logoPath As String
    Dim I As Long
    Dim detalle As String
    Dim Importe As String
    Dim nroCheque As String
    Dim banco As String
    Dim vto As String
    Dim cuenta As String
    Dim cuit As String
    Dim numero As String
    Dim fecha As String
    Dim descripcion As String
    Dim deudaRows As Long
    Dim chequeRows As Long
    Dim transfRows As Long
    Dim factRows As Long

    xLeft = 600
    xRight = 7800
    lineH = 240
    pageBottom = 10600

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Copies = 2

    y = 600

    ' Encabezado - Logo condicional si hay facturas
    If TieneFacturas Then
        logoPath = App.Path & "\Quilplac2.jpg"
        If Dir(logoPath) <> "" Then
            Printer.PaintPicture LoadPicture(logoPath), 600, 200, 7200, 1200
            y = 1600
        Else
            Printer.CurrentX = 120
            Printer.CurrentY = 520
            Printer.FontBold = True
            Printer.Print "ORDEN DE PAGO"
            Printer.FontBold = False
        End If
    Else
        Printer.CurrentX = 120
        Printer.CurrentY = 520
        Printer.FontBold = True
        Printer.Print "ORDEN DE PAGO"
        Printer.FontBold = False
    End If

    y = y + lineH * 2

    ' Nro de Orden, Fecha, Proveedor
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Nro: " & Trim$(txtNroOrden.text)
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print "Fecha: " & Trim$(txtFecha.text)
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Proveedor: " & Trim$(txtProveedor.text)
    y = y + lineH * 2

    ' Deuda
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "DEUDA"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    deudaRows = grdDeuda.Rows - 1
    If deudaRows >= 1 Then
        For I = 1 To deudaRows
            detalle = Trim$(grdDeuda.TextMatrix(I, 0))
            Importe = Trim$(grdDeuda.TextMatrix(I, 1))
            If Len(detalle) > 0 Or Len(Importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If
                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$(detalle, 55)
                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(Importe))
                y = y + lineH
            End If
        Next I
    End If

    y = y + lineH / 2
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Deuda:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.text))
    y = y + lineH * 2

    ' Cheques
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "CHEQUES"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(90, "-")
    y = y + lineH

    chequeRows = grdCheques.Rows - 1
    If chequeRows >= 1 Then
        For I = 1 To chequeRows
            banco = Trim$(grdCheques.TextMatrix(I, 0))
            nroCheque = Trim$(grdCheques.TextMatrix(I, 1))
            vto = Trim$(grdCheques.TextMatrix(I, 2))
            Importe = Trim$(grdCheques.TextMatrix(I, 3))
            If Len(banco) > 0 Or Len(nroCheque) > 0 Or Len(vto) > 0 Or Len(Importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If
                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$("Nro: " & nroCheque & "  Banco: " & banco & "  Vto: " & vto, 85)
                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(Importe))
                y = y + lineH
            End If
        Next I
    End If

    y = y + lineH / 2
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Cheques:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubCheques.text))
    y = y + lineH * 2

    ' Transferencia
    transfRows = grdTransferencia.Rows - 1
    If transfRows >= 1 And FilaTieneDatos(grdTransferencia, 1) Then
        If y > pageBottom Then
            Printer.NewPage
            y = 600
        End If

        Printer.FontBold = True
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "TRANSFERENCIA"
        Printer.FontBold = False
        y = y + lineH

        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print String$(90, "-")
        y = y + lineH

        For I = 1 To transfRows
            banco = Trim$(grdTransferencia.TextMatrix(I, 0))
            cuenta = Trim$(grdTransferencia.TextMatrix(I, 1))
            cuit = Trim$(grdTransferencia.TextMatrix(I, 2))
            Importe = Trim$(grdTransferencia.TextMatrix(I, 3))
            If Len(banco) > 0 Or Len(cuenta) > 0 Or Len(cuit) > 0 Or Len(Importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If
                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$(banco & " " & cuenta & " " & cuit, 55)
                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(Importe))
                y = y + lineH
            End If
        Next I

        y = y + lineH / 2
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Subtotal Transferencia:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubTransferencia.text))
        y = y + lineH * 2
    End If

    ' Facturas
    factRows = grdFacturas.Rows - 1
    If factRows >= 1 And FilaTieneDatos(grdFacturas, 1) Then
        If y > pageBottom Then
            Printer.NewPage
            y = 600
        End If

        Printer.FontBold = True
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "FACTURAS"
        Printer.FontBold = False
        y = y + lineH

        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print String$(90, "-")
        y = y + lineH

        For I = 1 To factRows
            numero = Trim$(grdFacturas.TextMatrix(I, 0))
            fecha = Trim$(grdFacturas.TextMatrix(I, 1))
            Importe = Trim$(grdFacturas.TextMatrix(I, 2))
            descripcion = Trim$(grdFacturas.TextMatrix(I, 3))
            If Len(numero) > 0 Or Len(fecha) > 0 Or Len(Importe) > 0 Then
                If y > pageBottom Then
                    Printer.NewPage
                    y = 600
                End If
                Printer.CurrentX = xLeft
                Printer.CurrentY = y
                Printer.Print Left$(numero & " " & fecha & " " & Left$(descripcion, 30), 55)
                Printer.CurrentX = xRight
                Printer.CurrentY = y
                Printer.Print FormatMoney(ParseCurrency(Importe))
                y = y + lineH
            End If
        Next I

        y = y + lineH / 2
        Printer.CurrentX = xLeft
        Printer.CurrentY = y
        Printer.Print "Subtotal Facturas:"
        Printer.CurrentX = xRight
        Printer.CurrentY = y
        Printer.Print FormatMoney(ParseCurrency(txtSubFacturas.text))
        y = y + lineH * 2
    End If

    ' Otros (sin Retencion IIBB)
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "OTROS"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Efectivo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtEfectivo.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Subtotal Otros:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubOtros.text))
    y = y + lineH * 2

    ' RESUMEN - SIN RetIIBB
    If y > pageBottom - (lineH * 10) Then
        Printer.NewPage
        y = 600
    End If

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "RESUMEN"
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Deuda:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Cheques:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubCheques.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Transferencia:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubTransferencia.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Facturas:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubFacturas.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Otros:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSubOtros.text))
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Efectivo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtEfectivo.text))
    y = y + lineH * 1.5

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(70, "-")
    y = y + lineH

    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "TOTAL PAGO:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtTotalPago.text))
    Printer.FontBold = False
    y = y + lineH

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Saldo:"
    Printer.CurrentX = xRight
    Printer.CurrentY = y
    Printer.Print FormatMoney(ParseCurrency(txtSaldo.text))
    y = y + lineH * 2

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Importe en letras:"
    y = y + lineH
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print Trim$(txtImporteLetras.text)
    y = y + lineH * 4

    ' Firma
    If y > pageBottom Then
        Printer.NewPage
        y = 600
    End If

    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print String$(35, "_")
    y = y + lineH
    Printer.CurrentX = xLeft
    Printer.CurrentY = y
    Printer.Print "Firma / Aclaracion"

    Printer.EndDoc
    Exit Sub

ErrHandler:
    MsgBox "Error al imprimir la orden de pago: " & Err.Description, vbExclamation, "Impresion"
End Sub

Private Function TieneFacturas() As Boolean
    Dim I As Integer
    For I = 1 To grdFacturas.Rows - 1
        If FilaTieneDatos(grdFacturas, I) Then
            TieneFacturas = True
            Exit Function
        End If
    Next I
    TieneFacturas = False
End Function
Private Function EnLetras(numero As String) As String
    
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        'If Mid(numero, paso, 1) = "." Then
        If Mid(numero, paso, 1) = "," Then
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
                'MsgBox (Mid(entero, 1, 1))
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "mill?n "
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

Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim subTransferencia As Currency
    Dim subFacturas As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim Saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)
    subTransferencia = SumarColumna(grdTransferencia, colTransfImporte)
    subFacturas = SumarColumna(grdFacturas, colFactImporte)

    efectivo = ParseCurrency(txtEfectivo.text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    totalPago = subCheques + subOtros + subTransferencia + subFacturas + efectivo + retIIBB
    Saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.text = FormatMoney(subDeuda)
    txtSubCheques.text = FormatMoney(subCheques)
    txtSubOtros.text = FormatMoney(subOtros)
    txtSubTransferencia.text = FormatMoney(subTransferencia)
    txtSubFacturas.text = FormatMoney(subFacturas)
    txtTotalPago.text = FormatMoney(totalPago)
    txtSaldo.text = FormatMoney(Saldo)
    txtImporteLetras.text = EnLetras(CStr(totalPago)) 'NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
'    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String
    Dim transfTxt As String, factTxt As String

    EnsureDatabaseReady

    nroOrden = CLng(Val(NzS(txtNroOrden.text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)
    transfTxt = SerializarGrid(grdTransferencia)
    factTxt = SerializarGrid(grdFacturas)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!nroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.text)
    rs!Proveedor = NzS(txtProveedor.text)

    rs!TotalDeuda = ParseCurrency(txtSubDeuda.text)
    rs!SubtotalCheques = ParseCurrency(txtSubCheques.text)
    rs!SubtotalOtros = ParseCurrency(txtSubOtros.text)
    rs!SubtotalTransferencia = ParseCurrency(txtSubTransferencia.text)
    rs!SubtotalFacturas = ParseCurrency(txtSubFacturas.text)
    rs!efectivo = ParseCurrency(txtEfectivo.text)
'   rs!RetencionIIBB = ParseCurrency(GetRetIIBBText())
    rs!totalPago = ParseCurrency(txtTotalPago.text)
    rs!Saldo = ParseCurrency(txtSaldo.text)
    rs!MontoLetras = NzS(txtImporteLetras.text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt
    rs!DetalleTransferencia = transfTxt
    rs!DetalleFacturas = factTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim Texto As String

    EnsureDatabaseReady

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        Texto = CStr(NzL(rs!nroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem Texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!nroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & CStr(Err.Number) & " - " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.text = CStr(NzL(rs!nroOrden))
    txtFecha.text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.text = NzS(rs!Proveedor)

    txtSubDeuda.text = FormatMoney(NzC(rs!TotalDeuda))
    txtSubCheques.text = FormatMoney(NzC(rs!SubtotalCheques))
    txtSubOtros.text = FormatMoney(NzC(rs!SubtotalOtros))
    txtSubTransferencia.text = FormatMoney(NzC(rs!SubtotalTransferencia))
    txtSubFacturas.text = FormatMoney(NzC(rs!SubtotalFacturas))
    txtEfectivo.text = FormatMoney(NzC(rs!efectivo))
    SetRetIIBBText FormatMoney(NzC(rs!RetencionIIBB))
    txtTotalPago.text = FormatMoney(NzC(rs!totalPago))
    txtSaldo.text = FormatMoney(NzC(rs!Saldo))
    txtImporteLetras.text = NzS(rs!MontoLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)
    DeserializarGrid grdTransferencia, NzS(rs!DetalleTransferencia)
    DeserializarGrid grdFacturas, NzS(rs!DetalleFacturas)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub



' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim IdSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    EnsureDatabaseReady

    IdSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(IdSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = IdSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim IdSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    IdSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(IdSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!IdSucursal = IdSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    Dim r As Integer, c As Integer
    Dim linea As String
    Dim salida As String

    salida = ""
    For r = 1 To g.Rows - 1
        If FilaTieneDatos(g, r) Then
            linea = ""
            For c = 0 To g.Cols - 1
                If c > 0 Then linea = linea & SEP_COL
                linea = linea & EscaparDato(g.TextMatrix(r, c))
            Next c
            If Len(salida) > 0 Then salida = salida & SEP_ROW
            salida = salida & linea
        End If
    Next r

    SerializarGrid = salida
End Function

Private Sub DeserializarGrid(ByRef g As MSFlexGrid, ByVal dataText As String)
    On Error GoTo EH

    Dim rowsArr() As String
    Dim colsArr() As String
    Dim r As Long, c As Long
    Dim targetRow As Integer

    ResetGridRows g

    If Trim$(dataText) = "" Then Exit Sub

    rowsArr = Split(dataText, SEP_ROW)
    g.Rows = UBound(rowsArr) + 2

    targetRow = 1
    For r = LBound(rowsArr) To UBound(rowsArr)
        colsArr = Split(rowsArr(r), SEP_COL)
        For c = 0 To g.Cols - 1
            If c <= UBound(colsArr) Then
                g.TextMatrix(targetRow, c) = DesescaparDato(colsArr(c))
            Else
                g.TextMatrix(targetRow, c) = ""
            End If
        Next c
        targetRow = targetRow + 1
    Next r

    Exit Sub
EH:
    MsgBox "Error deserializando grilla: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    Dim pDot As Long, pCom As Long, decSep As String
    Dim intPart As String, decPart As String
    Dim sign As Double
    
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseCurrency = 0
        Exit Function
    End If
    
    t = Replace$(t, " ", "")
    t = Replace$(t, "$", "")
    
    sign = 1
    If Left$(t, 1) = "-" Then
        sign = -1
        t = Mid$(t, 2)
    ElseIf Left$(t, 1) = "+" Then
        t = Mid$(t, 2)
    End If
    
    pDot = InStrRev(t, ".")
    pCom = InStrRev(t, ",")
    
    If pDot > 0 And pCom > 0 Then
        If pDot > pCom Then
            decSep = "."
        Else
            decSep = ","
        End If
    ElseIf pDot > 0 Then
        If Len(t) - pDot = 3 Then
            decSep = ""    ' miles
        Else
            decSep = "."   ' decimal
        End If
    ElseIf pCom > 0 Then
        If Len(t) - pCom = 3 Then
            decSep = ""    ' miles
        Else
            decSep = ","   ' decimal
        End If
    Else
        decSep = ""
    End If
    
    If decSep = "" Then
        t = Replace$(t, ".", "")
        t = Replace$(t, ",", "")
        If Len(t) = 0 Then t = "0"
        ParseCurrency = CCur(sign * Val(t))
        Exit Function
    End If
    
    If decSep = "." Then
        intPart = Left$(t, InStrRev(t, ".") - 1)
        decPart = Mid$(t, InStrRev(t, ".") + 1)
    Else
        intPart = Left$(t, InStrRev(t, ",") - 1)
        decPart = Mid$(t, InStrRev(t, ",") + 1)
    End If
    
    intPart = Replace$(intPart, ".", "")
    intPart = Replace$(intPart, ",", "")
    If Len(intPart) = 0 Then intPart = "0"
    
    If Len(decPart) = 0 Then
        decPart = "00"
    ElseIf Len(decPart) = 1 Then
        decPart = decPart & "0"
    ElseIf Len(decPart) > 2 Then
        decPart = Left$(decPart, 2)
    End If
    
    ParseCurrency = CCur(sign * Val(intPart & "." & decPart))
End Function

Private Function GroupThousands(ByVal digits As String) As String
    Dim I As Long, out As String, cnt As Long
    
    out = ""
    cnt = 0
    
    For I = Len(digits) To 1 Step -1
        out = Mid$(digits, I, 1) & out
        cnt = cnt + 1
        If cnt = 3 And I > 1 Then
            out = "." & out
            cnt = 0
        End If
    Next I
    
    GroupThousands = out
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    Dim isNeg As Boolean
    Dim absV As Currency
    Dim cents As Long
    Dim intPart As Long
    Dim decPart As Long
    Dim intTxt As String
    
    isNeg = (v < 0)
    absV = Abs(v)
    
    cents = CLng((absV * 100) + 0.5)
    intPart = cents \ 100
    decPart = cents Mod 100
    
    intTxt = GroupThousands(CStr(intPart))
    
    If isNeg Then
        FormatMoney = "-" & intTxt & "," & Right$("0" & CStr(decPart), 2)
    Else
        FormatMoney = intTxt & "," & Right$("0" & CStr(decPart), 2)
    End If
End Function

Private Sub FormatImporteTextBox(ByRef txt As TextBox)
    txt.text = FormatMoney(ParseCurrency(txt.text))
End Sub

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal Importe As Currency) As String
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(Importe)
End Function

' ------------------------------------------------------------
' Helpers de compatibilidad (control opcional txtRetIIBB)
' ------------------------------------------------------------
Private Sub EnsureDatabaseReady()
    On Error GoTo EH

    If BaseSPC Is Nothing Then
        Set BaseSPC = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    End If
    Exit Sub
EH:
    MsgBox "No se pudo abrir la base de datos: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

Private Function TieneControl(ByVal nombreControl As String) As Boolean
    On Error GoTo EH
    Dim ctl As Control
    Set ctl = Me.Controls(nombreControl)
    TieneControl = Not (ctl Is Nothing)
    Exit Function
EH:
    TieneControl = False
End Function

Private Function GetRetIIBBText() As String
    On Error GoTo EH
    If TieneControl("txtRetIIBB") Then
        GetRetIIBBText = CStr(Me.Controls("txtRetIIBB").text)
    Else
        GetRetIIBBText = "0,00"
    End If
    Exit Function
EH:
    GetRetIIBBText = "0,00"
End Function

Private Sub SetRetIIBBText(ByVal valor As String)
    On Error Resume Next
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").text = valor
    End If
End Sub

Private Sub txtEfectivo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtEfectivo
    RecalcularTodo
End Sub

Private Sub txtSaldo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSaldo
End Sub

Private Sub txtRetIIBB_LostFocus()
    If m_Cargando Then Exit Sub
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").text = FormatMoney(ParseCurrency(Me.Controls("txtRetIIBB").text))
    End If
    RecalcularTodo
End Sub
Private Sub txtSubCheques_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubCheques
End Sub
Private Sub txtSubDeuda_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubDeuda
End Sub

Private Sub txtSubOtros_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubOtros
End Sub

Private Sub txtTotalPago_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtTotalPago
End Sub
