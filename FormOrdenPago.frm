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
   Begin VB.Frame Fr_Base 
      Height          =   10455
      Left            =   120
      TabIndex        =   35
      Top             =   0
      Width           =   16215
      Begin VB.Frame FrReimpresion 
         Caption         =   "Reimpresi�n"
         Height          =   2415
         Left            =   8400
         TabIndex        =   63
         Top             =   6240
         Width           =   7455
         Begin VB.ComboBox cboBuscarOrden 
            Height          =   315
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   34
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
            TabIndex        =   64
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
            TabIndex        =   65
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
      Begin VB.Frame FrTransferencia 
         Caption         =   "Detalle de Transferencia"
         Height          =   2535
         Left            =   8400
         TabIndex        =   59
         Top             =   3600
         Width           =   7455
         Begin VB.TextBox txtGE_Transferencia 
            Height          =   315
            Left            =   0
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   960
            Visible         =   0   'False
            Width           =   1215
         End
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
            TabIndex        =   17
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
            TabIndex        =   18
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
            TabIndex        =   61
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
            TabIndex        =   60
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtSubTransferencia 
            Height          =   285
            Left            =   1440
            TabIndex        =   31
            Top             =   2040
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdTransferencia 
            Height          =   1575
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   2778
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
            TabIndex        =   62
            Top             =   2040
            Width           =   900
         End
      End
      Begin VB.Frame FrFacturas 
         Caption         =   "Detalle de Facturas"
         Height          =   2415
         Left            =   8400
         TabIndex        =   55
         Top             =   1080
         Width           =   7455
         Begin VB.TextBox txtGE_Facturas 
            Height          =   315
            Left            =   0
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
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
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   57
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
            TabIndex        =   56
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox txtSubFacturas 
            Height          =   285
            Left            =   1440
            TabIndex        =   29
            Top             =   1800
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdFacturas 
            Height          =   1455
            Left            =   480
            TabIndex        =   8
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
            TabIndex        =   58
            Top             =   1800
            Width           =   900
         End
      End
      Begin VB.Frame FrTotales 
         Caption         =   "Totales"
         Height          =   1335
         Left            =   600
         TabIndex        =   48
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
            TabIndex        =   54
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
            TabIndex        =   33
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
            TabIndex        =   53
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
            Left            =   4920
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
            TabIndex        =   52
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
            Left            =   5040
            TabIndex        =   51
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
            TabIndex        =   50
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
            TabIndex        =   49
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.Frame FrOtros 
         Caption         =   "Detalle de Otros"
         Height          =   2415
         Left            =   600
         TabIndex        =   44
         Top             =   6240
         Width           =   7455
         Begin VB.TextBox txtGE_Otros 
            Height          =   315
            Left            =   240
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
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
            TabIndex        =   22
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
            TabIndex        =   21
            Top             =   360
            Width           =   855
         End
         Begin VB.TextBox txtSubOtros 
            Height          =   285
            Left            =   1440
            TabIndex        =   32
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
            TabIndex        =   46
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
            TabIndex        =   45
            Top             =   600
            Width           =   1215
         End
         Begin MSFlexGridLib.MSFlexGrid grdOtros 
            Height          =   1455
            Left            =   480
            TabIndex        =   20
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
            TabIndex        =   47
            Top             =   1800
            Width           =   900
         End
      End
      Begin VB.Frame FrCheques 
         Caption         =   "Detalle de Cheques"
         Height          =   2535
         Left            =   600
         TabIndex        =   42
         Top             =   3600
         Width           =   7455
         Begin VB.TextBox txtGE_Cheques 
            Height          =   315
            Left            =   0
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
         End
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
            TabIndex        =   13
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
            TabIndex        =   14
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtSubCheques 
            Height          =   285
            Left            =   1440
            TabIndex        =   30
            Top             =   2040
            Width           =   1695
         End
         Begin MSFlexGridLib.MSFlexGrid grdCheques 
            Height          =   1575
            Left            =   480
            TabIndex        =   12
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
            TabIndex        =   43
            Top             =   2040
            Width           =   900
         End
      End
      Begin VB.Frame FrDeuda 
         Caption         =   "Detalle de Deuda"
         Height          =   2415
         Left            =   600
         TabIndex        =   40
         Top             =   1080
         Width           =   7455
         Begin VB.TextBox txtGE_Deuda 
            Height          =   315
            Left            =   0
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   720
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtSubDeuda 
            Height          =   285
            Left            =   1440
            TabIndex        =   28
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
            TabIndex        =   6
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
            TabIndex        =   5
            Top             =   480
            Width           =   855
         End
         Begin MSFlexGridLib.MSFlexGrid grdDeuda 
            Height          =   1575
            Left            =   480
            TabIndex        =   4
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
            TabIndex        =   41
            Top             =   1920
            Width           =   900
         End
      End
      Begin VB.Frame FrEncabezado 
         Height          =   855
         Left            =   600
         TabIndex        =   36
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
            TabIndex        =   39
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
            TabIndex        =   38
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
            TabIndex        =   37
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

' ============================================================
' NEW: Inline editing support (Phase 7)
' ============================================================
Private m_EditGrid As MSFlexGrid
Private m_EditRow As Integer
Private m_EditCol As Integer
Private m_EditingActive As Boolean
Private m_ValorAnterior As String
Private m_ActiveEditorTextBox As TextBox
Private m_SuppressEditorLostFocusCommit As Boolean
Private m_EditGridName As String
Private m_EditEditorName As String
Private m_FormattingEditorChange As Boolean

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
         If KeyCode <> vbKeyReturn Then Exit Sub

         Dim c As Control
         Set c = Me.ActiveControl

         If c Is Nothing Then Exit Sub

         Select Case TypeName(c)
             Case "MSFlexGrid", "MshFlexGrid", "VSFlexGrid"
                 ' En grilla se maneja aparte
             Case "TextBox"
                 If m_EditingActive Then
                     ' Durante edición inline, Enter/Tab se resuelve en ManejadorTextBoxKeyDown.
                     Exit Sub
                 End If
                 KeyCode = 0
                 Sendkeys "{TAB}"
             Case Else
                 KeyCode = 0
                 Sendkeys "{TAB}"
         End Select
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

Private Sub cmdAddOtros_Click()
    cmdAddOtro_Click
End Sub

Private Sub cmdDelOtros_Click()
    cmdDelOtro_Click
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
    g.col = 0
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

' ============================================================
' NEW: Inline editing system (Phase 7 - replaces InputBox)
' ============================================================

' Buscar recursivamente el Frame que contiene un grid (busca en Frames anidados)
Private Function FindParentFrame(ByRef grid As MSFlexGrid) As Frame
    On Error Resume Next

    If TypeName(grid.Container) = "Frame" Then
        Set FindParentFrame = grid.Container
    Else
        Set FindParentFrame = Nothing
    End If
End Function


Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim leftPos As Long, topPos As Long, wid As Long, hgt As Long
    Dim parentFr As Frame

    r = g.Row
    c = g.col

    If r < 1 Or c < 0 Then Exit Sub

    ' Guardar estado
    Set m_EditGrid = g
    m_EditRow = r
    m_EditCol = c
    m_ValorAnterior = g.TextMatrix(r, c)
    m_EditingActive = True
    m_EditGridName = g.Name
    m_EditEditorName = ""

    ' Calcular posición de la celda respecto del contenedor del grid
    leftPos = g.Left + g.CellLeft
    topPos = g.Top + g.CellTop

    wid = g.ColWidth(c)
    hgt = g.RowHeight(r)

    ' Margen
    Const MARGIN As Long = 2
    leftPos = leftPos + MARGIN
    topPos = topPos + MARGIN
    wid = wid - (MARGIN * 2)
    hgt = hgt - (MARGIN * 2)

    Set parentFr = FindParentFrame(g)
    If parentFr Is Nothing Then Exit Sub

    ' Posicionar TextBox en el mismo contenedor que el grid
    Select Case g.Name
        Case "grdDeuda"
            Set m_ActiveEditorTextBox = txtGE_Deuda
            m_EditEditorName = txtGE_Deuda.Name
            With txtGE_Deuda
                .Move leftPos, topPos, wid, hgt
                .text = g.TextMatrix(r, c)
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With

        Case "grdCheques"
            Set m_ActiveEditorTextBox = txtGE_Cheques
            m_EditEditorName = txtGE_Cheques.Name
            With txtGE_Cheques
                .Move leftPos, topPos, wid, hgt
                .text = g.TextMatrix(r, c)
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With

        Case "grdOtros"
            Set m_ActiveEditorTextBox = txtGE_Otros
            m_EditEditorName = txtGE_Otros.Name
            With txtGE_Otros
                .Move leftPos, topPos, wid, hgt
                .text = g.TextMatrix(r, c)
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With

        Case "grdTransferencia"
            Set m_ActiveEditorTextBox = txtGE_Transferencia
            m_EditEditorName = txtGE_Transferencia.Name
            With txtGE_Transferencia
                .Move leftPos, topPos, wid, hgt
                .text = g.TextMatrix(r, c)
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With

        Case "grdFacturas"
            Set m_ActiveEditorTextBox = txtGE_Facturas
            m_EditEditorName = txtGE_Facturas.Name
            With txtGE_Facturas
                .Move leftPos, topPos, wid, hgt
                .text = g.TextMatrix(r, c)
                .Visible = True
                .SetFocus
                .SelStart = 0
                .SelLength = Len(.text)
            End With
    End Select

    Exit Sub
EH:
    MsgBox "Error en EditarCeldaGrid: " & Err.Description, vbExclamation
End Sub

Private Function GetGridNameFromEditor(ByRef txt As TextBox) As String
    If txt Is Nothing Then Exit Function

    Select Case txt.Name
        Case "txtGE_Deuda": GetGridNameFromEditor = "grdDeuda"
        Case "txtGE_Cheques": GetGridNameFromEditor = "grdCheques"
        Case "txtGE_Otros": GetGridNameFromEditor = "grdOtros"
        Case "txtGE_Transferencia": GetGridNameFromEditor = "grdTransferencia"
        Case "txtGE_Facturas": GetGridNameFromEditor = "grdFacturas"
        Case Else: GetGridNameFromEditor = ""
    End Select
End Function

Private Function GetAddButtonByGridName(ByVal gridName As String) As CommandButton
    Select Case gridName
        Case "grdDeuda": Set GetAddButtonByGridName = cmdAddDeuda
        Case "grdCheques": Set GetAddButtonByGridName = cmdAddCheque
        Case "grdOtros": Set GetAddButtonByGridName = cmdAddOtros
        Case "grdTransferencia": Set GetAddButtonByGridName = cmdAddTransferencia
        Case "grdFacturas": Set GetAddButtonByGridName = cmdAddFacturas
        Case Else: Set GetAddButtonByGridName = Nothing
    End Select
End Function

Private Function ResolveGridByName(ByVal gridName As String) As MSFlexGrid
    Select Case gridName
        Case "grdDeuda": Set ResolveGridByName = grdDeuda
        Case "grdCheques": Set ResolveGridByName = grdCheques
        Case "grdOtros": Set ResolveGridByName = grdOtros
        Case "grdTransferencia": Set ResolveGridByName = grdTransferencia
        Case "grdFacturas": Set ResolveGridByName = grdFacturas
        Case Else: Set ResolveGridByName = Nothing
    End Select
End Function

Private Function IsEditorBoundToGrid(ByRef txt As TextBox, ByVal gridName As String) As Boolean
    If txt Is Nothing Then Exit Function
    IsEditorBoundToGrid = (StrComp(GetGridNameFromEditor(txt), gridName, vbTextCompare) = 0)
End Function

Private Sub ResetEditContext()
    Set m_EditGrid = Nothing
    Set m_ActiveEditorTextBox = Nothing
    m_EditingActive = False
    m_SuppressEditorLostFocusCommit = False
    m_EditGridName = ""
    m_EditEditorName = ""
    m_EditRow = 0
    m_EditCol = 0
    m_ValorAnterior = ""
End Sub

Private Sub txtGE_Deuda_KeyDown(KeyCode As Integer, Shift As Integer)
    ManejadorTextBoxKeyDown KeyCode, Shift
End Sub

Private Sub txtGE_Deuda_Change()
    ManejadorTextBoxChange
End Sub

Private Sub txtGE_Deuda_LostFocus()
    ManejadorTextBoxLostFocus
End Sub

Private Sub txtGE_Cheques_KeyDown(KeyCode As Integer, Shift As Integer)
    ManejadorTextBoxKeyDown KeyCode, Shift
End Sub

Private Sub txtGE_Cheques_Change()
    ManejadorTextBoxChange
End Sub

Private Sub txtGE_Cheques_LostFocus()
    ManejadorTextBoxLostFocus
End Sub

Private Sub txtGE_Otros_KeyDown(KeyCode As Integer, Shift As Integer)
    ManejadorTextBoxKeyDown KeyCode, Shift
End Sub

Private Sub txtGE_Otros_Change()
    ManejadorTextBoxChange
End Sub

Private Sub txtGE_Otros_LostFocus()
    ManejadorTextBoxLostFocus
End Sub

Private Sub txtGE_Transferencia_KeyDown(KeyCode As Integer, Shift As Integer)
    ManejadorTextBoxKeyDown KeyCode, Shift
End Sub

Private Sub txtGE_Transferencia_Change()
    ManejadorTextBoxChange
End Sub

Private Sub txtGE_Transferencia_LostFocus()
    ManejadorTextBoxLostFocus
End Sub

Private Sub txtGE_Facturas_KeyDown(KeyCode As Integer, Shift As Integer)
    ManejadorTextBoxKeyDown KeyCode, Shift
End Sub

Private Sub txtGE_Facturas_Change()
    ManejadorTextBoxChange
End Sub

Private Sub txtGE_Facturas_LostFocus()
    ManejadorTextBoxLostFocus
End Sub

' ============================================================
' txtGridEditor event handlers (Phase 7)
' ============================================================
' Event handlers para los TextBox de edición
' Se vinculan dinámicamente en Form_Load
Private Sub ManejadorTextBoxChange()
    On Error GoTo EH

    If m_FormattingEditorChange Then Exit Sub
    If Not m_EditingActive Or m_ActiveEditorTextBox Is Nothing Then Exit Sub
    If Not EsColumnaImporteEnEdicion() Then Exit Sub

    Dim inputText As String
    Dim formatted As String
    Dim oldSelStart As Long
    Dim newSelStart As Long
    Dim previewValue As String

    inputText = m_ActiveEditorTextBox.text
    oldSelStart = m_ActiveEditorTextBox.SelStart

    formatted = FormatearImporteEnVivo(inputText, oldSelStart, newSelStart)

    If StrComp(formatted, inputText, vbBinaryCompare) <> 0 Then
        m_FormattingEditorChange = True
        m_ActiveEditorTextBox.text = formatted
        m_ActiveEditorTextBox.SelStart = newSelStart
        m_FormattingEditorChange = False
    End If

    previewValue = ObtenerPreviewImporte(inputText, formatted)
    If Not m_EditGrid Is Nothing Then
        m_EditGrid.TextMatrix(m_EditRow, m_EditCol) = previewValue
    End If

    Exit Sub
EH:
    m_FormattingEditorChange = False
End Sub

Private Function EsColumnaImporteEnEdicion() As Boolean
    If Not m_EditGrid Is Nothing Then
        EsColumnaImporteEnEdicion = EsColumnaImportePorGrid(m_EditGrid.Name, m_EditCol)
    ElseIf Len(m_EditGridName) > 0 Then
        EsColumnaImporteEnEdicion = EsColumnaImportePorGrid(m_EditGridName, m_EditCol)
    End If
End Function

Private Function EsColumnaImportePorGrid(ByVal gridName As String, ByVal col As Integer) As Boolean
    Select Case gridName
        Case "grdDeuda"
            EsColumnaImportePorGrid = (col = colDeudaImporte)
        Case "grdCheques"
            EsColumnaImportePorGrid = (col = colChequeImporte)
        Case "grdOtros"
            EsColumnaImportePorGrid = (col = colOtroImporte)
        Case "grdTransferencia"
            EsColumnaImportePorGrid = (col = colTransfImporte)
        Case "grdFacturas"
            EsColumnaImportePorGrid = (col = colFactImporte)
    End Select
End Function

Private Function EsColumnaFechaPorGrid(ByVal gridName As String, ByVal col As Integer) As Boolean
    Select Case gridName
        Case "grdCheques"
            EsColumnaFechaPorGrid = (col = colChequeFecha)
        Case "grdFacturas"
            EsColumnaFechaPorGrid = (col = colFactFecha)
    End Select
End Function

Private Function FormatearImporteEnVivo(ByVal rawText As String, ByVal caretPos As Long, ByRef newCaretPos As Long) As String
    Dim I As Long
    Dim ch As String
    Dim digits As String
    Dim intDigits As String
    Dim decDigits As String
    Dim hadComma As Boolean
    Dim commaPos As Long
    Dim digitsLeftOfCaret As Long
    Dim outText As String

    If caretPos < 0 Then caretPos = 0
    If caretPos > Len(rawText) Then caretPos = Len(rawText)

    For I = 1 To Len(rawText)
        ch = Mid$(rawText, I, 1)

        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
            If I <= caretPos Then digitsLeftOfCaret = digitsLeftOfCaret + 1
        ElseIf ch = "," And Not hadComma Then
            hadComma = True
            commaPos = Len(digits)
        End If
    Next I

    If Len(digits) = 0 Then
        FormatearImporteEnVivo = ""
        newCaretPos = 0
        Exit Function
    End If

    If hadComma Then
        intDigits = Left$(digits, commaPos)
        decDigits = Mid$(digits, commaPos + 1)

        If Len(intDigits) = 0 Then intDigits = "0"
        If Len(decDigits) > 2 Then decDigits = Left$(decDigits, 2)

        outText = GroupThousands(intDigits)
        If commaPos >= 0 Then outText = outText & "," & decDigits
    Else
        outText = GroupThousands(digits)
    End If

    FormatearImporteEnVivo = outText
    newCaretPos = PosicionCursorPorDigitos(outText, digitsLeftOfCaret)
End Function

Private Function PosicionCursorPorDigitos(ByVal formattedText As String, ByVal digitsToLeft As Long) As Long
    Dim I As Long
    Dim seenDigits As Long
    Dim ch As String

    If digitsToLeft <= 0 Then
        PosicionCursorPorDigitos = 0
        Exit Function
    End If

    For I = 1 To Len(formattedText)
        ch = Mid$(formattedText, I, 1)
        If ch >= "0" And ch <= "9" Then
            seenDigits = seenDigits + 1
            If seenDigits >= digitsToLeft Then
                PosicionCursorPorDigitos = I
                Exit Function
            End If
        End If
    Next I

    PosicionCursorPorDigitos = Len(formattedText)
End Function

Private Function ObtenerPreviewImporte(ByVal originalText As String, ByVal formattedText As String) As String
    On Error GoTo EH

    If Len(formattedText) = 0 Then
        ObtenerPreviewImporte = ""
        Exit Function
    End If

    ObtenerPreviewImporte = FormatMoney(ParseCurrency(formattedText))
    Exit Function
EH:
    ObtenerPreviewImporte = formattedText
End Function

Private Sub ManejadorTextBoxKeyDown(KeyCode As Integer, Shift As Integer)
    If Not m_EditingActive Then Exit Sub

    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        Shift = 0
        Sendkeys "{TAB}"
    ElseIf KeyCode = vbKeyTab Then
        KeyCode = 0
        Shift = 0
        m_SuppressEditorLostFocusCommit = True
        FinalizarEdicion True, True
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        FinalizarEdicion False
    End If
End Sub

Private Sub ManejadorTextBoxLostFocus()
    If m_SuppressEditorLostFocusCommit Then
        m_SuppressEditorLostFocusCommit = False
        Exit Sub
    End If

    If m_EditingActive Then
        FinalizarEdicion True
    End If
End Sub

Private Sub FinalizarEdicion(ByVal guardar As Boolean, Optional ByVal navegarSiguiente As Boolean = False)
    On Error GoTo EH

    If Not m_EditingActive Then Exit Sub

    Dim valorNuevo As String
    Dim valorFinal As String
    Dim g As MSFlexGrid
    Dim r As Integer
    Dim c As Integer
    Dim capturedGridName As String
    Dim editorGridName As String
    Dim navGrid As MSFlexGrid

    Set g = m_EditGrid
    r = m_EditRow
    c = m_EditCol
    capturedGridName = m_EditGridName

    If m_ActiveEditorTextBox Is Nothing Then
        ResetEditContext
        Exit Sub
    End If

    If Len(m_EditEditorName) > 0 Then
        If StrComp(m_ActiveEditorTextBox.Name, m_EditEditorName, vbTextCompare) <> 0 Then
            EnfocarBotonAddPorGridName GetGridNameFromEditor(m_ActiveEditorTextBox)
            ResetEditContext
            Exit Sub
        End If
    End If

    editorGridName = GetGridNameFromEditor(m_ActiveEditorTextBox)
    If Len(editorGridName) = 0 Then
        If Len(capturedGridName) > 0 Then EnfocarBotonAddPorGridName capturedGridName
        ResetEditContext
        Exit Sub
    End If

    If Len(capturedGridName) = 0 Then capturedGridName = editorGridName

    If Not IsEditorBoundToGrid(m_ActiveEditorTextBox, capturedGridName) Then
        EnfocarBotonAddPorGridName editorGridName
        ResetEditContext
        Exit Sub
    End If

    If g Is Nothing Then
        Set g = ResolveGridByName(capturedGridName)
    End If
    If g Is Nothing Then
        EnfocarBotonAddPorGridName capturedGridName
        ResetEditContext
        Exit Sub
    End If

    Set navGrid = g

    If guardar Then
        valorNuevo = Trim$(m_ActiveEditorTextBox.text)

        ' Apply column-specific validation and formatting
        If Not ValidarYFormatearCelda(m_EditGrid, m_EditRow, m_EditCol, valorNuevo, valorFinal) Then
            ' Validation failed - keep editor open
            m_SuppressEditorLostFocusCommit = False
            Exit Sub
        End If

        ' Save to grid
        m_EditGrid.TextMatrix(m_EditRow, m_EditCol) = valorFinal

        ' Recalculate if importe column
        If Not m_Cargando Then
            RecalcularTodo
        End If
    End If

    ' Hide editor
    If Not m_ActiveEditorTextBox Is Nothing Then m_ActiveEditorTextBox.Visible = False
    ResetEditContext

    If guardar And navegarSiguiente Then
        NavegarPostCommit navGrid, r, c, capturedGridName
    End If

    Exit Sub
EH:
    MsgBox "Error finalizando edicion: " & Err.Description, vbExclamation, "Orden de Pago"
    If Not m_ActiveEditorTextBox Is Nothing Then m_ActiveEditorTextBox.Visible = False
    ResetEditContext
End Sub

Private Sub NavegarPostCommit(ByRef g As MSFlexGrid, ByVal r As Integer, ByVal c As Integer, ByVal gridName As String)
    On Error GoTo EH

    If g Is Nothing Then
        If Len(gridName) = 0 Then Exit Sub
        EnfocarBotonAddPorGridName gridName
        Exit Sub
    End If

    If Len(gridName) > 0 Then
        If StrComp(g.Name, gridName, vbTextCompare) <> 0 Then
            EnfocarBotonAddPorGridName gridName
            Exit Sub
        End If
    Else
        gridName = g.Name
    End If

    Dim nextCol As Integer
    nextCol = BuscarSiguienteColumnaEditable(g, c)

    If nextCol >= 0 Then
        g.Row = r
        g.col = nextCol
        EditarCeldaGrid g
    Else
        EnfocarBotonAddPorGridName gridName
    End If

    Exit Sub
EH:
    EnfocarBotonAddPorGridName gridName
End Sub

Private Function BuscarSiguienteColumnaEditable(ByRef g As MSFlexGrid, ByVal currentCol As Integer) As Integer
    Dim I As Integer

    BuscarSiguienteColumnaEditable = -1

    For I = currentCol + 1 To g.Cols - 1
        If EsColumnaEditable(g, I) Then
            BuscarSiguienteColumnaEditable = I
            Exit Function
        End If
    Next I
End Function

Private Function EsColumnaEditable(ByRef g As MSFlexGrid, ByVal col As Integer) As Boolean
    EsColumnaEditable = False

    If col < g.FixedCols Then Exit Function
    If g.ColWidth(col) <= 0 Then Exit Function

    EsColumnaEditable = True
End Function

Private Sub EnfocarBotonAddPorGrid(ByRef g As MSFlexGrid)
    On Error Resume Next

    If g Is Nothing Then Exit Sub

    EnfocarBotonAddPorGridName g.Name
End Sub

Private Sub EnfocarBotonAddPorGridName(ByVal gridName As String)
    On Error Resume Next

    Dim cmd As CommandButton
    Set cmd = GetAddButtonByGridName(gridName)
    If cmd Is Nothing Then Exit Sub
    cmd.SetFocus
End Sub

Private Function ValidarYFormatearCelda(ByRef g As MSFlexGrid, ByVal r As Integer, ByVal c As Integer, ByVal valorNuevo As String, ByRef valorFinal As String) As Boolean
    On Error GoTo EH

    ValidarYFormatearCelda = False
    valorFinal = valorNuevo

    If EsColumnaImportePorGrid(g.Name, c) Then
        valorFinal = FormatMoney(ParseCurrency(valorNuevo))
        ValidarYFormatearCelda = True
        Exit Function
    End If

    If EsColumnaFechaPorGrid(g.Name, c) Then
        If Len(valorNuevo) > 0 Then
            If Not TryNormalizarFecha(valorNuevo, valorFinal) Then
                MsgBox "Fecha invalida. Use formato DD/MM/YYYY o una fecha reconocible.", vbExclamation, "Validacion"
                m_ActiveEditorTextBox.SelStart = 0
                m_ActiveEditorTextBox.SelLength = Len(m_ActiveEditorTextBox.text)
                Exit Function
            End If
        End If
        ValidarYFormatearCelda = True
        Exit Function
    End If

    ValidarYFormatearCelda = True

    Exit Function
EH:
    MsgBox "Error en validacion: " & Err.Description, vbExclamation, "Validacion"
    ValidarYFormatearCelda = False
End Function

' ============================================================
' NEW: Validation helper functions (Phase 7)
' ============================================================
Private Function TryNormalizarFecha(ByVal fechaInput As String, ByRef fechaFormateada As String) As Boolean
    On Error GoTo EH

    Dim fechaLimpia As String
    Dim fechaValor As Date

    TryNormalizarFecha = False
    fechaFormateada = Trim$(fechaInput)
    fechaLimpia = Trim$(fechaInput)
    If Len(fechaLimpia) = 0 Then
        TryNormalizarFecha = True
        Exit Function
    End If

    If ValidarFechaDD_MM_YYYY(fechaLimpia) Then
        fechaValor = DateSerial(CInt(Mid$(fechaLimpia, 7, 4)), CInt(Mid$(fechaLimpia, 4, 2)), CInt(Mid$(fechaLimpia, 1, 2)))
        fechaFormateada = Format$(fechaValor, "dd/mm/yyyy")
        TryNormalizarFecha = True
        Exit Function
    End If

    If IsDate(fechaLimpia) Then
        fechaValor = CDate(fechaLimpia)
        fechaFormateada = Format$(fechaValor, "dd/mm/yyyy")
        TryNormalizarFecha = True
    End If
    Exit Function
EH:
    TryNormalizarFecha = False
End Function

Private Function ValidarFechaDD_MM_YYYY(ByVal fechaStr As String) As Boolean
    On Error GoTo EH

    Dim partes() As String
    Dim Dia As Integer, Mes As Integer, anno As Integer

    ValidarFechaDD_MM_YYYY = False

    If Len(fechaStr) <> 10 Then Exit Function
    If Mid(fechaStr, 3, 1) <> "/" Or Mid(fechaStr, 6, 1) <> "/" Then Exit Function

    partes = Split(fechaStr, "/")
    If UBound(partes) <> 2 Then Exit Function

    Dia = CInt(partes(0))
    Mes = CInt(partes(1))
    anno = CInt(partes(2))

    ' Basic range checks
    If Mes < 1 Or Mes > 12 Then Exit Function
    If Dia < 1 Or Dia > 31 Then Exit Function
    If anno < 1900 Or anno > 2100 Then Exit Function

    ' Try to create date (will fail if invalid like 31/02/2026)
    Dim testDate As Date
    testDate = DateSerial(anno, Mes, Dia)

    ValidarFechaDD_MM_YYYY = True
    Exit Function
EH:
    ValidarFechaDD_MM_YYYY = False
End Function

' ValidarNumeroFactura removed - invoice numbers now accept any format

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

    Dim Y As Single
    Const MM_TO_TWIPS As Single = 56.7
    Const EXTRA_LEFT_MM As Single = 10
    Const EXTRA_BOX_HEIGHT_MM As Single = 1.25
    Const EXTRA_LEFT_TWIPS As Single = EXTRA_LEFT_MM * MM_TO_TWIPS
    Const EXTRA_BOX_HEIGHT_TWIPS As Single = EXTRA_BOX_HEIGHT_MM * MM_TO_TWIPS
    Const EXTRA_BOX_TEXT_Y_TWIPS As Single = EXTRA_BOX_HEIGHT_TWIPS / 2
    Const PRINT_X_OFFSET As Single = 660 + EXTRA_LEFT_TWIPS
    Dim xOffset As Single
    Dim xLeft As Single, xRight As Single
    Dim BoxLeft As Single, boxRight As Single
    Dim lineH As Single
    Dim logoPath As String
    Dim I As Long
    Dim detalle As String
    Dim Importe As String
    Dim deudaMax As Long
    Dim sepY As Single
    Dim headerBandTop As Single
    Dim headerBandBottom As Single
    Dim titleY As Single

    xOffset = PRINT_X_OFFSET
    xLeft = 240 + xOffset
    xRight = 7200 + xOffset
    BoxLeft = 180 + xOffset
    boxRight = 9300 + xOffset
    lineH = 320

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Arial"
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.Copies = 1

    Y = 220

    logoPath = App.Path & "\Quilplac2.jpg"
    If Dir$(logoPath) <> "" Then
        Printer.PaintPicture LoadPicture(logoPath), xLeft, Y, 2600, 900
    End If

    Printer.FontSize = 14
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y + 980
    Printer.Print "ORDEN DE PAGO"
    Printer.FontBold = False
    Printer.FontSize = 11

    Printer.CurrentX = 6200 + xOffset
    Printer.CurrentY = Y + 980
    Printer.Print "FECHA:"
    Printer.CurrentX = 7700 + xOffset
    Printer.CurrentY = Y + 980
    Printer.Print Trim$(txtFecha.text)

    Y = Y + 1600
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "Nro: " & Trim$(txtNroOrden.text)

    Y = Y + 520
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "PROVEEDOR:"
    DrawBoxConBordeGrueso 2500 + xOffset, Y - 60, boxRight, Y + 360 + EXTRA_BOX_HEIGHT_TWIPS, 30
    Printer.CurrentX = 5200 + xOffset
    Printer.CurrentY = Y + 10 + EXTRA_BOX_TEXT_Y_TWIPS
    Printer.Print UCase$(Trim$(txtProveedor.text))

    Y = Y + 900
    DrawBoxConBordeGrueso BoxLeft, Y, boxRight, Y + 340 + EXTRA_BOX_HEIGHT_TWIPS, 30
    Printer.FontBold = True
    headerBandTop = Y + 48
    headerBandBottom = Y + 236 + EXTRA_BOX_TEXT_Y_TWIPS
    
    titleY = Y + 64 + EXTRA_BOX_TEXT_Y_TWIPS
    PrintBoxHeaderTitle "DEUDA", BoxLeft, boxRight, headerBandTop, headerBandBottom, titleY
    Printer.FontBold = False

    Y = Y + 700 + EXTRA_BOX_HEIGHT_TWIPS
    sepY = Y + 120
    Printer.Line (xLeft, sepY)-(boxRight - 180, sepY), vbBlack
    Y = Y + 360

    deudaMax = grdDeuda.Rows - 1
    If deudaMax > 4 Then deudaMax = 4

    For I = 1 To deudaMax
        detalle = Trim$(grdDeuda.TextMatrix(I, 0))
        Importe = Trim$(grdDeuda.TextMatrix(I, 1))
        If Len(detalle) > 0 Then
            Printer.CurrentX = xLeft
            Printer.CurrentY = Y
            Printer.Print Left$(UCase$(detalle), 45)
        End If
        If Len(Importe) > 0 Then
            Printer.CurrentX = xRight
            Printer.CurrentY = Y
            Printer.Print "$" & FormatMoney(ParseCurrency(Importe))
        End If
        Y = Y + lineH
    Next I

    Printer.CurrentX = xLeft
    Printer.CurrentY = Y + 120
    Printer.FontBold = True
    Printer.Print "TOTAL DEUDA:"
    Printer.CurrentX = xRight
    Printer.CurrentY = Y + 120
    Printer.Print "$" & FormatMoney(ParseCurrency(txtSubDeuda.text))
    Printer.FontBold = False

    Y = Y + 760
    DrawBoxConBordeGrueso BoxLeft, Y, boxRight, Y + 340 + EXTRA_BOX_HEIGHT_TWIPS, 30
    Printer.FontBold = True
    headerBandTop = Y + 48
    headerBandBottom = Y + 236 + EXTRA_BOX_TEXT_Y_TWIPS
    
    titleY = Y + 64 + EXTRA_BOX_TEXT_Y_TWIPS
    PrintBoxHeaderTitle "PAGO", BoxLeft, boxRight, headerBandTop, headerBandBottom, titleY
    Printer.FontBold = False

    Y = Y + 700 + EXTRA_BOX_HEIGHT_TWIPS
    sepY = Y + 120
    Printer.Line (xLeft, sepY)-(boxRight - 180, sepY), vbBlack

    Y = Y + 420
    Printer.CurrentX = xLeft: Printer.CurrentY = Y: Printer.Print "CHEQUES:"
    Printer.CurrentX = xRight: Printer.CurrentY = Y: Printer.Print "$" & FormatMoney(ParseCurrency(txtSubCheques.text))
    Y = Y + 520

    Printer.CurrentX = xLeft: Printer.CurrentY = Y: Printer.Print "EFECTIVO:"
    Printer.CurrentX = xRight: Printer.CurrentY = Y: Printer.Print "$" & FormatMoney(ParseCurrency(txtEfectivo.text))
    Y = Y + 520

    Printer.CurrentX = xLeft: Printer.CurrentY = Y: Printer.Print "TRANSFERENCIA:"
    Printer.CurrentX = xRight: Printer.CurrentY = Y: Printer.Print "$" & FormatMoney(ParseCurrency(txtSubTransferencia.text))
    Y = Y + 520

    Printer.CurrentX = xLeft: Printer.CurrentY = Y: Printer.Print "FACTURAS:"
    Printer.CurrentX = xRight: Printer.CurrentY = Y: Printer.Print "$" & FormatMoney(ParseCurrency(txtSubFacturas.text))
    Y = Y + 520

    Printer.CurrentX = xLeft: Printer.CurrentY = Y: Printer.Print "OTROS:"
    Printer.CurrentX = xRight: Printer.CurrentY = Y: Printer.Print "$" & FormatMoney(ParseCurrency(txtSubOtros.text))

    Y = Y + 700
    sepY = Y + 120
    Printer.Line (xLeft, sepY)-(boxRight - 180, sepY), vbBlack

    Y = Y + 420
    Printer.FontBold = True
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "TOTAL PAGO:"
    Printer.CurrentX = xRight
    Printer.CurrentY = Y
    Printer.Print "$" & FormatMoney(ParseCurrency(txtTotalPago.text))

    Y = Y + 440
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "SALDO:"
    Printer.CurrentX = xRight
    Printer.CurrentY = Y
    Printer.Print "$" & FormatMoney(ParseCurrency(txtSaldo.text))
    Printer.FontBold = False

    Y = Y + 560
    Printer.FontSize = 10
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "IMPORTE EN LETRAS: " & UCase$(Trim$(txtImporteLetras.text))

    Y = Y + 1200
    Printer.CurrentX = xLeft
    Printer.CurrentY = Y
    Printer.Print "Firma/Aclaracion:"

    Printer.EndDoc
    Exit Sub

ErrHandler:
    MsgBox "Error al imprimir la orden de pago: " & Err.Description, vbExclamation, "Impresion"
End Sub

Private Sub PrintBoxHeaderTitle(ByVal titleText As String, ByVal BoxLeft As Single, ByVal boxRight As Single, ByVal bandTop As Single, ByVal bandBottom As Single, ByVal titleY As Single)
    Dim fillStyleAnterior As Integer
    Dim fillColorAnterior As Long
    Dim textW As Single
    Dim titleX As Single
    Dim clearLeft As Single
    Dim clearRight As Single

    fillStyleAnterior = Printer.FillStyle
    fillColorAnterior = Printer.FillColor

    textW = Printer.TextWidth(titleText)
    titleX = ((BoxLeft + boxRight) - textW) / 2

    clearLeft = titleX - 120
    clearRight = titleX + textW + 120
    If clearLeft < (BoxLeft + 20) Then clearLeft = BoxLeft + 20
    If clearRight > (boxRight - 20) Then clearRight = boxRight - 20

    Printer.FillStyle = vbFSSolid
    Printer.FillColor = vbWhite
    Printer.Line (clearLeft, bandTop)-(clearRight, bandBottom), vbWhite, BF

    Printer.CurrentX = titleX
    Printer.CurrentY = titleY
    Printer.Print titleText

    Printer.FillStyle = fillStyleAnterior
    Printer.FillColor = fillColorAnterior
End Sub

Private Sub DrawBoxConBordeGrueso(ByVal x1 As Single, ByVal y1 As Single, ByVal x2 As Single, ByVal y2 As Single, Optional ByVal grosor As Integer = 10)
    Dim grosorBorde As Single
    Dim fillStyleAnterior As Integer
    Dim fillColorAnterior As Long

    grosorBorde = grosor
    If grosorBorde < 12 Then grosorBorde = 12

    fillStyleAnterior = Printer.FillStyle
    fillColorAnterior = Printer.FillColor

    Printer.FillStyle = vbFSSolid
    Printer.FillColor = vbBlack

    ' Borde superior
    Printer.Line (x1, y1)-(x2, y1 + grosorBorde), vbBlack, BF
    ' Borde inferior
    Printer.Line (x1, y2 - grosorBorde)-(x2, y2), vbBlack, BF
    ' Borde izquierdo
    Printer.Line (x1, y1)-(x1 + grosorBorde, y2), vbBlack, BF
    ' Borde derecho
    Printer.Line (x2 - grosorBorde, y1)-(x2, y2), vbBlack, BF

    Printer.FillStyle = fillStyleAnterior
    Printer.FillColor = fillColorAnterior
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
    'SetRetIIBBText FormatMoney(NzC(rs!RetencionIIBB))
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
    Dim result As Double

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
        result = sign * Val(t)
        ' Validate range for Currency type (VB6 Currency max: 922,337,203,685.4775)
        If Abs(result) > 922337203685.478 Then
            ParseCurrency = 922337203685.478 * Sgn(result)
        Else
            ParseCurrency = CCur(result)
        End If
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

    result = sign * Val(intPart & "." & decPart)
    ' Validate range for Currency type
    If Abs(result) > 922337203685.478 Then
        ParseCurrency = 922337203685.478 * Sgn(result)
    Else
        ParseCurrency = CCur(result)
    End If
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
    Dim intPart As Currency
    Dim decPart As Long
    Dim intTxt As String

    isNeg = (v < 0)
    absV = Abs(v)

    ' Separar parte entera y decimal sin multiplicar (evita overflow)
    intPart = Fix(absV)
    decPart = CLng((absV - intPart) * 100)

    intTxt = GroupThousands(CStr(CLng(intPart)))

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
