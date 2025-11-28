VERSION 5.00
Begin VB.Form FormChequesBapro 
   Caption         =   "Imprimir Cheque BANCO PROVINCIA Bs. As."
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4440
   ScaleWidth      =   10665
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   10335
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   6120
         TabIndex        =   19
         Top             =   3240
         Width           =   1695
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   615
         Left            =   2520
         TabIndex        =   18
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtImporte 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7080
         TabIndex        =   0
         Top             =   360
         Width           =   2775
      End
      Begin VB.TextBox txtEnLetras 
         Height          =   405
         Left            =   2760
         TabIndex        =   17
         Top             =   2400
         Width           =   7095
      End
      Begin VB.TextBox txtPagueseA 
         Height          =   405
         Left            =   1560
         TabIndex        =   15
         Top             =   1680
         Width           =   6855
      End
      Begin VB.ComboBox Combo6 
         Height          =   315
         Left            =   4320
         TabIndex        =   13
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         Left            =   1920
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.ComboBox Combo4 
         Height          =   315
         Left            =   720
         TabIndex        =   10
         Text            =   "Combo1"
         Top             =   1080
         Width           =   735
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   5640
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   570
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   3240
         TabIndex        =   2
         Text            =   "Combo1"
         Top             =   570
         Width           =   1935
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2040
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   570
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "LA CANTIDAD DE PESOS"
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
         TabIndex        =   16
         Top             =   2430
         Width           =   2235
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "PAGUESE A "
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
         TabIndex        =   14
         Top             =   1710
         Width           =   1125
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DE"
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
         Left            =   3960
         TabIndex        =   12
         Top             =   1110
         Width           =   270
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "DE"
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
         Left            =   1560
         TabIndex        =   9
         Top             =   1110
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "EL"
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
         TabIndex        =   8
         Top             =   1110
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "de"
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
         Left            =   5280
         TabIndex        =   7
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "de"
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
         Left            =   2880
         TabIndex        =   6
         Top             =   600
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Quilmes Oeste,"
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
         Left            =   600
         TabIndex        =   5
         Top             =   600
         Width           =   1290
      End
   End
End
Attribute VB_Name = "FormChequesBapro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function EnLetras(numero As String) As String
    
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        If Mid(numero, paso, 1) = "." Then
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
                MsgBox (Mid(entero, 1, 1))
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
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
Private Sub Text3_Change()

End Sub


Private Sub Text3_LostFocus()

    

End Sub


Private Sub txtImporte_LostFocus()

    txtEnLetras.Text = EnLetras(Trim(Str(txtImporte.Text)))

End Sub


