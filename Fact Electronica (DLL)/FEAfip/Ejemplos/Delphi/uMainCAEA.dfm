object Form3: TForm3
  Left = 0
  Top = 0
  Caption = 'Form3'
  ClientHeight = 357
  ClientWidth = 569
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 184
    Top = 24
    Width = 36
    Height = 13
    Caption = 'Per'#237'odo'
  end
  object Label2: TLabel
    Left = 184
    Top = 53
    Width = 30
    Height = 13
    Caption = 'Orden'
  end
  object btSolicitar: TButton
    Left = 40
    Top = 48
    Width = 75
    Height = 25
    Caption = 'Solicitar CAE'
    TabOrder = 0
    OnClick = btSolicitarClick
  end
  object btConsultar: TButton
    Left = 424
    Top = 48
    Width = 75
    Height = 25
    Caption = 'Consultar CAE'
    TabOrder = 1
    OnClick = btConsultarClick
  end
  object btInformar: TButton
    Left = 192
    Top = 280
    Width = 137
    Height = 25
    Caption = 'Informar Comprobante'
    TabOrder = 2
    OnClick = btInformarClick
  end
  object GroupBox1: TGroupBox
    Left = 192
    Top = 128
    Width = 185
    Height = 105
    Caption = 'CAE'
    TabOrder = 3
    object lbCAE: TLabel
      Left = 24
      Top = 48
      Width = 3
      Height = 13
    end
  end
  object edPeriodo: TEdit
    Left = 226
    Top = 23
    Width = 121
    Height = 21
    TabOrder = 4
    Text = '201506'
  end
  object edOrden: TEdit
    Left = 226
    Top = 50
    Width = 121
    Height = 21
    TabOrder = 5
    Text = '1'
  end
end
