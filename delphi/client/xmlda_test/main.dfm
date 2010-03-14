object Form1: TForm1
  Left = 296
  Top = 214
  Width = 586
  Height = 535
  Caption = 'OPC XML DA Core Sample'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -13
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  KeyPreview = True
  OldCreateOrder = False
  Position = poDesktopCenter
  Scaled = False
  OnKeyDown = FormKeyDown
  OnShow = FormShow
  PixelsPerInch = 120
  TextHeight = 16
  object Splitter1: TSplitter
    Left = 0
    Top = 257
    Width = 578
    Height = 8
    Cursor = crVSplit
    Align = alTop
  end
  object Panel1: TPanel
    Left = 0
    Top = 0
    Width = 578
    Height = 57
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 0
    object Label1: TLabel
      Left = 0
      Top = 8
      Width = 104
      Height = 16
      Caption = 'OPC Server URL '
    end
    object Label4: TLabel
      Left = 0
      Top = 32
      Width = 122
      Height = 16
      Caption = 'OPC XML Command'
    end
    object eURL: TEdit
      Left = 176
      Top = 4
      Width = 369
      Height = 24
      TabOrder = 0
      Text = 'http://localhost/opcxmlda/matrikon.opc.simulation.1'
    end
    object bGo: TButton
      Left = 488
      Top = 28
      Width = 57
      Height = 25
      Caption = 'F9 = Go'
      TabOrder = 1
      OnClick = bGoClick
    end
    object cXMLCommand: TComboBox
      Left = 176
      Top = 28
      Width = 305
      Height = 24
      Style = csDropDownList
      ItemHeight = 16
      TabOrder = 2
      OnChange = cXMLCommandChange
      Items.Strings = (
        'GetStatus'
        'Browse'
        'Read'
        'Write'
        'GetProperties'
        'Subscribe'
        'SubscriptionPolledRefresh'
        'SubscriptionCancel')
    end
  end
  object WebBrowser1: TWebBrowser
    Left = 0
    Top = 282
    Width = 578
    Height = 225
    Align = alClient
    DragMode = dmAutomatic
    TabOrder = 1
    ControlData = {
      4C000000CA2F00009B1200000000000000000000000000000000000000000000
      000000004C000000000000000000000001000000E0D057007335CF11AE690800
      2B2E126203000000000000004C0000000114020000000000C000000000000046
      8000000000000000000000000000000000000000000000000000000000000000
      00000000000000000100000000000000000000000000000000000000}
  end
  object Panel2: TPanel
    Left = 0
    Top = 265
    Width = 578
    Height = 17
    Align = alTop
    BevelOuter = bvNone
    TabOrder = 2
    object Label3: TLabel
      Left = 0
      Top = 0
      Width = 120
      Height = 16
      Caption = 'OPC XML DA Output'
    end
  end
  object Memo1: TMemo
    Left = 0
    Top = 57
    Width = 578
    Height = 200
    Align = alTop
    BorderStyle = bsNone
    ScrollBars = ssBoth
    TabOrder = 3
  end
end
