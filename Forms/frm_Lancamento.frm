VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Lancamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Lançamentos"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11625
   Icon            =   "frm_Lancamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   11625
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   1535
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   7
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fecha&r"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar Tela"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Novo"
            Key             =   "novo"
            Object.ToolTipText     =   "Adicionar Novo"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
      EndProperty
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "Lançamentos"
         Top             =   120
         Width           =   6135
      End
   End
   Begin VB.ComboBox txt_TEF_POS 
      Height          =   315
      ItemData        =   "frm_Lancamento.frx":27A2
      Left            =   4770
      List            =   "frm_Lancamento.frx":27AC
      TabIndex        =   2
      TabStop         =   0   'False
      Text            =   "POS"
      Top             =   1920
      Width           =   675
   End
   Begin Skin_Button.ctr_Button txt_Calc 
      Height          =   615
      Left            =   8880
      TabIndex        =   10
      Top             =   5880
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Calcular"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm_Lancamento.frx":27BA
      PICN            =   "frm_Lancamento.frx":27D6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   3195
      Left            =   5505
      TabIndex        =   32
      Top             =   1200
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   5636
      _Version        =   393216
      Cols            =   7
      ScrollBars      =   2
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin MSAdodcLib.Adodc adoForma 
      Height          =   330
      Left            =   2280
      Top             =   2640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoForma"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoCartao 
      Height          =   330
      Left            =   1920
      Top             =   1920
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoCartao"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc adoLogo 
      Height          =   330
      Left            =   3720
      Top             =   1200
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "adoLogo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin rdActiveText.ActiveText txt_tx 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   556
      Alignment       =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   5
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_Logo 
      Bindings        =   "frm_Lancamento.frx":2D19
      Height          =   315
      Left            =   1680
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1200
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "usl_nome"
      BoundColumn     =   "usl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo txt_Cartao 
      Bindings        =   "frm_Lancamento.frx":2D2F
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   1920
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "Cartão"
      BoundColumn     =   "ctl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_dias_V 
      Height          =   315
      Left            =   9600
      TabIndex        =   6
      Top             =   5280
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   556
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_tx_fixo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4560
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   556
      Alignment       =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "R$ 0,00"
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_tx_po 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      Height          =   315
      Left            =   9765
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4560
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   556
      Alignment       =   1
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   5
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_NDOC 
      Height          =   315
      Left            =   1680
      TabIndex        =   7
      Top             =   5520
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   556
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   20
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_FormaPg 
      Bindings        =   "frm_Lancamento.frx":2D47
      Height          =   315
      Left            =   1680
      TabIndex        =   3
      Top             =   2640
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "fpg_desc"
      BoundColumn     =   "fpg_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo txt_banco 
      Bindings        =   "frm_Lancamento.frx":2D5E
      Height          =   315
      Left            =   1680
      TabIndex        =   24
      Top             =   3360
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      ListField       =   "bco dep"
      BoundColumn     =   "ctl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo txt_tipoC 
      Bindings        =   "frm_Lancamento.frx":2D76
      Height          =   315
      Left            =   2160
      TabIndex        =   26
      Top             =   3480
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "ctl_tipoc"
      BoundColumn     =   "ctl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_dt_vnd 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   4800
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_FormaPg_Parc 
      Bindings        =   "frm_Lancamento.frx":2D8E
      Height          =   315
      Left            =   1920
      TabIndex        =   28
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "fpg_qt_parc"
      BoundColumn     =   "fpg_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo txt_FormaPg_Tipo 
      Bindings        =   "frm_Lancamento.frx":2DA5
      Height          =   315
      Left            =   3120
      TabIndex        =   29
      Top             =   2760
      Visible         =   0   'False
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "fpg_tIPO"
      BoundColumn     =   "fpg_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_Pre 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   5
      Top             =   5280
      Visible         =   0   'False
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_LbDoc 
      Bindings        =   "frm_Lancamento.frx":2DBC
      Height          =   315
      Left            =   3000
      TabIndex        =   33
      Top             =   3480
      Visible         =   0   'False
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "ctl_label_Ndoc"
      BoundColumn     =   "ctl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_Valor_Vnd 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   7320
      TabIndex        =   9
      Top             =   6240
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   9
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_Desc_Parc 
      Bindings        =   "frm_Lancamento.frx":2DD4
      Height          =   315
      Left            =   3480
      TabIndex        =   35
      Top             =   2040
      Visible         =   0   'False
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   556
      _Version        =   393216
      ListField       =   "ctl_des_parc"
      BoundColumn     =   "ctl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_NResumo 
      Height          =   315
      Left            =   1680
      TabIndex        =   8
      Top             =   6225
      Visible         =   0   'False
      Width           =   4155
      _ExtentX        =   7329
      _ExtentY        =   556
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   40
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_Valor_Entrada 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   6360
      TabIndex        =   37
      Top             =   5280
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxLength       =   9
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.Label lbEntrada 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Entrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   38
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label lbResumo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Resumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   36
      Top             =   6285
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Compra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   34
      Top             =   6240
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Dia(s)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   10335
      TabIndex        =   31
      Top             =   5340
      Width           =   555
   End
   Begin VB.Label lbPre 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pré-Datado p/:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4920
      TabIndex        =   30
      Top             =   5400
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data Protocolo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   27
      Top             =   4875
      Width           =   1305
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Banco"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   405
      TabIndex        =   25
      Top             =   3435
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Forma de Pagamento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   405
      TabIndex        =   23
      Top             =   2565
      Width           =   1215
   End
   Begin VB.Label lbDoc 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nº Doc"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   195
      TabIndex        =   22
      Top             =   5580
      Width           =   1425
   End
   Begin VB.Label lb_tx_po 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%-Adic"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8895
      TabIndex        =   21
      Top             =   4680
      Width           =   810
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Fixo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   20
      Top             =   4680
      Width           =   975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1º Parc daqui"
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
      Left            =   8235
      TabIndex        =   19
      Top             =   5400
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   420
      TabIndex        =   18
      Top             =   4155
      Width           =   1200
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   405
      TabIndex        =   17
      Top             =   1980
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   405
      TabIndex        =   16
      Top             =   1275
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   5775
      Left            =   120
      Top             =   960
      Width           =   11055
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   8760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":2DEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":3106
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":32E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":35FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":3914
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":3C2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":3F48
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento.frx":4122
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
   Begin VB.Menu mnuSep01 
      Caption         =   "        |"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuNovo 
      Caption         =   "&Novo"
   End
   Begin VB.Menu mnuSalvar 
      Caption         =   "&Salvar"
   End
   Begin VB.Menu mnuCancelar 
      Caption         =   "&Cancelar"
   End
   Begin VB.Menu mnusep02 
      Caption         =   "|"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frm_Lancamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_ValorBruto As Double
Dim w_Total_liq As Double
Dim w_Day As String
Dim w_taxa_adesao As Double
Dim w_valor_entrada_SEM_TAXA As Double



Private Sub Form_Load()
On Error GoTo err1

    MDI.TBar.Visible = False

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 100
    
    Set adoLogo.Recordset = w_ado_Logo.Clone
    
    mnuNovo_Click
    Format_Grid
    
    If Not w_Usu_Tipo = "L" Then
        lbResumo.Visible = True
        txt_NResumo.Visible = True
    End If
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub


Function InversaDate() As String

    InversaDate = 0
    If txt_FormaPg_Tipo = "D" Then InversaDate = txt_NDOC
    
    Select Case txt_tipoC
    Case "16":
        If txt_FormaPg_Tipo = "V" Then
            InversaDate = Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        Else
            InversaDate = "4" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        End If
    Case "43":
        If txt_FormaPg_Tipo = "V" Then
            InversaDate = Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        Else
            InversaDate = "4" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        End If
    Case "5":
        If txt_FormaPg_Tipo = "V" Then
            InversaDate = "5" & Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        Else
            InversaDate = "6" & Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
        End If
    Case "6":
            InversaDate = txt_NDOC
    Case "17":
            InversaDate = Format(txt_dt_vnd, "dd-mm-yyyy")
    Case "18":
            InversaDate = Format(txt_dt_vnd, "dd-mm-yyyy")
    Case "19":
            InversaDate = Format(txt_dt_vnd, "dd-mm-yyyy")
    Case "9": 'Hipercard
            InversaDate = "H" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "7": 'Sorocred
            InversaDate = "S" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "3": 'Amex
            InversaDate = "A" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "23": 'Credsystem
            InversaDate = "C" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "24": 'Depósito 1
            InversaDate = "D" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "25": 'Depósito 2
            InversaDate = "D" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "26": 'Depósito 3
            InversaDate = "D" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "27": 'Depósito 4
            InversaDate = "D" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case "28": 'Depósito 5
            InversaDate = "D" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    Case Else
            InversaDate = "9" & Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
    End Select
    
End Function

Sub Salvar()
On Error GoTo err1

If txt_NDOC = "" And (txt_tipoC >= 24 And txt_tipoC <= 28) Then
    txt_NDOC = "0"
End If


If CDbl(txt_Valor_Vnd) > 0 And txt_NDOC <> "" And Grid.TextMatrix(1, 2) <> "" Then
    
    txt_Calc_Click
    
    'Salvar o Cabeçalho
    w_NRESUMO = InversaDate()   'IIf(txt_Desc_Parc = "N", 0, 0)
    'If w_NRESUMO = 0 And txt_NResumo <> "0" Then w_NRESUMO = txt_NResumo
    If txt_NResumo <> "0" Then w_NRESUMO = txt_NResumo
    
    strSQL = "INSERT INTO tab_lanc(lnc_ndoc, lnc_loj, lnc_tipoc, lnc_formapg, lnc_dt_vnd, lnc_tx,  " _
           & "lnc_tx_fixo, lnc_tx_po, lnc_usu, lnc_dt_lanc, lnc_nresumo,lnc_vr,lnc_tipo, lnc_vr_liq, lnc_tef_pos) " _
           & "VALUES ('" & txt_NDOC & "', '" & txt_Logo.BoundText & "', '" & txt_tipoC & "', " _
           & "'" & txt_FormaPg.BoundText & "', '" & Format(txt_dt_vnd, "YYYY-MM-DD") & "', " _
           & "'" & Replace(CDbl(Format(txt_tx, "0.000#")), ",", ".") & "', '" & Replace(CDbl(txt_tx_fixo), ",", ".") & "', " _
           & "'" & Replace(CDbl(Format(txt_tx_po, "0.000#")), ",", ".") & "', '" & w_Usu_Cod & "', now(), " _
           & "'" & w_NRESUMO & "', '" & Replace(CDbl(txt_Valor_Vnd), ",", ".") & "','" & txt_FormaPg_Tipo & "', " _
           & "'" & Replace(CDbl(w_Total_liq), ",", ".") & "', '" & txt_TEF_POS & "')"
    'add lancamento
    
    Call ExecuteSQL(strSQL, wRegAf)
    
    If wRegAf = 1 Then 'Se foi salvo com sucesso, então salvar as parcelas
        wRegAf = 0
        w_NumLanc = ExecuteSQL("Select max(lnc_num) from tab_lanc WHERE lnc_loj = '" & txt_Logo.BoundText & "'").Fields(0)
        
        'Salva as Parcelas
        For i = 1 To Grid.Rows - 1
            Grid.Row = i
            strSQL = "INSERT INTO tab_lanc_parc(lcp_ndoc, lcp_parc, lcp_dt_vcto, lcp_vr_bto, lcp_vr_liq, lcp_num_lanc, lcp_tipo, lcp_nresumo, lcp_baixa) " _
                   & "VALUES ('" & Grid.TextMatrix(i, 5) & "', " & i & ", '" & Format(Grid.TextMatrix(i, 2), "YYYY-MM-DD") & "', " _
                   & "'" & Grid.TextMatrix(i, 3) & "', '" & Grid.TextMatrix(i, 6) & "', '" & w_NumLanc & "', " _
                   & "'" & txt_FormaPg_Tipo & "', '" & w_NRESUMO & "','0000-00-00')"
            'add parcelas
            Call ExecuteSQL(strSQL, wRegAf)
            If wRegAf = 0 Then MsgBox "Falha ao salvar a parcela " & i & " !", vbCritical
            
        Next i
        
        
        MsgBox "Registro Salvo com sucesso!", vbInformation
        mnuCancelar_Click
    Else
        MsgBox "Não foi pssível salvar!", vbCritical
    End If
Else
    If CDbl(txt_Valor_Vnd) = 0 Or txt_Valor_Vnd = "" Then
        MsgBox "Preencha o valor da compra!", vbExclamation
    ElseIf txt_NDOC = "" Then
        MsgBox "Preencha o Nº Doc.!", vbExclamation
    ElseIf Grid.TextMatrix(1, 2) = "" Then
        MsgBox "Clique em calcular parcelas!", vbExclamation
    End If
End If
    
    
sair:
    Exit Sub
err1:
On Error Resume Next
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub






Private Sub mnuCancelar_Click()
On Error GoTo err1

        mnuNovo_Click
        
sair:
    Exit Sub
err1:
    On Error Resume Next
    If err.Number <> 458 Then MsgBox msgErro(err), vbCritical
    If adoReg.Recordset.RecordCount > 0 Then adoReg.Recordset.MoveFirst
    Resume sair
End Sub

Private Sub mnuFechar_Click()
    Unload Me
End Sub

Private Sub mnuNovo_Click()
On Error GoTo err1
    
    txt_Logo.Enabled = True
    txt_dt_vnd = w_Data_Server
    txt_Cartao = ""
    txt_NDOC = ""
    txt_banco = ""
    txt_FormaPg = ""
    txt_dias_V = "0"
    txt_tx = "0"
    txt_tx_fixo = "0"
    txt_tx_po = "0"
    txt_Valor_Vnd = "0"
    txt_NResumo = "0"
    txt_TEF_POS = "POS"

On Error Resume Next
    
    If w_Usu_Tipo = "L" Then
        txt_Logo.BoundText = w_Usu_Cod
        txt_Logo.Enabled = False
        
        txt_Cartao.SetFocus
    Else
        txt_Logo.SetFocus
    End If

    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuSalvar_Click()
    Salvar
End Sub



Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    Case "novo": mnuNovo_Click
    Case "salvar": Salvar
    Case "cancelar": mnuCancelar_Click
    End Select
End Sub

Private Sub txt_bco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Sub txt_Calc_Click()
   
    Format_Grid
    
    'Se for Sorocred Ficha com entrada
    'If txt_tipoC = "6" And Left(txt_FormaPg, 2) = "0+" Then
    If Left(txt_FormaPg, 2) = "0+" Then
        'Não digitou entrada
        If (CDbl(txt_Valor_Entrada) <= 0) Then
            MsgBox "O valor de entrada não pode ser zero nessa condição de pagamento.", vbCritical + vbOKOnly, "Valor de entrada incorreto"
            txt_Valor_Entrada.SetFocus
            Exit Sub
        End If
        'Entrada > que venda
        If CDbl(txt_Valor_Entrada) >= CDbl(txt_Valor_Vnd) Then
            MsgBox "O valor da ENTRADA está incorreto. Você digitou errado, digite novamente!", vbCritical + vbOKOnly, "Valor de entrada incorreto"
            txt_Valor_Entrada.SetFocus
            Exit Sub
        End If
    End If
   
   'Valor digitado errado (consistência MENOR QUE 20 mil)
    If txt_Valor_Vnd >= 20000 Then
        MsgBox "O valor digitado é muito alto. Você digitou errado!", vbCritical + vbOKOnly, "Valor incorreto"
        txt_Valor_Vnd.SetFocus
        Exit Sub
    End If

    
    
    Calc_Grid
End Sub

Private Sub txt_Cartao_Change()
On Error GoTo err1

    txt_banco.BoundText = txt_Cartao.BoundText
    txt_tipoC.BoundText = txt_Cartao.BoundText
    txt_Desc_Parc.BoundText = txt_Cartao.BoundText
    txt_LbDoc.BoundText = txt_Cartao.BoundText
    lbDoc.Caption = txt_LbDoc
    
    txt_FormaPg = ""
    txt_dias_V = "0"
    txt_tx = "0"
    txt_tx_fixo = "0"
    txt_tx_po = "0"
    
    Set adoForma.Recordset = ExecuteSQL("SELECT tab_forma_pg.fpg_cod, tab_forma_pg.fpg_desc, tab_forma_pg.fpg_qt_parc, tab_forma_pg.fpg_Tipo FROM tab_forma_pg, tab_tipo_forma WHERE (tab_forma_pg.fpg_cod = tab_tipo_forma.fpg_cod) AND (tab_tipo_forma.tpc_cod = '" & txt_tipoC & "') ORDER BY tab_forma_pg.fpg_desc").Clone


    If txt_tipoC = "7" Or txt_tipoC = "6" Or txt_tipoC = "29" Or txt_tipoC = "31" Or txt_tipoC = "39" Then
        lbResumo.Caption = "Cliente"
        lbResumo.Visible = True
        txt_NResumo.Visible = True
    Else
        lbResumo.Caption = "Nº Resumo"
        If Not w_Usu_Tipo = "L" Then
            lbResumo.Visible = True
            txt_NResumo.Visible = True
        Else
            lbResumo.Visible = False
            txt_NResumo.Visible = False
        End If
        
    End If
    

    


    Format_Grid

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub txt_Cartao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub




Private Sub txt_dias_V_Validate(Cancel As Boolean)
    Format_Grid
    txt_Pre = CVDate(txt_dt_vnd) + CDbl(txt_dias_V)
End Sub

Private Sub txt_dt_vnd_Validate(Cancel As Boolean)
On Error GoTo err1

    Format_Grid

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub txt_FormaPg_Change()
    txt_FormaPg_Parc.BoundText = txt_FormaPg.BoundText
    txt_FormaPg_Tipo.BoundText = txt_FormaPg.BoundText
    
    adoCartao.Recordset.MoveFirst
    adoCartao.Recordset.Find "ctl_cod = " & IIf(txt_tipoC.BoundText = "", 0, txt_tipoC.BoundText) & ""
    
    If txt_FormaPg_Parc = "1" And txt_FormaPg_Tipo = "V" And Not adoCartao.Recordset.EOF Then
        txt_tx = Format(adoCartao.Recordset.Fields(3), "0.00%")
        txt_tx_fixo = adoCartao.Recordset.Fields(5)
        txt_dias_V = adoCartao.Recordset.Fields(4)
        txt_tx_po = Format(0, "0.00%")
        txt_tx_po.Visible = False
        lb_tx_po.Visible = False
    ElseIf Not adoCartao.Recordset.EOF Then
    
        taxaAlta_qt_parcelas = adoCartao.Recordset.Fields(15)
        qtParcelas = IIf(txt_FormaPg_Parc = "", 1, CInt(txt_FormaPg_Parc))
    
        ' Se qt de parcelas selecionada >= qt de parcelas com taxa de retenção maior, então pega taxa maior, senão pega taxa a prazo normal
        If qtParcelas >= taxaAlta_qt_parcelas And adoCartao.Recordset.Fields(15) > 0 And adoCartao.Recordset.Fields(16) > 0 Then
            txt_tx = Format(adoCartao.Recordset.Fields(16), "0.00%")
        Else
            txt_tx = Format(adoCartao.Recordset.Fields(6), "0.00%")
        End If
        txt_tx_fixo = adoCartao.Recordset.Fields(8)
        txt_dias_V = adoCartao.Recordset.Fields(7)
        txt_tx_po = Format(adoCartao.Recordset.Fields(9), "0.00%")
        txt_tx_po.Visible = True
        lb_tx_po.Visible = True
    End If
    If txt_FormaPg_Tipo = "D" Then
        lbPre.Visible = True
        txt_Pre.Visible = True
        txt_dias_V.Enabled = True
    Else
        lbPre.Visible = False
        txt_Pre.Visible = False
        txt_dias_V.Enabled = False
    End If
    If txt_tipoC = "6" And Left(txt_FormaPg, 2) = "0+" Then 'Sorocred Ficha com entrada
        lbEntrada.Visible = True
        txt_Valor_Entrada.Visible = True
    Else
        lbEntrada.Visible = False
        txt_Valor_Entrada.Visible = False
    End If
    
    Format_Grid
    txt_dt_vnd_Validate False
    
End Sub



Private Sub txt_FormaPg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub



Private Sub txt_Logo_Change()
    w_Usu = IIf(w_Usu_Tipo = "L", w_Usu_Nome, "%")
    If w_Usu = "%" Then
        Set adoCartao.Recordset = ExecuteSQL("SELECT tab_cartao_loja.ctl_cod, tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartão, tab_cartao_loja.ctl_txv AS `%-Vista`, tab_cartao_loja.ctl_dias_v AS `Dias-V`, tab_cartao_loja.ctl_vr_des_v AS `Vr Desc - V`, tab_cartao_loja.ctl_txp AS `%-Prazo`, tab_cartao_loja.ctl_dias_p AS `Dias-Pz`, tab_cartao_loja.ctl_vr_des_p AS `Vr Desc - Pz`, tab_cartao_loja.ctl_vr_po AS `%-Pz Adic`, tab_banco.bco_desc AS `Bco Dep`, tab_cartao_loja.ctl_loja, tab_cartao_loja.ctl_tipoc, tab_cartao_loja.ctl_label_ndoc, tab_cartao_loja.ctl_des_parc, tab_cartao_loja.ctl_parc_alta_qt, tab_cartao_loja.ctl_parc_alta_tx FROM tab_tipo_cartao, tab_usuario, { oj tab_cartao_loja LEFT OUTER JOIN tab_banco ON tab_cartao_loja.ctl_banco = tab_banco.bco_cod } WHERE (tab_cartao_loja.ctl_loja = tab_usuario.usl_cod) AND (tab_cartao_loja.ctl_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_usuario.usl_nome = '" & txt_Logo & "') ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc").Clone
    Else
        Set adoCartao.Recordset = w_ado_Cartao.Clone
    End If
End Sub

Private Sub txt_Logo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


'Procedimento q/ Formata o Grid
Sub Format_Grid()
On Error GoTo err1
    Grid.Rows = 0
    'Caption da Colunas
    Grid.Rows = IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc) + 1
    Grid.FixedRows = 1
    Grid.TextArray(1) = "Parc."
    Grid.TextArray(2) = "Dt. Vcto"
    Grid.TextArray(3) = "Vr Bruto"
    Grid.TextArray(4) = "Vr. Liq"
    Grid.TextArray(5) = "Num. Doc."
    Grid.TextArray(6) = "Vr. Liq com todas as casas decimais"
    
    'Formata a Largura das colunas
    Grid.ColWidth(0) = 250
    Grid.ColWidth(1) = 550
    Grid.ColWidth(2) = 1000
    Grid.ColWidth(3) = 1000
    Grid.ColWidth(4) = 1000
    Grid.ColWidth(5) = 1430
    Grid.ColWidth(6) = 0
    
    'Formata Alinhamento do Texto
    Grid.ColAlignment(1) = 6
    Grid.ColAlignment(2) = 4
    Grid.ColAlignment(3) = 6
    Grid.ColAlignment(4) = 6
    Grid.ColAlignment(5) = 6

    Grid.ColAlignmentFixed(1) = 6
    Grid.ColAlignmentFixed(2) = 4
    Grid.ColAlignmentFixed(3) = 6
    Grid.ColAlignmentFixed(4) = 6
    Grid.ColAlignmentFixed(5) = 6
    
    Grid.Row = 0
    For c = 1 To 5
        Grid.Col = c
        Grid.CellFontBold = True
    Next c
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Public Function ProximoDiaUtil(data As Variant) As Variant
    Dim dayOfWeek, dataConvertida
    
    dataConvertida = CVDate(data)
    dayOfWeek = Weekday(data)
    
    If (dayOfWeek = vbSaturday) Then
        ProximoDiaUtil = dataConvertida + 2
    ElseIf (dayOfWeek = vbSunday) Then
        ProximoDiaUtil = dataConvertida + 1
    Else
        ProximoDiaUtil = dataConvertida
    End If
End Function



'Procedimento :  Calcula as Parcelas e coloca no grid
Sub Calc_Grid()
Dim wDt As Date
Dim v_QtParcRest As Byte


On Error GoTo err1
    'Caption da Colunas
    If (txt_tipoC >= 17 And txt_tipoC <= 19) Then
        If (Day(txt_dt_vnd) >= 24 And Day(txt_dt_vnd) <= 31) Or (Day(txt_dt_vnd) >= 1 And Day(txt_dt_vnd) <= 3) Then
            w_Dt = "05/" & Format(CVDate(txt_dt_vnd) + 45, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 4 And Day(txt_dt_vnd) <= 13) Then
            w_Dt = "15/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 14 And Day(txt_dt_vnd) <= 23) Then
            w_Dt = "25/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        End If
    ElseIf txt_tipoC = 22 Then
        If (Day(txt_dt_vnd) >= 20 And Day(txt_dt_vnd) <= 31) Then
            w_Dt = "10/" & Format(CVDate(txt_dt_vnd) + 20, "mm/yyyy")
        Else
            w_Dt = "10/" & Format(CVDate(txt_dt_vnd) + 5, "mm/yyyy")
        End If
    ElseIf txt_tipoC = 23 Then 'Credsystem
        If (Day(txt_dt_vnd) >= 1 And Day(txt_dt_vnd) <= 7) Then
            w_Dt = "07/" & Format(CVDate(txt_dt_vnd) + 45, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 8 And Day(txt_dt_vnd) <= 15) Then
            w_Dt = "15/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 16 And Day(txt_dt_vnd) <= 22) Then
            w_Dt = "22/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 23 And Day(txt_dt_vnd) <= 31) Then
            w_Dt = "30/" & Format(CVDate(txt_dt_vnd) + 19, "mm/yyyy")
        End If
    ElseIf txt_tipoC = 31 Then 'Brasil Card
        If (Day(txt_dt_vnd) >= 1 And Day(txt_dt_vnd) <= 6) Then
            w_Dt = "13/" & Format(CVDate(txt_dt_vnd) + 45, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 7 And Day(txt_dt_vnd) <= 12) Then
            w_Dt = "20/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 13 And Day(txt_dt_vnd) <= 18) Then
            w_Dt = "26/" & Format(CVDate(txt_dt_vnd) + 32, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 19 And Day(txt_dt_vnd) <= 24) Then
            w_Dt = "03/" & Format(CVDate(txt_dt_vnd) + 45, "mm/yyyy")
        ElseIf (Day(txt_dt_vnd) >= 25 And Day(txt_dt_vnd) <= 31) Then
            w_Dt = "08/" & Format(CVDate(txt_dt_vnd) + 45, "mm/yyyy")
        End If
    Else
        w_Dt = CVDate(txt_dt_vnd)
    End If
    
    w_Dt = CVDate(w_Dt)
    
If CDbl(txt_Valor_Vnd) > 0 And txt_Valor_Vnd <> "" Then

    w_ValorBruto = CDbl(txt_Valor_Vnd)
    w_Total_liq = 0
    w_DT_Calc = CVDate(w_Dt)
    
    
    'For entre as Linhas - Parcelas do Grid
    For i = 1 To IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc)
        'Qtde de Parcelas q/ Resta
        v_QtParcRest = CDbl(txt_FormaPg_Parc) - (i - 1)
        
        Grid.TextMatrix(i, 1) = i & "º"   'Identifica a Parcela
                
        'se não for MaxCred
        If Not (txt_tipoC = "8") Then
            'Senão For DBCRED
            If Not (txt_tipoC >= 17 And txt_tipoC <= 19) Then
                
                If (txt_tipoC = 23) Then
                
                    If (Day(txt_dt_vnd) >= 1 And Day(txt_dt_vnd) <= 7) Then
                        w_Day = "07/"
                    ElseIf (Day(txt_dt_vnd) >= 8 And Day(txt_dt_vnd) <= 15) Then
                        w_Day = "15/"
                    ElseIf (Day(txt_dt_vnd) >= 16 And Day(txt_dt_vnd) <= 22) Then
                        w_Day = "22/"
                    ElseIf (Day(txt_dt_vnd) >= 23 And Day(txt_dt_vnd) <= 31) Then
                        w_Day = "30/"
                    End If
                
                    If i = 1 Then w_DT_Calc = CVDate(w_DT_Calc - 31)
                    w_DT_Calc = CVDate(w_DT_Calc + 31)
                    w_DT_Calc = CVDate(w_Day & Format(w_DT_Calc, "mm/yyyy"))
                    Grid.TextMatrix(i, 2) = ProximoDiaUtil(w_DT_Calc)
                                    
                Else
                
                'If Not (txt_tipoC = 7) Then 'se não for Sorocred
                    If txt_FormaPg_Tipo <> "D" Then   'V  ou  P
                        'Inserir a Data do Vcto da Parcela
                       If txt_dias_V = 30 Or txt_dias_V = 31 Then
                          'w_Dt = DateAdd("m", 1, w_Dt)
                          w_Dt = DateAdd("m", i, w_DT_Calc)
                       Else
                          w_Dt = w_Dt + CDbl(txt_dias_V)
                       End If
                        Grid.TextMatrix(i, 2) = ProximoDiaUtil(w_Dt)
                    Else 'Pre Datado
                        Grid.TextMatrix(i, 2) = ProximoDiaUtil(txt_Pre)
                    End If
                    
                    If (txt_tipoC = 22) Then
                        w_DT_Calc = CVDate(w_DT_Calc + 31)
                        w_DT_Calc = CVDate("10/" & Format(w_DT_Calc, "mm/yyyy"))
                        Grid.TextMatrix(i, 2) = ProximoDiaUtil(w_DT_Calc)
                    End If
                 'End If
                
            End If
            
            Else 'Se for DbCred
                Grid.TextMatrix(i, 2) = ProximoDiaUtil(w_Dt)
                w_Dt = Format(Day(w_Dt), "00") & "/" & Format(CVDate(w_Dt) + 32, "mm/yyyy")
            End If
            
        Else
            
                If Day(w_Dt) <= 20 Then
                    'Inserir a Data do Vcto da Parcela
                    w_DT_Calc = CVDate(w_Dt + 31)
                    w_DT_Calc = CVDate("05/" & Format(w_DT_Calc, "mm/yyyy"))
                Else
                    'Inserir a Data do Vcto da Parcela
                    w_DT_Calc = CVDate(w_Dt + 62)
                    w_DT_Calc = CVDate("05/" & Format(w_DT_Calc, "mm/yyyy"))
                End If
                w_Dt = w_DT_Calc
                Grid.TextMatrix(i, 2) = ProximoDiaUtil(w_Dt)
            
        End If
        
        'Inserir o Valor Bruto da Parcela
        If txt_tipoC = "6" And Left(txt_FormaPg, 2) = "0+" Then 'Se for Sorocred Ficha com entrada, somar com entrada
            w_taxa_adesao = 9.99
            w_valor_entrada_SEM_TAXA = txt_Valor_Entrada - w_taxa_adesao
            If i = 1 Then 'Se parcela 1
                w_taxa_adesao = 9.99
                Grid.TextMatrix(i, 3) = Format(((txt_Valor_Vnd - w_valor_entrada_SEM_TAXA) / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc) + w_valor_entrada_SEM_TAXA), "0.00")
            Else
                Grid.TextMatrix(i, 3) = Format(((txt_Valor_Vnd - w_valor_entrada_SEM_TAXA) / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc)), "0.00")
            End If
        Else
            Grid.TextMatrix(i, 3) = Format(txt_Valor_Vnd / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc), "0.00")
        End If
        Grid.TextMatrix(i, 3) = Replace(Grid.TextMatrix(i, 3), ",", ".")
        v_ValorCalc = Calc_ParcLiq(v_QtParcRest, i)
        
        'Inserir o Valor Liquido da Parcela -  Com as Retenções
        Grid.TextMatrix(i, 4) = Format(v_ValorCalc, "0.00")
        Grid.TextMatrix(i, 4) = Replace(Grid.TextMatrix(i, 4), ",", ".")

        'Inserir o Valor Liquido da Parcela -  Com as Retenções   na Com todas as Casas Decimais
        Grid.TextMatrix(i, 6) = v_ValorCalc
        Grid.TextMatrix(i, 6) = Replace(Grid.TextMatrix(i, 6), ",", ".")
       
        w_Total_liq = w_Total_liq + v_ValorCalc
        
        If txt_Desc_Parc = "N" Then
            'Inserir o Nº DOC
            Grid.TextMatrix(i, 5) = txt_NDOC
        Else
            'Inserir o Nº DOC
            Grid.TextMatrix(i, 5) = txt_NDOC + (i - 1)
        End If
        
    Next i
    
    w_ValorBruto = 0
    
Else
    MsgBox "Insira o Valor da venda e verifique se todos os campos estão preenchidos!", vbCritical
End If
    
sair:
    Exit Sub
err1:
    If err.Number = 13 Then
        MsgBox "Preencha todos os campos!", vbCritical
    Else
        MsgBox msgErro(err), vbCritical
    End If
    Resume sair
End Sub


Private Function Calc_ParcLiq(v_QtParc As Byte, Optional i)
Dim W_ValorL, w_ValorParcB, w_ValorParcL As Double

On Error GoTo err1

If IsMissing(i) Then i = 0

    If txt_Desc_Parc = "N" Then  'Desconto sobre o Total
    
        w_ValorParcB = txt_Valor_Vnd / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc)
        W_ValorL = txt_Valor_Vnd
        W_ValorL = W_ValorL - (W_ValorL * CDbl(Format(txt_tx, "0.000#")))
        W_ValorL = W_ValorL - CDbl(txt_tx_fixo)
        W_ValorL = W_ValorL - (W_ValorL * CDbl(Format(txt_tx_po, "0.000#")))
        If txt_tipoC = "6" And Left(txt_FormaPg, 2) = "0+" Then
            If i = 1 Then
                w_ValorParcL = (W_ValorL - (W_ValorL * (w_valor_entrada_SEM_TAXA / txt_Valor_Vnd))) / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc) + (W_ValorL * (w_valor_entrada_SEM_TAXA / txt_Valor_Vnd))
            Else
                w_ValorParcL = (W_ValorL - (W_ValorL * (w_valor_entrada_SEM_TAXA / txt_Valor_Vnd))) / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc)
            End If
        Else
            w_ValorParcL = W_ValorL / IIf(txt_FormaPg_Parc = "", 1, txt_FormaPg_Parc)
        End If
        
    Else '*** Desconto por Parcela
        
        w_ValorParcL = CDbl(w_ValorBruto) - (w_ValorBruto * CDbl(Format(txt_tx, "0.000#")))
        w_ValorParcL = w_ValorParcL - CDbl(txt_tx_fixo)
        w_ValorParcL = w_ValorParcL - (w_ValorParcL * CDbl(Format(txt_tx_po, "0.000#")))
        w_ValorBruto = w_ValorParcL
        w_ValorParcL = w_ValorBruto / v_QtParc
        w_ValorBruto = w_ValorBruto - w_ValorParcL
    End If
    
    Calc_ParcLiq = w_ValorParcL

sair:
    Exit Function
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Function


Private Sub txt_NDOC_Validate(Cancel As Boolean)
    Format_Grid
End Sub


Private Sub txt_Pre_Validate(Cancel As Boolean)
    If txt_Pre = "" Then
err:
        txt_Pre = ""
        MsgBox "A data pré-datada deve ser maior que a Data da compra!", vbCritical
        Cancel = True
        'txt_Pre.SetFocus
    ElseIf txt_Pre <> "" And CVDate(txt_Pre) > CVDate(txt_dt_vnd) Then
        txt_dias_V = CVDate(txt_Pre) - CVDate(txt_dt_vnd)
        Format_Grid
    Else
        GoTo err
    End If
End Sub


Private Sub txt_TEF_POS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub txt_Valor_Entrada_LostFocus()
    'Se for Sorocred Ficha com entrada e não tiver digitado a entrada
    If Not (txt_Valor_Entrada > "0") Then MsgBox "O valor de entrada não pode ser zero nessa condição de pagamento.", vbCritical + vbOKOnly, "Valor de entrada incorreto"
End Sub

Private Sub txt_Valor_Vnd_KeyDown(KeyCode As Integer, Shift As Integer)
Dim w_Pos As Byte

On Error GoTo err1

    If KeyCode = 13 Then
        w_Pos = 1
        txt_Calc_Click
        If txt_Logo.Enabled = False Then
            w_Pos = 2
            If txt_Cartao.Enabled = True Then txt_Cartao.SetFocus
        Else
            w_Pos = 3
            If txt_Logo.Enabled = True Then txt_Logo.SetFocus
        End If
        'w_Pos = 4
        'If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            w_Pos = 5
            'Salvar
            txt_Calc_Click
        'End If
    End If
    
sair:
    Exit Sub
err1:
    MsgBox w_Pos & " - " & err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub


