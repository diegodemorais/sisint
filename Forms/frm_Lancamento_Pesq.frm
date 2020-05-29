VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Lancamento_Pesq 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Lançamentos"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11670
   Icon            =   "frm_Lancamento_Pesq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   11670
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   1535
      ButtonWidth     =   1535
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
            Caption         =   "E&xcluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir o Lançamento"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Alterar"
            Key             =   "alterar"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Alt Todos"
            Key             =   "alterarTodos"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   10680
         Picture         =   "frm_Lancamento_Pesq.frx":27A2
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   44
         Top             =   120
         Width           =   480
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
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
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "Consulta de Lançamentos"
         Top             =   120
         Width           =   6615
      End
   End
   Begin VB.TextBox txt2 
      Height          =   495
      Left            =   4200
      TabIndex        =   46
      Text            =   "Text2"
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.TextBox txt 
      Height          =   495
      Left            =   3000
      TabIndex        =   45
      Text            =   "Text2"
      Top             =   0
      Width           =   7215
   End
   Begin VB.PictureBox pic_Pesq 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   6960
      Picture         =   "frm_Lancamento_Pesq.frx":2BE4
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   43
      Top             =   1050
      Visible         =   0   'False
      Width           =   480
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frm_Lancamento_Pesq.frx":2EEE
      Height          =   1815
      Left            =   120
      TabIndex        =   42
      Top             =   2280
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   3201
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   11
      BeginProperty Column00 
         DataField       =   "TEF_POS"
         Caption         =   "T/P"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "lnc_nresumo"
         Caption         =   "Nº Resumo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "lnc_ndoc"
         Caption         =   "Nº Doc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Logo"
         Caption         =   "Logo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "Cartao"
         Caption         =   "Cartão"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "FormaPG"
         Caption         =   "Forma Pg."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "lnc_vr"
         Caption         =   "Vr Compra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "lnc_tx"
         Caption         =   "%"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "lnc_tx_fixo"
         Caption         =   "Vr Fixo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "lnc_tx_po"
         Caption         =   "% Adic"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0,00%"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   5
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "lnc_dt_vnd"
         Caption         =   "Dt. Prot"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column05 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column06 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column07 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoGrid 
      Height          =   375
      Left            =   6600
      Top             =   5520
      Visible         =   0   'False
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   1
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
      Caption         =   "adoGrid"
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
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   6585
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   661
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
      MaxRecords      =   2
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
      Caption         =   "Registros: 0 de 0"
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
   Begin VB.Frame frFiltro 
      Caption         =   " Filtros "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   35
      Top             =   840
      Width           =   11415
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   1
         Left            =   9600
         Top             =   240
      End
      Begin MSAdodcLib.Adodc adoCartao 
         Height          =   330
         Left            =   3840
         Top             =   960
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
      Begin VB.CheckBox ck 
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
         Height          =   330
         Index           =   4
         Left            =   270
         TabIndex        =   4
         Top             =   1050
         Width           =   1575
      End
      Begin VB.CheckBox ck 
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
         Index           =   3
         Left            =   270
         TabIndex        =   3
         Top             =   870
         Width           =   1575
      End
      Begin VB.CheckBox ck 
         Caption         =   "Data Lanç."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   270
         TabIndex        =   2
         Top             =   630
         Width           =   1575
      End
      Begin VB.CheckBox ck 
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
         Height          =   270
         Index           =   1
         Left            =   270
         TabIndex        =   1
         Top             =   420
         Width           =   1275
      End
      Begin VB.CheckBox ck 
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
         Height          =   270
         Index           =   0
         Left            =   270
         TabIndex        =   0
         Top             =   210
         Width           =   1575
      End
      Begin rdActiveText.ActiveText Txt_DtI 
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
         Left            =   5820
         TabIndex        =   7
         Top             =   1005
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText Txt_DtF 
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
         Left            =   7185
         TabIndex        =   8
         Top             =   1005
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
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
         MaxLength       =   10
         TextMask        =   1
         RawText         =   1
         Mask            =   "##/##/####"
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin MSDataListLib.DataCombo txt_Logo_P 
         Bindings        =   "frm_Lancamento_Pesq.frx":2F03
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "usl_nome"
         BoundColumn     =   "usl_cod"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo txt_Cartao_P 
         Bindings        =   "frm_Lancamento_Pesq.frx":2F19
         Height          =   315
         Left            =   2280
         TabIndex        =   6
         Top             =   1005
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tpc_desc"
         BoundColumn     =   "tpc_cod"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin rdActiveText.ActiveText txt_NDOC_P 
         Height          =   315
         Left            =   8625
         TabIndex        =   9
         Top             =   420
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         MaxLength       =   20
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_NResumo_P 
         Height          =   315
         Left            =   8625
         TabIndex        =   10
         Top             =   1005
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
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
         MaxLength       =   20
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin Skin_Button.ctr_Button bt_Pesq 
         Height          =   1080
         Left            =   10110
         TabIndex        =   11
         Top             =   240
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1905
         BTYPE           =   2
         TX              =   "&Consultar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Lancamento_Pesq.frx":2F31
         PICN            =   "frm_Lancamento_Pesq.frx":2F4D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc adoLogo 
         Height          =   330
         Left            =   4200
         Top             =   360
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
      Begin VB.Label lb_Doc 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Doc"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   8625
         TabIndex        =   41
         Top             =   180
         Width           =   840
      End
      Begin VB.Label lb_resumo 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Resumo"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   8625
         TabIndex        =   40
         Top             =   765
         Width           =   1365
      End
      Begin VB.Label lb_Cartao 
         BackStyle       =   0  'Transparent
         Caption         =   "Cartão"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   39
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label lb_Logo 
         BackStyle       =   0  'Transparent
         Caption         =   "Logo"
         Enabled         =   0   'False
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
         Left            =   2280
         TabIndex        =   38
         Top             =   180
         Width           =   615
      End
      Begin VB.Label lb_Dt2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6765
         TabIndex        =   37
         Top             =   1050
         Width           =   480
      End
      Begin VB.Label lb_Dt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data"
         Enabled         =   0   'False
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
         Left            =   5640
         TabIndex        =   36
         Top             =   765
         Width           =   2760
      End
   End
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frm_Lancamento_Pesq.frx":3267
      Height          =   2220
      Left            =   6120
      TabIndex        =   34
      Top             =   4320
      Width           =   5355
      _ExtentX        =   9446
      _ExtentY        =   3916
      _Version        =   393216
      AllowUpdate     =   0   'False
      Enabled         =   -1  'True
      ForeColor       =   8421504
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   5
      BeginProperty Column00 
         DataField       =   "lcp_parc"
         Caption         =   "Parc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0º"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "lcp_dt_vcto"
         Caption         =   "Dt. Vcto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "dd/MM/yy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "lcp_vr_bto"
         Caption         =   "Vr. Bruto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "lcp_vr_liq"
         Caption         =   "Vr. Liq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   2
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "lcp_ndoc"
         Caption         =   "Nº Doc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column02 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column03 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column04 
            ColumnAllowSizing=   0   'False
         EndProperty
      EndProperty
   End
   Begin rdActiveText.ActiveText txt_tx 
      DataField       =   "lnc_tx"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   1200
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   6240
      Width           =   585
      _ExtentX        =   1032
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
      FontSize        =   8.25
   End
   Begin MSDataListLib.DataCombo txt_Logo 
      Bindings        =   "frm_Lancamento_Pesq.frx":327D
      DataField       =   "Logo"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   1200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4800
      Width           =   2025
      _ExtentX        =   3572
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "usl_nome"
      BoundColumn     =   "usl_nome"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSDataListLib.DataCombo txt_Cartao 
      Bindings        =   "frm_Lancamento_Pesq.frx":3293
      DataField       =   "Cartao"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   1200
      TabIndex        =   13
      Top             =   5280
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "tpc_desc"
      BoundColumn     =   "tpc_desc"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin rdActiveText.ActiveText txt_tx_fixo 
      DataField       =   "lnc_tx_fixo"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   2880
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6240
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
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txt_tx_po 
      DataField       =   "lnc_tx_po"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00%"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   5
      EndProperty
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   5040
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6225
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
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txt_NDOC 
      DataField       =   "lnc_ndoc"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   1200
      TabIndex        =   16
      Top             =   4320
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   556
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
      MaxLength       =   20
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin MSDataListLib.DataCombo txt_FormaPg 
      Bindings        =   "frm_Lancamento_Pesq.frx":32AB
      DataField       =   "formapg"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   4440
      TabIndex        =   14
      Top             =   5280
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "fpg_desc"
      BoundColumn     =   "fpg_desc"
      Text            =   ""
      Object.DataMember      =   "tab_forma_pg"
   End
   Begin rdActiveText.ActiveText txt_dt_vnd 
      DataField       =   "lnc_dt_vnd"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   3
      EndProperty
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   4440
      TabIndex        =   15
      Top             =   5760
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
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
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txt_Valor_Vnd 
      DataField       =   "lnc_vr"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   2
      EndProperty
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   1200
      TabIndex        =   17
      Top             =   5760
      Width           =   1305
      _ExtentX        =   2302
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
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txt_NResumo 
      DataField       =   "lnc_nresumo"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   4440
      TabIndex        =   32
      Top             =   4320
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   556
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
      MaxLength       =   20
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.Label Label4 
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
      Left            =   2955
      TabIndex        =   33
      Top             =   4380
      Width           =   1425
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vr Compra"
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
      Left            =   -285
      TabIndex        =   31
      Top             =   5835
      Width           =   1455
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
      Left            =   3075
      TabIndex        =   30
      Top             =   5835
      Width           =   1305
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
      Left            =   3165
      TabIndex        =   29
      Top             =   5205
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
      Left            =   -285
      TabIndex        =   28
      Top             =   4380
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
      Left            =   4260
      TabIndex        =   27
      Top             =   6300
      Width           =   705
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
      Left            =   1635
      TabIndex        =   26
      Top             =   6315
      Width           =   975
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
      Left            =   -60
      TabIndex        =   25
      Top             =   6315
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
      Left            =   -75
      TabIndex        =   24
      Top             =   5340
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
      Left            =   -75
      TabIndex        =   23
      Top             =   4875
      Width           =   1215
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
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":32BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":35D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":37B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":3ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":3DE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":40FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":4418
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":45F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Lancamento_Pesq.frx":490C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   2415
      Left            =   120
      Top             =   4200
      Width           =   11415
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
End
Attribute VB_Name = "frm_Lancamento_Pesq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_str_Det As String

Private Sub adoGrid_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    fCancelDisplay = True
End Sub


Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    adoReg.Caption = "Registro " & adoReg.Recordset.AbsolutePosition & " de " & adoReg.Recordset.RecordCount
    If Not adoReg.Recordset.EOF Then
        Set adoGrid.Recordset = ExecuteSQL("Select * FROM tab_lanc_parc WHERE lcp_num_lanc = '" & adoReg.Recordset.Fields("lnc_num") & "' ORDER BY lcp_parc").Clone
    Else
        Set adoGrid.Recordset = ExecuteSQL("Select * From tab_lanc_parc Where lcp_num_lanc = -999").Clone
    End If
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub


Private Sub bt_Pesq_Click()
Dim w_Str As String
On Error GoTo err1

    If ck(0).Value <> 0 And txt_Logo_P <> "" Then
        w_Str = "lnc_loj = " & txt_Logo_P.BoundText & ""
    End If
    If ck(1).Value <> 0 And txt_Cartao_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "lnc_tipoc = " & txt_Cartao_P.BoundText & ""
    End If
    
    If ck(2).Value <> 0 And Txt_DtI <> "" And Txt_DtF <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
    ElseIf ck(2).Value = 0 And w_Usu_Tipo = "L" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "lnc_dt_vnd >= '" & Format(w_Data_Server - 7, "yyyy-mm-dd") & "' and lnc_dt_vnd <= '" & Format(w_Data_Server, "yyyy-mm-dd") & "'"
    End If
    
    
    If ck(3).Value <> 0 And txt_NDOC_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "lnc_ndoc = '" & txt_NDOC_P & "'"
    End If
    If ck(4).Value <> 0 And txt_NResumo_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "lnc_nresumo = '" & txt_NResumo_P & "'"
    End If
   
    w_Str = "SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, tab_forma_pg.fpg_desc AS FormaPG, tab_lanc.lnc_ndoc, tab_lanc.lnc_dt_vnd, tab_lanc.lnc_tx, tab_lanc.lnc_tx_fixo, tab_lanc.lnc_tx_po, tab_lanc.lnc_vr, tab_lanc.lnc_vr_liq, tab_lanc.lnc_num, tab_lanc.lnc_nresumo, tab_lanc.lnc_tef_pos as TEF_POS FROM tab_lanc, tab_usuario, tab_tipo_cartao, tab_forma_pg WHERE " & w_Str & IIf(Len(w_Str) > 0, " and ", "") & "(tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod) ORDER BY tab_lanc.lnc_dt_vnd, tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc, tab_forma_pg.fpg_desc"
    w_str_Det = "SELECT * FROM tab_lanc_parc"
    
    txt2.text = w_Str

    pic_Pesq.Visible = True
    Set adoReg.Recordset = ExecuteSQL(w_Str).Clone

sair:
    pic_Pesq.Visible = False
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub ck_Click(Index As Integer)
On Error GoTo err1
    Select Case Index
    Case 0:
            lb_Logo.Enabled = ck(Index).Value
            txt_Logo_P.Enabled = ck(Index).Value
            If txt_Logo_P.Enabled = True Then txt_Logo_P.SetFocus
    Case 1:
            lb_Cartao.Enabled = ck(Index).Value
            txt_Cartao_P.Enabled = ck(Index).Value
            If txt_Cartao_P.Enabled = True Then txt_Cartao_P.SetFocus
    Case 2:
            lb_Dt.Enabled = ck(Index).Value
            lb_Dt2.Enabled = ck(Index).Value
            Txt_DtI.Enabled = ck(Index).Value
            Txt_DtF.Enabled = ck(Index).Value
            Txt_DtI = w_Data_Server
            Txt_DtF = w_Data_Server
            If Txt_DtI.Enabled = True Then Txt_DtI.SetFocus
    Case 3:
            lb_Doc.Enabled = ck(Index).Value
            txt_NDOC_P.Enabled = ck(Index).Value
            If txt_NDOC_P.Enabled = True Then txt_NDOC_P.SetFocus
    Case 4:
            lb_resumo.Enabled = ck(Index).Value
            txt_NResumo_P.Enabled = ck(Index).Value
            If txt_NResumo_P.Enabled = True Then txt_NResumo_P.SetFocus
    End Select

sair:
    Exit Sub
err1:
    'MsgBox ERR.Number & " : " & ERR.Description, vbCritical
    Resume sair
End Sub

Private Sub Form_Load()
On Error GoTo err1

    MDI.TBar.Visible = False

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 100
    
    w_Usu = IIf(w_Usu_Tipo = "L", w_Usu_Nome, "%")
    
    Set adoLogo.Recordset = w_ado_Logo.Clone
    Set adoCartao.Recordset = w_ado_CadCartao.Clone

    Txt_DtI = w_Data_Server
    Txt_DtF = w_Data_Server
    ck(2) = 1
   
    If w_Usu_Tipo = "L" Then
        ck(0) = 1
        txt_Logo_P.BoundText = w_Usu_Cod
        ck(0).Enabled = False
        txt_Logo_P.Enabled = False
    End If
    
    bt_Pesq_Click
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub

Private Sub grid_Error(ByVal DataError As Integer, Response As Integer)
    MsgBox DataError & "  : " & Response
End Sub




Private Sub mnuFechar_Click()
    Unload Me
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    Case "excluir": Excluir
    Case "alterar": Alterar
    Case "alterarTodos": alterarTodos
    End Select
End Sub

Private Sub txt_bco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub Timer1_Timer()
    bt_Pesq_Click
    Timer1.Enabled = False
End Sub

Private Sub txt_Cartao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_Cartao_P_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub



Private Sub Txt_DtI_Validate(Cancel As Boolean)
    
    If IsDate(Txt_DtI) Then
        Txt_DtF = Txt_DtI
    Else
        Txt_DtI = w_Data_Server
    End If
    
    If w_Usu_Tipo = "L" Then
        If (CVDate(Txt_DtI) < w_Data_Server - 7) Then
            Txt_DtI = ""
            MsgBox "Você não pode consultar lançamentos de mais de uma semana atrás!", vbCritical
            Txt_DtI.SetFocus
        End If
    End If
End Sub


Private Sub txt_FormaPg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_Logo_P_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub



Sub Excluir()
On Error GoTo err1
    If vbYes = MsgBox("O Documento Nº : " & adoReg.Recordset.Fields("lnc_ndoc") & " será excluído." & Chr(13) & "Tem certeza?", vbQuestion + vbYesNo + vbDefaultButton2) Then
           w_RegAf = 0
           'Excluir as Parcelas
           If adoGrid.Recordset.RecordCount > 0 Then
               Call ExecuteSQL("Delete From tab_lanc_parc Where lcp_num_lanc = " & adoReg.Recordset.Fields("lnc_num") & "", w_RegAf)
           Else
               w_RegAf = 1
           End If
           
           If w_RegAf > 0 Then
                'Excluir o Lançamento
                Call ExecuteSQL("Delete From tab_lanc Where lnc_num = " & adoReg.Recordset.Fields("lnc_num") & "", w_RegAf)
                bt_Pesq_Click
           Else
                MsgBox "Não foi possível excluir!", vbCritical
           End If

    End If
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Sub alterarTodos()
    
    controle = True 'Não exibir "REGISTRO SALVO COM SUCESSO!"
    adoReg.Recordset.MoveFirst
    Do While Not adoReg.Recordset.EOF
        'Verifica se possui alguma parcela baixada
        wResp = vbYes
        If 0 <> ExecuteSQL("Select Count(*) from tab_lanc_parc Where lcp_num_lanc = '" & adoReg.Recordset.Fields("lnc_num") & "' and lcp_baixa <> '0000-00-00'").Fields(0) Then
            wResp = vbNo
        '    wResp = MsgBox("Umas das parcelas já foi baixada, pois após a alteração as parcelas baixadas retornarão como não baixadas! " & Chr(13) & "Deseja realmente alterar?", vbQuestion + vbYesNo)
        End If
        If wResp = vbYes Then
            Alterar
            frm_Lancamento_Alt.Form_Activate
            frm_Lancamento_Alt.Salvar
            frm_Lancamento_Alt.mnuFechar_Click
        End If
        
        adoReg.Recordset.MoveNext
    Loop
    controle = False
End Sub

Sub Alterar()
On Error GoTo err1
    
    If IsEmpty(wResp) Then wResp = vbYes
    'Verifica se possui alguma parcela baixada
    If wResp = vbYes Then
    If 0 <> ExecuteSQL("Select Count(*) from tab_lanc_parc Where lcp_num_lanc = '" & adoReg.Recordset.Fields("lnc_num") & "' and lcp_baixa <> '0000-00-00'").Fields(0) Then
        wResp = MsgBox("Umas das parcelas já foi baixada, pois após a alteração as parcelas baixadas retornarão como não baixadas! " & Chr(13) & "Deseja realmente alterar?", vbQuestion + vbYesNo)
    End If
    End If
    
    If wResp = vbYes Then
        frm_Lancamento_Alt.Show
        frm_Lancamento_Alt.txt_NUM = adoReg.Recordset.Fields("lnc_num")
        frm_Lancamento_Alt.txt_Form = "Pesq"
'    Else
'        MsgBox "Você não pode alterar este lançamento, pois já foi dado baixa em alguma das parcelas!", vbExclamation
    End If

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub
