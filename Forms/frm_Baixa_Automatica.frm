VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Baixa_Automatica 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baixa Automática"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12840
   Icon            =   "frm_Baixa_Automatica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   12840
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   1535
      ButtonWidth     =   1693
      ButtonHeight    =   1429
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   6
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
            Caption         =   "&Alt. Vcto"
            Key             =   "alterar"
            Object.ToolTipText     =   "Alterar a Data de Vencimento da Parcela Selecionada"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "Alt. Lanç."
            Key             =   "altlanc"
            Object.ToolTipText     =   "Altera o Lançamento"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Relatório"
            Key             =   "relatorio"
            Object.ToolTipText     =   "Visualizar Relatório para Impressão"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin Skin_Button.ctr_Button btn_Acompanhamento 
         Height          =   735
         Left            =   8600
         TabIndex        =   42
         Top             =   45
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   "A&BS"
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
         MICON           =   "frm_Baixa_Automatica.frx":27A2
         PICN            =   "frm_Baixa_Automatica.frx":27BE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button1 
         Height          =   735
         Left            =   10365
         TabIndex        =   35
         TabStop         =   0   'False
         ToolTipText     =   "Cartões não Recebidos"
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   "Não Rec."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frm_Baixa_Automatica.frx":2C10
         PICN            =   "frm_Baixa_Automatica.frx":2C2C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_Rel 
         Height          =   735
         Left            =   9285
         TabIndex        =   34
         TabStop         =   0   'False
         ToolTipText     =   "Cartões Recebidos"
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1296
         BTYPE           =   2
         TX              =   "Rec."
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         MICON           =   "frm_Baixa_Automatica.frx":307E
         PICN            =   "frm_Baixa_Automatica.frx":309A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   3
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         Left            =   3720
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "Baixa Automática"
         Top             =   120
         Width           =   4005
      End
   End
   Begin VB.TextBox txtSql 
      Height          =   375
      Left            =   3480
      TabIndex        =   43
      Text            =   "[debug]"
      Top             =   2280
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox pic_Pesq 
      Height          =   555
      Left            =   5640
      Picture         =   "frm_Baixa_Automatica.frx":34EC
      ScaleHeight     =   495
      ScaleWidth      =   465
      TabIndex        =   41
      Top             =   4320
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   10080
      Top             =   6240
   End
   Begin VB.Frame fr 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Alterar Data de Vencimento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1350
      Left            =   840
      TabIndex        =   29
      Top             =   4080
      Visible         =   0   'False
      Width           =   3375
      Begin rdActiveText.ActiveText txt_DtVcto 
         Height          =   315
         Left            =   525
         TabIndex        =   30
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
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
         FontSize        =   8.25
      End
      Begin Skin_Button.ctr_Button bt_Sal_F 
         Height          =   525
         Left            =   1890
         TabIndex        =   31
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         BTYPE           =   2
         TX              =   ""
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Baixa_Automatica.frx":37F6
         PICN            =   "frm_Baixa_Automatica.frx":3812
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_Canc_F 
         Height          =   525
         Left            =   2370
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   480
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   926
         BTYPE           =   2
         TX              =   ""
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frm_Baixa_Automatica.frx":3B2C
         PICN            =   "frm_Baixa_Automatica.frx":3B48
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbResumo 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Vcto:"
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
         Left            =   570
         TabIndex        =   33
         Top             =   405
         Width           =   975
      End
   End
   Begin rdActiveText.ActiveText txt_Bto 
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
      Left            =   8670
      TabIndex        =   26
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
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
      Text            =   "0,00"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txt_Liq 
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
      Left            =   7560
      TabIndex        =   25
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
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
      Text            =   "0,00"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6630
      Width           =   12840
      _ExtentX        =   22648
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
   Begin MSDataGridLib.DataGrid grid 
      Bindings        =   "frm_Baixa_Automatica.frx":3E62
      Height          =   3690
      Left            =   0
      TabIndex        =   17
      Top             =   2520
      Width           =   12825
      _ExtentX        =   22622
      _ExtentY        =   6509
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
      ColumnCount     =   12
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
      BeginProperty Column02 
         DataField       =   "lcp_parc"
         Caption         =   "Parc."
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
      BeginProperty Column03 
         DataField       =   "vcto"
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
      BeginProperty Column04 
         DataField       =   "N"
         Caption         =   "Resumo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   """R$ ""#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "NDOC"
         Caption         =   "N. Doc"
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
      BeginProperty Column07 
         DataField       =   "vr_liq"
         Caption         =   "P. Liq"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column08 
         DataField       =   "vr_bto"
         Caption         =   "P. Bruto"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column09 
         DataField       =   "vr_compra"
         Caption         =   "Vr. Compra"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "#.##0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
      EndProperty
      BeginProperty Column10 
         DataField       =   "lnc_dt_vnd"
         Caption         =   "Dt. Vnd"
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
      BeginProperty Column11 
         DataField       =   "Baixa"
         Caption         =   "Baixa"
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         BeginProperty Column00 
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column04 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column05 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
            Object.Visible         =   -1  'True
         EndProperty
         BeginProperty Column06 
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column07 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column08 
            Alignment       =   1
            ColumnAllowSizing=   -1  'True
         EndProperty
         BeginProperty Column09 
            Alignment       =   1
         EndProperty
         BeginProperty Column10 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column11 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
         EndProperty
      EndProperty
   End
   Begin rdActiveText.ActiveText txt_Qtde 
      BeginProperty DataFormat 
         Type            =   0
         Format          =   """R$ ""#.##0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   0
      EndProperty
      Height          =   315
      Left            =   12030
      TabIndex        =   27
      Top             =   6240
      Width           =   600
      _ExtentX        =   1058
      _ExtentY        =   556
      Alignment       =   2
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
      Text            =   "0"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin rdActiveText.ActiveText txtDtAntecipado 
      Height          =   315
      Left            =   4920
      TabIndex        =   44
      Top             =   1560
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      BackColor       =   12632319
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
      Height          =   1635
      Left            =   0
      TabIndex        =   18
      Top             =   840
      Width           =   12765
      Begin VB.CheckBox ckAntecipado 
         Caption         =   "Antecipar?"
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
         Left            =   4680
         TabIndex        =   47
         Top             =   420
         Width           =   1335
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   1200
         Top             =   960
      End
      Begin MSAdodcLib.Adodc adoLogo 
         Height          =   330
         Left            =   2505
         Top             =   525
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
         Caption         =   "Adodc1"
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
         Left            =   3720
         Top             =   1110
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
         Height          =   315
         Index           =   4
         Left            =   225
         TabIndex        =   5
         Top             =   1290
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
         Left            =   225
         TabIndex        =   4
         Top             =   1080
         Width           =   1575
      End
      Begin VB.CheckBox ck 
         Caption         =   "Data Vcto"
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
         Left            =   225
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox ck 
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
         Height          =   300
         Index           =   6
         Left            =   225
         TabIndex        =   2
         Top             =   600
         Width           =   1800
      End
      Begin VB.CheckBox ck 
         Caption         =   "Mostrar Baixados?"
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
         Index           =   5
         Left            =   4680
         TabIndex        =   38
         Top             =   150
         Width           =   1935
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
         Left            =   225
         TabIndex        =   1
         Top             =   390
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
         Left            =   225
         TabIndex        =   0
         Top             =   170
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
         Left            =   7020
         TabIndex        =   10
         Top             =   1095
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
         Left            =   8385
         TabIndex        =   11
         Top             =   1080
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
         Bindings        =   "frm_Baixa_Automatica.frx":3E77
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   510
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
         Bindings        =   "frm_Baixa_Automatica.frx":3E8D
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   1110
         Width           =   4170
         _ExtentX        =   7355
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
         Left            =   9825
         TabIndex        =   12
         Top             =   510
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
         Left            =   9825
         TabIndex        =   13
         Top             =   1095
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
         Left            =   11520
         TabIndex        =   14
         Top             =   360
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
         MICON           =   "frm_Baixa_Automatica.frx":3EA5
         PICN            =   "frm_Baixa_Automatica.frx":3EC1
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin rdActiveText.ActiveText Txt_DtIP 
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
         Left            =   7020
         TabIndex        =   8
         Top             =   510
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
      Begin rdActiveText.ActiveText Txt_DtFP 
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
         Left            =   8385
         TabIndex        =   9
         Top             =   510
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
      Begin VB.Label lb_DtP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data Protocolo"
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
         Left            =   7080
         TabIndex        =   40
         Top             =   270
         Width           =   2310
      End
      Begin VB.Label lb_Dt2P 
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
         Left            =   7965
         TabIndex        =   39
         Top             =   555
         Width           =   480
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
         Left            =   9825
         TabIndex        =   24
         Top             =   270
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
         Left            =   9825
         TabIndex        =   23
         Top             =   855
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
         Left            =   2040
         TabIndex        =   22
         Top             =   870
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
         Left            =   2040
         TabIndex        =   21
         Top             =   270
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
         Left            =   7965
         TabIndex        =   20
         Top             =   1155
         Width           =   480
      End
      Begin VB.Label lb_Dt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data Vcto"
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
         Left            =   6840
         TabIndex        =   19
         Top             =   855
         Width           =   2385
      End
   End
   Begin VB.Label lbBaixarTodos 
      BackStyle       =   0  'Transparent
      Caption         =   "P/ Baixar TODOS pressione F8"
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   2520
      TabIndex        =   46
      Top             =   6240
      Width           =   2805
   End
   Begin VB.Label lbBaixarRemTodos 
      BackStyle       =   0  'Transparent
      Caption         =   "P/ Remover TODOS  CTRL + T"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   2520
      TabIndex        =   45
      Top             =   6420
      Width           =   2925
   End
   Begin VB.Label lbBaixarRem 
      BackStyle       =   0  'Transparent
      Caption         =   "P/ Remover   CTRL + R"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   180
      TabIndex        =   37
      Top             =   6420
      Width           =   2085
   End
   Begin VB.Label lbBaixar 
      BackStyle       =   0  'Transparent
      Caption         =   "P/ Baixar pressione  F5"
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
      ForeColor       =   &H00C00000&
      Height          =   240
      Left            =   180
      TabIndex        =   36
      Top             =   6240
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qtde Parcelas :"
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
      Left            =   10560
      TabIndex        =   28
      Top             =   6285
      Width           =   1335
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   8760
      Top             =   4200
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
            Picture         =   "frm_Baixa_Automatica.frx":41DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":44F5
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":46CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":49E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":4D03
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":501D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":5337
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":5511
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Baixa_Automatica.frx":582B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
   Begin VB.Menu mnuAlt 
      Caption         =   "|  &Alterar Vcto  |"
   End
   Begin VB.Menu mnuAltlanc 
      Caption         =   "Altera Lanç.  |"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuAc 
      Caption         =   "Ações"
      Begin VB.Menu mnuAcBaixar 
         Caption         =   "Baixar"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuAcRem 
         Caption         =   "Remover Baixa"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuAcBaixarTodos 
         Caption         =   "Baixar TODOS"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuAcRemTodos 
         Caption         =   "Remover TODOS"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frm_Baixa_Automatica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wRPT As Boolean
Dim wWherePG As String
Dim wWhereN As String
Dim wHabStatus As Boolean
Dim wAntecipado As Boolean

Private Sub adoGrid_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
    fCancelDisplay = True
End Sub


Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    adoReg.Caption = "Registro " & adoReg.Recordset.AbsolutePosition & " de " & adoReg.Recordset.RecordCount
    
    If wHabStatus = True And (txt_Cartao_P.BoundText = "17" Or txt_Cartao_P.BoundText = "18" Or txt_Cartao_P.BoundText = "19") Then
        MDI.StatusBar1.Panels(2).text = adoReg.Recordset.Fields("N")
    ElseIf wHabStatus = True Then
        MDI.StatusBar1.Panels(2).text = ""
    End If

sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub


Private Sub bt_Canc_F_Click()
    txt_DtVcto = ""
    fr.Visible = False
    frFiltro.Visible = True
    Grid.Enabled = True
    adoReg.Enabled = True
End Sub

Private Sub bt_Pesq_Click()
Dim w_Str As String
On Error GoTo err1

wHabStatus = False

    wWherePG = ""
    If ck(0).Value <> 0 And txt_Logo_P <> "" Then
        w_Str = "tab_lanc.lnc_loj = " & txt_Logo_P.BoundText & ""
    End If
    If ck(1).Value <> 0 And txt_Cartao_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "tab_lanc.lnc_tipoc = " & txt_Cartao_P.BoundText & ""
    End If
    wWherePG = w_Str
    
    If ck(2).Value <> 0 And Txt_DtI <> "" And Txt_DtF <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        
        'Mostrar Baixados não selecionado
        If ck(5).Value = 0 Then
            w_Str = w_Str & "tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
        
            wWherePG = IIf(Len(wWherePG) > 0, wWherePG & " and ", "")
            wWhereN = wWherePG
            wWherePG = wWherePG & "tab_lanc_parc.lcp_baixa = '0000-00-00' and tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
            
            wWhereN = wWhereN & "tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
            wWhereN = wWhereN & "and tab_lanc_parc.lcp_baixa = '0000-00-00'"
        Else
            w_Str = w_Str & "tab_lanc_parc.lcp_baixa >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_baixa <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
        
            wWherePG = IIf(Len(wWherePG) > 0, wWherePG & " and ", "")
            wWhereN = wWherePG
            wWherePG = wWherePG & "tab_lanc_parc.lcp_baixa >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_baixa <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
            
            wWhereN = wWhereN & "tab_lanc_parc.lcp_baixa >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' and tab_lanc_parc.lcp_baixa <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
            'wWhereN = wWhereN & "and tab_lanc_parc.lcp_baixa = '0000-00-00'"
        End If
    End If
    
    If ck(6).Value <> 0 And Txt_DtIP <> "" And Txt_DtFP <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtIP, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtFP, "yyyy-mm-dd") & "'"
        
        wWherePG = IIf(Len(wWherePG) > 0, wWherePG & " and ", "")
        wWhereN = wWherePG
        wWherePG = wWherePG & "tab_lanc_parc.lcp_baixa = '0000-00-00' and tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtIP, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtFP, "yyyy-mm-dd") & "'"
        
        wWhereN = wWhereN & "tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtIP, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtFP, "yyyy-mm-dd") & "'"
        If Not InStr(wWhereN, "and tab_lanc_parc.lcp_baixa = '0000-00-00'") Then
            wWhereN = wWhereN & "and tab_lanc_parc.lcp_baixa = '0000-00-00'"
        End If
    End If
    
    If ck(3).Value <> 0 And txt_NDOC_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "tab_lanc_parc.lcp_ndoc = '" & txt_NDOC_P & "'"
    End If
    If ck(4).Value <> 0 And txt_NResumo_P <> "" Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "tab_lanc_parc.lcp_nresumo = '" & txt_NResumo_P & "'"
    End If
    
    If ck(5).Value = 1 Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "not tab_lanc_parc.lcp_baixa = '0000-00-00'"
    ElseIf ck(5).Value = 0 Then
        w_Str = IIf(Len(w_Str) > 0, w_Str & " and ", "")
        w_Str = w_Str & "tab_lanc_parc.lcp_baixa = '0000-00-00'"
    End If
    
    wWhereN = wWherePG
    
    If txt_Cartao_P.BoundText = "6" Or txt_Cartao_P.BoundText = "17" Or txt_Cartao_P.BoundText = "18" Or txt_Cartao_P.BoundText = "19" Then
        w_Str = "SELECT tab_lanc.lnc_tipoc, tab_lanc.lnc_num, tab_lanc_parc.lcp_nresumo as N, tab_lanc_parc.lcp_ndoc as NDOC, tab_lanc_parc.lcp_dt_vcto AS vcto, tab_tipo_cartao.tpc_desc AS cartao,  tab_lanc_parc.lcp_parc,  tab_lanc.lnc_loj, tab_lanc.lnc_dt_vnd, tab_usuario.usl_nome AS Logo, tab_lanc_parc.lcp_baixa AS Baixa, tab_lanc.lnc_tef_pos AS TEF_POS, Sum(tab_lanc.lnc_vr) AS vr_compra, Sum(tab_lanc_parc.lcp_vr_bto) AS vr_bto, Sum(tab_lanc_parc.lcp_vr_liq) AS vr_liq " & _
                "From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc " & _
                "Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num  And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod  And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) " & _
                IIf(Len(w_Str) > 0, " and  (" & w_Str & ") ", "") & _
                "GROUP BY N, NDOC, tab_lanc_parc.lcp_dt_vcto, cartao, tab_lanc_parc.lcp_parc, tab_lanc.lnc_loj, tab_lanc.lnc_dt_vnd, Logo, Baixa, TEF_POS " & _
                "ORDER BY tab_lanc_parc.lcp_dt_vcto, tab_tipo_cartao.tpc_desc;"
    Else
        w_Str = "SELECT tab_lanc.lnc_tipoc, tab_lanc.lnc_num, tab_lanc_parc.lcp_nresumo as N, tab_lanc_parc.lcp_ndoc as NDOC, tab_lanc_parc.lcp_dt_vcto AS vcto, tab_tipo_cartao.tpc_desc AS cartao,  tab_lanc_parc.lcp_parc,  tab_lanc.lnc_loj, tab_lanc.lnc_dt_vnd, tab_usuario.usl_nome AS Logo, tab_lanc_parc.lcp_baixa AS Baixa, tab_lanc.lnc_tef_pos AS TEF_POS, Sum(tab_lanc.lnc_vr) AS vr_compra, Sum(tab_lanc_parc.lcp_vr_bto) AS vr_bto, Sum(tab_lanc_parc.lcp_vr_liq) AS vr_liq " & _
                "From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc " & _
                "Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num  And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod  And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) " & _
                IIf(Len(w_Str) > 0, " and  (" & w_Str & ") ", "") & _
                "GROUP BY N, NDOC, tab_lanc_parc.lcp_dt_vcto, cartao, tab_lanc_parc.lcp_parc, tab_lanc.lnc_loj, tab_lanc.lnc_dt_vnd, Logo, Baixa, TEF_POS " & _
                "ORDER BY tab_lanc_parc.lcp_dt_vcto, tab_tipo_cartao.tpc_desc;"
    End If

txtSql.text = w_Str
    pic_Pesq.Visible = True
    DoEvents
    Set adoReg.Recordset = ExecuteSQL(w_Str)

    Grid.Visible = False
    txt_Liq = 0
    txt_Bto = 0
    txt_Qtde = 0
    If Not adoReg.Recordset.EOF Then
        Do While Not adoReg.Recordset.EOF
            txt_Liq = CDbl(txt_Liq) + CDbl(adoReg.Recordset.Fields("vr_liq"))
            txt_Bto = CDbl(txt_Bto) + CDbl(adoReg.Recordset.Fields("vr_bto"))
            
            adoReg.Recordset.MoveNext
        Loop
        adoReg.Recordset.MoveFirst
        txt_Bto = Format(txt_Bto, "#,##0.00")
        txt_Liq = Format(txt_Liq, "#,##0.00")
        txt_Qtde = adoReg.Recordset.RecordCount
    End If
    
    Grid.Visible = False
    'If txt_Cartao_P.BoundText = 6 Or txt_Cartao_P.BoundText = 17 Or txt_Cartao_P.BoundText = 18 Or txt_Cartao_P.BoundText = 19 Then
    '    grid.Columns(4).Visible = False
    '    grid.Columns(5).Visible = True
    'Else
    '    grid.Columns(4).Visible = True
    '    grid.Columns(5).Visible = False
    'End If
    Grid.Visible = True
    
wHabStatus = True
    
sair:
    Grid.Visible = True
    pic_Pesq.Visible = False
    Exit Sub
err1:
   ' MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub bt_Rel_Click()
On Error GoTo err1

    
    w_SQL = "  SHAPE {SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, SUM(tab_lanc_parc.lcp_vr_liq) AS TotalRec, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc, tab_lanc_parc.lcp_tipo, tab_lanc.lnc_tef_pos as TEF_POS FROM tab_lanc_parc, tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num) AND (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) " & IIf(wWherePG <> "", " and " & wWherePG, "") & " and tab_lanc_parc.lcp_baixa <> '0000-00-00' GROUP BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc}  AS Rpt_Resumo_Recebimentos COMPUTE Rpt_Resumo_Recebimentos BY 'Logo'"
    'w_SQL = "SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, SUM(tab_lanc_parc.lcp_vr_liq) AS TotalRec, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc, tab_lanc_parc.lcp_tipo FROM tab_lanc_parc, tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num) AND (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) " & IIf(wWherePG <> "", " and " & wWherePG, "") & " and tab_lanc_parc.lcp_baixa <> '0000-00-00' GROUP BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc"
    
    
    Dim rs As New ADODB.Recordset
    Set rs = ExecuteSQL(w_SQL, , "MSDataShape").Clone
    
    If rs.RecordCount > 0 Then
        Set Rel_ResumoRec.DataSource = rs.Clone
        wRPT = True
        mnuAc.Visible = False
        mnuAlt.Visible = False
        If ck(0).Value = 0 Then
            Rel_ResumoRec.Sections("SecCab").Controls("LBTitulo").Caption = "Recebidos"
        Else
            Rel_ResumoRec.Sections("SecCab").Controls("LBTitulo").Caption = txt_Logo_P & " - Recebidos"
            'Rel_ResumoRec.Sections("SecDet").Controls("txtlogo").Visible = False
        End If
        Rel_ResumoRec.Show
        Rel_ResumoRec.WindowState = vbMaximized
    Else
        MsgBox "Nenhum registro encontrado para gerar o relatório!", vbExclamation
    End If

sair:
   
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub bt_Sal_F_Click()
Dim wAdo As ADODB.Recordset
   
   'If txt_Cartao_P = "" Then
    '    wCampo = "lcp_nresumo"
    'Else
    '    If txt_Cartao_P.BoundText = 6 Or txt_Cartao_P.BoundText = 17 Or txt_Cartao_P.BoundText = 18 Or txt_Cartao_P.BoundText = 19 Then
    '        wCampo = "lcp_ndoc"
    '    Else
    '        wCampo = "lcp_nresumo"
    '    End If
    'End If
    w_Str = "SELECT tab_lanc_parc.lcp_num " & _
            "From tab_lanc, tab_lanc_parc " & _
            "Where (tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num) " & _
            " and tab_lanc_parc.lcp_parc = " & adoReg.Recordset.Fields("lcp_parc") & _
            " and tab_lanc_parc.lcp_nresumo = '" & adoReg.Recordset.Fields("N") & _
            "' and tab_lanc_parc.lcp_ndoc = '" & adoReg.Recordset.Fields("NDOC") & _
            "' and tab_lanc.lnc_loj = " & adoReg.Recordset.Fields("lnc_loj") & _
            " and tab_lanc.lnc_tipoc = " & adoReg.Recordset.Fields("lnc_tipoc") & _
            " and tab_lanc.lnc_dt_vnd = '" & adoReg.Recordset.Fields("lnc_dt_vnd") & "'"
        
    Set wAdo = ExecuteSQL(w_Str).Clone
    
    Do While Not wAdo.EOF
        Call ExecuteSQL("UPDATE tab_lanc_parc SET lcp_dt_vcto = '" & Format(txt_DtVcto, "yyyy-mm-dd") & "'  WHERE (lcp_num = " & wAdo.Fields(0) & ")")
        wAdo.MoveNext
    Loop
    
    Grid.Enabled = True
    adoReg.Enabled = True
    fr.Visible = False
    frFiltro.Visible = True
    
    bt_Pesq_Click
End Sub

Private Sub btn_Acompanhamento_Click()
    frm_Acompanhamento.Show
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
    Case 5:
            If ck(5).Value = 0 Then
                ctr_Button1.Enabled = True
                bt_Rel.Enabled = False
                ck(2).Caption = "Data Vcto"
                lb_Dt.Caption = "Escopo de Data Vcto"
            Else
                ctr_Button1.Enabled = False
                bt_Rel.Enabled = True
                ck(2).Caption = "Data Baixa"
                lb_Dt.Caption = "Escopo de Data Baixa"
            End If
    Case 6:
            lb_DtP.Enabled = ck(Index).Value
            lb_Dt2P.Enabled = ck(Index).Value
            Txt_DtIP.Enabled = ck(Index).Value
            Txt_DtFP.Enabled = ck(Index).Value
            Txt_DtIP = w_Data_Server
            Txt_DtFP = w_Data_Server
            If Txt_DtIP.Enabled = True Then Txt_DtIP.SetFocus
            
    End Select

sair:
    Exit Sub
err1:
    'MsgBox ERR.Number & " : " & ERR.Description, vbCritical
    Resume sair
End Sub



Private Sub ckAntecipado_Click()
    If ckAntecipado.Value Then
        txtDtAntecipado.Visible = True
        txtDtAntecipado.SetFocus
    Else
        txtDtAntecipado.Visible = False
    End If
End Sub

Private Sub ctr_Button1_Click()
Dim rs As Object
On Error GoTo err1
    

    w_SQL = "SHAPE {SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, SUM(tab_lanc_parc.lcp_vr_liq) AS TotalRec, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc, tab_lanc_parc.lcp_tipo, tab_lanc.lnc_tef_pos as TEF_POS FROM tab_lanc_parc, tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num) AND (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) " & IIf(wWhereN <> "", " and " & wWhereN, "") & " GROUP BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc, tab_lanc.lnc_loj, tab_lanc.lnc_tipoc, TEF_POS ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc}  AS Rpt_Resumo_Recebimentos COMPUTE Rpt_Resumo_Recebimentos BY 'Logo'"
    
    Set rs = New ADODB.Recordset
    Set rs = ExecuteSQL(w_SQL, , "MSDataShape").Clone
    
    If rs.RecordCount > 0 Then
        wRPT = True
        mnuAc.Visible = False
        mnuAlt.Visible = False
        Set Rel_ResumoRec.DataSource = rs.Clone
        If ck(0).Value = 0 Then
            Rel_ResumoRec.Sections("SecCab").Controls("LBTitulo").Caption = "Não Recebidos"
        Else
            Rel_ResumoRec.Sections("SecCab").Controls("LBTitulo").Caption = txt_Logo_P & " - Não Recebidos"
            'Rel_ResumoRec.Sections("SecDet").Controls("txtlogo").Visible = False
        End If
        Rel_ResumoRec.Show
        Rel_ResumoRec.WindowState = vbMaximized
    Else
        MsgBox "Nenhum registro encontrado para gerar o relatório!", vbExclamation
    End If

sair:
    
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
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


Private Sub grid_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuAc
End Sub



Sub mnuAcBaixar_Click()
On Error Resume Next
    
    wAntecipado = False
    If ckAntecipado.Value Then
        If Not IsDate(txtDtAntecipado) Then
            MsgBox "Preencha o campo da Data de Antecipação ou retire a seleção do item 'Antecipado?'", vbCritical, "Data de antecipação não preenchida"
            Exit Sub
        Else
            wAntecipado = True
        End If
    End If
  
    w_Pos = adoReg.Recordset.AbsolutePosition - 1
    Dim wAdo As ADODB.Recordset
    
    'If txt_Cartao_P.BoundText = "6" Or txt_Cartao_P.BoundText = "17" Or txt_Cartao_P.BoundText = "18" Or txt_Cartao_P.BoundText = "19" Then
    '    wCampo = "tab_lanc_parc.lcp_ndoc"
    'Else
    '    wCampo = "tab_lanc_parc.lcp_nresumo"
    'End If
    
    'Set wAdo = ExecuteSQL("SELECT lcp_num FROM tab_lanc , tab_lanc_parc WHERE tab_lanc.lnc_num = tab_lanc_parc.lcp_num_lanc AND " & wCampo & "  = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc.lnc_tipoc = " & adoReg.Recordset.Fields("lnc_tipoc") & " AND tab_lanc.lnc_loj = " & adoReg.Recordset.Fields("lnc_loj") & " AND lcp_Baixa = '0000-00-00' AND tab_lanc_parc.lcp_parc = " & adoReg.Recordset.Fields("lcp_parc") & " AND tab_lanc_parc.lcp_dt_vcto = '" & adoReg.Recordset.Fields("vcto") & "'").Clone
    Set wAdo = ExecuteSQL("SELECT lcp_num FROM tab_lanc , tab_lanc_parc WHERE tab_lanc.lnc_num = tab_lanc_parc.lcp_num_lanc AND tab_lanc_parc.lcp_nresumo = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc_parc.lcp_ndoc = '" & adoReg.Recordset.Fields("ndoc") & "'AND tab_lanc.lnc_tipoc = " & adoReg.Recordset.Fields("lnc_tipoc") & " AND tab_lanc.lnc_loj = " & adoReg.Recordset.Fields("lnc_loj") & " AND lcp_Baixa = '0000-00-00' AND tab_lanc_parc.lcp_parc = " & adoReg.Recordset.Fields("lcp_parc") & " AND tab_lanc_parc.lcp_dt_vcto = '" & adoReg.Recordset.Fields("vcto") & "'").Clone
    'txtSql.Text = "SELECT lcp_num FROM tab_lanc , tab_lanc_parc WHERE tab_lanc.lnc_num = tab_lanc_parc.lcp_num_lanc AND " & wCampo & "  = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc.lnc_tipoc = " & adoReg.Recordset.Fields("lnc_tipoc") & " AND tab_lanc.lnc_loj = " & adoReg.Recordset.Fields("lnc_loj") & " AND lcp_Baixa = '0000-00-00' AND tab_lanc_parc.lcp_parc = " & adoReg.Recordset.Fields("lcp_parc")
    'txtSql.Visible = True
    Do While Not wAdo.EOF
        wNum = wNum & IIf(Len(wNum) > 0, ",", "") & wAdo.Fields(0)
        wAdo.MoveNext
    Loop
    
    If Not IsEmpty(wNum) Then
        Call ExecuteSQL("UPDATE tab_lanc_parc SET lcp_Baixa = '" & IIf(wAntecipado, Format(txtDtAntecipado, "yyyy-mm-dd"), Format(adoReg.Recordset.Fields("VCTO"), "yyyy-mm-dd")) & "', lcp_usu_baixa = " & w_Usu_Cod & " WHERE (tab_lanc_parc.lcp_nresumo = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc_parc.lcp_ndoc = '" & adoReg.Recordset.Fields("ndoc") & "' AND lcp_num IN(" & wNum & "))")
        bt_Pesq_Click
        If Not adoReg.Recordset.EOF Then adoReg.Recordset.Move w_Pos
        Grid.SetFocus
    Else
        MsgBox "Não foi possível baixar, especifique na consulta qual cartão deseja!", vbCritical
    End If
   
End Sub

Private Sub mnuAcBaixarTodos_Click()
Dim wQtRegistros As Integer

On Error Resume Next

    If Not adoReg.Recordset.EOF Then adoReg.Recordset.MoveFirst
    
    Do While Not adoReg.Recordset.EOF
        wQtRegistros = adoReg.Recordset.RecordCount
        mnuAcBaixar_Click
        If wQtRegistros = adoReg.Recordset.RecordCount Then
            MsgBox "Não foi possível baixar algum cartão. Cancelando!", vbCritical, "Erro de baixa"
            Exit Sub
        End If
    Loop
End Sub

Sub mnuAcRem_Click()
On Error Resume Next

    w_Pos = adoReg.Recordset.AbsolutePosition - 1
    
    
    Dim wAdo As ADODB.Recordset
    
    'If txt_Cartao_P.BoundText = "6" Or txt_Cartao_P.BoundText = "17" Or txt_Cartao_P.BoundText = "18" Or txt_Cartao_P.BoundText = "19" Then
    '    wCampo = "tab_lanc_parc.lcp_ndoc"
    'Else
    '    wCampo = "tab_lanc_parc.lcp_nresumo"
    'End If
    
    Set wAdo = ExecuteSQL("SELECT lcp_num FROM tab_lanc , tab_lanc_parc WHERE tab_lanc.lnc_num = tab_lanc_parc.lcp_num_lanc AND tab_lanc_parc.lcp_nresumo = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc_parc.lcp_ndoc = '" & adoReg.Recordset.Fields("ndoc") & "' AND tab_lanc.lnc_tipoc = " & adoReg.Recordset.Fields("lnc_tipoc") & " AND tab_lanc.lnc_loj = " & adoReg.Recordset.Fields("lnc_loj") & " AND lcp_Baixa <> '0000-00-00'  AND tab_lanc_parc.lcp_parc = " & adoReg.Recordset.Fields("lcp_parc")).Clone
   
    Do While Not wAdo.EOF
        wNum = wNum & IIf(Len(wNum) > 0, ",", "") & wAdo.Fields(0)
        wAdo.MoveNext
    Loop
    
    
    If Not IsEmpty(wNum) Then
        Call ExecuteSQL("UPDATE tab_lanc_parc SET lcp_Baixa = '0000-00-00' WHERE (tab_lanc_parc.lcp_nresumo = '" & adoReg.Recordset.Fields("n") & "' AND tab_lanc_parc.lcp_ndoc = '" & adoReg.Recordset.Fields("ndoc") & "' AND lcp_num IN(" & wNum & "))")
        bt_Pesq_Click
        If Not adoReg.Recordset.EOF Then adoReg.Recordset.Move w_Pos
        Grid.SetFocus
    Else
        MsgBox "Não foi possível remover a baixa, especifique na consulta qual cartão deseja!", vbCritical
    End If
    
End Sub

Private Sub mnuAcRemTodos_Click()
Dim wQtRegistros As Integer

On Error Resume Next

    If Not adoReg.Recordset.EOF Then adoReg.Recordset.MoveFirst
    
    Do While Not adoReg.Recordset.EOF
        wQtRegistros = adoReg.Recordset.RecordCount
        mnuAcRem_Click
        If wQtRegistros = adoReg.Recordset.RecordCount Then
            MsgBox "Não foi possível remover a baixa algum cartão. Cancelando!", vbCritical, "Erro de remoção de baixa"
            Exit Sub
        End If
    Loop
End Sub

Private Sub mnuAlt_Click()
On Error GoTo err1

    If adoReg.Recordset.Fields("Baixa") = "0000-00-00" Or IsNull(adoReg.Recordset.Fields("Baixa")) Then
        frFiltro.Visible = False
        fr.Visible = True
        txt_DtVcto = Format(adoReg.Recordset.Fields("vcto"), "dd/mm/yyyy")
        Grid.Enabled = False
        adoReg.Enabled = False
    Else
        MsgBox "Não é possível alterar o vcto porque esta parcela já foi baixada!", vbCritical
    End If

sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuRelatorio_Click()
    
Dim w_Rec As New Recordset

On Error Resume Next
 
    Set w_Rec = adoReg.Recordset.Clone

    Set Rel_Baixa_Automatica.DataSource = w_Rec.Clone
    Rel_Baixa_Automatica.WindowState = vbMaximized
    'Se "mostrar baixados", muda título para "Pendência do Relatório das Baixas"
    If ck(5).Value = 1 Then Rel_Baixa_Automatica.Sections("Section4").Controls("lbTitulo").Caption = "Pendência do Relatório das Baixas"
    Rel_Baixa_Automatica.Show

End Sub


Private Sub mnuAltlanc_Click()
On Error GoTo err1
    
    wResp = vbYes
    'Verifica se possui alguma parcela baixada
    If 0 <> ExecuteSQL("Select Count(*) from tab_lanc_parc Where lcp_num_lanc = '" & adoReg.Recordset.Fields("lnc_num") & "' and lcp_baixa <> '0000-00-00'").Fields(0) Then
        wResp = MsgBox("Umas das parcelas já foi baixada, pois após a alteração as parcelas baixadas retonarão como não baixadas! " & Chr(13) & "Deseja realmente alterar?", vbQuestion + vbYesNo)
    End If
    
    If wResp = vbYes Then
        frm_Lancamento_Alt.Show
        frm_Lancamento_Alt.txt_NUM = adoReg.Recordset.Fields("lnc_num")
        frm_Lancamento_Alt.txt_Form = "Baixa"
    End If

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuFechar_Click()
    MDI.StatusBar1.Panels(2).text = ""
    If wRPT = True Then
        Unload Rel_ResumoRec
        wRPT = False
        mnuAc.Visible = True
        mnuAlt.Visible = True

    Else
        Unload Me
    End If
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "fechar": mnuFechar_Click
        Case "alterar": mnuAlt_Click
        Case "altlanc": mnuAltlanc_Click
        Case "relatorio": mnuRelatorio_Click
    End Select
End Sub

Private Sub txt_bco_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_Cartao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub Text2_GotFocus()
    Grid.SetFocus
End Sub

Private Sub Timer1_Timer()
    lbBaixar.Enabled = Not lbBaixar.Enabled
    lbBaixarRem.Enabled = Not lbBaixarRem.Enabled
    
    lbBaixarTodos.Enabled = Not lbBaixarTodos.Enabled
    lbBaixarRemTodos.Enabled = Not lbBaixarRemTodos.Enabled
End Sub

Private Sub Timer2_Timer()
    bt_Pesq_Click
    Timer2.Enabled = False
End Sub

Private Sub txt_Cartao_P_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_FormaPg_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub Txt_DtI_Validate(Cancel As Boolean)
    If IsDate(Txt_DtI) Then
        Txt_DtF = Txt_DtI
    Else
        Txt_DtF = w_Data_Server
    End If
End Sub

Private Sub Txt_DtIP_Validate(Cancel As Boolean)
    If IsDate(Txt_DtIP) Then
        Txt_DtFP = Txt_DtIP
    Else
        Txt_DtFP = w_Data_Server
    End If
End Sub


Private Sub txt_Logo_P_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub
