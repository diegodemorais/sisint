VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Total_Lanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Total Lançamentos - Diários"
   ClientHeight    =   5434
   ClientLeft      =   39
   ClientTop       =   611
   ClientWidth     =   9737
   Icon            =   "frm_Total_Lanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5434
   ScaleWidth      =   9737
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   819
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9737
      _ExtentX        =   17181
      _ExtentY        =   1438
      ButtonWidth     =   1244
      ButtonHeight    =   1376
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   1
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fecha&r"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar Tela"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
      EndProperty
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   23.77
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   1470
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Total Lançamentos - Diários"
         Top             =   120
         Width           =   7455
      End
   End
   Begin Skin_Button.ctr_Button bt_RPT 
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   4785
      Width           =   1935
      _ExtentX        =   3403
      _ExtentY        =   1078
      BTYPE           =   2
      TX              =   "Relatório"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
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
      MICON           =   "frm_Total_Lanc.frx":27A2
      PICN            =   "frm_Total_Lanc.frx":27BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame fr_Resumo 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Inserir Resumo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12.23
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   2160
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   5055
      Begin MSAdodcLib.Adodc adoCartao 
         Height          =   330
         Left            =   2640
         Top             =   480
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2109
         _ExtentY        =   575
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
            Size            =   7.47
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataListLib.DataCombo txt_cartao 
         Bindings        =   "frm_Total_Lanc.frx":3AA0
         Height          =   286
         Left            =   1365
         TabIndex        =   9
         Top             =   481
         Width           =   3328
         _ExtentX        =   6134
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tpc_desc"
         BoundColumn     =   "tpc_cod"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin rdActiveText.ActiveText txt_NResumo 
         Height          =   315
         Left            =   1365
         TabIndex        =   10
         Top             =   1080
         Width           =   1995
         _ExtentX        =   3522
         _ExtentY        =   551
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxLength       =   20
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   7.472
      End
      Begin Skin_Button.ctr_Button bt_Sal_F 
         Height          =   525
         Left            =   3690
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   885
         Width           =   495
         _ExtentX        =   863
         _ExtentY        =   935
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.47
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
         MICON           =   "frm_Total_Lanc.frx":3AB8
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
         Left            =   4200
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   885
         Width           =   495
         _ExtentX        =   863
         _ExtentY        =   935
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.47
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
         MICON           =   "frm_Total_Lanc.frx":3AD4
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
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Resumo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   1065
      End
      Begin VB.Label lbCartao_pesq 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartão"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   630
         TabIndex        =   14
         Top             =   555
         Width           =   675
      End
   End
   Begin VB.Frame fr_Principal 
      Height          =   3870
      Left            =   120
      TabIndex        =   16
      Top             =   870
      Width           =   9495
      Begin Skin_Button.ctr_Button bt_Pesq 
         Height          =   585
         Left            =   7725
         TabIndex        =   4
         Top             =   285
         Width           =   1635
         _ExtentX        =   2875
         _ExtentY        =   1030
         BTYPE           =   2
         TX              =   "Pesquisar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.47
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
         MICON           =   "frm_Total_Lanc.frx":3AF0
         PICN            =   "frm_Total_Lanc.frx":3B0C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSAdodcLib.Adodc adoReg 
         Height          =   375
         Left            =   3480
         Top             =   3240
         Visible         =   0   'False
         Width           =   1935
         _ExtentX        =   3403
         _ExtentY        =   671
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
            Size            =   7.47
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid grid 
         Bindings        =   "frm_Total_Lanc.frx":404F
         Height          =   2775
         Left            =   120
         TabIndex        =   7
         Top             =   1005
         Width           =   9255
         _ExtentX        =   16318
         _ExtentY        =   4888
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "logo"
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
         BeginProperty Column01 
            DataField       =   "cartao"
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
         BeginProperty Column02 
            DataField       =   "tipo"
            Caption         =   "Tipo"
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
            DataField       =   "qtde_lanc"
            Caption         =   "Qtde Lanç."
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
            DataField       =   "total"
            Caption         =   "Vr. Lanç"
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
         BeginProperty Column05 
            DataField       =   "Resumo"
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
         BeginProperty Column06 
            DataField       =   "Dt"
            Caption         =   "Data Vnd"
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
         BeginProperty Column07 
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
               Object.Visible         =   -1  'True
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column05 
               Alignment       =   2
            EndProperty
            BeginProperty Column06 
               Alignment       =   1
            EndProperty
            BeginProperty Column07 
            EndProperty
         EndProperty
      End
      Begin MSDataListLib.DataCombo txt_Logo 
         Bindings        =   "frm_Total_Lanc.frx":4064
         Height          =   286
         Left            =   234
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   546
         Width           =   3289
         _ExtentX        =   6062
         _ExtentY        =   503
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "usl_nome"
         BoundColumn     =   "usl_cod"
         Text            =   ""
         Object.DataMember      =   ""
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
         Left            =   4290
         TabIndex        =   2
         Top             =   510
         Width           =   1035
         _ExtentX        =   1821
         _ExtentY        =   551
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         FontSize        =   7.472
      End
      Begin MSAdodcLib.Adodc adoLogo 
         Height          =   375
         Left            =   1200
         Top             =   285
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2133
         _ExtentY        =   671
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
            Size            =   7.47
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
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
         Left            =   5655
         TabIndex        =   3
         Top             =   510
         Width           =   1035
         _ExtentX        =   1821
         _ExtentY        =   551
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         FontSize        =   7.472
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Logo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   2205
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4110
         TabIndex        =   18
         Top             =   270
         Width           =   2760
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "à"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5250
         TabIndex        =   17
         Top             =   600
         Width           =   480
      End
   End
   Begin Skin_Button.ctr_Button btResumo 
      Height          =   615
      Left            =   7320
      TabIndex        =   6
      Top             =   4785
      Width           =   2295
      _ExtentX        =   4050
      _ExtentY        =   1078
      BTYPE           =   2
      TX              =   "Resumo"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
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
      MICON           =   "frm_Total_Lanc.frx":407A
      PICN            =   "frm_Total_Lanc.frx":4096
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      ForeColor       =   &H00FF0000&
      Height          =   364
      ItemData        =   "frm_Total_Lanc.frx":43B8
      Left            =   5610
      List            =   "frm_Total_Lanc.frx":43C2
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   4890
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   705
      Left            =   2160
      TabIndex        =   21
      Top             =   4710
      Width           =   1335
      Begin VB.OptionButton OpOrient 
         Caption         =   "Paisagem"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   405
         Width           =   1095
      End
      Begin VB.OptionButton OpOrient 
         Caption         =   "Retrato"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   165
         Value           =   -1  'True
         Width           =   1095
      End
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   8760
      Top             =   240
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":43E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":46FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":48D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":4BEE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":4F08
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":5222
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":553C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Total_Lanc.frx":5716
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
End
Attribute VB_Name = "frm_Total_Lanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_RPT As Boolean

   
   
Private Sub bt_Canc_F_Click()
    btResumo.Visible = True
    fr_Resumo.Visible = False
    fr_Principal.Enabled = True
End Sub

Private Sub bt_Pesq_Click()
inicio:
On Error GoTo err1
    
If (CVDate(Txt_DtI) > w_Data_Server - 7 And w_Usu_Tipo = "L") Or w_Usu_Tipo <> "L" Then
     
    Set adoReg.Recordset = ExecuteSQL("SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, SUM(tab_lanc.lnc_vr) AS Total, COUNT(1) AS Qtde_Lanc, tab_lanc.lnc_tipoc, tab_lanc.lnc_nresumo AS Resumo, tab_lanc.lnc_dt_vnd AS Dt, tab_lanc.lnc_tipo AS TIPO, tab_lanc.lnc_tef_pos as TEF_POS FROM tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "') GROUP BY tab_tipo_cartao.tpc_desc, tab_usuario.usl_nome, tab_lanc.lnc_tipoc, tab_lanc.lnc_nresumo, tab_lanc.lnc_dt_vnd, tab_lanc.lnc_tipo, TEF_POS HAVING (tab_usuario.usl_nome LIKE '" & txt_Logo & "%') ORDER BY tab_usuario.usl_nome, tab_lanc.lnc_dt_vnd, tab_tipo_cartao.tpc_desc").Clone 'de.rsSql_Total_Diario.Clone
    adoReg.Recordset.Filter = "TIPO <> 'D'"

Else
    MsgBox "Não é permitido verificar o total diário de mais de uma semana atrás!", vbExclamation
End If

sair:
    Exit Sub
err1:
     w_msg = msgErro(err)
     If w_msg = "" Then
        GoTo inicio
    Else
        MsgBox msgErro(err), vbCritical
    End If
    Resume sair
End Sub



Private Sub bt_RPT_Click()
Dim w_SQL As String
Dim w_Rec As New Recordset
Dim w_Resumo As Boolean

On Error GoTo err1
    w_Resumo = True
    adoReg.Recordset.MoveFirst
    Do While Not adoReg.Recordset.EOF
        If adoReg.Recordset.Fields("resumo") = "" Or adoReg.Recordset.Fields("resumo") = "0" Then
            w_Resumo = False
        End If
        adoReg.Recordset.MoveNext
    Loop




If (w_Usu_Tipo <> "L" Or (w_Usu_Tipo = "L" And CVDate(Txt_DtI) > w_Data_Server - 7)) And w_Resumo = True Then
   
    w_SQL = "SHAPE {SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, tab_forma_pg.fpg_desc AS FormaPG, tab_lanc.lnc_ndoc, tab_lanc.lnc_dt_vnd, tab_lanc.lnc_tx, tab_lanc.lnc_tx_fixo, tab_lanc.lnc_tx_po, tab_lanc.lnc_vr, tab_lanc.lnc_vr_liq, tab_lanc.lnc_num, tab_lanc.lnc_nresumo, tab_lanc.lnc_tef_pos as TEF_POS FROM tab_lanc, tab_usuario, tab_tipo_cartao, tab_forma_pg WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod) AND (tab_usuario.usl_nome = '" & txt_Logo & "') AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "') ORDER BY tab_lanc.lnc_dt_vnd, tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc, tab_forma_pg.fpg_desc}  AS Sql_Lanc_Resumo COMPUTE Sql_Lanc_Resumo BY 'Logo','Cartao', 'TEF_POS'"

    Set w_Rec = ExecuteSQL(w_SQL, , "MSDataShape").Clone
   
    If w_Rec.RecordCount > 0 Then
        Set Rel_Vendas_Resumo.DataSource = w_Rec.Clone
        Rel_Vendas_Resumo.Sections("SecCab").Controls("lbTitulo").Caption = "Relatório das Vs - " & txt_Logo
        
        If OpOrient(0).Value = 0 Then
            Rel_Vendas_Resumo.Orientation = rptOrientLandscape
        Else
            Rel_Vendas_Resumo.Orientation = rptOrientPortrait
        End If
        
               
        Rel_Vendas_Resumo.WindowState = vbMaximized
        Rel_Vendas_Resumo.Show
        
        w_RPT = True
    Else
        MsgBox "Nenhum registro encontrado para gerar o relatório!", vbExclamation
    End If
ElseIf w_Resumo = False Then
    MsgBox "Você precisa inserir o resumo dos cartões, antes de imprimir o resumo!", vbExclamation
Else
    MsgBox "Não é permitido verificar o total diário de mais de uma semana atrás!", vbExclamation
End If
    
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub





Private Sub bt_Sal_F_Click()
On Error GoTo err1
    
   
    
    If vbYes = MsgBox("Será atualizado o resumo do dia " & adoReg.Recordset.Fields("DT") & " do cartão " & txt_cartao & "!", vbQuestion + vbYesNo + vbDefaultButton1, " Logo - " & txt_Logo) Then
        
        
    Dim w_ado As ADODB.Recordset
        'Pega numeros do lançamentos p/ atualizar suas parcelas
        Set w_ado = ExecuteSQL("Select lnc_num from tab_lanc WHERE (lnc_loj = " & txt_Logo.BoundText & " and lnc_tipoc = " & txt_cartao.BoundText & " and lnc_dt_vnd = '" & Format(adoReg.Recordset.Fields("DT"), "yyyy-mm-dd") & "' and lnc_nresumo = '" & adoReg.Recordset.Fields("resumo") & "' and lnc_tipo = '" & adoReg.Recordset.Fields("tipo") & "' and lnc_tef_pos = '" & adoReg.Recordset.Fields("tef_pos") & "')").Clone

        'Atualiza o Resumo das Parcelas
        Do While Not w_ado.EOF
            'Tab_lanc_parc
            Call ExecuteSQL("UPDATE tab_lanc_parc SET lcp_nresumo = '" & txt_NResumo & "' WHERE (lcp_num_lanc = " & w_ado.Fields("lnc_num") & ")")
            w_ado.MoveNext
        Loop
   
        'Tab_lanc
        Call ExecuteSQL("UPDATE tab_lanc SET lnc_nresumo = '" & txt_NResumo & "' WHERE (lnc_loj = " & txt_Logo.BoundText & " and lnc_tipoc = " & txt_cartao.BoundText & " and lnc_dt_vnd = '" & Format(adoReg.Recordset.Fields("DT"), "yyyy-mm-dd") & "' and lnc_nresumo = '" & adoReg.Recordset.Fields("resumo") & "' and lnc_tipo = '" & adoReg.Recordset.Fields("tipo") & "' and lnc_tef_pos = '" & adoReg.Recordset.Fields("tef_pos") & "')", w_RegAf)
    
        If w_RegAf = 0 Then MsgBox "Não foi possível atualizar o Nº de Resumo para este Cartão!", vbCritical
        bt_Canc_F_Click
        bt_Pesq_Click
    
    End If
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub btResumo_Click()
On Error GoTo err1
    

    
    If Not adoReg.Recordset.EOF Then
        btResumo.Visible = False
        fr_Resumo.Visible = True
        txt_NResumo = adoReg.Recordset.Fields("Resumo")
        txt_cartao.BoundText = adoReg.Recordset.Fields("lnc_tipoc")
        txt_NResumo.SetFocus
        fr_Principal.Enabled = False
        txt_dt_vnd = adoReg.Recordset.Fields("Dt")
        
        If adoReg.Recordset.Fields("TEF_POS") = "POS" Then
            Select Case txt_cartao.BoundText
            Case "4":
                If txt_FormaPg_Tipo = "V" Then
                    txt_NResumo = Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Else
                    txt_NResumo = "4" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                End If
            Case "5":
                If txt_FormaPg_Tipo = "V" Then
                    txt_NResumo = "5" & Format(txt_dt_vnd, "yy") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                End If
            Case "7": 'Hipercard
                    txt_NResumo = "H" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
            Case "8": 'Sorocred
                    txt_NResumo = "S" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
            Case "6": 'Amex
                    txt_NResumo = "A" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
            End Select
        Else ' se for TEF
            Select Case txt_cartao.BoundText
                Case "9": 'Hipercard
                     txt_NResumo = "H" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "7": 'Sorocred
                     txt_NResumo = "S" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "3": 'Amex
                     txt_NResumo = "A" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "2": 'Redeshop
                     txt_NResumo = "R" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "21": 'Redeshop/Eletron
                     txt_NResumo = "RE" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "1": 'Credcard
                     txt_NResumo = "C" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "20": 'Credcard/Visa
                     txt_NResumo = "CV" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "5": 'Eletron
                     txt_NResumo = "E" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "16": 'Visa
                     txt_NResumo = "V" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
                Case "22": 'Nelycard
                     txt_NResumo = "N" & Format(txt_dt_vnd, "YY") & Format(txt_dt_vnd, "MM") & Format(txt_dt_vnd, "dd")
            End Select
        End If
    
    End If
    
sair:
    Exit Sub
err1:
    If err.Number = 91 Then
        MsgBox "Preencha os campos e faça a consulta !", vbCritical
        
    Else
     MsgBox msgErro(err), vbCritical
    End If
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
     
    If UCase(w_Usu_Nome) = "PL" Then bt_Rpt_Geral.Visible = True

    If w_Usu = "%" Then
        txt_Logo.Enabled = True

    Else
        txt_Logo.Enabled = False
        txt_Logo = w_Usu
        Grid.Columns(0).Visible = False
        Grid.Left = 850
        Grid.Width = 7741
        bt_Pesq_Click
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

Private Sub List1_Click()
    Grid.SetFocus
End Sub

Private Sub List1_GotFocus()
    Grid.SetFocus
End Sub

Private Sub mnuFechar_Click()
On Error Resume Next
    If w_RPT = True Then
        Unload Rel_Vendas_Resumo
        Unload Rel_Vendas_Resumo_Geral
        w_RPT = False
    Else
        Unload Me
    End If
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    End Select
End Sub




Private Sub Txt_DtI_Validate(Cancel As Boolean)
    If IsDate(Txt_DtI) Then
        Txt_DtF = Txt_DtI
    Else
        Txt_DtI = w_Data_Server
    End If
End Sub

Private Sub txt_Logo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_NResumo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then bt_Sal_F.SetFocus
End Sub
