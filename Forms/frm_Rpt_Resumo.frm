VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Rpt_Resumo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "R. Resumo"
   ClientHeight    =   5434
   ClientLeft      =   39
   ClientTop       =   611
   ClientWidth     =   11531
   Icon            =   "frm_Rpt_Resumo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5434
   ScaleWidth      =   11531
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   871
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11531
      _ExtentX        =   20344
      _ExtentY        =   1534
      ButtonWidth     =   1376
      ButtonHeight    =   1429
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
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "R. Resumo"
         Top             =   120
         Width           =   6735
      End
   End
   Begin VB.Frame fr_Principal 
      Height          =   4470
      Left            =   120
      TabIndex        =   7
      Top             =   870
      Width           =   11235
      Begin rdActiveText.ActiveText txt_Bto 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   9240
         TabIndex        =   18
         Top             =   4080
         Width           =   765
         _ExtentX        =   1342
         _ExtentY        =   551
         Alignment       =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0,00"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   7.472
      End
      Begin VB.Frame Painel 
         Caption         =   "Agrupar por: "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   5880
         TabIndex        =   20
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton op2 
            Caption         =   "Logo"
            Height          =   225
            Left            =   255
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   540
            Width           =   1095
         End
         Begin VB.OptionButton op1 
            Caption         =   "Cartão"
            Height          =   225
            Left            =   255
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   285
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin MSAdodcLib.Adodc adoGrid 
         Height          =   375
         Left            =   7680
         Top             =   3000
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
         Caption         =   ""
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Bindings        =   "frm_Rpt_Resumo.frx":27A2
         Height          =   2820
         Left            =   5880
         TabIndex        =   16
         Top             =   1200
         Width           =   5160
         _ExtentX        =   9106
         _ExtentY        =   4984
         _Version        =   393216
         Cols            =   5
         AllowBigSelection=   0   'False
         FillStyle       =   1
         ScrollBars      =   2
         SelectionMode   =   1
         RowSizingMode   =   1
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
      Begin MSAdodcLib.Adodc adoCartao 
         Height          =   330
         Left            =   3675
         Top             =   1920
         Visible         =   0   'False
         Width           =   1215
         _ExtentX        =   2133
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
      Begin VB.ListBox List_Cartao 
         Appearance      =   0  'Flat
         Height          =   3146
         ItemData        =   "frm_Rpt_Resumo.frx":27B8
         Left            =   2520
         List            =   "frm_Rpt_Resumo.frx":27BA
         Style           =   1  'Checkbox
         TabIndex        =   3
         Top             =   1200
         Width           =   2505
      End
      Begin VB.ListBox List_loja 
         Appearance      =   0  'Flat
         Height          =   3731
         ItemData        =   "frm_Rpt_Resumo.frx":27BC
         Left            =   240
         List            =   "frm_Rpt_Resumo.frx":27BE
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   525
         Width           =   1335
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
         Left            =   2490
         TabIndex        =   1
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
         Left            =   960
         Top             =   120
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
         Left            =   3975
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
      Begin Skin_Button.ctr_Button bt_Rpt_Geral 
         Height          =   615
         Left            =   9360
         TabIndex        =   4
         Top             =   240
         Width           =   1650
         _ExtentX        =   2899
         _ExtentY        =   1078
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":27C0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_STodos 
         Height          =   495
         Left            =   1560
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona todos"
         Top             =   510
         Width           =   525
         _ExtentX        =   935
         _ExtentY        =   863
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":27DC
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_RTodos 
         Height          =   495
         Left            =   1560
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Retira Selecão de todos"
         Top             =   1080
         Width           =   525
         _ExtentX        =   935
         _ExtentY        =   863
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":27F8
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_STodos_Cart 
         Height          =   495
         Left            =   5025
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Seleciona todos"
         Top             =   1200
         Width           =   525
         _ExtentX        =   935
         _ExtentY        =   863
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":2814
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_RTodos_Cart 
         Height          =   495
         Left            =   5025
         TabIndex        =   15
         TabStop         =   0   'False
         ToolTipText     =   "Retira Selecão de todos"
         Top             =   1770
         Width           =   525
         _ExtentX        =   935
         _ExtentY        =   863
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":2830
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button bt_Consultar 
         Height          =   615
         Left            =   7560
         TabIndex        =   17
         Top             =   240
         Width           =   1650
         _ExtentX        =   2899
         _ExtentY        =   1078
         BTYPE           =   2
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
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
         MICON           =   "frm_Rpt_Resumo.frx":284C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin rdActiveText.ActiveText txt_Liq 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0,00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   1
         EndProperty
         Height          =   315
         Left            =   10020
         TabIndex        =   19
         Top             =   4080
         Width           =   765
         _ExtentX        =   1342
         _ExtentY        =   551
         Alignment       =   1
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.4717
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0,00"
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   7.472
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cartão :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2520
         TabIndex        =   11
         Top             =   960
         Width           =   2760
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
         TabIndex        =   10
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
         Left            =   2310
         TabIndex        =   9
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
         Left            =   3510
         TabIndex        =   8
         Top             =   600
         Width           =   480
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
            Picture         =   "frm_Rpt_Resumo.frx":2868
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":2B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":2D5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":3076
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":3390
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":36AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":39C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Resumo.frx":3B9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
End
Attribute VB_Name = "frm_Rpt_Resumo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_RPT As Boolean

   
   
Private Sub bt_Consultar_Click()
Dim w_SQL As String
Dim w_Rec As New Recordset

On Error GoTo err1

    w_logo = ""

    'Pega o Codigo de todas as lojas selecionadas
    For i = 0 To List_loja.ListCount - 1
        If List_loja.Selected(i) = True Then
            adoLogo.Recordset.MoveFirst
            Call adoLogo.Recordset.Move(i)
            w_logo = w_logo & IIf(Len(w_logo) > 0, ",", "") & adoLogo.Recordset.Fields(0)
        End If
    Next i
    
    w_cartao = ""
    'Pega o Codigo de Todos os Cartões Selecionados
    For i = 0 To List_cartao.ListCount - 1
        If List_cartao.Selected(i) = True Then
            adoCartao.Recordset.MoveFirst
            Call adoCartao.Recordset.Move(i)
            w_cartao = w_cartao & IIf(Len(w_cartao) > 0, ",", "") & adoCartao.Recordset.Fields(0)
        End If
    Next i


If Len(w_cartao) > 0 And Len(w_logo) > 0 Then
   
    If op1.Value = True Then
        w_SQL = "SELECT tab_tipo_cartao.tpc_desc AS Cartao, tab_usuario.usl_nome AS Logo, ROUND(Sum(tab_lanc.lnc_vr)) as Vr, ROUND(Sum(tab_lanc.lnc_vr_liq)) as VrLiq FROM tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_loj IN (" & w_logo & ")) AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' and tab_tipo_cartao.tpc_cod IN (" & w_cartao & ")) GROUP BY CARTAO, LOGO ORDER BY tab_tipo_cartao.tpc_desc, tab_usuario.usl_nome"
    Else
        w_SQL = "SELECT tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartao, ROUND(Sum(tab_lanc.lnc_vr)) as Vr, ROUND(Sum(tab_lanc.lnc_vr_liq)) as VrLiq FROM tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_loj IN (" & w_logo & ")) AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' and tab_tipo_cartao.tpc_cod IN (" & w_cartao & ")) GROUP BY LOGO, CARTAO ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc"
    End If
    
    Set w_Rec = ExecuteSQL(w_SQL, , "MSDataShape").Clone
   
    If w_Rec.RecordCount > 0 Then
        Set adoGrid.Recordset = w_Rec.Clone
        Formata_Grid
        txt_Liq = 0
        txt_Bto = 0
        Do While Not adoGrid.Recordset.EOF
            txt_Liq = txt_Liq + adoGrid.Recordset("VRLIQ")
            txt_Bto = txt_Bto + adoGrid.Recordset("VR")
            adoGrid.Recordset.MoveNext
        Loop
        
    Else
        MsgBox "Nenhum registro encontrado para gerar o relatório!", vbExclamation
    End If
    
ElseIf Len(w_logo) = 0 And Len(w_cartao) = 0 Then
    MsgBox "Escolha a loja e o cartão desejado!", vbExclamation
ElseIf Len(w_logo) = 0 Then
    MsgBox "Escolha a loja desejada!", vbExclamation
ElseIf Len(w_cartao) = 0 Then
    MsgBox "Escolha o cartão desejado!", vbExclamation
End If
    
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub bt_Rpt_Geral_Click()
Dim w_SQL As String
Dim w_Rec As New Recordset

On Error GoTo err1

    w_logo = ""

    'Pega o Codigo de todas as lojas selecionadas
    For i = 0 To List_loja.ListCount - 1
        If List_loja.Selected(i) = True Then
            adoLogo.Recordset.MoveFirst
            Call adoLogo.Recordset.Move(i)
            w_logo = w_logo & IIf(Len(w_logo) > 0, ",", "") & adoLogo.Recordset.Fields(0)
        End If
    Next i
    
    w_cartao = ""
    'Pega o Codigo de Todos os Cartões Selecionados
    For i = 0 To List_cartao.ListCount - 1
        If List_cartao.Selected(i) = True Then
            adoCartao.Recordset.MoveFirst
            Call adoCartao.Recordset.Move(i)
            w_cartao = w_cartao & IIf(Len(w_cartao) > 0, ",", "") & adoCartao.Recordset.Fields(0)
        End If
    Next i


If Len(w_cartao) > 0 And Len(w_logo) > 0 Then
   
    If op1.Value = True Then
        w_SQL = "SHAPE {SELECT tab_tipo_cartao.tpc_desc AS Cartao, tab_usuario.usl_nome AS Logo,  Sum(tab_lanc.lnc_vr) as Vr, Sum(tab_lanc.lnc_vr_liq) as VrLiq FROM tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_loj IN (" & w_logo & ")) AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' and tab_tipo_cartao.tpc_cod IN (" & w_cartao & ")) GROUP BY CARTAO, LOGO ORDER BY tab_tipo_cartao.tpc_desc, tab_usuario.usl_nome}  AS Sql_Lanc_Resumo COMPUTE Sql_Lanc_Resumo BY 'Cartao'"
    Else
        w_SQL = "SHAPE {SELECT tab_tipo_cartao.tpc_desc AS Cartao, tab_usuario.usl_nome AS Logo,  Sum(tab_lanc.lnc_vr) as Vr, Sum(tab_lanc.lnc_vr_liq) as VrLiq FROM tab_lanc, tab_usuario, tab_tipo_cartao WHERE (tab_lanc.lnc_loj = tab_usuario.usl_cod) AND (tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_lanc.lnc_loj IN (" & w_logo & ")) AND (tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "') AND (tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' and tab_tipo_cartao.tpc_cod IN (" & w_cartao & ")) GROUP BY LOGO, CARTAO ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc}  AS Sql_Lanc_Resumo COMPUTE Sql_Lanc_Resumo BY 'Logo'"
    End If
    Set w_Rec = ExecuteSQL(w_SQL, , "MSDataShape").Clone
   
    If w_Rec.RecordCount > 0 Then
    
        If op1.Value = True Then
            Set Rel_Vendas_Resumo_Geral_Cartao.DataSource = w_Rec.Clone
            
            Rel_Vendas_Resumo_Geral_Cartao.WindowState = vbMaximized
            Rel_Vendas_Resumo_Geral_Cartao.Show
        Else
            Set Rel_Vendas_Resumo_Geral_Logo.DataSource = w_Rec.Clone
            
            Rel_Vendas_Resumo_Geral_Logo.WindowState = vbMaximized
            Rel_Vendas_Resumo_Geral_Logo.Show
        End If
        w_RPT = True
    Else
        MsgBox "Nenhum registro encontrado para gerar o relatório!", vbExclamation
    End If
    
ElseIf Len(w_logo) = 0 And Len(w_cartao) = 0 Then
    MsgBox "Escolha a loja e o cartão desejado!", vbExclamation
ElseIf Len(w_logo) = 0 Then
    MsgBox "Escolha a loja desejada!", vbExclamation
ElseIf Len(w_cartao) = 0 Then
    MsgBox "Escolha o cartão desejado!", vbExclamation
End If
    
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub





Private Sub bt_RTodos_Click()
    For i = List_loja.ListCount - 1 To 0 Step -1
        List_loja.Selected(i) = False
    Next i
    List_loja.Selected(0) = False
End Sub

Private Sub bt_STodos_Click()
    For i = List_loja.ListCount - 1 To 0 Step -1
        List_loja.Selected(i) = True
    Next i
    List_loja.Selected(0) = True
End Sub

Private Sub bt_RTodos_Cart_Click()
    For i = List_cartao.ListCount - 1 To 0 Step -1
        List_cartao.Selected(i) = False
    Next i
    List_cartao.Selected(0) = False
End Sub

Private Sub bt_STodos_Cart_Click()
    For i = List_cartao.ListCount - 1 To 0 Step -1
        List_cartao.Selected(i) = True
    Next i
    List_cartao.Selected(0) = True
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
    txt_Situacao = "Todos"
    
    'If UCase(w_Usu_Nome) = "PL" Then bt_Rpt_Geral.Visible = True

    
    'monta lista das lojas
    For i = 1 To adoLogo.Recordset.RecordCount
        Call List_loja.AddItem(adoLogo.Recordset.Fields("USL_NOME"), List_loja.ListCount)
        adoLogo.Recordset.MoveNext
    Next i
    
    'monta lista dos cartões
    For i = 1 To adoCartao.Recordset.RecordCount
        Call List_cartao.AddItem(adoCartao.Recordset.Fields("tpc_desc"), List_cartao.ListCount)
        adoCartao.Recordset.MoveNext
    Next i
    
    
    Formata_Grid
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








Private Sub List_Cartao_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then bt_Rpt_Geral.SetFocus
End Sub

Private Sub List_loja_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Txt_DtI.SetFocus
End Sub

Private Sub mnuFechar_Click()
    If w_RPT = True Then
        If op1.Value = True Then
            Unload Rel_Vendas_Resumo_Geral_Cartao
        Else
            Unload Rel_Vendas_Resumo_Geral_Logo
        End If
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







Sub Formata_Grid()
    Grid.Visible = False
    Grid.MergeCells = flexMergeRestrictColumns
    '*** Colunas p/ Agrupar/Mesclar
    Grid.MergeCol(1) = True

    Grid.ColWidth(0) = 200
    
    If op1.Value = True Then
        Grid.ColWidth(1) = 2500
        Grid.ColWidth(2) = 700
    Else
        Grid.ColWidth(2) = 2500
        Grid.ColWidth(1) = 700
    End If
    
    Grid.ColWidth(3) = 700
    Grid.ColWidth(4) = 700
    
    Grid.TextArray(1) = "Cartão"
    Grid.TextArray(2) = "Logo"
    Grid.TextArray(3) = "Bruto"
    Grid.TextArray(4) = "Líquido"
    
    Grid.ColAlignment(2) = 3
    Grid.ColAlignment(3) = 6
    Grid.ColAlignment(4) = 6
    
    Grid.Redraw = True
    Grid.Refresh
    Grid.Visible = True
    
    
End Sub
