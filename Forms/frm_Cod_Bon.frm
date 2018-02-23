VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Cod_Bon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Código / Bônus"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   10125
   Icon            =   "frm_Cod_Bon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   10125
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10125
      _ExtentX        =   17859
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
         Left            =   5040
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "Código / Bônus"
         Top             =   120
         Width           =   3135
      End
   End
   Begin rdActiveText.ActiveText txt_vr_cred_acum 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   3960
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Text            =   "0"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_vr_cred 
      Height          =   375
      Left            =   3360
      TabIndex        =   6
      Top             =   3960
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
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
      Text            =   "0"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin VB.TextBox txt_obs 
      Height          =   780
      Left            =   2760
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   4815
   End
   Begin MSAdodcLib.Adodc adoGrid 
      Height          =   330
      Left            =   705
      Top             =   7440
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
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
   Begin MSDataGridLib.DataGrid Grid 
      Bindings        =   "frm_Cod_Bon.frx":27A2
      Height          =   3000
      Left            =   240
      TabIndex        =   16
      Top             =   4680
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   5292
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
      Caption         =   "Lançamentos do Dia"
      ColumnCount     =   10
      BeginProperty Column00 
         DataField       =   "usl_cod"
         Caption         =   "usl_cod"
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
         DataField       =   "dt_vnd"
         Caption         =   "Data"
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
         DataField       =   "vr_vnd"
         Caption         =   "Vnd"
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
         DataField       =   "vr_bonus"
         Caption         =   "Bônus"
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
         DataField       =   "vr_acum"
         Caption         =   "V.Acum"
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
         DataField       =   "dt_cad"
         Caption         =   "dt_cad"
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
         DataField       =   "codigo"
         Caption         =   "codigo"
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
         DataField       =   "obs"
         Caption         =   "Obs"
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
      BeginProperty Column08 
         DataField       =   "vr_cred"
         Caption         =   "V.Cred"
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
      BeginProperty Column09 
         DataField       =   "vr_cred_acum"
         Caption         =   "V.Cred.Acum"
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
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column01 
            Alignment       =   2
            Object.Visible         =   -1  'True
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Object.Visible         =   -1  'True
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            ColumnWidth     =   1094,74
         EndProperty
         BeginProperty Column05 
            Object.Visible         =   0   'False
            ColumnWidth     =   1995,024
         EndProperty
         BeginProperty Column06 
            Object.Visible         =   0   'False
            ColumnWidth     =   1739,906
         EndProperty
         BeginProperty Column07 
            ColumnWidth     =   2294,929
         EndProperty
         BeginProperty Column08 
            Alignment       =   2
            ColumnWidth     =   900,284
         EndProperty
         BeginProperty Column09 
            Alignment       =   2
            ColumnWidth     =   1200,189
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adoLogo 
      Height          =   330
      Left            =   3960
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
   Begin MSDataListLib.DataCombo txt_Logo 
      Bindings        =   "frm_Cod_Bon.frx":27B8
      Height          =   315
      Left            =   3195
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1440
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "usl_nome"
      BoundColumn     =   "usl_cod"
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
      Left            =   6030
      TabIndex        =   1
      Top             =   1440
      Width           =   1275
      _ExtentX        =   2249
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
      MaxLength       =   10
      TextMask        =   1
      RawText         =   1
      Mask            =   "##/##/####"
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_Valor_Vnd 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1046
         SubFormatType   =   1
      EndProperty
      Height          =   315
      Left            =   3195
      TabIndex        =   2
      Top             =   1890
      Width           =   1695
      _ExtentX        =   2990
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
      Text            =   "0"
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_bonus 
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
      Left            =   6030
      TabIndex        =   3
      Top             =   1890
      Width           =   1275
      _ExtentX        =   2249
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
      Text            =   "0"
      RawText         =   0
      eAuto           =   1
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin rdActiveText.ActiveText txt_acum 
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
      Left            =   4845
      TabIndex        =   4
      Top             =   2370
      Visible         =   0   'False
      Width           =   1305
      _ExtentX        =   2302
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
      Text            =   "R$ 0,00"
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8,25
   End
   Begin MSDataListLib.DataCombo txt_Usu_AC 
      Bindings        =   "frm_Cod_Bon.frx":27CE
      Height          =   315
      Left            =   3315
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "usl_ac"
      BoundColumn     =   "usl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin Skin_Button.ctr_Button bt_Exc_F 
      Height          =   1500
      Left            =   9000
      TabIndex        =   17
      TabStop         =   0   'False
      ToolTipText     =   "Excluir Lançamento selecionado"
      Top             =   6180
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   2646
      BTYPE           =   2
      TX              =   "&Excluir"
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
      MICON           =   "frm_Cod_Bon.frx":27E4
      PICN            =   "frm_Cod_Bon.frx":2800
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button bt_Edit 
      Height          =   1485
      Left            =   9000
      TabIndex        =   19
      TabStop         =   0   'False
      ToolTipText     =   "Editar Lançamento selecionado"
      Top             =   4695
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   2619
      BTYPE           =   2
      TX              =   "&Editar"
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
      MICON           =   "frm_Cod_Bon.frx":2B1A
      PICN            =   "frm_Cod_Bon.frx":2B36
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin rdActiveText.ActiveText ActiveText1 
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
      Left            =   0
      TabIndex        =   20
      Top             =   0
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
      FontSize        =   8,25
   End
   Begin VB.Label lbl_vr_cred_acum 
      Caption         =   "V.Cred.Acum"
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
      Left            =   4920
      TabIndex        =   22
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "V.Cred"
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
      Left            =   2640
      TabIndex        =   21
      Top             =   4080
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Observação"
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
      Left            =   2760
      TabIndex        =   18
      Top             =   2805
      Width           =   1590
   End
   Begin VB.Shape Shape2 
      Height          =   3015
      Left            =   240
      Top             =   4680
      Width           =   9615
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   2400
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label lbAcum 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Acumulado"
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
      Left            =   3300
      TabIndex        =   14
      Top             =   2445
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Bônus"
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
      Left            =   4485
      TabIndex        =   13
      Top             =   1965
      Width           =   1455
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Vnd"
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
      Left            =   1560
      TabIndex        =   12
      Top             =   1965
      Width           =   1455
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Data"
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
      Left            =   4560
      TabIndex        =   11
      Top             =   1560
      Width           =   1305
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
      Left            =   1890
      TabIndex        =   10
      Top             =   1515
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
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":2E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":316A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":3344
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":365E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":3978
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":3C92
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":3FAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cod_Bon.frx":4186
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
   Begin VB.Menu mnuSep01 
      Caption         =   "         |"
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
   Begin VB.Menu mnuSep02 
      Caption         =   "|"
   End
End
Attribute VB_Name = "frm_Cod_Bon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim wCodigo
Dim wNovo As Boolean



Private Sub bt_Edit_Click()

    txt_Logo.Enabled = False
    bt_Edit.Enabled = False
    wNovo = False
    
    wCodigo = adoGrid.Recordset.Fields("codigo")
    
    txt_acum = adoGrid.Recordset.Fields("vr_acum")
    txt_bonus = adoGrid.Recordset.Fields("vr_bonus")
    txt_dt_vnd = adoGrid.Recordset.Fields("dt_vnd")
    txt_Logo.BoundText = adoGrid.Recordset.Fields("usl_cod")
    txt_obs = adoGrid.Recordset.Fields("obs")
    txt_Valor_Vnd = adoGrid.Recordset.Fields("vr_vnd")
    txt_vr_cred = adoGrid.Recordset.Fields("vr_cred")
    txt_vr_cred_acum = adoGrid.Recordset.Fields("vr_cred_acum")
    

    
End Sub

Private Sub bt_Exc_F_Click()
    If vbYes = MsgBox("Deseja realmente excluir?", vbQuestion + vbYesNo) Then
        ExecuteSQL "DELETE FROM tab_vnds_bonus Where codigo = " & adoGrid.Recordset.Fields("codigo")
        mnuCancelar_Click
    End If
End Sub

Private Sub Form_Load()
On Error GoTo err1

    MDI.TBar.Visible = False

    Set adoLogo.Recordset = w_ado_Logo.Clone

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2)
    
    mnuCancelar_Click

    'Joga o Logo do usuario logado
    txt_Logo.BoundText = w_Usu_Cod
    txt_Usu_AC.BoundText = txt_Logo.BoundText
    'Habilita ou desabilita Logo
    txt_Logo.Enabled = Not (w_Usu_Tipo = "L")
    wNovo = True
   

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub




Private Sub mnuCancelar_Click()
On Error Resume Next
    
    If txt_dt_vnd = "" Then txt_dt_vnd = w_Data_Server
    txt_Valor_Vnd = 0
    txt_bonus = 0
    txt_acum = 0
    txt_obs = ""
    txt_vr_cred = 0
    txt_vr_cred_acum = 0
    
    If txt_Logo.Enabled Then
        txt_Logo.SetFocus
    Else
        txt_dt_vnd.SetFocus
    End If
    
    If w_Usu_Tipo = "L" Then
        bt_Edit.Visible = False
    Else
        bt_Edit.Visible = True
        bt_Edit.Enabled = True
        txt_Logo.Enabled = True
    End If
    
    If Not txt_Logo.BoundText = "" And Not w_Usu_Tipo = "L" Then
        
        Set adoGrid.Recordset = ExecuteSQL("SELECT * FROM tab_vnds_bonus WHERE usl_cod = " & txt_Logo.BoundText & " AND dt_cad >= '" & Format(w_Data_Server - 3, "yyyy-mm-dd") & "'").Clone
    
    ElseIf Not txt_Logo.BoundText = "" And w_Usu_Tipo = "L" Then
        Set adoGrid.Recordset = ExecuteSQL("SELECT * FROM tab_vnds_bonus WHERE usl_cod = " & txt_Logo.BoundText & " AND dt_cad >= '" & Format(w_Data_Server, "yyyy-mm-dd") & "'").Clone
    End If
End Sub

Private Sub mnuFechar_Click()
    Unload Me
End Sub

Private Sub mnuNovo_Click()
    wNovo = True
    mnuCancelar_Click
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


Private Sub txt_acum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then KeyAscii = 46
End Sub


Private Sub txt_bonus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Then KeyAscii = 46
End Sub



Private Sub txt_dt_vnd_Validate(Cancel As Boolean)
On Error GoTo sair
    If Not ((Day(CVDate(txt_dt_vnd)) = 1) Or (Day(CVDate(txt_dt_vnd)) = 30) Or (Day(CVDate(txt_dt_vnd)) = 31)) Then
        lbAcum.Visible = (txt_Usu_AC = "S")
        txt_acum.Visible = lbAcum.Visible
        lbl_vr_cred_acum.Visible = txt_acum.Visible
        txt_vr_cred_acum.Visible = lbl_vr_cred_acum.Visible
    Else
        If Day(CVDate(w_Data_Server)) = 1 Or Day(CVDate(w_Data_Server)) = 30 Or Day(CVDate(w_Data_Server)) = 31 Then
            'Or (Day(w_Data_Server) = 1 And Weekday(w_Data_Server, vbSunday) = 2)
            lbAcum.Visible = True
            txt_acum.Visible = lbAcum.Visible
            lbl_vr_cred_acum.Visible = txt_acum.Visible
            txt_vr_cred_acum.Visible = lbl_vr_cred_acum.Visible
        Else
            lbAcum.Visible = (txt_Usu_AC = "S")
            txt_acum.Visible = lbAcum.Visible
            lbl_vr_cred_acum.Visible = txt_acum.Visible
            txt_vr_cred_acum.Visible = lbl_vr_cred_acum.Visible
        End If
    End If
    
    If w_Usu_Tipo = "A" Then
    
        lbAcum.Visible = True
        txt_acum.Visible = True
        lbl_vr_cred_acum.Visible = True
        txt_vr_cred_acum.Visible = True
    
    End If

sair:
End Sub

Private Sub txt_Logo_Change()
    
    txt_Usu_AC.BoundText = txt_Logo.BoundText
    txt_dt_vnd_Validate (True)
        'Se o Proximo dia for dia 1 or Se o Dia Atual for dia 1 e for Segunda - Feira, então ...
        '.. libera o txt_acum
                    '    If Day(w_Data_Server + 1) = 1 Or _
                    '      (Day(w_Data_Server) = 1 And Weekday(w_Data_Server, vbSunday) = 2) Then
                    '
                    '        lbAcum.Visible = True
                    '        txt_acum.Visible = lbAcum.Visible
                    '        lbl_vr_cred_acum.Visible = txt_acum.Visible
                    '        txt_vr_cred_acum.Visible = lbl_vr_cred_acum.Visible
                    '        If Day(w_Data_Server + 1) = 1 Then
                    '            txt_dt_vnd = w_Data_Server
                    '        Else
                    '            txt_dt_vnd = w_Data_Server - 1
                    '        End If
                    '    Else
                    '
                    '        lbAcum.Visible = (txt_Usu_AC = "S")
                    '        txt_acum.Visible = lbAcum.Visible
                    '        lbl_vr_cred_acum.Visible = txt_acum.Visible
                    '        txt_vr_cred_acum.Visible = lbl_vr_cred_acum.Visible
                    '
                    '    End If

'    lbAcum.Visible = (txt_Usu_AC = "S")
'    txt_acum.Visible = lbAcum.Visible
    mnuCancelar_Click
    
    
End Sub

Private Sub txt_Logo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then KeyEnter KeyCode
End Sub



Sub Salvar()
On Error GoTo err1

If Left(txt_obs, 1) = Chr(13) And Len(txt_obs) <= 2 Then txt_obs = ""

If CDbl(txt_Valor_Vnd) > 0 Or txt_obs <> "" Then
    
    If wNovo = True Then
        
        wCod = ExecuteSQL("SELECT max(CODIGO) FROM tab_vnds_bonus").Fields(0)
        If IsNull(wCod) Then wCod = 0
        wCod = wCod + 1
            
        'Pega a Qtde de Bonus Salvo... Para saber se já foi salvo ou não
        wQtCad = ExecuteSQL("SELECT count(CODIGO) FROM tab_vnds_bonus Where usl_cod = '" & txt_Logo.BoundText & "' and dt_vnd = '" & Format(txt_dt_vnd, "yyyy-mm-dd") & "'").Fields(0)
        
        If wQtCad = 0 Then
            Call ExecuteSQL("INSERT INTO tab_vnds_bonus (usl_cod, dt_vnd, vr_vnd, vr_bonus, vr_acum, dt_cad, codigo, obs, vr_cred, vr_cred_acum) " _
                          & "VALUES ('" & txt_Logo.BoundText & "','" & Format(txt_dt_vnd, "yyyy-mm-dd") & "','" & txt_Valor_Vnd & "','" & txt_bonus & "','" & txt_acum & "', NOW(),'" & wCod & "','" & txt_obs & "','" & txt_vr_cred & "','" & txt_vr_cred_acum & "')", wRegAf)
        Else
            MsgBox "O código/bonus do dia " & txt_dt_vnd & " já foi lançado!", vbCritical
            Exit Sub
        End If
        
        If Not wRegAf = 1 Then MsgBox "Não foi possível salvar!" & Chr(13) & "Tente novamente!", vbCritical
        mnuCancelar_Click
    
    Else
        
        wSQL = "Update tab_vnds_bonus set dt_vnd = '" & Format(txt_dt_vnd, "yyyy-mm-dd") & "', " & _
        "vr_vnd = '" & txt_Valor_Vnd & "', vr_bonus = '" & txt_bonus & "', vr_acum = '" & txt_acum & "', " & _
        "obs = '" & txt_obs & "', vr_cred = '" & txt_vr_cred & "',vr_cred_acum = '" & txt_vr_cred_acum & "' " & _
        "Where codigo = " & wCodigo & ""
        Call ExecuteSQL(wSQL, wRegAf)
        If Not wRegAf = 1 Then MsgBox "Não foi possível salvar!" & Chr(13) & "Tente novamente!", vbCritical
        mnuCancelar_Click
        
    End If
    
Else
    If CDbl(txt_Valor_Vnd) = 0 Then
        MsgBox "Preencha o valor vnd!", vbExclamation
    ElseIf CDbl(txt_bonus) = 0 Then
        MsgBox "Preencha o valor do bônus!", vbExclamation
    End If
End If
    
    
sair:
    Exit Sub
err1:
    If err.Number = -2147217900 Then
        MsgBox "O código/bonus do dia " & txt_dt_vnd & " já foi lançado!", vbCritical
    Else
On Error Resume Next
     MsgBox err.Description, vbCritical
    End If
    Resume sair
End Sub




Private Sub txt_obs_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then Salvar
    End If
End Sub



Private Sub txt_Valor_Vnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 46 Then KeyAscii = 0
End Sub



Private Sub txt_vr_cred_acum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 46 Then KeyAscii = 0
End Sub



Private Sub txt_vr_cred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 44 Or KeyAscii = 46 Then KeyAscii = 0
End Sub
