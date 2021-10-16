VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Usuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuários"
   ClientHeight    =   7007
   ClientLeft      =   39
   ClientTop       =   611
   ClientWidth     =   7085
   Icon            =   "frm_Usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7007
   ScaleWidth      =   7085
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   819
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7085
      _ExtentX        =   12508
      _ExtentY        =   1438
      ButtonWidth     =   1482
      ButtonHeight    =   1376
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   8
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
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   3
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
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Excluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
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
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "Usuários"
         Top             =   120
         Width           =   1935
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5535
      Left            =   120
      TabIndex        =   17
      Top             =   960
      Width           =   6615
      _ExtentX        =   11670
      _ExtentY        =   9753
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ficha Individual"
      TabPicture(0)   =   "frm_Usuario.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFieldLabel(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Shape1(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblFieldLabel(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFieldLabel(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFieldLabel(1)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label3"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblFieldLabel(12)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lblFieldLabel(13)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblFieldLabel(14)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblFieldLabel(15)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblFieldLabel(16)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblFieldLabel(17)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblFieldLabel(18)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txt_Tipo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txt_PassConf"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_Pass"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_Usu"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_grupo"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_ac"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_ordem"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_Rpt"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txt_versao"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txt_versao_login"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Text2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).ControlCount=   27
      TabCaption(1)   =   "Grade"
      TabPicture(1)   =   "frm_Usuario.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Painel"
      Tab(1).Control(1)=   "Grid"
      Tab(1).ControlCount=   2
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         DataField       =   "usl_windows"
         DataSource      =   "adoReg"
         Height          =   285
         Left            =   2925
         TabIndex        =   41
         Top             =   4995
         Width           =   1980
      End
      Begin VB.TextBox txt_versao_login 
         Alignment       =   2  'Center
         DataField       =   "usl_versao_login"
         DataSource      =   "adoReg"
         Height          =   285
         Left            =   4950
         TabIndex        =   10
         Top             =   4605
         Width           =   1020
      End
      Begin VB.TextBox txt_versao 
         Alignment       =   2  'Center
         DataField       =   "usl_versao"
         DataSource      =   "adoReg"
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   4605
         Width           =   1020
      End
      Begin VB.ComboBox txt_Rpt 
         DataField       =   "usl_Rpt"
         DataSource      =   "adoReg"
         Height          =   315
         ItemData        =   "frm_Usuario.frx":27DA
         Left            =   2565
         List            =   "frm_Usuario.frx":27E4
         TabIndex        =   6
         Top             =   3360
         Width           =   660
      End
      Begin VB.TextBox txt_ordem 
         DataField       =   "usl_ordem"
         DataSource      =   "adoReg"
         Height          =   285
         Left            =   5325
         TabIndex        =   5
         Top             =   2760
         Width           =   660
      End
      Begin VB.ComboBox txt_ac 
         DataField       =   "usl_ac"
         DataSource      =   "adoReg"
         Height          =   315
         ItemData        =   "frm_Usuario.frx":27EE
         Left            =   5325
         List            =   "frm_Usuario.frx":27F8
         TabIndex        =   7
         Top             =   3360
         Width           =   660
      End
      Begin VB.ComboBox txt_grupo 
         DataField       =   "usl_grupo"
         DataSource      =   "adoReg"
         Height          =   315
         ItemData        =   "frm_Usuario.frx":2802
         Left            =   2565
         List            =   "frm_Usuario.frx":2815
         TabIndex        =   4
         Top             =   2760
         Width           =   660
      End
      Begin VB.TextBox txt_Usu 
         DataField       =   "usl_nome"
         DataSource      =   "adoReg"
         Height          =   285
         Left            =   3525
         TabIndex        =   1
         Top             =   840
         Width           =   1620
      End
      Begin VB.TextBox txt_Pass 
         DataField       =   "usl_pwd"
         DataSource      =   "adoReg"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3525
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1590
         Width           =   1620
      End
      Begin VB.TextBox txt_PassConf 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   3525
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2055
         Width           =   1620
      End
      Begin VB.ComboBox txt_Tipo 
         DataField       =   "usl_tipo"
         DataSource      =   "adoReg"
         Height          =   315
         ItemData        =   "frm_Usuario.frx":2828
         Left            =   3525
         List            =   "frm_Usuario.frx":2838
         TabIndex        =   8
         Top             =   3855
         Width           =   660
      End
      Begin VB.Frame Painel 
         Caption         =   "Tipo de Filtro"
         Height          =   930
         Left            =   -74040
         TabIndex        =   18
         Top             =   360
         Width           =   3870
         Begin VB.OptionButton Usu_Nome 
            Caption         =   "Nome"
            Height          =   255
            Left            =   840
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   600
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton usu_tipo 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   2160
            TabIndex        =   13
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.TextBox txt_filtro 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   ">"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
            DataSource      =   "adoReg"
            Height          =   285
            Left            =   840
            TabIndex        =   11
            Top             =   240
            Width           =   2175
         End
         Begin Skin_Button.ctr_Button btFiltrar 
            Height          =   735
            Left            =   3045
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   150
            Width           =   735
            _ExtentX        =   1294
            _ExtentY        =   1294
         End
         Begin VB.Label lblFieldLabel 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Digite:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.83
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   120
            TabIndex        =   19
            Top             =   285
            Width           =   690
         End
      End
      Begin MSDataGridLib.DataGrid Grid 
         Bindings        =   "frm_Usuario.frx":2848
         Height          =   3975
         Left            =   -74880
         TabIndex        =   15
         Top             =   1320
         Width           =   6375
         _ExtentX        =   11238
         _ExtentY        =   7021
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
            Name            =   "Arial"
            Size            =   6.7925
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   7
         BeginProperty Column00 
            DataField       =   "usl_cod"
            Caption         =   "Cód."
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
            DataField       =   "usl_nome"
            Caption         =   "Nome"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   ">"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1046
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "usl_grupo"
            Caption         =   "Grupo"
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
            DataField       =   "usl_ordem"
            Caption         =   "Ordem"
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
            DataField       =   "usl_ac"
            Caption         =   "Acum."
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
            DataField       =   "usl_rpt"
            Caption         =   "Rpt"
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
            DataField       =   "usl_Tipo"
            Caption         =   "Tipo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   ">"
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
               ColumnAllowSizing=   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Windows:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   18
         Left            =   1680
         TabIndex        =   42
         Top             =   5010
         Width           =   1155
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ver. Login:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   17
         Left            =   3540
         TabIndex        =   40
         Top             =   4620
         Width           =   1320
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ver. EXE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   16
         Left            =   825
         TabIndex        =   39
         Top             =   4620
         Width           =   1200
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "S - SuperV )"
         Height          =   255
         Left            =   4500
         TabIndex        =   38
         Top             =   4200
         Width           =   975
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Rpt Cód/Bônus ?:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   585
         TabIndex        =   37
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   4320
         TabIndex        =   36
         Top             =   2775
         Width           =   885
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acumulado ?:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   13
         Left            =   3480
         TabIndex        =   35
         Top             =   3360
         Width           =   1650
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   12
         Left            =   1560
         TabIndex        =   34
         Top             =   2760
         Width           =   840
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "( A - Administrador"
         Height          =   255
         Left            =   1200
         TabIndex        =   33
         Top             =   4200
         Width           =   1380
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "L - Logo"
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "U - Usuário"
         Height          =   255
         Left            =   3480
         TabIndex        =   31
         Top             =   4200
         Width           =   855
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   1560
         TabIndex        =   30
         Top             =   855
         Width           =   1500
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   1560
         TabIndex        =   29
         Top             =   1605
         Width           =   1725
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmação:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   1560
         TabIndex        =   28
         Top             =   2070
         Width           =   1725
      End
      Begin VB.Shape Shape1 
         Height          =   4920
         Index           =   0
         Left            =   360
         Top             =   495
         Width           =   5895
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12.23
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   1560
         TabIndex        =   27
         Top             =   3855
         Width           =   600
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   -73920
         TabIndex        =   26
         Top             =   1125
         Width           =   825
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nome:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   10
         Left            =   -73785
         TabIndex        =   25
         Top             =   1770
         Width           =   690
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contato:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   -73965
         TabIndex        =   24
         Top             =   2370
         Width           =   870
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   -74280
         TabIndex        =   23
         Top             =   2985
         Width           =   1185
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Telefone 2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   -69840
         TabIndex        =   22
         Top             =   3000
         Width           =   1185
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   -73755
         TabIndex        =   21
         Top             =   3630
         Width           =   660
      End
      Begin VB.Label lblFieldLabel 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Faz Assistência :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.83
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   -72225
         TabIndex        =   20
         Top             =   4215
         Width           =   1770
      End
      Begin VB.Shape Shape1 
         Height          =   4575
         Index           =   1
         Left            =   -74640
         Top             =   600
         Width           =   8535
      End
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   377
      Left            =   0
      Top             =   6630
      Width           =   7085
      _ExtentX        =   12508
      _ExtentY        =   671
      ConnectMode     =   0
      CursorLocation  =   2
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   2
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
      Caption         =   "Registro(s): 0 / 0"
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
   Begin ComctlLib.ImageList IMG 
      Left            =   6120
      Top             =   600
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":285D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":2B77
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":2E91
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":31AB
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":34C5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Usuario.frx":37DF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
   Begin VB.Menu mnuSep01 
      Caption         =   "      |"
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
   Begin VB.Menu mnuExcluir 
      Caption         =   "&Excluir"
   End
   Begin VB.Menu mnuSep02 
      Caption         =   "|"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frm_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_New As Boolean
Dim w_Alt_Senha As Boolean


Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    adoReg.Caption = "Registro(s) : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount
    txt_PassConf = ""
    
sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox msgErro(err), vbCritical

    Resume sair
End Sub


Private Sub btFiltrar_Click()
    
    If Usu_Nome.Value = True Then
        w_filtro = "Usl_nome LIKE '%" & txt_filtro & "%'"
    ElseIf usu_tipo.Value = True Then
        w_filtro = "usl_tipo LIKE '%" & txt_filtro & "%'"
    End If
    

    If txt_filtro = "" Or IsEmpty(w_filtro) Then
        Set adoReg.Recordset = de.rstab_usuario.Clone
    Else
        adoReg.Recordset.Filter = IIf(txt_filtro = "" Or IsEmpty(w_filtro), 0, w_filtro)
    End If
End Sub




Private Sub Grid_DblClick()
    
    SSTab1.Tab = 0

End Sub







Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    If SSTab1.Tab = 0 Then
        txt_Pass.SetFocus
        txt_Usu.SetFocus
    End If
End Sub



Private Sub Text1_GotFocus()
txt_Usu.SetFocus
End Sub


Private Sub txt_ac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_filtro_Change()
    btFiltrar_Click
End Sub


Private Sub Form_Activate()
    w_Alt_Senha = False
End Sub

Private Sub Form_Load()
On Error GoTo err1

    MDI.TBar.Visible = False

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.92) - (Me.Height / 2) - 100
    
    Set adoReg.Recordset = ExecuteSQL("SELECT tab_usuario.* FROM tab_usuario WHERE (usl_nome LIKE '" & IIf(w_Usu_Tipo = "L", w_Usu_Nome, "%") & "') ORDER BY usl_nome").Clone
    If Not w_Usu_Tipo = "A" Then
       TBar.Buttons("novo").Visible = False
       TBar.Buttons("excluir").Visible = False
       mnuNovo.Visible = False
       mnuExcluir.Visible = False
       txt_Tipo.Enabled = False
       txt_Usu.Enabled = False
       txt_grupo.Visible = False
       txt_ordem.Visible = False
       txt_ac.Visible = False
       txt_Rpt.Visible = False
       txt_versao.Enabled = False
       txt_versao_login.Enabled = False
       
       lblFieldLabel(12).Visible = False
       lblFieldLabel(13).Visible = False
       lblFieldLabel(14).Visible = False
       lblFieldLabel(15).Visible = False
       
       
       SSTab1.TabVisible(1) = False
       adoReg.Recordset.Filter = "usl_nome = '" & w_Usu_Nome & "'"
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

Private Sub mnuFechar_Click()
    Unload Me
End Sub
Private Sub mnuNovo_Click()
On Error GoTo err1
    
        Set txt_Usu.DataSource = Nothing
        Set txt_Pass.DataSource = Nothing
        Set txt_Tipo.DataSource = Nothing
        Set txt_grupo.DataSource = Nothing
        Set txt_ordem.DataSource = Nothing
        Set txt_ac.DataSource = Nothing
        Set txt_Rpt.DataSource = Nothing
        Set txt_versao.DataSource = Nothing
        Set txt_versao_login.DataSource = Nothing

        txt_Usu = ""
        txt_Pass = ""
        txt_Tipo = "U"
        txt_grupo = ""
        txt_ordem = ""
        txt_ac = ""
        txt_Rpt = ""
        txt_versao = "0.0.0"
        txt_versao_login = "0.0.0"
    
    SSTab1.Tab = 0
    txt_Usu.SetFocus
    w_New = True

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub
Private Sub mnuCancelar_Click()
On Error GoTo err1

          If Not w_Usu_Tipo = "L" Then
          
                w_Pos = adoReg.Recordset.AbsolutePosition
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_usuario ORDER BY usl_nome").Clone
                adoReg.Recordset.Move w_Pos - 1
                
          Else
                
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_usuario Where usl_nome = '" & w_Usu_Nome & "' ORDER BY usl_nome").Clone
                
          End If
          w_New = False
          w_Alt_Senha = False
          
        Set txt_Usu.DataSource = adoReg
        Set txt_Pass.DataSource = adoReg
        Set txt_Tipo.DataSource = adoReg
        Set txt_grupo.DataSource = adoReg
        Set txt_ordem.DataSource = adoReg
        Set txt_ac.DataSource = adoReg
        Set txt_Rpt.DataSource = adoReg
        Set txt_versao.DataSource = adoReg
        Set txt_versao_login.DataSource = adoReg
          
sair:
    Exit Sub
err1:
    On Error Resume Next
     MsgBox msgErro(err), vbCritical
    If Not (err.Number = 0 Or err.Number = 91 Or err.Number = -2147217871 Or err.Number = -2147467259) Then If adoReg.Recordset.RecordCount > 0 Then adoReg.Recordset.MoveFirst
    Resume sair
End Sub
Sub Salvar()
On Error GoTo err1

    If (txt_Pass = txt_PassConf And txt_PassConf <> "") Or w_Alt_Senha = False Then
            
        If Not txt_Usu.DataSource Is Nothing Then
             
             strSQL = "UPDATE tab_usuario SET usl_nome = '" & txt_Usu & "', usl_pwd = '" & txt_Pass & "', " & _
              "usl_tipo = '" & txt_Tipo & "', " & _
              "usl_grupo = '" & txt_grupo & "', usl_ordem = '" & txt_ordem & "', usl_ac = '" & txt_ac & "', " & _
              "usl_rpt = '" & txt_Rpt & "', " & _
              "usl_versao = '" & txt_versao & "', usl_versao_login = '" & txt_versao_login & "' " & _
              "WHERE usl_cod = '" & adoReg.Recordset.Fields("usl_cod") & "'"
        
        Else
             
            If adoReg.Recordset.RecordCount > 0 Then w_cod = ExecuteSQL("Select MAX(usl_cod) as MCodigo from tab_usuario").Fields(0) + 1
             
             strSQL_Fields = "usl_nome, usl_pwd, usl_tipo, usl_grupo, usl_ordem, usl_ac, usl_rpt, usl_versao, usl_versao_login, usl_cod"
             
             strSql_Values = "'" & txt_Usu & "', '" & txt_Pass & "', '" & txt_Tipo & "', '" & txt_grupo & "', " & _
              "'" & txt_ordem & "', '" & txt_ac & "', '" & txt_Rpt & "', '" & txt_versao & "',  '" & txt_versao_login & "', '" & w_cod & "'"
            
             
             strSQL = "INSERT INTO tab_usuario (" & strSQL_Fields & ") " & _
             "VALUES (" & strSql_Values & ")"
        
        End If
    
        ExecuteSQL strSQL, , , False
        
        Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_usuario ORDER BY usl_nome").Clone
    
        Set txt_Usu.DataSource = adoReg
        Set txt_Pass.DataSource = adoReg
        Set txt_Tipo.DataSource = adoReg
        Set txt_grupo.DataSource = adoReg
        Set txt_ordem.DataSource = adoReg
        Set txt_ac.DataSource = adoReg
        Set txt_Rpt.DataSource = adoReg
        Set txt_versao.DataSource = adoReg
        Set txt_versao_login.DataSource = adoReg
        
        w_Alt_Senha = False
        txt_PassConf = ""
        MsgBox "Registro Salvo com sucesso!", vbInformation
    Else
        MsgBox "Senhas não conferem!" & Chr(13) & "Redigite às!", vbExclamation
        txt_Pass = ""
        txt_PassConf = ""
        txt_Pass.SetFocus
    End If

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub
Private Sub mnuExcluir_Click()
On Error GoTo err1

        If vbYes = MsgBox("Deseja excluir?", vbQuestion + vbYesNo + vbDefaultButton1) Then
            ExecuteSQL "DELETE FROM tab_usuario WHERE usl_cod = " & adoReg.Recordset.Fields("usl_cod")
            adoReg.Recordset.MovePrevious
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
    Case "excluir": mnuExcluir_Click
    End Select
End Sub



Private Sub txt_grupo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub txt_ordem_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"

End Sub

Private Sub txt_Pass_Change()
    w_Alt_Senha = True
End Sub

Private Sub txt_Pass_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub
Private Sub txt_PassConf_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And txt_Tipo.Enabled = False Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            Salvar
            txtfor_Desc.SetFocus
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub



Private Sub txt_Rpt_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_Tipo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            Salvar
            txt_Usu.SetFocus
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub txt_Usu_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub



Private Sub Usu_Nome_Click()
    txt_filtro.SetFocus
    adoReg.Recordset.Sort = "USL_NOME"
End Sub

Private Sub usu_tipo_Click()
    txt_filtro.SetFocus
    adoReg.Recordset.Sort = "USL_TIPO"
End Sub
