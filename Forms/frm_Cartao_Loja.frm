VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSAdoDc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDatLst.Ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Cartao_Loja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Taxas Cartão / Loja"
   ClientHeight    =   7155
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   11685
   Icon            =   "frm_Cartao_Loja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7155
   ScaleWidth      =   11685
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   1535
      ButtonWidth     =   1720
      ButtonHeight    =   1429
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
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
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Excluir"
            Key             =   "excluir"
            Object.ToolTipText     =   "Excluir Registro"
            Object.Tag             =   ""
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Listagem"
            Key             =   "rpt"
            Object.ToolTipText     =   "Listagem das Taxas"
            Object.Tag             =   ""
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Copy"
            Key             =   "copy"
            Object.ToolTipText     =   "Copia taxas p/ outras lojas"
            Object.Tag             =   ""
            ImageIndex      =   12
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
         Left            =   7470
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "Taxas Cartão / Loja"
         Top             =   120
         Width           =   4095
      End
   End
   Begin MSAdodcLib.Adodc adoReg 
      Height          =   375
      Left            =   480
      Top             =   6240
      Visible         =   0   'False
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   11430
      _ExtentX        =   20161
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ficha Individual"
      TabPicture(0)   =   "frm_Cartao_Loja.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label14"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label13"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label10"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label8"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label4"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label3"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Shape1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label11"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label12"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label15"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label17"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txt_parc_alta_qt"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txt_parc_alta_tx"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txt_label_ndoc"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "txt_bco_add"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txt_tx_Po_add"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "txt_vr_desc_Pz_add"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "txt_dias_P_add"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "txt_tx_Pz_add"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "txt_vr_desc_V_add"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txt_dias_V_add"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txt_Cartao_add"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txt_Logo_add"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txt_txV_add"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txt_Des_Parc"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "adoLogo"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "adoCartao"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).ControlCount=   34
      TabCaption(1)   =   "Grade"
      TabPicture(1)   =   "frm_Cartao_Loja.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txt_Status_Filtro"
      Tab(1).Control(1)=   "Painel"
      Tab(1).Control(2)=   "grid"
      Tab(1).ControlCount=   3
      Begin MSAdodcLib.Adodc adoCartao 
         Height          =   375
         Left            =   1560
         Top             =   2040
         Visible         =   0   'False
         Width           =   1200
         _ExtentX        =   2117
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
      Begin VB.TextBox txt_Status_Filtro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   645
         Left            =   -74880
         TabIndex        =   38
         Text            =   "Nenhum registro foi encontrado!"
         Top             =   3240
         Visible         =   0   'False
         Width           =   10935
      End
      Begin MSAdodcLib.Adodc adoLogo 
         Height          =   330
         Left            =   3240
         Top             =   1320
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
      Begin VB.ComboBox txt_Des_Parc 
         DataField       =   "ctl_des_parc"
         DataSource      =   "adoReg"
         Height          =   315
         ItemData        =   "frm_Cartao_Loja.frx":27DA
         Left            =   2445
         List            =   "frm_Cartao_Loja.frx":27E4
         TabIndex        =   10
         Text            =   "N"
         Top             =   4485
         Width           =   495
      End
      Begin VB.Frame Painel 
         Height          =   1035
         Left            =   -74880
         TabIndex        =   0
         Top             =   465
         Width           =   6975
         Begin VB.CheckBox op 
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
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   35
            Top             =   600
            Width           =   975
         End
         Begin VB.CheckBox op 
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
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   240
            Width           =   975
         End
         Begin Skin_Button.ctr_Button btFiltrar 
            Default         =   -1  'True
            Height          =   855
            Left            =   6045
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   120
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   1508
            BTYPE           =   9
            TX              =   "&Filtrar"
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
            COLTYPE         =   2
            FOCUSR          =   -1  'True
            BCOL            =   -2147483633
            BCOLO           =   -2147483624
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "frm_Cartao_Loja.frx":27EE
            PICN            =   "frm_Cartao_Loja.frx":280A
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   2
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSDataListLib.DataCombo txt_logo_pesq 
            Bindings        =   "frm_Cartao_Loja.frx":2B24
            Height          =   315
            Left            =   2280
            TabIndex        =   30
            Top             =   210
            Width           =   2085
            _ExtentX        =   3678
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
         Begin MSDataListLib.DataCombo txt_cartao_pesq 
            Bindings        =   "frm_Cartao_Loja.frx":2B3A
            Height          =   315
            Left            =   2280
            TabIndex        =   32
            Top             =   600
            Width           =   3450
            _ExtentX        =   6085
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
         Begin VB.Label lbCartao_pesq 
            BackStyle       =   0  'Transparent
            Caption         =   "Cartão:"
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
            Height          =   255
            Left            =   1575
            TabIndex        =   33
            Top             =   675
            Width           =   675
         End
         Begin VB.Label lbLogo_pesq 
            BackStyle       =   0  'Transparent
            Caption         =   "Logo:"
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
            Height          =   255
            Left            =   1680
            TabIndex        =   31
            Top             =   270
            Width           =   615
         End
      End
      Begin MSDataGridLib.DataGrid grid 
         Bindings        =   "frm_Cartao_Loja.frx":2B52
         Height          =   3975
         Left            =   -74880
         TabIndex        =   5
         Top             =   1680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   7011
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
         ColumnCount     =   10
         BeginProperty Column00 
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
         BeginProperty Column01 
            DataField       =   "Cartão"
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
            DataField       =   "%-Vista"
            Caption         =   "%-Vista"
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
         BeginProperty Column03 
            DataField       =   "Dias-V"
            Caption         =   "Dias-V"
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
            DataField       =   "Vr Desc - V"
            Caption         =   "Vr Desc - V"
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
            DataField       =   "%-Prazo"
            Caption         =   "%-Prazo"
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
         BeginProperty Column06 
            DataField       =   "Dias-Pz"
            Caption         =   "Dias-Pz"
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
            DataField       =   "Vr Desc - Pz"
            Caption         =   "Vr Desc - Pz"
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
         BeginProperty Column08 
            DataField       =   "%-Pz Adic"
            Caption         =   "%-Pz Adic"
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
         BeginProperty Column09 
            DataField       =   "Bco Dep"
            Caption         =   "Bco Dep"
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
            AllowRowSizing  =   0   'False
            AllowSizing     =   0   'False
            BeginProperty Column00 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1080
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column06 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column07 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column08 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column09 
               ColumnAllowSizing=   0   'False
            EndProperty
         EndProperty
      End
      Begin rdActiveText.ActiveText txt_txV_add 
         DataField       =   "ctl_txv"
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
         Left            =   8280
         TabIndex        =   11
         Top             =   1680
         Width           =   735
         _ExtentX        =   1296
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
         MaxLength       =   7
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin MSDataListLib.DataCombo txt_Logo_add 
         Bindings        =   "frm_Cartao_Loja.frx":2B67
         DataField       =   "ctl_loja"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   720
         TabIndex        =   6
         Top             =   1320
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "usl_nome"
         BoundColumn     =   "usl_cod"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin MSDataListLib.DataCombo txt_Cartao_add 
         Bindings        =   "frm_Cartao_Loja.frx":2B7D
         DataField       =   "ctl_tipoc"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Top             =   2040
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   556
         _Version        =   393216
         MatchEntry      =   -1  'True
         Style           =   2
         ListField       =   "tpc_desc"
         BoundColumn     =   "tpc_cod"
         Text            =   ""
         Object.DataMember      =   ""
      End
      Begin rdActiveText.ActiveText txt_dias_V_add 
         DataField       =   "ctl_dias_v"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   8970
         TabIndex        =   12
         Top             =   1680
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
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_vr_desc_V_add 
         DataField       =   "ctl_vr_des_v"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   9630
         TabIndex        =   13
         Top             =   1680
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   556
         Alignment       =   1
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
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_tx_Pz_add 
         DataField       =   "ctl_txp"
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
         Left            =   7155
         TabIndex        =   14
         Top             =   3000
         Width           =   660
         _ExtentX        =   1164
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
         MaxLength       =   7
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_dias_P_add 
         DataField       =   "ctl_dias_p"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   7800
         TabIndex        =   15
         Top             =   3000
         Width           =   765
         _ExtentX        =   1349
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
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_vr_desc_Pz_add 
         DataField       =   "ctl_vr_des_p"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   1
         EndProperty
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   8550
         TabIndex        =   16
         Top             =   3000
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         Alignment       =   1
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
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_tx_Po_add 
         DataField       =   "ctl_vr_po"
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
         Left            =   9750
         TabIndex        =   17
         Top             =   3000
         Width           =   945
         _ExtentX        =   1667
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
         MaxLength       =   7
         RawText         =   0
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin MSDataListLib.DataCombo txt_bco_add 
         Bindings        =   "frm_Cartao_Loja.frx":2B95
         DataField       =   "ctl_banco"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   720
         TabIndex        =   8
         Top             =   2760
         Width           =   3360
         _ExtentX        =   5927
         _ExtentY        =   556
         _Version        =   393216
         ListField       =   "bco_desc"
         BoundColumn     =   "bco_cod"
         Text            =   ""
         Object.DataMember      =   "tab_banco"
      End
      Begin rdActiveText.ActiveText txt_label_ndoc 
         DataField       =   "ctl_label_ndoc"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   720
         TabIndex        =   9
         Top             =   3600
         Width           =   2910
         _ExtentX        =   5133
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
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_parc_alta_tx 
         DataField       =   "ctl_parc_alta_tx"
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
         Left            =   9240
         TabIndex        =   42
         Top             =   4080
         Width           =   660
         _ExtentX        =   1164
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
         MaxLength       =   7
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin rdActiveText.ActiveText txt_parc_alta_qt 
         DataField       =   "ctl_parc_alta_qt"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   7200
         TabIndex        =   43
         Top             =   4080
         Width           =   765
         _ExtentX        =   1349
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
         FontSize        =   8.25
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   9240
         TabIndex        =   41
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "A partir da parcela"
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
         Left            =   7200
         TabIndex        =   40
         Top             =   3840
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Retenção - Qtde maior de parcelas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   255
         Left            =   6960
         TabIndex        =   39
         Top             =   3480
         Width           =   3855
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "O Desconto é aplicado em cada Parcela?"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   720
         TabIndex        =   29
         Top             =   4275
         Width           =   2325
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Nº Doc. deste Cartão"
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
         Left            =   720
         TabIndex        =   28
         Top             =   3360
         Width           =   2865
      End
      Begin VB.Shape Shape1 
         Height          =   4095
         Left            =   360
         Top             =   840
         Width           =   10695
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Height          =   255
         Left            =   720
         TabIndex        =   26
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   8250
         TabIndex        =   25
         Top             =   1440
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
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
         Left            =   8970
         TabIndex        =   24
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vr Desc"
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
         Left            =   9675
         TabIndex        =   23
         Top             =   1440
         Width           =   990
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
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
         Height          =   255
         Left            =   7050
         TabIndex        =   22
         Top             =   2760
         Width           =   750
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dias"
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
         Left            =   7800
         TabIndex        =   21
         Top             =   2760
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vr Desc"
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
         Left            =   8640
         TabIndex        =   20
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label9 
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
         Left            =   9765
         TabIndex        =   19
         Top             =   2760
         Width           =   915
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Bco Dep"
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
         Left            =   720
         TabIndex        =   18
         Top             =   2520
         Width           =   2850
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Retenção - À Vista"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   1050
         Left            =   8040
         TabIndex        =   36
         Top             =   1080
         Width           =   2850
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Retenção - À Prazo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   2175
         Left            =   6960
         TabIndex        =   37
         Top             =   2400
         Width           =   3885
      End
   End
   Begin MSAdodcLib.Adodc adoSQL 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   6825
      Width           =   11685
      _ExtentX        =   20611
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
      Caption         =   "adoSQL"
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
         NumListImages   =   13
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":2BA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":2EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":309A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":33B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":36CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":39E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":3D02
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":3EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":41F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":43D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":46EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":4A04
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Cartao_Loja.frx":4D1E
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
   Begin VB.Menu mnuExcluir 
      Caption         =   "&Excluir"
   End
   Begin VB.Menu mnuSep03 
      Caption         =   "|"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuRpt 
      Caption         =   "&Listagem"
   End
End
Attribute VB_Name = "frm_Cartao_Loja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_Usu As String
Dim w_cod As Long

Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    'adoReg.Caption = "Registro(s) : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount


sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub



Private Sub adoSQL_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1
    If Not adoSQL.Recordset.EOF And Not adoSQL.Recordset.BOF Then
        adoReg.Recordset.MoveFirst
        adoReg.Recordset.Find "ctl_cod = " & adoSQL.Recordset.Fields("ctl_cod") & ""
    End If
    adoSQL.Caption = "Registro(s) : " & adoSQL.Recordset.AbsolutePosition & " / " & adoSQL.Recordset.RecordCount
    
sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox err.Number & " : " & err.Description, vbCritical
    mnuCancelar_Click
    Resume sair
End Sub


Private Sub btFiltrar_Click()
On Error GoTo err1
    
    If op(0).Value = 1 Then w_filtro = "ctl_loja = " & txt_logo_pesq.BoundText & ""
    If op(1).Value = 1 Then w_filtro = w_filtro & IIf(w_filtro = "", "", " and ") & "ctl_tipoc = " & txt_cartao_pesq.BoundText & ""

    If op(0).Value = 0 And op(1).Value = 0 Then

        adoReg.Recordset.Filter = 0
        adoSQL.Recordset.Filter = 0
       ' adoReg.Recordset.CancelUpdate
        adoReg.Recordset.MoveFirst
        Set adoSQL.Recordset = de.rsSql_Cartao_Loja.Clone
    
    Else
        adoReg.Recordset.Filter = 0
        adoSQL.Recordset.Filter = 0
       ' adoReg.Recordset.CancelUpdate
        adoReg.Recordset.Filter = w_filtro
        adoSQL.Recordset.Filter = w_filtro
        
    End If
    
    txt_Status_Filtro.Visible = adoReg.Recordset.EOF
    
sair:
    Exit Sub
err1:
    'MsgBox ERR.Number & " : " & ERR.Description, vbCritical
        On Error Resume Next
        adoReg.Recordset.Filter = 0
        adoSQL.Recordset.Filter = 0
        txt_Status_Filtro.Visible = adoReg.Recordset.EOF
    
    Resume sair
End Sub

Private Sub Form_Activate()
    If mnuRpt.Visible = False Then Habilita_Menus True
End Sub

Private Sub Form_Load()
On Error GoTo err1
    
    MDI.TBar.Visible = False
    
    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 100
    
    w_Usu = IIf(w_Usu_Tipo = "L", w_Usu_Nome, "%")

    Set adoLogo.Recordset = w_ado_Logo.Clone
    Set adoCartao.Recordset = ExecuteSQL("Select * from tab_tipo_cartao Order By tpc_desc").Clone
     
    str_SQL = "SELECT tab_cartao_loja.* FROM tab_cartao_loja, tab_usuario WHERE tab_cartao_loja.ctl_loja = tab_usuario.usl_cod AND (tab_usuario.usl_nome LIKE '" & w_Usu & "') ORDER BY tab_usuario.usl_nome"
    Set adoReg.Recordset = ExecuteSQL(str_SQL).Clone
    
    At_Grid

sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub

Sub At_Grid()
On Error GoTo err1

    If w_Usu_Tipo = "L" Then
        w_Usu = w_Usu_Nome
    Else
        w_Usu = "%"
    End If
    
    str_SQL = "SELECT tab_cartao_loja.ctl_cod, tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartão, tab_cartao_loja.ctl_txv AS `%-Vista`, tab_cartao_loja.ctl_dias_v AS `Dias-V`, tab_cartao_loja.ctl_vr_des_v AS `Vr Desc - V`, tab_cartao_loja.ctl_txp AS `%-Prazo`, tab_cartao_loja.ctl_dias_p AS `Dias-Pz`, tab_cartao_loja.ctl_vr_des_p AS `Vr Desc - Pz`, tab_cartao_loja.ctl_vr_po AS `%-Pz Adic`, tab_banco.bco_desc AS `Bco Dep`, tab_cartao_loja.ctl_loja, tab_cartao_loja.ctl_tipoc, tab_cartao_loja.ctl_label_ndoc, tab_cartao_loja.ctl_des_parc FROM tab_tipo_cartao, tab_usuario, { oj tab_cartao_loja LEFT OUTER JOIN tab_banco ON tab_cartao_loja.ctl_banco = tab_banco.bco_cod } WHERE (tab_cartao_loja.ctl_loja = tab_usuario.usl_cod) AND (tab_cartao_loja.ctl_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_usuario.usl_nome LIKE '" & w_Usu & "') ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc"
    Set adoSQL.Recordset = ExecuteSQL(str_SQL).Clone
    
sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub


Sub Salvar()
On Error GoTo err1

    If Not txt_bco_add.DataSource Is Nothing Then
         
         strSQL = "UPDATE tab_cartao_loja SET ctl_loja = '" & txt_Logo_add.BoundText & "', " & _
         "ctl_tipoc = '" & txt_Cartao_add.BoundText & "', ctl_txv = '" & Replace(Format(txt_txV_add, "0.00##"), ",", ".") & "', " & _
         "ctl_txp = '" & Replace(Format(txt_tx_Pz_add, "0.00##"), ",", ".") & "', ctl_dias_v = '" & txt_dias_V_add & "', " & _
         "ctl_dias_p = '" & txt_dias_P_add & "', ctl_vr_des_v = '" & Replace(CDbl(txt_vr_desc_V_add), ",", ".") & "', " & _
         "ctl_vr_des_p = '" & Replace(CDbl(txt_vr_desc_Pz_add), ",", ".") & "', ctl_vr_po = '" & Replace(Format(txt_tx_Po_add, "0.00##"), ",", ".") & "', " & _
         "ctl_banco = '" & txt_bco_add.BoundText & "', ctl_label_ndoc = '" & txt_label_ndoc & "', " & _
         "ctl_parc_alta_qt = '" & txt_parc_alta_qt & "', ctl_parc_alta_tx = '" & Replace(Format(txt_parc_alta_tx, "0.00##"), ",", ".") & "', " & _
         "ctl_des_parc = '" & Replace(txt_Des_Parc, ",", ".") & "', ctl_cod = '" & adoReg.Recordset.Fields("ctl_cod") & "'" & _
         " WHERE ctl_cod = " & adoReg.Recordset.Fields("ctl_cod")

    Else
         
         strSQL = "INSERT INTO tab_cartao_loja (ctl_loja , ctl_tipoc, ctl_txv, ctl_txp, ctl_dias_v, ctl_dias_p," & _
         "ctl_vr_des_v, ctl_vr_des_p, ctl_vr_po, ctl_banco, ctl_label_ndoc, ctl_des_parc, ctl_cod, ctl_parc_alta_qt, ctl_parc_alta_tx) " & _
         "VALUES ('" & txt_Logo_add.BoundText & "', '" & txt_Cartao_add.BoundText & "', " & _
         "'" & Replace(Format(txt_txV_add, "0.00##"), ",", ".") & "', '" & Replace(Format(txt_tx_Pz_add, "0.00##"), ",", ".") & "', " & _
         "'" & txt_dias_V_add & "', '" & txt_dias_P_add & "', '" & Replace(CDbl(txt_vr_desc_V_add), ",", ".") & "', " & _
         "'" & Replace(CDbl(txt_vr_desc_Pz_add), ",", ".") & "', '" & Replace(Format(txt_tx_Po_add, "0.00##"), ",", ".") & "', " & _
         "'" & txt_bco_add.BoundText & "', '" & txt_label_ndoc & "', '" & Replace(txt_Des_Parc, ",", ".") & "', " & _
         "'" & w_cod & "', '" & txt_parc_alta_qt & "', '" & Replace(Format(txt_parc_alta_tx, "0.00##"), ",", ".") & "')"
         
    End If
    
    
    ExecuteSQL strSQL, w_RegAf, , False
    
    MsgBox "Registro Salvo com sucesso!", vbInformation
    
    str_SQL = "SELECT tab_cartao_loja.* FROM tab_cartao_loja, tab_usuario WHERE tab_cartao_loja.ctl_loja = tab_usuario.usl_cod AND (tab_usuario.usl_nome LIKE '" & w_Usu & "') ORDER BY tab_usuario.usl_nome"
    Set adoReg.Recordset = ExecuteSQL(str_SQL).Clone
    
    
    Set txt_bco_add.DataSource = adoReg
    Set txt_Cartao_add.DataSource = adoReg
    Set txt_Des_Parc.DataSource = adoReg
    Set txt_dias_P_add.DataSource = adoReg
    Set txt_dias_V_add.DataSource = adoReg
    Set txt_label_ndoc.DataSource = adoReg
    Set txt_Logo_add.DataSource = adoReg
    Set txt_tx_Po_add.DataSource = adoReg
    Set txt_tx_Pz_add.DataSource = adoReg
    Set txt_txV_add.DataSource = adoReg
    Set txt_vr_desc_Pz_add.DataSource = adoReg
    Set txt_vr_desc_V_add.DataSource = adoReg
    Set txt_parc_alta_tx.DataSource = adoReg
    Set txt_parc_alta_qt.DataSource = adoReg
    
    
    If adoSQL.Recordset.Filter <> 0 Then w_filtro = adoSQL.Recordset.Filter
    At_Grid
    If w_filtro <> 0 Then adoSQL.Recordset.Filter = w_filtro
    
    
sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub



Private Sub Grid_DblClick()
On Error GoTo err1
    If Not adoReg.Recordset.EOF Then
        w_cod = adoSQL.Recordset.Fields("ctl_cod")
        SSTab1.Tab = 0
        adoReg.Recordset.MoveFirst
        adoReg.Recordset.Find "ctl_cod = " & w_cod & ""
        txt_Logo_add.SetFocus
        Sendkeys "{esc}"
    End If
sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub



Private Sub mnuCancelar_Click()
On Error GoTo err1

    Set txt_bco_add.DataSource = adoReg
    Set txt_Cartao_add.DataSource = adoReg
    Set txt_Des_Parc.DataSource = adoReg
    Set txt_dias_P_add.DataSource = adoReg
    Set txt_dias_V_add.DataSource = adoReg
    Set txt_label_ndoc.DataSource = adoReg
    Set txt_Logo_add.DataSource = adoReg
    Set txt_tx_Po_add.DataSource = adoReg
    Set txt_tx_Pz_add.DataSource = adoReg
    Set txt_txV_add.DataSource = adoReg
    Set txt_vr_desc_Pz_add.DataSource = adoReg
    Set txt_vr_desc_V_add.DataSource = adoReg
    
    
    
    str_SQL = "SELECT tab_cartao_loja.* FROM tab_cartao_loja, tab_usuario WHERE tab_cartao_loja.ctl_loja = tab_usuario.usl_cod AND (tab_usuario.usl_nome LIKE '" & w_Usu & "') ORDER BY tab_usuario.usl_nome"
    Set adoReg.Recordset = ExecuteSQL(str_SQL).Clone
       
    
    
          If adoReg.Recordset.Fields(9) <> "" Then
                w_Pos = adoSQL.Recordset.AbsolutePosition - IIf(adoSQL.Recordset.Fields(1) = "", 1, 0)
                At_Grid
                adoSQL.Recordset.MoveFirst
                adoSQL.Recordset.Move w_Pos - 1
          End If
          
          
sair:
    Exit Sub
err1:
    If err.Number <> 458 Then MsgBox msgErro(err), vbCritical

    On Error Resume Next
    'If adoReg.Recordset.RecordCount > 0 Then adoReg.Recordset.MoveFirst
    Resume sair
End Sub

Private Sub mnuExcluir_Click()
On Error GoTo err1
        
    If vbYes = MsgBox("Deseja excluir?", vbQuestion + vbYesNo + vbDefaultButton1) Then
        w_cod = 0
        w_cod = adoSQL.Recordset.Fields("ctl_cod")
        ExecuteSQL "Delete From tab_cartao_loja Where ctl_cod = " & w_cod
        
        mnuCancelar_Click
        
    End If
        
sair:
    Exit Sub
err1:
   
    If err.Number = -2147217871 Then
        MsgBox "Antes de Excluir o tipo de cartão, você precisa excluir as formas de pagamentos relacionadas a mesma!", vbInformation
    Else
         MsgBox msgErro(err), vbCritical
    End If
    Resume sair
End Sub

Private Sub mnuFechar_Click()
If mnuRpt.Visible = True Then
    Unload Me
Else
    Unload Rel_Cartao_Loja
    Habilita_Menus True
End If
End Sub

Private Sub mnuNovo_Click()
On Error GoTo err1
    
    SSTab1.Tab = 0
    txt_Logo_add.Enabled = True
    
    w_cod = ExecuteSQL("Select Max(ctl_Cod) as UltCod from tab_cartao_loja").Fields(0) + 1
        
    Set txt_bco_add.DataSource = Nothing
    Set txt_Cartao_add.DataSource = Nothing
    Set txt_Des_Parc.DataSource = Nothing
    Set txt_dias_P_add.DataSource = Nothing
    Set txt_dias_V_add.DataSource = Nothing
    Set txt_label_ndoc.DataSource = Nothing
    Set txt_Logo_add.DataSource = Nothing
    Set txt_tx_Po_add.DataSource = Nothing
    Set txt_tx_Pz_add.DataSource = Nothing
    Set txt_txV_add.DataSource = Nothing
    Set txt_vr_desc_Pz_add.DataSource = Nothing
    Set txt_vr_desc_V_add.DataSource = Nothing
    
    
    txt_bco_add = ""
    txt_Cartao_add = ""
    txt_Des_Parc = ""
    txt_dias_P_add = ""
    txt_dias_V_add = ""
    txt_label_ndoc = ""
    txt_Logo_add = ""
    txt_tx_Po_add = ""
    txt_tx_Pz_add = ""
    txt_txV_add = ""
    txt_vr_desc_Pz_add = ""
    txt_vr_desc_V_add = ""
                
    If w_Usu_Tipo = "L" Then txt_Logo_add.BoundText = w_Usu_Cod
    txt_Des_Parc = "N"
    
    txt_Logo_add.SetFocus
    
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuRpt_Click()
On Error GoTo err1
    
    If w_Usu_Tipo = "L" Then
        w_Usu = w_Usu_Nome
    Else
        w_Usu = "%"
    End If
    
    If op(0).Value = 1 Then w_filtro = " AND tab_usuario.usl_COD = " & txt_logo_pesq.BoundText & ""
    If op(1).Value = 1 Then w_filtro = w_filtro & " AND ctl_tipoc = " & txt_cartao_pesq.BoundText & ""
    
    w_SQL = "SHAPE {SELECT tab_cartao_loja.ctl_cod, tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartão, tab_cartao_loja.ctl_txv AS `%-Vista`, tab_cartao_loja.ctl_dias_v AS `Dias-V`, tab_cartao_loja.ctl_vr_des_v AS `Vr Desc - V`, tab_cartao_loja.ctl_txp AS `%-Prazo`, tab_cartao_loja.ctl_dias_p AS `Dias-Pz`, tab_cartao_loja.ctl_vr_des_p AS `Vr Desc - Pz`, tab_cartao_loja.ctl_vr_po AS `%-Pz Adic`, tab_banco.bco_desc AS `Bco Dep`, tab_cartao_loja.ctl_loja, tab_cartao_loja.ctl_tipoc, tab_cartao_loja.ctl_label_ndoc, tab_cartao_loja.ctl_des_parc FROM tab_tipo_cartao, tab_usuario, { oj tab_cartao_loja LEFT OUTER JOIN tab_banco ON tab_cartao_loja.ctl_banco = tab_banco.bco_cod } WHERE (tab_cartao_loja.ctl_loja = tab_usuario.usl_cod) AND (tab_cartao_loja.ctl_tipoc = tab_tipo_cartao.tpc_cod) " & w_filtro & " ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc}  AS Rpt_SQL_Cartao_Loja COMPUTE Rpt_SQL_Cartao_Loja BY 'Logo'"
    
    Dim w_Rec As Recordset
    Set w_Rec = ExecuteSQL(w_SQL, , "MSDataShape").Clone
   
    If w_Rec.RecordCount > 0 Then
        Set Rel_Cartao_Loja.DataSource = w_Rec.Clone

        Habilita_Menus False

        Rel_Cartao_Loja.WindowState = vbMaximized
        Rel_Cartao_Loja.Show
    End If
sair:
    Exit Sub
err1:
    
    Habilita_Menus True
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub


Sub Habilita_Menus(op As Boolean)
    mnuCancelar.Visible = op
    mnuSalvar.Visible = op
    mnuExcluir.Visible = op
    mnuNovo.Visible = op
    mnuRpt.Visible = op
    mnuSep01.Visible = op
    mnuSep03.Visible = op
End Sub

Private Sub mnuSalvar_Click()
    Salvar
End Sub



Private Sub op_Click(Index As Integer)
    Select Case Index
    Case 0:
            txt_logo_pesq.Enabled = op(0).Value
            lbLogo_pesq.Enabled = op(0).Value
            If lbLogo_pesq.Enabled = True Then txt_logo_pesq.SetFocus
    Case 1:
            txt_cartao_pesq.Enabled = op(1).Value
            lbCartao_pesq.Enabled = op(1).Value
            If txt_cartao_pesq.Enabled = True Then txt_cartao_pesq.SetFocus
    End Select
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        txt_Logo_add.SetFocus
        btFiltrar.Default = False
    Else
        btFiltrar.Default = True
        op(0).SetFocus
    End If
End Sub


Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    Case "novo": mnuNovo_Click
    Case "salvar": Salvar
    Case "cancelar": mnuCancelar_Click
    Case "excluir": mnuExcluir_Click
    Case "rpt": mnuRpt_Click
    Case "copy":
                  frm_Copy_CartLoja.Show
                  frm_Copy_CartLoja.txt_Logo = txt_Logo_add
                  frm_Copy_CartLoja.txt_Cartao = txt_Cartao_add
                  
    End Select
End Sub




Private Sub Text1_GotFocus()
 If SSTab1.Tab = 0 Then
    txtfor_Desc.SetFocus
 Else
    txt_filtro.SetFocus
 End If
End Sub

Private Sub tlj_cod_Click()
    adoReg.Recordset.Sort = "tpc_cod"
    txt_filtro.SetFocus
End Sub

Private Sub tlj_desc_Click()
    adoReg.Recordset.Sort = "tpc_desc"
    txt_filtro.SetFocus
End Sub



Private Sub txt_bco_add_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_Cartao_add_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub


Private Sub txt_Des_Parc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_filtro_Change()
    btFiltrar_Click
End Sub





Private Sub txt_Logo_add_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Sendkeys "{tab}"
End Sub

Private Sub txt_tx_Po_add_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            txt_Logo_add.SetFocus
            Salvar
        End If

    End If
End Sub

