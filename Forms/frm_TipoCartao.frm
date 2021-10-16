VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_TipoCartao 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tipo de Cartão"
   ClientHeight    =   7215
   ClientLeft      =   39
   ClientTop       =   611
   ClientWidth     =   8255
   Icon            =   "frm_TipoCartao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   8255
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   819
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   8255
      _ExtentX        =   14569
      _ExtentY        =   1438
      ButtonWidth     =   1667
      ButtonHeight    =   1429
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Tipo de Cartão"
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5775
      Left            =   285
      TabIndex        =   8
      Top             =   960
      Width           =   7485
      _ExtentX        =   13203
      _ExtentY        =   10184
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ficha Individual"
      TabPicture(0)   =   "frm_TipoCartao.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFieldLabel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFieldLabel(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtfor_Desc"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtfor_Cod"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "frame1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "Grade"
      TabPicture(1)   =   "frm_TipoCartao.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid"
      Tab(1).Control(1)=   "Painel"
      Tab(1).ControlCount=   2
      Begin VB.Frame frame1 
         BackColor       =   &H80000000&
         Caption         =   " Formas de Pagamento "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.47
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   720
         TabIndex        =   14
         Top             =   2880
         Width           =   5655
         Begin MSAdodcLib.Adodc adoFormaPg 
            Height          =   375
            Left            =   1560
            Top             =   480
            Visible         =   0   'False
            Width           =   1200
            _ExtentX        =   2109
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
         Begin MSDataListLib.DataCombo txt_FormaPg 
            Bindings        =   "frm_TipoCartao.frx":27DA
            Height          =   286
            Left            =   260
            TabIndex        =   20
            Top             =   494
            Visible         =   0   'False
            Width           =   4082
            _ExtentX        =   7524
            _ExtentY        =   503
            _Version        =   393216
            MatchEntry      =   -1  'True
            Style           =   2
            ListField       =   "fpg_desc"
            BoundColumn     =   "fpg_cod"
            Text            =   ""
            Object.DataMember      =   ""
         End
         Begin Skin_Button.ctr_Button bt_Add_F 
            Height          =   525
            Left            =   4875
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   1440
            Width           =   495
            _ExtentX        =   863
            _ExtentY        =   935
         End
         Begin MSAdodcLib.Adodc adoFormas 
            Height          =   330
            Left            =   2760
            Top             =   2280
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
         Begin Skin_Button.ctr_Button bt_Exc_F 
            Height          =   525
            Left            =   4875
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   2040
            Width           =   495
            _ExtentX        =   863
            _ExtentY        =   935
         End
         Begin Skin_Button.ctr_Button bt_Sal_F 
            Height          =   525
            Left            =   4395
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   863
            _ExtentY        =   935
         End
         Begin Skin_Button.ctr_Button bt_Canc_F 
            Height          =   525
            Left            =   4875
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   270
            Visible         =   0   'False
            Width           =   495
            _ExtentX        =   863
            _ExtentY        =   935
         End
         Begin MSDataGridLib.DataGrid Grid_Formas 
            Bindings        =   "frm_TipoCartao.frx":27F3
            Height          =   2385
            Left            =   255
            TabIndex        =   15
            Top             =   270
            Width           =   4560
            _ExtentX        =   8051
            _ExtentY        =   4217
            _Version        =   393216
            AllowUpdate     =   0   'False
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.4717
               Charset         =   0
               Weight          =   400
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
            ColumnCount     =   1
            BeginProperty Column00 
               DataField       =   "Formas"
               Caption         =   "Formas Pagamentos"
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
               ScrollBars      =   2
               BeginProperty Column00 
                  ColumnAllowSizing=   0   'False
                  ColumnWidth     =   3966.803
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Painel 
         Caption         =   "Tipo de Filtro"
         Height          =   1035
         Left            =   -74280
         TabIndex        =   2
         Top             =   465
         Width           =   5655
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
            Left            =   870
            TabIndex        =   5
            Top             =   585
            Width           =   3855
         End
         Begin VB.OptionButton tlj_desc 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   2040
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   255
            Width           =   1215
         End
         Begin VB.OptionButton tlj_cod 
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   255
            Value           =   -1  'True
            Width           =   1095
         End
         Begin Skin_Button.ctr_Button btFiltrar 
            Height          =   855
            Left            =   4845
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   120
            Width           =   735
            _ExtentX        =   1294
            _ExtentY        =   1510
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
            Left            =   150
            TabIndex        =   12
            Top             =   630
            Width           =   690
         End
      End
      Begin MSDataGridLib.DataGrid Grid 
         Bindings        =   "frm_TipoCartao.frx":280B
         Height          =   4080
         Left            =   -74280
         TabIndex        =   7
         Top             =   1575
         Width           =   5655
         _ExtentX        =   9968
         _ExtentY        =   7189
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
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   "tpc_cod"
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
            DataField       =   "tpc_desc"
            Caption         =   "Descrição"
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
         EndProperty
      End
      Begin VB.TextBox txtfor_Cod 
         DataField       =   "tpc_cod"
         DataSource      =   "adoReg"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Top             =   1200
         Width           =   1020
      End
      Begin VB.TextBox txtfor_Desc 
         DataField       =   "tpc_desc"
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
         Left            =   2640
         TabIndex        =   1
         Top             =   1845
         Width           =   3375
      End
      Begin VB.Shape Shape1 
         Height          =   1935
         Left            =   720
         Top             =   720
         Width           =   5655
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição:"
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
         Index           =   1
         Left            =   1125
         TabIndex        =   10
         Top             =   1890
         Width           =   1140
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
         Index           =   0
         Left            =   1440
         TabIndex        =   9
         Top             =   1245
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   377
      Left            =   0
      Top             =   6838
      Width           =   8255
      _ExtentX        =   14569
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
            Picture         =   "frm_TipoCartao.frx":2820
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":2B3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":2D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":302E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":3348
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":3662
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":397C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_TipoCartao.frx":3B56
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
End
Attribute VB_Name = "frm_TipoCartao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    adoReg.Caption = "Registro(s) : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount
    At_Formas

sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox msgErro(err), vbCritical

    Resume sair
End Sub

Sub At_Formas()
On Error GoTo err1
    
    If Not adoReg.Recordset.BOF And Not adoReg.Recordset.EOF And Not IsNull(adoReg.Recordset.Fields("tpc_cod")) Then
        Set adoFormas.Recordset = ExecuteSQL("SELECT tab_tipo_forma.tpc_cod AS CodC, tab_forma_pg.fpg_cod AS CodF, tab_forma_pg.fpg_desc AS Formas FROM tab_tipo_forma, tab_forma_pg WHERE (tab_tipo_forma.fpg_cod = tab_forma_pg.fpg_cod) AND (tab_tipo_forma.tpc_cod = '" & adoReg.Recordset.Fields("tpc_cod") & "') ORDER BY tab_forma_pg.fpg_desc").Clone
    ElseIf Not adoReg.Recordset.BOF And Not adoReg.Recordset.EOF Then
        Set adoFormas.Recordset = ExecuteSQL("SELECT tab_tipo_forma.tpc_cod AS CodC, tab_forma_pg.fpg_cod AS CodF, tab_forma_pg.fpg_desc AS Formas FROM tab_tipo_forma, tab_forma_pg WHERE (tab_tipo_forma.fpg_cod = tab_forma_pg.fpg_cod) AND (tab_tipo_forma.tpc_cod = '0') ORDER BY tab_forma_pg.fpg_desc").Clone
    End If
    
sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox msgErro(err), vbCritical

    Resume sair
End Sub



Private Sub bt_Add_F_Click()
On Error Resume Next

    Grid_Formas.Height = 1695
    Grid_Formas.Top = 960
    
    txt_FormaPg.Visible = True
    bt_Sal_F.Visible = True
    bt_Canc_F.Visible = True
    bt_Exc_F.Enabled = False
    bt_Add_F.Enabled = False
    
End Sub

Private Sub bt_Canc_F_Click()
On Error Resume Next
    
    Grid_Formas.Height = 2385
    Grid_Formas.Top = 270
    txt_FormaPg = ""
    txt_FormaPg.Visible = False
    bt_Sal_F.Visible = False
    bt_Canc_F.Visible = False
    bt_Exc_F.Enabled = True
    bt_Add_F.Enabled = True
End Sub

Private Sub bt_Exc_F_Click()
On Error GoTo err1
   
    If vbYes = MsgBox("Deseja Excluir?", vbQuestion + vbYesNo) Then
        
        Call ExecuteSQL("Delete  From tab_tipo_forma Where fpg_cod = " & adoFormas.Recordset.Fields("CodF") & " and tpc_cod = " & adoReg.Recordset.Fields("tpc_cod") & "", w_RegAf)
        If w_RegAf = 0 Then MsgBox "Não foi possível excluir!", vbInformation
        At_Formas
        
    End If
   
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub bt_Sal_F_Click()
On Error GoTo err1

    If IsNull(adoReg.Recordset.Fields(0)) Then Salvar
    
    If Not txt_FormaPg = "" Then
        
        'Salvar
        ExecuteSQL "INSERT INTO tab_tipo_forma(fpg_cod, tpc_cod) VALUES (" & txt_FormaPg.BoundText & ", " & adoReg.Recordset.Fields("tpc_cod") & ")"
        'Deixar Invisivel os Add
        bt_Canc_F_Click
        At_Formas
    Else
        MsgBox "Escolha a Forma de Pagamento!", vbInformation
        txt_FormaPg.SetFocus
    End If
    
sair:
    Exit Sub
err1:
    If err.Number = -2147217900 Then
        MsgBox "Forma de Pagamento já cadastrada para este tipo de Cartão!", vbCritical
    Else
       MsgBox msgErro(err), vbCritical
    End If
    Resume sair
End Sub

Private Sub btFiltrar_Click()
On Error GoTo err1
    
    If tlj_cod.Value = True And IsNumeric(txt_filtro) Then
        w_filtro = "tpc_cod LIKE " & txt_filtro & ""
    ElseIf tlj_desc.Value = True Then
        w_filtro = "tpc_desc LIKE '" & txt_filtro & "%'"
    End If
    

    If txt_filtro = "" Or IsEmpty(w_filtro) Then
        Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone
    Else
        adoReg.Recordset.Filter = IIf(txt_filtro = "" Or IsEmpty(w_filtro), 0, w_filtro)
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
    
    Set adoFormaPg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg ORDER BY fpg_desc").Clone
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone


sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub

Sub Salvar()
On Error GoTo err1



    If Not txtfor_Desc.DataSource Is Nothing Then
         
         strSQL = "UPDATE tab_tipo_cartao SET tpc_desc = '" & txtfor_Desc & "' WHERE tpc_cod = '" & txtfor_Cod & "'"
    
    Else
         
         strSQL_Fields = "tpc_desc"
         strSql_Values = "'" & txtfor_Desc & "'"
         strSQL = "INSERT INTO tab_tipo_cartao (" & strSQL_Fields & ") " & _
         "VALUES (" & strSql_Values & ")"
    
    End If

    ExecuteSQL strSQL, , , False
    
    MsgBox "Registro Salvo com sucesso!", vbInformation
    
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone


    Set txtfor_Cod.DataSource = adoReg
    Set txtfor_Desc.DataSource = adoReg

sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub



Private Sub Grid_DblClick()
    
    SSTab1.Tab = 0
    txtfor_Desc.SetFocus
    
End Sub

Private Sub mnuCancelar_Click()
On Error GoTo err1

          If Not IsNull(adoReg.Recordset.Fields(1)) Then

                w_Pos = adoReg.Recordset.AbsolutePosition - IIf(adoReg.Recordset.Fields(1) = "", 1, 0)
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone
                adoReg.Recordset.MoveFirst
                adoReg.Recordset.Move w_Pos - 1
                
          Else
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone
          End If
          
          
    Set txtfor_Cod.DataSource = adoReg
    Set txtfor_Desc.DataSource = adoReg
          
sair:
    Exit Sub
err1:
    On Error Resume Next
     MsgBox msgErro(err), vbCritical
    If adoReg.Recordset.RecordCount > 0 Then adoReg.Recordset.MoveFirst
    Resume sair
End Sub

Private Sub mnuExcluir_Click()
On Error GoTo err1
        
    If vbYes = MsgBox("Deseja excluir?", vbQuestion + vbYesNo + vbDefaultButton1) Then
        w_cod = adoReg.Recordset.Fields("tpc_Cod")
        ExecuteSQL "DELETE FROM tab_tipo_cartao Where tpc_cod = " & w_cod
        Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc").Clone

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
    Unload Me
End Sub

Private Sub mnuNovo_Click()
On Error GoTo err1
    SSTab1.Tab = 0

    Set txtfor_Cod.DataSource = Nothing
    Set txtfor_Desc.DataSource = Nothing

    txtfor_Cod = ""
    txtfor_Desc = ""

    txtfor_Desc.Enabled = True
    txtfor_Desc.SetFocus
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuSalvar_Click()
    Salvar
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        txtfor_Desc.SetFocus
    Else
        txt_filtro.SetFocus
    End If
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

Private Sub txt_filtro_Change()
    btFiltrar_Click
End Sub


Private Sub txtfor_Ass_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 13 And vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
        Salvar
        txtfor_Desc.SetFocus
    End If
End Sub

Private Sub txtfor_Contato_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub txtfor_Desc_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            Salvar
            txtfor_Desc.SetFocus
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub

Private Sub txtfor_Email_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub txtfor_Tel_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub

Private Sub txtfor_Tel2_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub
