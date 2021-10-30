VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Forma_Pg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forma de Pagamento"
   ClientHeight    =   6994
   ClientLeft      =   39
   ClientTop       =   611
   ClientWidth     =   8255
   Icon            =   "frm_Forma_Pg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6994
   ScaleWidth      =   8255
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   819
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8255
      _ExtentX        =   14569
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
         Left            =   5280
         Locked          =   -1  'True
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "Forma de Pg."
         Top             =   120
         Width           =   3015
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5415
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   7725
      _ExtentX        =   13635
      _ExtentY        =   9561
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Ficha Individual"
      TabPicture(0)   =   "frm_Forma_Pg.frx":27A2
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblFieldLabel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lblFieldLabel(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Shape1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblFieldLabel(2)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblFieldLabel(3)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtfor_Desc"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtfor_Cod"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txt_Tipo"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txt_Qt"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      TabCaption(1)   =   "Grade"
      TabPicture(1)   =   "frm_Forma_Pg.frx":27BE
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Grid"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Painel"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin rdActiveText.ActiveText txt_Qt 
         DataField       =   "fpg_qt_parc"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   2640
         TabIndex        =   2
         Top             =   2445
         Width           =   495
         _ExtentX        =   863
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
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   7.472
      End
      Begin VB.ComboBox txt_Tipo 
         DataField       =   "fpg_Tipo"
         DataSource      =   "adoReg"
         Height          =   273
         ItemData        =   "frm_Forma_Pg.frx":27DA
         Left            =   5490
         List            =   "frm_Forma_Pg.frx":27E7
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2445
         Width           =   525
      End
      Begin VB.Frame Painel 
         Caption         =   "Tipo de Filtro"
         Height          =   1035
         Left            =   -74280
         TabIndex        =   4
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
            TabIndex        =   7
            Top             =   585
            Width           =   3855
         End
         Begin VB.OptionButton tlj_desc 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   2040
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   255
            Width           =   1215
         End
         Begin VB.OptionButton tlj_cod 
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            TabStop         =   0   'False
            Top             =   255
            Value           =   -1  'True
            Width           =   1095
         End
         Begin Skin_Button.ctr_Button btFiltrar 
            Height          =   855
            Left            =   4845
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   120
            Width           =   735
            _ExtentX        =   1294
            _ExtentY        =   1510
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
            MICON           =   "frm_Forma_Pg.frx":27F4
            PICN            =   "frm_Forma_Pg.frx":2810
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
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
            TabIndex        =   14
            Top             =   630
            Width           =   690
         End
      End
      Begin MSDataGridLib.DataGrid Grid 
         Bindings        =   "frm_Forma_Pg.frx":2B2A
         Height          =   4080
         Left            =   -74280
         TabIndex        =   9
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
         ColumnCount     =   4
         BeginProperty Column00 
            DataField       =   "fpg_cod"
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
            DataField       =   "fpg_desc"
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
         BeginProperty Column02 
            DataField       =   "fpg_qt_parc"
            Caption         =   "Qt Parc."
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
            DataField       =   "fpg_Tipo"
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
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column03 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtfor_Cod 
         DataField       =   "fpg_cod"
         DataSource      =   "adoReg"
         Enabled         =   0   'False
         Height          =   315
         Left            =   2640
         TabIndex        =   0
         Top             =   1200
         Width           =   1020
      End
      Begin rdActiveText.ActiveText txtfor_Desc 
         DataField       =   "fpg_desc"
         DataSource      =   "adoReg"
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   1800
         Width           =   3375
         _ExtentX        =   5943
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
         TextCase        =   3
         RawText         =   0
         eAuto           =   1
         FontName        =   "MS Sans Serif"
         FontSize        =   7.472
      End
      Begin VB.Label Label1 
         Caption         =   "V (À Vista)"
         Height          =   255
         Left            =   4770
         TabIndex        =   20
         Top             =   3300
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "P (Prazo)"
         Height          =   255
         Left            =   4770
         TabIndex        =   19
         Top             =   3060
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "D ( Pré-Datado)"
         Height          =   255
         Left            =   4770
         TabIndex        =   18
         Top             =   2820
         Width           =   1215
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo :"
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
         Index           =   3
         Left            =   4725
         TabIndex        =   17
         Top             =   2490
         Width           =   615
      End
      Begin VB.Label lblFieldLabel 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde de Parc.:"
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
         Index           =   2
         Left            =   1005
         TabIndex        =   16
         Top             =   2490
         Width           =   1500
      End
      Begin VB.Shape Shape1 
         Height          =   3015
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
         Left            =   1365
         TabIndex        =   12
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
         Left            =   1680
         TabIndex        =   11
         Top             =   1245
         Width           =   825
      End
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   377
      Left            =   0
      Top             =   6617
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
            Picture         =   "frm_Forma_Pg.frx":2B3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":2E59
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":3033
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":334D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":3667
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":3981
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":3C9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Forma_Pg.frx":3E75
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
Attribute VB_Name = "frm_Forma_Pg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    adoReg.Caption = "Registro(s) : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount

sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox msgErro(err), vbCritical
    Resume sair
End Sub





Private Sub btFiltrar_Click()
On Error GoTo err1
    
    If tlj_cod.Value = True And IsNumeric(txt_filtro) Then
        w_filtro = "fpg_cod LIKE " & txt_filtro & ""
    ElseIf tlj_desc.Value = True Then
        w_filtro = "fpg_desc LIKE '" & txt_filtro & "%'"
    End If
    

    If txt_filtro = "" Or IsEmpty(w_filtro) Then
        Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone
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
    
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone

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
         
         strSQL = "UPDATE tab_forma_pg SET fpg_desc = '" & txtfor_Desc & "', fpg_qt_parc = '" & txt_Qt & "', fpg_tipo = '" & txt_Tipo & "' WHERE fpg_cod = '" & txtfor_Cod & "'"
    
    Else
         
         strSQL_Fields = "fpg_desc, fpg_qt_parc, fpg_tipo"
         strSql_Values = "'" & txtfor_Desc & "', '" & txt_Qt & "', '" & txt_Tipo & "'"
         strSQL = "INSERT INTO tab_forma_pg (" & strSQL_Fields & ") " & _
         "VALUES (" & strSql_Values & ")"
    
    End If

    ExecuteSQL strSQL, , , False
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone
        
    
    MsgBox "Registro Salvo com sucesso!", vbInformation


    Set txt_Qt.DataSource = adoReg
    Set txt_Tipo.DataSource = adoReg
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
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone
                adoReg.Recordset.MoveFirst
                adoReg.Recordset.Move w_Pos - 1
          Else
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone
          End If
          
    Set txt_Qt.DataSource = adoReg
    Set txt_Tipo.DataSource = adoReg
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
        ExecuteSQL "DELETE FROM tab_forma_pg WHERE fpg_cod = " & txtfor_Cod

                w_Pos = adoReg.Recordset.AbsolutePosition - IIf(adoReg.Recordset.Fields(1) = "", 1, 0)
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_forma_pg").Clone
                adoReg.Recordset.MoveFirst
                adoReg.Recordset.Move w_Pos - 1
    End If
        
sair:
    Exit Sub
err1:
     MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub mnuFechar_Click()
    Unload Me
End Sub

Private Sub mnuNovo_Click()
On Error GoTo err1
    SSTab1.Tab = 0

    Set txt_Qt.DataSource = Nothing
    Set txt_Tipo.DataSource = Nothing
    Set txtfor_Cod.DataSource = Nothing
    Set txtfor_Desc.DataSource = Nothing

    txt_Qt = ""
    txt_Tipo = "V"
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
    adoReg.Recordset.Sort = "fpg_cod"
    txt_filtro.SetFocus
End Sub

Private Sub tlj_desc_Click()
    adoReg.Recordset.Sort = "fpg_desc"
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

Private Sub txt_SN_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            Salvar
            txtfor_Desc.SetFocus
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub

