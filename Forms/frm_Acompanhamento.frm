VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Acompanhamento 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acompanhamento"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   11190
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   11190
      _ExtentX        =   19738
      _ExtentY        =   1111
      ButtonWidth     =   1244
      ButtonHeight    =   1005
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
      Begin VBXPolitec.ocxProgressBarTexto pgBar 
         Height          =   420
         Left            =   6120
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   741
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColorFundo  =   -2147483643
         MaxProgress     =   100
      End
      Begin VB.TextBox lblTitle 
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
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "Acompanhamento"
         Top             =   120
         Width           =   3975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Relatório "
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
      Left            =   5160
      TabIndex        =   16
      Top             =   720
      Width           =   6015
      Begin VB.CheckBox ck 
         Caption         =   "Venda"
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
         Left            =   3000
         TabIndex        =   9
         Top             =   480
         Width           =   1080
      End
      Begin VB.CheckBox ck 
         Caption         =   "Recebimento"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   840
         Width           =   1575
      End
      Begin VB.CheckBox ck 
         Caption         =   "Saldo"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   11
         Top             =   1200
         Width           =   1320
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
         Left            =   360
         TabIndex        =   7
         Top             =   1200
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
         FontSize        =   8,25
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
         Left            =   1725
         TabIndex        =   8
         Top             =   1185
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
         FontSize        =   8,25
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
         Left            =   360
         TabIndex        =   5
         Top             =   600
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
         FontSize        =   8,25
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
         Left            =   1725
         TabIndex        =   6
         Top             =   600
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
         FontSize        =   8,25
      End
      Begin Skin_Button.ctr_Button bt_Gerar 
         Height          =   1080
         Left            =   4680
         TabIndex        =   12
         Top             =   360
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   1905
         BTYPE           =   2
         TX              =   "&Gerar"
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
         MICON           =   "frm_Acompanhamento.frx":0000
         PICN            =   "frm_Acompanhamento.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lb_Dt 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data Vcto"
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
         Left            =   375
         TabIndex        =   20
         Top             =   945
         Width           =   2385
      End
      Begin VB.Label lb_Dt2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "à"
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
         Left            =   1305
         TabIndex        =   19
         Top             =   1245
         Width           =   480
      End
      Begin VB.Label lb_Dt2P 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "à"
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
         Left            =   1305
         TabIndex        =   18
         Top             =   645
         Width           =   480
      End
      Begin VB.Label lb_DtP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Escopo de Data Protocolo"
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
         Left            =   420
         TabIndex        =   17
         Top             =   360
         Width           =   2310
      End
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
      TabIndex        =   0
      Top             =   720
      Width           =   5085
      Begin VB.TextBox txt 
         Height          =   285
         Left            =   2760
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   2295
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
         TabIndex        =   1
         Top             =   360
         Width           =   1095
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
         TabIndex        =   3
         Top             =   720
         Width           =   1155
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   4200
         Top             =   360
      End
      Begin MSAdodcLib.Adodc adoLogo 
         Height          =   330
         Left            =   1785
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
         Left            =   3000
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
      Begin MSDataListLib.DataCombo txt_Logo_P 
         Bindings        =   "frm_Acompanhamento.frx":0336
         Height          =   315
         Left            =   1440
         TabIndex        =   2
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
         Bindings        =   "frm_Acompanhamento.frx":034C
         Height          =   315
         Left            =   1440
         TabIndex        =   4
         Top             =   1110
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
      Begin VB.Label lb_Logo 
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
         Left            =   1440
         TabIndex        =   15
         Top             =   270
         Width           =   615
      End
      Begin VB.Label lb_Cartao 
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
         Left            =   1440
         TabIndex        =   14
         Top             =   870
         Width           =   1215
      End
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   4920
      Top             =   360
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
            Picture         =   "frm_Acompanhamento.frx":0364
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":067E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":0858
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":0B72
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":0E8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":169A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Acompanhamento.frx":19B4
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frm_Acompanhamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub bt_Gerar_Click()
Dim w_SQL_Venda, w_SQL_Saldo, w_SQL_Recebido
Dim w_Rec_Venda As New Recordset
Dim w_Rec_Saldo As New Recordset
Dim w_Rec_Recebido As New Recordset

On Error Resume Next
   
If (ck(6).Value <> 0 And Txt_DtIP <> "" And Txt_DtFP <> "") _
    Or (ck(2).Value <> 0 And Txt_DtI <> "" And Txt_DtF <> "") _
    Or (ck(3).Value <> 0 And Txt_DtF <> "") Then
    
    pgBar.Visible = True
    pgBar.Text = "................................ Gerando Relatório ..............................."
    pgBar.Value = 1
    
    If (ck(6).Value <> 0 And Txt_DtIP <> "" And Txt_DtFP <> "") Then
        'Progresso
        pgBar.Value = pgBar.Value + 7
        'Venda
        w_SQL_Venda = ""
        
        w_SQL_Venda = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_tipo_cartao.tpc_desc as cartao, sum(tab_lanc_parc.lcp_vr_bto) as vr_bto, sum(tab_lanc_parc.lcp_vr_liq) as vr_liq, Month(tab_lanc.lnc_dt_vnd) as mes, Monthname(tab_lanc.lnc_dt_vnd) as mesNome " & _
                " From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc" & _
                " Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj" & _
                       " And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num" & _
                       " And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod" & _
                       " And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod )"
                       
        If ck(0).Value <> 0 And txt_Logo_P <> "" Then
            w_SQL_Venda = w_SQL_Venda & " And tab_lanc.lnc_loj = " & txt_Logo_P.BoundText & ""
        End If
        
        If ck(1).Value <> 0 And txt_Cartao_P <> "" Then
            w_SQL_Venda = w_SQL_Venda & " And tab_lanc.lnc_tipoc = " & txt_Cartao_P.BoundText & ""
        End If
        
        If ck(6).Value <> 0 And Txt_DtIP <> "" And Txt_DtFP <> "" Then
            w_SQL_Venda = w_SQL_Venda & " And tab_lanc.lnc_dt_vnd >= '" & Format(Txt_DtIP, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(Txt_DtFP, "yyyy-mm-dd") & "'"
        End If
         
        w_SQL_Venda = w_SQL_Venda & " GROUP BY logo,cartao,mes ORDER BY tab_usuario.usl_cod,tab_tipo_cartao.tpc_desc,Month(tab_lanc.lnc_dt_vnd)} AS Sql_VendaR COMPUTE Sql_VendaR BY 'Logo'"
        
        'Progresso
        pgBar.Value = pgBar.Value + 7

        Set w_Rec_Venda = ExecuteSQL(w_SQL_Venda, , "MSDataShape").Clone
           
        'Progresso
        pgBar.Value = pgBar.Value + 7
        Set Rel_Venda.DataSource = w_Rec_Venda.Clone
        Rel_Venda.WindowState = vbMaximized
        
        Rel_Venda.Sections("Section4").Controls("lbDT").Caption = "De " & Txt_DtIP & "  à  " & Txt_DtFP
        
        Rel_Venda.Show
        'Progresso
        pgBar.Value = pgBar.Value + 7
    
    End If
    
        
    If ck(2).Value <> 0 And Txt_DtI <> "" And Txt_DtF <> "" Then
        'Progresso
        pgBar.Value = pgBar.Value + 7
         'Recebido
         w_SQL_Recebido = ""
         w_SQL_Recebido = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_tipo_cartao.tpc_desc as cartao, sum(tab_lanc_parc.lcp_vr_bto) as vr_bto" & _
                 " From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc" & _
                 " Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj" & _
                        " And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num" & _
                        " And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod" & _
                        " And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod )"
             
             If ck(0).Value <> 0 And txt_Logo_P <> "" Then
                 w_SQL_Recebido = w_SQL_Recebido & " And tab_lanc.lnc_loj = " & txt_Logo_P.BoundText & ""
             End If
             
             If ck(1).Value <> 0 And txt_Cartao_P <> "" Then
                 w_SQL_Recebido = w_SQL_Recebido & " And tab_lanc.lnc_tipoc = " & txt_Cartao_P.BoundText & ""
             End If
             
             If ck(2).Value <> 0 And Txt_DtI <> "" And Txt_DtF <> "" Then
                 w_SQL_Recebido = w_SQL_Recebido & " And tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "'" & " And tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
             End If
         w_SQL_Recebido = w_SQL_Recebido & " And tab_lanc_parc.lcp_baixa <> '0000-00-00'"
         w_SQL_Recebido = w_SQL_Recebido & " GROUP BY logo,cartao ORDER BY tab_usuario.usl_cod,tab_tipo_cartao.tpc_desc} AS Sql_Recebido COMPUTE Sql_Recebido BY 'Logo'"
        
        'Progresso
        pgBar.Value = pgBar.Value + 7
        
         Set w_Rec_Recebido = ExecuteSQL(w_SQL_Recebido, , "MSDataShape").Clone
        
        'Progresso
        pgBar.Value = pgBar.Value + 7
        
         Set Rel_Recebido.DataSource = w_Rec_Recebido.Clone
         Rel_Recebido.WindowState = vbMaximized
         
         Rel_Recebido.Sections("Section4").Controls("lbDT").Caption = "De " & Txt_DtI & "  à  " & Txt_DtF
         
         Rel_Recebido.Show
        
        'Progresso
        pgBar.Value = pgBar.Value + 7
    End If
    
    If ck(3).Value <> 0 And Txt_DtF <> "" Then
        'Progresso
        pgBar.Value = pgBar.Value + 7
         'Saldo
         w_SQL_Saldo = ""
         w_SQL_Saldo = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_tipo_cartao.tpc_desc as cartao, sum(tab_lanc_parc.lcp_vr_bto) as vr_bto, sum(tab_lanc_parc.lcp_vr_liq) as vr_liq" & _
                 " From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc" & _
                 " Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj" & _
                        " And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num" & _
                        " And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod" & _
                        " And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod )"
             
             If ck(0).Value <> 0 And txt_Logo_P <> "" Then
                 w_SQL_Saldo = w_SQL_Saldo & " And tab_lanc.lnc_loj = " & txt_Logo_P.BoundText & ""
             End If
             
             If ck(1).Value <> 0 And txt_Cartao_P <> "" Then
                 w_SQL_Saldo = w_SQL_Saldo & " And tab_lanc.lnc_tipoc = " & txt_Cartao_P.BoundText & ""
             End If
             
             If ck(3).Value <> 0 And Txt_DtIP <> "" And Txt_DtFP <> "" Then
                 w_SQL_Saldo = w_SQL_Saldo & " And tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtF, "yyyy-mm-dd") & "'"
             End If
        
        'Progresso
        pgBar.Value = pgBar.Value + 7
        
         w_SQL_Saldo = w_SQL_Saldo & " GROUP BY logo,cartao ORDER BY tab_usuario.usl_cod,tab_tipo_cartao.tpc_desc} AS Sql_Saldo COMPUTE Sql_Saldo BY 'Logo'"
         Set w_Rec_Saldo = ExecuteSQL(w_SQL_Saldo, , "MSDataShape").Clone
         
        'Progresso
        pgBar.Value = pgBar.Value + 7
         
         Set Rel_Saldo.DataSource = w_Rec_Saldo.Clone
         Rel_Saldo.WindowState = vbMaximized
         
         
         Rel_Saldo.Sections("Section4").Controls("lbDT").Caption = "A partir de " & Txt_DtF
         
         Rel_Saldo.Show
         
        'Progresso
        pgBar.Value = pgBar.Value + 7
    End If
    
        pgBar.Value = 0
        pgBar.Visible = False
    
Else
    MsgBox "É necessário escolher pelo menos um tipo de relatório [Venda | Recebido | Saldo]!", vbCritical
End If

txt.Text = w_SQL_Venda
End Sub

Private Sub ck_Click(Index As Integer)
        txt_Logo_P.Enabled = ck(0).Value
        txt_Cartao_P.Enabled = ck(1).Value
        
        'Venda
        Txt_DtIP.Enabled = ck(6).Value
        Txt_DtFP.Enabled = ck(6).Value
        
        'Recebido
        Txt_DtI.Enabled = ck(2).Value
        
        'Recebido e Saldo
        If (ck(2).Value Or ck(3).Value) Then
            Txt_DtF.Enabled = True
        Else
            Txt_DtF.Enabled = False
        End If
        
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
    
    Txt_DtIP = w_Data_Server
    Txt_DtFP = w_Data_Server
   
sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair

End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    End Select
End Sub

Private Sub mnuFechar_Click()
        Unload Me
End Sub
