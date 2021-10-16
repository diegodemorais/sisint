VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Rpt_Total_Receber 
   Caption         =   "Teste"
   ClientHeight    =   8216
   ClientLeft      =   1833
   ClientTop       =   2548
   ClientWidth     =   5356
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8216
   ScaleWidth      =   5356
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   624
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   5356
      _ExtentX        =   9441
      _ExtentY        =   1102
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "Relatório Teste"
         Top             =   0
         Width           =   2895
      End
   End
   Begin VB.ListBox List_loja 
      Appearance      =   0  'Flat
      Height          =   3731
      ItemData        =   "frm_Rpt_Total_Receber.frx":0000
      Left            =   1080
      List            =   "frm_Rpt_Total_Receber.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   3255
      Width           =   1335
   End
   Begin VB.TextBox txt 
      Height          =   285
      Left            =   1920
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   960
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.OptionButton opt2 
      Caption         =   "&Mês a mês"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1440
      Width           =   1455
   End
   Begin VB.OptionButton opt1 
      Caption         =   "&Dia a Dia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   1440
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VBXPolitec.ocxProgressBarTexto pgBar 
      Height          =   300
      Left            =   240
      TabIndex        =   0
      Top             =   7680
      Visible         =   0   'False
      Width           =   4815
      _ExtentX        =   8483
      _ExtentY        =   527
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   8.8302
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   "................................ Gerando Relatório ..............................."
      Text            =   "................................ Gerando Relatório ..............................."
      BackColorFundo  =   -2147483643
      MaxProgress     =   100
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
      Left            =   1080
      TabIndex        =   1
      Top             =   2280
      Width           =   1275
      _ExtentX        =   2252
      _ExtentY        =   551
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.8302
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
      FontSize        =   8.83
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
      Left            =   2880
      TabIndex        =   2
      Top             =   2280
      Width           =   1275
      _ExtentX        =   2252
      _ExtentY        =   551
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.8302
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
      FontSize        =   8.83
   End
   Begin Skin_Button.ctr_Button bt_Pesq 
      Height          =   1080
      Left            =   3720
      TabIndex        =   3
      Top             =   6240
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   1917
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
      MICON           =   "frm_Rpt_Total_Receber.frx":0004
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc adoLogo 
      Height          =   375
      Left            =   1080
      Top             =   7050
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2348
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
   Begin Skin_Button.ctr_Button bt_STodos 
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      TabStop         =   0   'False
      ToolTipText     =   "Seleciona todos"
      Top             =   3360
      Width           =   1965
      _ExtentX        =   3475
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
      MICON           =   "frm_Rpt_Total_Receber.frx":0020
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
      Left            =   2520
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Retira Selecão de todos"
      Top             =   3960
      Width           =   1965
      _ExtentX        =   3475
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
      MICON           =   "frm_Rpt_Total_Receber.frx":003C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Descontar (B)s:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   15
      Top             =   2880
      Width           =   2145
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   11
      Top             =   1080
      Width           =   1065
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   4320
      Top             =   1320
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   9
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":0058
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":0372
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":054C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":0866
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":0B80
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":0E9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":138E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Total_Receber.frx":16A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lb_Dt2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "à"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2385
      TabIndex        =   5
      Top             =   2325
      Width           =   480
   End
   Begin VB.Label lb_Dt 
      BackStyle       =   0  'Transparent
      Caption         =   "Período"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.83
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   600
      TabIndex        =   4
      Top             =   1920
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      Height          =   6660
      Left            =   240
      Top             =   840
      Width           =   4815
   End
End
Attribute VB_Name = "frm_Rpt_Total_Receber"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_RPT As Boolean


Private Sub bt_Pesq_Click()
On Error Resume Next
       
Dim w_SQL As String
Dim w_Rec As New Recordset
Dim wNenhuma As Boolean
Dim wDescontar As String

    pgBar.Value = 0
    pgBar.Visible = True
    
    If opt1.Value Then
        pgBar.Value = 0
        wDescontar = "Descontar (B)s:"
    
        'w_SQL = "SELECT tab_lanc_parc.lcp_dt_vcto as Data, Sum(tab_lanc_parc.lcp_vr_bto) as Bruto, Sum(tab_lanc_parc.lcp_vr_liq) as Liquido From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) and tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' And tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' GROUP BY tab_lanc_parc.lcp_dt_vcto ORDER BY tab_lanc_parc.lcp_dt_vcto;"
        w_SQL = "SELECT parc.lcp_dt_vcto as Data, " & _
                "sum((SELECT Sum(tab_lanc_parc.lcp_vr_bto) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
                      "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num  And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
                      "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc not like '%DEPÓSITO%')) as CartaoBto, " & _
                "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
                      "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num  And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
                      "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc not like '%DEPÓSITO%')) as CartaoLiq, " & _
                "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
                      "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
                      "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc like '%DEPÓSITO%')) as Deposito, " & _
                "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc " & _
                      "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc_parc.lcp_num = parc.lcp_num)) as Total " & _
                "From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc parc " & _
                "Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And parc.lcp_num_lanc =  tab_lanc.lnc_num " & _
                      "And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod " & _
                      "And parc.lcp_baixa = '0000-00-00') " & _
                      "and parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' " & _
                      "And parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' " & _
                      "And tab_usuario.usl_nome NOT IN ("
                      
                wNenhuma = True
                For i = 0 To List_loja.ListCount - 1
                    pgBar.Value = pgBar.Value + 1
                    'Salva configuração
                    ExecuteSQL ("UPDATE tab_usuario SET usl_saldo = " & List_loja.Selected(i) & " WHERE usl_nome = '" & List_loja.List(i) & "'")
                    If List_loja.Selected(i) = True Then
                        If wNenhuma = False Then
                            w_SQL = w_SQL & ","
                            wDescontar = wDescontar & ","
                        End If
                        w_SQL = w_SQL & "'" & List_loja.List(i) & "'"
                        wDescontar = wDescontar & " " & List_loja.List(i)
                        wNenhuma = False
                    End If
                Next i
                If wNenhuma Then w_SQL = w_SQL & "''"
                w_SQL = w_SQL & ") GROUP BY parc.lcp_dt_vcto ORDER BY parc.lcp_dt_vcto;"
        txt = w_SQL

        Set w_Rec = ExecuteSQL(w_SQL, w_RegAf, "MSDataShape").Clone
        
        pgBar.Value = 50
       
        If w_RegAf > 0 Then
            Set Rel_Total_Receber.DataSource = w_Rec.Clone
            'Rel_Total_Receber.WindowState = vbMaximized
            pgBar.Value = 60
    
            Rel_Total_Receber.Sections("Cabecalho").Controls("lbDT").Caption = "De " & Txt_DtI & "  à  " & Txt_DtF
            Rel_Total_Receber.Sections("Cabecalho").Controls("lbLojas").Caption = wDescontar
    
            Rel_Total_Receber.Show
            pgBar.Value = 80
            w_RPT = True

        End If
       
        'Relatório das lojas descontadas
        w_SQL = Replace(w_SQL, "And tab_usuario.usl_nome NOT IN (", "And tab_usuario.usl_nome IN (")
        wDescontar = Replace(wDescontar, "Descontar (B)s:", "Somente (B)s:")
        Set w_Rec = ExecuteSQL(w_SQL, w_RegAf, "MSDataShape").Clone
      
        If w_RegAf > 0 Then
            Set Rel_Total_Receber2.DataSource = w_Rec.Clone
            'Rel_Total_Receber2.WindowState = vbMaximized
    
            Rel_Total_Receber2.Sections("Cabecalho").Controls("lbDT").Caption = "De " & Txt_DtI & "  à  " & Txt_DtF
            Rel_Total_Receber2.Sections("Cabecalho").Controls("lbLojas").Caption = wDescontar
    
            Rel_Total_Receber2.Show
            w_RPT = True

        End If
        
        pgBar.Value = 100
        pgBar.Visible = False
        
    Else
        pgBar.Value = 0
        
        wDescontar = "Descontar (B)s:"
        'w_SQL = "SELECT Month(tab_lanc_parc.lcp_dt_vcto) as Mes, Year(tab_lanc_parc.lcp_dt_vcto) as Ano, Sum(tab_lanc_parc.lcp_vr_bto) as Bruto, Sum(tab_lanc_parc.lcp_vr_liq) as Liquido From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) and tab_lanc_parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' And tab_lanc_parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' GROUP BY Month(tab_lanc_parc.lcp_dt_vcto), Year(tab_lanc_parc.lcp_dt_vcto)  ORDER BY Year(tab_lanc_parc.lcp_dt_vcto),Month(tab_lanc_parc.lcp_dt_vcto);"
        w_SQL = "SELECT Month(parc.lcp_dt_vcto) as Mes, Year(parc.lcp_dt_vcto) as Ano, " & _
        "sum((SELECT Sum(tab_lanc_parc.lcp_vr_bto) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
              "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
              "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc not like '%DEPÓSITO%')) as CartaoBto, " & _
        "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
              "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
              "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc not like '%DEPÓSITO%')) as CartaoLiq, " & _
        "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc, tab_tipo_cartao " & _
              "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod " & _
              "And tab_lanc_parc.lcp_num = parc.lcp_num and tab_tipo_cartao.tpc_desc like '%DEPÓSITO%')) as Deposito, " & _
        "sum((SELECT Sum(tab_lanc_parc.lcp_vr_liq) From  tab_lanc,  tab_lanc_parc " & _
              "Where tab_lanc_parc.lcp_num_lanc = tab_lanc.lnc_num And tab_lanc_parc.lcp_num = parc.lcp_num)) as Total " & _
        "From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc parc " & _
        "Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And parc.lcp_num_lanc =  tab_lanc.lnc_num " & _
              "And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod " & _
              "And parc.lcp_baixa = '0000-00-00') " & _
              "and parc.lcp_dt_vcto >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' " & _
              "And parc.lcp_dt_vcto <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "' " & _
              "And tab_usuario.usl_nome NOT IN ("
                
              wNenhuma = True
              For i = 0 To List_loja.ListCount - 1
                  pgBar.Value = pgBar.Value + 1
                  'Salva configuração
                  ExecuteSQL ("UPDATE tab_usuario SET usl_saldo = " & List_loja.Selected(i) & " WHERE usl_nome = '" & List_loja.List(i) & "'")
                  If List_loja.Selected(i) = True Then
                      If wNenhuma = False Then
                          w_SQL = w_SQL & ","
                          wDescontar = wDescontar & ","
                      End If
                      w_SQL = w_SQL & "'" & List_loja.List(i) & "'"
                      wDescontar = wDescontar & " " & List_loja.List(i)
                      wNenhuma = False
                  End If
              Next i
              If wNenhuma Then w_SQL = w_SQL & "''"
              w_SQL = w_SQL & ") GROUP BY Month(parc.lcp_dt_vcto), Year(parc.lcp_dt_vcto) " & _
                                "ORDER BY Year(parc.lcp_dt_vcto),Month(parc.lcp_dt_vcto); "
    
        
        Set w_Rec = ExecuteSQL(w_SQL, w_RegAf, "MSDataShape").Clone
        
        pgBar.Value = 50
       
        If w_RegAf > 0 Then
            Set Rel_Total_Receber_Mes.DataSource = w_Rec.Clone
            pgBar.Value = 60
            'Rel_Total_Receber_Mes.WindowState = vbMaximized
    
            Rel_Total_Receber_Mes.Sections("Cabecalho").Controls("lbDT").Caption = "De " & Txt_DtI & "  à  " & Txt_DtF
            Rel_Total_Receber_Mes.Sections("Cabecalho").Controls("lbLojas").Caption = wDescontar
    
            Rel_Total_Receber_Mes.Show
            pgBar.Value = 80
            w_RPT = True
       
        End If


        'Relatório das lojas descontadas
        w_SQL = Replace(w_SQL, "And tab_usuario.usl_nome NOT IN (", "And tab_usuario.usl_nome IN (")
        wDescontar = Replace(wDescontar, "Descontar (B)s:", "Somente (B)s:")
        Set w_Rec = ExecuteSQL(w_SQL, w_RegAf, "MSDataShape").Clone
      
        If w_RegAf > 0 Then
            Set Rel_Total_Receber_Mes2.DataSource = w_Rec.Clone
            'Rel_Total_Receber_Mes2.WindowState = vbMaximized
    
            Rel_Total_Receber_Mes2.Sections("Cabecalho").Controls("lbDT").Caption = "De " & Txt_DtI & "  à  " & Txt_DtF
            Rel_Total_Receber_Mes2.Sections("Cabecalho").Controls("lbLojas").Caption = wDescontar
    
            Rel_Total_Receber_Mes2.Show
            w_RPT = True
        End If
  
        pgBar.Value = 100
        pgBar.Visible = False
    End If
               
        txt.Text = w_SQL
    
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
End Sub

Private Sub Form_Load()
        
    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.92) - (Me.Height / 2) - 100
      
    Txt_DtI = w_Data_Server - 1
    Txt_DtF = w_Data_Server - 1
    
    If Weekday(w_Data_Server) = vbMonday Then
        Txt_DtI = w_Data_Server - 3
        Txt_DtF = w_Data_Server - 1
    Else
        Txt_DtI = w_Data_Server - 1
        Txt_DtF = w_Data_Server - 1
    End If


    Set adoLogo.Recordset = ExecuteSQL("SELECT usl_cod, usl_nome, usl_tipo, usl_ac, usl_saldo FROM tab_usuario WHERE (usl_tipo = 'L')", , , False).Clone
    'Set adoLogo.Recordset = w_ado_Logo.Clone
    
    'monta lista das lojas
    For i = 1 To adoLogo.Recordset.RecordCount
        Call List_loja.AddItem(adoLogo.Recordset.Fields("USL_NOME"), List_loja.ListCount)
        If adoLogo.Recordset.Fields("USL_SALDO") Then List_loja.Selected(List_loja.ListCount - 1) = True
        adoLogo.Recordset.MoveNext
    Next i
    
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
        Case "fechar": mnuFechar_Click
    End Select
End Sub

Private Sub mnuFechar_Click()
    Unload Me
End Sub


Private Sub Text1_Change()
    If w_RPT Then
        Unload Rel_Rotal_Receber
    End If
    Unload Me
End Sub

