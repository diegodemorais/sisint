VERSION 5.00
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Menu Principal"
   ClientHeight    =   6840
   ClientLeft      =   165
   ClientTop       =   810
   ClientWidth     =   13470
   Icon            =   "mdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   1535
      ButtonWidth     =   2302
      ButtonHeight    =   1429
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   12
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Sair"
            Key             =   "sair"
            Object.ToolTipText     =   "Sair do Sistema"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Lançar"
            Key             =   "lançar"
            Object.ToolTipText     =   "Lançar Cartões"
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cons. Lanç"
            Key             =   "con_lanc"
            Object.ToolTipText     =   "Consulta Lançamentos"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Resumo"
            Key             =   "resumo"
            Object.ToolTipText     =   "Inserir Resumo"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cód/Bônus"
            Key             =   "cod_bon"
            Object.ToolTipText     =   "Código/Bônus"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Rel. Cod/Bôn"
            Key             =   "rel_codbon"
            Object.ToolTipText     =   "Relatório de Código e Bônus"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Baixa Auto"
            Key             =   "baixa"
            Object.ToolTipText     =   "Baixa Automática"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Caption         =   "Rel. Resumo"
            Key             =   "rpt_resumo"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin Skin_Button.ctr_Button btnSorocredFicha 
         Height          =   765
         Left            =   10200
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1349
         BTYPE           =   2
         TX              =   "Sorocred &Ficha"
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
         MICON           =   "mdi.frx":08CA
         PICN            =   "mdi.frx":08E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button btn_TotRec 
         Height          =   765
         Left            =   11280
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1349
         BTYPE           =   2
         TX              =   "&Teste"
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
         MICON           =   "mdi.frx":1BC8
         PICN            =   "mdi.frx":1BE4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin Skin_Button.ctr_Button ctr_Button1 
         Height          =   765
         Left            =   3240
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2040
         Visible         =   0   'False
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1349
         BTYPE           =   2
         TX              =   "Enviar Caixa"
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
         MICON           =   "mdi.frx":2036
         PICN            =   "mdi.frx":2052
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000080FF&
         BorderStyle     =   0  'None
         Height          =   135
         Left            =   5760
         ScaleHeight     =   135
         ScaleWidth      =   975
         TabIndex        =   3
         Top             =   6600
         Width           =   975
      End
      Begin VBXPolitec.ocxProgressBarTexto pgBar 
         Height          =   315
         Left            =   8400
         TabIndex        =   2
         Top             =   480
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   " Atualizando ....................."
         Text            =   " Atualizando ....................."
         ForeColorFundo  =   33023
         BackColorFundo  =   -2147483643
         BackColorProgress=   33023
      End
   End
   Begin VB.Timer timer_at 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1200
      Top             =   960
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6585
      Width           =   13470
      _ExtentX        =   23760
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   9
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Usuário : "
            TextSave        =   "Usuário : "
            Object.Tag             =   ""
            Object.ToolTipText     =   "Usuário Logado"
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   10292
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "progbar"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Alignment       =   1
            Object.Width           =   1764
            MinWidth        =   1764
            Text            =   "Ver. 0.0.0"
            TextSave        =   "Ver. 0.0.0"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Versão do Sistema"
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   635
            MinWidth        =   617
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "11/04/2020"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "09:17"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":2CA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":38F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":3C40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":4892
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":4BAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":4D86
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":59D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":752A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":8204
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":8D5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "mdi.frx":8FB0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSis 
      Caption         =   "Sistema"
      Begin VB.Menu mnuAtu 
         Caption         =   "Atualização"
         Enabled         =   0   'False
         Begin VB.Menu mnuAtBM 
            Caption         =   "Baixar Arqs. MOV"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnuAtLog 
            Caption         =   "Login"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnuSisPar 
         Caption         =   "Parametrização"
         Begin VB.Menu mnuConf 
            Caption         =   "Config"
         End
         Begin VB.Menu mnubco 
            Caption         =   "Banco"
         End
         Begin VB.Menu mnuCartLoja 
            Caption         =   "Taxa Cartão / Loja"
         End
         Begin VB.Menu mnuFPg 
            Caption         =   "Forma de Pagamento"
         End
         Begin VB.Menu mnuTpC 
            Caption         =   "Tipo de Cartão"
         End
         Begin VB.Menu mnuSisUsu 
            Caption         =   "Usuários"
         End
      End
      Begin VB.Menu mnuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLanc 
         Caption         =   "Lançar"
      End
      Begin VB.Menu mnuBaixa 
         Caption         =   "Baixa Automática"
      End
      Begin VB.Menu mnuConLanc 
         Caption         =   "Consulta Lançamentos"
      End
      Begin VB.Menu mnuTot 
         Caption         =   "Resumo"
      End
      Begin VB.Menu mnusep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodBon 
         Caption         =   "Código / Bônus"
      End
      Begin VB.Menu mnuRpt 
         Caption         =   "Relatório de Código / Bônus"
      End
      Begin VB.Menu mnuSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSisSair 
         Caption         =   "Sair do Sistema"
      End
   End
End
Attribute VB_Name = "MDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub btEnvCX_Click()
    frm_EnvioMov.Show
End Sub

Private Sub btn_TotRec_Click()
    frm_Rpt_Total_Receber.Show
End Sub

Private Sub btnSorocredFicha_Click()
    Dim w_dtIni, w_dtFim As Date
    Dim rsSorocredFicha As Recordset
    Dim sqlSorocredFicha As String
    
        w_dtIni = InputBox("Data Inicial: ", "Data inicial")
        w_dtFim = InputBox("Data Final: ", "Data final")
        

        If IsDate(w_dtIni) And IsDate(w_dtFim) Then
            'sqlSorocredFicha = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_tipo_cartao.tpc_desc as cartao, sum(tab_lanc_parc.lcp_vr_bto) as vr_bto, Month(tab_lanc.lnc_dt_vnd) as mes, Monthname(tab_lanc.lnc_dt_vnd) as mesNome  From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc Where (tab_usuario.usl_cod =  tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num And tab_lanc.lnc_tipoc =  tab_tipo_cartao.tpc_cod And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) And tab_lanc.lnc_tipoc = 7 And tab_lanc.lnc_dt_vnd >= '" & Format(w_dtIni, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(w_dtFim, "yyyy-mm-dd") & "' GROUP BY logo,cartao,mes ORDER BY tab_usuario.usl_cod,tab_tipo_cartao.tpc_desc,Month(tab_lanc.lnc_dt_vnd)} AS Sql_VendaR COMPUTE Sql_VendaR BY 'Logo'"
            'sqlSorocredFicha = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_lanc.lnc_ndoc as doc, tab_lanc.lnc_dt_vnd as dtVenda, sum(tab_lanc_parc.lcp_vr_bto) As vr_bto From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc Where (tab_usuario.usl_cod = tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num And tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod  And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) And tab_lanc.lnc_tipoc = 7 And tab_lanc.lnc_dt_vnd >= '" & Format(w_dtIni, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(w_dtFim, "yyyy-mm-dd") & "' GROUP BY logo, dtVenda ASC, vr_bto DESC ORDER BY tab_usuario.usl_cod ASC, dtVenda ASC, vr_bto DESC} AS Sql_VendaR COMPUTE Sql_VendaR BY 'Logo'"
            sqlSorocredFicha = "SHAPE {SELECT tab_usuario.usl_nome as logo, tab_lanc.lnc_ndoc as doc, tab_lanc.lnc_dt_vnd as dtVenda, sum(tab_lanc_parc.lcp_vr_bto) As vr_bto From tab_usuario, tab_lanc, tab_forma_pg, tab_tipo_cartao, tab_lanc_parc Where (tab_usuario.usl_cod = tab_lanc.lnc_loj And tab_lanc_parc.lcp_num_lanc =  tab_lanc.lnc_num And tab_lanc.lnc_tipoc = tab_tipo_cartao.tpc_cod  And tab_lanc.lnc_formapg = tab_forma_pg.fpg_cod ) And tab_lanc.lnc_tipoc = 6 And tab_lanc.lnc_dt_vnd >= '" & Format(w_dtIni, "yyyy-mm-dd") & "' and tab_lanc.lnc_dt_vnd <= '" & Format(w_dtFim, "yyyy-mm-dd") & "' GROUP BY logo, dtVenda, doc ORDER BY tab_usuario.usl_cod, dtVenda, vr_bto DESC} AS sqlSorocredFicha COMPUTE sqlSorocredFIcha BY 'logo'"
        End If
         
        Set rsSorocredFicha = ExecuteSQL(sqlSorocredFicha, , "MSDataShape").Clone
        'Set rsSorocredFicha = ExecuteSQL(sqlSorocredFicha).Clone
           
        Set Rel_SorocredFicha.DataSource = rsSorocredFicha.Clone
        
        'rsSorocredFicha w_dtIni, w_dtFim
        'Set Rel_SorocredFicha.DataSource = de.rssqlSorocredFicha
        
        Rel_SorocredFicha.WindowState = vbMaximized
        
        Rel_SorocredFicha.Sections("Section4").Controls("lbDT").Caption = "De " & w_dtIni & "  à  " & w_dtFim
        
        Rel_SorocredFicha.Show
End Sub

Private Sub Command1_Click()
Dim wAdo As ADODB.Recordset

    wSQL = "SELECT lnc_num, lnc_nresumo FROM `rpaps_2`.`tab_lanc` where lnc_formapg = 11 and lnc_tipoc = 5"
    
    Set wAdo = ExecuteSQL(wSQL)
    
    Do While Not wAdo.EOF
        wRegAf = 0
            
            'Atualiza o Resumo do Cabeçalho
            wSQL = "UPDATE tab_lanc SET lnc_nresumo = '6" & Mid(wAdo.Fields("lnc_nresumo"), 2) & "' " & _
                   "WHERE lnc_num = " & wAdo.Fields("lnc_num")
            ExecuteSQL wSQL, wRegAf
            
            'Atualiza o Resumo das Parcelas
            wSQL = "UPDATE tab_lanc_parc SET lcp_nresumo = '6" & Mid(wAdo.Fields("lnc_nresumo"), 2) & "' " & _
                   "WHERE lcp_num_lanc = " & wAdo.Fields("lnc_num")
            ExecuteSQL wSQL, wRegAf
            
                
        wAdo.MoveNext
    Loop
    

End Sub





Private Sub MDIForm_Load()
    StatusBar1.Panels(1) = "Usuário: " & w_Usu_Nome
    
    If w_Usu_Tipo = "U" Or w_Usu_Tipo = "A" Then
        btn_TotRec.Visible = True
    End If
    
    If w_Usu_Nome = "KELY" Then
        btnSorocredFicha.Visible = True
    End If
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If vbYes = MsgBox("Deseja Sair do Sistema?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") Then
        End
    Else
        Cancel = 1
    End If
End Sub



Private Sub mnuAtBM_Click()
    Call Baixar_FTP_MOV(Inet1, pgBar)
End Sub

Private Sub mnuAtLog_Click()
    Call Baixar_FTP("prlogin", pgBar)
End Sub

Private Sub mnuBaixa_Click()
    frm_Baixa_Automatica.Show
End Sub

Private Sub mnubco_Click()
    frm_Banco.Show
End Sub

Private Sub mnuCartLoja_Click()
    frm_Cartao_Loja.Show
End Sub

Private Sub mnuCodBon_Click()
    frm_Cod_Bon.Show
End Sub

Private Sub mnuConf_Click()
    frm_Config.Show
End Sub

Private Sub mnuConLanc_Click()
    frm_Lancamento_Pesq.Show
End Sub

Private Sub mnuFPg_Click()
    frm_Forma_Pg.Show
End Sub

Private Sub mnuLanc_Click()
    frm_Lancamento.Show
End Sub

Private Sub mnuRpt_Click()
    frm_Rpt_Cod_Bon.Show
End Sub

Private Sub mnuSisUsu_Click()
    frm_Usuario.Show
End Sub


Private Sub mnuSisSair_Click()
    If vbYes = MsgBox("Deseja Sair do Sistema?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") Then End
End Sub


Private Sub mnuTot_Click()
    frm_Total_Lanc.Show
End Sub

Private Sub mnuTpC_Click()
    frm_TipoCartao.Show
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "sair": mnuSisSair_Click
    Case "lançar": frm_Lancamento.Show
    Case "con_lanc": frm_Lancamento_Pesq.Show
    Case "resumo": frm_Total_Lanc.Show
    Case "cod_bon": frm_Cod_Bon.Show
    Case "rel_codbon": frm_Rpt_Cod_Bon.Show
    Case "baixa": frm_Baixa_Automatica.Show
    Case "rpt_resumo": frm_Rpt_Resumo.Show
    End Select
End Sub


Private Sub timer_at_Timer()
    'Call Baixar_FTP_MOV(Inet1, pgBar)
    timer_at.Enabled = False
End Sub
