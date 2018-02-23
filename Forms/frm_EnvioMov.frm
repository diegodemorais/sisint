VERSION 5.00
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form frm_EnvioMov 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envio de Movimento"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   Icon            =   "frm_EnvioMov.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   600
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VBXPolitec.ocxProgressBarTexto pbar 
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Text            =   " Preparando Caixa  -  CREDIÁRIOS"
      Text            =   " Preparando Caixa  -  CREDIÁRIOS"
      BackColorFundo  =   -2147483643
      MaxProgress     =   100
   End
   Begin Skin_Button.ctr_Button bt_Enviar 
      Height          =   615
      Left            =   2880
      TabIndex        =   2
      Top             =   270
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      BTYPE           =   2
      TX              =   "Enviar"
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
      MICON           =   "frm_EnvioMov.frx":08CA
      PICN            =   "frm_EnvioMov.frx":08E6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin rdActiveText.ActiveText txt_dt 
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
      Left            =   1800
      TabIndex        =   0
      Top             =   570
      Width           =   1035
      _ExtentX        =   1826
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Entre c/ Data do Movimento: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   1785
   End
End
Attribute VB_Name = "frm_EnvioMov"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Conexões com os Arquivos de Movimento
Dim DB_PX_OF As DAO.Database
Dim db_ACCESS As Database
Dim wks As Workspace

Private Sub Form_Load()
On Error Resume Next
    
    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 1000
    
      
    Kill App.Path & "\CMOV\*.*"
    'Cria Arquivo BAT p/ Criar os Diretorios
    Set FSO = CreateObject("Scripting.FileSystemObject")
    caminho = App.Path & "\cmov\copiar.bat"   'especifique aqui o caminho onde ficará/está o TXT
    Set gravar = FSO.CreateTextFile(caminho, True)  'Arquivo Criado
    gravar.Write ("copy " & App.Path & "\modelo\modelo1.mdb " & App.Path & "\modelo\modelo.mdb /y" & vbCrLf)
    gravar.Write ("exit")
    gravar.Close
    
    
    'Roda BAT
     Shell App.Path & "\cmov\copiar.bat", vbHide
    
    txt_dt = Date - 1
    ' Open the first workspace.
    Set wks = DBEngine.Workspaces(0)
    ' Open the database object.
    Set db_ACCESS = wks.OpenDatabase(App.Path & "\MODELO\modelo.mdb")
        
    'Cria as Conexões PX e DBF
    Set DB_PX_OF = DBEngine.OpenDatabase(w_Caminho_SCL, False, False, "Paradox 7.x;pwd=ASPEN_PRESENCE")
        
End Sub


Private Sub bt_Enviar_Click()
Dim w_Access As Access.Application
Set w_Access = New Access.Application

On Error Resume Next
    bt_Enviar.Enabled = False
    txt_dt.Enabled = False
    
    Call Copia_Movs 'Copia os Movs. do Scl p/ pasta CMOV
    
    pbar.Visible = True
    pbar.Value = 1
    pbar.Text = " Preparando Caixa  -  CREDIÁRIOS"
    Pause 1
    
    'Inserir os Contratos Alterados nesta data
    Call Insert_Cred_DataAT 'Inseri os contratos , parcelas e pagamentos ... Dos q/ Foram Atualizados
    Call Insert_Cred_todos 'Inseri todos os contratos , parcelas e pagamentos .... dos Atualizados na Data requerida
    
    pbar.Value = 1
    pbar.Text = "Preparando Caixa - Clientes"
    Call Insert_Clientes 'Inseri os contratos , parcelas e pagamentos ... Dos q/ Foram Atualizados
    
    
    
'**** EXPORTAÇÃO MDB P/ OUTRAS ****
    w_Access.OpenCurrentDatabase App.Path & "\modelo\modelo.mdb", False
    Dim w_Tab As String
    
    pbar.Text = "Exportando Tabelas..........."
    pbar.Value = 1
    Pause 1
    
    'Export Crediarios - Paradox
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "O" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "Paradox 5.x", App.Path & "\CMOV", acTable, "mxxO0000", w_Tab
    
    pbar.Value = 10
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "S" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "Paradox 5.x", App.Path & "\CMOV", acTable, "mxxS0000", w_Tab
    
    pbar.Value = 20
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "P" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "Paradox 5.x", App.Path & "\CMOV", acTable, "mxxP0000", w_Tab
    
    'Export Clientes - DBase
    pbar.Value = 30
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "L" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "DBase IV", App.Path & "\CMOV", acTable, "mxxL0000", w_Tab
    pbar.Value = 40
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "N" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "DBase IV", App.Path & "\CMOV", acTable, "mxxN0000", w_Tab
    pbar.Value = 50
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "T" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "DBase IV", App.Path & "\CMOV", acTable, "mxxT0000", w_Tab
    pbar.Value = 60
     w_Tab = "m" & Right(w_Usu_Nome, 2) & "D" & Format(txt_dt, "ddmm")
    w_Access.DoCmd.TransferDatabase acExport, "DBase IV", App.Path & "\CMOV", acTable, "mxxD0000", w_Tab
    pbar.Value = 70
    
    
    Call Prepara_Arquivos 'Compacta arquivos
    
    
    'deleta o bat de copia
    Kill App.Path & "\cmov\copiar.bat"
    'deleta o bat de compactação
    Kill App.Path & "\cmov\compacta.bat"
    
    
    'Call Send_FTP_MOV(MDI.Inet1, pbar)
    Call SendEmail
    
    
    pbar.Value = 100

    
    pbar.Visible = False
    bt_Enviar.Enabled = True
    txt_dt.Enabled = True
    
End Sub


Sub Insert_Cred_todos()
On Error Resume Next

    wSQL = "Select Cred_loja, N_Cred FROM mxxP0000 Group BY Cred_loja, N_Cred"
    Set w_Rec = db_ACCESS.OpenRecordset(wSQL)
    If IsObject(w_Rec) Then
        Do While Not w_Rec.EOF
            Call LocInsert_081(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            Call LocInsert_082(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            Call LocInsert_118(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            w_Rec.MoveNext
        Loop
    End If
    pbar.Value = 75
    
    wSQL = "Select Cred_loja, N_Cred FROM mxxO0000 Group BY Cred_loja, N_Cred"
    Set w_Rec = db_ACCESS.OpenRecordset(wSQL)
    If IsObject(w_Rec) Then
        Do While Not w_Rec.EOF
            Call LocInsert_082(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            Call LocInsert_118(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            w_Rec.MoveNext
        Loop
    End If
    pbar.Value = 85

    wSQL = "Select Cred_loja, N_Cred FROM mxxS0000 Group BY Cred_loja, N_Cred"
    Set w_Rec = db_ACCESS.OpenRecordset(wSQL)
    If IsObject(w_Rec) Then
        Do While Not w_Rec.EOF
            Call LocInsert_081(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            Call LocInsert_118(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            Call LocInsert_082(w_Rec.Fields("cred_loja"), w_Rec.Fields("n_Cred"))
            w_Rec.MoveNext
        Loop
    End If
    pbar.Value = 95
    
End Sub


Sub Insert_Cred_DataAT()
Dim w_Mov_081, w_Mov_082, w_Mov_118, w_Rec

    Set w_Mov_081 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb081 WHERE (data_at = #" & Format(txt_dt, "mm/dd/yyyy") & "#)")
    
    'INCLUINDO OS 081
    Do While Not w_Mov_081.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_081.Fields("DATA_AT") = CVDate(txt_dt) Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXO0000 " & _
                    "(CRED_LOJA, N_CRED, CLI_LOJA, CODIGO, COND_PGT, N_PARC, DATA_VND, " & _
                    "VALOR_COMPRA, ENTRADA, SALDO, TOTAL, TIPO_DOC, NUMERO, EXCLUIDO, " & _
                    "NOME, CPF_RG, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_081.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_081.Fields("N_CRED") & "#@#, #@#" & w_Mov_081.Fields("CLI_LOJA") & "#@#, #@#" & w_Mov_081.Fields("CODIGO") & "#@#, #@#" & _
                    w_Mov_081.Fields("COND_PGT") & "#@#, #@#" & w_Mov_081.Fields("N_PARC") & "#@#, #@#" & Format(w_Mov_081.Fields("DATA_VND"), "DD/MM/YYYY") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_081.Fields("VALOR_COMPRA")), "0", w_Mov_081.Fields("VALOR_COMPRA")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_081.Fields("ENTRADA")), "0", w_Mov_081.Fields("ENTRADA")), ",", ",") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_081.Fields("SALDO")), "0", w_Mov_081.Fields("SALDO")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_081.Fields("TOTAL")), "0", w_Mov_081.Fields("TOTAL")), ",", ",") & "#@#, #@#" & w_Mov_081.Fields("TIPO_DOC") & "#@#, #@#" & _
                    w_Mov_081.Fields("NUMERO") & "#@#, #@#" & w_Mov_081.Fields("EXCLUIDO") & "#@#, #@#" & IIf(IsNull(w_Mov_081.Fields("NOME")), "", w_Mov_081.Fields("NOME")) & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_081.Fields("CPF_RG")), "", w_Mov_081.Fields("CPF_RG")) & "#@#, #@#" & Format(w_Mov_081.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_081.MoveNext
        
    Loop
    
    pbar.Value = 20
    
    
    
    Set w_Mov_082 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb082 WHERE (data_at = #" & Format(txt_dt, "mm/dd/yyyy") & "#)")
    'INCLUINDO OS 082
    Do While Not w_Mov_082.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_082.Fields("DATA_AT") = CVDate(txt_dt) Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXS0000 " & _
                    "(CRED_LOJA, N_CRED, PARCELA, SITUACAO, CARTORIO, AG_COBRAD, DATA_VNC, " & _
                    "VALOR, SALDO, SPC, EXCLUIDO, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_082.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_082.Fields("N_CRED") & "#@#, #@#" & w_Mov_082.Fields("PARCELA") & "#@#, #@#" & w_Mov_082.Fields("SITUACAO") & "#@#, #@#" & w_Mov_082.Fields("CARTORIO") & "#@#, #@#" & w_Mov_082.Fields("AG_COBRAD") & "#@#, #@#" & Format(w_Mov_082.Fields("DATA_VNC"), "DD/MM/YYYY") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_082.Fields("VALOR")), "0", w_Mov_082.Fields("VALOR")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_082.Fields("SALDO")), "0", w_Mov_082.Fields("SALDO")), ",", ",") & "#@#, #@#" & w_Mov_082.Fields("SPC") & "#@#, #@#" & w_Mov_082.Fields("EXCLUIDO") & "#@#, #@#" & Format(w_Mov_082.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_082.MoveNext
        
    Loop
    
    pbar.Value = 45
    
    Set w_Mov_118 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb118 WHERE (data_at = #" & Format(txt_dt, "mm/dd/yyyy") & "#)")
    'INCLUINDO OS 118
    Do While Not w_Mov_118.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_118.Fields("DATA_AT") = CVDate(txt_dt) Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXP0000 " & _
                    "(CRED_LOJA, N_CRED, PARCELA, CONTROLE, VALOR, JUROS, DATA_PGT, " & _
                    "COD_LOJ, EXCLUIDO, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_118.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_118.Fields("N_CRED") & "#@#, #@#" & w_Mov_118.Fields("PARCELA") & "#@#, #@#" & w_Mov_118.Fields("CONTROLE") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_118.Fields("VALOR")), "0", w_Mov_118.Fields("VALOR")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_118.Fields("JUROS")), "0", w_Mov_118.Fields("JUROS")), ",", ",") & "#@#, #@#" & Format(w_Mov_118.Fields("DATA_PGT"), "YYYY/MM/DD") & "#@#, #@#" & _
                    w_Mov_118.Fields("COD_LOJ") & "#@#, #@#" & w_Mov_118.Fields("EXCLUIDO") & "#@#,  #@#" & Format(w_Mov_118.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_118.MoveNext
        
    Loop
    
    pbar.Value = 65
    
End Sub

Sub LocInsert_081(p_Cred_Loja As String, p_N_Cred As Long)
Dim w_Mov_081
On Error Resume Next

    'Set w_Mov_081 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb081")
    Set w_Mov_081 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb081 WHERE (cred_loja = '" & p_Cred_Loja & "' and n_cred = " & p_N_Cred & ")")
    
    'INCLUINDO OS 081
    Do While Not w_Mov_081.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_081.Fields("Cred_loja") = p_Cred_Loja And w_Mov_081.Fields("N_Cred") = p_N_Cred Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXO0000 " & _
                    "(CRED_LOJA, N_CRED, CLI_LOJA, CODIGO, COND_PGT, N_PARC, DATA_VND, " & _
                    "VALOR_COMPRA, ENTRADA, SALDO, TOTAL, TIPO_DOC, NUMERO, EXCLUIDO, " & _
                    "NOME, CPF_RG, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_081.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_081.Fields("N_CRED") & "#@#, #@#" & w_Mov_081.Fields("CLI_LOJA") & "#@#, #@#" & w_Mov_081.Fields("CODIGO") & "#@#, #@#" & _
                    w_Mov_081.Fields("COND_PGT") & "#@#, #@#" & w_Mov_081.Fields("N_PARC") & "#@#, #@#" & Format(w_Mov_081.Fields("DATA_VND"), "DD/MM/YYYY") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_081.Fields("VALOR_COMPRA")), "0", w_Mov_081.Fields("VALOR_COMPRA")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_081.Fields("ENTRADA")), "0", w_Mov_081.Fields("ENTRADA")), ",", ",") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_081.Fields("SALDO")), "0", w_Mov_081.Fields("SALDO")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_081.Fields("TOTAL")), "0", w_Mov_081.Fields("TOTAL")), ",", ",") & "#@#, #@#" & w_Mov_081.Fields("TIPO_DOC") & "#@#, #@#" & _
                    w_Mov_081.Fields("NUMERO") & "#@#, #@#" & w_Mov_081.Fields("EXCLUIDO") & "#@#, #@#" & IIf(IsNull(w_Mov_081.Fields("NOME")), "", w_Mov_081.Fields("NOME")) & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_081.Fields("CPF_RG")), "", w_Mov_081.Fields("CPF_RG")) & "#@#, #@#" & Format(w_Mov_081.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_081.MoveNext
    Loop
    
End Sub


Sub LocInsert_082(p_Cred_Loja As String, p_N_Cred As Long)
Dim w_Mov_082
On Error Resume Next

'    Set w_Mov_082 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb082")
    Set w_Mov_082 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb082 WHERE (cred_loja = '" & p_Cred_Loja & "' and n_cred = " & p_N_Cred & ")")
    
    'INCLUINDO OS 082
    Do While Not w_Mov_082.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_082.Fields("Cred_loja") = p_Cred_Loja And w_Mov_082.Fields("N_Cred") = p_N_Cred Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXS0000 " & _
                    "(CRED_LOJA, N_CRED, PARCELA, SITUACAO, CARTORIO, AG_COBRAD, DATA_VNC, " & _
                    "VALOR, SALDO, SPC, EXCLUIDO, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_082.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_082.Fields("N_CRED") & "#@#, #@#" & w_Mov_082.Fields("PARCELA") & "#@#, #@#" & w_Mov_082.Fields("SITUACAO") & "#@#, #@#" & w_Mov_082.Fields("CARTORIO") & "#@#, #@#" & w_Mov_082.Fields("AG_COBRAD") & "#@#, #@#" & Format(w_Mov_082.Fields("DATA_VNC"), "DD/MM/YYYY") & "#@#, #@#" & _
                    Replace(IIf(IsNull(w_Mov_082.Fields("VALOR")), "0", w_Mov_082.Fields("VALOR")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_082.Fields("SALDO")), "0", w_Mov_082.Fields("SALDO")), ",", ",") & "#@#, #@#" & w_Mov_082.Fields("SPC") & "#@#, #@#" & w_Mov_082.Fields("EXCLUIDO") & "#@#, #@#" & Format(w_Mov_082.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_082.MoveNext
    Loop
    
End Sub


Sub LocInsert_118(p_Cred_Loja As String, p_N_Cred As Long)
Dim w_Mov_118
On Error Resume Next

'    Set w_Mov_118 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb118")
    Set w_Mov_118 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb118 WHERE (cred_loja = '" & p_Cred_Loja & "' and n_cred = " & p_N_Cred & ")")
    
    'INCLUINDO OS 118
    Do While Not w_Mov_118.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_118.Fields("Cred_loja") = p_Cred_Loja And w_Mov_118.Fields("N_Cred") = p_N_Cred Then
            w_SQL = ""
            w_SQL = "INSERT INTO MXXP0000 " & _
                    "(CRED_LOJA, N_CRED, PARCELA, CONTROLE, VALOR, JUROS, DATA_PGT, " & _
                    "COD_LOJ, EXCLUIDO, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_118.Fields("CRED_LOJA") & "#@#, #@#" & w_Mov_118.Fields("N_CRED") & "#@#, #@#" & w_Mov_118.Fields("PARCELA") & "#@#, #@#" & w_Mov_118.Fields("CONTROLE") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_118.Fields("VALOR")), "0", w_Mov_118.Fields("VALOR")), ",", ",") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_118.Fields("JUROS")), "0", w_Mov_118.Fields("JUROS")), ",", ",") & "#@#, #@#" & Format(w_Mov_118.Fields("DATA_PGT"), "YYYY/MM/DD") & "#@#, #@#" & _
                    w_Mov_118.Fields("COD_LOJ") & "#@#, #@#" & w_Mov_118.Fields("EXCLUIDO") & "#@#,  #@#" & Format(w_Mov_118.Fields("DATA_AT"), "DD/MM/YYYY") & "#@#)"
            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_118.MoveNext
    Loop
    
End Sub



'*********************** clientes - 108,109,110,111,112 ******************************************

Sub Insert_Clientes()
Dim w_Mov_108, w_Rec
    Pause 1
    
    pbar.Value = 20
    '*** Inseri os Cadastros alterados ***
    Set w_Mov_108 = DB_PX_OF.OpenRecordset("SELECT cli_loja, codigo FROM lojb108 WHERE (data_at = #" & Format(txt_dt, "mm/dd/yyyy") & "#)")
    Do While Not w_Mov_108.EOF
        Call LocInsert_108(w_Mov_108.Fields("cli_loja"), w_Mov_108.Fields("codigo"))
        Call LocInsert_109(w_Mov_108.Fields("cli_loja"), w_Mov_108.Fields("codigo"))
        Call LocInsert_110(w_Mov_108.Fields("cli_loja"), w_Mov_108.Fields("codigo"))
        Call LocInsert_111(w_Mov_108.Fields("cli_loja"), w_Mov_108.Fields("codigo"))
        w_Mov_108.MoveNext
        pbar.Value = pbar.Value + 5
    Loop
    
    pbar.Value = 40
    '*** Pega todos os Clientes dos Contratos do Caixa ***
    wSQL = "Select Cli_Loja, Codigo FROM mxxO0000 Group BY Cli_Loja, Codigo"
    Set w_Rec = db_ACCESS.OpenRecordset(wSQL)
    Do While Not w_Rec.EOF
        Call LocInsert_108(w_Rec.Fields("cli_loja"), w_Rec.Fields("codigo"))
        Call LocInsert_109(w_Rec.Fields("cli_loja"), w_Rec.Fields("codigo"))
        Call LocInsert_110(w_Rec.Fields("cli_loja"), w_Rec.Fields("codigo"))
        Call LocInsert_111(w_Rec.Fields("cli_loja"), w_Rec.Fields("codigo"))
        w_Rec.MoveNext
        pbar.Value = pbar.Value + 5
    Loop
    
    pbar.Value = 100
    
End Sub


Sub LocInsert_108(p_Cli_Loja As String, p_Codigo As Long)
Dim w_Mov_108
On Error Resume Next

'    Set w_Mov_108 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb108")
    Set w_Mov_108 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb108 WHERE (cli_loja = '" & p_Cli_Loja & "' and codigo = " & p_Codigo & ")")
    
    'INCLUINDO OS 108
    Do While Not w_Mov_108.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_108.Fields("Cli_loja") = p_Cli_Loja And w_Mov_108.Fields("codigo") = p_Codigo Then
            
            w_SQL = ""
            If IsNull(w_Mov_108.Fields("LIMITE_CR")) Then
                wLimitCr = "0"
            Else
                wLimitCr = Replace(w_Mov_108.Fields("LIMITE_CR"), ",", ",")
            End If
            
            w_SQL = "INSERT INTO mxxl0000 " & _
                    "(CLI_LOJA, CODIGO, COD_AUX, RAZAO, NOME, CREDITO, F_J, " & _
                    "SEXO, NACIONALID, ESTADO_CV, A_V, CPF, TELECHEQUE, " & _
                    "SPC, CODLIBERA, RG, ORGAORG, CONSCGC, CONTATO, " & _
                    "DTNASC, `NATURAL`, UF_NASC, EMAIL, END_CLIENT, " & _
                    "COMPLEM, BAIRRO, CIDADE, ESTADO, CEP, DDD, TELRES, " & _
                    "DDDCOM, TELCOM, DDDFAX, FAX, ENDCARTAS, DTCMP, " & _
                    "CARTAO_PRO, CARTAO_CRD, NRCARTAO, VENDEDOR, " & _
                    "REPRES, REGIAO, PRICMP, TIPOMOV, LANCADO, EXCLUSIVO, " & _
                    "DATA_INC, DATA_REV, LIMITE_CR, NOMEPAI, NOMEMAE, " & _
                    "SITUACAO, CARROPROP, RESPROP, RESTEMPO, FINANC, " & _
                    "OBSERV, TIPO_EMPR, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & w_Mov_108.Fields("CLI_LOJA") & "#@#, #@#" & w_Mov_108.Fields("CODIGO") & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("COD_AUX")), 0, w_Mov_108.Fields("COD_AUX")) & "#@#, #@#" & w_Mov_108.Fields("RAZAO") & "#@#, #@#" & w_Mov_108.Fields("NOME") & "#@#, #@#" & w_Mov_108.Fields("CREDITO") & "#@#, #@#" & w_Mov_108.Fields("F_J") & "#@#, #@#" & _
                    w_Mov_108.Fields("SEXO") & "#@#, #@#" & w_Mov_108.Fields("NACIONALID") & "#@#, #@#" & w_Mov_108.Fields("ESTADO_CV") & "#@#, #@#" & w_Mov_108.Fields("A_V") & "#@#, #@#" & w_CPF & "#@#, #@#" & w_Mov_108.Fields("TELECHEQUE") & "#@#, #@#" & _
                    w_Mov_108.Fields("SPC") & "#@#, #@#" & w_Mov_108.Fields("CODLIBERA") & "#@#, #@#" & w_Mov_108.Fields("RG") & "#@#, #@#" & w_Mov_108.Fields("ORGAORG") & "#@#, #@#" & w_Mov_108.Fields("CONSCGC") & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("CONTATO")), "", w_Mov_108.Fields("CONTATO")) & "#@#, " & _
                    IIf(IsNull(w_Mov_108.Fields("DTNASC")), "null", "#@#" & Format(w_Mov_108.Fields("DTNASC"), "YYYY/MM/DD") & "#@#") & ", #@#" & w_Mov_108.Fields("NATURAL") & "#@#, #@#" & w_Mov_108.Fields("UF_NASC") & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("EMAIL")), "", w_Mov_108.Fields("EMAIL")) & "#@#, #@#" & w_Mov_108.Fields("END_CLIENT") & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("COMPLEM")), "", w_Mov_108.Fields("COMPLEM")) & "#@#, #@#" & w_Mov_108.Fields("BAIRRO") & "#@#, #@#" & w_Mov_108.Fields("CIDADE") & "#@#, #@#" & w_Mov_108.Fields("ESTADO") & "#@#, #@#" & w_Mov_108.Fields("CEP") & "#@#, #@#" & w_Mov_108.Fields("DDD") & "#@#, #@#" & w_Mov_108.Fields("TELRES") & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("DDDCOM")), "", w_Mov_108.Fields("DDDCOM")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("TELCOM")), "", w_Mov_108.Fields("TELCOM")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("DDDFAX")), "", w_Mov_108.Fields("DDDFAX")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("FAX")), "", w_Mov_108.Fields("FAX")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("ENDCARTAS")), "", w_Mov_108.Fields("ENDCARTAS")) & "#@#, " & IIf(IsNull(w_Mov_108.Fields("DTCMP")), "null", "#@#" & Format(w_Mov_108.Fields("DTCMP"), "YYYY/MM/DD") & "#@#") & ", #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("CARTAO_PRO")), "", w_Mov_108.Fields("CARTAO_PRO")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("CARTAO_CRD")), "", w_Mov_108.Fields("CARTAO_CRD")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("NRCARTAO")), "", w_Mov_108.Fields("NRCARTAO")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("VENDEDOR")), "", w_Mov_108.Fields("VENDEDOR")) & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("REPRES")), "", w_Mov_108.Fields("REPRES")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("REGIAO")), "", w_Mov_108.Fields("REGIAO")) & "#@#, " & IIf(IsNull(w_Mov_108.Fields("PRICMP")), "null", "#@#" & Format(w_Mov_108.Fields("PRICMP"), "YYYY/MM/DD") & "#@#") & ", #@#" & IIf(IsNull(w_Mov_108.Fields("TIPOMOV")), "", w_Mov_108.Fields("TIPOMOV")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("LANCADO")), "", w_Mov_108.Fields("LANCADO")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("EXCLUSIVO")), "", w_Mov_108.Fields("EXCLUSIVO")) & "#@#, #@#" & _
                    Format(w_Mov_108.Fields("DATA_INC"), "YYYY/MM/DD") & "#@#, " & IIf(IsNull(w_Mov_108.Fields("DATA_REV")), "null", "#@#" & Format(w_Mov_108.Fields("DATA_REV"), "YYYY/MM/DD") & "#@#") & ", " & wLimitCr & ", #@#" & IIf(IsNull(w_Mov_108.Fields("NOMEPAI")), "", w_Mov_108.Fields("NOMEPAI")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("NOMEMAE")), "", w_Mov_108.Fields("NOMEMAE")) & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("SITUACAO")), "", w_Mov_108.Fields("SITUACAO")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("CARROPROP")), "", w_Mov_108.Fields("CARROPROP")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("RESPROP")), "", w_Mov_108.Fields("RESPROP")) & "#@#, " & IIf(IsNull(w_Mov_108.Fields("RESTEMPO")), "null", "#@#" & w_Mov_108.Fields("RESTEMPO") & "#@#") & ", #@#" & IIf(IsNull(w_Mov_108.Fields("FINANC")), "", w_Mov_108.Fields("FINANC")) & "#@#, #@#" & _
                    IIf(IsNull(w_Mov_108.Fields("OBSERV")), "", w_Mov_108.Fields("OBSERV")) & "#@#, #@#" & IIf(IsNull(w_Mov_108.Fields("TIPO_EMPR")), "", w_Mov_108.Fields("TIPO_EMPR")) & "#@#, #@#" & Format(w_Mov_108.Fields("DATA_AT"), "YYYY/MM/DD") & "#@#)"
        

            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_108.MoveNext
    Loop
    
End Sub



Sub LocInsert_109(p_Cli_Loja As String, p_Codigo As Long)
Dim w_Mov_109
On Error Resume Next

'    Set w_Mov_109 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb109")
    Set w_Mov_109 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb109 WHERE (cli_loja = '" & p_Cli_Loja & "' and codigo = " & p_Codigo & ")")
    
    'INCLUINDO OS 109
    Do While Not w_Mov_109.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_109.Fields("Cli_loja") = p_Cli_Loja And w_Mov_109.Fields("codigo") = p_Codigo Then
            
            w_SQL = ""
            w_SQL = "INSERT INTO mxxn0000 " & _
                    "(CLI_LOJA, CODIGO, CONTROLE, TIPO_REF, LOJA_BANCO, AGENCIA, CONTA, " & _
                    "BAIRRO, DDD, TELEFONE, CONTATO, ESPECIAL, LIM_CREDIT, " & _
                    "ULTCMP, ENDERECO, CIDADE, ESTADO, CEP, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & p_Cli_Loja & "#@#, #@#" & p_Codigo & "#@#, #@#" & w_Mov_109.Fields("CONTROLE") & "#@#, #@#" & w_Mov_109.Fields("TIPO_REF") & "#@#, #@#" & w_Mov_109.Fields("LOJA_BANCO") & "#@#, #@#" & w_Mov_109.Fields("AGENCIA") & "#@#, #@#" & w_Mov_109.Fields("CONTA") & "#@#, #@#" & _
                    w_Mov_109.Fields("BAIRRO") & "#@#, #@#" & w_Mov_109.Fields("DDD") & "#@#, #@#" & w_Mov_109.Fields("TELEFONE") & "#@#, #@#" & w_Mov_109.Fields("CONTATO") & "#@#, #@#" & w_Mov_109.Fields("ESPECIAL") & "#@#, #@#" & Replace(IIf(IsNull(w_Mov_109.Fields("LIM_CREDIT")), "0", w_Mov_109.Fields("LIM_CREDIT")), ",", ",") & "#@#, " & _
                    IIf(IsNull(w_Mov_109.Fields("ULTCMP")), "null", "#@#" & Format(w_Mov_109.Fields("ULTCMP"), "yyyy/mm/dd") & "#@#") & ", " & _
                    "#@#" & w_Mov_109.Fields("ENDERECO") & "#@#, #@#" & w_Mov_109.Fields("CIDADE") & "#@#, #@#" & w_Mov_109.Fields("ESTADO") & "#@#, #@#" & w_Mov_109.Fields("CEP") & "#@#, #@#" & Format(w_Mov_109.Fields("DATA_AT"), "YYYY/MM/DD") & "#@#)"

            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_109.MoveNext
    Loop
    
End Sub

Sub LocInsert_110(p_Cli_Loja As String, p_Codigo As Long)
Dim w_Mov_110
On Error Resume Next

'    Set w_Mov_110 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb110")
    Set w_Mov_110 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb110 WHERE (cli_loja = '" & p_Cli_Loja & "' and codigo = " & p_Codigo & ")")
    
    'INCLUINDO OS 110
    Do While Not w_Mov_110.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_110.Fields("Cli_loja") = p_Cli_Loja And w_Mov_110.Fields("codigo") = p_Codigo Then
            
            w_SQL = ""
            w_SQL = "INSERT INTO mxxt0000 " & _
                    "(CLI_LOJA, CODIGO, TITNOME, TITRG, TITCPF, TITDTNASC, TITEND, " & _
                    "TITCOMPLEM, TITBAIRRO, TITCIDADE, TITESTADO, TITCEP, " & _
                    "TITDDD, TITTELRES, ESTADO_CV, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & p_Cli_Loja & "#@#, #@#" & p_Codigo & "#@#, #@#" & w_Mov_110.Fields("TITNOME") & "#@#, #@#" & w_Mov_110.Fields("TITRG") & "#@#, #@#" & w_Mov_110.Fields("TITCPF") & "#@#, #@#" & Format(w_Mov_110.Fields("TITDTNASC"), "yyyy/mm/dd") & "#@#, #@#" & w_Mov_110.Fields("TITEND") & "#@#, #@#" & _
                    w_Mov_110.Fields("TITCOMPLEM") & "#@#, #@#" & w_Mov_110.Fields("TITBAIRRO") & "#@#, #@#" & w_Mov_110.Fields("TITCIDADE") & "#@#, #@#" & w_Mov_110.Fields("TITESTADO") & "#@#, #@#" & w_Mov_110.Fields("TITCEP") & "#@#, #@#" & w_Mov_110.Fields("TITDDD") & "#@#, #@#" & _
                    w_Mov_110.Fields("TITTELRES") & "#@#, #@#" & w_Mov_110.Fields("ESTADO_CV") & "#@#, #@#" & Format(w_Mov_110.Fields("DATA_AT"), "YYYY/MM/DD") & "#@#)"

            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_110.MoveNext
    Loop
    
End Sub

Sub LocInsert_111(p_Cli_Loja As String, p_Codigo As Long)
Dim w_Mov_111
On Error Resume Next

'    Set w_Mov_111 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb111")
    Set w_Mov_111 = DB_PX_OF.OpenRecordset("SELECT * FROM lojb111 WHERE (cli_loja = '" & p_Cli_Loja & "' and codigo = " & p_Codigo & ")")
    
    'INCLUINDO OS 111
    Do While Not w_Mov_111.EOF
    
        On Error Resume Next
        'SE A DATA_AT FOR IGUAL A DESEJADA DO CAIXA
        If w_Mov_111.Fields("Cli_loja") = p_Cli_Loja And w_Mov_111.Fields("codigo") = p_Codigo Then
            
            w_SQL = ""
            w_SQL = "INSERT INTO mxxd0000 " & _
                    "(CLI_LOJA, CODIGO, CONTROLE, NOME, SEXO, DTNASC, DEPENDENCI, DATA_AT) " & _
                    " VALUES " & _
                    "(#@#" & p_Cli_Loja & "#@#, #@#" & p_Codigo & "#@#, #@#" & w_Mov_111.Fields("CONTROLE") & "#@#, #@#" & _
                    w_Mov_111.Fields("NOME") & "#@#, #@#" & w_Mov_111.Fields("SEXO") & "#@#, #@#" & _
                    Format(w_Mov_111.Fields("DTNASC"), "yyyy/mm/dd") & "#@#, #@#" & w_Mov_111.Fields("DEPENDENCI") & "#@#, #@#" & _
                    Format(w_Mov_111.Fields("DATA_AT"), "YYYY/MM/DD") & "#@#)"

            w_SQL = ReplaceSQL(w_SQL)
            db_ACCESS.Execute w_SQL
        End If
        
        w_Mov_111.MoveNext
    Loop
    
End Sub




'********************** send caixa ****************************************************************

Public Sub Send_FTP_MOV(ByRef Inet1, ByRef pgBar)
On Error Resume Next

On Error GoTo ErroGeral

    Inet1.URL = strFTPHost
    Inet1.UserName = strFTPUser
    Inet1.Password = strFTPPassW
    pgBar.Value = 0
    pgBar.MaxProgress = 10000
    pgBar.Text = "Enviando Caixa ....................."
    Pause 1
    pgBar.Visible = True
        w_Arq = LCase("m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj")
        w_CArq = LCase(App.Path & "\cmov\m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj")
        wTime = Time
        'Cria o Diretorio MODELO - q/ Guarda os Movimentos
        Inet1.Execute , "SEND " & w_CArq & " " & strFTPDir & "/caixas/" & w_Arq
        
        Do
            DoEvents
            pgBar.Value = pgBar.Value + 2
            If pgBar.Value >= 10000 Then pgBar.Value = 0
            If CVDate(Time - wTime) >= CVDate("00:01:00") Then
            
                If vbYes = MsgBox("Sua Conexão está lenta, deseja continuar fazendo atualização?", vbQuestion + vbYesNo) Then
                    wTime = Time()
                Else
                    Exit Do
                End If
                
            End If
        Loop Until Not Inet1.StillExecuting
    
    
    If Inet1.ResponseCode = 12003 Then
        MsgBox "Não foi possível fazer a atualização!" & Chr(13) & Chr(13) & "O arquivo a ser baixado FTP, não foi encontrado!", vbExclamation
    ElseIf Inet1.ResponseCode = 0 Then
    
        pgBar.Value = pgBar.MaxProgress
        pgBar.Text = "Atualizado com sucesso!"
        Pause 0.5
    End If

ErroGeral:
    pgBar.Visible = False
    Pause 0.2
    Set p_fileObj = Nothing

End Sub



Sub Copia_Movs()
On Error Resume Next

   
    'Cria Arquivo BAT p/ Criar os Diretorios
    Set FSO = CreateObject("Scripting.FileSystemObject")
    caminho = App.Path & "\cmov\copiar.bat"   'especifique aqui o caminho onde ficará/está o TXT
    Set gravar = FSO.CreateTextFile(caminho, True)  'Arquivo Criado
    gravar.Write ("copy " & w_Caminho_SCL & "mov\m" & Right(w_Usu_Nome, 2) & "?" & Format(txt_dt, "ddmm") & ".* " & App.Path & "\CMOV\*.* /y" & vbCrLf)
    gravar.Write ("exit")
    gravar.Close
    
    
    'Roda BAT
     Shell App.Path & "\cmov\copiar.bat", vbHide
    
    Pause 5
    

End Sub


Sub Prepara_Arquivos()
On Error Resume Next

    'Cria Arquivo BAT p/ Criar os Diretorios
    Set FSO = CreateObject("Scripting.FileSystemObject")
    caminho = App.Path & "\cmov\compacta.bat"   'especifique aqui o caminho onde ficará/está o TXT
    Set gravar = FSO.CreateTextFile(caminho, True)  'Arquivo Criado
    gravar.Write ("cd " & App.Path & "\cmov" & vbCrLf)
    gravar.Write (LCase("arj a m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj M*" & Format(txt_dt, "ddmm") & ".* ") & vbCrLf)
    gravar.Write ("exit")
    gravar.Close
    
    'Roda BAT
    Shell App.Path & "\cmov\compacta.bat", vbHide
    Pause 5
    
End Sub




Private Sub SendEmail()
Dim Mailer As MailSender  'as Object
Dim recipient
Dim sender
Dim subject
Dim Message
Dim mailserver
Dim result
    result = False
    Set Mailer = New MailSender
    Mailer.Host = "smtp.amr.terra.com.br"
    
    Mailer.AddAddress "stworks@terra.com.br", "RP ASSESSORIA"
    Mailer.MailFrom = "stworks@terra.com.br"
    Mailer.FromName = Right(w_Usu_Nome, 2)
    Mailer.subject = "Caixa - " & Right(w_Usu_Nome, 2) & " - Sistema Integrado" 'Assunto do e-mail
    Mailer.Body = "Caixa do dia " & txt_dt & " da Loja " & Right(w_Usu_Nome, 2) 'Corpo do e-mail
    Mailer.UserName = "stworks"
    Mailer.Password = "d16m12"
    Mailer.AddAttachment App.Path & "\cmov\m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj"  'Pega o arquivo gerado acima como anexo do e-mail
    result = Mailer.Send()
    
    If result = True Then
        MsgBox "Email enviado com sucesso!", vbInformation
    Else
        MsgBox "Erro ao enviar o email , este é a messagem :  " & result, vbCritical
    End If


    Set Mailer = Nothing
End Sub

Private Sub SendEmail_DLL()
On Error GoTo err1

    Dim Mail As vbSendMail.clsSendMail     'aqui começam os códigos para envio por e-mail
    Set Mail = New clsSendMail

    
    Mail.SMTPHost = "smtp.amr.terra.com.br" 'Servidor SMPT
    Mail.From = "stworks@terra.com.br" 'Meu e-mail como "FROM". AI DE VCS SE USAREM O MEU!!
    Mail.FromDisplayName = Right(w_Usu_Nome, 2)  'Nome de exibição pro e-mail do FROM
    Mail.recipient = "stworks@terra.com.br" 'Destinatário do e-mail
    Mail.RecipientDisplayName = "RP ASSESSORIA" 'Nome de exibição do e-mail do destinatário
    Mail.subject = "Caixa - " & Right(w_Usu_Nome, 2) & " - Sistema Integrado" 'Assunto do e-mail
    Mail.Attachment = App.Path & "\cmov\m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj"  'Pega o arquivo gerado acima como anexo do e-mail
    Mail.Message = "Caixa do dia " & txt_dt & " da Loja " & Right(w_Usu_Nome, 2)  'Corpo do e-mail
    
    Mail.SMTPHostValidation = VALIDATE_HOST_SYNTAX
    
    Mail.UseAuthentication = True
    Mail.UserName = "stworks"
    Mail.Password = "d16m12"
    Mail.Send

    
sair:
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
End Sub

Private Sub SendEmail_MAPI()
On Error GoTo err1

  MAPISession1.SignOn

  MAPIMessages1.SessionID = MAPISession1.SessionID

  MAPIMessages1.Compose
  MAPIMessages1.RecipAddress = "stworks@terra.com.br"
  MAPIMessages1.MsgSubject = "Caixa - " & Right(w_Usu_Nome, 2) & " - Sistema Integrado" 'Assunto do e-mail
  MAPIMessages1.MsgNoteText = "Caixa do dia " & txt_dt & " da Loja " & Right(w_Usu_Nome, 2)  'Corpo do e-mail

  'anexa no final da mensagem
  MAPIMessages1.AttachmentPosition = Len(MAPIMessages1.MsgNoteText)
  'define o tipo de dados do anexo
  MAPIMessages1.AttachmentType = mapData
  'da um nome ao anexo
  MAPIMessages1.AttachmentName = "m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj"
  'define o caminho e nome do arquivo a anexar
  MAPIMessages1.AttachmentPathName = App.Path & "\cmov\m" & Right(w_Usu_Nome, 2) & Format(txt_dt, "ddmm") & ".arj"  'Pega o arquivo gerado acima como anexo do e-mail

  'envia o arquivo
  MAPIMessages1.Send False 'True
  MAPISession1.SignOff

sair:
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
End Sub





Public Function ReplaceSQL(p_SQL) As String
Dim w_Str, i, w_Pos

    w_Str = Replace(p_SQL, "\", "\\")
    w_Str = Replace(w_Str, "'", "|+")
    w_Str = Replace(w_Str, "#@#", "'")
    w_Str = Replace(w_Str, "'',", "null,")
    w_Str = Replace(w_Str, "'')", "null)")
    w_Str = Replace(w_Str, "=''", "=null")
    w_Str = Replace(w_Str, "|+", "''")
    ReplaceSQL = w_Str


End Function
