Attribute VB_Name = "Unit"
Public w_Caminho_SCL As String
Public strLoja As String  'define qual a loja instalada
Public W_VER As String    'Versao do EXE
Public w_Data_Server As Date 'Data do servidor  utilizada p/ ninguem fraudar o sistema mudando a data do windows
Public w_ado_Execute As ADODB.Recordset
Public w_SQLString As String

Public w_Usu_Nome As String
Public w_Usu_Cod As String
Public w_Usu_Pass As String
Public w_Usu_Tipo As String
Public w_Usu_Ac As String
Public w_Usu_Rpt As String

Public controle As Boolean

Public Conexão As ADODB.Connection   ' Object
Public strConectaMySQL As String
Public timeCNC As Date
Private i_Provider As String

Private Const timeOut = "00:00:10"
Public LD_Thread As Boolean   'Identificador True ou False da Thread

'Dados Banco Dados
    Public strBDHost, _
           strBDUser, _
           strBDPW, _
           strBDDataBase
    
'Dados p/ Conexao FTP
    Public strFTPHost, _
        strFTPUser, _
        strFTPPassW, _
        strFTPDir

Public w_ado_Logo As ADODB.Recordset, _
       w_ado_Cartao As ADODB.Recordset, _
       w_ado_CadCartao As ADODB.Recordset

Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
ByVal lpKeyName As String, ByVal lpDefault As String, ByVal _
lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName _
As String) As Long

Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'O problema do FileCopy do VB é q ele não mostra visualmente a
'operação (barra de progresso e etc) com no Explorer. Ao invés
'de copiar arquivos com o FileCopy, use a rotina API abaixo:

'Num módulo:

Public Declare Function SHFileOperation Lib "shell32.dll" _
       Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) _
       As Long

Public Const FO_COPY As Long = &H2
Public Const FOF_ALLOWUNDO As Long = &H40

Public Type SHFILEOPSTRUCT
  hwnd As Long
  wFunc As Long
  pFrom As String
  pTo As String
  fFlags As Integer
  fAnyOperationsAborted As Boolean
  hNameMappings As Long
  lpszProgressTitle As String
End Type

Public Sub CopiarArq(Origem As String, Destino As String)
  Dim RST As Long
  Dim FLOP As SHFILEOPSTRUCT

  FLOP.hwnd = 0
  FLOP.wFunc = FO_COPY

  'Arquivo de origem:
  FLOP.pFrom = Origem & vbNullChar & vbNullChar

  'Para copiar TODOS os arquivos, use:
  'FLOP.pFrom = "C:\*.*" & vbNullChar & vbNullChar

  'Diretório ou arquivo de destino:
  FLOP.pTo = Destino & vbNullChar & vbNullChar

  FLOP.fFlags = FOF_ALLOWUNDO
  RST = SHFileOperation(FLOP)
  If RST <> 0 Then
    'Erro na cópia
    MsgBox err.LastDllError, vbCritical Or vbOKOnly
  Else
    If FLOP.fAnyOperationsAborted <> 0 Then
      MsgBox "Falha na cópia!!!", vbCritical Or vbOKOnly
    End If
  End If
End Sub









Sub Pause(Seconds As Single)
Dim EndTime As Date

    EndTime = DateAdd("s", Seconds, Now)
    
    Do
    DoEvents
    Loop Until Now >= EndTime

End Sub


Sub Desligar()
Select Case MsgBox("Deseja Desligar o Computador?", vbInformation + vbYesNoCancel + vbDefaultButton2)
Case 6
    Call ExitWindowsEx(1, 1)
Case 7
    End
End Select

End Sub


Public Function ExecuteSQL(SQLString, Optional ByRef w_RegAf, Optional w_Provider As String, Optional w_ShowProgBar As Boolean = True) As ADODB.Recordset
    Dim i As Byte, w_Err As Byte
    
conectar:
On Error GoTo errCNC
    If Not (InStr(UCase(SQLString), "SELECT") > 0 Or InStr(UCase(SQLString), "CALL")) Then w_ShowProgBar = False
    
    
   'Se o tempo da ultima conexao p/ a q/ será realizada for maior q/ o timeOut
   'entao -  feche e abra a conexao
    If (w_Provider <> "" And w_Provider <> i_Provider) Or (timeOut <= Format(Time() - timeCNC, "hh:mm:ss") And timeCNC <> "00:00:00") Then
        Call FechaConexao(Conexão)
        Call AbreConexao(Conexão, w_Provider)
    End If
    
   timeCNC = Time()
    
    '2º Executando SQL
    w_RegAf = 0
    If InStr(UCase(SQLString), "SELECT") > 0 Or InStr(UCase(SQLString), "CALL") Then
        Set ExecuteSQL = Conexão.Execute(SQLString, w_RegAf).Clone
        w_RegAf = ExecuteSQL.RecordCount
        Set ExecuteSQL.ActiveConnection = Nothing
    Else
        If InStr(SQLString, "DROP") Then On Error Resume Next
        'On Error Resume Next
        Conexão.Execute SQLString, w_RegAf
    End If
sair:
    Exit Function
errCNC:

        Call FechaConexao(Conexão)
        Call AbreConexao(Conexão, w_Provider)
    timeCNC = Time
    w_Err = w_Err + 1
    Pause 0.3
    If w_Err <= 5 Then Resume conectar
    MsgBox msgErro(err), vbCritical, "ExecuteSQL"
    'MsgBox "Erro no ExecutarSQL " & Chr(13) & Chr(13) & "Err: " & Error$, vbCritical, "ExecuteSQL"
    Resume sair
End Function

Private Sub FechaConexao(ByRef Conexão)
On Error Resume Next
        Conexão.Close
        Set Conexão = Nothing
End Sub
Public Sub AbreConexao(ByRef Conexão, Optional ByRef w_Provider)
On Error Resume Next
    '1º - Abrindo Conexão
    Set Conexão = New ADODB.Connection
    
    If (Not w_Provider = "" And i_Provider <> w_Provider) Or w_Provider = "MSDataShape" Then
        Conexão.Provider = w_Provider
    Else
        Conexão.Provider = "MSDASQL.1"
    End If
    i_Provider = w_Provider
    
    Conexão.CursorLocation = adUseClient
    Conexão.Open strConectaMySQL
End Sub


Public Sub KeyEnter(key)
    'Utilizar assim : KeyEnter (KeyAscii)
    Select Case key
    Case 13 'enter
        Sendkeys "{tab}{home}+{end}"
    Case 38 'seta para baixo
        Sendkeys "+{tab}{home}+{end}"
    Case 40 'seta para cima
        Sendkeys "{tab}{home}+{end}"
    End Select
End Sub



Public Function DiaSemana(Num As Integer, Abreviado As Boolean) As String
        Select Case Num 'Verifica Nº do dia da semana para escrever o dia referente
          Case 1:
                 If Abreviado = True Then
                    DiaSemana = "Dom"
                 Else
                    DiaSemana = "Domingo"
                 End If
          Case 2:
                 If Abreviado = True Then
                    DiaSemana = "Seg"
                 Else
                    DiaSemana = "2ª Feira"
                 End If
          Case 3:
                 If Abreviado = True Then
                    DiaSemana = "Ter"
                 Else
                    DiaSemana = "3ª Feira"
                 End If
          Case 4:
                 If Abreviado = True Then
                    DiaSemana = "Qua"
                 Else
                    DiaSemana = "4ª Feira"
                 End If
          Case 5:
                 If Abreviado = True Then
                    DiaSemana = "Qui"
                 Else
                    DiaSemana = "5ª Feira"
                 End If
          Case 6:
                 If Abreviado = True Then
                    DiaSemana = "Sex"
                 Else
                    DiaSemana = "6ª Feira"
                 End If
          Case 7:
                 If Abreviado = True Then
                    DiaSemana = "Sab"
                 Else
                    DiaSemana = "Sábado"
                 End If

         End Select
End Function



Function GetIni(section, key, arq)
    'section = É o que está entre []
    'key = É o nome que se encontra antes do sinal de igual (=)
    'arq = É o nome do arquivo INI
    
    Dim Val As String
    Dim valor As Integer
    
    Val = String$(255, 0)
    valor = GetPrivateProfileString(section, key, "", Val, Len(Val), arq)
    
    If valor = 0 Then
    GetIni = ""
    Else
    GetIni = Left(Val, valor)
    End If

End Function






Public Sub AtualizarGeral()
On Error GoTo err1
    Static w_umaVez As Byte
    Dim w_Usu As String
    
If w_umaVez = 0 Then
       
    frmSplash.pb.MaxProgress = 100
    frmSplash.pb.Value = 5
    frmSplash.pb.Text = "Carregando " & frmSplash.pb.Value & "%"

       
    'ABRE AS TABELAS
       'MySql
       'Se a Loja logar consulta somente os Cartões q/ pertence a ela  -  Loja
       'Senão libera todos os Cartão e de todas lojas  - Adm. ou Usu
        w_Usu = IIf(w_Usu_Tipo = "L", w_Usu_Nome, "%")
        Set w_ado_Logo = ExecuteSQL("SELECT usl_cod, usl_nome, usl_tipo, usl_ac FROM tab_usuario WHERE (usl_tipo = 'L') AND (usl_nome LIKE '" & w_Usu & "')", , , False).Clone
        
    frmSplash.pb.Value = 20
    frmSplash.pb.Text = "Carregando " & frmSplash.pb.Value & "%"
        
        Set w_ado_Cartao = ExecuteSQL("SELECT tab_cartao_loja.ctl_cod, tab_usuario.usl_nome AS Logo, tab_tipo_cartao.tpc_desc AS Cartão, tab_cartao_loja.ctl_txv AS `%-Vista`, tab_cartao_loja.ctl_dias_v AS `Dias-V`, tab_cartao_loja.ctl_vr_des_v AS `Vr Desc - V`, tab_cartao_loja.ctl_txp AS `%-Prazo`, tab_cartao_loja.ctl_dias_p AS `Dias-Pz`, tab_cartao_loja.ctl_vr_des_p AS `Vr Desc - Pz`, tab_cartao_loja.ctl_vr_po AS `%-Pz Adic`, tab_banco.bco_desc AS `Bco Dep`, tab_cartao_loja.ctl_loja, tab_cartao_loja.ctl_tipoc, tab_cartao_loja.ctl_label_ndoc, tab_cartao_loja.ctl_des_parc, tab_cartao_loja.ctl_parc_alta_qt, tab_cartao_loja.ctl_parc_alta_tx FROM tab_tipo_cartao, tab_usuario, { oj tab_cartao_loja LEFT OUTER JOIN tab_banco ON tab_cartao_loja.ctl_banco = tab_banco.bco_cod } WHERE (tab_cartao_loja.ctl_loja = tab_usuario.usl_cod) AND (tab_cartao_loja.ctl_tipoc = tab_tipo_cartao.tpc_cod) AND (tab_usuario.usl_nome LIKE '" & w_Usu & "') ORDER BY tab_usuario.usl_nome, tab_tipo_cartao.tpc_desc", , , False).Clone
        
    frmSplash.pb.Value = 50
    frmSplash.pb.Text = "Carregando " & frmSplash.pb.Value & "%"
        
        Set w_ado_CadCartao = ExecuteSQL("SELECT * FROM tab_tipo_cartao ORDER BY tpc_desc", , , False).Clone
       
       
    frmSplash.pb.Value = 100
    frmSplash.pb.Text = "Carregando " & frmSplash.pb.Value & "%"
       
    Pause 0.5
       
       w_umaVez = 1
End If

sair:
    Exit Sub
err1:
    
    If CDbl(err.Number) = CDbl(424) Then
        MsgBox "Favor fazer a importação de tabelas!", vbCritical
    Else
         'MsgBox Error$ & dePP.Recordsets(i).Source, vbCritical
    End If
    Resume sair
End Sub




Public Function msgErro(ByRef wErr) As String
Dim i, w_count
On Error Resume Next
    


On Error GoTo err1

    If wErr.Number = 0 Or wErr.Number = 91 Or wErr.Number = -2147217871 Or wErr.Number = -2147467259 Then
        msgErro = "Erro: " & wErr.Number & "   -  Perda de conexão com a internet! " & Chr(13) & Chr(13) & "   - Verifique se sua conexão de Internet está funcionando!" & Chr(13) & Chr(13) & "Por Favor feche e abra novamente o sistema!"
    ElseIf wErr = -2147217900 Then
        msgErro = "Cartão já cadastrado p/ este Logo!"
    ElseIf wErr = -2147217871 Then
        MsgBox "Antes de Excluir o tipo de cartão, você precisa excluir as formas de pagamentos relacionadas a mesma!", vbInformation
    Else
        msgErro = wErr.Number & " : " & wErr.Description
    End If

sair:
    Exit Function
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Function



'Procedimento : Cria a Tabela Temporario p/ manipulação do relatorio de Código/Bônus
'Criado em: 23/06/06  por  Rafael Bianchin
Public Sub Monta_SQL_Tab_Tmp(DtI As Date, DtF As Date)
Dim w_Dias As Integer, i As Integer
Dim strSQL As String



On Error GoTo err1
    'strBaseName
    strSQL = "CREATE TABLE `" & strBDDataBase & "`.`tab_tmp_" & w_Usu_Cod & "` ( `Loja` INTEGER UNSIGNED NOT NULL, `usl_cod` DOUBLE , `usl_grupo` DOUBLE , `usl_ordem` DOUBLE"

    w_Dias = DtF - DtI
    
    'Codigo
    For i = 0 To w_Dias
           strSQL = strSQL & ", `class_C_" & Format(DtI + i, "dd") & "` VARCHAR(3)"
           strSQL = strSQL & ", `vr_C_" & Format(DtI + i, "dd") & "` DOUBLE"
    Next i
    strSQL = strSQL & ", `Tot_C` DOUBLE"
    strSQL = strSQL & ", `Tot_Ac_C` DOUBLE"
    
    'Bonus
    For i = 0 To w_Dias
           strSQL = strSQL & ", `class_B_" & Format(DtI + i, "dd") & "` VARCHAR(3)"
           strSQL = strSQL & ", `vr_B_" & Format(DtI + i, "dd") & "` DOUBLE"
    Next i
    strSQL = strSQL & ", `Tot_B` DOUBLE"
    
    'Crediario
    For i = 0 To w_Dias
           strSQL = strSQL & ", `class_D_" & Format(DtI + i, "dd") & "` VARCHAR(3)"
           strSQL = strSQL & ", `vr_D_" & Format(DtI + i, "dd") & "` DOUBLE"
    Next i
    strSQL = strSQL & ", `Tot_D` DOUBLE"
    strSQL = strSQL & ", `Tot_Ac_D` DOUBLE) ENGINE=MyIsam"
    
   
On Error Resume Next
    ExecuteSQL "DROP TABLE `" & strBDDataBase & "`.`tab_tmp_" & w_Usu_Cod & "`", , , False
    
On Error GoTo err1
    frm_Rpt_Cod_Bon.pgBar.Value = 5
    ExecuteSQL strSQL, , , False
    frm_Rpt_Cod_Bon.pgBar.Value = 10
    
sair:
    Inserir_Dados_tab_tmp DtI, DtF
    
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub


Public Sub Inserir_Dados_tab_tmp(DtI As Date, DtF As Date)
Dim wAdoDados As ADODB.Recordset
Dim wAdoLojas As ADODB.Recordset
Dim w_Dias As Integer, w_SumC As Double, w_SumB As Double
Dim strSQL As String

On Error GoTo err1

    Set wAdoDados = ExecuteSQL("SELECT tab_usuario.usl_cod, tab_usuario.usl_nome, tab_usuario.usl_grupo, tab_usuario.usl_ordem, tab_usuario.usl_ac, tab_vnds_bonus.dt_vnd, tab_vnds_bonus.vr_vnd, tab_vnds_bonus.vr_bonus, tab_vnds_bonus.vr_acum, tab_vnds_bonus.vr_cred, tab_vnds_bonus.vr_cred_acum FROM tab_vnds_bonus, tab_usuario WHERE tab_vnds_bonus.usl_cod = tab_usuario.usl_cod AND (tab_vnds_bonus.dt_vnd >= '" & Format(DtI, "yyyy-mm-dd") & "' AND tab_vnds_bonus.dt_vnd <= '" & Format(DtF, "yyyy-mm-dd") & "') ORDER BY tab_usuario.usl_grupo, tab_usuario.usl_ordem, tab_usuario.usl_nome, tab_vnds_bonus.dt_vnd", , , False).Clone
    Set wAdoLojas = ExecuteSQL("SELECT usl_cod, usl_nome, usl_grupo, usl_ordem FROM tab_usuario WHERE usl_tipo = 'L' ORDER BY usl_grupo, usl_ordem", , , False)
    w_Dias = DtF - DtI
    
    Do While Not wAdoLojas.EOF
        'Inseri a Loja na TMP
        strSQL = "INSERT INTO tab_tmp_" & w_Usu_Cod & " (usl_cod, loja, usl_grupo, usl_ordem) VALUES (" & wAdoLojas.Fields("usl_cod") & ", " & Left(wAdoLojas.Fields("usl_nome"), 2) & " , " & wAdoLojas.Fields("usl_grupo") & " , " & wAdoLojas.Fields("usl_ordem") & " )"
        ExecuteSQL strSQL, , , False
        
        wAdoDados.Filter = "usl_cod = '" & wAdoLojas.Fields("usl_cod") & "'"
        
        w_SumC = 0
        w_SumB = 0
        w_SumD = 0
        Do While Not wAdoDados.EOF
            'UPDATE nos campos da TMP
            w_DD = Format(wAdoDados.Fields("dt_vnd"), "dd")
            strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET vr_C_" & w_DD & " = '" & Replace(wAdoDados.Fields("vr_vnd"), ",", ".") & "', vr_B_" & w_DD & " = '" & Replace(wAdoDados.Fields("vr_bonus"), ",", ".") & "', Tot_Ac_C = '" & Replace(wAdoDados.Fields("vr_acum"), ",", ".") & "', vr_D_" & w_DD & " = '" & Replace(wAdoDados.Fields("vr_cred"), ",", ".") & "', Tot_Ac_D = '" & Replace(wAdoDados.Fields("vr_cred_acum"), ",", ".") & "' WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
            
            ExecuteSQL strSQL, , , False
            
            w_SumC = w_SumC + CDbl(wAdoDados.Fields("vr_vnd"))
            w_SumB = w_SumB + CDbl(wAdoDados.Fields("vr_bonus"))
            w_SumD = w_SumD + CDbl(wAdoDados.Fields("vr_cred"))
            wAdoDados.MoveNext
        Loop
        'UPDATE  Tot_C, Tot_B e Tot_D
        strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET Tot_C = '" & Replace(w_SumC, ",", ".") & "', Tot_B = '" & Replace(w_SumB, ",", ".") & "', Tot_D = '" & Replace(w_SumD, ",", ".") & "' WHERE (usl_cod = " & wAdoLojas.Fields("usl_cod") & ")"
        ExecuteSQL strSQL, , , False
        
        
        w_Porc = ((wAdoLojas.AbsolutePosition / wAdoLojas.RecordCount) * 100) * 0.3
        frm_Rpt_Cod_Bon.pgBar.Value = 10 + Int(w_Porc)
        
        wAdoLojas.MoveNext

    Loop
    

sair:
    Classifica_tab_tmp DtI, DtF
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub


'Procedimento :  Classifica as Posições da tab_tmp
Public Sub Classifica_tab_tmp(DtI As Date, DtF As Date)
Dim w_qtGrp As Byte, w_Dias As Byte, G As Byte, d As Byte
Dim wAdoDados As ADODB.Recordset
Dim w_StrFiltro As String

On Error GoTo err1
    
    w_qtGrp = ExecuteSQL("SELECT max(usl_grupo) FROM tab_tmp_" & w_Usu_Cod & "", , , False).Fields(0)
    w_Dias = DtF - DtI
        
    'Looping entre os Grupos
    For G = 1 To w_qtGrp
        
        For d = 0 To w_Dias
            w_Valor = 0
            
'**** Classificando Valor da Venda ****
            'Pega o ultimo Colocado
            Set wAdoDados = ExecuteSQL("SELECT MIN(vr_C_" & Format(DtI + d, "DD") & ") as menor, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (Min(vr_C_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_C_" & Format(DtI + d, "DD") & "", , , False).Clone
                            
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS ULTIMOS COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_C
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_C_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "***", IIf(U = 2, "**", "*")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    ExecuteSQL strSQL, , , False
                    If Not w_Valor = wAdoDados.Fields("menor") Then
                        w_Valor = wAdoDados.Fields("menor")
                    Else
                        U = U - 1
                    End If
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
            
            
            '*** PEGA OS PRIMEIROs COLOCADOS ***
            w_Valor = 0
            
            'Pega o Primeiros Colocado
            Set wAdoDados = ExecuteSQL("SELECT MAX(vr_C_" & Format(DtI + d, "DD") & ") as Maior, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (MAX(vr_C_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_C_" & Format(DtI + d, "DD") & " DESC", , , False).Clone
                            
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS PRIMEIROs COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_C
                    If Not w_Valor = wAdoDados.Fields("maior") Then
                        w_Valor = wAdoDados.Fields("maior")
                    Else
                        U = U - 1
                    End If
                    
                    strSQL = ""
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_C_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "1º", IIf(U = 2, "2º", "3º")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    Call ExecuteSQL(strSQL, , , False)
                    
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
        
        
'**** Classificando Bônus ****
            'Pega o ultimo Colocado
            Set wAdoDados = ExecuteSQL("SELECT MIN(vr_B_" & Format(DtI + d, "DD") & ") as menor, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (Min(vr_B_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_B_" & Format(DtI + d, "DD") & "", , , False).Clone
            w_Valor = 0
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS ULTIMOS COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_C
                    strSQL = ""
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_B_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "***", IIf(U = 2, "**", "*")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    ExecuteSQL strSQL, , , False
                    If Not w_Valor = wAdoDados.Fields("menor") Then
                        w_Valor = wAdoDados.Fields("menor")
                    Else
                        U = U - 1
                    End If
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
            
            
            '*** PEGA OS PRIMEIROs COLOCADOS ***
            w_Valor = 0
            
            'Pega o Primeiros Colocado
            Set wAdoDados = ExecuteSQL("SELECT MAX(vr_B_" & Format(DtI + d, "DD") & ") as Maior, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (MAX(vr_B_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_B_" & Format(DtI + d, "DD") & " DESC", , , False).Clone
                            
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS PRIMEIROs COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_C
                    strSQL = ""
                    If Not w_Valor = wAdoDados.Fields("maior") Then
                        w_Valor = wAdoDados.Fields("maior")
                    Else
                        U = U - 1
                    End If
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_B_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "1º", IIf(U = 2, "2º", "3º")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    ExecuteSQL strSQL, , , False
                    
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
        
        
    '**** Classificando Valor do Crediario ****
            'Pega o ultimo Colocado
            Set wAdoDados = ExecuteSQL("SELECT MIN(vr_D_" & Format(DtI + d, "DD") & ") as menor, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (Min(vr_D_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_D_" & Format(DtI + d, "DD") & "", , , False).Clone
                            
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS ULTIMOS COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_D
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_D_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "***", IIf(U = 2, "**", "*")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    ExecuteSQL strSQL, , , False
                    If Not w_Valor = wAdoDados.Fields("menor") Then
                        w_Valor = wAdoDados.Fields("menor")
                    Else
                        U = U - 1
                    End If
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
            
            
            '*** PEGA OS PRIMEIROs COLOCADOS ***
            w_Valor = 0
            
            'Pega o Primeiros Colocado
            Set wAdoDados = ExecuteSQL("SELECT MAX(vr_D_" & Format(DtI + d, "DD") & ") as Maior, usl_cod FROM tab_tmp_" & w_Usu_Cod & " WHERE usl_grupo = " & G & " GROUP BY usl_cod Having (Not (MAX(vr_D_" & Format(DtI + d, "DD") & ") Is Null)) ORDER BY vr_D_" & Format(DtI + d, "DD") & " DESC", , , False).Clone
                            
            If wAdoDados.RecordCount > 0 Then
                '*** PEGA OS PRIMEIROs COLOCADOS ***
                For U = 1 To 3   'As 3  posições
                    'ATUALIZA O CAMPOS CLASS_D
                    If Not w_Valor = wAdoDados.Fields("maior") Then
                        w_Valor = wAdoDados.Fields("maior")
                    Else
                        U = U - 1
                    End If
                    
                    strSQL = ""
                    strSQL = "UPDATE tab_tmp_" & w_Usu_Cod & " SET class_D_" & Format(DtI + d, "DD") & "  = '" & IIf(U = 1, "1º", IIf(U = 2, "2º", "3º")) & "'  WHERE (usl_cod = " & wAdoDados.Fields("usl_cod") & ")"
                    Call ExecuteSQL(strSQL, , , False)
                    
                    wAdoDados.MoveNext
                    If wAdoDados.EOF Then Exit For
                Next U
            End If
        
        
        
        
        
        Next d
        
        frm_Rpt_Cod_Bon.pgBar.Value = frm_Rpt_Cod_Bon.pgBar.Value + 5
    Next G
        
sair:
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub


Public Sub CRIA_RPT_EXCEL(DtI As Date, DtF As Date)
Dim xlA As New Excel.Application
Dim wAdoDados As ADODB.Recordset
Dim w_qtDias As Byte

Dim wPlanName As String
    
On Error GoTo err1


    wPlanName = "Plan"
    'wPlanName = "Sheet"
    w_qtDias = DtF - DtI

    'Pega os registros
    Set wAdoDados = ExecuteSQL("SELECT * FROM tab_tmp_" & w_Usu_Cod & "", , , False).Clone


    'Cria Aplication
    Set xlA = CreateObject("Excel.application")
    
    
    
    'Add Novo arquivo do Excel
    Set xl = xlA.Workbooks.Add()
    
    'xlA.Visible = True
    
    'Muda o Nome da Planilha
    xl.Sheets(wPlanName & "1").Select
    xl.Sheets.Item(wPlanName & "1").Name = "Cod"
    'DESATIVA O DISPLAY DE ALERTA DO EXCEL
    xlA.DisplayAlerts = False
    'apaga as outras planilhas
    xl.Sheets.Item(wPlanName & "2").Delete
    xl.Sheets.Item(wPlanName & "3").Delete
    xlA.DisplayAlerts = True
    
    xlA.Cells.Select
    xlA.Selection.Font.Bold = True
    
    Inseri_Cab_CodBon_Excel wAdoDados, xl, xlA, IIf(w_qtDias > 0, True, False)
    If w_qtDias > 0 Then
        Inseri_Dados_CodBon_Excel_Vnd DtI, DtF, wAdoDados, xl, xlA
    Else
        Inseri_Dados_CodBon_Excel_Vnd_Bonus DtI, DtF, wAdoDados, xl, xlA
    End If
    
sair:
    Exit Sub
err1:
    MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub

Sub Dimensiona_Margem_Print(ByRef wXLA)
    With wXLA.ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    wXLA.ActiveSheet.PageSetup.PrintArea = ""
    With wXLA.ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.78740157480315)
        .RightMargin = Application.InchesToPoints(0.78740157480315)
        .TopMargin = Application.InchesToPoints(0.393700787401575)
        .BottomMargin = Application.InchesToPoints(0.393700787401575)
        .HeaderMargin = Application.InchesToPoints(0.511811023622047)
        .FooterMargin = Application.InchesToPoints(0.511811023622047)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 300
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 75
    End With
End Sub


Sub Inseri_Cab_CodBon_Excel(wAdo As ADODB.Recordset, ByRef wXl, ByRef wXLA, wComTotAcum As Boolean)
Dim w_Col As Byte

    wXl.Sheets(1).Columns(1).ColumnWidth = 3
    
    
'MUDA O TAMANHO DA FONTE DAS CELULAS
    wXLA.Cells.Select
    With wXLA.Selection.Font
        .Name = "Arial"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = xlAutomatic
    End With



    
    
If wComTotAcum Then
    'Looping entre todos os campos
    For i = 0 To wAdo.Fields.Count - 1
        'Insere somente os campos vr_C  -  Colocar a Data deles
        If (Left(wAdo.Fields.Item(i).Name, 4) = "vr_C") Then
            w_Col = w_Col + 2
            wXl.Sheets(1).Cells(1, w_Col).Select
            wXl.Sheets(1).Columns(w_Col).ColumnWidth = 4
            wXl.Sheets(1).Columns(w_Col + 1).ColumnWidth = 17
            
            'Alinhamento da Coluna Classificação - Center
            wXl.Sheets(1).Columns(w_Col).Select
            With wXLA.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
            End With
            
            'Alinhamento da Coluna VALOR - Right
            wXl.Sheets(1).Columns(w_Col + 1).Select
            With wXLA.Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
            End With
            
            
            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(1, w_Col).Value = Mid(wAdo.Fields.Item(i).Name, 6)
            wXl.Sheets(1).Cells(1, w_Col).Select
            
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            wXl.Sheets(1).Cells(1, w_Col + 1).Select
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            
            '*** Mescla Celula ***
            wXLA.Range(Cel_Ini & ":" & Cel_Fim).Select
            With wXLA.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
                .Font.ColorIndex = 3
                .Font.Bold = True
            End With
            wXLA.Selection.Merge
          
        End If
    Next i
    
    If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
     
        'Insere o Total, o ACumulado e o Bonus
        w_Col = w_Col + 2
        wXl.Sheets(1).Columns(w_Col).ColumnWidth = 10
        'Pega os Endereços das Celulas
        wXl.Sheets(1).Cells(1, w_Col).Value = "TT"
        wXl.Sheets(1).Columns(w_Col).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
        
        wXl.Sheets(1).Cells(1, w_Col + 1).Value = "Acum"
        wXl.Sheets(1).Columns(w_Col + 1).ColumnWidth = 10
        wXl.Sheets(1).Columns(w_Col + 1).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
        
        wXl.Sheets(1).Cells(1, w_Col + 2).Value = "Bonus"
        wXl.Sheets(1).Columns(w_Col + 2).ColumnWidth = 10
        wXl.Sheets(1).Columns(w_Col + 2).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If
    
Else 'wComTotAcum = false

    
    'Looping entre todos os campos
    For i = 0 To wAdo.Fields.Count - 1
        'Insere somente os campos vr_C -  Colocar a Data deles
        If (Left(wAdo.Fields.Item(i).Name, 4) = "vr_C") Then
            w_Col = w_Col + 2
            wXl.Sheets(1).Cells(1, w_Col).Select
            wXl.Sheets(1).Columns(w_Col).ColumnWidth = 4
            wXl.Sheets(1).Columns(w_Col + 1).ColumnWidth = 17
            wXl.Sheets(1).Columns(w_Col + 2).ColumnWidth = 4
            wXl.Sheets(1).Columns(w_Col + 2 + 1).ColumnWidth = 17
            wXl.Sheets(1).Columns(w_Col + 4).ColumnWidth = 3
            wXl.Sheets(1).Columns(w_Col + 5).ColumnWidth = 4
            wXl.Sheets(1).Columns(w_Col + 5 + 1).ColumnWidth = 17
            
            For B = 0 To 2
                'Alinhamento da Coluna Classificação - Center
                c = B
                If c = 2 Then c = 2.5
                wXl.Sheets(1).Columns(w_Col + (c * 2)).Select
                With wXLA.Selection
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                End With
                
                'Alinhamento da Coluna Classificação - Right
                wXl.Sheets(1).Columns(w_Col + (B * 2) + 1).Select
                With wXLA.Selection
                    .HorizontalAlignment = xlRight
                    .VerticalAlignment = xlBottom
                    .WrapText = False
                    .Orientation = 0
                    .AddIndent = False
                    .ShrinkToFit = False
                End With
                
                
            Next B
            
            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(1, 2).Value = Mid(wAdo.Fields.Item(i).Name, 6)
            wXl.Sheets(1).Cells(1, 2).Select
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            wXl.Sheets(1).Cells(1, 8).Select
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            
            
            '*** Mescla Celula ***
            wXLA.Range(Cel_Ini & ":" & Cel_Fim).Select
            With wXLA.Selection
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .ShrinkToFit = False
                .MergeCells = False
                .Font.ColorIndex = 3
                .Font.Bold = True
            End With
            wXLA.Selection.Merge
          
        End If
    Next i


    If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
     
        'Insere o Total, o ACumulado e o Bonus
        w_Col = w_Col + 4
        wXl.Sheets(1).Columns(w_Col).ColumnWidth = 10
        'Pega os Endereços das Celulas
        wXl.Sheets(1).Cells(1, w_Col).Value = "TT"
        wXl.Sheets(1).Columns(w_Col).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
        
        wXl.Sheets(1).Cells(1, w_Col + 1).Value = "Acum"
        wXl.Sheets(1).Columns(w_Col + 1).ColumnWidth = 10
        wXl.Sheets(1).Columns(w_Col + 1).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
        
        wXl.Sheets(1).Cells(1, w_Col + 2).Value = "Bonus"
        wXl.Sheets(1).Columns(w_Col + 2).ColumnWidth = 10
        wXl.Sheets(1).Columns(w_Col + 2).Select
        With wXLA.Selection
            .HorizontalAlignment = xlRight
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .ShrinkToFit = False
            .Font.ColorIndex = 3
            .Font.Bold = True
        End With
    End If

End If


End Sub


Sub Inseri_Dados_CodBon_Excel_Vnd(DtI As Date, DtF As Date, wAdo As ADODB.Recordset, ByRef wXl, ByRef wXLA)
Dim w_qtGrp  As Byte, G As Byte, w_Row As Integer
Dim w_SomaC As Double
Dim w_arr_Tot() As Double
Dim w_arr_Tot_G() As Double

    w_qtGrp = ExecuteSQL("SELECT max(usl_grupo) FROM tab_tmp_" & w_Usu_Cod & "", , , False).Fields(0)
    w_Dias = DtF - DtI
    w_Row = 2
    ReDim w_arr_Tot_G(w_Dias + 3)
    
    For G = 1 To w_qtGrp
        wAdo.Filter = "usl_grupo = " & G
        w_SomaC = 0
        ReDim w_arr_Tot(w_Dias + 3)
        
        Do While Not wAdo.EOF
            If frm_Rpt_Cod_Bon.ckLogo.Value = 1 And frm_Rpt_Cod_Bon.ckSup.Value = 0 And G = 5 Then Exit Do
        
            w_Col = 0
            wXl.Sheets(1).Cells(w_Row, w_Col + 1).Value = wAdo.Fields("Loja")
            
            
            For d = 0 To w_Dias
            
                w_Col = w_Col + 2
                
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col).Value = IIf(IsNull(wAdo.Fields("class_C_" & Format(DtI + d, "dd"))), "", wAdo.Fields("class_C_" & Format(DtI + d, "dd")))
                wXl.Sheets(1).Cells(w_Row, w_Col + 1).Value = Formata_Valor_Logo(IIf(IsNull(wAdo.Fields("vr_C_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_C_" & Format(DtI + d, "dd"))))
                w_arr_Tot(d) = w_arr_Tot(d) + IIf(IsNull(wAdo.Fields("vr_C_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_C_" & Format(DtI + d, "dd")))
                
                If Right(wXl.Sheets(1).Cells(w_Row, w_Col).Value, 1) = "º" And Not G = 5 Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + i).Select
                        With wXLA.Selection.Interior
                            .ColorIndex = 15
                            .Pattern = xlSolid
                            
                            If "1º" = wXl.Sheets(1).Cells(w_Row, w_Col).Value Then
                                'muda a cor do 1º coloca e tambem a loja
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                                
                                wXl.Sheets(1).Cells(w_Row, 1).Select
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                            End If
                        End With
                        
                    Next i
                End If
                
                'Grupo 5  CC   e  MC
                'Muda a cor da Letra pra Vermelho
                If G = 5 Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + i).Select
                        wXLA.Selection.Font.ColorIndex = 3
                    Next i
                End If
                
                
            Next d
            
            wXl.Sheets(1).Cells(w_Row, w_Col + 2).Select
            wXLA.Selection.Font.ColorIndex = 7
            If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
                wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value = Formata_Valor_Logo(wAdo.Fields("Tot_C"))
                wXl.Sheets(1).Cells(w_Row, w_Col + 3).Value = Formata_Valor_Logo(wAdo.Fields("Tot_AC_C"))
                wXl.Sheets(1).Cells(w_Row, w_Col + 4).Value = Formata_Valor_Logo(wAdo.Fields("Tot_B"))
            End If
            w_arr_Tot(w_Dias + 1) = w_arr_Tot(w_Dias + 1) + IIf(IsNull(wAdo.Fields("Tot_C")), 0, wAdo.Fields("Tot_C"))
            w_arr_Tot(w_Dias + 2) = w_arr_Tot(w_Dias + 2) + IIf(IsNull(wAdo.Fields("Tot_AC_C")), 0, wAdo.Fields("Tot_AC_C"))
            w_arr_Tot(w_Dias + 3) = w_arr_Tot(w_Dias + 3) + IIf(IsNull(wAdo.Fields("Tot_B")), 0, wAdo.Fields("Tot_B"))
            
            w_Row = w_Row + 1
            
            wAdo.MoveNext
        Loop
        'Acabou o Grupo - Coloca o Total
        w_Col = 1
        For d = 0 To w_Dias + IIf(frm_Rpt_Cod_Bon.ckAcum.Value = 1, 3, 0)
            If d <= w_Dias Then
                w_Col = w_Col + 2
            Else
                w_Col = w_Col + 1
            End If
            
            If Not G = 5 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
                wXl.Sheets(1).Cells(w_Row, w_Col).Value = Formata_Valor_Logo(w_arr_Tot(d))
                wXl.Sheets(1).Cells(w_Row, w_Col).Select
                wXLA.Selection.Font.ColorIndex = 5
            End If
            w_arr_Tot_G(d) = w_arr_Tot_G(d) + w_arr_Tot(d)
            
            If G = 4 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
                w_Row = w_Row + 1
                wXl.Sheets(1).Cells(w_Row, w_Col).Value = Formata_Valor_Logo(w_arr_Tot_G(d))
                wXl.Sheets(1).Cells(w_Row, w_Col).Select
                wXLA.Selection.Font.ColorIndex = 3
                w_Row = w_Row - 1
            End If
        Next d
        If G = 4 Then

            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(1, 1).Select
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
                wXl.Sheets(1).Cells(w_Row, ((w_Dias + 1) * 2) + 4).Select
            Else
                wXl.Sheets(1).Cells(w_Row, ((w_Dias + 1) * 2) + 1).Select
            End If
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            Dim strCel_INI_FIM As String
            strCel_INI_FIM = Cel_Ini & ":" & Cel_Fim
            Inseri_Borda strCel_INI_FIM, wXLA, "G"
            
            w_Row = w_Row + 1
        
        End If
        w_Row = w_Row + 1
        frm_Rpt_Cod_Bon.pgBar.Value = frm_Rpt_Cod_Bon.pgBar.Value + 5
    Next G
    
    'inseri os bonus
    Inseri_Dados_CodBon_Excel_Bonus DtI, DtF, wAdo, wXl, wXLA, w_Row
    'Dimensiona_Margem_Print wXLA
    wXLA.Visible = True
    
End Sub

Sub Inseri_Dados_CodBon_Excel_Bonus(DtI As Date, DtF As Date, wAdo As ADODB.Recordset, ByRef wXl, ByRef wXLA, ByRef w_Row)
Dim w_qtGrp  As Byte, G As Byte, w_Row_I As Integer
Dim w_SomaC As Double
Dim w_arr_Tot() As Double
Dim w_arr_Tot_G() As Double
    
    w_qtGrp = ExecuteSQL("SELECT max(usl_grupo) FROM tab_tmp_" & w_Usu_Cod & "", , , False).Fields(0)
    w_Dias = DtF - DtI
    w_Row_I = w_Row
    ReDim w_arr_Tot_G(w_Dias + 3)
    
    For G = 1 To 4
        wAdo.Filter = "usl_grupo = " & G
        w_SomaC = 0
        ReDim w_arr_Tot(w_Dias + 3)
        
        Do While Not wAdo.EOF
            If frm_Rpt_Cod_Bon.ckLogo.Value = 1 And frm_Rpt_Cod_Bon.ckSup.Value = 0 And G = 5 Then Exit Do
            
            w_Col = 0
            wXl.Sheets(1).Cells(w_Row, w_Col + 1).Value = wAdo.Fields("Loja")
            For d = 0 To w_Dias
            
                w_Col = w_Col + 2
                
                wXl.Sheets(1).Cells(w_Row, w_Col).Value = IIf(IsNull(wAdo.Fields("class_B_" & Format(DtI + d, "dd"))), "", wAdo.Fields("class_B_" & Format(DtI + d, "dd")))
                wXl.Sheets(1).Cells(w_Row, w_Col + 1).Value = CDbl((IIf(IsNull(wAdo.Fields("vr_B_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_B_" & Format(DtI + d, "dd")))))
                w_arr_Tot(d) = w_arr_Tot(d) + IIf(IsNull(wAdo.Fields("vr_B_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_B_" & Format(DtI + d, "dd")))
                
                If Right(wXl.Sheets(1).Cells(w_Row, w_Col).Value, 1) = "º" Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + i).Select
                        With wXLA.Selection.Interior
                            .ColorIndex = 15
                            .Pattern = xlSolid
                            If "1º" = wXl.Sheets(1).Cells(w_Row, w_Col).Value Then
                                'muda a cor do 1º coloca e tambem a loja
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                                
                                wXl.Sheets(1).Cells(w_Row, 1).Select
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                            End If
                        End With
                    Next i
                End If
                
            Next d
            
            If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
                wXl.Sheets(1).Cells(w_Row, w_Col + 2).Select
                wXLA.Selection.Font.ColorIndex = 7
                wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value = CDbl((wAdo.Fields("Tot_B")))
            End If
            
            w_arr_Tot(w_Dias + 1) = w_arr_Tot(w_Dias + 1) + IIf(IsNull(wAdo.Fields("Tot_B")), 0, wAdo.Fields("Tot_B"))
            
            w_Row = w_Row + 1
            
            wAdo.MoveNext
        Loop
        'Acabou o Grupo - Coloca o Total
        w_Col = 1
        For d = 0 To w_Dias + IIf(frm_Rpt_Cod_Bon.ckAcum.Value = 1, 3, 0)
            If d <= w_Dias Then
                w_Col = w_Col + 2
            Else
                w_Col = w_Col + 1
            End If
            
            If Not G = 5 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
                wXl.Sheets(1).Cells(w_Row, w_Col).Value = CDbl(w_arr_Tot(d))
                wXl.Sheets(1).Cells(w_Row, w_Col).Select
                wXLA.Selection.Font.ColorIndex = 5
            End If
            w_arr_Tot_G(d) = w_arr_Tot_G(d) + w_arr_Tot(d)
            
            If G = 4 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
                w_Row = w_Row + 1
                wXl.Sheets(1).Cells(w_Row, w_Col).Value = CDbl(w_arr_Tot_G(d))
                wXl.Sheets(1).Cells(w_Row, w_Col).Select
                wXLA.Selection.Font.ColorIndex = 3
                w_Row = w_Row - 1
            End If
        Next d
        
        If G = 4 Then

            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(w_Row_I, 1).Select
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            If frm_Rpt_Cod_Bon.ckAcum.Value = 1 Then
                wXl.Sheets(1).Cells(w_Row, ((w_Dias + 1) * 2) + 4).Select
            Else
                wXl.Sheets(1).Cells(w_Row, ((w_Dias + 1) * 2) + 1).Select
            End If
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            Dim strCel_INI_FIM As String
            strCel_INI_FIM = Cel_Ini & ":" & Cel_Fim
            Inseri_Borda strCel_INI_FIM, wXLA, "G"
            
            
            w_Row = w_Row + 1
        
        End If
        w_Row = w_Row + 1
        
    frm_Rpt_Cod_Bon.pgBar.Value = frm_Rpt_Cod_Bon.pgBar.Value + 5
    Next G
    
    
End Sub




Sub Inseri_Dados_CodBon_Excel_Vnd_Bonus(DtI As Date, DtF As Date, wAdo As ADODB.Recordset, ByRef wXl, ByRef wXLA)
Dim w_qtGrp  As Byte, G As Byte, w_Row As Integer
Dim w_Tot_C As Double, w_Tot_B As Double, w_Tot_C_G As Double, w_Tot_b_G As Double, w_Tot_Ac As Double, w_Tot_C_Ac As Double
Dim w_Dia As String, w_Dias As Byte

On Error GoTo err1

    w_Dia = Format(DtI, "dd")
    w_qtGrp = ExecuteSQL("SELECT max(usl_grupo) FROM tab_tmp_" & w_Usu_Cod & "", , , False).Fields(0)
    w_Dias = DtF - DtI
    w_Row = 2
    
    For G = 1 To w_qtGrp
        wAdo.Filter = "usl_grupo = " & G
        w_SomaC = 0
        w_Tot_C = 0
        w_Tot_B = 0
        w_Tot_C_Ac = 0
        w_Tot_D = 0
        w_Tot_D_Ac = 0
        
        'Looping entre os registro do Grupo Atual
        Do While Not wAdo.EOF
                If frm_Rpt_Cod_Bon.ckLogo.Value = 1 And frm_Rpt_Cod_Bon.ckSup.Value = 0 And G = 5 Then Exit Do
                w_Col = 0
                wXl.Sheets(1).Cells(w_Row, 1).Value = wAdo.Fields("Loja")
                wXl.Sheets(1).Cells(w_Row, 6).Value = wAdo.Fields("Loja")

                w_Col = w_Col + IIf(w_Col = 0, 2, 4)
                'INSERI Class. da Venda , Valor da Vendas, Class. do Bonus e o valor do Bonus
                
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col).Value = IIf(IsNull(wAdo.Fields("class_C_" & w_Dia)), "", wAdo.Fields("class_C_" & w_Dia))
                wXl.Sheets(1).Cells(w_Row, w_Col + 1).Value = Formata_Valor_Logo(IIf(IsNull(wAdo.Fields("vr_C_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_C_" & Format(DtI + d, "dd"))))
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value = IIf(IsNull(wAdo.Fields("class_B_" & w_Dia)), "", wAdo.Fields("class_B_" & w_Dia))
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col + 3).Value = CDbl(IIf(IsNull(wAdo.Fields("vr_B_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_B_" & Format(DtI + d, "dd"))))
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value = IIf(IsNull(wAdo.Fields("class_D_" & w_Dia)), "", wAdo.Fields("class_D_" & w_Dia))
                If Not G = 5 Then wXl.Sheets(1).Cells(w_Row, w_Col + 6).Value = CDbl(IIf(IsNull(wAdo.Fields("vr_D_" & Format(DtI + d, "dd"))), 0, wAdo.Fields("vr_D_" & Format(DtI + d, "dd"))))
                
                If frm_Rpt_Cod_Bon.ckAcum = 1 Then
                    wXl.Sheets(1).Cells(w_Row, w_Col + 4).Value = Formata_Valor_Logo(wAdo.Fields("Tot_C"))
                    wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value = Formata_Valor_Logo(wAdo.Fields("Tot_Ac_C"))
                    wXl.Sheets(1).Cells(w_Row, w_Col + 6).Value = Formata_Valor_Logo(wAdo.Fields("Tot_B"))
                End If
                
                
                'Somatoria Acumulativa Venda e Bonus
                w_Tot_C = w_Tot_C + IIf(IsNull(wAdo.Fields("vr_C_" & w_Dia)), 0, wAdo.Fields("vr_C_" & w_Dia))
                w_Tot_B = w_Tot_B + IIf(IsNull(wAdo.Fields("vr_B_" & w_Dia)), 0, wAdo.Fields("vr_B_" & w_Dia))
                w_Tot_C_Ac = w_Tot_C_Ac + IIf(IsNull(wAdo.Fields("Tot_Ac_C")), 0, wAdo.Fields("Tot_Ac_C"))
                w_Tot_D = w_Tot_D + IIf(IsNull(wAdo.Fields("vr_D_" & w_Dia)), 0, wAdo.Fields("vr_D_" & w_Dia))
                w_Tot_D_Ac = w_Tot_D_Ac + IIf(IsNull(wAdo.Fields("Tot_Ac_D")), 0, wAdo.Fields("Tot_Ac_D"))
                
                'vENDA - CODIGO
                'SE TIVER ALGUMA POSIÇÃO NA CLASSIFICAÇÃO - ENTÃO PINTA O FUNDO E MUDAR A COR DO 1º P/ VERMELHO
                If Right(wXl.Sheets(1).Cells(w_Row, w_Col).Value, 1) = "º" And Not G = 5 Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + i).Select
                        With wXLA.Selection.Interior
                            .ColorIndex = 15
                            .Pattern = xlSolid
                            
                            If "1º" = wXl.Sheets(1).Cells(w_Row, w_Col).Value Then
                                'muda a cor do 1º coloca e tambem a loja
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                                
                                wXl.Sheets(1).Cells(w_Row, 1).Select
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                            End If

                        End With
                    Next i
                End If
                
                'bONUS
                'SE TIVER ALGUMA POSIÇÃO NA CLASSIFICAÇÃO - ETNÃO PINTA O FUNDO E MUDAR A COR DO 1º P/ VERMELHO
                If Right(wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value, 1) = "º" And Not G = 5 Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + 2 + i).Select
                        With wXLA.Selection.Interior
                            .ColorIndex = 15
                            .Pattern = xlSolid
                            If "1º" = wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value Then
                                'muda a cor do 1º coloca e tambem a loja
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                            End If
                        End With
                    Next i
                End If
                      
                'Crediario
                'SE TIVER ALGUMA POSIÇÃO NA CLASSIFICAÇÃO - ETNÃO PINTA O FUNDO E MUDAR A COR DO 1º P/ VERMELHO
                If Right(wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value, 1) = "º" And Not G = 5 Then
                    For i = 0 To 1
                        wXl.Sheets(1).Cells(w_Row, w_Col + 5 + i).Select
                        With wXLA.Selection.Interior
                            .ColorIndex = 15
                            .Pattern = xlSolid
                            If "1º" = wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value Then
                                'muda a cor do 1º coloca e tambem a loja
                                wXLA.Selection.Font.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 2, 3)
                                wXLA.Selection.Interior.ColorIndex = IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 1, 1, 15)
                                wXLA.Selection.Interior.Pattern = xlSolid
                            End If
                        End With
                    Next i
                End If
                      
                      
                'Grupo 5  CC   e  MC
                'Muda a cor da Letra pra Vermelho
                If G = 5 Then
                    For i = 0 To 2
                        wXl.Sheets(1).Cells(w_Row, w_Col + 1 + i).Select
                        wXLA.Selection.Font.ColorIndex = 3
                    Next i
                End If
                      
                      
            w_Row = w_Row + 1
            
            wAdo.MoveNext
        Loop
              
        
        'Acabou o Grupo - Coloca o Total
        w_Col = 3

        'GUARDA O TOTAL DO GRUPO
        w_Tot_C_G = w_Tot_C_G + w_Tot_C
        w_Tot_b_G = w_Tot_b_G + w_Tot_B
        w_Tot_Ac = w_Tot_Ac + w_Tot_C_Ac
        w_Tot_d_G = w_Tot_d_G + w_Tot_D
        w_Tot_Ac_cred = w_Tot_Ac_cred + w_Tot_D_Ac
        
        
        If Not G = 5 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
            'INSERI O TOTAL DO GRUPO ATUAL  -  VENDA, BONUS E CREDIARIO
            wXl.Sheets(1).Cells(w_Row, w_Col).Value = Formata_Valor_Logo(w_Tot_C)
            wXl.Sheets(1).Cells(w_Row, w_Col).Select
            wXLA.Selection.Font.ColorIndex = 5
            
            wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value = CDbl(w_Tot_B)
            wXl.Sheets(1).Cells(w_Row, w_Col + 2).Select
            wXLA.Selection.Font.ColorIndex = 5
            
            wXl.Sheets(1).Cells(w_Row, w_Col + 2 + 3).Value = CDbl(w_Tot_D)
            wXl.Sheets(1).Cells(w_Row, w_Col + 2 + 3).Select
            wXLA.Selection.Font.ColorIndex = 5
            
            If frm_Rpt_Cod_Bon.ckAcum = 1 Then
                wXl.Sheets(1).Cells(w_Row, w_Col + 3).Value = Formata_Valor_Logo(w_Tot_C)
                wXl.Sheets(1).Cells(w_Row, w_Col + 3).Select
                wXLA.Selection.Font.ColorIndex = 5
                        
                wXl.Sheets(1).Cells(w_Row, w_Col + 4).Value = Formata_Valor_Logo(w_Tot_C_Ac)
                wXl.Sheets(1).Cells(w_Row, w_Col + 4).Select
                wXLA.Selection.Font.ColorIndex = 5
                
                wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value = Formata_Valor_Logo(w_Tot_B)
                wXl.Sheets(1).Cells(w_Row, w_Col + 5).Select
                wXLA.Selection.Font.ColorIndex = 5
                       
            End If
        End If
                   

        If G = 4 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then 'QUANDO COMPLETOU O GRUPO 4 - ENT
            w_Row = w_Row + 1
            
            wXl.Sheets(1).Cells(w_Row, w_Col).Value = Formata_Valor_Logo(w_Tot_C_G)
            wXl.Sheets(1).Cells(w_Row, w_Col).Select
            wXLA.Selection.Font.ColorIndex = 3
            
            'Bonus
            wXl.Sheets(1).Cells(w_Row, w_Col + 2).Value = CDbl(w_Tot_b_G)
            wXl.Sheets(1).Cells(w_Row, w_Col + 2).Select
            wXLA.Selection.Font.ColorIndex = 3
            
            wXl.Sheets(1).Cells(w_Row, w_Col + 2 + 3).Value = CDbl(w_Tot_d_G)
            wXl.Sheets(1).Cells(w_Row, w_Col + 2 + 3).Select
            wXLA.Selection.Font.ColorIndex = 3
            
            If frm_Rpt_Cod_Bon.ckAcum = 1 Then
                wXl.Sheets(1).Cells(w_Row, w_Col + 3).Value = Formata_Valor_Logo(w_Tot_C_G)
                wXl.Sheets(1).Cells(w_Row, w_Col + 3).Select
                wXLA.Selection.Font.ColorIndex = 3
                
                wXl.Sheets(1).Cells(w_Row, w_Col + 4).Value = Formata_Valor_Logo(w_Tot_Ac)
                wXl.Sheets(1).Cells(w_Row, w_Col + 4).Select
                wXLA.Selection.Font.ColorIndex = 3
                
                wXl.Sheets(1).Cells(w_Row, w_Col + 5).Value = Formata_Valor_Logo(w_Tot_B)
                wXl.Sheets(1).Cells(w_Row, w_Col + 5).Select
                wXLA.Selection.Font.ColorIndex = 3
            End If
            
            w_Row = w_Row - 1
        End If
            
        
        
        If G = 4 Then
            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(1, 1).Select
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            
            If frm_Rpt_Cod_Bon.ckAcum = 1 Then
                wXl.Sheets(1).Cells(w_Row, 8).Select
            Else
                wXl.Sheets(1).Cells(w_Row, 5).Select
            End If
            
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            Dim strCel_INI_FIM As String
            strCel_INI_FIM = Cel_Ini & ":" & Cel_Fim
            Inseri_Borda strCel_INI_FIM, wXLA, "G"
            
            w_Row = w_Row + 1
        End If
        
        'Inseri moldura só do total e grupo 5
        If G = 5 And frm_Rpt_Cod_Bon.ckLogo.Value = 0 Then
            'Pega os Endereços das Celulas
            wXl.Sheets(1).Cells(w_Row - 1, 1).Select
            Cel_Ini = wXLA.ActiveCell.AddressLocal
            
            If frm_Rpt_Cod_Bon.ckAcum = 1 Then
                wXl.Sheets(1).Cells(w_Row - 1, 8).Select
            Else
                wXl.Sheets(1).Cells(w_Row - 1, 5).Select
            End If
            
            Cel_Fim = wXLA.ActiveCell.AddressLocal
            strCel_INI_FIM = Cel_Ini & ":" & Cel_Fim
            Inseri_Borda strCel_INI_FIM, wXLA, "M"
            
        End If

        w_Row = w_Row + IIf(frm_Rpt_Cod_Bon.ckLogo.Value = 0, 2, 1)
        frm_Rpt_Cod_Bon.pgBar.Value = frm_Rpt_Cod_Bon.pgBar.Value + 5
    Next G
    
    
    wXLA.Visible = True
    
  Exit Sub
err1:
    MsgBox err.Description, vbCritical
End Sub




'No tipo   G - Grade   ou   M - Moldura
Public Sub Inseri_Borda(Celulas As String, ByRef wXLA, p_Tipo)

If p_Tipo = "G" Then
    
    wXLA.Range(Celulas).Select
    wXLA.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    wXLA.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With wXLA.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

Else 'Tipo Moldura
    wXLA.Range(Celulas).Select
    wXLA.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    wXLA.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With wXLA.Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With wXLA.Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    wXLA.Selection.Borders(xlInsideVertical).LineStyle = xlNone
    wXLA.Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

End If

End Sub

'Se estiver selecionado o Logo do codigo/ bonus  - então formata os valores
Public Function Formata_Valor_Logo(pp_valor, Optional Tipo As String = "C") As Variant
Dim p_valor     As Double

    
    If IsNull(pp_valor) Then pp_valor = 0
    p_valor = pp_valor
    If frm_Rpt_Cod_Bon.ckLogo.Value = 0 And frm_Rpt_Cod_Bon.ckPL.Value = 0 Then
        Formata_Valor_Logo = p_valor
    Else

        If Tipo = "C" Then p_valor = p_valor / 1000
        If Int(p_valor) > 0 Then
            w_Valor = Mid(p_valor, InStr(p_valor, ",") + 1)
            If p_valor = Int(p_valor) Then
                Formata_Valor_Logo = Int(p_valor) & "=00"
            Else
                If Len(Mid(p_valor, InStr(p_valor, ",") + 1, 2)) = 1 Then
                    Formata_Valor_Logo = Int(p_valor) & "=" & Mid(p_valor, InStr(p_valor, ",") + 1, 2) & "0"
                Else
                    Formata_Valor_Logo = Int(p_valor) & "=" & Mid(p_valor, InStr(p_valor, ",") + 1, 2)
                End If
            End If
        Else
            If Mid(p_valor, 3, 2) = "" Then
                Formata_Valor_Logo = "0=00"
            Else
                Formata_Valor_Logo = "0=" & Mid(p_valor, 3, 2)
                If Len(Mid(p_valor, 3, 2)) = 1 Then Formata_Valor_Logo = Formata_Valor_Logo & "0"
              
            End If
        End If
    End If

End Function

Public Sub Put_System_CNC()
Dim wbd, wftp

    'bd="0-host | 1-user | 2-pw | 3database"
    wbd = Split(GetIni("SYSTEM", "bd", App.Path & "\System_cnc.INI"), "|")
    
    'ftp = "0-host | 1-user | 2-pw | 3-dir"
    wftp = Split(GetIni("SYSTEM", "ftp", App.Path & "\System_cnc.INI"), "|")
    
    'Dados Banco Dados
    strBDHost = wbd(0)
    strBDUser = wbd(1)
    strBDPW = wbd(2)
    strBDDataBase = wbd(3)
    
    'Dados p/ Conexao FTP
    strFTPHost = wftp(0)
    strFTPUser = wftp(1)
    strFTPPassW = wftp(2)
    strFTPDir = wftp(3)
    
   
   strConectaMySQL = "DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & strBDHost & _
                      ";PORT=3306" & _
                      ";DATABASE=" & strBDDataBase & _
                      "; USER=" & strBDUser & _
                      ";PASSWORD=" & strBDPW & _
                      ";OPTION=3;"
                      
   ' strConectaMySQL = "Provider=MSDASQL.1;Persist Security Info=true;Data Source=odbc_cartao"
End Sub



Public Sub Baixar_FTP(w_File As String, ByRef pgBar)
Dim bRet As Boolean
Dim hOpen As Long, hConnection As Long
Dim p_arquivo, p_fileObj
Dim w_File_exe
    
    w_File_exe = w_File & ".exe"
    
    p_arquivo = App.Path & "\" & w_File & "_new.exe"
    Set p_fileObj = CreateObject("Scripting.FileSystemObject")
    
    If p_fileObj.FileExists(p_arquivo) = True Then
        Kill p_arquivo
    End If
        
    hOpen = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strFTPHost, 0, strFTPUser, strFTPPassW, 1, &H8000000, 0)
    bRet = FtpSetCurrentDirectory(hConnection, strFTPDir)
    bRet = FtpGetFile(hConnection, w_File_exe, p_arquivo, False, &H80000000, &H2, 0)
       
    If bRet = False Then
        'MsgBox err.LastDllError
    End If
       
    
    If p_fileObj.FileExists(p_arquivo) = True Then
            pgBar.Value = pgBar.MaxProgress

            Kill App.Path & "\" & w_File_exe  'Exclui o arquivo velho
            FileCopy p_arquivo, App.Path & "\" & w_File_exe 'Copia o novo com outro nome
            Kill p_arquivo 'Deleta o novo chamdo New
                    
            pgBar.Text = "Atualizado com sucesso!"
            Pause 0.5
    Else
            MsgBox "Não foi possível baixar a atualização!", vbCritical
    End If
    
    
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
        
End Sub






Public Sub Baixar_FTP2(ByRef Inet1, w_File As String, ByRef pgBar)
On Error Resume Next
    w_File = LCase(w_File)
    Kill App.Path & "\" & w_File & "_new.exe"


On Error GoTo ErroGeral

    Inet1.URL = strFTPHost
    Inet1.UserName = strFTPUser
    Inet1.Password = strFTPPassW
    pgBar.Value = 0
    pgBar.MaxProgress = 10000
    pgBar.Text = "Atualizando ................."
    Pause 1
    pgBar.Visible = True
    wTime = Time
    
    Inet1.Execute Inet1.URL, "GET " & strFTPDir & w_File & ".exe " & App.Path & "\" & w_File & "_new.exe"
    
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
        p_arquivo = App.Path & "\" & w_File & "_new.exe"
        Set p_fileObj = CreateObject("Scripting.FileSystemObject")
        If p_fileObj.FileExists(p_arquivo) = True Then
    
            pgBar.Value = pgBar.MaxProgress
                       
            Kill App.Path & "\" & w_File & ".exe" 'Exclui o arquivo velho
            FileCopy App.Path & "\" & w_File & "_new.exe", App.Path & "\" & w_File & ".exe" 'Copia o novo com outro nome
            Kill App.Path & "\" & w_File & "_new.exe" 'Deleta o novo chamdo New
                                        
            pgBar.Text = "Atualizado com sucesso!"
            Pause 0.5
        Else
            MsgBox "Não foi possível baixar a atualização!", vbCritical
        End If
       
    End If

ErroGeral:
    pgBar.Visible = False
    Pause 0.2
End Sub



Public Sub Baixar_FTP_MOV(ByRef Inet1, ByRef pgBar)
On Error Resume Next
    w_File = LCase(w_File)
    Kill App.Path & "\MODELO\M*.*"


On Error GoTo ErroGeral

    Inet1.URL = strFTPHost
    Inet1.UserName = strFTPUser
    Inet1.Password = strFTPPassW
    pgBar.Value = 0
    pgBar.MaxProgress = 10000
    pgBar.Text = "Atualizando ................."
    Pause 1
    pgBar.Visible = True
    
    
    'Cria Arquivo BAT p/ Criar os Diretorios
    Set FSO = CreateObject("Scripting.FileSystemObject")
    caminho = App.Path & "\CDIR.bat"   'especifique aqui o caminho onde ficará/está o TXT
    Set gravar = FSO.CreateTextFile(caminho, True)  'Arquivo Criado
    gravar.Write ("MD " & App.Path & "\MODELO" & vbCrLf)
    gravar.Write ("MD " & App.Path & "\CMOV" & vbCrLf)
    gravar.Write ("exit")
    gravar.Close
    
    'Cria o Diretorio MCAIXA - q/ Guarda os Movimentos diarios q/ serão enviados
    Shell App.Path & "\CDIR.bat", vbHide
    Pause 2
    Kill App.Path & "\CDIR.bat"
    
    
    Dim wDiretorio As String
    Dim SysInfo As SO
    Set SysInfo = New SO
    If InStr(SysInfo.OSPlatform, "XP") > 0 Then
        wDiretorio = "system32"
    Else
        wDiretorio = "system"
    End If
    
    
    For i = 1 To 4
        
        Select Case i
        Case 1: w_Arq = "modelo1.mdb"
        Case 2: w_Arq = "MSMAPI32.OCX"
        Case 3: w_Arq = "MSWINSCK.OCX"
        Case 4: w_Arq = "AspEmail.dll"
        End Select
        
        wTime = Time
        'Cria o Diretorio MODELO - q/ Guarda os Movimentos
        
        
        If i = 1 Then
            Inet1.Execute Inet1.URL, "GET " & strFTPDir & "/sisint/" & w_Arq & " " & App.Path & "\MODELO\" & w_Arq
        Else
            Inet1.Execute Inet1.URL, "GET " & strFTPDir & "/sisint/" & w_Arq & " " & "c:\windows\" & wDiretorio & "\" & w_Arq
        End If
        
        
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
    
    Next i
    
    
    'Cria Arquivo BAT p/ registrar Componentes e Dlls
    Set FSO = CreateObject("Scripting.FileSystemObject")
    caminho = App.Path & "\reg.bat"   'especifique aqui o caminho onde ficará/está o TXT
    Set gravar = FSO.CreateTextFile(caminho, True)  'Arquivo Criado
    gravar.Write ("c:" & vbCrLf)
    gravar.Write ("cd \windows\" & wDiretorio & vbCrLf)
   
    gravar.Write ("regsvr32 MSWINSCK.OCX" & vbCrLf)
    gravar.Write ("regsvr32 AspEmail.dll" & vbCrLf)
    gravar.Write ("regsvr32 MSMAPI32.OCX" & vbCrLf)
    
    gravar.Write ("exit")
    gravar.Close
    
    'Cria o Diretorio MCAIXA - q/ Guarda os Movimentos diarios q/ serão enviados
    Shell App.Path & "\reg.bat", vbHide
    Pause 2
    
    
    
    If Inet1.ResponseCode = 12003 Then
        MsgBox "Não foi possível fazer a atualização!" & Chr(13) & Chr(13) & "O arquivo a ser baixado FTP, não foi encontrado!", vbExclamation
    ElseIf Inet1.ResponseCode = 0 Then
    
        pgBar.Value = pgBar.MaxProgress
        p_arquivo = App.Path & "\MODELO\modelo1.mdb"
        Set p_fileObj = CreateObject("Scripting.FileSystemObject")
        If p_fileObj.FileExists(p_arquivo) = True Then
    
            pgBar.Value = pgBar.MaxProgress
            
            pgBar.Text = "Atualizado com sucesso!"
            Pause 0.5
        Else
            MsgBox "Não foi possível baixar a atualização!", vbCritical
        End If
       
    End If

ErroGeral:
    pgBar.Visible = False
    Pause 0.2
    Set p_fileObj = Nothing

End Sub


Public Sub Sendkeys(Text$, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys Text, wait
   Set WshShell = Nothing
End Sub

