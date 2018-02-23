Attribute VB_Name = "Unit_Login"
Public strLoja As String  'define qual a loja instalada
Public W_VER As String    'Versao do EXE
Public w_Data_Server As Date 'Data do servidor  utilizada p/ ninguem fraudar o sistema mudando a data do windows

Public w_Usu_Nome As String
Public w_Usu_Cod As String
Public w_Usu_Pass As String
Public w_Usu_Tipo As String
Public w_Usu_Ac As String
Public w_Usu_Rpt As String


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

Public strConectaMySQL As String, strConectaMySQL2 As String


Declare Function GetPrivateProfileString Lib "Kernel32" Alias _
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
  hWnd As Long
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

  FLOP.hWnd = 0
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
    MsgBox Err.LastDllError, vbCritical Or vbOKOnly
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


Public Function ExecuteSQL(SQLString, Optional ByRef w_RegAf, Optional w_Provider As String) As ADODB.Recordset
    Dim Conexão As ADODB.Connection
    Dim i As Byte, w_Err As Byte

w_RegAf = 0
w_Err = 0
conectar:
On Error GoTo errCNC

    '1º - Abrindo Conexão
    Set Conexão = New ADODB.Connection


    'Conexão.Provider = "MSDASQL.1"
    If Not w_Provider = "" Then Conexão.Provider = w_Provider

    Conexão.CursorLocation = adUseClient
    Conexão.Open strConectaMySQL
   

    '2º Executando SQL
    
    If InStr(UCase(SQLString), "SELECT") > 0 Or InStr(UCase(SQLString), "CALL") Then
        Set ExecuteSQL = Conexão.Execute(SQLString, w_RegAf).Clone
        Set ExecuteSQL.ActiveConnection = Nothing
    Else
        Conexão.Execute SQLString, w_RegAf
    End If



sair:
    On Error Resume Next
    Conexão.Close
    Set Conexão = Nothing
    Exit Function
    
errCNC:
    strConectaMySQL = strConectaMySQL2
    w_Err = w_Err + 1
    Pause 0.3
    If w_Err <= 5 Then
        Resume conectar
    Else
        strConectaMySQL = strConectaMySQL2
    End If
    MsgBox "Erro no ExecutarSQL " & Chr(13) & Chr(13) & "Err: " & Error$, vbCritical, "ExecuteSQL"
    Resume sair
End Function

Public Sub AbreConexao(ByRef Conexão)
On Error Resume Next
    Set Conexão = New ADODB.Connection
    Conexão.CursorLocation = adUseClient
    Conexão.ConnectionString = strConectaMySQL
    Conexão.Open
End Sub

Public Sub FechaConexao(ByRef Conexão)
 On Error Resume Next
    Conexão.Close
    Set Conexão = Nothing
End Sub
   

Public Sub KeyEnter(key)
    'Utilizar assim : KeyEnter (KeyAscii)
    Select Case key
    Case 13 'enter
        SendKeys "{tab}{home}+{end}"
    Case 38 'seta para baixo
        SendKeys "+{tab}{home}+{end}"
    Case 40 'seta para cima
        SendKeys "{tab}{home}+{end}"
    End Select
End Sub


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



Public Function msgErro(ByRef wErr) As String
Dim i, w_count

On Error GoTo err1

    If wErr = 0 Or wErr = 91 Or wErr = -2147217871 Or wErr = -2147467259 Then
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
    MsgBox Err.Number & " : " & Err.Description, vbCritical
    Resume sair
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
    
    strConectaMySQL = "driver={MySQL ODBC 3.51 Driver}; server=" & strBDHost & _
                      "; uid=" & strBDUser & _
                      ";Pwd=" & strBDPW & _
                      ";database=" & strBDDataBase


    strConectaMySQL2 = "Provider=MSDASQL.1;Persist Security Info=False; " & _
                      "User ID=rpaps_2;Data Source=odbc_cartao"


End Sub
