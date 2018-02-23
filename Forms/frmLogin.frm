VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Integrado - Login"
   ClientHeight    =   1830
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1081.224
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VBXPolitec.ocxProgressBarTexto pgBar 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   3735
      _ExtentX        =   6588
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
      Text            =   " Atualizando Versao................."
      Text            =   " Atualizando Versao................."
      BackColorFundo  =   -2147483643
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   1680
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   4
      RemoteHost      =   "www.rpaps.locaweb.com.br"
      URL             =   "http://www.rpaps.locaweb.com.br/sisint.exe"
      Document        =   "/sisint.exe"
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Text            =   "Loja"
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   600
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "Usuário:"
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      Caption         =   "&Senha:"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'*** Botões ***
Private Sub cmdOK_Click()
Dim W As Variant, W_VER As String
Dim w_Arr, w_Usu
Dim w_Ver_Loja, w_Ver_Config
Dim w_adoUsu As ADODB.Recordset
Dim w_RegAf
On Error GoTo err1
    On Error Resume Next
    Set w_adoUsu = ExecuteSQL("SELECT usl_versao, usl_cod FROM tab_usuario WHERE (usl_nome = '" & txtUserName & "' ) AND (usl_pwd = '" & txtPassword & "')", w_RegAf).Clone
    
    w_Ver_Loja = w_adoUsu.Fields(0)
   'Check de Senha e Nome
    If w_RegAf > 0 Then

       W_VER = App.Major & "." & App.Minor & "." & App.Revision

       w_Ver_Config = ExecuteSQL("SELECT conf_versao FROM tab_config").Fields(0)

       'If w_Ver_Loja < w_Ver_Config Then

       '     pgBar.Value = 0
       '     pgBar.Visible = True
       '     Pause 0.2
       '     pgBar.MaxProgress = 10000

            
       '     Call Baixar_FTP
            

       'End If
       'ExecuteSQL "UPDATE tab_usuario SET usl_versao_login = '" & W_VER & "' WHERE usl_cod = " & w_adoUsu.Fields(1)
        
        'Abre a tela de splash
        Shell App.Path & "\SisInt.exe " & w_adoUsu.Fields(1), vbHide
        'Fecha o Form
        Unload frmLogin
        
   'Se não for nenhuma das comparações
    Else
        MsgBox "Senha ou Nome Inválido, Tente Novamente!", vbCritical, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If

sair:
    Exit Sub
err1:
    MsgBox msgErro(Err), vbCritical
    Resume sair
End Sub

Private Sub cmdCancel_Click()
    End
End Sub






Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
On Error GoTo err1

    'Pega os Dir. do Arq. INI
    strLoja = GetIni("SYSTEM", "Loja", App.Path & "\System.INI")
    txtUserName = strLoja
    Call Put_System_CNC
    
sair:
    Exit Sub
err1:
    MsgBox Err.Number & " : " & Err.Description, vbCritical
End Sub

Private Sub txtPassword_GotFocus()
    SendKeys "{home}+{end}"
End Sub

'*** Keydowns das txt
Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter (KeyCode)
End Sub

Private Sub txtUserName_GotFocus()
    SendKeys "{home}+{end}"
End Sub

Private Sub txtUserName_KeyDown(KeyCode As Integer, Shift As Integer)
KeyEnter (KeyCode)
End Sub

Private Sub Baixar_FTP()
Dim bRet As Boolean
Dim hOpen As Long, hConnection As Long
Dim p_arquivo, p_fileObj

    p_arquivo = App.Path & "\sisint_new.exe"
    Set p_fileObj = CreateObject("Scripting.FileSystemObject")
    
    If p_fileObj.FileExists(p_arquivo) = True Then
        Kill p_arquivo
    End If
        
    hOpen = InternetOpen("vb wininet", 1, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strFTPHost, 0, strFTPUser, strFTPPassW, 1, &H8000000, 0)
    bRet = FtpSetCurrentDirectory(hConnection, strFTPDir)
    bRet = FtpGetFile(hConnection, "sisint.exe", App.Path & "\sisint_new.exe", False, &H80000000, &H2, 0)
       
    If bRet = False Then
        'MsgBox Err.LastDllError
    End If
       
    If p_fileObj.FileExists(p_arquivo) = True Then
            pgBar.Value = pgBar.MaxProgress

            Kill App.Path & "\sisint.exe"  'Exclui o arquivo velho
            FileCopy p_arquivo, App.Path & "\sisint.exe" 'Copia o novo com outro nome
            Kill p_arquivo 'Deleta o novo chamdo New
                    
            pgBar.Text = "Atualizado com sucesso!"
            Pause 0.5
    Else
            Call Baixar_FTP2
            'MsgBox "Não foi possível baixar a atualização!", vbCritical
    End If
    
    
    If hConnection <> 0 Then InternetCloseHandle hConnection
    hConnection = 0
        
End Sub


Private Sub Baixar_FTP2()
Dim p_arquivo, p_fileObj, wTime
On Error Resume Next
  
    Kill App.Path & "\sisint_new.exe"
    
'On Error GoTo ErroGeral
    
    Inet1.Protocol = icFTP 'icHTTP
    Inet1.AccessType = icUseDefault
    Inet1.URL = strFTPHost
    Inet1.UserName = strFTPUser
    Inet1.Password = strFTPPassW

    wTime = Time()
    Inet1.Execute Inet1.URL, "GET " & strFTPDir & "sisint.exe " & App.Path & "\sisint_new.exe"

    Do
        DoEvents
        pgBar.Value = pgBar.Value + 2
        If pgBar.Value >= 10000 Then pgBar.Value = 0
        If CVDate(Time - wTime) >= CVDate("00:03:00") Then
        
            If vbYes = MsgBox("Sua Conexão está lenta, deseja continuar fazendo atualização?", vbQuestion + vbYesNo) Then
                wTime = Time()
            Else
                Exit Do
            End If
            
        End If
    Loop Until Not Inet1.StillExecuting


    If Inet1.ResponseCode = 12003 Then
        'MsgBox "Não foi possível fazer a atualização!" & Chr(13) & Chr(13) & "O arquivo a ser baixado FTP, não foi encontrado!", vbExclamation
    ElseIf Inet1.ResponseCode = 0 And InStr(Inet1.ResponseInfo, "concluída") > 0 Then

        p_arquivo = App.Path & "\sisint_new.exe"
   
        Set p_fileObj = CreateObject("Scripting.FileSystemObject")
           
        If p_fileObj.FileExists(p_arquivo) = True Then
    
            pgBar.Value = pgBar.MaxProgress

            Kill App.Path & "\sisint.exe"  'Exclui o arquivo velho
            FileCopy App.Path & "\sisint_new.exe", App.Path & "\sisint.exe" 'Copia o novo com outro nome
            Kill App.Path & "\sisint_new.exe" 'Deleta o novo chamdo New
                    
            pgBar.Text = "Atualizado com sucesso!"
            Pause 0.5
        Else
            'MsgBox "Não foi possível baixar a atualização!", vbCritical
        End If
        
    Else
        'MsgBox "ERRO : " & Inet1.ResponseInfo & Chr(13) & "   URL: " & Inet1.URL, vbCritical
    End If


Exit Sub
'ErroGeral:
    'MsgBox "Erro : " & Inet1.ResponseCode & " :  " & Inet1.ResponseInfo & Chr(13) & "   URL: " & Inet1.URL, vbCritical
End Sub
