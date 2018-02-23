VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1515
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   895.112
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSAdodcLib.Adodc w_Ado 
      Height          =   330
      Left            =   1920
      Top             =   600
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
      Left            =   495
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
      Caption         =   "&Password:"
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
Dim W As Variant
Dim w_Arr
    
On Error GoTo err1

    If de.cncCartao.State = 0 Then de.cncCartao.Open
    On Error Resume Next
    w_Arr = de.cncCartao.Execute("SELECT * FROM tab_usuario WHERE (usl_nome = '" & txtUserName & "') AND (usl_pwd = '" & txtPassword & "')").GetRows
   
    w_Data_Server = CVDate(Format(de.cncCartao.Execute("SELECT { fn NOW() } AS DT FROM tab_banco GROUP BY { fn NOW() }").Fields(0), "dd/mm/yyyy"))
   
   'Check de Senha e Nome é do Comum
    If IsArray(w_Arr) Then
        
        On Error Resume Next
        w_Usu_Nome = UCase(w_Arr(1, 0))
        w_Usu_Cod = UCase(w_Arr(0, 0))
        w_Usu_Pass = UCase(w_Arr(2, 0))
        w_Usu_Tipo = UCase(w_Arr(3, 0))
             
       'Fecha o Form
         Hide 'Unload Me
       'Abre a tela de splash
         frmSplash.Show
     
   'Se não for nenhuma das comparações
    Else
        MsgBox "Senha ou Nome Inválido, Tente Novamente!", vbCritical, "Login"
        txtPassword.SetFocus
        SendKeys "{Home}+{End}"
    End If

sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub cmdCancel_Click()
    End
End Sub






Private Sub Form_Activate()
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    'Pega os Dir. do Arq. INI
    strLoja = GetIni("SYSTEM", "Loja", App.Path & "\System.INI")

    txtUserName = strLoja

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
