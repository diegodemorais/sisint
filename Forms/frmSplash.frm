VERSION 5.00
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.Ocx"
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4425
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   2160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      Protocol        =   2
      RemotePort      =   21
      URL             =   "ftp://"
   End
   Begin VB.Timer timInicio 
      Interval        =   1000
      Left            =   6840
      Top             =   120
   End
   Begin VBXPolitec.ocxProgressBarTexto PB 
      Height          =   360
      Left            =   5160
      TabIndex        =   3
      Top             =   3795
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorFundo  =   16576
      BackColorFundo  =   -2147483643
      BackColorProgress=   16576
      MaxProgress     =   50
   End
   Begin VBXPolitec.ocxProgressBarTexto pgBarAT 
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   635
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColorFundo  =   16576
      BackColorFundo  =   -2147483643
      BackColorProgress=   16576
      MaxProgress     =   50
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão : 1.0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   285
      Left            =   5145
      TabIndex        =   2
      Top             =   3450
      Width           =   2190
   End
   Begin VB.Label lblWarning 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   3960
      Visible         =   0   'False
      Width           =   4815
   End
   Begin VB.Label lblProductName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Sistema Integrado "
      BeginProperty Font 
         Name            =   "Monotype Corsiva"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   855
      Left            =   1650
      TabIndex        =   0
      Top             =   1005
      Width           =   5835
   End
   Begin VB.Image Image1 
      Height          =   4515
      Left            =   -45
      Picture         =   "frmSplash.frx":548A
      Stretch         =   -1  'True
      Top             =   -45
      Width           =   7560
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_time As Byte
Dim w_Ver_Loja
Dim w_FazerAT As Boolean

Option Explicit

Private Sub Form_Load()
Dim w_Arr
Dim sParametro As String
    controle = False

  timeCNC = "00:00:00"
  sParametro = Command$
  If sParametro = "" Then GoTo errUser

    w_Caminho_SCL = Left(App.Path, 3) & "SCL\LOJA\"
    
    Call Put_System_CNC
    W_VER = App.Major & "." & App.Minor & "." & App.Revision
    
    lblVersion.Caption = "Versão: " & W_VER
    
    Call AbreConexao(Conexão, "")
    w_Data_Server = CVDate(Format(ExecuteSQL("SELECT NOW() AS DT FROM tab_banco GROUP BY NOW()", , , False).Fields(0), "dd/mm/yyyy"))
    
    w_Arr = ExecuteSQL("SELECT * FROM tab_usuario WHERE (usl_cod = " & sParametro & ")", , , False).GetRows
    If IsArray(w_Arr) Then

        On Error Resume Next
        w_Ver_Loja = w_Arr(9, 0)
        w_Usu_Cod = UCase(w_Arr(0, 0))
        w_Usu_Nome = UCase(w_Arr(1, 0))
        w_Usu_Pass = UCase(w_Arr(2, 0))
        w_Usu_Tipo = UCase(w_Arr(3, 0))
        w_Usu_Ac = UCase(w_Arr(6, 0))
        w_Usu_Rpt = UCase(w_Arr(7, 0))
        
        'Se for usuario , então deleta historico dos Cod/Bonus de mais de 3 dias
        If w_Usu_Tipo = "U" Then
            Call ExecuteSQL("DELETE FROM tab_vnds_bonus WHERE (dt_vnd < '" & Format(CVDate(w_Data_Server - 5), "yyyy-mm-dd") & "')", , , False)
        End If
    End If
    
    Dim SysInfo As SO
    Set SysInfo = New SO

    Dim wVer As String
    wVer = ExecuteSQL("Select USL_VERSAO FROM tab_usuario WHERE USL_COD = " & w_Usu_Cod).Fields(0)
    
    ExecuteSQL "UPDATE tab_usuario SET USL_VERSAO = '" & W_VER & "', USL_WINDOWS = '" & SysInfo.OSPlatform & "' WHERE USL_COD = " & w_Usu_Cod
    
    Set SysInfo = Nothing
    
    If wVer < W_VER Then
        w_FazerAT = True
    End If
    
    frmSplash.lblWarning = frmSplash.lblWarning & " ******** BASE OFICIAL ! MySQL *********"
       
    timInicio.Enabled = True 'ligar atualização

sair:
    Exit Sub
errUser:
    MsgBox "Usuário Inválido!", vbCritical
    End
End Sub


Private Sub timInicio_Timer()
Dim w_Ver_Config
On Error GoTo err1
 
    'instanciação / atualização
    timInicio.Enabled = False

    w_Ver_Config = ExecuteSQL("SELECT conf_versao_login FROM tab_config", , , False).Fields(0)
    
    'Verificando a Versão do prlogin  p/ atualizar
    If w_Ver_Loja < w_Ver_Config Then
        'Call Baixar_FTP("prlogin", pgBarAT)
    End If

    AtualizarGeral 'abrir as tabelas e instanciar objetos
    'de.cncCartao.open strConectaMySQL
    
sair:
    frmSplash.Hide
    Liberar_Menu
    MDI.Show
    
Dim W_VER As String
    W_VER = App.Major & "." & App.Minor & "." & App.Revision
    
    MDI.StatusBar1.Panels(4).text = "Ver. " & W_VER
    MDI.timer_at.Enabled = w_FazerAT
    
    Unload frmSplash
    
    Exit Sub
err1:
'If CDbl(err.Number) = CDbl(424) Then
        'MsgBox "Favor fazer a importação de tabelas!", vbCritical
    'Else
         'MsgBox Error$, vbCritical
    'End If
    Resume sair
    
End Sub

Sub Liberar_Menu()
    If w_Usu_Tipo = "U" Then
        MDI.mnuBaixa.Visible = False
        'MDI.TBar.Buttons("baixa").Visible = False
    ElseIf w_Usu_Tipo = "L" Then
        MDI.mnuBaixa.Visible = False
        MDI.mnubco.Visible = False
        MDI.mnuCartLoja.Visible = False
        MDI.mnuFPg.Visible = False
        MDI.mnuTpC.Visible = False
        MDI.mnuRpt.Visible = (w_Usu_Rpt = "S")
        MDI.TBar.Buttons("baixa").Visible = False
        MDI.mnuConf.Visible = False
    ElseIf w_Usu_Tipo = "S" Then
        MDI.mnuBaixa.Visible = False
        MDI.mnubco.Visible = False
        MDI.mnuCartLoja.Visible = False
        MDI.mnuFPg.Visible = False
        MDI.mnuTpC.Visible = False
        MDI.mnuRpt.Visible = (w_Usu_Rpt = "S")
        MDI.mnuLanc.Visible = False
        MDI.mnuTot.Visible = False
        MDI.mnuConLanc.Visible = False
        MDI.mnuCodBon.Visible = False
        MDI.mnuSep01.Visible = False
        MDI.mnuConf.Visible = False
        
        MDI.TBar.Buttons("lançar").Visible = False
        MDI.TBar.Buttons("con_lanc").Visible = False
        MDI.TBar.Buttons("resumo").Visible = False
        MDI.TBar.Buttons("cod_bon").Visible = False
        MDI.TBar.Buttons("baixa").Visible = False
    End If
   
    If w_Usu_Cod = 40 Or w_Usu_Cod = 38 Then MDI.TBar.Buttons.Item("rpt_resumo").Visible = True
   
End Sub


