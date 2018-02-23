VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDI 
   BackColor       =   &H8000000C&
   Caption         =   "Menu Principal"
   ClientHeight    =   6855
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8475
   Icon            =   "mdi.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   6600
      Width           =   8475
      _ExtentX        =   14949
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   3175
            MinWidth        =   3175
            Text            =   "Usuário : "
            TextSave        =   "Usuário : "
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   5847
            Object.Tag             =   ""
            Object.ToolTipText     =   "Status"
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   900
            MinWidth        =   882
            TextSave        =   "CAPS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            AutoSize        =   2
            Object.Width           =   873
            MinWidth        =   882
            TextSave        =   "NUM"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   635
            MinWidth        =   617
            TextSave        =   "INS"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            AutoSize        =   2
            Object.Width           =   1773
            MinWidth        =   1764
            TextSave        =   "20/06/2006"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "11:43"
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuSis 
      Caption         =   "Sistema"
      Begin VB.Menu mnuSisPar 
         Caption         =   "Parametrização"
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
         Caption         =   "Lançamentos"
      End
      Begin VB.Menu mnuBaixa 
         Caption         =   "Baixa Automática"
      End
      Begin VB.Menu mnuConLanc 
         Caption         =   "Consulta Lançamentos"
      End
      Begin VB.Menu mnuTot 
         Caption         =   "Total Diário"
      End
      Begin VB.Menu mnusep02 
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
Private Sub MDIForm_Load()
    StatusBar1.Panels(1) = "Usuário: " & w_Usu_Nome
    
    '*** Direitos de Acesso ***
    If w_Usu_Tipo = "L" Then    'Lojas
               
    ElseIf w_Usu_Tipo = "U" Then    'Usuarios
    
    ElseIf w_Usu_Tipo = "A" Then    'ADMINISTRADOR
    
    End If
    
    
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
    If vbYes = MsgBox("Deseja Sair do Sistema?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmação") Then
        End
    Else
        Cancel = 1
    End If
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

Private Sub mnuConLanc_Click()
    frm_Lancamento_Pesq.Show
End Sub

Private Sub mnuFPg_Click()
    frm_Forma_Pg.Show
End Sub

Private Sub mnuLanc_Click()
    frm_Lancamento.Show
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
