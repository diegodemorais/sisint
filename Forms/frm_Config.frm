VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "ACTIVETEXT.OCX"
Begin VB.Form frm_Config 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Config"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4455
   Icon            =   "frm_Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   1535
      ButtonWidth     =   1667
      ButtonHeight    =   1429
      ImageList       =   "IMG"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   4
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Fecha&r"
            Key             =   "fechar"
            Object.ToolTipText     =   "Fechar Tela"
            Object.Tag             =   ""
            ImageIndex      =   1
            Object.Width           =   1e-4
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Salvar"
            Key             =   "salvar"
            Object.ToolTipText     =   "Salvar Registro"
            Object.Tag             =   ""
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Cancelar"
            Key             =   "cancelar"
            Object.ToolTipText     =   "Cancelar Alteração"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
      EndProperty
      Begin VB.TextBox Text1 
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
         Left            =   3075
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "Config"
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.TextBox txt_versao 
      Alignment       =   2  'Center
      DataField       =   "conf_versao"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1140
   End
   Begin MSAdodcLib.Adodc adoReg 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      Top             =   2955
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
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
      Caption         =   "Registro(s): 0 / 0"
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
   Begin rdActiveText.ActiveText txt_versao_login 
      DataField       =   "conf_versao_login"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   2400
      TabIndex        =   3
      Top             =   2040
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   556
      Alignment       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TextCase        =   1
      RawText         =   0
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão  EXE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   885
      TabIndex        =   5
      Top             =   1485
      Width           =   1380
   End
   Begin VB.Label lblFieldLabel 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Versão Login:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   4
      Top             =   2130
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   480
      Top             =   960
      Width           =   3495
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   8760
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":2ABC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":2C96
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":2FB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":32CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":35E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":38FE
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Config.frx":3AD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFechar 
      Caption         =   "Fecha&r"
   End
   Begin VB.Menu mnuSep01 
      Caption         =   "        |"
      Enabled         =   0   'False
   End
   Begin VB.Menu mnuSalvar 
      Caption         =   "&Salvar"
   End
   Begin VB.Menu mnuCancelar 
      Caption         =   "&Cancelar"
   End
   Begin VB.Menu mnuSep03 
      Caption         =   "|"
      Enabled         =   0   'False
   End
End
Attribute VB_Name = "frm_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub adoReg_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
On Error GoTo err1

    adoReg.Caption = "Registro(s) : " & adoReg.Recordset.AbsolutePosition & " / " & adoReg.Recordset.RecordCount

sair:
    Exit Sub
err1:
    If Not err.Number = -2147217885 Then MsgBox err.Number & " : " & err.Description, vbCritical
    Resume sair
End Sub


Private Sub Form_Load()
On Error GoTo err1

    MDI.TBar.Visible = False

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 100
    
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_config").Clone


sair:
    Exit Sub
err1:
    MsgBox msgErro(err), vbCritical
    Resume sair
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub

Sub Salvar()
On Error GoTo err1
    If Not txt_versao.DataSource Is Nothing Then
         
         strSQL = "UPDATE tab_config SET " & _
                  "conf_versao = '" & txt_versao & "', conf_versao_login = '" & txt_versao_login & "'"
                    
    Else
         
         strSQL_Fields = "conf_versao, conf_versao_login"
         strSql_Values = "'" & txt_versao & "','" & txt_versao_login & "'"
         strSQL = "INSERT INTO tab_banco (" & strSQL_Fields & ") VALUES (" & strSql_Values & ")"
    
    End If
    
    
    ExecuteSQL strSQL, , , False
    Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_config").Clone
    
    Set txt_versao.DataSource = adoReg
    Set txt_versao_login.DataSource = adoReg
    
    MsgBox "Registro Salvo com sucesso!", vbInformation
    
sair:
    Exit Sub
err1:
    If Not err.Number = 0 Then MsgBox msgErro(err), vbCritical
    Resume sair
End Sub


Private Sub mnuCancelar_Click()
On Error GoTo err1

          Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_config").Clone

          If Not IsNull(adoReg.Recordset.Fields(1)) Then

                w_Pos = adoReg.Recordset.AbsolutePosition - IIf(adoReg.Recordset.Fields(1) = "", 1, 0)
                adoReg.Recordset.MoveFirst
                adoReg.Recordset.Move w_Pos - 1
          
          Else
                Set adoReg.Recordset = ExecuteSQL("SELECT * FROM tab_config").Clone
          End If
sair:
    Exit Sub
err1:

    MsgBox msgErro(err), vbCritical
    If Not (err.Number = -2147217871 Or err.Number = 91 Or err.Number = 0) Then If adoReg.Recordset.RecordCount > 0 Then adoReg.Recordset.MoveFirst
    Resume sair
End Sub


Private Sub mnuFechar_Click()
    Unload Me
End Sub


Private Sub mnuSalvar_Click()
    Salvar
End Sub

Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFechar_Click
    Case "salvar": Salvar
    Case "cancelar": mnuCancelar_Click
    End Select
End Sub




Private Sub txt_versao_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyEnter KeyCode
End Sub


Private Sub txt_versao_login_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If vbYes = MsgBox("Deseja Salvar?", vbQuestion + vbYesNo + vbDefaultButton1, "Estoque") Then
            Salvar
            txt_versao.SetFocus
        End If
    Else
        KeyEnter KeyCode
    End If
End Sub
