VERSION 5.00
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.5#0"; "COMCTL32.OCX"
Object = "{4E6B00F6-69BE-11D2-885A-A1A33992992C}#2.6#0"; "activetext.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Rpt_Cod_Bon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Relat�rio de C�digo / B�nus"
   ClientHeight    =   2951
   ClientLeft      =   156
   ClientTop       =   546
   ClientWidth     =   6604
   Icon            =   "frm_Rpt_Cod_Bon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2951
   ScaleWidth      =   6604
   ShowInTaskbar   =   0   'False
   Begin ComctlLib.Toolbar TBar 
      Align           =   1  'Align Top
      Height          =   819
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   6604
      _ExtentX        =   11646
      _ExtentY        =   1438
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
         Height          =   675
         Left            =   750
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "Relat�rio de C�digo / B�nus"
         Top             =   60
         Width           =   5775
      End
   End
   Begin VB.CheckBox ckPL 
      Caption         =   "PL?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3555
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2205
      Width           =   735
   End
   Begin VBXPolitec.ocxProgressBarTexto pgBar 
      Height          =   300
      Left            =   1185
      TabIndex        =   10
      Top             =   2610
      Visible         =   0   'False
      Width           =   4335
      _ExtentX        =   7644
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
      Text            =   "................................ Gerando Relat�rio ..............................."
      Text            =   "................................ Gerando Relat�rio ..............................."
      BackColorFundo  =   -2147483643
      MaxProgress     =   100
   End
   Begin VB.CheckBox ckSup 
      Caption         =   "Super?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2385
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2205
      Width           =   975
   End
   Begin VB.CheckBox ckAcum 
      Caption         =   "Acum?"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4440
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.CheckBox ckLogo 
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1455
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2205
      Width           =   840
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
      Left            =   1245
      TabIndex        =   0
      Top             =   1515
      Width           =   1035
      _ExtentX        =   1821
      _ExtentY        =   551
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
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
      FontSize        =   7.472
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
      Left            =   2610
      TabIndex        =   1
      Top             =   1515
      Width           =   1035
      _ExtentX        =   1821
      _ExtentY        =   551
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.4717
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
      FontSize        =   7.472
   End
   Begin Skin_Button.ctr_Button bt_Pesq 
      Height          =   845
      Left            =   4446
      TabIndex        =   4
      Top             =   1157
      Width           =   1040
      _ExtentX        =   1845
      _ExtentY        =   1486
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
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
      MICON           =   "frm_Rpt_Cod_Bon.frx":27A2
      PICN            =   "frm_Rpt_Cod_Bon.frx":27BE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button bt_RPT_OBS 
      Height          =   840
      Left            =   3675
      TabIndex        =   11
      Top             =   1155
      Width           =   675
      _ExtentX        =   1198
      _ExtentY        =   1486
      BTYPE           =   2
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.47
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
      MICON           =   "frm_Rpt_Cod_Bon.frx":2D01
      PICN            =   "frm_Rpt_Cod_Bon.frx":2D1D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Shape Shape2 
      Height          =   450
      Left            =   1185
      Top             =   2085
      Width           =   4335
   End
   Begin ComctlLib.ImageList IMG 
      Left            =   8760
      Top             =   240
      _ExtentX        =   1006
      _ExtentY        =   1006
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   8
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":3037
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":3351
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":352B
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":3845
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":3B5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":3E79
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":4193
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frm_Rpt_Cod_Bon.frx":436D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Shape Shape1 
      Height          =   900
      Left            =   1185
      Top             =   1125
      Width           =   4335
   End
   Begin VB.Label lb_Dt 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Escopo de Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1260
      TabIndex        =   3
      Top             =   1275
      Width           =   2385
   End
   Begin VB.Label lb_Dt2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "�"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.47
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   2190
      TabIndex        =   2
      Top             =   1560
      Width           =   480
   End
   Begin VB.Menu mnuFec 
      Caption         =   "Fecha&r"
   End
End
Attribute VB_Name = "frm_Rpt_Cod_Bon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim w_RPT As Boolean

Private Sub bt_Pesq_Click()
On Error Resume Next
    pgBar.Value = 0
    pgBar.Visible = True
    
    Call ExecuteSQL("DROP TABLE `" & strBDDataBase & "`.`tab_tmp_" & w_Usu_Cod & "`")
    
    'Cria a Tabela tab_tmp  baseado na data do escopo
    Monta_SQL_Tab_Tmp Txt_DtI, Txt_DtF   '60% do tempo
    CRIA_RPT_EXCEL Txt_DtI, Txt_DtF     '40% do Tempo
    
    pgBar.Value = 94
    Call ExecuteSQL("DROP TABLE `" & strBDDataBase & "`.`tab_tmp_" & w_Usu_Cod & "`")
    
    pgBar.Value = 100
    pgBar.Visible = False
End Sub

Private Sub bt_RPT_OBS_Click()
Dim w_SQL As String
Dim w_Rec As New Recordset

On Error Resume Next

    w_SQL = "SELECT tab_vnds_bonus.dt_vnd as Data, tab_usuario.usl_nome as Logo, tab_vnds_bonus.obs as OBS FROM tab_vnds_bonus, tab_usuario WHERE (tab_vnds_bonus.usl_cod = tab_usuario.usl_cod) AND (tab_vnds_bonus.dt_vnd >= '" & Format(Txt_DtI, "yyyy-mm-dd") & "' AND tab_vnds_bonus.dt_vnd <= '" & Format(Txt_DtF, "yyyy-mm-dd") & "') AND (NOT (tab_vnds_bonus.obs IS NULL)) AND (tab_vnds_bonus.obs <> '') ORDER BY tab_vnds_bonus.dt_vnd, tab_usuario.usl_nome"

    Set w_Rec = ExecuteSQL(w_SQL, w_RegAf, "MSDataShape").Clone
   
    If w_RegAf > 0 Then
        Set Rel_Obs_CodBon.DataSource = w_Rec.Clone
        Rel_Obs_CodBon.WindowState = vbMaximized
        Rel_Obs_CodBon.Show
        w_RPT = True
    Else
        MsgBox "Nenhum Logo com Observa��o foi encontado!", vbExclamation
    End If
End Sub



Private Sub ckLogo_Click()
    If ckLogo.Value = 1 Then ckPL.Value = 0
    If ckSup.Value = 1 And ckLogo.Value = 0 Then ckSup.Value = 0
End Sub

Private Sub ckPL_Click()
    If ckPL.Value = 1 Then
        ckLogo.Value = 0
        ckSup.Value = 0
    End If
    If ckAcum.Visible = True Then ckAcum.Value = ckPL.Value
    
End Sub

Private Sub ckSup_Click()
    ckLogo.Value = ckSup.Value
End Sub

Private Sub Form_Load()
    MDI.TBar.Visible = False
    
    If w_Usu_Tipo = "L" Or w_Usu_Tipo = "S" Then Height = 3465
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
    
    If w_Usu_Tipo = "L" Then  'Lojas
    
        ckLogo.Visible = False
        ckLogo.Value = 1
        Txt_DtI.Enabled = False
        Txt_DtF.Enabled = False
        ckAcum.Visible = False
        ckSup.Visible = False
        bt_RPT_OBS.Visible = False
        Shape2.Visible = False
        ckPL.Visible = False
        
    ElseIf w_Usu_Tipo = "S" Then    'Supervisor
        
        ckLogo.Visible = False
        ckLogo.Value = 1
        Txt_DtI.Enabled = False
        Txt_DtF.Enabled = False
        ckAcum.Visible = False
        ckSup.Visible = False
        bt_RPT_OBS.Visible = False
        ckSup.Value = 1
        ckPL.Visible = False
        
        Shape2.Visible = False
        
    End If
    
    If w_Usu_Tipo = "A" Then ckAcum.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDI.TBar.Visible = True
End Sub

Private Sub mnuFec_Click()
    If w_RPT = True Then
        Unload Rel_Obs_CodBon
        w_RPT = False
    Else
        Unload Me
    End If
End Sub


Private Sub TBar_ButtonClick(ByVal Button As ComctlLib.Button)
    Select Case Button.key
    Case "fechar": mnuFec_Click
    End Select
End Sub


Private Sub Txt_DtI_Validate(Cancel As Boolean)
On Error Resume Next
    Txt_DtF = CVDate(Txt_DtI)
End Sub
