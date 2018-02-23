VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{9A4D18F7-4EC7-11D5-9E33-0040C78773FC}#1.0#0"; "VBxPOLITEC.ocx"
Object = "{83E7A33D-84B8-4C96-9A60-2290FFC1A9A1}#2.0#0"; "Skin_Button.ocx"
Begin VB.Form frm_Copy_CartLoja 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copia taxa de uma loja p/ outras Lojas"
   ClientHeight    =   7275
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4530
   Icon            =   "frm_Copy_CartLoja.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7275
   ScaleWidth      =   4530
   Begin VBXPolitec.ocxProgressBarTexto pb 
      Height          =   390
      Left            =   0
      TabIndex        =   11
      Top             =   6825
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   688
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
      Text            =   "Copiando taxa p/ Loja :"
      Text            =   "Copiando taxa p/ Loja :"
      BackColorFundo  =   -2147483643
      MaxProgress     =   3
   End
   Begin Skin_Button.ctr_Button bt_Copiar 
      Height          =   735
      Left            =   2550
      TabIndex        =   10
      Top             =   5895
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   1296
      BTYPE           =   2
      TX              =   "Copiar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
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
      MICON           =   "frm_Copy_CartLoja.frx":0442
      PICN            =   "frm_Copy_CartLoja.frx":045E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSAdodcLib.Adodc adoCartao 
      Height          =   330
      Left            =   960
      Top             =   1440
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
   Begin VB.ListBox List_loja 
      Appearance      =   0  'Flat
      Height          =   3855
      ItemData        =   "frm_Copy_CartLoja.frx":08B0
      Left            =   360
      List            =   "frm_Copy_CartLoja.frx":08B2
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   2805
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc adoLogo 
      Height          =   375
      Left            =   360
      Top             =   6600
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Skin_Button.ctr_Button bt_STodos 
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Seleciona todos"
      Top             =   2760
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Selecionar Todos"
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
      MICON           =   "frm_Copy_CartLoja.frx":08B4
      PICN            =   "frm_Copy_CartLoja.frx":08D0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Skin_Button.ctr_Button bt_RTodos 
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Retira Selecão de todos"
      Top             =   3360
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   873
      BTYPE           =   2
      TX              =   "Remover Seleção"
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
      MICON           =   "frm_Copy_CartLoja.frx":0BEA
      PICN            =   "frm_Copy_CartLoja.frx":0C06
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSDataListLib.DataCombo txt_Cartao 
      Bindings        =   "frm_Copy_CartLoja.frx":0F20
      DataField       =   "ctl_tipoc"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "tpc_desc"
      BoundColumn     =   "tpc_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   960
      Top             =   720
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
      Caption         =   "adoLogo"
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
   Begin MSDataListLib.DataCombo txt_Logo 
      Bindings        =   "frm_Copy_CartLoja.frx":0F38
      DataField       =   "ctl_loja"
      DataSource      =   "adoReg"
      Height          =   315
      Left            =   360
      TabIndex        =   8
      Top             =   720
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      MatchEntry      =   -1  'True
      Style           =   2
      ListField       =   "usl_nome"
      BoundColumn     =   "usl_cod"
      Text            =   ""
      Object.DataMember      =   ""
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Cartão"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Origem"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1935
   End
   Begin VB.Line Line1 
      X1              =   -1875
      X2              =   4500
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Logo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   360
      TabIndex        =   3
      Top             =   2550
      Width           =   2205
   End
End
Attribute VB_Name = "frm_Copy_CartLoja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bt_Copiar_Click()
Dim wReg As ADODB.Recordset
Dim wCodMax As Long
    
    pb.Visible = True
    wCodMax = ExecuteSQL("Select Max(ctl_cod) from tab_cartao_loja").Fields(0)
    'Retorna Array do Select
    Set wReg = ExecuteSQL("Select * from tab_cartao_loja Where ctl_loja = " & txt_Logo.BoundText & " and ctl_tipoc = " & txt_Cartao.BoundText & "").Clone
    If wReg.EOF Then
        pb.Text = "Nada consta!"
        Exit Sub
    End If
    For i = 0 To List_loja.ListCount - 1
        If List_loja.Selected(i) = True Then
            pb.Value = 0
            pb.Text = "Copiando taxa p/ Loja : " & List_loja.List(i)
            wCodMax = wCodMax + 1
            adoLogo.Recordset.MoveFirst
            adoLogo.Recordset.Find "usl_nome = '" & List_loja.List(i) & "'"
            Pause 0.5
            pb.Value = 1
            
            If List_loja.List(i) = adoLogo.Recordset.Fields("usl_nome") Then
                Pause 0.5
                wLogo = adoLogo.Recordset.Fields("usl_cod")
                
                'If Null(wLogo) or Null(txt_cartao.BoundText) or Null(wReg.Fields(2)) or Null(wReg.Fields(3)) or Null(wReg.Fields(4)) or Null(wReg.Fields(5)) or Null(wReg.Fields(6)) or Null(wReg.Fields(7)) or Null(wReg.Fields(8)) or Null(wReg.Fields(9)) or Null(wReg.Fields(10)) or Null(wReg.Fields(11)) or Null(wCodMax) Then
                    wSQL2 = "DELETE FROM tab_cartao_loja WHERE " & wReg.Fields(0).Name & " = " & wLogo & " AND " & wReg.Fields(1).Name & " = " & txt_Cartao.BoundText
                    Call ExecuteSQL(wSQL2, wRegAf2)
                    Pause 0.5
                    wSQL = "INSERT INTO tab_cartao_loja " & _
                                  "(" & wReg.Fields(0).Name & " # " & wReg.Fields(1).Name & " # " & wReg.Fields(2).Name & " # " & wReg.Fields(3).Name & " # " & wReg.Fields(4).Name & " # " & wReg.Fields(5).Name & " # " & wReg.Fields(6).Name & _
                                  " # " & wReg.Fields(7).Name & " # " & wReg.Fields(8).Name & " # " & wReg.Fields(9).Name & " # " & wReg.Fields(10).Name & " # " & wReg.Fields(11).Name & " # " & wReg.Fields(12).Name & " ) " & _
                           "VALUES " & _
                                  "(" & wLogo & " # " & txt_Cartao.BoundText & " # '" & wReg.Fields(2) & "' # '" & wReg.Fields(3) & "' # '" & wReg.Fields(4) & "' # '" & wReg.Fields(5) & "' # '" & wReg.Fields(6) & _
                                  "' # '" & wReg.Fields(7) & "' # '" & wReg.Fields(8) & "' # '" & wReg.Fields(9) & "' # '" & wReg.Fields(10) & "' # '" & wReg.Fields(11) & "' # " & wCodMax & ")"
                    wSQL = Replace(wSQL, ",", ".")
                    wSQL = Replace(wSQL, "#", ",")
                    Pause 0.5
                    pb.Value = 2
                    Call ExecuteSQL(wSQL, wRegAf)
                'End If
                If wRegAf = 0 Then MsgBox "Erro ao copiar a taxa p/ a loja " & txt_Logo, vbCritical
                Pause 0.5
                pb.Value = 3
            End If
        End If
    Next i
    pb.Text = "Copia concluída!"
    pb.Value = 3
End Sub

Private Sub bt_RTodos_Click()
    For i = List_loja.ListCount - 1 To 0 Step -1
        List_loja.Selected(i) = False
    Next i
    List_loja.Selected(0) = False
End Sub

Private Sub bt_STodos_Click()
    For i = List_loja.ListCount - 1 To 0 Step -1
        List_loja.Selected(i) = True
    Next i
    
End Sub


Private Sub Form_Load()
    Set adoLogo.Recordset = w_ado_Logo.Clone
    Set adoCartao.Recordset = w_ado_CadCartao.Clone

    Left = (MDI.Width / 2 * 0.98) - (Me.Width / 2)
    Top = ((MDI.Height / 2) * 0.89) - (Me.Height / 2) - 100
    
    'monta lista das lojas
    For i = 1 To adoLogo.Recordset.RecordCount
        Call List_loja.AddItem(adoLogo.Recordset.Fields("USL_NOME"), List_loja.ListCount)
        adoLogo.Recordset.MoveNext
    Next i
End Sub

Private Sub List_loja_Click()
    If (List_loja.List(0) = txt_Logo) Then
        List_loja.Selected(0) = False
    End If
End Sub

Private Sub List_loja_Validate(Cancel As Boolean)
    If (List_loja.List(0) = txt_Logo) Then
        List_loja.Selected(0) = False
    End If
End Sub
