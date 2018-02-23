Attribute VB_Name = "func_int"
Public myThread_enable As Boolean


Private Const WM_USER As Long = &H400
Private Const SB_GETRECT As Long = (WM_USER + 10)

Private Type RECT
 Left As Long
 Top As Long
 Right As Long
 Bottom As Long
End Type

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SendMessageAny Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub Sleep Lib "Kernel32" (ByVal dwMilliseconds As Long)

Public Sub ShowProgressInStatusBar(oStatus As StatusBar, oPicture As PictureBox, Optional iPanelIndex As Integer = 1, Optional bShowPicture As Boolean = True)

'Public Sub ShowProgressInStatusBar(oPicture As PictureBox, Optional iPanelIndex As Integer = 1, Optional bShowPicture As Boolean = True)
' --------------------------------------------------------------------
' Lembre-se que o indice do painel na api é sempre 1
' a menos que seu numero real (ou seja, 2 é 1)
' --------------------------------------------------------------------
Dim tRC As RECT

If (oStatus Is Nothing) Or (oPicture Is Nothing) Or (iPanelIndex < 1) Or (iPanelIndex > oStatus.Panels.Count) Then
  Exit Sub
End If

If bShowPicture Then
  SendMessageAny oStatus.hwnd, SB_GETRECT, iPanelIndex - 1, tRC
  With tRC
   .Top = (.Top * Screen.TwipsPerPixelY)
   .Left = (.Left * Screen.TwipsPerPixelX)
   .Bottom = (.Bottom * Screen.TwipsPerPixelY) - .Top
   .Right = (.Right * Screen.TwipsPerPixelX) - .Left
  End With

  With oPicture
   SetParent .hwnd, oStatus.hwnd
   .Move tRC.Left + 10, tRC.Top + 10, tRC.Right - 15, tRC.Bottom - 25
   .Visible = True
  End With
Else
  SetParent oPicture.hwnd, 0
  oPicture.Visible = False
End If

End Sub



Public Sub Progredindo_BAR()
Dim Max As Integer, wValue As Integer

On Error Resume Next

Max = MDI.Picture1.Width

Do While myThread_enable = True
        DoEvents
        If wValue >= Max Then
            wValue = 0
        ElseIf (wValue + 100) > Max Then
            wValue = Max
        Else
            wValue = wValue + 50
        End If
        
        MDI.Picture1.Width = wValue
        MDI.Picture1.Refresh
        Sleep 200

Loop


End Sub

Public Sub Thread_Start(ByRef myThread)
    ShowProgressInStatusBar MDI.StatusBar1, MDI.Picture1, 3, True
    Set myThread = New clsThreads
    With myThread
        .Initialize AddressOf ExecuteSql_2
        .Enabled = True
        myThread_enable = True
    End With
    Call Progredindo_BAR
End Sub


Public Sub Thread_Stop(ByRef myThread)

    myThread_enable = False
    Set myThread = Nothing
    ShowProgressInStatusBar MDI.StatusBar1, MDI.Picture1, 3, False

End Sub


Public Sub Thread_Start_PgBar(ByRef myThread)
    ShowProgressInStatusBar MDI.StatusBar1, MDI.Picture1, 3, True
    Set myThread = New clsThreads
    With myThread
        .Initialize AddressOf Progredindo_BAR
        .Enabled = True
        myThread_enable = True
    End With
    'Call Progredindo_BAR
End Sub
