Attribute VB_Name = "Threads"
Option Explicit
Public Declare Function CreateThread Lib "kernel32" _
    (ByVal lpThreadAttributes As Any, ByVal dwStackSize As _
    Long, ByVal lpStartAddress As Long, lpParameter As _
    Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SetEvent Lib "kernel32.dll" (ByVal hEvent As Long) As Long

Public Declare Function TerminateThread Lib "kernel32" _
    (ByVal hThread As Long, ByVal dwExitCode As Long) As Long

Public Declare Function TerminateProcess Lib "kernel32" _
    (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
    
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long

Public Declare Function WaitForSingleObject Lib "kernel32.dll" _
    (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Function CreateEvent& Lib "kernel32" _
    Alias "CreateEventA" (ByVal lpEventAttributes As Long, _
    ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String)

Public lThreadHandle As Long
Public lEventHandle As Long


Private Sub Progredindo_BAR()
Dim ret As Long

    Call ShowProgressInStatusBar(MDI.StatusBar1, MDI.pgBarProc, 3, True)
    Pause 0.1
    MDI.pgBarProc.Max = 10
    Do
        'DoEvents
        If MDI.pgBarProc.Value = MDI.pgBarProc.Max Then MDI.pgBarProc.Value = 0
        MDI.pgBarProc.Value = MDI.pgBarProc.Value + 1
        ret = WaitForSingleObject(lEventHandle, 500)
    Loop Until ret = 0
    Call ShowProgressInStatusBar(MDI.StatusBar1, MDI.pgBarProc, 3, False)

End Sub

Public Sub Thread_Start_ProgBar()

Dim lpThreadID As Long

    lThreadHandle = CreateThread(ByVal 0&, ByVal 0&, AddressOf Progredindo_BAR, ByVal 0&, 0, lpThreadID)
    lEventHandle = CreateEvent(ByVal 0&, False, False, ByVal 0&)

End Sub

Public Sub Thread_Stop_ProgBar()
Dim lRC As Long

    lRC = SetEvent(lEventHandle)
    
    If lThreadHandle > 0 Then
        Call TerminateThread(lThreadHandle, ByVal 0&)
    End If

    lThreadHandle = 0
    
    'Colocar no Unload do MDI
    'Call TerminateProcess(GetCurrentProcess, ByVal 0&)

End Sub

