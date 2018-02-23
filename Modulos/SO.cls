VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Coloque um CLASS MODULE , não é MODULE. (CLASS MODULE é diferente de MODULE com Name "SO")

Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Private lngTotalMemory As Long
Private lngAvailableMemory As Long
Private strOSVersion As String
Private sngOSBuild As Single
Private strOSPlatform As String
Private lngProcessor As Long

Private Type MEMORYSTATUS
    Length As Long
    MemoryLoad As Long
    TotalPhys As Long
    AvailPhys As Long
    TotalPageFile As Long
    AvailPageFile As Long
    TotalVirtual As Long
    AvailVirtual As Long
End Type


Private Declare Function GetVersionEx Lib "kernel32" _
Alias "GetVersionExA" (lpOSInfo As OSVERINFO) As Boolean

Private Type OSVERINFO ' Operating System Version Information
    OSVerInfoSize As Long
    MajorVer As Long
    MinorVer As Long
    BuildNo As Long
    PlatformId As Long
    strReserved As String * 128
End Type


Private Declare Sub GetSystemInfo Lib "kernel32" _
(lpSystemInfo As SYSTEM_INFO)

Private Type SYSTEM_INFO
    OEMId As Long
    PageSize As Long
    MinimumApplicationAddress As Long
    MaximumApplicationAddress As Long
    ActiveProcessorMark As Long
    NumberOfProcessors As Long
    ProcessorType As Long
    AllocationGranularity As Long
    lngReserved As Long
End Type

Public Property Get TotalMemory() As Long
    TotalMemory = lngTotalMemory
End Property

Public Property Get AvailableMemory() As Long
    AvailableMemory = lngAvailableMemory
End Property

Public Property Get OSVersion() As String
    OSVersion = strOSVersion
End Property

Public Property Get OSBuild() As Single
    OSBuild = sngOSBuild
End Property

Public Property Get OSPlatform() As String
    OSPlatform = strOSPlatform
End Property

Public Property Get Processor() As Long
    Processor = lngProcessor
End Property

Private Sub Class_Initialize()

' Really all the work is done here
' It makes all the API calls

' Getting free memory
Dim MemStat As MEMORYSTATUS
MemStat.Length = Len(MemStat)
GlobalMemoryStatus MemStat

lngTotalMemory = Format(MemStat.TotalPhys, "Standard")
lngAvailableMemory = Format(MemStat.AvailPhys, "Standard")

' Getting version info
Dim OSInfo As OSVERINFO
OSInfo.OSVerInfoSize = Len(OSInfo)

If GetVersionEx(OSInfo) Then
    strOSVersion = OSInfo.MajorVer & "." & OSInfo.MinorVer
    sngOSBuild = OSInfo.BuildNo And &HFFFF&
    If OSInfo.PlatformId = 0 Then
        strOSPlatform = "Windows 95"
    ElseIf OSInfo.PlatformId = 1 Then
        strOSPlatform = "Windows 98"
    Else
        strOSPlatform = "Windows XP/NT/2000"
    End If
End If

' Getting system info
Dim SI As SYSTEM_INFO
GetSystemInfo SI
lngProcessor = SI.ProcessorType
End Sub



