VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Rem --------------------------- For RAM
Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Private Declare Function GetProcessMemoryInfo Lib "PSAPI.DLL" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Rem --------------------------- For CPU
Private Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'Private Declare Function GetTickCount& Lib "kernel32" ()
' ----
'Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As FILETIME, lpExitTime As FILETIME, lpKernelTime As FILETIME, lpUserTime As FILETIME) As Long
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
Private lngThreadID As Long
Private hProcess As Long
Dim dblLastKernel As Double
Dim dblLastTick As Double
Dim lngHWnd As Long
Public Sub Init(ByVal lngWin As Long)
Dim lngPid As Long
Const PROCESS_QUERY_INFORMATION = 1024
Const PROCESS_VM_READ = 16
    
    If lngWin <> lngHWnd Then
        If lngHWnd <> 0 Then Call CloseHandle(hProcess)
        lngThreadID = GetWindowThreadProcessId(lngWin, lngPid)
        hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, lngPid&)
    End If

End Sub

Public Function CPUUsage() As Long
Rem - Return the current CPU usage in percent
Dim dblPerc As Double
Dim FT0 As FILETIME, FT1 As FILETIME ', ST As SYSTEMTIME
Dim KernelT As FILETIME, UserT As FILETIME

    Rem - Open process and return its handle to variable hProcess...
    Call GetProcessTimes(hProcess, FT1, FT0, KernelT, UserT)

    Dim dblKTime As Double, dblUTime As Double
    Dim lngMin As Long, lngHours As Long
    dblKTime = p_UIntToDbl(KernelT.dwLowDateTime)
    dblUTime = p_UIntToDbl(UserT.dwLowDateTime)
    dblKTime = p_UIntToDbl(KernelT.dwHighDateTime) * 429496.7296 + dblKTime / 10000
    dblUTime = p_UIntToDbl(UserT.dwHighDateTime) * 429496.7296 + dblUTime / 10000
    dblKTime = (dblKTime + dblUTime) / 1000 ' Total time in ms
        
    'If dblLastTick > 0 Then lngPerc = CLng(CDbl(dblKTime - dblLastKernel) / 10 * CDbl(GetTickCount - dblLastTick))
    dblPerc = CDbl(dblKTime - dblLastKernel) / 10 * CDbl(GetTickCount - dblLastTick)
    Select Case True
        Case (dblPerc < 0) Or (dblPerc > 100): dblPerc = 0
    End Select
    CPUUsage = CLng(dblPerc)
    dblLastTick = GetTickCount
    dblLastKernel = dblKTime


End Function
Private Function p_UIntToDbl(ByVal Value As Long) As Double
    
    If ((Value And &H80000000) <> 0) Then
      p_UIntToDbl = Value And &H7FFFFFFF
      p_UIntToDbl = p_UIntToDbl + 2147483648#
    Else
        p_UIntToDbl = Value
    End If
      
End Function

Public Sub Unload()
    If hProcess <> 0 Then Call CloseHandle(hProcess)
End Sub

Public Function CPUTime() As String
Rem - Return the process CPU time as string description
Dim FT0 As FILETIME, FT1 As FILETIME ', ST As SYSTEMTIME
Dim KernelT As FILETIME, UserT As FILETIME

    Call GetProcessTimes(hProcess, FT1, FT0, KernelT, UserT)
    
    Dim dblKTime As Double, dblUTime As Double
    Dim lngMin As Long, lngHours As Long
    dblKTime = p_UIntToDbl(KernelT.dwLowDateTime)
    dblUTime = p_UIntToDbl(UserT.dwLowDateTime)
    dblKTime = p_UIntToDbl(KernelT.dwHighDateTime) * 429496.7296 + dblKTime / 10000
    dblUTime = p_UIntToDbl(UserT.dwHighDateTime) * 429496.7296 + dblUTime / 10000
    dblKTime = (dblKTime + dblUTime) / 1000 ' Total time in ms

    lngMin = Fix(dblKTime / 60)
    If lngMin > 0 Then dblKTime = dblKTime - lngMin * 60
    
    Select Case True
        Case lngMin > 59
            lngHours = lngMin \ 60
            lngMin = lngMin Mod 60
            CPUTime = CStr(lngHours) & "h" & Format(lngMin, "00") & "m"
        Case lngMin > 0
            CPUTime = CStr(lngMin) & "m" & Format(dblKTime, "00") + "s"
        Case Else
            CPUTime = Format(dblKTime, "0.0") + "s"
    End Select

End Function

Public Function RAM()
Rem - Return the RAM Usage as a string
Dim pmc As PROCESS_MEMORY_COUNTERS
    pmc.cb = LenB(pmc)
    Call GetProcessMemoryInfo(hProcess, pmc, pmc.cb)
    RAM = Format(pmc.WorkingSetSize / 1024000, "0.0") & "mb"
    
End Function

