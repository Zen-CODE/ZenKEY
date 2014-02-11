VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsWMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem - WMI
Private objMem As SWbemObject
Private objRefresher As SWbemRefresher
Private objCPU As SWbemObjectSet
Private lngMemTot As Long
Private strUser As String
Private strComputer As String
Private Sub Class_Initialize()
Dim objCimv2 As SWbemServicesEx

    Set objCimv2 = GetObject("winmgmts:root\cimv2")
    Set objRefresher = CreateObject("WbemScripting.SWbemRefresher")
    Rem - For whole ObjectSet
    'Private objMemory As SWbemObjectSet
    'Set objMemory = objRefresher.AddEnum(objCimv2, "Win32_PerfFormattedData_PerfOS_Memory").ObjectSet
    Set objMem = objRefresher.Add(objCimv2, "Win32_PerfFormattedData_PerfOS_Memory.Name='AvailableMBytes'").Object
    'Set objCPU = objRefresher.AddEnum(objCimv2, "Win32_PerfFormattedData_PerfOS_Processor").ObjectSet
    Set objCPU = objRefresher.AddEnum(objCimv2, "Win32_PerfFormattedData_PerfOS_Processor").ObjectSet
    
    objRefresher.Refresh

    Dim objMemTot As SWbemObjectSet
    Dim oMem As SWbemObject 'WMI Object, in this case, local CPUs

    'objMemTot = objCimv2.ExecQuery("Select * from Win32_Process where name = '" & strprocess & "'")
    Set objMemTot = objCimv2.ExecQuery("Select * from Win32_ComputerSystem")
    For Each oMem In objMemTot
        lngMemTot = Val(oMem.TotalPhysicalMemory) / 1048576
        strUser = oMem.UserName
        If InStr(strUser, "\") > 0 Then strUser = Mid(strUser, InStr(strUser, "\") + 1)
        strComputer = oMem.Name
    Next oMem

End Sub



Public Function GetMeter(ByVal Name As String) As String
Rem -----------------------------------------------------------------------------------
Rem - WMI for CPU
Rem -----------------------------------------------------------------------------------
    Dim oCpu As SWbemObject 'WMI Object, in this case, local CPUs
    Dim strTemp As String
    
    On Error GoTo Errortrap:
    
    Select Case Name
        Case "CPU"
            Rem - CPU Stats
            objRefresher.Refresh
            If objCPU.Count = 2 Then ' 1 CPU. 2nd Reading = _Total for all
                For Each oCpu In objCPU
                    GetMeter = oCpu.PercentProcessorTime & " % CPU"
                Next oCpu
            Else
                Dim k As Long
                For Each oCpu In objCPU
                    GetMeter = GetMeter & oCpu.PercentProcessorTime & "%"
                    If k >= objCPU.Count - 2 Then
                        GetMeter = GetMeter & " CPU"
                        Exit For
                    Else
                        k = k + 1
                        GetMeter = GetMeter & ", "
                    End If
                Next oCpu
            End If
            
        Case "RAM"
            Rem - RAM Stats
            GetMeter = CStr(lngMemTot - objMem.AvailableMBytes) & " / " & CStr(lngMemTot) & " MB used"
            'For Each oCpu In objMemory ' For object set
            '    GetMeter = CStr(lngMemTot - oCpu.AvailableMBytes) & " / " & CStr(lngMemTot) & " MB used"
            'Next oCpu
        Case "USER"
            GetMeter = strUser
        Case "COMPUTER"
            GetMeter = strComputer
    End Select
    Exit Function
    
Errortrap:
    GetMeter = "WMI Error " & CStr(Err.Number) & ", " & Err.Description
    Err.Clear
End Function
