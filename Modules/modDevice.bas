Attribute VB_Name = "modDevice"
Private Declare Function QueryDosDeviceW Lib "kernel32.dll" ( _
    ByVal lpDeviceName As Long, _
    ByVal lpTargetPath As Long, _
    ByVal ucchMax As Long _
    ) As Long
Private Const MAX_PATH = 260
Private Declare Function GetLogicalDriveStringsA Lib "kernel32" (ByVal nBufferLength As Long, lpBuffer As Any) As Long
Private Function GetNtDeviceNameForDrive(ByVal sDrive As String) As String
Dim colDrives As Collection
Dim bDrive() As Byte
Dim bResult() As Byte
Dim lR As Long
Dim sDeviceName As String

   If Right(sDrive, 1) = "\" Then
      If Len(sDrive) > 1 Then
         sDrive = left(sDrive, Len(sDrive) - 1)
      End If
   End If
   bDrive = sDrive
   ReDim Preserve bDrive(0 To UBound(bDrive) + 2) As Byte
   ReDim bResult(0 To MAX_PATH * 2 + 1) As Byte
   lR = QueryDosDeviceW(VarPtr(bDrive(0)), VarPtr(bResult(0)), MAX_PATH)
   If (lR > 2) Then
      sDeviceName = bResult
      sDeviceName = left(sDeviceName, lR - 2)
      GetNtDeviceNameForDrive = sDeviceName
   End If
   
End Function


Private Function GetDrives() As Collection
Dim colDrives As New Collection
Dim lSize As Long
Dim lR As Long
Dim iLastPos As Long
Dim iPos As Long
Dim sDrive As String
Dim sDriveStrings As String

   lSize = GetLogicalDriveStringsA(0, ByVal 0&)
   sDriveStrings = String(lSize + 1, 0)
   lR = GetLogicalDriveStringsA(lSize, ByVal sDriveStrings)
   iLastPos = 1
   Do
      iPos = InStr(iLastPos, sDriveStrings, vbNullChar)
      If Not (iPos = 0) Then
         sDrive = Mid$(sDriveStrings, iLastPos, iPos - iLastPos)
         iLastPos = iPos + 1
      Else
         sDrive = Mid$(sDriveStrings, iLastPos)
      End If
      If Len(sDrive) > 0 Then
         colDrives.Add sDrive
      End If
   Loop While Not (iPos = 0)
   Set GetDrives = colDrives
   
End Function
Public Function GetDriveForNtDeviceName(ByVal sDeviceName As String) As String
Dim sFoundDrive As String
Dim colDrives As Collection
Dim vDrive As Variant

    'TODO : Optimize - Kee[p dictionary, lookup as needed
   For Each vDrive In GetDrives()
      If (GetNtDeviceNameForDrive(vDrive) = sDeviceName) Then
         sFoundDrive = vDrive
         Exit For
      End If
   Next
   GetDriveForNtDeviceName = sFoundDrive
   
End Function

