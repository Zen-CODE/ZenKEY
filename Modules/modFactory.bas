Attribute VB_Name = "modFactory"
Option Explicit
Private Const ToolSetPath$ = "\" 'a single Backslash resolves to App.Path
Private Declare Function GetInstanceEx Lib "DirectCom" _
                        (StrPtr_FName As Long, StrPtr_ClassName As Long, _
                         Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
Private Declare Function GetInstanceOld Lib "DirectCom" Alias "GETINSTANCE" _
                        (FName As String, ClassName As String) As Object
Private Declare Function GETINSTANCELASTERROR Lib "DirectCom" () As String

'used, to preload DirectCOM.dll from a given Path, before we try our calls
Private Declare Function LoadLibraryW Lib "kernel32.dll" _
                        (ByVal LibFilePath As Long) As Long

Public Sub dhc_Remove(ByRef cCol As Collection, ByRef Item As Variant)
Dim k As Long

    For k = cCol.Count - 1 To 0 Step -1
        If cCol.ItemByIndex(k) = Item Then Call cCol.RemoveByIndex(k)
    Next k

End Sub

Public Sub dhc_Add(ByRef cCol As Collection, ByRef Item As Variant)
Dim k As Long

    For k = cCol.Count - 1 To 0 Step -1
        If cCol.ItemByIndex(k) = Item Then Exit Sub
    Next k
    cCol.Add Item

End Sub

Public Function dhc_Contains(ByRef cCol As Collection, ByRef Item As Variant) As Boolean
' A helper function to detect the precense of items in a collection
Dim k As Long
    
    For k = cCol.Count - 1 To 0 Step -1
        If Item = cCol.ItemByIndex(k) Then
            dhc_Contains = True
            Exit Function
        End If
    Next k

End Function
'deliver the constructor-helper from the factory
Public Property Get New_c() As cConstructor
  Set New_c = dhF.C
End Property

'deliver the regfree-"namespace" from the factory
Public Property Get regfree() As cRegFree
  Set regfree = dhF.regfree
End Property

Public Property Get dhF() As cFactory
Static F As cFactory, hLib As Long
  'if we already have an instance, we pass it and return immediately
    If Not F Is Nothing Then
        Set dhF = F
    Else
        #If IDE = 1 Then
            Set F = New cFactory '"normal" instancing, using VBs 'New'-Operator
            Set dhF = F
        #Else
            Dim RegFreePath As String
            RegFreePath = ToolSetPath
            If left$(RegFreePath, 1) = "\" And left$(RegFreePath, 2) <> "\\" Then
                RegFreePath = App.Path & RegFreePath 'use expansion to the App.Path
            End If
            If Right$(RegFreePath, 1) <> "\" Then RegFreePath = RegFreePath & "\"
            If hLib = 0 Then hLib = LoadLibraryW(StrPtr(RegFreePath & "DirectCOM.dll"))                '<-preload
            Set F = GetInstance(RegFreePath & "dhRichClient3.dll", "cFactory", True)
            Set dhF = F
        #End If
    End If
End Property

'The new GetInstance-Wrapper-Proc, which is using the new DirectCOM.dll (March 2009 and newer)
'with the new Unicode-capable GetInstanceEx-Call (which now supports the AlteredSearchPath-Flag as well) -
'If you omit that optional param or set it to True, then LoadLibraryExW is used with the appropriate
'Flag. If the Param was set to False, then the behaviour is the same as with the former
'DirectCOM.dll-GETINSTANCE-Call - only that LoadLibraryW is used instead of LoadLibraryA.
'This routine also tries a fallback to the former DirectCOM.dll-GETINSTANCE-Call, in case
'you are using it against an older version of this small regfree-helper-lib.
Private Function GetInstance(DllFileName As String, ClassName As String, Optional ByVal UseAlteredSearchPath As Boolean = True) As Object
  On Error Resume Next
    Set GetInstance = GetInstanceEx(StrPtr(DllFileName), StrPtr(ClassName), UseAlteredSearchPath)
  If Err.Number = 453 Then 'GetInstanceEx not available, probably an older DirectCOM.dll...
    Err.Clear
    Set GetInstance = GetInstanceOld(DllFileName, ClassName) 'so let's try the older GETINSTANCE-call
  End If
  If Err Then
    Dim Error As String
    Error = Err.Description
    On Error GoTo 0: Err.Raise vbObjectError, , Error
  Else
    If GetInstance Is Nothing Then
      On Error GoTo 0: Err.Raise vbObjectError, , GETINSTANCELASTERROR()
    End If
  End If
End Function

