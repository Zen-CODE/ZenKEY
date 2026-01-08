Attribute VB_Name = "modFileBrowse"
Option Explicit
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
'Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long

Private Type OPENFILENAME
    lStructSize As Long
    hWndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Rem - For Directory browsing
Const MAX_PATH = 260
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = 1
'Private Const MAX_PATH = 260
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function FBR_GetLastFolder(ByVal Path As String) As String
Dim k As Long, lngPos As Long
Dim lngMax As Long

    Do
        k = InStr(k + 1, Path, "\")
        If k <> 0 Then lngPos = k
    Loop Until k = 0
    FBR_GetLastFolder = Mid$(Path, lngPos + 1)
    
    
End Function

Private Function Extract(ByVal Sentance As String, ByVal AfterNthSep As Long, ByVal Separator As String) As String
Rem - Pumps the pipe separated items into Items()
Dim k As Integer, intEnd As Integer

    intEnd = InStr(Sentance, Separator)
    For k = 0 To AfterNthSep - 1
        If intEnd > 0 Then
            Sentance = Mid$(Sentance, intEnd + 1)
        Else
            Sentance = vbNullString
        End If
        intEnd = InStr(Sentance, Separator)
    Next k
    intEnd = InStr(Sentance, Separator)
    If intEnd > 0 Then Extract = left$(Sentance, intEnd - 1) Else Extract = Sentance

End Function
Public Function FBR_GetOFName(ByVal Caption As String, ByRef FName As String, ParamArray FileType()) As Boolean
Dim OFName As OPENFILENAME

    
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = MainForm.hwnd
    OFName.hInstance = App.hInstance
    Rem - Select a filter
    'OFName.lpstrFilter = "Text Files (*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
    'If Len(FileType) = 0 Then
    If UBound(FileType) < 0 Then
        OFName.lpstrFilter = "All Files (*.*)" + Chr$(0)
    Else
        'OFName.lpstrFilter = FileType + Chr$(0)
        Dim k As Integer
        For k = 0 To UBound(FileType())
            'OFName.lpstrFilter = FileType + Chr$(0) & Extract(Extract(FileType, 1, "("), 0, ")")
            OFName.lpstrFilter = OFName.lpstrFilter & FileType(k) + Chr$(0) & Extract(Extract(FileType(k), 1, "("), 0, ")") + Chr$(0)
        Next k
    End If
    'create a buffer for the file
    OFName.lpstrFile = Space$(254)
    'set the maximum length of a returned file
    OFName.nMaxFile = 255
    'Create a buffer for the file title
    OFName.lpstrFileTitle = Space$(254)
    'Set the maximum length of a returned file title
    OFName.nMaxFileTitle = 255
    'Set the initial directory
    If Len(FName) > 0 Then
        OFName.lpstrInitialDir = FName
    Else
        OFName.lpstrInitialDir = App.Path ' "C:\"
    End If
    'Set the title
    OFName.lpstrTitle = Caption ' "Select File - Powerkey"
    'No flags
    OFName.flags = 0

    'Show the 'Open File'-dialog
    FBR_GetOFName = GetOpenFileName(OFName)
    If FBR_GetOFName Then
        FName = Trim(OFName.lpstrFile)
        FName = left$(FName, Len(FName) - 1)
    End If

End Function




Public Function FBR_BrowseForFolder(ByVal Caption As String, ByRef strFName As String) As Boolean
Dim iNull As Integer, lpIDList As Long, lResult As Long
Dim sPath As String, udtBI As BrowseInfo

    With udtBI
        Rem - Set the owner window
        .hWndOwner = MainForm.hwnd
        Rem - lstrcat appends the two strings and returns the memory address
        '.lpszTitle = lstrcat("C:\", "")
        .lpszTitle = lstrcat(Caption, "")
        Rem - Return only if the user selected a directory
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With

    Rem - Show the 'Browse for folder' dialog
    lpIDList = SHBrowseForFolder(udtBI)
    FBR_BrowseForFolder = CBool(lpIDList <> 0)
    If FBR_BrowseForFolder Then
        Rem - Get the path from the IDList
        sPath = String$(MAX_PATH, 0)
        SHGetPathFromIDList lpIDList, sPath
        'free the block of memory
        CoTaskMemFree lpIDList
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            'sPath = Left$(sPath, iNull - 1)
            strFName = left$(sPath, iNull - 1)
        End If
    End If

End Function

