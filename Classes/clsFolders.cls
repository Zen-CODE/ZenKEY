VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private Const CSIDL_ADMINTOOLS = &H30
'Private Const CSIDL_ALTSTARTUP = &H1D
'Private Const CSIDL_APPDATA = &H1A
'Private Const CSIDL_BITBUCKET = &HA
'Private Const CSIDL_COMMON_ADMINTOOLS = &H2F
'Private Const CSIDL_COMMON_ALTSTARTUP = &H1E
'Private Const CSIDL_COMMON_APPDATA = &H23
'Private Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
'Private Const CSIDL_COMMON_DOCUMENTS = &H2E
'Private Const CSIDL_COMMON_FAVORITES = &H1F
'Private Const CSIDL_COMMON_PROGRAMS = &H17
'Private Const CSIDL_COMMON_STARTMENU = &H16
'Private Const CSIDL_COMMON_STARTUP = &H18
'Private Const CSIDL_COMMON_TEMPLATES = &H2D
'Private Const CSIDL_CONNECTIONS = &H31
'Private Const CSIDL_CONTROLS = &H3
'Private Const CSIDL_COOKIES = &H21
'Private Const CSIDL_DESKTOP = &H0
'Private Const CSIDL_DESKTOPDIRECTORY = &H10
'Private Const CSIDL_DRIVES = &H11
'Private Const CSIDL_FAVORITES = &H6
'Private Const CSIDL_FONTS = &H14
'Private Const CSIDL_HISTORY = &H22
'Private Const CSIDL_INTERNET = &H1
'Private Const CSIDL_INTERNET_CACHE = &H20
'Private Const CSIDL_LOCAL_APPDATA = &H1C
'Private Const CSIDL_MYPICTURES = &H27
'Private Const CSIDL_NETHOOD = &H13
'Private Const CSIDL_NETWORK = &H12
'Private Const CSIDL_PERSONAL = &H5
'Private Const CSIDL_PRINTERS = &H4
'Private Const CSIDL_PRINTHOOD = &H1B
'Private Const CSIDL_PROFILE = &H28
'Private Const CSIDL_PROGRAM_FILES = &H26
'Private Const CSIDL_PROGRAM_FILES_COMMON = &H2B
'Private Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
'Private Const CSIDL_PROGRAM_FILESX86 = &H2A
'Private Const CSIDL_PROGRAMS = &H2
'Private Const CSIDL_RECENT = &H8
'Private Const CSIDL_SENDTO = &H9
'Private Const CSIDL_STARTMENU = &HB
'Private Const CSIDL_STARTUP = &H7
'Private Const CSIDL_SYSTEM = &H25
'Private Const CSIDL_SYSTEMX86 = &H29
'Private Const CSIDL_TEMPLATES = &H15
'Private Const CSIDL_WINDOWS = &H24

Const CSIDL_DESKTOP = &H0
Const CSIDL_PROGRAMS = &H2
Const CSIDL_CONTROLS = &H3
Const CSIDL_PRINTERS = &H4
Const CSIDL_PERSONAL = &H5
Const CSIDL_FAVORITES = &H6
Const CSIDL_STARTUP = &H7
Const CSIDL_RECENT = &H8
Const CSIDL_SENDTO = &H9
Const CSIDL_BITBUCKET = &HA
Const CSIDL_STARTMENU = &HB
Const CSIDL_DESKTOPDIRECTORY = &H10
Const CSIDL_DRIVES = &H11
Const CSIDL_NETWORK = &H12
Const CSIDL_NETHOOD = &H13
Const CSIDL_FONTS = &H14
Const CSIDL_TEMPLATES = &H15
Const MAX_PATH = 260
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
Private Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
'Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Function GetSpecialfolder(ByVal CSIDL As Long) As String
Dim r As Long
Dim Path As String
Dim IDL As ITEMIDLIST
    
    r = SHGetSpecialFolderLocation(100, CSIDL, IDL)
    If r = 0 Then
        Rem - No error. Create a buffer
        Path = Space$(512)
        r = SHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal Path$)
        GetSpecialfolder = Left$(Path, InStr(Path, Chr$(0)) - 1)
        Exit Function
    End If
    GetSpecialfolder = ""
End Function


Public Sub DoAction(ByVal Action As String)
Dim FName As String, FNum As Long
    'Call MsgBox("Stub DoAction fired! - " & Action, vbInformation)
    Select Case Action
        Case "DESKTOP": FNum = &H0
        Case "PROGRAMS": FNum = &H2
        Case "CONTROLS": FNum = &H3
        Case "PRINTERS": FNum = &H4
        Case "PERSONAL": FNum = &H5
        Case "FAVORITES": FNum = &H6
        Case "STARTUP": FNum = &H7
        Case "RECENT": FNum = &H8
        Case "SENDTO": FNum = &H9
        Case "BITBUCKET": FNum = &HA
        Case "STARTMENU": FNum = &HB
        Case "MUSIC": FNum = &HD
        Case "DESKTOPDIRECTORY": FNum = &H10
        Case "DRIVES": FNum = &H11
        Case "NETWORK": FNum = &H12
        Case "NETHOOD": FNum = &H13
        Case "FONTS": FNum = &H14
        Case "TEMPLATES": FNum = &H1
        Case "HISTORY": FNum = &H22
        Case "MYPICTURES": FNum = &H27
        Case "PROGRAM_FILES": FNum = &H26
        Case "PROGRAM_FILES_COMMON": FNum = &H2B
        Case "WINDOWS": FNum = &H24
    End Select
    FName = Trim(GetSpecialfolder(FNum))
    If Len(FName) = 0 Then
        Call MsgBox("And Ichi asked 'If a folder cannot be found, is it still there?' Fomo smiled. His student was beginning to understand '" & Action & "' = " & FName, vbInformation, ZenCODECap)
    Else
        Call ShellExecute(Forms(0).hwnd, "open", "explorer.exe", FName, vbNullString, SW_NORMAL)
    End If
End Sub
