VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    dwTypeData As String
    cch As Long
End Type
Private Declare Function CreatePopupMenu Lib "user32" () As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function GetMenu Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
'Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
'Private Const MF_MENUBARBREAK = &H20&
'Private Const MF_MENUBREAK = &H40&
'Private Const MF_BITMAP = &H4
Private Const MF_BYPOSITION = &H400
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8
'Private Const MFS_GRAYED = &H3&
'Private Const MFS_DISABLED = MFS_GRAYED
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4
Private Const GWL_WNDPROC = -4
Private MenuForm As Form
Rem - DTMMenu
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Public Sub UnHook()

    If MENU_OldProc > 0 Then Call SetWindowLong(MenuForm.hwnd, GWL_WNDPROC, MENU_OldProc)

End Sub

Public Function SetItemProp(ByRef hParent As Long, ByVal Property As String, ByRef Data As Variant, ByVal ID As Long) As Long
Const MF_BYCOMMAND = &H0&
Const MF_CHECKED = &H8&
Const MF_UNCHECKED = &H0&
Const MF_STRING = &H0&
    
    Select Case UCase(Property)
        Case "PICTURE"
            Call SetMenuItemBitmaps(hParent, ID, MF_BYCOMMAND, Data.Picture, Data.Picture)
        Case "BITMAP"
            Const MF_BITMAP = &H4
            Call ModifyMenu(hParent, ID, MF_BITMAP, ID, CLng(Data.Picture))
        Case "CAPTION"
            Call ModifyMenu(hParent, ID, MF_BYCOMMAND, ID, CStr(Data))
        Case "REMOVE"
            Rem - DTMMenu
            Call RemoveMenu(hParent, ID, MF_BYCOMMAND)
        Case "CHECK"
            Call ModifyMenu(hParent, ID, MF_STRING Or MF_CHECKED, ID, CStr(Data))
        Case "UNCHECK"
            Call ModifyMenu(hParent, ID, MF_STRING Or MF_UNCHECKED, ID, CStr(Data))
'        Case "ADDSUBMENU"
'            Const MASK_TEXT = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
'            Dim MItem As MENUITEMINFO
'            With MItem
'                .cbSize = Len(MItem)
'                .fMask = MASK_TEXT Or MIIM_SUBMENU
'                .hSubMenu = ID
'                .fType = MFT_STRING
'                .fState = MFS_ENABLED
'                .wID = ID
'                .dwTypeData = CStr(Data)
'                .cch = Len(.dwTypeData)
'            End With
'            Call InsertMenuItem(hParent, 0, True, MItem)
        
    End Select
    
End Function






Public Function GetFormMenu(ByVal TheForm As Form) As Long
    
    GetFormMenu = GetMenu(TheForm.hwnd)

End Function

Public Function Add(ByVal hParent As Long, ByVal Caption As String, ByVal SubMenu As Boolean, ByVal ID As Long, Optional ForceFirst As Boolean) As Long
Rem - If SubMenu is True, function resturns the submenu handle
Const MASK_TEXT = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
Dim MItem As MENUITEMINFO
Dim k As Long

    With MItem
        .cbSize = Len(MItem)
        If SubMenu Then
            .fMask = MASK_TEXT Or MIIM_SUBMENU
            .hSubMenu = CreatePopupMenu()
            Add = .hSubMenu
        Else
            .fMask = MASK_TEXT
        End If
        Rem - Regular text item or separator
        If Caption = "-" Then .fType = MFT_SEPARATOR Else .fType = MFT_STRING
        .fState = MFS_ENABLED
        .wID = ID
        .dwTypeData = Caption
        .cch = Len(.dwTypeData)
    End With
    
    Rem - Locate the specific item to which we must append this menu
    'Call InsertMenuItem(lngCurrentMenu, MenuPos(lngMaxLevel), 1, MItem)
    Call InsertMenuItem(hParent, 0, False, MItem)
    

End Function

Public Sub SetHook(ByRef TheForm As Form)

    Rem - Set up the procedure to hook the menu commands
    Set MenuForm = TheForm
    
    'Here
    #If IDE <> 1 Then
        MENU_OldProc = SetWindowLong(TheForm.hwnd, GWL_WNDPROC, AddressOf WindowMenuProc) '
    #End If
    'Here

End Sub

