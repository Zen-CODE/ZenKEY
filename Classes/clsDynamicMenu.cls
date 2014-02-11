VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsDynamicMenu"
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
Private Declare Function InsertMenu Lib "user32" Alias "InsertMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function GetMenu Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function GetMenuItemID Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
'Private Declare Function ModifyMenuBynum Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Long) As Long
Private Declare Function GetSubMenu Lib "user32" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal Hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal Hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
'Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
'Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPosition As Long, lpmii As MENUITEMINFO) As Long
Private Declare Function ModifyMenu Lib "user32" Alias "ModifyMenuA" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpString As Any) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Const MF_MENUBARBREAK = &H20&
Private Const MF_MENUBREAK = &H40&
Private Const MF_BITMAP = &H4
Private Const MF_BYPOSITION = &H400
Private Const MFT_SEPARATOR = &H800
Private Const MFT_STRING = &H0
Private Const MFS_ENABLED = &H0
Private Const MFS_CHECKED = &H8
Private Const MFS_GRAYED = &H3&
Private Const MFS_DISABLED = MFS_GRAYED
Private Const MIIM_STATE = &H1
Private Const MIIM_ID = &H2
Private Const MIIM_TYPE = &H10
Private Const MIIM_SUBMENU = &H4
'Private Const MIIM_BITMAP = &H80
'Private Const MIIM_FTYPE = &H100
'Private Const MF_CHECKED = &H8&
'Private Const MF_APPEND = &H100&
'Private Const MF_DISABLED = &H2&
'Private Const MF_GRAYED = &H1&
'Private Const MF_SEPARATOR = &H800&
'Private Const MF_STRING = &H0&
'Private Const MF_MENUBARBREAK = &H20&
'Private Const MF_MENUBREAK = &H40&
'Private Const MF_BITMAP = &H4
'Private Const MF_BYPOSITION = &H400

Private Const GWL_WNDPROC = -4
Public MENU_OldProc As Long  ' pointer to Form1's previous window procedure
Private MenuForm As Form
Private lngMenuHandle As Long

Public Sub Init(ByRef TheForm As Form)
    
    lngMenuHandle = GetMenu(TheForm.Hwnd)
    lngMenuHandle = GetSubMenu(lngMenuHandle, 0) ' And get a handle to the Floating popup menu.
    If lngMenuHandle = 0 Then
        Call MsgBox("The form that has been passed to the 'clsDynamicMenu' initialisation has no Menu!", vbExclamation, ZenKEYCap)
    Else
        Rem - Set up the procedure to hook the menu commands
        Set MenuForm = TheForm
        MENU_OldProc = SetWindowLong(TheForm.Hwnd, GWL_WNDPROC, AddressOf WindowProc) '
    End If
    

End Sub

'
Public Sub AddMenuItem(ByVal Caption As String, ByVal Properties As String, ByVal ID As Long, ParamArray MenuPos())
Rem - The MenuPos array contains the depth of the menu e.g. ..1) is a 1st level menu item, ..2,1) adds
Rem - a subitem to the 2nd Menu item
Const MASK_TEXT = MIIM_STATE Or MIIM_ID Or MIIM_TYPE
Dim MItem As MENUITEMINFO
Dim k As Long
Dim lngMaxLevel As Long
Dim lngCurrentMenu As Long

    With MItem
        .cbSize = Len(MItem)
        If InStr(UCase(Properties), "SUBMENU") > 0 Then
            .fMask = MASK_TEXT Or MIIM_SUBMENU
            .hSubMenu = CreatePopupMenu()
        Else
            .fMask = MASK_TEXT
        End If
        Rem - Regular text item or separator
        If Caption = "-" Then .fType = MFT_SEPARATOR Else .fType = MFT_STRING
        
        If InStr(UCase(Properties), "DISABLED") = 0 Then
            .fState = MFS_ENABLED
        Else
            .fState = MFS_DISABLED
        End If
        .wID = ID
        .dwTypeData = Caption
        .cch = Len(.dwTypeData)
    End With
    
    Rem - Locate the specific item to which we must append this menu
    lngMaxLevel = UBound(MenuPos())
    lngCurrentMenu = lngMenuHandle
    For k = 1 To lngMaxLevel
        lngCurrentMenu = GetSubMenu(lngCurrentMenu, MenuPos(k - 1))
    Next k
    If lngCurrentMenu = 0 Then
        Call MsgBox("Invalid MenuPos specification! Sub AddMenuItem", vbInformation, "ZenKEY")
    Else
        Call InsertMenuItem(lngCurrentMenu, MenuPos(lngMaxLevel), 1, MItem)
    End If
    


End Sub

Public Sub UnHook()

    If MENU_OldProc > 0 Then Call SetWindowLong(MenuForm.Hwnd, GWL_WNDPROC, MENU_OldProc)

End Sub

Public Sub SetItemProp(ByRef Property As String, ByRef Data As Variant, ByVal ID As Long)
Const MF_BYCOMMAND = &H0&

    Select Case ID
        Case -1 ' First item
            Call SetMenuItemBitmaps(lngMenuHandle, 0, MF_BYPOSITION, Data.Picture, Data.Picture)
            Exit Sub
        Case -2 ' Last item
            Call SetMenuItemBitmaps(lngMenuHandle, GetMenuItemCount(lngMenuHandle) - 1, MF_BYPOSITION, Data.Picture, Data.Picture)
            Exit Sub
    End Select
    
    Select Case UCase(Property)
        Case "PICTURE"
            Call SetMenuItemBitmaps(lngMenuHandle, ID, MF_BYCOMMAND, Data.Picture, Data.Picture)
        Case "CAPTION"
            Call ModifyMenu(lngMenuHandle, ID, MF_BYCOMMAND, ID, CStr(Data))
    End Select
    
End Sub



Public Function GetItemHandle(Optional ByVal Index As Long = -1) As Long
    If Index > -1 Then
        GetItemHandle = GetSubMenu(lngMenuHandle, Index)
    Else
        GetItemHandle = lngMenuHandle
    End If
End Function


