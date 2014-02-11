; Defien global variable for installation settings
Var /GLOBAL zk_Shortcut
Var /GLOBAL zk_UserPath
Var /GLOBAL zk_Startup

Function SettingsInit
    ; Setting of 0 = Disabled , 1 = Enabled
    ; For Desktop shortcut
    StrCpy $zk_Shortcut "1"
    ReadRegStr $0 HKLM Software\ZenCODE\ZenKEY "ShortCut"
    StrCmp $0 "" +2 0
        StrCpy $zk_Shortcut $0
    
    ; For UserPath / Individual settigns
    StrCpy $zk_UserPath "1"
    ReadRegStr $0 HKLM Software\ZenCODE\ZenKEY "UserPath"
    StrCmp $0 "" +2 0
        StrCpy $zk_UserPath $0
    
    ; For loading on startup
    StrCpy $zk_Startup "1"
    ReadRegStr $0 HKLM Software\ZenCODE\ZenKEY "Startup"
    StrCmp $0 "" +2 0
        StrCpy $zk_Startup $0

FunctionEnd

Function SettingsSave

    ; Save the settings
    WriteRegStr HKLM Software\ZenCODE\ZenKEY "ShortCut" $zk_Shortcut
    WriteRegStr HKLM Software\ZenCODE\ZenKEY "UserPath" $zk_UserPath
    WriteRegStr HKLM Software\ZenCODE\ZenKEY "Startup" $zk_Startup

    ; Not make sure they are implemented
    ${If} $zk_Shortcut == "1"
        SetOutPath "$INSTDIR"
        CreateShortCut "$DESKTOP\ZenKEY.lnk" "$INSTDIR\ZenKEY.exe" "" "$INSTDIR\ZenKEY.ico" 
    ${Else}
        IfFileExists "$DESKTOP\ZenKEY.lnk" 0 +2
            Delete "$DESKTOP\ZenKEY.lnk"
    ${EndIf}

    ${If} $zk_Startup == "1"
        WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "ZenKEY" $INSTDIR\ZenKEY.exe
    ${Else}
        DeleteRegValue HKCU "Software\Microsoft\Windows\CurrentVersion\Run" "ZenKEY"
        ClearErrors
    ${EndIf}
    
FunctionEnd

