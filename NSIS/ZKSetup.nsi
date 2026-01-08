;--------------------------------
;Include Modern UI + DLL Library installer
!include "MUI2.nsh"
!include "LogicLib.nsh"
;--------------------------------
;Name and file
Name "ZenKEY"
OutFile "ZKSetup.exe"

;Default installation folder
InstallDir "$PROGRAMFILES\ZenKEY"

;Get installation folder from registry if available
InstallDirRegKey HKLM "Software\ZenCODE\ZenKEY" "ZKPath"
Var StartMenuFolder

;Request application privileges for Windows Vista
RequestExecutionLevel admin
!include Settings.nsi
;--------------------------------
;Interface Settings
!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "ZenKEY.bmp" ; optional
!define MUI_WELCOMEFINISHPAGE_BITMAP "ZenKEY Cross.bmp"
!define MUI_WELCOMEPAGE_TITLE  "Welcome to the ZenKEY 2.5.1 Setup Wizard"

;!define MUI_HEADERIMAGE_BITMAP "Z:\Zen\My Documents\ZenCODE\NSIS\Y4_Medium.bmp" ; optional
!define MUI_ABORTWARNING
!define MUI_STARTMENUPAGE_TEXT_TOP "Please select the start menu folder."
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "ZenCode\ZenKEY"
;!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_TEXT_CHECKBOX "Do no create Startmenu entries"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT HKLM
!define MUI_STARTMENUPAGE_REGISTRY_KEY "Software\ZenCODE\ZenKEY"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "StartMenu"
!define MUI_FINISHPAGE_RUN $INSTDIR\ZenKEY.exe
!define MUI_FINISHPAGE_RUN_TEXT "Run ZenKEY now"

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_WELCOME
  !insertmacro MUI_PAGE_LICENSE "License.rtf"
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_STARTMENU "ZenKEY" $StartMenuFolder
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  !insertmacro MUI_PAGE_FINISH
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections

Section "ZenKEY" ZKDescrip
  SectionIn RO
  SetOutPath "$INSTDIR"
  
  ;ADD YOUR OWN FILES HERE...
  # Dependencies
  
  # First test to see if ZenKEY is in use...
  ClearErrors
  FileOpen $R0 $INSTDIR\ZenKEY.exe w
  ${If} ${Errors}
     # File is locked.
     MessageBox MB_OK "ZenKEY is currently running. Thank you ;-). Please exit ZenKEY and then click OK."
  ${Else}
     FileClose $R0
  ${EndIf}
  # Do things this way to allow portability on XP upwards
  File "C:\Windows\System32\mscomctl.ocx"
  File "C:\Windows\System32\msvbvm60.dll"
  IfFileExists $SYSDIR\mscomctl.ocx +2 0 # If not in Sys32, copy there
	CopyFiles $INSTDIR\mscomctl.ocx $SYSDIR\mscomctl.ocx
  RegDLL $SYSDIR\mscomctl.ocx
  IfFileExists $SYSDIR\msvbvm60.dll +2 0
	CopyFiles "C:\Windows\System32\msvbvm60.dll" $SYSDIR\msvbvm60.dll
  RegDLL $SYSDIR\msvbvm60.dll
  
  # Executables + Manifests
  File "Z:\ZenKEY\ZenKEY.exe"
  File "Z:\ZenKEY\ZenKEY.exe.manifest"
  File "Z:\ZenKEY\ZenDim.exe"
  File "Z:\ZenKEY\ZenKP.exe"
  File "Z:\ZenKEY\ZenWiz.exe"
  File "Z:\ZenKEY\ZenWiz.exe.manifest"
  File "Z:\ZenKEY\ZKConfig.exe"
  File "Z:\ZenKEY\ZKConfig.exe.manifest"
  File "Z:\ZenKEY\ZenKEY.ico"
  
  # INI Files
  File "Z:\ZenKEY\Actions.ini"
  File "Z:\ZenKEY\Default_Complete.ini"
  File "Z:\ZenKEY\Default_Complete7.ini"
  IfFileExists $INSTDIR\ZenKEY.ini +2 0
	File "/oname=ZenKEY.ini" "Z:\ZenKEY\Default_Complete.ini"
  File "Z:\ZenKEY\Default.ini"
  File "Z:\ZenKEY\DTMMenu.ini"
  File "Z:\ZenKEY\Search.ini"
  File "Z:\ZenKEY\SetList.ini"
  File "Z:\ZenKEY\SetDef.ini"
  IfFileExists $INSTDIR\Settings.ini +2 0
	File "/oname=Settings.ini" "Z:\ZenKEY\SetDef.ini"
  
  # Subfolders
  File /r "Z:\ZenKEY\Help"
  File /r "Z:\ZenKEY\Quotes"
  File /r "Z:\ZenKEY\Skins"
  CreateDirectory $INSTDIR\Icons
  
  # Other files
  File "Z:\ZenKEY\Show Desktop.scf"
  File "Z:\ZenKEY\Document.rtf"
  
  ; Save and implement install options
  WriteRegStr HKLM "Software\ZenCODE\ZenKEY" "ZKPath" $INSTDIR  
  WriteRegStr HKLM "Software\ZenCODE\ZenKEY" "StartMenu" $StartMenuFolder  
  Call SettingsSave
  
  ; Create start menu items
  !insertmacro MUI_STARTMENU_WRITE_BEGIN ZenKEY
    CreateDirectory "$SMPROGRAMS\$StartMenuFolder"
	CreateShortCut "$SMPROGRAMS\$StartMenuFolder\ZenKEY.lnk" "$INSTDIR\ZenKEY.exe" "" "$INSTDIR\ZenKEY.ico" 
	CreateShortCut "$SMPROGRAMS\$StartMenuFolder\ZenKEY Configuration Utility.lnk" "$INSTDIR\ZKConfig.exe" "" "$INSTDIR\ZenKEY.ico" 
	CreateShortCut "$SMPROGRAMS\$StartMenuFolder\ZenKEY Wizard.lnk" "$INSTDIR\ZenWiz.exe" "" "$INSTDIR\ZenKEY.ico" 
	CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Help.lnk" "$INSTDIR\Help\Index.htm"  
    CreateShortCut "$SMPROGRAMS\$StartMenuFolder\Uninstall.lnk" "$INSTDIR\Uninstall.exe"  
  !insertmacro MUI_STARTMENU_WRITE_END  
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "DisplayName" "ZenKEY"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "DisplayVersion" "2.5.1"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "DisplayIcon" "$INSTDIR\ZenKEY.exe,0"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "UninstallString" "$INSTDIR\Uninstall.exe"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "InstallLocation" "$INSTDIR"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "Publisher" "ZenCODE"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "HelpLink" "http://www.camiweb.com/zenkey"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "URLInfoAbout" "http://www.camiweb.com/zenkey"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "URLUpdateInfo" "http://www.camiweb.com/zenkey/Download.htm"
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY" "NoRepair" 1
  
SectionEnd

Section "Per-user settings" UserPath
    ; Done on .onSelChange change as this does not fire if unticked
SectionEnd

Section "Destop shortcut" Shortcut
    ; Done on .onSelChange change as this does not fire if unticked
SectionEnd

Section "Load on startup" Startup
    ; Done on .onSelChange change as this does not fire if unticked
SectionEnd

Function .onInit

    Call SettingsInit
    SectionSetFlags ${UserPath} $zk_UserPath
    SectionSetFlags ${Shortcut} $zk_Shortcut
    SectionSetFlags ${Startup} $zk_Startup
    
FunctionEnd 

Function .onSelChange

    SectionGetFlags ${UserPath} $zk_UserPath
    SectionGetFlags ${Shortcut} $zk_Shortcut
    SectionGetFlags ${Startup} $zk_Startup
    
FunctionEnd

;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_ZKDescrip ${LANG_ENGLISH} "ZenKEY program files"
  LangString DESC_UserPath ${LANG_ENGLISH} "Stores settings for each user in thier own folder. All settings are otherwise stored in the ZenKEY folder (untick for portability or global settings)"
  LangString DESC_Shortcut ${LANG_ENGLISH} "Places a ZenKEY icon on your desktop"
  LangString DESC_Startup ${LANG_ENGLISH} "Automatically runs ZenKEY when Windows starts"

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
  !insertmacro MUI_DESCRIPTION_TEXT ${ZKDescrip} $(DESC_ZKDescrip)
  !insertmacro MUI_DESCRIPTION_TEXT ${UserPath} $(DESC_UserPath)
  !insertmacro MUI_DESCRIPTION_TEXT ${Shortcut} $(DESC_Shortcut)
  !insertmacro MUI_DESCRIPTION_TEXT ${Startup} $(DESC_Startup)
  !insertmacro MUI_FUNCTION_DESCRIPTION_END

;--------------------------------
;Uninstaller Section

Section "Uninstall"

  ;ADD YOUR OWN FILES HERE...
  Delete "$INSTDIR\*.*"  
  RMDir /r "$INSTDIR\Help"
  RMDir /r "$INSTDIR\Quotes"
  RMDir /r "$INSTDIR\Skins"
  RMDir /r "$INSTDIR\Icons"
  RMDir "$INSTDIR"
  DeleteRegKey /ifempty HKLM "Software\ZenCODE\ZenKEY"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\ZenKEY"
  
  #!insertmacro MUI_STARTMENU_GETFOLDER page_id $R0
  !insertmacro MUI_STARTMENU_GETFOLDER "ZenKEY" $R0
  Delete "$SMPROGRAMS\$R0\ZenKEY.lnk"
  Delete "$SMPROGRAMS\$R0\ZenKEY Configuration Utility.lnk"
  Delete "$SMPROGRAMS\$R0\ZenKEY Wizard.lnk"
  Delete "$SMPROGRAMS\$R0\Help.lnk"
  Delete "$SMPROGRAMS\$R0\Uninstall.lnk"
  Delete "$DESKTOP\ZenKEY.lnk"

SectionEnd