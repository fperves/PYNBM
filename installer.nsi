;NSIS Modern User Interface
;Basic Example Script
;Written by Joost Verburg

;--------------------------------
;Include Modern UI

  !include "MUI2.nsh"
  !include WinVer.nsh
;--------------------------------
;General

  ;Name and file
  Name "PYNBM-${VERSION}"
  OutFile "PYNBM-${VERSION}.exe"
  RequestExecutionLevel admin 
  ;Default installation folder
  InstallDir "$PROGRAMFILES\PYNBM"
  
  ;Get installation folder from registry if available
  InstallDirRegKey HKCU "Software\PYNBM" ""


  ;Request application privileges for Windows Vista
  RequestExecutionLevel admin
  SetCompressor /SOLID lzma

;--------------------------------
;Interface Settings

  !define MUI_ABORTWARNING

;--------------------------------
;Pages

  !insertmacro MUI_PAGE_LICENSE License.txt
  !insertmacro MUI_PAGE_COMPONENTS
  !insertmacro MUI_PAGE_DIRECTORY
  !insertmacro MUI_PAGE_INSTFILES
  
  !insertmacro MUI_UNPAGE_CONFIRM
  !insertmacro MUI_UNPAGE_INSTFILES
  
;--------------------------------
;Languages
 
  !insertmacro MUI_LANGUAGE "English"

;--------------------------------
;Installer Sections

Section "PYNBM Application" SecApplication

  SetOutPath "$INSTDIR\Images"
  File images\*.*  
  
  SetOutPath "$APPDATA\PYNBM"
  
  SetOutPath "$INSTDIR"
  File dist\*.*
  File License.txt
  File Extensions.conf_readme.txt
  File Extensions.defaults
  ;${If} ${IsWin2008}
  ;${OrIf} ${IsWin2008R2}
  ;${OrIf} ${IsWin7}
  ;${OrIf} ${IsWinVista}
  
	;File PYNBM.exe.manifest
  ;${EndIf}

  ;Store installation folder
  WriteRegStr HKCU "Software\PYNBM" "" $INSTDIR
  
  ;Add entry to windows add/remove programs
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"  \
					"DisplayName" "PYNBM"  										
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"  \					
					"DisplayIcon" "$INSTDIR\Images\Polling.ico"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"  \					
					"DisplayVersion" "${VERSION}"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"  \					
					"Publisher" "Florian PERVES"
  WriteRegStr HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"  \
					"UninstallString" "$INSTDIR\uninstall.exe"
  
  ;Create uninstaller
  WriteUninstaller "$INSTDIR\Uninstall.exe"


SectionEnd

; Optional section (can be disabled by the user)
Section "Start Menu Shortcuts" SecShortcuts

  CreateDirectory "$SMPROGRAMS\PYNBM"  
  CreateShortCut "$SMPROGRAMS\PYNBM\PYNBM.lnk" "$INSTDIR\PYNBM.exe" "	" "$INSTDIR\Images\Polling.ico" 0
  CreateShortCut "$SMPROGRAMS\PYNBM\Uninstall.lnk" "$INSTDIR\uninstall.exe" "" "$INSTDIR\uninstall.exe" 0
  
SectionEnd

Section "Source Code" SecSource
  SetOutPath "$INSTDIR\source"
  File PYNBM.py
  File pynbm_*.py
  File license.txt
  
  
SectionEnd


;--------------------------------
;Descriptions

  ;Language strings
  LangString DESC_SecApplication ${LANG_ENGLISH} "Install the Application."
  LangString DESC_SecSource ${LANG_ENGLISH} "Install the Source Code."
    LangString DESC_SecShortcuts ${LANG_ENGLISH} "Install Start Menu Shortcuts."

  ;Assign language strings to sections
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${SecApplication} $(DESC_SecApplication)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecSource} $(DESC_SecSource)
	!insertmacro MUI_DESCRIPTION_TEXT ${SecShortcuts} $(DESC_SecShortcuts)	
  !insertmacro MUI_FUNCTION_DESCRIPTION_END





;--------------------------------
;Uninstaller Section

Section "Uninstall"

  Delete "$INSTDIR\*.*"
  Delete "$INSTDIR\source\*.*"
  RMDir "$INSTDIR\source"
  Delete "$INSTDIR\Images\*.*"
  RMDir "$INSTDIR\Images"
  RMDir "$INSTDIR"

    ; Remove shortcuts, if any
  Delete "$SMPROGRAMS\PYNBM\*.*"

  ; Remove directories used
  RMDir "$SMPROGRAMS\PYNBM"
  
  Delete "$APPDATA\PYNBM\*.*"
  RMDir "$APPDATA\PYNBM"
  
  DeleteRegKey HKCU "Software\PYNBM"
  DeleteRegKey HKCU "Software\Microsoft\Windows\CurrentVersion\Uninstall\PYNBM"
SectionEnd