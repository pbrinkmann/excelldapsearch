;Excel LDAP Search
;Copyright (C) 2008 Paul Brinkmann <paul@paulb.org>
;
;This program is free software; you can redistribute it and/or modify it under the terms of the GNU General Public License as published by 'the Free Software Foundation; either version 2 of the License, r (at your option) any later version.
;
;This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details.
;
;You should have received a copy of the GNU General Public License along with this program; if not, write to the Free Software Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA 02111-1307 USA



;
; Excel LDAP Search install script
;

;--------------------------------

!include "install_lib.nsh"

SetCompressor /SOLID lzma

; The name of the installer
Name "Excel LDAP Search"

; Version #
VIProductVersion "0.0.5.2"
VIAddVersionKey /LANG=${LANG_ENGLISH} "ProductName" "Excel LDAP Search"
VIAddVersionKey /LANG=${LANG_ENGLISH} "Comments" "An Excel add-in to perform LDAP searches"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileDescription" "Excel LDAP Search installer"
VIAddVersionKey /LANG=${LANG_ENGLISH} "FileVersion" "0.0.5.2"

; The file to write
OutFile "ExcelLdapSearchInstall_0_52.exe"


; The default installation directory
;InstallDir "c:\source\excel ldap search\nsis installer\testdest"
InstallDir "$PROGRAMFILES\Excel LDAP Search"


InstallDirRegKey HKLM "Software\Excel LDAP Search" "installDir"

; Request application privileges for Windows Vista
RequestExecutionLevel user

; license agreement
LicenseData "..\distrib\license.txt"

;--------------------------------

; Pages

Page license
Page directory
Page instfiles
UninstPage uninstConfirm
UninstPage instfiles

;--------------------------------

; The stuff to install
Section "Program Install" ;No components page, name is not important

fileInUseCheck:
    ;
    ; First check for locked files (the two DLLs and the macro file)
    ;
	!insertmacro IfFileInUse "$INSTDIR\ldapquery_addin.xla" fileInUse 0
    !insertmacro IfFileInUse "$INSTDIR\LdapQuery.dll" fileInUse 0
    !insertmacro IfFileInUse "$INSTDIR\nsldap32v50.dll" fileInUse noFileInUse

fileInUse:
    DetailPrint "Detected files in use by another process"
    MessageBox MB_ICONSTOP|MB_RETRYCANCEL "It appears some of the files are in use.  If you're running Excel, please exit now.  Click 'retry' to try again or 'cancel' to abort the installation" IDRETRY fileInUseCheck IDCANCEL 0
    DetailPrint "User aborted install"
    Abort 
noFileInUse:

  ;
  ; Set output path to the installation directory.
  ;
  SetOutPath $INSTDIR
  
  ;
  ; save our installed location
  ;
  WriteRegStr HKLM "Software\Excel LDAP Search" "installDir" $INSTDIR
  
  ;
  ; see if there's an existing "ldap_params.ini" file and rename it if there is
  ;
  IfFileExists "$INSTDIR\ldap_params.ini" 0 ldapParamsRenameEnd
    IntOp $R0  0 + 0  ; this is dumb, but is there a better way of "$R0 = 0"?

    ; this'll find a number that doesn't exist    
backupLoop:
    IfFileExists "$INSTDIR\ldap_params_OLD$R0.ini" 0 backupLoopEnd
        IntOp $R0 $R0 + 1
        Goto backupLoop
backupLoopEnd: 

    Rename "$INSTDIR\ldap_params.ini" "$INSTDIR\ldap_params_OLD$R0.ini"
    MessageBox MB_OK "Your existing ldap_params.ini file was renamed with a '_OLD$R0' suffix"    
 
 ldapParamsRenameEnd:

  ;
  ; grab the files in the "distrib" directory
  ;
  File /r /x .svn /x install_script.vbs /x web /x copyfiles.bat ..\distrib\*
  
  IfErrors 0 afterFileErrors
	MessageBox MB_OK "File errors ohmy"
  
  
  
 afterFileErrors:
  
  ;
  ; introduce our LDAP-searching COM object
  ;
  RegDLL $INSTDIR\LdapQuery.dll
  
  ;
  ; install the Excel menu item
  ;
  ExecWait 'cscript "$INSTDIR\install_excel_menuitem.vbs"' $0
  
  IntCmp $0 0 menuRegOK
  DetailPrint "Installing the Excel menu item failed"
  Abort
    
menuRegOK:
  
  ;
  ; have the user customize their directory server parameters
  ;
  MessageBox MB_OK "Please change the connection parameters to match your needs"
  ExecShell "Open" $INSTDIR\ldap_params.ini
  
  ;
  ; Write the uninstall keys for Windows
  ;
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel LDAP Search" "DisplayName" "Excel LDAP Search"
  WriteRegStr HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel LDAP Search" "UninstallString" '"$INSTDIR\uninstall.exe"'
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel LDAP Search" "NoModify" 1
  WriteRegDWORD HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel LDAP Search" "NoRepair" 1
  
  ;
  ; and finally, create an uninstaller
  ;
  WriteUninstaller "uninstall.exe"
  
SectionEnd ; end the section

;--------------------------------

; Uninstaller

Section "Uninstall"
  
fileInUseCheckB:
    ;
    ; First check for locked files (the two DLLs and the macro file)
    ;
    !insertmacro IfFileInUse "$INSTDIR\macro_template.xlt" fileInUseB 0   ; left in for legacy purposes
	!insertmacro IfFileInUse "$INSTDIR\ldapquery_addin.xla" fileInUseB 0
    !insertmacro IfFileInUse "$INSTDIR\LdapQuery.dll" fileInUseB 0
    !insertmacro IfFileInUse "$INSTDIR\nsldap32v50.dll" fileInUseB noFileInUseB

fileInUseB:
    DetailPrint "Detected files in use by another process"
    MessageBox MB_ICONSTOP|MB_ABORTRETRYIGNORE "It appears some of the files are in use.  If you're running Excel, please exit now.  Click 'abort' to abort the uninstall, click 'retry' to try again or 'ignore' to continue the uninstall and require a reboot" IDIGNORE noFileInUseB IDRETRY fileInUseCheckB
    DetailPrint "User aborted uninstall"
    Abort 
noFileInUseB:
  
  ;These will be locked if Excel is open and a query has been run
  Delete /REBOOTOK $INSTDIR\LdapQuery.dll
  Delete /REBOOTOK $INSTDIR\macro_template.xlt	; left in for legacy purposes
  Delete /REBOOTOK $INSTDIR\ldapquery_addin.xla
  Delete /REBOOTOK $INSTDIR\nsldap32v50.dll
  Delete /REBOOTOK $INSTDIR\msvcp71.dll
  Delete /REBOOTOK $INSTDIR\msvcr71.dll
  
    
  
  ; remove the Excel menu item
  ExecWait 'cscript "$INSTDIR\remove_excel_menuitem.vbs"'
  
  ; Remove registry keys
  DeleteRegKey HKLM "Software\Excel LDAP Search"
  DeleteRegKey HKLM "Software\Microsoft\Windows\CurrentVersion\Uninstall\Excel LDAP Search"

  ; unregister our dll
  UnRegDLL $INSTDIR\LdapQuery.dll

  ; manually delete all the files 
  Delete $INSTDIR\readme_images\column_offset.png
  Delete $INSTDIR\readme_images\data_range.png
  Delete $INSTDIR\readme_images\per_row_results.png
  Delete $INSTDIR\readme_images\query_string.png
  Delete $INSTDIR\readme_images\return_attributes.png
  Delete $INSTDIR\readme_images\return_attributes_freeform.png
  Delete $INSTDIR\readme_images\search_attribute.png
  Delete $INSTDIR\readme_images\blank_demo_thumb.png
  Delete $INSTDIR\readme_images\existing_demo_thumb.png
  Delete $INSTDIR\install_excel_menuitem.vbs
  Delete $INSTDIR\remove_excel_menuitem.vbs
  Delete $INSTDIR\ldap_params.ini
  Delete $INSTDIR\license.txt
  Delete $INSTDIR\COPYING
  Delete $INSTDIR\msvcp71.dll
  Delete $INSTDIR\msvcr71.dll
  Delete $INSTDIR\readme.html
  Delete $INSTDIR\TODO.txt
  Delete $INSTDIR\uninstall.exe
    
  ; Remove directories used
  RMDir $INSTDIR\readme_images
  RMDir "$INSTDIR"
    
    
    IfRebootFlag 0 uninstallEnd
        MessageBox MB_YESNO|MB_ICONQUESTION "A reboot is required to complete the uninstall.  Do you wish to reboot now?" IDYES 0 IDNO uninstallEnd
        Reboot
         
    
uninstallEnd:

SectionEnd


