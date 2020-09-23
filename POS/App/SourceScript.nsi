!define APP_NAME "PBSP POS v1.0"
!define COMP_NAME "PBSP System Development"
!define VERSION "1.0.2.0"
!define COPYRIGHT "PBSP Dev. © 2010-2011"
!define DESCRIPTION "POS System Setup"
!define LICENSE_TXT "G:\subjects\Prog2\Projects\final\POS\App\License.txt"
!define INSTALLER_NAME "G:\subjects\Prog2\Projects\final\POS\App\Setup.exe"
!define MAIN_APP_EXE "POS.exe"
!define INSTALL_TYPE "SetShellVarContext all"
!define REG_ROOT "HKLM"
!define REG_APP_PATH "Software\Microsoft\Windows\CurrentVersion\App Paths\${MAIN_APP_EXE}"
!define UNINSTALL_PATH "Software\Microsoft\Windows\CurrentVersion\Uninstall\${APP_NAME}"

######################################################################

VIProductVersion  "${VERSION}"
VIAddVersionKey "ProductName"  "${APP_NAME}"
VIAddVersionKey "CompanyName"  "${COMP_NAME}"
VIAddVersionKey "LegalCopyright"  "${COPYRIGHT}"
VIAddVersionKey "FileDescription"  "${DESCRIPTION}"
VIAddVersionKey "FileVersion"  "${VERSION}"

######################################################################

SetCompressor ZLIB
Name "${APP_NAME}"
Caption "${APP_NAME}"
OutFile "${INSTALLER_NAME}"
BrandingText "${APP_NAME}"
XPStyle on
InstallDirRegKey "${REG_ROOT}" "${REG_APP_PATH}" ""
InstallDir "$PROGRAMFILES\POS"

######################################################################

!include "MUI.nsh"
!include "MUI2.nsh"

!define MUI_HEADERIMAGE
!define MUI_HEADERIMAGE_BITMAP "G:\subjects\Prog2\Projects\final\POS\App\win.bmp"

!define MUI_ABORTWARNING
!define MUI_UNABORTWARNING
!define MUI_ICON "G:\subjects\Prog2\Projects\final\POS\App\Install.ico"
!define MUI_UNINSTALLICON "G:\subjects\Prog2\Projects\final\POS\App\UnInstall.ico"
######################################################################
  ;Remember the installer language
  !define MUI_LANGDLL_REGISTRY_ROOT "HKCU" 
  !define MUI_LANGDLL_REGISTRY_KEY "Software\POS" 
  !define MUI_LANGDLL_REGISTRY_VALUENAME "Installer Language"
######################################################################

!insertmacro MUI_PAGE_WELCOME

!ifdef LICENSE_TXT
!insertmacro MUI_PAGE_LICENSE "${LICENSE_TXT}"
!endif

!insertmacro MUI_PAGE_COMPONENTS

!insertmacro MUI_PAGE_DIRECTORY

!ifdef REG_START_MENU
!define MUI_STARTMENUPAGE_NODISABLE
!define MUI_STARTMENUPAGE_DEFAULTFOLDER "POS"
!define MUI_STARTMENUPAGE_REGISTRY_ROOT "${REG_ROOT}"
!define MUI_STARTMENUPAGE_REGISTRY_KEY "${UNINSTALL_PATH}"
!define MUI_STARTMENUPAGE_REGISTRY_VALUENAME "${REG_START_MENU}"
!insertmacro MUI_PAGE_STARTMENU Application $SM_Folder
!endif

!insertmacro MUI_PAGE_INSTFILES

!define MUI_FINISHPAGE_RUN "$INSTDIR\${MAIN_APP_EXE}"
!define MUI_FINISHPAGE_SHOWREADME "$INSTDIR\Readme.txt"
!insertmacro MUI_PAGE_FINISH

!insertmacro MUI_UNPAGE_CONFIRM

!insertmacro MUI_UNPAGE_INSTFILES



#####################################################################################

  !insertmacro MUI_LANGUAGE "English" ;first language is the default language
  !insertmacro MUI_LANGUAGE "French"
  !insertmacro MUI_LANGUAGE "German"
  !insertmacro MUI_LANGUAGE "Spanish"
  !insertmacro MUI_LANGUAGE "SpanishInternational"
  !insertmacro MUI_LANGUAGE "SimpChinese"
  !insertmacro MUI_LANGUAGE "TradChinese"
  !insertmacro MUI_LANGUAGE "Japanese"
  !insertmacro MUI_LANGUAGE "Korean"
  !insertmacro MUI_LANGUAGE "Italian"
  !insertmacro MUI_LANGUAGE "Dutch"
  !insertmacro MUI_LANGUAGE "Danish"
  !insertmacro MUI_LANGUAGE "Swedish"
  !insertmacro MUI_LANGUAGE "Norwegian"
  !insertmacro MUI_LANGUAGE "NorwegianNynorsk"
  !insertmacro MUI_LANGUAGE "Finnish"
  !insertmacro MUI_LANGUAGE "Greek"
  !insertmacro MUI_LANGUAGE "Russian"
  !insertmacro MUI_LANGUAGE "Portuguese"
  !insertmacro MUI_LANGUAGE "PortugueseBR"
  !insertmacro MUI_LANGUAGE "Polish"
  !insertmacro MUI_LANGUAGE "Ukrainian"
  !insertmacro MUI_LANGUAGE "Czech"
  !insertmacro MUI_LANGUAGE "Slovak"
  !insertmacro MUI_LANGUAGE "Croatian"
  !insertmacro MUI_LANGUAGE "Bulgarian"
  !insertmacro MUI_LANGUAGE "Hungarian"
  !insertmacro MUI_LANGUAGE "Thai"
  !insertmacro MUI_LANGUAGE "Romanian"
  !insertmacro MUI_LANGUAGE "Latvian"
  !insertmacro MUI_LANGUAGE "Macedonian"
  !insertmacro MUI_LANGUAGE "Estonian"
  !insertmacro MUI_LANGUAGE "Turkish"
  !insertmacro MUI_LANGUAGE "Lithuanian"
  !insertmacro MUI_LANGUAGE "Slovenian"
  !insertmacro MUI_LANGUAGE "Serbian"
  !insertmacro MUI_LANGUAGE "SerbianLatin"
  !insertmacro MUI_LANGUAGE "Arabic"
  !insertmacro MUI_LANGUAGE "Farsi"
  !insertmacro MUI_LANGUAGE "Hebrew"
  !insertmacro MUI_LANGUAGE "Indonesian"
  !insertmacro MUI_LANGUAGE "Mongolian"
  !insertmacro MUI_LANGUAGE "Luxembourgish"
  !insertmacro MUI_LANGUAGE "Albanian"
  !insertmacro MUI_LANGUAGE "Breton"
  !insertmacro MUI_LANGUAGE "Belarusian"
  !insertmacro MUI_LANGUAGE "Icelandic"
  !insertmacro MUI_LANGUAGE "Malay"
  !insertmacro MUI_LANGUAGE "Bosnian"
  !insertmacro MUI_LANGUAGE "Kurdish"
  !insertmacro MUI_LANGUAGE "Irish"
  !insertmacro MUI_LANGUAGE "Uzbek"
  !insertmacro MUI_LANGUAGE "Galician"
  !insertmacro MUI_LANGUAGE "Afrikaans"
  !insertmacro MUI_LANGUAGE "Catalan"
  !insertmacro MUI_LANGUAGE "Esperanto"

############################################################################

;Installer Functions

Function .onInit

  !insertmacro MUI_LANGDLL_DISPLAY

FunctionEnd

#############################################################################

section "Graphics" SecImg
SetOutPath "$INSTDIR\Images\System\Icons"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\accept.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\add.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\application_add.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\application_home.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\application_key.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\cart.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\clipboard_text.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\cog.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\coins.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\cursor.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\close.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\delivery.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\exclamation_octagon_fram.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\expenses.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\find copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\help.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\house.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\layout.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\lock.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\money.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\cancel.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\package.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\page_white_delete.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\page_white_edit.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\page_white_find.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\page_white_gear.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\pencil.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\pill.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\printer.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\reportdelivery.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\reportexpense.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\reports.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\reportsales.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\reporttrans.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\rptdelivery.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\rptinvntory.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\rptsales.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\sched.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\stocks.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\supprofile.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\transact.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\user_business_boss.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\user_female.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\arrow_refresh copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\warnings.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\application_view_list copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\database.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\calculator.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\Case Info.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\client_info.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\Clients.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\disk.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\edit.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\exclamation.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\notes.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\bin_closed copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\form_icon.ico"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\lock_unlock.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\Log-out.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\magnifier copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\new.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\page_copy.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\Perpetrator.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\PI Info.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\report.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\Trial Info.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\arrow_undo.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\cashier.ico"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Icons\drink.bmp"
SetOutPath "$INSTDIR\Images\System\MDI Icons"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\User Accounts.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Settings.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Delete Document.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Coinstack-32x32.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Text-Bubble-32x32.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Search.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Inventory.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Ok.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Suppliers.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Administrator.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Add Stocks.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Users.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Stocks.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Exit.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Reports.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Locate.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Files.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Find.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Home.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Cashier.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Sales.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Sales report.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Warning.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\About.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Help.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Attach.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\MDI Icons\Delete-32x32.bmp"
SetOutPath "$INSTDIR\Images\System\Background"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\System\Background\SPLASH.bmp"
SetOutPath "$INSTDIR\Images\Accounts"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Accounts\1_160585304l.jpg"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Accounts\april.jpg"
SetOutPath "$INSTDIR\Images\Products"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Products\sample.bmp"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Products\binary-addition.gif"
SetOutPath "$INSTDIR\Images\Suppliers"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Suppliers\Sunset.jpg"
File "G:\subjects\Prog2\Projects\final\POS\App\Images\Suppliers\IMG_0204.JPG"
SectionEnd

######################################################################
;Descriptions
  !insertmacro MUI_FUNCTION_DESCRIPTION_BEGIN
    !insertmacro MUI_DESCRIPTION_TEXT ${SecImg} "Include System Graphics on installation"
  !insertmacro MUI_FUNCTION_DESCRIPTION_END
 
#######################################################################

Section -MainProgram
${INSTALL_TYPE}
SetOverwrite ifnewer
SetOutPath "$INSTDIR"
File "G:\subjects\Prog2\Projects\final\POS\App\POS.exe"
File "G:\subjects\Prog2\Projects\final\POS\App\POS.mdb"
File "G:\subjects\Prog2\Projects\final\POS\App\Settings.ini"
File "G:\subjects\Prog2\Projects\final\POS\App\POS.exe.manifest"
File "G:\subjects\Prog2\Projects\final\POS\App\help.chm"
File "G:\subjects\Prog2\Projects\final\POS\App\Notes.log"
File "G:\subjects\Prog2\Projects\final\POS\App\James.acs"
File "G:\subjects\Prog2\Projects\final\POS\App\License.txt"
File "G:\subjects\Prog2\Projects\final\POS\App\Readme.txt"
File "G:\subjects\Prog2\Projects\final\POS\App\cPopMenu6.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\CtrlLine.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\ctrlButton.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\SSubTmr6.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Run.bat"

RegDll "$INSTDIR\cPopMenu6.ocx"
RegDll "$INSTDIR\CtrlLine.ocx"
RegDll "$INSTDIR\SSubTmr6.dll"
RegDll "$INSTDIR\ctrlButton.ocx"

SetOutPath "$INSTDIR\Images\Accounts"
SetOutPath "$INSTDIR\Images\Products"
SetOutPath "$INSTDIR\Images\Suppliers"
SetOutPath "$INSTDIR\Backup"
SetOutPath "$INSTDIR\Components"

File "G:\subjects\Prog2\Projects\final\POS\App\Components\Comdlg32.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\asycfilt.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\COMCTL32.OCX"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\msado15.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\MSBIND.DLL"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\MSCOMCT2.OCX"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\mscomctl.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\MSDBRPT.DLL"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\MSDBRPTR.DLL"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\msderun.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\msjro.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\msmask32.ocx"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\MSSTDFMT.DLL"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\msvbvm60.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\oleaut32.dll"
File "G:\subjects\Prog2\Projects\final\POS\App\Components\olepro32.dll"

RegDll "$INSTDIR\Components\Comdlg32.ocx"
RegDll "$INSTDIR\Components\asycfilt.dll"
RegDll "$INSTDIR\Components\COMCTL32.OCX"
RegDll "$INSTDIR\Components\msado15.dll"
RegDll "$INSTDIR\Components\MSBIND.DLL"
RegDll "$INSTDIR\Components\MSCOMCT2.OCX"
RegDll "$INSTDIR\Components\mscomctl.ocx"
RegDll "$INSTDIR\Components\MSDBRPT.DLL"
RegDll "$INSTDIR\Components\MSDBRPTR.DLL"
RegDll "$INSTDIR\Components\msderun.dll"
RegDll "$INSTDIR\Components\msjro.dll"
RegDll "$INSTDIR\Components\msmask32.ocx"
RegDll "$INSTDIR\Components\MSSTDFMT.DLL"
RegDll "$INSTDIR\Components\msvbvm60.dll"
RegDll "$INSTDIR\Components\oleaut32.dll"
RegDll "$INSTDIR\Components\olepro32.dll"

execshell "open" "$INSTDIR\Run.bat"

SectionEnd

######################################################################

Section -Icons_Reg
SetOutPath "$INSTDIR"
WriteUninstaller "$INSTDIR\uninstall.exe"

!ifdef REG_START_MENU
!insertmacro MUI_STARTMENU_WRITE_BEGIN Application
CreateDirectory "$SMPROGRAMS\$SM_Folder"
CreateShortCut "$SMPROGRAMS\$SM_Folder\${APP_NAME}.lnk" "$INSTDIR\${MAIN_APP_EXE}"
CreateShortCut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${MAIN_APP_EXE}"
!ifdef WEB_SITE
WriteIniStr "$INSTDIR\${APP_NAME} website.url" "InternetShortcut" "URL" "${WEB_SITE}"
CreateShortCut "$SMPROGRAMS\$SM_Folder\${APP_NAME} Website.lnk" "$INSTDIR\${APP_NAME} website.url"
!endif
!insertmacro MUI_STARTMENU_WRITE_END
!endif

!ifndef REG_START_MENU
CreateDirectory "$SMPROGRAMS\POS"
CreateShortCut "$SMPROGRAMS\POS\${APP_NAME}.lnk" "$INSTDIR\${MAIN_APP_EXE}"
CreateShortCut "$DESKTOP\${APP_NAME}.lnk" "$INSTDIR\${MAIN_APP_EXE}"
!ifdef WEB_SITE
WriteIniStr "$INSTDIR\${APP_NAME} website.url" "InternetShortcut" "URL" "${WEB_SITE}"
CreateShortCut "$SMPROGRAMS\POS\${APP_NAME} Website.lnk" "$INSTDIR\${APP_NAME} website.url"
!endif
!endif

WriteRegStr ${REG_ROOT} "${REG_APP_PATH}" "" "$INSTDIR\${MAIN_APP_EXE}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "DisplayName" "${APP_NAME}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "UninstallString" "$INSTDIR\uninstall.exe"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "DisplayIcon" "$INSTDIR\${MAIN_APP_EXE}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "DisplayVersion" "${VERSION}"
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "Publisher" "${COMP_NAME}"

!ifdef WEB_SITE
WriteRegStr ${REG_ROOT} "${UNINSTALL_PATH}"  "URLInfoAbout" "${WEB_SITE}"
!endif
SectionEnd

######################################################################

Section Uninstall
${INSTALL_TYPE}
Delete "$INSTDIR\${MAIN_APP_EXE}"
Delete "$INSTDIR\POS.mdb"
Delete "$INSTDIR\RegCtrl.bat"
Delete "$INSTDIR\Settings.ini"
Delete "$INSTDIR\SSubTmr6.dll"
Delete "$INSTDIR\POS.exe.manifest"
Delete "$INSTDIR\cPopMenu6.ocx"
Delete "$INSTDIR\CtrlLine.ocx"
Delete "$INSTDIR\ctrlButton.ocx"
Delete "$INSTDIR\help.chm"
Delete "$INSTDIR\Notes.log"
Delete "$INSTDIR\SourceScript.nsi"
Delete "$INSTDIR\James.acs"
Delete "$INSTDIR\License.txt"
Delete "$INSTDIR\Readme.txt"
Delete "$INSTDIR\Uninstall.ico"
Delete "$INSTDIR\win.bmp"
Delete "$INSTDIR\Install.ico"
Delete "$INSTDIR\Images\System\Icons\accept.bmp"
Delete "$INSTDIR\Images\System\Icons\add.bmp"
Delete "$INSTDIR\Images\System\Icons\application_add.bmp"
Delete "$INSTDIR\Images\System\Icons\application_home.bmp"
Delete "$INSTDIR\Images\System\Icons\application_key.bmp"
Delete "$INSTDIR\Images\System\Icons\cart.bmp"
Delete "$INSTDIR\Images\System\Icons\clipboard_text.bmp"
Delete "$INSTDIR\Images\System\Icons\cog.bmp"
Delete "$INSTDIR\Images\System\Icons\coins.bmp"
Delete "$INSTDIR\Images\System\Icons\cursor.bmp"
Delete "$INSTDIR\Images\System\Icons\close.bmp"
Delete "$INSTDIR\Images\System\Icons\delivery.bmp"
Delete "$INSTDIR\Images\System\Icons\exclamation_octagon_fram.bmp"
Delete "$INSTDIR\Images\System\Icons\expenses.bmp"
Delete "$INSTDIR\Images\System\Icons\find copy.bmp"
Delete "$INSTDIR\Images\System\Icons\help.bmp"
Delete "$INSTDIR\Images\System\Icons\house.bmp"
Delete "$INSTDIR\Images\System\Icons\layout.bmp"
Delete "$INSTDIR\Images\System\Icons\lock.bmp"
Delete "$INSTDIR\Images\System\Icons\money.bmp"
Delete "$INSTDIR\Images\System\Icons\cancel.bmp"
Delete "$INSTDIR\Images\System\Icons\package.bmp"
Delete "$INSTDIR\Images\System\Icons\page_white_delete.bmp"
Delete "$INSTDIR\Images\System\Icons\page_white_edit.bmp"
Delete "$INSTDIR\Images\System\Icons\page_white_find.bmp"
Delete "$INSTDIR\Images\System\Icons\page_white_gear.bmp"
Delete "$INSTDIR\Images\System\Icons\pencil.bmp"
Delete "$INSTDIR\Images\System\Icons\pill.bmp"
Delete "$INSTDIR\Images\System\Icons\printer.bmp"
Delete "$INSTDIR\Images\System\Icons\reportdelivery.bmp"
Delete "$INSTDIR\Images\System\Icons\reportexpense.bmp"
Delete "$INSTDIR\Images\System\Icons\reports.bmp"
Delete "$INSTDIR\Images\System\Icons\reportsales.bmp"
Delete "$INSTDIR\Images\System\Icons\reporttrans.bmp"
Delete "$INSTDIR\Images\System\Icons\rptdelivery.bmp"
Delete "$INSTDIR\Images\System\Icons\rptinvntory.bmp"
Delete "$INSTDIR\Images\System\Icons\rptsales.bmp"
Delete "$INSTDIR\Images\System\Icons\sched.bmp"
Delete "$INSTDIR\Images\System\Icons\stocks.bmp"
Delete "$INSTDIR\Images\System\Icons\supprofile.bmp"
Delete "$INSTDIR\Images\System\Icons\transact.bmp"
Delete "$INSTDIR\Images\System\Icons\user_business_boss.bmp"
Delete "$INSTDIR\Images\System\Icons\user_female.bmp"
Delete "$INSTDIR\Images\System\Icons\arrow_refresh copy.bmp"
Delete "$INSTDIR\Images\System\Icons\warnings.bmp"
Delete "$INSTDIR\Images\System\Icons\application_view_list copy.bmp"
Delete "$INSTDIR\Images\System\Icons\database.bmp"
Delete "$INSTDIR\Images\System\Icons\calculator.bmp"
Delete "$INSTDIR\Images\System\Icons\Case Info.bmp"
Delete "$INSTDIR\Images\System\Icons\client_info.bmp"
Delete "$INSTDIR\Images\System\Icons\Clients.bmp"
Delete "$INSTDIR\Images\System\Icons\disk.bmp"
Delete "$INSTDIR\Images\System\Icons\edit.bmp"
Delete "$INSTDIR\Images\System\Icons\exclamation.bmp"
Delete "$INSTDIR\Images\System\Icons\notes.bmp"
Delete "$INSTDIR\Images\System\Icons\bin_closed copy.bmp"
Delete "$INSTDIR\Images\System\Icons\form_icon.ico"
Delete "$INSTDIR\Images\System\Icons\lock_unlock.bmp"
Delete "$INSTDIR\Images\System\Icons\Log-out.bmp"
Delete "$INSTDIR\Images\System\Icons\magnifier copy.bmp"
Delete "$INSTDIR\Images\System\Icons\new.bmp"
Delete "$INSTDIR\Images\System\Icons\page_copy.bmp"
Delete "$INSTDIR\Images\System\Icons\Perpetrator.bmp"
Delete "$INSTDIR\Images\System\Icons\PI Info.bmp"
Delete "$INSTDIR\Images\System\Icons\report.bmp"
Delete "$INSTDIR\Images\System\Icons\Trial Info.bmp"
Delete "$INSTDIR\Images\System\Icons\arrow_undo.bmp"
Delete "$INSTDIR\Images\System\Icons\cashier.ico"
Delete "$INSTDIR\Images\System\Icons\drink.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\User Accounts.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Settings.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Delete Document.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Coinstack-32x32.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Text-Bubble-32x32.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Search.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Inventory.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Ok.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Suppliers.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Administrator.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Add Stocks.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Users.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Stocks.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Exit.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Reports.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Locate.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Files.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Find.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Home.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Cashier.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Sales.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Sales report.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Warning.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\About.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Help.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Attach.bmp"
Delete "$INSTDIR\Images\System\MDI Icons\Delete-32x32.bmp"
Delete "$INSTDIR\Images\System\Background\SPLASH.bmp"
Delete "$INSTDIR\Images\Accounts\1_160585304l.jpg"
Delete "$INSTDIR\Images\Accounts\april.jpg"
Delete "$INSTDIR\Images\Products\sample.bmp"
Delete "$INSTDIR\Images\Products\binary-addition.gif"
Delete "$INSTDIR\Images\Suppliers\Sunset.jpg"
Delete "$INSTDIR\Images\Suppliers\IMG_0204.JPG"

Delete "$INSTDIR\Components\Comdlg32.ocx"
Delete "$INSTDIR\Components\MSCOMCT2.OCX"
Delete "$INSTDIR\Components\mscomctl.ocx"
Delete "$INSTDIR\Components\MSDBRPT.DLL"
Delete "$INSTDIR\Components\msmask32.ocx"
Delete "$INSTDIR\Components\TABCTL32.OCX"
Delete "$INSTDIR\Components\SSubTmr6.dll"
Delete "$INSTDIR\Components\MSSTDFMT.DLL"
Delete "$INSTDIR\Components\COMCTL32.OCX"
Delete "$INSTDIR\Components\msado15.dll"
Delete "$INSTDIR\Components\MSDBRPTR.DLL"
Delete "$INSTDIR\Components\msjro.dll"
 
RmDir "$INSTDIR\Images\Suppliers"
RmDir "$INSTDIR\Images\Products"
RmDir "$INSTDIR\Images\Accounts"
RmDir "$INSTDIR\Images\System\Background"
RmDir "$INSTDIR\Images\System\MDI Icons"
RmDir "$INSTDIR\Images\System\Icons"
 
Delete "$INSTDIR\uninstall.exe"
!ifdef WEB_SITE
Delete "$INSTDIR\${APP_NAME} website.url"
!endif

RmDir "$INSTDIR"

!ifdef REG_START_MENU
!insertmacro MUI_STARTMENU_GETFOLDER "Application" $SM_Folder
Delete "$SMPROGRAMS\$SM_Folder\${APP_NAME}.lnk"
!ifdef WEB_SITE
Delete "$SMPROGRAMS\$SM_Folder\${APP_NAME} Website.lnk"
!endif
Delete "$DESKTOP\${APP_NAME}.lnk"

RmDir "$SMPROGRAMS\$SM_Folder"
!endif

!ifndef REG_START_MENU
Delete "$SMPROGRAMS\POS\${APP_NAME}.lnk"
!ifdef WEB_SITE
Delete "$SMPROGRAMS\POS\${APP_NAME} Website.lnk"
!endif
Delete "$DESKTOP\${APP_NAME}.lnk"

RmDir "$SMPROGRAMS\POS"
!endif

DeleteRegKey ${REG_ROOT} "${REG_APP_PATH}"
DeleteRegKey ${REG_ROOT} "${UNINSTALL_PATH}"
SectionEnd

######################################################################

