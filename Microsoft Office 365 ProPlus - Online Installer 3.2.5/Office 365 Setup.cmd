:: http://zone94.com
@ECHO OFF
TITLE Microsoft Office 365 ProPlus - Online Installer
PushD "%~dp0"
Set "Version=3.2.5"
CHCP 1252 1>NUL
Set "BatchDir=%~dp0"
If [%BatchDir:~-1%]==[\] Set "BatchDir=%BatchDir:~0,-1%"

Set "on=(YES)"
Set "off=(NO) "
Set "opt1=%on%"
Set "opt2=%on%"
Set "opt3=%on%"
Set "opt4=%on%"
Set "opt5=%on%"
Set "opt6=%on%"
Set "opt7=%on%"
Set "opt8=%on%"
Set "opt9=%on%"
Set "optP=%on%"
Set "optT=%on%"
Set "optD=%on%"
If "%Processor_Architecture%"=="x86" Set "optB=32-bit"
If "%Processor_Architecture%"=="AMD64" Set "optB=64-bit"
If "%Processor_Architecture%"=="ARM64" Set "optB=64-bit"
Set "optO=%off%"
Set "optS=%on%"

:MENU
CLS
ECHO                                    \\!//
ECHO                                    (o o)
ECHO -------------------------------oOOo-(_)-oOOo-------------------------------
ECHO      Microsoft Office 365 Professional Plus - Online Installer v%Version%
ECHO      INFO: After the Installation run the 'KMS Activation Script' from
ECHO     http://zone94.com/downloads/135-windows-and-office-activation-script
ECHO ===========================================================================
FLTMC >NUL 2>&1 || Set _=^"Set "_ELEV=1" ^& CD /d """%~dp0""" ^& "%~f0" %*^" && ((If "%_ELEV%"=="" (ECHO. & ECHO Requesting administrator privileges . . . & ((PowerShell -nop -c START "" CMD -args '/d/x/s/v:off/r',$env:_ -Verb RunAs >NUL 2>NUL) || (mshta vbscript:execute^("CreateObject(""Shell.Application"").ShellExecute(""CMD"",""/d/x/s/v:off/r"" & CreateObject(""WScript.Shell"").Environment(""PROCESS"")(""_""),,""RunAs"",1)(Window.Close)"^) >NUL 2>&1))) Else (ECHO. & ECHO This script requires administrator privileges. & PAUSE)) & EXIT /b)

WHERE /q /r ".\Office\Data" "stream.*.x-none.dat" 2>NUL && (Set "OFiles=%on%") || (Set "OFiles=%off%")
If "%optO%,%OFiles%"=="%on%,%on%" Set "OMessage=Delete Offline Files"
If "%optS%,%OFiles%"=="%on%,%on%" Set "SMessage=Offline Installation"
If "%optO%,%OFiles%"=="%off%,%off%" Set "OMessage=Download Offline Files (x32 ^& x64 = 6,40 GB)"
If "%optS%,%OFiles%"=="%off%,%off%" Set "SMessage=Online Installation"
If "%optO%,%OFiles%"=="%on%,%off%" Set "OMessage=Download Offline Files (x32 ^& x64 = 6,40 GB)"
If "%optS%,%OFiles%"=="%on%,%off%" Set "SMessage=Online Installation"
If "%optO%,%OFiles%"=="%off%,%on%" Set "OMessage=Delete Offline Files"
If "%optS%,%OFiles%"=="%off%,%on%" Set "SMessage=Offline Installation"
If "%optB%"=="64-bit" Set "BMessage=Here you can switch between a 32-bit or 64-bit installation"
If "%optB%"=="32-bit" Set "BMessage=Here you can switch between a 32-bit or 64-bit installation"
If "%optB%,%Processor_Architecture%"=="64-bit,x86" Set "optB=32-bit" & Set "BMessage=A 64-bit Application cannot be installed on a 32-bit Windows"

ECHO.
<NUL Set/P=[1] & (If "%opt1%"=="%on%" (Call :EchoColor "%opt1%" 0a) Else (<NUL Set/P="%opt1%")) & ECHO  Microsoft Office Word.
<NUL Set/P=[2] & (If "%opt2%"=="%on%" (Call :EchoColor "%opt2%" 0a) Else (<NUL Set/P="%opt2%")) & ECHO  Microsoft Office Excel.
<NUL Set/P=[3] & (If "%opt3%"=="%on%" (Call :EchoColor "%opt3%" 0a) Else (<NUL Set/P="%opt3%")) & ECHO  Microsoft Office PowerPoint.
<NUL Set/P=[4] & (If "%opt4%"=="%on%" (Call :EchoColor "%opt4%" 0a) Else (<NUL Set/P="%opt4%")) & ECHO  Microsoft Office Outlook.
<NUL Set/P=[5] & (If "%opt5%"=="%on%" (Call :EchoColor "%opt5%" 0a) Else (<NUL Set/P="%opt5%")) & ECHO  Microsoft Office OneNote.
<NUL Set/P=[6] & (If "%opt6%"=="%on%" (Call :EchoColor "%opt6%" 0a) Else (<NUL Set/P="%opt6%")) & ECHO  Microsoft Office Publisher
<NUL Set/P=[7] & (If "%opt7%"=="%on%" (Call :EchoColor "%opt7%" 0a) Else (<NUL Set/P="%opt7%")) & ECHO  Microsoft Office Access.
<NUL Set/P=[8] & (If "%opt8%"=="%on%" (Call :EchoColor "%opt8%" 0a) Else (<NUL Set/P="%opt8%")) & ECHO  Microsoft Office Visio.
<NUL Set/P=[9] & (If "%opt9%"=="%on%" (Call :EchoColor "%opt9%" 0a) Else (<NUL Set/P="%opt9%")) & ECHO  Microsoft Office Project.
<NUL Set/P=[P] & (If "%optP%"=="%on%" (Call :EchoColor "%optP%" 0a) Else (<NUL Set/P="%optP%")) & ECHO  Microsoft Office Proofing Tools.
<NUL Set/P=[T] & (If "%optT%"=="%on%" (Call :EchoColor "%optT%" 0a) Else (<NUL Set/P="%optT%")) & ECHO  Microsoft Teams.
<NUL Set/P=[D] & (If "%optD%"=="%on%" (Call :EchoColor "%optD%" 0a) Else (<NUL Set/P="%optD%")) & ECHO  Microsoft OneDrive.
ECHO --- Office will matches it's language with your Windows system language.
<NUL Set/P=[B] & (If "%optB%"=="%optB%" (Call :EchoColor "(%optB%)" 06) Else (<NUL Set/P="%optB%")) & ECHO  %BMessage%.
<NUL Set/P=[O] & (If "%optO%"=="%on%" (Call :EchoColor "%optO%" 09) Else (<NUL Set/P="%optO%")) & ECHO  %OMessage%.
<NUL Set/P=[S] & (If "%optS%"=="%on%" (Call :EchoColor "%optS%" 09) Else (<NUL Set/P="%optS%")) & ECHO  %SMessage%.

ECHO.
CHOICE /c 123456789PTDBOSXZ /n /t 300 /d X /m "--> Toggle your option(s) and toggle [Z] to Start: "

If ERRORLEVEL 17 GoTo CONTINUE
If ERRORLEVEL 16 EXIT
If ERRORLEVEL 15 (If "%optS%"=="%on%" (Set "optS=%off%" & Set "optO=%on%") Else (Set "optS=%on%" & Set "optO=%off%")) & GoTo MENU
If ERRORLEVEL 14 (If "%optO%"=="%on%" (Set "optO=%off%" & Set "optS=%on%") Else (Set "optO=%on%" & Set "optS=%off%")) & GoTo MENU
If ERRORLEVEL 13 (If "%optB%"=="64-bit" (Set "optB=32-bit") Else (Set "optB=64-bit")) & GoTo MENU
If ERRORLEVEL 12 (If "%optD%"=="%on%" (Set "optD=%off%") Else (Set "optD=%on%")) & GoTo MENU
If ERRORLEVEL 11 (If "%optT%"=="%on%" (Set "optT=%off%") Else (Set "optT=%on%")) & GoTo MENU
If ERRORLEVEL 10 (If "%optP%"=="%on%" (Set "optP=%off%") Else (Set "optP=%on%")) & GoTo MENU
If ERRORLEVEL 9 (If "%opt9%"=="%on%" (Set "opt9=%off%") Else (Set "opt9=%on%")) & GoTo MENU
If ERRORLEVEL 8 (If "%opt8%"=="%on%" (Set "opt8=%off%") Else (Set "opt8=%on%")) & GoTo MENU
If ERRORLEVEL 7 (If "%opt7%"=="%on%" (Set "opt7=%off%") Else (Set "opt7=%on%")) & GoTo MENU
If ERRORLEVEL 6 (If "%opt6%"=="%on%" (Set "opt6=%off%") Else (Set "opt6=%on%")) & GoTo MENU
If ERRORLEVEL 5 (If "%opt5%"=="%on%" (Set "opt5=%off%") Else (Set "opt5=%on%")) & GoTo MENU
If ERRORLEVEL 4 (If "%opt4%"=="%on%" (Set "opt4=%off%") Else (Set "opt4=%on%")) & GoTo MENU
If ERRORLEVEL 3 (If "%opt3%"=="%on%" (Set "opt3=%off%") Else (Set "opt3=%on%")) & GoTo MENU
If ERRORLEVEL 2 (If "%opt2%"=="%on%" (Set "opt2=%off%") Else (Set "opt2=%on%")) & GoTo MENU
If ERRORLEVEL 1 (If "%opt1%"=="%on%" (Set "opt1=%off%") Else (Set "opt1=%on%")) & GoTo MENU

:EchoColor (Text, Color)
  MkDir "%Temp%\_%1" 1>NUL
  PushD "%Temp%\_%1"
  For /f %%a in ('Echo PROMPT $H ^| "CMD"') do Set "bs=%%a"
  <NUL Set /P="_" >"%1"
  FindStr /l /i /b /p /a:%2 /s /c:"_" "%1"
  <NUL Set /P=%bs%%bs%
  PushD "%~dp0"
  RmDir /s /q "%Temp%\_%1"
GoTo :EOF

:CONTINUE
ECHO.
ECHO ---------------------------------------------------------------------------
Echo %opt1% %opt2% %opt3% %opt4% %opt5% %opt6% %opt7% %opt8% %opt9% %optP% %optT% %optD% | FindStr /l /i "%on%" >nul 2>&1 && GoTo SKIP
ECHO.
ECHO No option were selected. O_o ?
GoTo END
:SKIP


GoTo SKIP
:CONFIG
ECHO Creating Configuration File for Microsoft Office 365 ProPlus %CPU%-bit . . .
ping -n 3 localhost 1>NUL
SETLOCAL
Set "OCS=".\Office 365 Setup Config.xml""
                                     >%OCS% ECHO ^<Configuration^>
If "%optO%,%OFiles%"=="%on%,%off%"  >>%OCS% ECHO   ^<Add OfficeClientEdition="%CPU%" Channel="Current"^>
If "%optS%,%OFiles%"=="%on%,%off%"  >>%OCS% ECHO   ^<Add OfficeClientEdition="%CPU%" Channel="Current" MigrateArch="TRUE"^>
If "%optS%,%OFiles%"=="%on%,%on%"   >>%OCS% ECHO   ^<Add OfficeClientEdition="%CPU%" Channel="Current" MigrateArch="TRUE" SourcePath="%BatchDir%"^>
                                    >>%OCS% ECHO     ^<Product ID="O365ProPlusRetail"^>
                                    >>%OCS% ECHO       ^<Language ID="MatchOS" Fallback="en-US" /^>
If "%opt1%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Word" /^>
If "%opt2%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Excel" /^>
If "%opt3%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="PowerPoint" /^>
If "%opt4%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Outlook" /^>
If "%opt5%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="OneNote" /^>
If "%opt6%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Publisher" /^>
If "%opt7%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Access" /^>
If "%optT%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="Teams" /^>
If "%optD%"=="%off%"                >>%OCS% ECHO       ^<ExcludeApp ID="OneDrive" /^>
                                    >>%OCS% ECHO       ^<ExcludeApp ID="Lync" /^>
                                    >>%OCS% ECHO       ^<ExcludeApp ID="Groove" /^>
                                    >>%OCS% ECHO       ^<ExcludeApp ID="Bing" /^>
                                    >>%OCS% ECHO     ^</Product^>
If "%opt8%"=="%on%"                 >>%OCS% ECHO     ^<Product ID="VisioProRetail"^>
If "%opt8%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="MatchOS" Fallback="en-US" /^>
If "%opt8%"=="%on%"                 >>%OCS% ECHO       ^<ExcludeApp ID="Bing" /^>
If "%opt8%"=="%on%"                 >>%OCS% ECHO     ^</Product^>
If "%opt9%"=="%on%"                 >>%OCS% ECHO     ^<Product ID="ProjectProRetail"^>
If "%opt9%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="MatchOS" Fallback="en-US" /^>
If "%opt9%"=="%on%"                 >>%OCS% ECHO       ^<ExcludeApp ID="Bing" /^>
If "%opt9%"=="%on%"                 >>%OCS% ECHO     ^</Product^>
If "%optP%"=="%on%"                 >>%OCS% ECHO     ^<Product ID="ProofingTools"^>
If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="MatchOS" Fallback="en-US" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="af-ZA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sq-AL" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ar-SA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="hy-AM" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="as-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="az-Latn-AZ" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="bn-BD" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="bn-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="eu-ES" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="bs-Latn-BA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="bg-BG" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ca-ES" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="zh-CN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="zh-TW" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="hr-HR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="cs-CZ" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="da-DK" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="nl-NL" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="en-US" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="et-EE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="fi-FI" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="fr-FR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="gl-ES" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ka-GE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="de-DE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="el-GR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="gu-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ha-Latn-NG" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="he-IL" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="hi-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="hu-HU" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="is-IS" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ig-NG" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="id-ID" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ga-IE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="xh-ZA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="zu-ZA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="it-IT" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ja-JP" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="kn-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="kk-KZ" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="rw-RW" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sw-KE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="kok-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ko-KR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ky-KG" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="lv-LV" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="lt-LT" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="lb-LU" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="mk-MK" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ms-MY" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ml-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="mt-MT" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="mi-NZ" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="mr-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ne-NP" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="nb-NO" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="nn-NO" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="or-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ps-AF" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="fa-IR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="pl-PL" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="pt-BR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="pt-PT" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="pa-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ro-RO" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="rm-CH" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ru-RU" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="gd-GB" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sr-Cyrl-BA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sr-Cyrl-RS" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sr-Latn-RS" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="nso-ZA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="tn-ZA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="si-LK" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sk-SK" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sl-SI" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="es-ES" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="sv-SE" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ta-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="tt-RU" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="te-IN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="th-TH" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="tr-TR" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="uk-UA" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ur-PK" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="uz-Latn-UZ" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="ca-ES-Valencia" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="vi-VN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="cy-GB" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="wo-SN" /^>
::If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<Language ID="yo-NG" /^>
If "%optP%"=="%on%"                 >>%OCS% ECHO       ^<ExcludeApp ID="Bing" /^>
If "%optP%"=="%on%"                 >>%OCS% ECHO     ^</Product^>
                                    >>%OCS% ECHO     ^<Product ID="LanguagePack"^>
                                    >>%OCS% ECHO       ^<Language ID="MatchInstalled" /^>
                                    >>%OCS% ECHO       ^<ExcludeApp ID="Bing" /^>
                                    >>%OCS% ECHO     ^</Product^>
                                    >>%OCS% ECHO   ^</Add^>
If "%optS%"=="%on%"                 >>%OCS% ECHO   ^<Updates Enabled="TRUE" /^>
If "%optS%"=="%on%"                 >>%OCS% ECHO   ^<Display Level="Full" AcceptEULA="TRUE" /^>
If "%optS%"=="%on%"                 >>%OCS% ECHO   ^<Property Name="ForceAppShutdown" Value="TRUE" /^>
If "%optS%"=="%on%"                 >>%OCS% ECHO   ^<AppSettings^>
If "%optS%,%opt1%"=="%on%,%on%"     >>%OCS% ECHO     ^<User Key="Software\Microsoft\Office\16.0\Word\Options" Name="defaultformat" Value="" Type="REG_SZ" App="word16" Id="L_SaveWordfilesas" /^>
If "%optS%,%opt2%"=="%on%,%on%"     >>%OCS% ECHO     ^<User Key="Software\Microsoft\Office\16.0\Excel\Options" Name="defaultformat" Value="51" Type="REG_DWORD" App="excel16" Id="L_SaveExcelfilesas" /^>
If "%optS%,%opt3%"=="%on%,%on%"     >>%OCS% ECHO     ^<User Key="Software\Microsoft\Office\16.0\PowerPoint\Options" Name="defaultformat" Value="27" Type="REG_DWORD" App="ppt16" Id="L_SavePowerPointfilesas" /^>
If "%optS%"=="%on%"                 >>%OCS% ECHO   ^</AppSettings^>
                                    >>%OCS% ECHO ^</Configuration^>
ENDLOCAL
GoTo :EOF
:SKIP


If Not "%optO%,%OFiles%"=="%on%,%off%" GoTo SKIP
ECHO.
ECHO INFO: This can take a few minutes, depending on your Internet Speed.
ECHO.
Set "CPU=32" & CALL :CONFIG
ECHO.
ECHO Downloading Microsoft Office 365 ProPlus %CPU%-bit (Size 3,08 GB) . . .
ping -n 3 localhost 1>NUL
START "" /WAIT /B ".\setup.exe" /download ".\Office 365 Setup Config.xml"
ECHO.
Set "CPU=64" & CALL :CONFIG
ECHO.
ECHO Downloading Microsoft Office 365 ProPlus %CPU%-bit (Size 3,32 GB) . . .
ping -n 3 localhost 1>NUL
START "" /WAIT /B ".\setup.exe" /download ".\Office 365 Setup Config.xml"
:SKIP

If Not "%optS%"=="%on%" GoTo SKIP
ECHO.
If "%optB%"=="64-bit" Set "CPU=64"
If "%optB%"=="32-bit" Set "CPU=32"
CALL :CONFIG
ECHO.
ECHO Installing Microsoft Office 365 ProPlus %CPU%-bit . . .
ping -n 3 localhost 1>NUL
START "" /WAIT /B ".\setup.exe" /configure ".\Office 365 Setup Config.xml"

ECHO.
ECHO Disabling Microsoft Office 365 Telemetry . . .
ping -n 3 localhost 1>NUL
REG ADD "HKLM\SOFTWARE\Microsoft\Office\Common\ClientTelemetry" /v "DisableTelemetry" /t REG_DWORD /d "00000001" /f 1>NUL
:SKIP

If Not "%optO%,%OFiles%"=="%on%,%on%" GoTo SKIP
ECHO.
ECHO Deleting Microsoft Office 365 ProPlus Installation Files . . .
ping -n 3 localhost 1>NUL
RmDir /s /q ".\Office" 2>NUL
ECHO.
ECHO ---------------------------------[Finished]--------------------------------
ping -n 5 localhost 1>NUL
Set "optO=%off%"
Set "optS=%on%"
GoTo MENU
:SKIP


:END
ECHO.
ECHO --------------------------------[Completed]--------------------------------
ECHO.
Pause
DEL /f ".\Office 365 Setup Config.xml" 2>NUL
EXIT
