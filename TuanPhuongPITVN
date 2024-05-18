CHCP 1258 >nul 2>&1
CHCP 65001 >nul 2>&1
@echo off
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"
if '%errorlevel%' NEQ '0' (
    echo  Run CMD as Administrator...
    goto goUAC 
) else (
 goto goADMIN )

:goUAC
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    set params = %*:"=""
    echo UAC.ShellExecute "cmd.exe", "/c %~s0 %params%", "", "runas", 1 >> "%temp%\getadmin.vbs"
    "%temp%\getadmin.vbs"
    del "%temp%\getadmin.vbs"
    exit /B

:goADMIN
    pushd "%CD%"
    CD /D "%~dp0"
	


====================================================================
title Convert cac phien ban office!
color f0


:MainMenu
mode con: cols=65 lines=25
echo. 
cls
echo.
echo.                          == MENU ==
echo.      
echo.      [  1. Convert Office    : Nhan so 1  ] 
echo.
echo.      [  2. Exit              : Nhan so 2  ]
echo.	  

echo.	  
echo.        
echo.
@echo ===========================
Choice /N /C 12 /M "* Nhap lua chon cua ban: 


if ERRORLEVEL 2 goto:Exit
if ERRORLEVEL 1 goto:convert









:============================================================================================================
:in_shortcut_office
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\W*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\E*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\P*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\V*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\A*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\O*.lnk" "%AllUsersProfile%\Desktop"
COPY /Y "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\T*.lnk" "%AllUsersProfile%\Desktop"
goto:MainMenu


:======================================================================================================================================================

:convert
mode con: cols=70 lines=27
echo. 
cls
set id1=O365ProPlusRetail
set id2=ProPlus2019Retail
set id3=ProPlus2019Volume
set id4=ProPlusRetail
set id5=Office16.PROPLUS
set id6=Office15.PROPLUSR
set id7=Office14.PROPLUSR
set id8=Officel4.PROPLUS
set id9=ProPlus2021Retail
set id10=ProPlus2021Volume

for %%b in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%%b" /v "Displayname"') do cls&set a1=%%A&set b1=%%b)
for %%b in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\%%b - en-us" /v "Displayname"') do cls&set a1=%%A&set b1=%%b)
for %%c in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%%c" /v "Displayname"') do cls&set a1=%%A&set b1=%%c)
for %%c in (%id1%,%id2%,%id3%,%id4%,%id5%,%id6%,%id7%,%id8%,%id9%,%id10%) do (for /f "delims=" %%A in ('reg query "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\%%c - en-us" /v "Displayname"') do cls&set a1=%%A&set b1=%%c)

:menu_convert
cls
echo.
echo.
if [%b1%] EQU [O365ProPlusRetail] echo.  Dang su dung:       Microsoft 365 Apps for enterprise&set ak47=365
if [%b1%] EQU [ProPlus2019Retail] echo.  Dang su dung:     Office Professional Plus 2019 (Retail)&set ak47=19
if [%b1%] EQU [ProPlus2019Volume] echo.  Dang su dung:     Office Professional Plus 2019 (Volume)&set ak47=19
if [%b1%] EQU [ProPlusRetail]     echo.  Dang su dung:     Office Professional Plus 2016 (Retail)&set ak47=16
if [%b1%] EQU [Office16.PROPLUS]  echo.  Dang su dung:     Office Professional Plus 2016 (Volume)&set ak47=16
if [%b1%] EQU [Office15.PROPLUSR] echo.  Dang su dung:     Office Professional Plus 2013 (Retail)
if [%b1%] EQU [Office14.PROPLUSR] echo.  Dang su dung:     Office Professional Plus 2010 (Retail)
if [%b1%] EQU [Office14.PROPLUS]  echo.  Dang su dung:     Office Professional Plus 2010 (Volume)
if [%b1%] EQU [ProPlus2021Retail] echo.  Dang su dung:     Office Professional Plus 2021 (Retail)&set ak47=21
if [%b1%] EQU [ProPlus2021Volume] echo.  Dang su dung:     Office LTSC Professional Plus 2021 (Volume)&set ak47=21




set off21=""
echo.
echo.
echo.      
echo.      [  1. Retail sang Volume                      : Nhan so 1  ] 
echo.
echo.      [  2. Volume sang Retail                      : Nhan so 2  ]
echo.
echo.      [  3. ProPlus sang "Home and Student"         : Nhan so 3  ]
echo.
echo.      [  4. ProPlus sang "Home and Bussiness"       : Nhan so 4  ]
echo.
echo.      [  5. "Student" or "Bussiness" sang Pro Plus  : Nhan so 5  ]
echo.
echo.           -------------------------------------------
echo.
echo.               [  6.Quay lai: Nhan so 6  ]
echo.	  
echo.	  
echo.        
echo.
@echo ===========================
Choice /N /C 123456 /M "* Nhap lua chon cua ban: 

if ERRORLEVEL 6 goto:menu_convert  
if ERRORLEVEL 5 set option=retail&goto:khoidau      
if ERRORLEVEL 4 set option=buss&set off21=no&goto:khoidau      
if ERRORLEVEL 3 set option=student&set off21=no&goto:khoidau      
if ERRORLEVEL 2 set option=retail&goto:khoidau    
if ERRORLEVEL 1 set option=vl&goto:khoidau




:khoidau
if [%ak47%] EQU [16] goto:office2016
if [%ak47%] EQU [19] goto:office2019
if [%ak47%] EQU [21] goto:office2021
if [%ak47%] EQU [365] echo.&echo. Khong Duoc Phep Convert!!!&timeout 5&goto:menu_convert


:office2019
if [%option%] EQU [vl] goto:vl19
if [%option%] EQU [retail] goto:retail19
if [%option%] EQU [student] goto:student19
if [%option%] EQU [buss] goto:buss19

:retail19
set Key=NW69C-TFYXD-YBDBD-KBMHD-MDYCT
set Description=ProPlus2019MSDNR_Retail
goto:ketthuc
:student19
set Key=FMTDN-KMP4R-H983J-6J87C-YKPY7
set Description=HomeStudent2019R_OEM_Perp
goto:ketthuc
:buss19
set Key=BNH3K-7JG4D-C7XF8-6FM6B-8FJFG
set Description=HomeBusiness2019R_Retail
goto:ketthuc
:vl19
set Key=W9HYN-C8J79-2YGTT-JVQW8-K2GT3
set Description=ProPlus2019VL_MAK_AE
goto:ketthuc



:office2021
if [%option%] EQU [vl] goto:vl21
if [%option%] EQU [retail] goto:retail21
if [%off21%] EQU [no] goto:echo.&echo. Khong Co Data de Convert!!!&timeout 5&goto:menu_convert
if [%off21%] EQU [no] goto:echo.&echo. Khong Co Data de Convert!!!&timeout 5&goto:menu_convert

:retail21
set Key=9D9QG-NKMFF-GGGR8-TYF9Y-3GRY9
set Description=ProPlus2021MSDNR_Retail
goto:ketthuc
:vl21
set Key=XN3GG-8V6WR-GXYKF-3MX6H-JQP89
set Description=ProPlus2021VL_MAK_AE
goto:ketthuc



:offfice2016
if [%option%] EQU [vl] goto:vl16
if [%option%] EQU [retail] goto:retail16
if [%option%] EQU [student] goto:student16
if [%option%] EQU [buss] goto:buss16

:retail16
set Key=BQ8H7-N8MWT-9FD39-24MDC-TMVHC
set Description=Office16_ProPlusMSDNR_Retail
goto:ketthuc
:student16
set Key=QN97V-HRX9H-QQVB6-8Q9BM-YWRFK
set Description=Office16_HomeStudentR_Retail
goto:ketthuc
:buss16
set Key=WBXCT-N9PQ9-WJG7Q-CYH4D-BG9VX
set Description=Office16_HomeBusinessR_Retail
goto:ketthuc
:vl16
set Key=KC7N8-WXH8P-8M8R7-QWYKK-JXCQY
set Description=Office16_ProPlusVL_MAK
goto:ketthuc




:ketthuc
mode con: cols=80 lines=25
reg Delete HKLM\Software\Wow6432Node\Microsoft\Office\16.0\Common\OEM /f
reg Delete HKLM\Software\Microsoft\Office\16.0\Common\OEM /f
for %%a in (4,5,6) do (if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles%\Microsoft Office\Office1%%a")
If exist "%ProgramFiles% (x86)\Microsoft Office\Office1%%a\ospp.vbs" (cd /d "%ProgramFiles% (x86)\Microsoft Office\Office1%%a"))&cls
for /f "tokens= 8" %%b in ('cscript //nologo OSPP.VBS /dstatus ^| findstr /b /c:"Last 5"') do (cscript //nologo ospp.vbs /unpkey:%%b)
for /f %%i in ('dir /b ..\root\Licenses16\%Description%*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%%i"
cscript ospp.vbs /remhst
cscript ospp.vbs /ckms-domain
cscript ospp.vbs /inpkey:%key%
cscript //nologo ospp.vbs /act
echo.
echo.=========================
echo. Da convert thanh cong!
echo.=========================
echo.
echo.
timeout 7
goto:menu_convert













:======================================================================================================================================================
:Exit
echo. Tam biet!
timeout 3
exit




