:: Changes to the Office directory according to the installed version
cd /d "%ProgramFiles%\Microsoft Office\Office16"
if not exist ospp.vbs (
    cd /d "%ProgramFiles(x86)%\Microsoft Office\Office16"
)
if not exist ospp.vbs (
    echo Didnt find ospp.vbs in the common paths
    pause
    exit /b
)

:: Obtains the last 5 characters of the installed product key
for /f "tokens=*" %%A in ('cscript //nologo ospp.vbs /dstatus ^| findstr /i "Last 5 characters"') do (
    set "line=%%A"
    set "key=!line:~-5!"
)

:: Checks if the value is obtained
if "%key%"=="" (
    echo Cannot obtain the installed key
    pause
    exit /b
)

:: uninstalls the current key 
echo uninstalling key: %key%
cscript //nologo ospp.vbs /unpkey:%key%

:: Activates Office
echo Activating Office...
cscript //nologo ospp.vbs /act

echo process completed.
pause