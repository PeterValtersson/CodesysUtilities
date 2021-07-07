
set destination=C:\ProgramData\CODESYS\Script Commands

if not exist "%destination%" mkdir "%destination%"

copy "config.json" "%destination%\config.json"
copy "CoDeSys.ico" "%destination%\CoDeSys.ico"
xcopy "%cd%\*.py" "%destination%" /y