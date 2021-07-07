
set destination="C:\ProgramData\CODESYS\Script Commands"

if not exist %destination%\NUL mkdir %destination%

copy "config.json" %destination%\config.json
copy "CoDeSys.ico" %destination%\CoDeSys.ico
for . %%f in (*.py) do copy %%f %destination%