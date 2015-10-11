@echo off
for %%* in (.) do (
SET projectname=%%~n*
)
rmdir /S /Q "Z:\Builds\%projectname%"
pyinstaller -D --icon=icon.ico --noconsole "main.py"
ren dist\main %projectname%
ren dist\%projectname%\main.exe %projectname%.exe
FOR /F %%a in (build_files.txt) DO XCOPY %%a dist\%projectname%
MOVE /Y "dist\%projectname%" "Z:\Builds\%projectname%"
rmdir /S /Q dist 
rmdir /S /Q build
echo 'Build complete'
pause