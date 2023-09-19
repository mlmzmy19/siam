@echo off

timeout /t 30 /nobreak
FOR %%X IN ("C:\MisProgramas\*.*") DO rundll32 shell32.dll,ShellExec_RunDLL %%X

c:\ActSIAM.bat