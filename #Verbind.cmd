@Echo off
REM Dit bestand NIET zelf aanpassen!!!
REM ==================================
set _PoSh=%systemroot%\System32\WindowsPowerShell\v1.0\Powershell.exe
set _Edir=%CD%
set _Pdir=%_EDir%\.Progs
set _Cdir=%_Edir%\.Cache


%_PoSh% -ExecutionPolicy Unrestricted -NoLogo -File %_Pdir%\Verbindmappen.ps1 -fIPAddress "*" %1 %2 %3 %4 %5 %6 %7 %8 %9

set _PoSh=
set _Edir=
set _Pdir=
set _Cdir=


