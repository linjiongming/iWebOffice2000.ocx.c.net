@echo off
copy iWebOffice2000.ocx %windir%\system32\iWebOffice2000.ocx
copy SignatureSetEnv.exe %windir%\system32\SignatureSetEnv.exe
regsvr32 %windir%\system32\iWebOffice2000.ocx "-u" "-s"
call %windir%\system32\SignatureSetEnv.exe
regsvr32 %windir%\system32\iWebOffice2000.ocx
pause