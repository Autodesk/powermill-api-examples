cd C:\WINDOWS\Microsoft.NET\Framework64\v4.0.30319\
regasm.exe "C:\Users\greniem\Delcam\PowerMILL\Macros\Drilling Automation\Drilling Automation\bin\Debug\DrillingAutomation.dll" /register /codebase
REG ADD "HKCR\CLSID\{7893D42C-0A4D-44E4-9B06-8C9E2E8F2BE2}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:32 /f
REG ADD "HKCR\CLSID\{7893D42C-0A4D-44E4-9B06-8C9E2E8F2BE2}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f

