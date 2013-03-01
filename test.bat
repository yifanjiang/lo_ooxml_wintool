@echo off

taskkill /F /IM POWERPNT.EXE & taskkill /F /IM WINWORD.EXE & taskkill /F /IM EXCEL.EXE

set INPUT="C:\fooooxml\reference"
set OUTPUT="C:\fooooxml\test"

FOR /R "%INPUT%" %%F in (*.doc) do cscript /nologo "ms2pdf.vbs" /nologo "%%F" /o:"%OUTPUT%"
FOR /R "%INPUT%" %%F in (*.ppt) do cscript /nologo "ms2pdf.vbs" /nologo "%%F" /o:"%OUTPUT%"
FOR /R "%INPUT%" %%F in (*.xls) do cscript /nologo "ms2pdf.vbs" /nologo "%%F" /o:"%OUTPUT%"

taskkill /F /IM POWERPNT.EXE & taskkill /F /IM WINWORD.EXE & taskkill /F /IM EXCEL.EXE