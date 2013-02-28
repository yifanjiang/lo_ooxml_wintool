@echo off
set INPUT=C:\fooooxml
set OUTPUT=C:\fooooxml\reference

FOR /R "%INPUT%" %%F in (*.doc) do cscript /nologo "ms2pdf.vbs" /nologo %%F /o:"%OUTPUT%"
FOR /R "%INPUT%" %%F in (*.ppt) do cscript /nologo "ms2pdf.vbs" /nologo %%F /o:"%OUTPUT%"
FOR /R "%INPUT%" %%F in (*.xls) do cscript /nologo "ms2pdf.vbs" /nologo %%F /o:"%OUTPUT%"
