@echo off
 
set ps_fn=ofd.ps1
echo [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") ^| out-null > %ps_fn%
echo $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog >> %ps_fn%
echo $OpenFileDialog.initialDirectory = "%USERPROFILE%\Downloads" >> %ps_fn%
echo $OpenFileDialog.filter = "Excel workbooks (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv" >> %ps_fn%
echo $OpenFileDialog.ShowDialog() >> %ps_fn%
echo $OpenFileDialog.filename >> %ps_fn%
 
for /F "tokens=* usebackq" %%a in (`powershell -executionpolicy bypass -file %ps_fn%`) do if not "%%a" == "" if not "%%a" == "OK" set filename=%%a
del %ps_fn%
 
if not "%filename%"=="" echo %filename%