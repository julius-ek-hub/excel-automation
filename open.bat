<# : chooser.bat

:: drop file to execute, or open a file chooser dialog window to execute.
:: code from https://stackoverflow.com/a/15885133/1683264

@ECHO OFF

if "%~1" == "" goto SELECT
bin\dosomething "%~1"
goto :EOF

:SELECT

setlocal
for /f "delims=" %%I in ('powershell -noprofile "iex (${%~f0} | out-string)"') do (
    echo %%~I
)
goto :EOF

: end Batch portion / begin PowerShell hybrid chimera #>

Add-Type -AssemblyName System.Windows.Forms
$f = new-object Windows.Forms.OpenFileDialog
$f.InitialDirectory = pwd
$f.Filter = "Excel Files |*.xlsx|CSV Files |*.csv"
$f.ShowHelp = $true
$f.Multiselect = $false
[void]$f.ShowDialog()
if ($f.Multiselect) { $f.FileNames } else { $f.FileName }