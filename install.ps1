#Requires -RunAsAdministrator
$installPath = "C:\intelliansys\LibraryForExcel"

if (Test-Path $installPath)
{
	rm $installPath -Recurse
}
mkdir $installPath

$filenames = ls -n .\Release
foreach ($filename in $filenames)
{
	cp .\Release\$filename $installPath
}

$dotnetVersion = Get-ChildItem $env:windir\Microsoft.NET\Framework64 -Name | Select-String "v" | Select -last 1
$exePath = echo $env:windir\Microsoft.Net\Framework64\${dotnetVersion}\RegAsm.exe

& $exePath $installPath\CSharpLibraryForExcel.dll /tlb /codebase

Read-Host "`n`nPress Enter to End Application" 