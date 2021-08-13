#clear screen
Clear-Host


#function to verify 8 html files are in input directory

#function to verify that html string is in all files in input directory

Write-Host "  Sparky -> Starting Parser  "  -BackgroundColor "Green" -ForegroundColor "Black"



<#
$ie = new-object -ComObject "InternetExplorer.Application"
$requestUri = "$PSSCriptRoot\test.html"
$ie.silent = $true
$ie.navigate($requestUri)
echo "navigating"
while($ie.Busy) { Start-Sleep -Milliseconds 100 }
$doc = $ie.Document

echo "setting default"
$pdfPrinter = Get-WmiObject -Class Win32_Printer | Where{$_.Name -eq "Microsoft Print to PDF"}
$pdfPrinter.SetDefaultPrinter() | Out-Null

echo "start sleep"
Start-Sleep -Milliseconds 500
echo "start print"
$ie.ExecWB(6,2) #prints from default printer
Start-Sleep -Milliseconds 500

$wshell = New-Object -com WScript.Shell
Start-Sleep -Milliseconds 100
$wshell.sendkeys("$PSSCriptRoot\temp.pdf")
Start-Sleep -Milliseconds 50
$wshell.sendkeys("{ENTER}")
#>