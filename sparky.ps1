#setup vars
#$stringToFind = '    <font color="red">\r\n    <P>Learning HTML will enable you to:\r\n    <UL>\r\n    <LI>create your own simple pages\r\n    <LI>read and appreciate pages created by others\r\n    <LI>develop an understanding of the creative and literary implications of web-texts\r\n    <LI>have the confidence to branch out into more complex web design \r\n    </UL></P>\r\n    </font>'
$stringToFind = '    <font color="red">'
$inputDir = "$PSSCriptRoot\input"
$outputDir = "$PSSCriptRoot\output"
$expectedDirCount = 8
$logStart = "  Sparky ->"

#clear screen
Clear-Host

#verify 8 html or htm files are in input directory
function verifyInput {
    Write-Host "$logStart Starting Input Verification.  "  -BackgroundColor "Green" -ForegroundColor "Black"
    $directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
    $count = $directoryArray.Count
    if ($count -ne $expectedDirCount) {
        Write-Host "$logStart Input Verification Failed.  "  -BackgroundColor "Red" -ForegroundColor "Black"
        Write-Host "$logStart Found $count Input Files. Expected Value = $expectedDirCount.  "  -BackgroundColor "Red" -ForegroundColor "Black"
        Write-Host "$logStart Aborting.  "  -BackgroundColor "Red" -ForegroundColor "Black"
        exit
    }
    Write-Host "$logStart Input Verification Complete. Found $count HTML/HTM Files.  "  -BackgroundColor "Green" -ForegroundColor "Black"
}

#verify that html string is in all files in input directory
function verifyString {
    Write-Host "$logStart Starting String Verification.  "  -BackgroundColor "Green" -ForegroundColor "Black"
    $directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
    $textCounter = 1
    foreach ($file in $directoryArray) {
        $fileContent = Get-Content -Path "$inputDir\$file" -Raw
        if (Select-String -InputObject $fileContent -Pattern $stringtoFind) {
            Write-Host "$logStart File $textCounter '$($file.Name)' Verified.  "  -BackgroundColor "Green" -ForegroundColor "Black"
            $textCounter++
        } else {
            Write-Host "$logStart File $textCounter Failed.  "  -BackgroundColor "Red" -ForegroundColor "Black"
            Write-Host "$logStart File $textCounter '$($file.Name)' Does NOT Contain Expected String.  "  -BackgroundColor "Red" -ForegroundColor "Black"
            Write-Host "$logStart Aborting.  "  -BackgroundColor "Red" -ForegroundColor "Black"
            exit
        }
        
    }
}

Write-Host "$logStart Starting Parser.  `n"  -BackgroundColor "Green" -ForegroundColor "Black"

#run verification
verifyInput
verifyString

Write-Host "$logStart Verification Functions Complete.  "  -BackgroundColor "Green" -ForegroundColor "Black"
Write-Host "`n$logStart Starting Conversion.  "  -BackgroundColor "Green" -ForegroundColor "Black"

#modify the web page files to remove string
$directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
foreach ($file in $directoryArray) {

    if ($file.Name -eq 'File-1.html') {

    
    Write-Host "`n$logStart Starting Conversion of $file.  "  -BackgroundColor "Green" -ForegroundColor "Black"
    $tempContent = Get-Content -Path "$inputDir\$file"
    $tempContent2 = $tempContent.Replace($stringToFind,"")
    $tempContent2
    #$tempContent
    #$file.FullName
    #Out-File -FilePath $element.FullName -InputObject $tempWeb


    }
}

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