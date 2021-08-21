#if not running as admin then start a new powershell process as admin with script assuming logged in user has rights
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator"))  
{  
  $arguments = "& '" +$myinvocation.mycommand.definition + "'"
  Start-Process powershell -Verb runAs -ArgumentList $arguments
  Break
}

#start time of script running
$startTime = (Get-Date)

#clear screen so output text is readable
Clear-Host

#setup vars
$stringToFind = '    <font color="red">\r\n    <P>Learning HTML will enable you to:\r\n    <UL>\r\n    <LI>create your own simple pages\r\n    <LI>read and appreciate pages created by others\r\n    <LI>develop an understanding of the creative and literary implications of web-texts\r\n    <LI>have the confidence to branch out into more complex web design \r\n    </UL></P>\r\n    </font>'
$inputDir = "$PSSCriptRoot\input"
$outputDir = "$PSSCriptRoot\output"
$expectedDirCount = 8
$logStart = "  Sparky ->"

#function to output message with highlighted text based on error state
function Write-HostAdv {

    param(
        [Parameter(Mandatory)]
        [int]$code,
        [String]$messageMain,
        [String]$messageStatus
    )

    if ($code -eq 0) { #if code is 0 then no error, so set text to green
      $textColor = "Green"  
    } else {
        $textColor = "Red"
    }

    Write-Host "$logStart $messageMain" -NoNewline
    Write-Host "  $messageStatus  " -BackgroundColor $textColor -ForegroundColor "Black"

}

#function to verify 8 html or htm files are in input directory
function verifyInput {
    Write-Host "$logStart Starting Input Verification.  "  -BackgroundColor "Yellow" -ForegroundColor "Black"
    $directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
    $count = $directoryArray.Count
    if ($count -ne $expectedDirCount) {
        Write-HostAdv -code 1 -messageMain "Input Verification Complete. Status:     " -messageStatus "FAILED"
        Write-Host "$logStart Found $count Input Files. Expected Value = $expectedDirCount.  "
        Write-Host "$logStart Aborting.  "
        exit
    }
    Write-HostAdv -code 0 -messageMain "Input Verification Complete. Status:     " -messageStatus "SUCCESS"
    Write-Host "$logStart Found $count HTML/HTM Files.  "
}

#function to verify that html string is in all files in input directory
function verifyString {
    Write-Host "`n$logStart Starting String Verification.  "  -BackgroundColor "Yellow" -ForegroundColor "Black"
    $directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
    $textCounter = 1
    foreach ($file in $directoryArray) {
        $fileContent = Get-Content -Path "$inputDir\$file" -Raw

        if (Select-String -InputObject $fileContent -Pattern $stringtoFind) {
            Write-HostAdv -code 0 -messageMain "File $textCounter '$($file.Name)' Verification:     " -messageStatus "SUCCESS"
            $textCounter++
        } else {
            Write-HostAdv -code 1 -messageMain "File $textCounter '$($file.Name)' Verification:     " -messageStatus "FAILED"
            Write-Host "$logStart File $textCounter '$($file.Name)' Does NOT Contain Expected String.  "
            Write-Host "$logStart Aborting.  "
            exit
        }
        
    }
}

Write-Host "$logStart Starting Parser.  `n"  -BackgroundColor "Yellow" -ForegroundColor "Black"

#run verification functions
verifyInput
verifyString

Write-Host "`n$logStart All Verification Functions Complete.  "  -BackgroundColor "Yellow" -ForegroundColor "Black"
Write-Host "`n$logStart Starting Conversion.  "  -BackgroundColor "Yellow" -ForegroundColor "Black"

#clear output directory before converting or modifying any files
Remove-Item -Path "$outputDir\*"

#modify the web page files to remove string
$directoryArray = Get-ChildItem -Path $inputDir | Where-Object { $_.Extension -eq '.html' -or $_.Extension -eq '.htm'}
foreach ($file in $directoryArray) {

    Write-Host "$logStart Converting $file.  "

    #put file into arraylist
    $fileContent = [System.Collections.ArrayList](Get-Content -Path "$inputDir\$file")
    #remove 9 elements starting at element 16
    $fileContent.RemoveRange(16, 9)

    Out-File -FilePath "$outputDir\$($file.Name)" -InputObject $fileContent

    Start-Sleep -Seconds 1

    $ie = new-object -ComObject "InternetExplorer.Application"
    $requestUri = "$outputDir\$file"
    $ie.silent = $true
    $ie.navigate($requestUri)

    while ($ie.Busy) { Start-Sleep -Seconds 1 }

    $pdfPrinter = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Name -eq "Microsoft Print to PDF"}
    $pdfPrinter.SetDefaultPrinter() | Out-Null

    Start-Sleep -Seconds 1
    $ie.ExecWB(6,2) #prints from default printer
    Start-Sleep -Seconds 2

    $wshell = New-Object -com WScript.Shell
    Start-Sleep -Seconds 1
    $wshell.sendkeys("$outputDir\$($file.BaseName).pdf")
    Start-Sleep -Seconds 1
    $wshell.sendkeys("{ENTER}")
    Start-Sleep -Seconds 1

    Remove-Item -Path "$outputDir\$($file.Name)" #remove HTML file from output

    #check for PDF and report
    if ((Test-Path -Path "$outputDir\$($file.Basename).pdf") -eq $true) {
        Write-HostAdv -code 0 -messageMain "File '$($file.Name)' Conversion:     " -messageStatus "SUCCESS"
    } else {
        Write-HostAdv -code 1 -messageMain "File '$($file.Name)' Conversion:     " -messageStatus "FAILED"
        Write-Host "$logStart Aborting.  "
        exit
    }

    Start-Sleep -Seconds 1

    #kill ie and cleanup
    $ie.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ie) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

}

#end time of script running
$endTime = (Get-Date)

Write-Host "`n$logStart Parser has Completed in $(($endTime-$startTime).totalseconds) Seconds.  "  -BackgroundColor "Yellow" -ForegroundColor "Black"
Write-Host "$logStart Find Finished PDF Files at `"$outputDir`"  "  -BackgroundColor "Yellow" -ForegroundColor "Black"

pause