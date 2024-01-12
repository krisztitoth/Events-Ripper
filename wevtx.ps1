param(
    [Parameter(Mandatory)][string]$FolderPath,
    [string]$TempFolder,
    [string]$OutputFile,
    [switch]$KeepTemporaryFiles
    )

if (-not (Test-Path $FolderPath)) {
    Write-Host "Folder '$FolderPath' does not exist, exiting" -ForegroundColor Red
    exit
}

if ($OutputFile -eq "") {
    $OutputFile = "$env:USERPROFILE\Desktop\RippedEventsReport.txt"
}

if ($TempFolder -eq "") {
    New-Item -Path "$env:USERPROFILE\Desktop\EventsRipper-TEMP" -ItemType "directory" -Force
    if (-not (Test-Path "$env:USERPROFILE\Desktop\EventsRipper-TEMP")) {
        Write-Host "Folder '$env:USERPROFILE\Desktop\EventsRipper-TEMP' cannot be created, exiting" -ForegroundColor Red
        exit
    }
    else
    { 
        $TempFolder = "$env:USERPROFILE\Desktop\EventsRipper-TEMP" 
        Write-Host "Temporary folder $TempFolder has been created."
    }
}

$MyCounter = 1
Get-ChildItem -Path $FolderPath | ForEach-Object {
    try {
        Write-Host $_.FullName
        $myPath = $_.FullName
        Start-Process -FilePath "$PWD\logparser.exe" -ArgumentList "-i:evt -o:csv `"Select RecordNumber,TO_UTCTIME(TimeGenerated),EventID,SourceName,ComputerName,SID,Strings from $myPath`"" -RedirectStandardOutput "$TempFolder\$MyCounter-firstStep.txt" -WindowStyle Hidden -Wait
        Start-Process -FilePath "$PWD\evtxparse.exe" -ArgumentList "$TempFolder\$MyCounter-firstStep.txt"  -RedirectStandardOutput "$TempFolder\$MyCounter-secondStep.txt" -WindowStyle Hidden -Wait
        Get-Content -Path "$TempFolder\$MyCounter-secondStep.txt" -Raw | Add-Content -Path "$TempFolder\Summary.txt"
        $MyCounter += 1

    }
    catch {
        Write-Host "Error processing $($_.FullName): $_" -ForegroundColor DarkRed
    }
}

Start-Process -FilePath "$PWD\erip.exe" -ArgumentList "-f $TempFolder\Summary.txt -a"  -RedirectStandardOutput $OutputFile -WindowStyle Hidden -Wait
if (-not ($KeepTemporaryFiles)) {
    Get-ChildItem "$TempFolder\" -Recurse | Remove-Item -Confirm
}
$MyCounter -= 1
Write-Host "Processing of $MyCounter file(s) is completed. See $OutputFile."

