# ============================================================================
# Configuration
# ============================================================================
$baseFolder = $PSScriptRoot
$winrarExe = "C:\Program Files\WinRAR\WinRAR.exe"
$maxSizeBytes = 10GB

# ============================================================================
# Script Logic
# ============================================================================
Clear-Host
Write-Host "Starting the zipping process..." -ForegroundColor Green
Write-Host "Processing folder: `"$($baseFolder)`""
Write-Host ""

if (-not (Test-Path -Path $winrarExe)) {
    Write-Host "ERROR: WinRAR not found at `"$winrarExe`". Please fix the path." -ForegroundColor Red
    Start-Sleep -Seconds 10
    exit
}

Get-ChildItem -Path $baseFolder -Directory | ForEach-Object {
    $personFolder = $_
    Write-Host "=====================================================" -ForegroundColor Cyan
    Write-Host "Processing subfolder: `"$($personFolder.Name)`"" -ForegroundColor Cyan
    Write-Host "====================================================="

    # --- THE FIX: Step into the person's folder ---
    Push-Location -Path $personFolder.FullName

    $cr3Files = Get-ChildItem -Path "." -Filter "*.CR3" | Sort-Object Name

    if ($cr3Files) {
        $currentZipSize = 0
        $filesForThisZip = New-Object System.Collections.ArrayList
        $startFileNum = $null
        $tempListFile = "temp_zip_list.txt"

        foreach ($file in $cr3Files) {
            $fileNum = $file.BaseName.Substring($file.BaseName.Length - 4)
            if ($startFileNum -eq $null) { $startFileNum = $fileNum }
            $endFileNum = $fileNum
            $filesForThisZip.Add($file.Name) | Out-Null
            $currentZipSize += $file.Length

            if ($currentZipSize -ge $maxSizeBytes) {
                $zipName = "$($startFileNum)-$($endFileNum).zip"
                Write-Host "Creating archive: $zipName"
                $filesForThisZip | Set-Content -Path $tempListFile
                & $winrarExe a -afzip -ep -m0 $zipName "@$tempListFile" | Out-Null
                Remove-Item $tempListFile
                $filesForThisZip.Clear()
                $currentZipSize = 0
                $startFileNum = $null
            }
        }

        if ($filesForThisZip.Count -gt 0) {
            $zipName = "$($startFileNum)-$($endFileNum).zip"
            Write-Host "Creating final archive: $zipName"
            $filesForThisZip | Set-Content -Path $tempListFile
            & $winrarExe a -afzip -ep -m0 $zipName "@$tempListFile" | Out-Null
            Remove-Item $tempListFile
        }
    } else {
        Write-Host "No .CR3 files found in this folder. Skipping." -ForegroundColor Yellow
    }

    # --- THE FIX: Step back out to the parent folder ---
    Pop-Location
    Write-Host ""
}

Write-Host "=====================================================" -ForegroundColor Green
Write-Host "All folders processed. The script has finished." -ForegroundColor Green
Write-Host "====================================================="