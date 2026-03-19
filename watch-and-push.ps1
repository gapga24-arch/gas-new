# Watch .gs / .html / .json and run clasp push on change
# Usage: .\watch-and-push.ps1 (keep this terminal open)

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$projectRoot = $PSScriptRoot
Set-Location $projectRoot

Write-Host "Running initial clasp push..." -ForegroundColor Yellow
& clasp push 2>&1
if ($LASTEXITCODE -eq 0) { Write-Host "OK" -ForegroundColor Green } else { Write-Host "Check clasp login / network" -ForegroundColor Red }
Write-Host ""

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $projectRoot
$watcher.Filter = "*.*"
$watcher.IncludeSubdirectories = $false
$watcher.EnableRaisingEvents = $true

$debounceSeconds = 2
$global:lastPushTime = [DateTime]::MinValue

$action = {
    $name = $Event.SourceEventArgs.Name
    $changeType = $Event.SourceEventArgs.ChangeType

    if ($name -match '\.(gs|html|json)$' -and ($changeType -eq 'Changed' -or $changeType -eq 'Created')) {
        $now = Get-Date
        if (($now - $global:lastPushTime).TotalSeconds -lt 2) {
            return
        }
        $global:lastPushTime = $now
        Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Detected change: $name -> running clasp push..."
        & clasp push 2>&1
        if ($LASTEXITCODE -eq 0) {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] Pushed to GAS OK" -ForegroundColor Green
        } else {
            Write-Host "[$(Get-Date -Format 'HH:mm:ss')] clasp push failed" -ForegroundColor Red
        }
    }
}

$null = Register-ObjectEvent -InputObject $watcher -EventName Changed -Action $action
$null = Register-ObjectEvent -InputObject $watcher -EventName Created -Action $action

Write-Host "Watching for changes. Save code.gs / index.html / appsscript.json to push to GAS." -ForegroundColor Cyan
Write-Host "Stop with Ctrl+C" -ForegroundColor Gray
try {
    while ($true) { Start-Sleep -Seconds 1 }
} finally {
    Get-EventSubscriber | Unregister-Event
    $watcher.Dispose()
}
