# Path to your bookmarks file
$bookmarkFile = "./bookmarks.html"

# Log file with current date
$logFile = "./access_report_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

# Ensure log file is empty at start
New-Item -Path $logFile -ItemType File -Force | Out-Null

# Retry settings
$maxRetries = 3
$retryDelaySeconds = 2

# Function to write to console and log
function Write-Log {
    param (
        [string]$Message,
        [ConsoleColor]$Color = [ConsoleColor]::White
    )
    Write-Host $Message -ForegroundColor $Color
    Add-Content -Path $logFile -Value $Message
}

# Read the file
$html = Get-Content -Path $bookmarkFile -Raw

# Regex patterns
$groupPattern = '<H3[^>]*>([^<]+)</H3>'
$urlPattern   = '<A[^>]+HREF="([^"]+)"[^>]*>([^<]+)</A>'

# Stack for nested groups
$groupStack = New-Object System.Collections.Stack

Write-Log "============================= Network Access Check =============================" Gray
Write-Log "[Make sure there is a bookmarks.html file in the same directory as this script.]`n" Gray

foreach ($line in ($html -split "`n")) {

    # Detect group (folder start)
    if ($line -match $groupPattern) {
        $title = $matches[1].Trim()
        
        # Build full path including this folder
        $fullPath = if ($groupStack.Count -gt 0) {
            (@($groupStack.ToArray())[-1..0] -join " > ") + " > $title"
        } else {
            $title
        }

        Write-Log "`n### $fullPath [Folder start]" Cyan

        # Push the folder onto the stack after logging
        $groupStack.Push($title)
        continue
    }

    # Detect URL
    if ($line -match $urlPattern) {
        $url   = $matches[1].Trim()
        $label = $matches[2].Trim()
        $attempt = 0
        $totalDuration = 0
        $finalStatus = ""
        $errorMsg = ""
        $ipAddress = "N/A"

        # Attempt to get IP address
        try {
            $uri = [System.Uri] $url
            $ipAddresses = [System.Net.Dns]::GetHostAddresses($uri.Host)
            if ($ipAddresses.Length -gt 0) { $ipAddress = $ipAddresses[0].ToString() }
        } catch {
            $ipAddress = "Unable to resolve"
        }

        # Retry loop
        while ($attempt -lt $maxRetries) {
            $attempt++
            $startTime = Get-Date
            try {
                $resp = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 10
                $endTime = Get-Date
                $totalDuration += ($endTime - $startTime).TotalSeconds

                if ($resp.StatusCode -ge 200 -and $resp.StatusCode -lt 400) {
                    $finalStatus = "OK"
                } else {
                    $finalStatus = "FAIL"
                }

                break
            } catch {
                $endTime = Get-Date
                $totalDuration += ($endTime - $startTime).TotalSeconds
                $errorMsg = $_.Exception.Message

                if ($errorMsg -match "trust relationship") {
                    $finalStatus = "SSL"
                    break
                } else {
                    $finalStatus = "ERROR"
                    if ($attempt -lt $maxRetries) { Start-Sleep -Seconds $retryDelaySeconds }
                }
            }
        }

        # Use full folder path for URL
        $fullPath = @($groupStack.ToArray())[-1..0] -join " > "

        # Log final result including IP
        switch ($finalStatus) {
            "OK"    { Write-Log ("- {0,-5} [{1}] {2} -> Success | IP: {3} | Attempts: {4} | Duration: {5:N2}s" -f $finalStatus, $label, $url, $ipAddress, $attempt, $totalDuration) Green }
            "FAIL"  { Write-Log ("- {0,-5} [{1}] {2} -> HTTP error | IP: {3} | Attempts: {4} | Duration: {5:N2}s" -f $finalStatus, $label, $url, $ipAddress, $attempt, $totalDuration) Red }
            "SSL"   { Write-Log ("- {0,-5} [{1}] {2} -> SSL error | IP: {3} | Attempts: {4} | Duration: {5:N2}s" -f $finalStatus, $label, $url, $ipAddress, $attempt, $totalDuration) DarkYellow }
            "ERROR" { Write-Log ("- {0,-5} [{1}] {2} -> {4} | IP: {3} | Attempts: {5} | Duration: {6:N2}s" -f $finalStatus, $label, $url, $ipAddress, $errorMsg, $attempt, $totalDuration) Red }
        }
        continue
    }

    # Detect folder end
    if ($line -match '</DL>') {
        if ($groupStack.Count -gt 0) {
            # Pop the folder we are closing
            $folderName = $groupStack.Pop()

            # Build full path including this folder
            $fullPath = if ($groupStack.Count -gt 0) {
                (@($groupStack.ToArray())[-1..0] -join " > ") + " > $folderName"
            } else {
                $folderName
            }

            Write-Log "### $fullPath [Folder end]" Cyan
        }
        continue
    }
}

Write-Log "`n=== Test Finished. Press any key to exit... ===" Gray
[void][System.Console]::ReadKey($true)