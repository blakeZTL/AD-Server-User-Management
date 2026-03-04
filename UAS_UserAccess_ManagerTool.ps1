#Requires -Modules ImportExcel, ActiveDirectory

Import-Module ImportExcel
Import-Module ActiveDirectory

# -------------------------------
# Config
# -------------------------------
$group = "V-AFG001-APP-Part48-GASA-Access"
$filePath = ".\AFG Employee Roster.xlsx"
$sheetName = "Roster"
$emailColumnName = "Email Address Work"
$jobSeriesColumnName = "Occupational Series"
$targetJobSeries = @("1825", "1802")   # keep as string to avoid Excel numeric quirks

# If you want to actually apply changes, set to $true (requires permissions)
$applyChanges = $false

# Progress/log tuning (group is ~1500 now; this scales well as it grows)
$progressEvery = 100     # update Write-Progress every N items
$logEvery = 250     # heartbeat log every N items

# -------------------------------
# Logging helpers
# -------------------------------
$sw = [System.Diagnostics.Stopwatch]::StartNew()

function Log([string]$msg) {
    Write-Host ("[{0:hh\:mm\:ss}] {1}" -f $sw.Elapsed, $msg)
}

# Optional: creates a full console log file in the working directory
$transcriptPath = ".\GroupSync-$($group)-$(Get-Date -Format yyyyMMdd-HHmmss).log" -replace '[\\/:*?"<>|]', '_'
Start-Transcript -Path $transcriptPath | Out-Null
Log "Transcript: $transcriptPath"

try {
    # -------------------------------
    # Load and filter Excel rows
    # -------------------------------
    Log "Loading Excel: $filePath (sheet: $sheetName)..."
    $rows = Import-Excel $filePath -WorksheetName $sheetName
    Log "Excel rows loaded: $($rows.Count)"

    Log "Filtering rows where '$jobSeriesColumnName' in $($targetJobSeries -join ', ') ..."
    # build normalized set for fast, case-insensitive lookup
    $targetJobSeriesSet = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    foreach ($ts in $targetJobSeries) {
        if ($null -ne $ts -and -not [string]::IsNullOrWhiteSpace($ts)) { [void]$targetJobSeriesSet.Add($ts.ToString().Trim()) }
    }

    $filtered = $rows | Where-Object {
        $js = ($_.("$jobSeriesColumnName")).ToString().Trim()
        $targetJobSeriesSet.Contains($js)
    }
    Log "Filtered rows: $($filtered.Count)"

    # Unique email set (lowercased)
    $targetEmails = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)

    foreach ($r in $filtered) {
        $email = ($r.($emailColumnName)).ToString().Trim()
        if ([string]::IsNullOrWhiteSpace($email)) { continue }
        [void]$targetEmails.Add($email)
    }

    Log "Unique target emails: $($targetEmails.Count)"
    Write-Host ""

    # -------------------------------
    # Resolve target emails to AD users
    #   - Store by "key" (mail if present else UPN) for reference
    #   - Build target DN set for fast membership diffing
    # -------------------------------
    Log "Resolving target emails to AD users..."
    $targetsByKey = @{}   # key => ADUser
    $targetDns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $notFound = New-Object System.Collections.Generic.List[string]

    $i = 0
    $total = $targetEmails.Count

    foreach ($email in $targetEmails) {
        $i++

        if (($i % $progressEvery) -eq 0 -or $i -eq 1 -or $i -eq $total) {
            Write-Progress -Activity "Resolving Excel emails in AD" -Status "$i / $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
        }
        if (($i % $logEvery) -eq 0) {
            Log "Resolved $i / $total emails..."
        }

        $emailTrim = $email.Trim()
        if ([string]::IsNullOrWhiteSpace($emailTrim)) { continue }

        # Prefer mail match; fallback to UPN match
        $user = Get-ADUser -Filter "mail -eq '$emailTrim'" -Properties mail, UserPrincipalName, DistinguishedName -ErrorAction SilentlyContinue
        if ($null -eq $user) {
            $user = Get-ADUser -Filter "UserPrincipalName -eq '$emailTrim'" -Properties mail, UserPrincipalName, DistinguishedName -ErrorAction SilentlyContinue
        }

        if ($null -eq $user) {
            $notFound.Add($emailTrim) | Out-Null
            continue
        }

        # Key by mail if present else UPN
        $key = if ($user.mail) { $user.mail } else { $user.UserPrincipalName }

        # Keep last one if duplicates exist (rare, but possible)
        $targetsByKey[$key.ToLowerInvariant()] = $user

        if ($user.DistinguishedName) {
            [void]$targetDns.Add($user.DistinguishedName)
        }
    }

    Write-Progress -Activity "Resolving Excel emails in AD" -Completed
    Log "Resolved in AD: $($targetsByKey.Count) / $($targetEmails.Count)"
    Log "Target DN count: $($targetDns.Count)"
    if ($notFound.Count -gt 0) { Log "Not found in AD: $($notFound.Count)" }
    Write-Host ""

    # -------------------------------
    # Get current group members (FAST path)
    #   - Use DNs directly to diff membership
    #   - Avoid N+1 Get-ADUser calls
    # -------------------------------
    Log "Getting current members of group '$group' (recursive)..."
    $rawMembers = Get-ADGroupMember -Identity $group -Recursive
    Log "Retrieved raw members: $($rawMembers.Count) (users + groups + others)"

    $currentUserMembers = $rawMembers | Where-Object { $_.objectClass -eq "user" }
    Log "User objects in membership: $($currentUserMembers.Count)"

    $currentDns = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $i = 0
    $total = $currentUserMembers.Count

    foreach ($m in $currentUserMembers) {
        $i++
        if (($i % $progressEvery) -eq 0 -or $i -eq 1 -or $i -eq $total) {
            Write-Progress -Activity "Building current membership set (DNs)" -Status "$i / $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
        }
        if (($i % $logEvery) -eq 0) {
            Log "Processed $i / $total group user members..."
        }

        if ($m.DistinguishedName) {
            [void]$currentDns.Add($m.DistinguishedName)
        }
    }

    Write-Progress -Activity "Building current membership set (DNs)" -Completed
    Log "Current membership DN count: $($currentDns.Count)"
    Write-Host ""

    # -------------------------------
    # Compute delta: add / remove / skip
    # -------------------------------
    Log "Computing delta (target list is source of truth)..."

    $toAddDn = New-Object System.Collections.Generic.List[string]
    $toRemoveDn = New-Object System.Collections.Generic.List[string]
    $skippedCount = 0

    foreach ($dn in $targetDns) {
        if ($currentDns.Contains($dn)) { $skippedCount++ }
        else { $toAddDn.Add($dn) | Out-Null }
    }

    foreach ($dn in $currentDns) {
        if (-not $targetDns.Contains($dn)) { $toRemoveDn.Add($dn) | Out-Null }
    }

    Log "Delta computed."
    Log "  Would add   : $($toAddDn.Count)"
    Log "  Would remove: $($toRemoveDn.Count)"
    Log "  Skipped     : $skippedCount"
    Write-Host ""

    # -------------------------------
    # Apply changes (or dry run)
    # -------------------------------
    $addedCount = 0
    $removedCount = 0
    $addFailed = 0
    $removeFailed = 0

    if ($applyChanges) {
        Log "Applying changes to AD group (applyChanges = $applyChanges)..."
    }
    else {
        Log "Dry run only (applyChanges = $applyChanges). No changes will be made."
    }

    # Adds
    $i = 0
    $total = $toAddDn.Count
    foreach ($dn in $toAddDn) {
        $i++
        if (($i % $progressEvery) -eq 0 -or $i -eq 1 -or $i -eq $total) {
            Write-Progress -Activity "Adding members" -Status "$i / $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
        }

        try {
            if ($applyChanges) {
                Add-ADGroupMember -Identity $group -Members $dn -ErrorAction Stop
            }
            $addedCount++
            if (($i -le 10) -or (($i % $logEvery) -eq 0) -or ($i -eq $total)) {
                Log "ADD    : $dn"
            }
        }
        catch {
            $addFailed++
            Write-Warning "FAILED ADD $dn : $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "Adding members" -Completed

    # Removes
    $i = 0
    $total = $toRemoveDn.Count
    foreach ($dn in $toRemoveDn) {
        $i++
        if (($i % $progressEvery) -eq 0 -or $i -eq 1 -or $i -eq $total) {
            Write-Progress -Activity "Removing members" -Status "$i / $total" -PercentComplete ([int](100 * $i / [math]::Max(1, $total)))
        }

        try {
            if ($applyChanges) {
                Remove-ADGroupMember -Identity $group -Members $dn -Confirm:$false -ErrorAction Stop
            }
            $removedCount++
            if (($i -le 10) -or (($i % $logEvery) -eq 0) -or ($i -eq $total)) {
                Log "REMOVE : $dn"
            }
        }
        catch {
            $removeFailed++
            Write-Warning "FAILED REMOVE $dn : $($_.Exception.Message)"
        }
    }
    Write-Progress -Activity "Removing members" -Completed

    # -------------------------------
    # Summary
    # -------------------------------
    Write-Host ""
    Write-Host "========== SUMMARY =========="
    Write-Host "Group: $group"
    Write-Host "Target Job Series: $($targetJobSeries -join ', ')"
    Write-Host "Excel rows loaded: $($rows.Count)"
    Write-Host "Target rows (Job Series match): $($filtered.Count)"
    Write-Host "Target list (unique emails): $($targetEmails.Count)"
    Write-Host "Resolved in AD (unique users): $($targetsByKey.Count)"
    Write-Host "Not found in AD: $($notFound.Count)"
    Write-Host "Already in group (skipped): $skippedCount"
    Write-Host "Would add: $($toAddDn.Count)    Executed add count: $addedCount    Add failures: $addFailed"
    Write-Host "Would remove: $($toRemoveDn.Count) Executed remove count: $removedCount Remove failures: $removeFailed"
    Write-Host "Apply changes: $applyChanges"
    Write-Host "Elapsed: $($sw.Elapsed.ToString())"
    Write-Host "============================="

    if ($notFound.Count -gt 0) {
        Write-Host ""
        Write-Host "Not found emails (first 50 shown):"
        $notFound | Sort-Object | Select-Object -First 50 | ForEach-Object { Write-Host "  - $_" }

        if ($notFound.Count -gt 50) {
            Write-Host "  ... and $($notFound.Count - 50) more"
        }
    }

}
finally {
    Stop-Transcript | Out-Null
    Log "Done."
}