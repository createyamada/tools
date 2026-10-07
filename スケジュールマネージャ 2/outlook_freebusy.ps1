# Read classic Outlook free/busy data for the Python application.
$ErrorActionPreference = 'Stop'
[Console]::InputEncoding = New-Object System.Text.UTF8Encoding($false)
[Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)

function Get-FreeBusyView {
    param($Recipient, [string]$Label, [DateTime]$First, [DateTime]$Last)

    $view = New-Object System.Text.StringBuilder
    $day = $First
    while ($day -le $Last) {
        # FreeBusy returns about one month starting at midnight.
        $freeBusy = [string]$Recipient.FreeBusy($day, 30, $true)
        $availableDays = [Math]::Floor($freeBusy.Length / 48)
        if ($availableDays -lt 1) {
            throw "Free/busy data is unavailable: $Label ($($day.ToString('yyyy-MM-dd')))"
        }
        $remainingDays = ($Last - $day).Days + 1
        $daysToTake = [int][Math]::Min($availableDays, $remainingDays)
        $part = $freeBusy.Substring(0, $daysToTake * 48)
        if ($part -match '[^0-4]') {
            throw "Free/busy data has an unknown value: $Label"
        }
        [void]$view.Append($part)
        $day = $day.AddDays($daysToTake)
    }
    return $view.ToString()
}

try {
    $request = [Console]::In.ReadToEnd() | ConvertFrom-Json
    $culture = [System.Globalization.CultureInfo]::InvariantCulture
    $first = [DateTime]::ParseExact([string]$request.start_date, 'yyyy-MM-dd', $culture)
    $last = [DateTime]::ParseExact([string]$request.end_date, 'yyyy-MM-dd', $culture)
    if ($last -lt $first -or @($request.emails).Count -eq 0) {
        throw 'Invalid search conditions.'
    }

    $outlook = New-Object -ComObject Outlook.Application
    $mapi = $outlook.GetNamespace('MAPI')
    $selfRecipient = $mapi.CurrentUser
    if ($null -eq $selfRecipient) {
        throw 'Outlook could not identify the current user.'
    }
    $selfView = Get-FreeBusyView -Recipient $selfRecipient -Label 'current user' -First $first -Last $last
    $results = New-Object 'System.Collections.Generic.List[object]'

    foreach ($email in @($request.emails)) {
        $recipient = $mapi.CreateRecipient([string]$email)
        if (-not $recipient.Resolve()) {
            throw "Outlook could not resolve the email address: $email"
        }
        $view = Get-FreeBusyView -Recipient $recipient -Label ([string]$email) -First $first -Last $last
        [void]$results.Add([pscustomobject]@{ email = [string]$email; view = $view })
    }

    [pscustomobject]@{ ok = $true; self_view = $selfView; values = @($results.ToArray()) } |
        ConvertTo-Json -Depth 4 -Compress
} catch {
    [pscustomobject]@{ ok = $false; error = $_.Exception.Message } |
        ConvertTo-Json -Depth 4 -Compress
    exit 1
}
