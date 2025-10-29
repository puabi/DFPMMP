# Requires: ImportExcel module
# Purpose: Generate dashlaneflags_mailmerge.xlsx from team-members.csv for users with weak passwords

# Get script directory
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Definition

# Define input/output paths
$csvPath = Join-Path $scriptDir 'team-members.csv'
$xlsxPath = Join-Path $scriptDir 'dashlaneflags_mailmerge.xlsx'

# Ensure ImportExcel module is loaded
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Host "Installing ImportExcel module..." -ForegroundColor Yellow
    Install-Module ImportExcel -Scope CurrentUser -Force
}

# Import CSV
if (-not (Test-Path $csvPath)) {
    Write-Error "File not found: $csvPath"
    exit 1
}

$data = Import-Csv -Path $csvPath

# Filter valid entries with numeric password_health_score < 60
$filtered = $data | Where-Object {
    [int]::TryParse($_.password_health_score, [ref]$null) -and
    ([int]$_.password_health_score -lt 60)
}

# Sort by password_health_score ascending
$filtered = $filtered | Sort-Object {[int]$_.password_health_score}

# Create objects for Excel
$mailMergeList = foreach ($entry in $filtered) {
    $email = $entry.'login email'
    if (-not [string]::IsNullOrWhiteSpace($email)) {
        # Extract first name (text before first period)
        $firstPart = ($email -split '\.')[0]
        if ($firstPart) {
            $firstName = ($firstPart.Substring(0,1).ToUpper() + $firstPart.Substring(1).ToLower())
        } else {
            $firstName = ""
        }

        [PSCustomObject]@{
            FirstName = $firstName
            Email     = $email
        }
    }
}

# Export to Excel
$mailMergeList | Export-Excel -Path $xlsxPath -WorksheetName "MailMerge" -AutoSize -AutoFilter -BoldTopRow

Write-Host "✅ Mail merge file created at: $xlsxPath" -ForegroundColor Green
