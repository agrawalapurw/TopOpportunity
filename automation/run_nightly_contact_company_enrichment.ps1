param(
    [string]$PythonExe = "python",
    [string]$ProjectRoot = "C:\Users\agrawalapurw\Documents\All_Projects\BDP\contact_company_enrichment",
    [string]$ShareRoot = "\\SINSDV030.infineon.com\DEM_AP_MM_SALES\01_SHARING\Digitalisation\AI_Automation\Top_Opportunity",
    [string]$DiveUid = "",
    [string]$DivePwd = "",
    [switch]$EnableEmailNotifications,
    [string]$MailFrom = "R-IFX-DEMMA@infineon.com",
    [string]$MailTo = "",
    [string]$MailCc = "",
    [string]$SmtpServer = "",
    [int]$SmtpPort = 25,
    [switch]$UseSsl
)

$ErrorActionPreference = "Stop"

$scriptPath = Join-Path $ProjectRoot "run_contact_company_enrichment.py"
$inputDir = Join-Path $ShareRoot "input"
$outputDir = Join-Path $ShareRoot "output"
$archiveRoot = Join-Path $ShareRoot "archive"
$logDir = Join-Path $ProjectRoot "logs"

if (-not (Test-Path $scriptPath)) {
    throw "Runner script not found at: $scriptPath"
}
if (-not (Test-Path $inputDir)) {
    throw "Input folder not found: $inputDir"
}
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}
if (-not (Test-Path $archiveRoot)) {
    New-Item -ItemType Directory -Path $archiveRoot -Force | Out-Null
}
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -Force | Out-Null
}

$runStamp = Get-Date -Format "yyyyMMdd_HHmmss"
$runArchiveDir = Join-Path $archiveRoot $runStamp
$logFile = Join-Path $logDir "nightly_run_$runStamp.log"

function Get-LatestOutputForWorkbook {
    param(
        [Parameter(Mandatory = $true)] [string]$OutputFolder,
        [Parameter(Mandatory = $true)] [string]$WorkbookName,
        [Parameter(Mandatory = $true)] [datetime]$RunStart
    )

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($WorkbookName)
    $pattern = "$stem*_company_enriched*.xlsx"
    $candidates = Get-ChildItem -Path $OutputFolder -Filter $pattern -File -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -ge $RunStart } |
        Sort-Object LastWriteTime -Descending

    if ($candidates -and $candidates.Count -gt 0) {
        return $candidates[0].FullName
    }

    return ""
}

function Send-OutputReadyEmail {
    param(
        [Parameter(Mandatory = $true)] [string]$From,
        [Parameter(Mandatory = $true)] [string[]]$To,
        [string[]]$Cc,
        [Parameter(Mandatory = $true)] [string]$Server,
        [Parameter(Mandatory = $true)] [int]$Port,
        [Parameter(Mandatory = $true)] [bool]$EnableSsl,
        [Parameter(Mandatory = $true)] [string]$InputFileName,
        [Parameter(Mandatory = $true)] [string]$OutputFilePath,
        [Parameter(Mandatory = $true)] [string]$ShareOutputFolder
    )

    $mail = New-Object System.Net.Mail.MailMessage
    $mail.From = $From
    foreach ($addr in $To) {
        if ($addr) { [void]$mail.To.Add($addr) }
    }
    foreach ($addr in $Cc) {
        if ($addr) { [void]$mail.CC.Add($addr) }
    }

    $outputName = [System.IO.Path]::GetFileName($OutputFilePath)
    $mail.Subject = "Top Opportunity output ready: $outputName"
    $mail.Body = @"
Hello,

The enrichment output file is ready.

Input file: $InputFileName
Output file: $outputName
Output folder: $ShareOutputFolder
Generated at: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")

Regards,
Top Opportunity Automation
"@

    $smtp = New-Object System.Net.Mail.SmtpClient($Server, $Port)
    $smtp.EnableSsl = $EnableSsl
    $smtp.Send($mail)
}

Start-Transcript -Path $logFile | Out-Null

try {
    Set-Location $ProjectRoot

    if ($DiveUid) {
        $env:DIVE_UID = $DiveUid
    }
    if ($DivePwd) {
        $env:DIVE_PWD = $DivePwd
    }

    if ($EnableEmailNotifications) {
        if ([string]::IsNullOrWhiteSpace($SmtpServer)) {
            throw "Email notifications enabled but SmtpServer is empty."
        }
        if ([string]::IsNullOrWhiteSpace($MailTo)) {
            throw "Email notifications enabled but MailTo is empty."
        }
    }

    $batchFiles = Get-ChildItem -Path $inputDir -Filter "*.xlsx" -File |
        Where-Object { -not $_.Name.StartsWith("~$") } |
        Sort-Object Name

    if (-not $batchFiles -or $batchFiles.Count -eq 0) {
        Write-Host "No input files found. Exiting successfully."
        exit 0
    }

    New-Item -ItemType Directory -Path $runArchiveDir -Force | Out-Null

    $processed = 0
    $failed = 0

    foreach ($file in $batchFiles) {
        Write-Host "Processing file: $($file.Name)"
        $fileRunStart = Get-Date

        $args = @(
            $scriptPath,
            "--input-dir", $inputDir,
            "--output-dir", $outputDir,
            "--workbook", $file.Name
        )

        & $PythonExe @args
        $code = $LASTEXITCODE

        if ($code -eq 0) {
            $destination = Join-Path $runArchiveDir $file.Name
            Move-Item -Path $file.FullName -Destination $destination -Force
            Write-Host "Success: archived input file to $destination"

            if ($EnableEmailNotifications) {
                try {
                    $latestOutput = Get-LatestOutputForWorkbook -OutputFolder $outputDir -WorkbookName $file.Name -RunStart $fileRunStart
                    if (-not [string]::IsNullOrWhiteSpace($latestOutput)) {
                        $toList = $MailTo.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        $ccList = $MailCc.Split(",") | ForEach-Object { $_.Trim() } | Where-Object { $_ }
                        Send-OutputReadyEmail -From $MailFrom -To $toList -Cc $ccList -Server $SmtpServer -Port $SmtpPort -EnableSsl:$UseSsl.IsPresent -InputFileName $file.Name -OutputFilePath $latestOutput -ShareOutputFolder $outputDir
                        Write-Host "Notification email sent for $($file.Name)"
                    }
                    else {
                        Write-Host "Warning: output file not detected for email notification: $($file.Name)"
                    }
                }
                catch {
                    Write-Host "Warning: failed to send email for $($file.Name). Error: $($_.Exception.Message)"
                }
            }

            $processed++
        }
        else {
            Write-Host "Failed (exit code $code): $($file.Name)"
            $failed++
        }
    }

    Write-Host "Run complete. Processed: $processed, Failed: $failed"

    if ($failed -gt 0) {
        exit 1
    }

    exit 0
}
catch {
    Write-Error $_
    exit 1
}
finally {
    Stop-Transcript | Out-Null
}
