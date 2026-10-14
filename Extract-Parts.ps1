# Extract-Parts.ps1
# Windows PowerShell 5.1 compatible.
# Produced and iteratively refined with GitHub Copilot.
# Features:
# - Adaptively schedules RAR archive sets using Start-Job with no fixed job ceiling.
# - Flatten extracted files into a user-selected destination folder.
# - Convert extracted and previously-existing .nsp/.xci files to .nsz/.xcz with adaptive concurrency.
# - Ask before installing missing dependencies (WinRAR via winget; NSZ via its official PyPI package in a virtual environment).
# - Ask the user to select an existing legal prod.keys/keys.txt file if none is found.
# - Show a color-coded live dashboard with queue previews and tool-reported progress when available.
# - Remove source RAR volume sets only when extraction, conversion (when applicable), and destination copy succeed.
# - Optionally delete standalone .nsp/.xci originals only after a verified output has been copied successfully.
# - Timestamped diagnostic log, transcript, and per-archive WinRAR logs for troubleshooting.
#
# Usage examples:
#   .\Extract-Parts.ps1
#   .\Extract-Parts.ps1 -SourceFolder "C:\Games\Roms" -OutputFolder "C:\Games\Output" -DeleteArchives -DeleteStandalone
#   .\Extract-Parts.ps1 -SourceFolder "C:\Games\Roms" -OutputFolder "C:\Games\Output" -NonInteractive -DryRun
#   .\Extract-Parts.ps1 -SourceFolder "C:\Games\Roms" -OutputFolder "C:\Games\Output" -NonInteractive -DeleteArchives
#   .\Extract-Parts.ps1 -SourceFolder "C:\Games\Roms" -OutputFolder "C:\Games\Output" -FixedConcurrency
#
# Behavior:
#   - If command-line flags are provided, they take precedence.
#   - Otherwise, the script asks interactively to confirm deletions.
#   - Non-interactive runs default to safe behavior and skip destructive cleanup unless explicitly enabled.
#   - Dry-run mode scans inputs without installing dependencies or processing files.

# --------------------------- Configuration ---------------------------
param(
    [string]$SourceFolder = "",
    [string]$OutputFolder,
    [switch]$DeleteArchives,
    [switch]$DeleteStandalone,
    [switch]$DryRun,
    [switch]$NoDelete,
    [switch]$NonInteractive,
    [switch]$Force,
    [switch]$FixedConcurrency,
    [switch]$NoCompression,
    [switch]$QuickVerify,
    [ValidateRange(1, 22)][int]$CompressionLevel = 18
)

$MaxParallel = 6
$MaxParallelConversions = 2
$script:AdaptiveConcurrencyEnabled = -not [bool]$FixedConcurrency
$script:EffectiveMaxParallel = if ($script:AdaptiveConcurrencyEnabled) { 1 } else { $MaxParallel }
$script:EffectiveMaxParallelConversions = if ($script:AdaptiveConcurrencyEnabled) { 1 } else { $MaxParallelConversions }
$script:AdaptiveNextSampleAt = [datetime]::MinValue
$script:AdaptiveHealthySamples = 0
$script:AdaptivePressureSamples = 0
$script:AdaptiveNextScaleStage = 'Extraction'
$script:AdaptiveMetrics = $null
$script:AdaptiveMetricsJob = $null
$script:AdaptiveStatus = if ($script:AdaptiveConcurrencyEnabled) { 'Sampling system resources' } else { 'Fixed concurrency limits selected' }
$script:AdaptiveMetricsUnavailable = $false
$script:AdaptiveLowSpace = $false
$script:DiskReserveBlocked = $false
$script:DiskReserveUnavailable = $false
$script:DiskReserveNextSampleAt = [datetime]::MinValue
$script:DiskReserveVolumes = @()
$script:CancellationKeyUnavailable = $false
$script:PendingConversions = 0
$script:PendingConversionRecords = @()
$script:PendingConversionNextIndex = 0
$QueuePreviewLimit = 5
$DeleteArchivesOnSuccess = $null
$DefaultOutputFolder = if ([string]::IsNullOrWhiteSpace($SourceFolder)) { Join-Path $env:USERPROFILE 'Extracted' } else { Join-Path $SourceFolder 'Extracted' }
$NszInstallFolder = Join-Path $env:LOCALAPPDATA "ExtractParts\tools"
$NszExecutableName = "nsz-cli-windows-x64.exe"
$NszPythonVenvFolder = Join-Path $NszInstallFolder "venv"

# ------------------------------ Setup -------------------------------
$ErrorActionPreference = 'Stop'
$script:RarExe = $null
$script:NszExe = $null
$script:KeysFile = $null
$script:CompressionEnabled = $false
$script:CompressionRequested = -not [bool]$NoCompression
$script:QuickVerify = [bool]$QuickVerify
$script:ProcessingFinishedArchiveEntries = $false
$script:DeleteStandaloneSources = $false
$script:DeleteArchivesConfirmed = $false
$script:DryRun = [bool]$DryRun
$script:NoDelete = [bool]$NoDelete
$script:NonInteractive = [bool]$NonInteractive
$script:Force = [bool]$Force

if ($DeleteArchives) {
    $DeleteArchivesOnSuccess = $true
}
if ($DeleteStandalone) {
    $script:DeleteStandaloneSources = $true
}
if ($script:DryRun -or $script:NoDelete) {
    $DeleteArchivesOnSuccess = $false
    $script:DeleteStandaloneSources = $false
}
$script:StopProcessing = $false
$script:StopReason = $null
$script:LastCopyDestinationPath = $null
$script:LastCopyDestinationHash = $null
$script:ProgressDashboardTop = $null
$script:ProgressDashboardLineCount = 0
$script:ProgressDashboardRawUI = $null
$script:ProgressDashboardInteractive = $false
$script:ProgressDashboardCursorFailed = $false
$script:RunStartedAt = Get-Date
$script:DashboardState = [pscustomobject]@{
    CurrentWork = 'Preparing'
    OperationPercent = $null
    OperationLabel = $null
    Rows = @()
    LastRenderedAt = $null
}
$script:ProgressStage = 'Archives'
$script:Active = @()
$script:ActiveConversions = @()
$script:NextArchiveIndex = 0
$script:CompletedArchives = 0
$script:CopiedCount = 0
$script:SkippedCount = 0
$script:FailedCopyCount = 0
$script:ErrorCount = 0
$script:WarningCount = 0
$script:RunExitCode = 0
$script:DeletedArchiveCount = 0
$script:ConversionSuccessCount = 0
$script:ConversionErrorCount = 0
$script:StandaloneProcessedCount = 0
$script:StandaloneQueue = @()
$script:StandaloneNextIndex = 0
$script:TotalArchives = 0
$script:TotalStandalone = 0
$script:Archives = @()
$script:RunLogFolder = Join-Path $env:LOCALAPPDATA 'ExtractParts\Logs'
$script:RunId = '{0}_{1}' -f (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), ([guid]::NewGuid().ToString('N').Substring(0,8))
try {
    New-Item -ItemType Directory -Path $script:RunLogFolder -Force -ErrorAction Stop | Out-Null
} catch {
    $script:RunLogFolder = Join-Path $env:TEMP 'ExtractParts_Logs'
    New-Item -ItemType Directory -Path $script:RunLogFolder -Force -ErrorAction SilentlyContinue | Out-Null
}
$script:DiagnosticsLogPath = Join-Path $script:RunLogFolder ("ExtractParts_{0}.log" -f $script:RunId)
$script:TranscriptPath = Join-Path $script:RunLogFolder ("ExtractParts_{0}_transcript.txt" -f $script:RunId)
$script:CurrentProcessStartTimeUtcTicks = (Get-Process -Id $PID -ErrorAction Stop).StartTime.ToUniversalTime().Ticks
$script:TranscriptStarted = $false
$script:LogWriteFailureReported = $false
$script:WorkerLogCounter = 0
try {
    Start-Transcript -Path $script:TranscriptPath -Force | Out-Null
    $script:TranscriptStarted = $true
} catch {
    # Diagnostics logging remains available if transcripts are unsupported in this host.
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet('DEBUG','INFO','WARN','ERROR')][string]$Level = 'INFO'
    )
    $line = '[{0}][{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Level, $Message
    try {
        Add-Content -LiteralPath $script:DiagnosticsLogPath -Value $line -Encoding UTF8 -ErrorAction Stop
    } catch {
        if (-not $script:LogWriteFailureReported) {
            $script:LogWriteFailureReported = $true
            Write-Warning "Could not write diagnostic log '$script:DiagnosticsLogPath': $($_.Exception.Message)"
        }
    }
}

function Stop-RunLogging {
    if ($script:TranscriptStarted) {
        try {
            Stop-Transcript | Out-Null
        } catch {
            Write-Log -Message "Could not stop the run transcript: $($_.Exception.Message)" -Level WARN
        }
        $script:TranscriptStarted = $false
    }
}

function Get-AdaptiveResourceSnapshot {
    $cpu = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'" -ErrorAction Stop |
        Select-Object -First 1
    $memory = Get-CimInstance -ClassName Win32_OperatingSystem -ErrorAction Stop
    $disk = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'" -ErrorAction Stop |
        Select-Object -First 1
    if ($null -eq $cpu -or $null -eq $memory -or $null -eq $disk) {
        throw 'Windows returned incomplete CPU, memory, or disk performance data.'
    }
    if ($memory.TotalVisibleMemorySize -le 0 -or $null -eq $memory.FreePhysicalMemory) {
        throw 'Windows returned invalid physical-memory totals.'
    }

    $roots = @()
    foreach ($path in @($script:OutputFolder, $env:TEMP)) {
        if ([string]::IsNullOrWhiteSpace($path)) { continue }
        $root = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($path))
        if ($root -notmatch '^[A-Za-z]:\\$') {
            throw "Cannot safely measure free space for storage path '$path'."
        }
        $root = $root.TrimEnd('\')
        if ($roots -notcontains $root) { $roots += $root }
    }
    if ($roots.Count -eq 0) { throw 'Could not identify the output and temporary storage volumes.' }

    $logicalDisks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3' -ErrorAction Stop)
    $storageVolumes = @()
    foreach ($root in $roots) {
        $volume = $logicalDisks | Where-Object { $_.DeviceID -ieq $root } | Select-Object -First 1
        if ($null -eq $volume -or $volume.Size -le 0 -or $null -eq $volume.FreeSpace) {
            throw "Windows could not report free space for volume '$root'."
        }
        $reserveBytes = [long][Math]::Max(10GB, [double]$volume.Size * 0.05)
        $freeBytes = [long]$volume.FreeSpace
        $storageVolumes += [pscustomobject]@{
            DeviceID = [string]$volume.DeviceID
            FreePercent = 100.0 * $freeBytes / [double]$volume.Size
            BelowReserve = $freeBytes -lt $reserveBytes
        }
    }

    $minimumFreePercent = ($storageVolumes | Measure-Object -Property FreePercent -Minimum).Minimum
    return [pscustomobject]@{
        CpuPercent = [Math]::Min(100, [Math]::Max(0, [double]$cpu.PercentProcessorTime))
        MemoryFreePercent = 100.0 * [double]$memory.FreePhysicalMemory / [double]$memory.TotalVisibleMemorySize
        DiskBusyPercent = [Math]::Min(100, [Math]::Max(0, [double]$disk.PercentDiskTime))
        DiskQueueLength = [double]$disk.CurrentDiskQueueLength
        MinimumDiskFreePercent = [double]$minimumFreePercent
        BelowDiskReserve = [bool](@($storageVolumes | Where-Object { $_.BelowReserve }).Count -gt 0)
    }
}

function Update-DiskReserveStatus {
    param([switch]$ForceRefresh)
    if (-not $ForceRefresh -and (Get-Date) -lt $script:DiskReserveNextSampleAt) { return }
    $script:DiskReserveNextSampleAt = (Get-Date).AddSeconds(15)
    try {
        $roots = @()
        foreach ($path in @($script:OutputFolder, $env:TEMP)) {
            if ([string]::IsNullOrWhiteSpace($path)) { continue }
            $root = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($path))
            if ($root -notmatch '^[A-Za-z]:\\$') {
                throw "Cannot safely measure free space for storage path '$path'."
            }
            $root = $root.TrimEnd('\')
            if ($roots -notcontains $root) { $roots += $root }
        }
        if ($roots.Count -eq 0) { throw 'Could not identify output and temporary storage volumes.' }
        $volumes = @()
        foreach ($root in $roots) {
            $drive = [System.IO.DriveInfo]::new("$root\")
            if (-not $drive.IsReady -or $drive.TotalSize -le 0) {
                throw "Windows could not report free space for volume '$root'."
            }
            $reserveBytes = [long][Math]::Max(10GB, [double]$drive.TotalSize * 0.05)
            $freeBytes = [long]$drive.AvailableFreeSpace
            $volumes += [pscustomobject]@{
                DeviceID = [string]$root
                FreeBytes = $freeBytes
                ReserveBytes = $reserveBytes
                BelowReserve = $freeBytes -lt $reserveBytes
            }
        }
        $wasBlocked = $script:DiskReserveBlocked
        $wasUnavailable = $script:DiskReserveUnavailable
        $script:DiskReserveVolumes = $volumes
        $script:DiskReserveBlocked = [bool](@($volumes | Where-Object { $_.BelowReserve }).Count -gt 0)
        $script:DiskReserveUnavailable = $false
        if ($script:DiskReserveBlocked -and -not $wasBlocked) {
            Write-Log -Message ("Free disk space is below the 10 GB / 5% safety reserve; no new extraction or conversion jobs will start. {0}" -f (($volumes | ForEach-Object { "$($_.DeviceID) free=$([Math]::Round($_.FreeBytes / 1GB, 1)) GB reserve=$([Math]::Round($_.ReserveBytes / 1GB, 1)) GB" }) -join '; ')) -Level WARN
            if ($script:NonInteractive) {
                if (-not $script:StopProcessing) {
                    $script:StopProcessing = $true
                    $script:StopReason = 'disk safety reserve'
                    $script:ErrorCount++
                }
            }
        } elseif (-not $script:DiskReserveBlocked -and $wasBlocked) {
            Write-Log -Message 'Free disk space is above the safety reserve again; queued work may resume.' -Level INFO
        } elseif ($wasUnavailable) {
            Write-Log -Message 'Free disk space can be measured again; the disk reserve gate is active.' -Level INFO
        }
    } catch {
        $wasUnavailable = $script:DiskReserveUnavailable
        $script:DiskReserveUnavailable = $true
        $script:DiskReserveBlocked = $false
        $script:DiskReserveVolumes = @()
        if (-not $wasUnavailable) {
            Write-Log -Message "Could not verify free disk space; pausing new work until the reserve can be checked. $($_.Exception.Message)" -Level WARN
        }
        if ($script:NonInteractive) {
            if (-not $script:StopProcessing) {
                $script:StopProcessing = $true
                $script:StopReason = 'disk space could not be verified'
                $script:ErrorCount++
            }
        }
    }
}

function Update-AdaptiveConcurrency {
    if (-not $script:AdaptiveConcurrencyEnabled) { return }

    $metrics = $null
    if ($script:AdaptiveMetricsJob) {
        if ($script:AdaptiveMetricsJob.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) { return }
        try {
            $metrics = @(Receive-Job -Job $script:AdaptiveMetricsJob -ErrorAction Stop |
                Where-Object { $_ -and $_.Kind -eq 'Metrics' } | Select-Object -Last 1)[0]
            if ($script:AdaptiveMetricsJob.State -ne 'Completed' -or $null -eq $metrics) {
                throw 'The asynchronous Windows resource sample did not return metrics.'
            }
        } catch {
            $metrics = $null
            $metricError = $_.Exception.Message
        } finally {
            Remove-BackgroundJob -Job $script:AdaptiveMetricsJob
            $script:AdaptiveMetricsJob = $null
        }
    } elseif ((Get-Date) -ge $script:AdaptiveNextSampleAt) {
        $script:AdaptiveNextSampleAt = (Get-Date).AddSeconds(15)
        $sampleScript = {
            param([string[]]$Paths)
            $ErrorActionPreference = 'Stop'
            $cpu = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfOS_Processor -Filter "Name='_Total'"
            $memory = Get-CimInstance -ClassName Win32_OperatingSystem
            $disk = Get-CimInstance -ClassName Win32_PerfFormattedData_PerfDisk_PhysicalDisk -Filter "Name='_Total'"
            if ($null -eq $cpu -or $null -eq $memory -or $null -eq $disk) {
                throw 'Windows returned incomplete CPU, memory, or disk performance data.'
            }
            $roots = @()
            foreach ($path in $Paths) {
                if ([string]::IsNullOrWhiteSpace($path)) { continue }
                $root = [System.IO.Path]::GetPathRoot([System.IO.Path]::GetFullPath($path))
                if ($root -notmatch '^[A-Za-z]:\\$') { throw "Cannot safely measure storage path '$path'." }
                $root = $root.TrimEnd('\')
                if ($roots -notcontains $root) { $roots += $root }
            }
            if ($roots.Count -eq 0) { throw 'Could not identify output and temporary storage volumes.' }
            $logicalDisks = @(Get-CimInstance -ClassName Win32_LogicalDisk -Filter 'DriveType=3')
            $volumes = @()
            foreach ($root in $roots) {
                $volume = $logicalDisks | Where-Object { $_.DeviceID -ieq $root } | Select-Object -First 1
                if ($null -eq $volume -or $volume.Size -le 0 -or $null -eq $volume.FreeSpace) {
                    throw "Windows could not report free space for volume '$root'."
                }
                $reserveBytes = [long][Math]::Max(10GB, [double]$volume.Size * 0.05)
                $volumes += [pscustomobject]@{
                    DeviceID = [string]$volume.DeviceID
                    FreePercent = 100.0 * [double]$volume.FreeSpace / [double]$volume.Size
                    BelowReserve = [long]$volume.FreeSpace -lt $reserveBytes
                }
            }
            if ($memory.TotalVisibleMemorySize -le 0 -or $null -eq $memory.FreePhysicalMemory) {
                throw 'Windows returned invalid physical-memory totals.'
            }
            [pscustomobject]@{
                Kind = 'Metrics'
                CpuPercent = [Math]::Min(100, [Math]::Max(0, [double]$cpu.PercentProcessorTime))
                MemoryFreePercent = 100.0 * [double]$memory.FreePhysicalMemory / [double]$memory.TotalVisibleMemorySize
                DiskBusyPercent = [Math]::Min(100, [Math]::Max(0, [double]$disk.PercentDiskTime))
                DiskQueueLength = [double]$disk.CurrentDiskQueueLength
                MinimumDiskFreePercent = [double](($volumes | Measure-Object -Property FreePercent -Minimum).Minimum)
                BelowDiskReserve = [bool](@($volumes | Where-Object BelowReserve).Count -gt 0)
            }
        }
        try {
            $script:AdaptiveMetricsJob = Start-Job -ScriptBlock $sampleScript -ArgumentList (, @($script:OutputFolder, $env:TEMP)) -ErrorAction Stop
        } catch {
            $metricError = $_.Exception.Message
        }
        if ($script:AdaptiveMetricsJob) { return }
        $metrics = $null
    } else {
        return
    }

    if ($null -eq $metrics) {
        $wasUnavailable = $script:AdaptiveMetricsUnavailable
        $script:AdaptiveMetricsUnavailable = $true
        $script:AdaptiveLowSpace = $false
        $script:AdaptiveMetrics = $null
        $script:EffectiveMaxParallel = 1
        $script:EffectiveMaxParallelConversions = 1
        $script:AdaptiveHealthySamples = 0
        $script:AdaptivePressureSamples = 0
        $script:AdaptiveStatus = 'Resource metrics unavailable; continuing at 1 RAR + 1 NSZ slot'
        if (-not $wasUnavailable) {
            Write-Log -Message "Adaptive concurrency could not read Windows resource metrics; continuing without stopping at one extraction and one conversion slot. $metricError" -Level WARN
        }
        return
    }

    if ($script:AdaptiveMetricsUnavailable) {
        Write-Log -Message 'Windows resource metrics are available again; adaptive concurrency has resumed.' -Level INFO
    }
    $script:AdaptiveMetricsUnavailable = $false
    $script:AdaptiveMetrics = $metrics
    $wasLowSpace = $script:AdaptiveLowSpace
    $script:AdaptiveLowSpace = [bool]$metrics.BelowDiskReserve

    if ($metrics.BelowDiskReserve) {
        $script:EffectiveMaxParallel = 1
        $script:EffectiveMaxParallelConversions = 1
        $script:AdaptiveHealthySamples = 0
        $script:AdaptivePressureSamples = 0
        $script:AdaptiveStatus = 'Low disk reserve; continuing with one slot per workload'
        if (-not $wasLowSpace) {
            Write-Log -Message ("Free space is below the 10 GB / 5% safety reserve on an output or temporary volume; limiting new work to one extraction and one conversion. Minimum free space={0:N1} percent." -f $metrics.MinimumDiskFreePercent) -Level WARN
        }
        return
    }
    if ($wasLowSpace) {
        Write-Log -Message 'Free space on the output and temporary volumes is above the adaptive concurrency reserve again.' -Level INFO
    }

    $cpuPressure = $metrics.CpuPercent -ge 85
    $memoryPressure = $metrics.MemoryFreePercent -le 15
    $diskPressure = $metrics.DiskBusyPercent -ge 90 -or $metrics.DiskQueueLength -ge 4
    $severePressure = $metrics.CpuPercent -ge 95 -or $metrics.MemoryFreePercent -le 8 -or
        $metrics.DiskBusyPercent -ge 98 -or $metrics.DiskQueueLength -ge 8
    $pressure = $cpuPressure -or $memoryPressure -or $diskPressure
    $healthy = $metrics.CpuPercent -le 65 -and $metrics.MemoryFreePercent -ge 25 -and
        $metrics.DiskBusyPercent -le 70 -and $metrics.DiskQueueLength -lt 4

    if ($pressure) {
        $script:AdaptiveHealthySamples = 0
        $script:AdaptivePressureSamples++
        if ($severePressure -or $script:AdaptivePressureSamples -ge 2) {
            $changed = $false
            if ($diskPressure -and -not $cpuPressure -and -not $memoryPressure) {
                if ($script:EffectiveMaxParallel -gt 1) {
                    $script:EffectiveMaxParallel--
                    $changed = $true
                } elseif ($script:EffectiveMaxParallelConversions -gt 1) {
                    $script:EffectiveMaxParallelConversions--
                    $changed = $true
                }
            } elseif ($script:EffectiveMaxParallelConversions -gt 1) {
                $script:EffectiveMaxParallelConversions--
                $changed = $true
            } elseif ($script:EffectiveMaxParallel -gt 1) {
                $script:EffectiveMaxParallel--
                $changed = $true
            }
            $script:AdaptivePressureSamples = 0
            if ($cpuPressure) { $reason = 'CPU pressure' }
            elseif ($memoryPressure) { $reason = 'low available memory' }
            else { $reason = 'disk I/O pressure' }
            if ($changed) {
                $script:AdaptiveStatus = "Reducing new work for $reason; running jobs continue"
                Write-Log -Message ("Adaptive concurrency reduced a slot for {0}. Current capacities: extraction={1}, compression={2}; no fixed job ceiling." -f $reason, $script:EffectiveMaxParallel, $script:EffectiveMaxParallelConversions) -Level INFO
            } else {
                $script:AdaptiveStatus = "Resource pressure ($reason); at minimum concurrency, continuing"
            }
        } else {
            $script:AdaptiveStatus = 'Resource pressure detected; confirming before reducing concurrency'
        }
        return
    }

    $script:AdaptivePressureSamples = 0
    if (-not $healthy) {
        $script:AdaptiveHealthySamples = 0
        $script:AdaptiveStatus = 'Resource use moderate; holding current concurrency'
        return
    }

    $extractionQueued = $script:NextArchiveIndex -lt $script:Archives.Count
    $conversionQueued = $script:PendingConversions -gt 0 -or
        ($script:ProgressStage -eq 'Standalone' -and $script:StandaloneNextIndex -lt $script:StandaloneQueue.Count)
    if (-not $extractionQueued -and -not $conversionQueued) {
        $script:AdaptiveHealthySamples = 0
        $script:AdaptiveStatus = 'Resources healthy; no queued work to scale up'
        return
    }

    $script:AdaptiveHealthySamples++
    if ($script:AdaptiveHealthySamples -lt 2) {
        $script:AdaptiveStatus = 'Resources have headroom; confirming before increasing concurrency'
        return
    }

    $changed = $false
    $stage = $script:AdaptiveNextScaleStage
    for ($attempt = 0; $attempt -lt 2 -and -not $changed; $attempt++) {
        if ($stage -eq 'Extraction' -and $extractionQueued) {
            $script:EffectiveMaxParallel++
            $changed = $true
            $script:AdaptiveNextScaleStage = 'Compression'
        } elseif ($stage -eq 'Compression' -and $conversionQueued) {
            $script:EffectiveMaxParallelConversions++
            $changed = $true
            $script:AdaptiveNextScaleStage = 'Extraction'
        }
        if ($stage -eq 'Extraction') { $stage = 'Compression' } else { $stage = 'Extraction' }
    }
    $script:AdaptiveHealthySamples = 0
    if ($changed) {
        if ($script:AdaptiveNextScaleStage -eq 'Compression') { $increasedStage = 'extraction' } else { $increasedStage = 'compression' }
        $script:AdaptiveStatus = "Resources healthy; increased $increasedStage capacity by one"
        Write-Log -Message "Adaptive concurrency increased capacity. Current capacities: extraction=$script:EffectiveMaxParallel; compression=$script:EffectiveMaxParallelConversions; no fixed job ceiling." -Level INFO
    } else {
        $script:AdaptiveStatus = 'Resources healthy; holding current concurrency'
    }
}

Write-Log -Message "Run started. PowerShell=$($PSVersionTable.PSVersion); Source=$SourceFolder; adaptiveConcurrency=$($script:AdaptiveConcurrencyEnabled); fixedExtractionCeiling=$MaxParallel; fixedCompressionCeiling=$MaxParallelConversions" -Level INFO
Write-Log -Message "Diagnostic log: $script:DiagnosticsLogPath; transcript: $script:TranscriptPath" -Level INFO

function Test-DeleteAllowed {
    if ($script:DryRun -or $script:NoDelete) {
        return $false
    }
    return $true
}

function Test-DeleteCandidate {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [object[]]$RequiredOutputs = @(),
        [switch]$SkipRequiredOutputVerification
    )

    if ([string]::IsNullOrWhiteSpace($SourcePath)) { return $false }
    if (-not (Test-Path -LiteralPath $SourcePath -PathType Leaf)) { return $false }

    $sourceItem = Get-Item -LiteralPath $SourcePath -ErrorAction SilentlyContinue
    if ($null -eq $sourceItem -or $sourceItem.Length -le 0) { return $false }

    if ($SkipRequiredOutputVerification) { return $true }

    foreach ($requiredOutput in $RequiredOutputs) {
        if ($null -eq $requiredOutput -or
            -not $requiredOutput.PSObject.Properties['Path'] -or
            -not $requiredOutput.PSObject.Properties['Hash'] -or
            [string]::IsNullOrWhiteSpace([string]$requiredOutput.Path) -or
            [string]::IsNullOrWhiteSpace([string]$requiredOutput.Hash)) { return $false }
        $outputPath = [string]$requiredOutput.Path
        if (-not (Test-Path -LiteralPath $outputPath -PathType Leaf)) { return $false }
        try {
            $outputItem = Get-Item -LiteralPath $outputPath -ErrorAction Stop
            if ($outputItem.Length -le 0) { return $false }
            $outputHash = Get-FileHashWithDashboard -LiteralPath $outputPath
            if ($outputHash -ne [string]$requiredOutput.Hash) {
                Write-Log -Message "Source deletion blocked because destination content changed after placement: '$outputPath'." -Level WARN
                return $false
            }
        } catch {
            Write-Log -Message "Source deletion blocked because destination verification failed for '$outputPath': $($_.Exception.Message)" -Level WARN
            return $false
        }
    }

    return $true
}

function Remove-TemporaryFolder {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Purpose
    )
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return }
    try {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
    } catch {
        $script:WarningCount++
        Write-Host "Could not remove temporary $Purpose folder '$Path': $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log -Message "Temporary $Purpose folder cleanup failed: path='$Path'; exception=$($_ | Out-String)" -Level WARN
    }
}

function New-TemporaryWorkFolder {
    param([Parameter(Mandatory)][string]$Path)
    New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    $ownerMarkerPath = Join-Path $Path '.extractparts-owner.json'
    try {
        [ordered]@{
            ProcessId = $PID
            ProcessStartTimeUtcTicks = $script:CurrentProcessStartTimeUtcTicks
            RunId = $script:RunId
        } | ConvertTo-Json | Set-Content -LiteralPath $ownerMarkerPath -Encoding UTF8 -ErrorAction Stop
    } catch {
        Remove-TemporaryFolder -Path $Path -Purpose 'untracked work'
        throw "Could not mark temporary folder '$Path' as owned by this run: $($_.Exception.Message)"
    }
}

function Clear-StaleTemporaryFolders {
    param(
        [Parameter(Mandatory)][string[]]$Roots
    )

    try {
        $processes = @(Get-CimInstance -ClassName Win32_Process -ErrorAction Stop)
    } catch {
        $script:WarningCount++
        Write-Host "Could not check whether temporary work folders are in use; leaving them untouched. $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log -Message "Startup temporary-folder cleanup skipped because process ownership could not be checked: $($_ | Out-String)" -Level WARN
        return
    }
    $otherScriptRunning = @($processes | Where-Object {
        $_.ProcessId -ne $PID -and $_.CommandLine -match '(?i)Extract-Parts\.ps1'
    }).Count -gt 0
    $deletedCount = 0
    $skippedCount = 0
    $legacyFolders = @()

    foreach ($root in @($Roots | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique)) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) { continue }
        try {
            $candidateFolders = @(Get-ChildItem -LiteralPath $root -Directory -Force -ErrorAction Stop | Where-Object {
                $_.Name -match '^(ArchiveExtract_[0-9a-f]{32}|\.ExtractParts_ArchiveExtract_[0-9a-f]{32}|\.ExtractParts_NszOutput_[0-9a-f]{32})$'
            })
        } catch {
            $script:WarningCount++
            Write-Host "Could not inspect temporary work-folder candidates in '$root': $($_.Exception.Message)" -ForegroundColor Yellow
            Write-Log -Message "Startup temporary-folder cleanup could not inspect '$root': $($_ | Out-String)" -Level WARN
            continue
        }
        foreach ($folder in $candidateFolders) {
            if (($folder.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) {
                $skippedCount++
                Write-Log -Message "Startup cleanup skipped reparse-point work-folder candidate '$($folder.FullName)'." -Level WARN
                continue
            }

            $markerPath = Join-Path $folder.FullName '.extractparts-owner.json'
            $hasValidMarker = $false
            $ownerIsRunning = $false
            if (Test-Path -LiteralPath $markerPath -PathType Leaf) {
                try {
                    $owner = Get-Content -LiteralPath $markerPath -Raw -ErrorAction Stop | ConvertFrom-Json -ErrorAction Stop
                    if ($null -eq $owner.ProcessId -or $null -eq $owner.ProcessStartTimeUtcTicks -or [string]::IsNullOrWhiteSpace([string]$owner.RunId)) {
                        throw 'Ownership marker is missing required fields.'
                    }
                    $hasValidMarker = $true
                    $ownerProcessRecord = @($processes | Where-Object { $_.ProcessId -eq [int]$owner.ProcessId } | Select-Object -First 1)
                    if ($ownerProcessRecord.Count -gt 0) {
                        $ownerProcess = Get-Process -Id ([int]$owner.ProcessId) -ErrorAction Stop
                        $ownerIsRunning = $ownerProcess.StartTime.ToUniversalTime().Ticks -eq [long]$owner.ProcessStartTimeUtcTicks
                    }
                } catch {
                    $skippedCount++
                    Write-Log -Message "Startup cleanup skipped work folder with an unreadable ownership marker '$markerPath': $($_.Exception.Message)" -Level WARN
                    continue
                }
            }

            $referencedByProcess = @($processes | Where-Object {
                -not [string]::IsNullOrWhiteSpace($_.CommandLine) -and
                $_.CommandLine.IndexOf($folder.FullName, [System.StringComparison]::OrdinalIgnoreCase) -ge 0
            }).Count -gt 0
            if ($ownerIsRunning -or $referencedByProcess) {
                $skippedCount++
                Write-Log -Message "Startup cleanup preserved active work folder '$($folder.FullName)' (ownerActive=$ownerIsRunning; processReference=$referencedByProcess)." -Level INFO
                continue
            }
            if (-not $hasValidMarker) {
                $legacyFolders += $folder
                continue
            }

            try {
                Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop
                $deletedCount++
                Write-Log -Message "Removed stale temporary work folder '$($folder.FullName)'." -Level INFO
            } catch {
                $skippedCount++
                $script:WarningCount++
                Write-Host "Could not remove stale temporary work folder '$($folder.FullName)': $($_.Exception.Message)" -ForegroundColor Yellow
                Write-Log -Message "Startup temporary-folder cleanup failed for '$($folder.FullName)': $($_ | Out-String)" -Level WARN
            }
        }
    }

    if ($legacyFolders.Count -gt 0) {
        $legacyCleanupConfirmed = $false
        if ($otherScriptRunning) {
            Write-Host "Found $($legacyFolders.Count) unmarked temporary folder(s) from older runs. Another Extract-Parts process may be active, so these folders will be preserved." -ForegroundColor Yellow
            Write-Log -Message "Preserved $($legacyFolders.Count) unmarked legacy work folder(s) because another Extract-Parts process appears active." -Level WARN
        } else {
            $legacyCleanupConfirmed = Read-YesNo -Prompt "Found $($legacyFolders.Count) temporary folder(s) from older script versions without ownership markers. Remove them only if no older Extract-Parts run is still working?" -DefaultNo:$true
        }
        if ($legacyCleanupConfirmed) {
            foreach ($folder in $legacyFolders) {
                try {
                    Remove-Item -LiteralPath $folder.FullName -Recurse -Force -ErrorAction Stop
                    $deletedCount++
                    Write-Log -Message "Removed user-confirmed legacy temporary work folder '$($folder.FullName)'." -Level INFO
                } catch {
                    $skippedCount++
                    $script:WarningCount++
                    Write-Host "Could not remove legacy temporary work folder '$($folder.FullName)': $($_.Exception.Message)" -ForegroundColor Yellow
                    Write-Log -Message "Legacy temporary-folder cleanup failed for '$($folder.FullName)': $($_ | Out-String)" -Level WARN
                }
            }
        } else {
            $skippedCount += $legacyFolders.Count
            Write-Log -Message "Preserved $($legacyFolders.Count) unmarked legacy work folder(s); cleanup was not confirmed." -Level WARN
        }
    }

    Write-Log -Message "Startup temporary-folder cleanup completed. Removed=$deletedCount; preservedOrFailed=$skippedCount." -Level INFO
    if ($deletedCount -gt 0) {
        Write-Host "Startup cleanup removed $deletedCount stale temporary work folder(s)." -ForegroundColor Cyan
    }
}

function Remove-BackgroundJob {
    param([System.Management.Automation.Job]$Job)
    if (-not $Job) { return }
    try {
        Remove-Job -Job $Job -Force -ErrorAction Stop | Out-Null
    } catch {
        $script:WarningCount++
        Write-Host "Could not remove background job $($Job.Id): $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Log -Message "Background job cleanup failed: jobId=$($Job.Id); exception=$($_ | Out-String)" -Level WARN
    }
}

function Stop-ActiveWork {
    if ($script:AdaptiveMetricsJob) {
        try {
            if ($script:AdaptiveMetricsJob.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) {
                Stop-Job -Job $script:AdaptiveMetricsJob -ErrorAction Stop
            }
        } catch {
            Write-Log -Message "Could not stop the resource-metrics job during cleanup: $($_.Exception.Message)" -Level WARN
        }
        Remove-BackgroundJob -Job $script:AdaptiveMetricsJob
        $script:AdaptiveMetricsJob = $null
    }
    foreach ($task in @($script:ActiveConversions)) {
        if ($task.Job) {
            try {
                if ($task.Job.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) {
                    Stop-Job -Job $task.Job -ErrorAction Stop
                    Write-Log -Message "Stopped active NSZ job during cleanup: jobId=$($task.Job.Id)." -Level WARN
                }
            } catch {
                Write-Log -Message "Could not stop NSZ job $($task.Job.Id) during cleanup: $($_.Exception.Message)" -Level WARN
            }
            Remove-BackgroundJob -Job $task.Job
        }
        if ($task.WorkFolder -and (Test-Path -LiteralPath $task.WorkFolder)) {
            Remove-TemporaryFolder -Path $task.WorkFolder -Purpose 'conversion'
        }
    }
    foreach ($entry in @($script:Active)) {
        if ($entry.Job) {
            try {
                if ($entry.Job.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) {
                    Stop-Job -Job $entry.Job -ErrorAction Stop
                    Write-Log -Message "Stopped active extraction job during cleanup: jobId=$($entry.Job.Id)." -Level WARN
                }
            } catch {
                Write-Log -Message "Could not stop extraction job $($entry.Job.Id) during cleanup: $($_.Exception.Message)" -Level WARN
            }
            Remove-BackgroundJob -Job $entry.Job
        }
        if ($entry.TempFolder -and (Test-Path -LiteralPath $entry.TempFolder)) {
            Remove-TemporaryFolder -Path $entry.TempFolder -Purpose 'extraction'
        }
    }
    $script:ActiveConversions = @()
    $script:Active = @()
}

function Read-YesNo {
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [bool]$DefaultNo = $true
    )
    if ($script:NonInteractive -and -not $script:Force) {
        Write-Host "Skipping interactive prompt in non-interactive mode: $Prompt" -ForegroundColor Yellow
        Write-Log -Message "Skipped interactive prompt in non-interactive mode: '$Prompt'" -Level WARN
        return $false
    }

    $suffix = if ($DefaultNo) { '[y/N]' } else { '[Y/n]' }
    $answer = (Read-Host "$Prompt $suffix").Trim()
    if ([string]::IsNullOrWhiteSpace($answer)) { return (-not $DefaultNo) }
    return ($answer -match '^(?i:y|yes)$')
}

function Resolve-RarExecutable {
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($env:ProgramFiles)) { $candidates += (Join-Path $env:ProgramFiles 'WinRAR\Rar.exe') }
    $programFilesX86 = [Environment]::GetEnvironmentVariable('ProgramFiles(x86)')
    if (-not [string]::IsNullOrWhiteSpace($programFilesX86)) { $candidates += (Join-Path $programFilesX86 'WinRAR\Rar.exe') }
    if (-not [string]::IsNullOrWhiteSpace($env:LOCALAPPDATA)) { $candidates += (Join-Path $env:LOCALAPPDATA 'Programs\WinRAR\Rar.exe') }
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    $command = Get-Command 'Rar.exe' -ErrorAction SilentlyContinue
    if ($command -and $command.Source) { return $command.Source }
    return $null
}

function Install-WinRAR {
    Write-Host "WinRAR's command-line executable is required to extract the RAR archives." -ForegroundColor Yellow
    Write-Host "The script can try to install WinRAR using the Windows package manager (winget)." -ForegroundColor Yellow
    if (-not (Read-YesNo -Prompt 'Install WinRAR now using winget?')) { Write-Log -Message 'User declined WinRAR installation.' -Level WARN; return $null }

    $winget = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
    if (-not $winget) { $winget = Get-Command 'winget' -ErrorAction SilentlyContinue }
    if (-not $winget -or -not $winget.Source) {
        Write-Host "winget is unavailable. Install Microsoft's App Installer (which provides winget), or install WinRAR manually from https://www.win-rar.com/, then run this script again." -ForegroundColor Red
        return $null
    }

    try {
        Write-Host 'Installing WinRAR with winget...' -ForegroundColor Cyan
        Write-Log -Message 'Starting winget install for RARLab.WinRAR.' -Level INFO
        $wingetArguments = @('install', '--exact', '--id', 'RARLab.WinRAR', '--accept-package-agreements', '--accept-source-agreements')
        & $winget.Source @wingetArguments
        $wingetExitCode = $LASTEXITCODE
        Write-Log -Message "WinRAR winget install exit code: $wingetExitCode" -Level INFO
        if ($wingetExitCode -ne 0) {
            Write-Host "winget returned exit code $wingetExitCode. WinRAR may not have been installed." -ForegroundColor Red
        }
    } catch {
        Write-Host "WinRAR installation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "WinRAR installation exception: $($_ | Out-String)" -Level ERROR
    }

    $rar = Resolve-RarExecutable
    if (-not $rar) {
        Write-Host "WinRAR's Rar.exe was not detected after the installation attempt. Archive extraction will be skipped." -ForegroundColor Red
        Write-Log -Message 'WinRAR executable was not detected after installation attempt.' -Level ERROR
    } else { Write-Log -Message "WinRAR detected after install attempt: '$rar'." -Level INFO }
    return $rar
}

function Test-NszExecutable {
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $false }
    $oldPreference = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $outputLines = @()
    $code = 999
    try {
        $outputLines = @(& $Path --help 2>&1 | ForEach-Object { [string]$_ })
        $code = $LASTEXITCODE
    } catch {
        $outputLines += $_.Exception.Message
    } finally {
        $ErrorActionPreference = $oldPreference
    }
    if ($code -eq 0) {
        Write-Log -Message "NSZ executable self-test succeeded: '$Path'" -Level INFO
        return $true
    }
    $detail = ($outputLines | Select-Object -Last 12) -join ' | '
    Write-Log -Message "NSZ executable self-test FAILED: path='$Path'; exit=$code; output=$detail" -Level ERROR
    return $false
}

function Resolve-NszExecutable {
    $candidates = @(
        (Join-Path $NszPythonVenvFolder 'Scripts\nsz.exe'),
        (Join-Path $NszInstallFolder $NszExecutableName),
        (Join-Path $NszInstallFolder 'nsz.exe')
    )
    foreach ($candidate in $candidates) {
        if ((Test-Path -LiteralPath $candidate -PathType Leaf) -and (Test-NszExecutable -Path $candidate)) {
            return $candidate
        }
    }

    foreach ($name in @('nsz.exe', 'nsz')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source -and (Test-NszExecutable -Path $command.Source)) { return $command.Source }
    }
    return $null
}

function Resolve-PythonExecutable {
    $candidates = @()
    foreach ($name in @('python.exe', 'python3.exe', 'python', 'python3')) {
        $command = Get-Command $name -ErrorAction SilentlyContinue
        if ($command -and $command.Source) { $candidates += $command.Source }
    }
    foreach ($candidate in @(
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python311\python.exe'),
        (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python310\python.exe'),
        (Join-Path $env:ProgramFiles 'Python312\python.exe'),
        (Join-Path $env:ProgramFiles 'Python311\python.exe'),
        (Join-Path $env:ProgramFiles 'Python310\python.exe')
    )) { if ($candidate -and (Test-Path -LiteralPath $candidate -PathType Leaf)) { $candidates += $candidate } }

    $seen = @{}
    foreach ($candidate in $candidates) {
        $key = ([string]$candidate).ToLowerInvariant()
        if ($seen.ContainsKey($key)) { continue }
        $seen[$key] = $true
        try {
            $oldPreference = $ErrorActionPreference
            $ErrorActionPreference = 'Continue'
            $versionOutput = @(& $candidate --version 2>&1 | ForEach-Object { [string]$_ })
            $code = $LASTEXITCODE
            $ErrorActionPreference = $oldPreference
            $versionText = $versionOutput -join ' '
            if ($code -eq 0 -and $versionText -match 'Python\s+(?<major>\d+)\.(?<minor>\d+)') {
                $major = [int]$Matches['major']; $minor = [int]$Matches['minor']
                if ($major -gt 3 -or ($major -eq 3 -and $minor -ge 8)) {
                    Write-Log -Message "Found usable Python: '$candidate' ($versionText)" -Level INFO
                    return $candidate
                }
                Write-Log -Message "Python is too old for this setup: '$candidate' ($versionText); requiring Python 3.8+" -Level WARN
            } else {
                Write-Log -Message "Python candidate failed version test: '$candidate'; exit=$code; output=$versionText" -Level WARN
            }
        } catch {
            $ErrorActionPreference = $oldPreference
            Write-Log -Message "Python candidate test failed: '$candidate'; error=$($_.Exception.Message)" -Level WARN
        }
    }

    # Use the py launcher if installed, without assuming its default version.
    $launcher = Get-Command 'py.exe' -ErrorAction SilentlyContinue
    if (-not $launcher) { $launcher = Get-Command 'py' -ErrorAction SilentlyContinue }
    if ($launcher -and $launcher.Source) {
        try {
            $oldPreference = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
            $resolved = @(& $launcher.Source -3 -c 'import sys; print(sys.executable); print(sys.version_info[0]); print(sys.version_info[1])' 2>&1)
            $code = $LASTEXITCODE
            $ErrorActionPreference = $oldPreference
            if ($code -eq 0 -and $resolved.Count -ge 3) {
                $pyPath = [string]$resolved[0]
                $major = [int]$resolved[1]; $minor = [int]$resolved[2]
                if ((Test-Path -LiteralPath $pyPath -PathType Leaf) -and ($major -gt 3 -or ($major -eq 3 -and $minor -ge 8))) {
                    Write-Log -Message "Resolved Python through py launcher: '$pyPath' ($major.$minor)" -Level INFO
                    return $pyPath
                }
            }
        } catch {
            $ErrorActionPreference = $oldPreference
            Write-Log -Message "Python launcher detection failed: $($_.Exception.Message)" -Level WARN
        }
    }
    return $null
}

function Install-Nsz {
    Write-Host "NSZ is needed to convert .nsp to .nsz and .xci to .xcz." -ForegroundColor Yellow
    Write-Host "The standalone NSZ executable appears to be unreliable on some Windows systems. This fallback installs the official NSZ Python package in an isolated virtual environment." -ForegroundColor Yellow
    Write-Host "Installation folder: $NszPythonVenvFolder" -ForegroundColor Yellow
    if (-not (Read-YesNo -Prompt 'Install NSZ and its required Python packages now?')) {
        Write-Log -Message 'User declined NSZ installation.' -Level WARN
        return $null
    }

    $pythonPath = Resolve-PythonExecutable
    if (-not $pythonPath) {
        Write-Host 'Python 3.8+ was not found. The script can install Python 3.12 using winget first.' -ForegroundColor Yellow
        if (-not (Read-YesNo -Prompt 'Install Python 3.12 using winget?')) {
            Write-Log -Message 'Python installation was declined; NSZ will be unavailable.' -Level WARN
            return $null
        }
        $winget = Get-Command 'winget.exe' -ErrorAction SilentlyContinue
        if (-not $winget) { $winget = Get-Command 'winget' -ErrorAction SilentlyContinue }
        if (-not $winget -or -not $winget.Source) {
            Write-Host 'winget is unavailable. Install Python 3.12 manually from https://www.python.org/downloads/windows/ and rerun this script.' -ForegroundColor Red
            Write-Log -Message 'Python install failed: winget not available.' -Level ERROR
            return $null
        }
        try {
            Write-Host 'Installing Python 3.12 with winget...' -ForegroundColor Cyan
            Write-Log -Message 'Starting winget install for Python.Python.3.12.' -Level INFO
            & $winget.Source install --exact --id Python.Python.3.12 --accept-package-agreements --accept-source-agreements
            $pythonInstallCode = $LASTEXITCODE
            Write-Log -Message "winget Python install exit code: $pythonInstallCode" -Level INFO
        } catch {
            Write-Log -Message "winget Python installation exception: $($_.Exception.Message)" -Level ERROR
        }
        # A newly installed interpreter may not yet be present in this session's PATH.
        $pythonPath = Resolve-PythonExecutable
        if (-not $pythonPath) {
            foreach ($guess in @(
                (Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'),
                (Join-Path $env:ProgramFiles 'Python312\python.exe')
            )) {
                if ($guess -and (Test-Path -LiteralPath $guess -PathType Leaf)) { $pythonPath = $guess; break }
            }
        }
        if (-not $pythonPath) {
            Write-Host 'Python was not detected after installation. Restart PowerShell and rerun the script, or install Python manually.' -ForegroundColor Red
            Write-Log -Message 'Python still not found after winget installation.' -Level ERROR
            return $null
        }
    }

    $oldInstallPreference = $ErrorActionPreference
    try {
        $ErrorActionPreference = 'Continue'
        New-Item -ItemType Directory -Path $NszInstallFolder -Force -ErrorAction Stop | Out-Null
        $venvPython = Join-Path $NszPythonVenvFolder 'Scripts\python.exe'
        if (-not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
            Write-Host 'Creating an isolated Python environment for NSZ...' -ForegroundColor Cyan
            Write-Log -Message "Creating virtual environment with '$pythonPath' at '$NszPythonVenvFolder'." -Level INFO
            & $pythonPath -m venv $NszPythonVenvFolder 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("PythonVenv_{0}.log" -f $script:RunId)) -Append
            if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $venvPython -PathType Leaf)) {
                throw "Python failed to create the virtual environment (exit code $LASTEXITCODE). See the diagnostic log folder: $script:RunLogFolder"
            }
        }
        Write-Host 'Installing/updating the official NSZ package and dependencies (pycryptodome, zstandard, enlighten)...' -ForegroundColor Cyan
        Write-Log -Message "Updating pip in venv '$NszPythonVenvFolder'." -Level INFO
        & $venvPython -m pip install --upgrade pip 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("NszInstall_{0}.log" -f $script:RunId)) -Append
        if ($LASTEXITCODE -ne 0) { throw "pip upgrade failed with exit code $LASTEXITCODE." }

        Write-Log -Message 'Installing/updating official PyPI package: nsz.' -Level INFO
        & $venvPython -m pip install --upgrade nsz 2>&1 | Tee-Object -FilePath (Join-Path $script:RunLogFolder ("NszInstall_{0}.log" -f $script:RunId)) -Append
        if ($LASTEXITCODE -ne 0) { throw "Installing the nsz package failed with exit code $LASTEXITCODE." }

        $venvNsz = [string](Join-Path $NszPythonVenvFolder 'Scripts\nsz.exe')
        if ([string]::IsNullOrWhiteSpace($venvNsz)) { throw 'NSZ installation produced an empty executable path.' }
        if (-not (Test-Path -LiteralPath $venvNsz -PathType Leaf)) { throw "The pip installation completed but nsz.exe was not found at '$venvNsz'." }
        if (-not (Test-NszExecutable -Path $venvNsz)) { throw 'The installed NSZ Python package failed its --help startup test.' }
        Write-Host "NSZ and its dependencies are installed: $venvNsz" -ForegroundColor Green
        Write-Log -Message "NSZ Python package installed successfully at '$venvNsz'." -Level INFO
        return $venvNsz.Trim()
    } catch {
        Write-Host "NSZ installation failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "See installation details under: $script:RunLogFolder" -ForegroundColor Yellow
        Write-Log -Message "NSZ installation failed: $($_ | Out-String)" -Level ERROR
        return $null
    } finally {
        $ErrorActionPreference = $oldInstallPreference
    }
}

function Find-KeysFile {
    $nszPath = $script:NszExe
    if ([string]::IsNullOrWhiteSpace([string]$nszPath)) {
        $nszDirectory = $NszInstallFolder
    } else {
        $nszDirectory = Split-Path -Parent $nszPath
        if ([string]::IsNullOrWhiteSpace([string]$nszDirectory)) {
            $nszDirectory = $NszInstallFolder
        }
    }

    $candidates = @(
        (Join-Path $HOME '.switch\prod.keys'),
        (Join-Path $HOME '.switch\keys.txt'),
        (Join-Path $NszInstallFolder 'prod.keys'),
        (Join-Path $NszInstallFolder 'keys.txt'),
        (Join-Path $nszDirectory 'prod.keys'),
        (Join-Path $nszDirectory 'keys.txt'),
        (Join-Path $PSScriptRoot 'prod.keys'),
        (Join-Path $PSScriptRoot 'keys.txt'),
        (Join-Path (Get-Location).Path 'prod.keys'),
        (Join-Path (Get-Location).Path 'keys.txt')
    )
    foreach ($candidate in $candidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate) -and (Test-Path -LiteralPath $candidate -PathType Leaf)) {
            return $candidate
        }
    }
    return $null
}

function Select-KeysFile {
    if ($script:NonInteractive -and -not $script:Force) {
        Write-Host 'No keys file was detected; skipping the file picker in non-interactive mode.' -ForegroundColor Yellow
        Write-Log -Message 'No keys file detected; file picker skipped in non-interactive mode.' -Level WARN
        return $null
    }
    Write-Host ''
    Write-Host 'NSZ requires a compatible keys file that you must obtain legally. The script will not download keys.' -ForegroundColor Yellow
    Write-Host 'Choose an existing prod.keys or keys.txt file in the file picker, or cancel to skip compression.' -ForegroundColor Yellow
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = 'Select your existing prod.keys or keys.txt file'
    $dialog.Filter = 'Switch keys files (prod.keys or keys.txt)|prod.keys;keys.txt|All files (*.*)|*.*'
    $dialog.CheckFileExists = $true
    $dialog.Multiselect = $false
    $result = $dialog.ShowDialog()
    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        $chosen = $dialog.FileName
        $leaf = [System.IO.Path]::GetFileName($chosen)
        if (($leaf -ieq 'prod.keys' -or $leaf -ieq 'keys.txt') -and (Test-Path -LiteralPath $chosen -PathType Leaf)) {
            $dialog.Dispose()
            return $chosen
        }
        Write-Host 'The selected file must be named prod.keys or keys.txt. Compression will be skipped.' -ForegroundColor Red
    }
    $dialog.Dispose()
    return $null
}

function Select-Folder {
    param(
        [Parameter(Mandatory)][string]$Description,
        [Parameter(Mandatory)][string]$InitialFolder
    )
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.SelectedPath = $InitialFolder
    $dialog.ShowNewFolderButton = -not $script:DryRun
    $result = $dialog.ShowDialog()
    $selected = $null
    if ($result -eq [System.Windows.Forms.DialogResult]::OK -and -not [string]::IsNullOrWhiteSpace($dialog.SelectedPath)) {
        $selected = $dialog.SelectedPath
    }
    $dialog.Dispose()
    return $selected
}

function Get-ArchiveStarts {
    param([Parameter(Mandatory)][string]$Folder)
    @(
        Get-ChildItem -LiteralPath $Folder -File -ErrorAction Stop |
            Where-Object {
                ($_.Name -match '(?i)\.part0*1\.rar$') -or
                (($_.Extension -ieq '.rar') -and ($_.Name -notmatch '(?i)\.part\d+\.rar$'))
            } |
            Sort-Object Name
    )
}

function Get-LooseGameFiles {
    param([string[]]$Folders)
    $seen = @{}
    $files = @()
    foreach ($folder in $Folders) {
        if ([string]::IsNullOrWhiteSpace($folder) -or -not (Test-Path -LiteralPath $folder -PathType Container)) { continue }
        foreach ($file in (Get-ChildItem -LiteralPath $folder -File -ErrorAction Stop | Where-Object { $_.Extension -in @('.nsp', '.xci') })) {
            $key = $file.FullName.ToLowerInvariant()
            if (-not $seen.ContainsKey($key)) {
                $seen[$key] = $true
                $files += $file
            }
        }
    }
    @($files | Sort-Object FullName)
}

function Get-EligibleFiles {
    param([Parameter(Mandatory)][string]$Folder)
    if (-not (Test-Path -LiteralPath $Folder -PathType Container)) { return @() }
    $ownerMarkerPath = Join-Path ([System.IO.Path]::GetFullPath($Folder)) '.extractparts-owner.json'
    @(
        Get-ChildItem -LiteralPath $Folder -File -Recurse -Force -ErrorAction Stop |
            Where-Object {
                $_.FullName -ine $ownerMarkerPath -and
                $_.Extension -ine '.url' -and
                $_.FullName -notmatch '(?i)(^|[\\/])Internet\s*Shortcuts?([\\/]|$)'
            }
    )
}

function Get-ArchiveFilesToDelete {
    param([Parameter(Mandatory)][System.IO.FileInfo]$Archive)

    $allFiles = @(Get-ChildItem -LiteralPath $SourceFolder -File -ErrorAction Stop)
    $partOne = [regex]::Match($Archive.Name, '^(?<base>.+)\.part0*1\.rar$', [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
    if ($partOne.Success) {
        $base = $partOne.Groups['base'].Value
        $pattern = '^' + [regex]::Escape($base) + '\.part\d+\.rar$'
        return @($allFiles | Where-Object { $_.Name -match $pattern } | Sort-Object Name)
    }

    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($Archive.Name)
    $oldVolumePattern = '^' + [regex]::Escape($baseName) + '\.(?:r|s|t|u|v|w|x|y|z)\d{2}$'
    $matching = @($allFiles | Where-Object { $_.Name -ieq $Archive.Name -or $_.Name -match $oldVolumePattern })
    if ($matching.Count -gt 0) { return $matching }
    return @($Archive)
}

function Copy-FlatFile {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [switch]$MoveSource
    )
    $script:LastCopyDestinationPath = $null
    $script:LastCopyDestinationHash = $null
    if ($script:StopProcessing) { return 'Quit' }

    $destination = Join-Path $script:OutputFolder $File.Name
    $sourceLength = $File.Length
    $operation = if ($MoveSource) { 'Move' } else { 'Copy' }
    Set-DashboardWork -Text "$operation output: $($File.Name)" -ClearOperationProgress
    if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
    if ($script:StopProcessing) { return 'Quit' }
    Write-Log -Message "$operation destination handling started: source='$($File.FullName)'; target='$destination'; bytes=$sourceLength." -Level DEBUG
    $exists = Test-Path -LiteralPath $destination
    $action = if ($MoveSource) { 'Move' } else { 'Copy' }
    $incomingHash = $null
    $verifyForDeletion = (Test-DeleteAllowed) -and ($DeleteArchivesOnSuccess -or $script:DeleteStandaloneSources)
    if ($exists) {
        try {
            $existingFile = Get-Item -LiteralPath $destination -ErrorAction Stop
            if ($existingFile.Length -eq $File.Length) {
                $existingHash = Get-FileHashWithDashboard -LiteralPath $destination
                $incomingHash = Get-FileHashWithDashboard -LiteralPath $File.FullName
                if ($existingHash -eq $incomingHash) {
                    $script:LastCopyDestinationPath = $destination
                    $script:LastCopyDestinationHash = $existingHash
                    $script:SkippedCount++
                    Write-Host "Identical file already exists; skipped incoming copy: $($File.Name)" -ForegroundColor DarkYellow
                    Write-Log -Message "Identical destination skipped after SHA256 comparison: '$destination'." -Level INFO
                    return 'Identical'
                }
            }
            $action = 'Overwrite'
            Write-Host "Different file already exists; overwriting: $($File.Name)" -ForegroundColor Yellow
            Write-Log -Message "Destination exists with different content; incoming file will overwrite it: '$destination'." -Level WARN
        } catch {
            $script:FailedCopyCount++
            Write-Log -Message "Could not compare duplicate destination: source='$($File.FullName)'; destination='$destination'; exception=$($_ | Out-String)" -Level ERROR
            Write-Host "Could not compare existing and incoming files for '$($File.Name)': $($_.Exception.Message)" -ForegroundColor Red
            return 'Failed'
        }
    }

    try {
        if ($verifyForDeletion -and $null -eq $incomingHash) {
            $incomingHash = Get-FileHashWithDashboard -LiteralPath $File.FullName
        }
        if ($MoveSource) {
            Move-Item -LiteralPath $File.FullName -Destination $destination -Force -ErrorAction Stop
        } elseif ($action -eq 'Overwrite') {
            Copy-FileWithDashboard -SourcePath $File.FullName -DestinationPath $destination -Overwrite
        } else {
            Copy-FileWithDashboard -SourcePath $File.FullName -DestinationPath $destination
        }
        $copiedFile = Get-Item -LiteralPath $destination -ErrorAction Stop
        if ($copiedFile.Length -ne $sourceLength) {
            throw "$operation file size mismatch: expected $sourceLength bytes, found $($copiedFile.Length) bytes."
        }
        if ($verifyForDeletion) {
            $destinationHash = Get-FileHashWithDashboard -LiteralPath $destination
            if ($destinationHash -ne $incomingHash) {
                throw "$operation file SHA256 mismatch; source data will be retained."
            }
            $script:LastCopyDestinationHash = $destinationHash
        }
        $script:LastCopyDestinationPath = $destination
        $script:CopiedCount++
        Write-Log -Message "$operation operation ($action) succeeded: '$($File.FullName)' -> '$destination' ($sourceLength bytes)." -Level INFO
        return $action
    } catch {
        $script:FailedCopyCount++
        Write-Log -Message "$operation failed: source='$($File.FullName)'; destination='$destination'; exception=$($_ | Out-String)" -Level ERROR
        Write-Host "$operation failed for '$($File.Name)': $($_.Exception.Message)" -ForegroundColor Red
        return 'Failed'
    }
}

# Invoked in a background job so the dashboard can keep refreshing while NSZ works.
$script:NszWorker = {
    param([string]$Executable, [string]$SourceFile, [string]$OutputDirectory, [string]$KeysPath, [bool]$QuickVerify, [int]$CompressionLevel)
    $ErrorActionPreference = 'Continue'
    $lines = @()
    $processOutputQueue = New-Object 'System.Collections.Generic.Queue[string]'
    $lastProgressKey = $null
    $code = 999
    try {
        $verificationArguments = if ($QuickVerify) { @('--quick-verify') } else { @('--verify', '--keep') }
        $arguments = @('--keys', $KeysPath, '--level', [string]$CompressionLevel) + $verificationArguments + @('--minimal-output', '--output', $OutputDirectory, '-C', $SourceFile)
        $lines += "Executable: $Executable"
        $lines += "Source: $SourceFile"
        $lines += "Output directory: $OutputDirectory"
        $lines += "Keys file: $KeysPath"
        $lines += "Arguments: $($arguments -join ' ')"
        $lines += 'NSZ output capture is capped at the last 200 non-progress lines.'
        & $Executable @arguments 2>&1 | ForEach-Object {
            $line = [string]$_
            $progressMatch = [regex]::Match($line.Trim(), '^\s*[\|/\\-]?\s*(?<percent>\d{1,3})%\s*(?<step>.*?)\s*$')
            if ($progressMatch.Success) {
                $percent = [Math]::Min(100, [int]$progressMatch.Groups['percent'].Value)
                $step = [string]$progressMatch.Groups['step'].Value
                if ([string]::IsNullOrWhiteSpace($step)) { $step = 'Working' }
                $progressKey = "$percent|$step"
                if ($progressKey -ne $lastProgressKey) {
                    [pscustomobject]@{ Kind = 'Progress'; Percent = $percent; Step = $step }
                    $lastProgressKey = $progressKey
                }
                return
            }
            if (-not [string]::IsNullOrWhiteSpace($line)) {
                $processOutputQueue.Enqueue($line)
                if ($processOutputQueue.Count -gt 200) { [void]$processOutputQueue.Dequeue() }
            }
        }
        $code = $LASTEXITCODE
        $processOutput = @($processOutputQueue.ToArray())
        $lines += $processOutput
        $lines += "NSZ exit code: $code"
    } catch {
        $processOutput = @($processOutputQueue.ToArray())
        $lines += "Exception invoking NSZ: $($_.Exception.Message)"
        $lines += ($_ | Out-String)
    }
    $detailLines = @($processOutput | Where-Object {
        $_ -match '(?i)(traceback|error|exception|no space left|errno|failed)' -and
        $_ -notmatch '^\s*[\|/\\-]?\s*\d+%\s'
    } | Select-Object -Last 12)
    if ($detailLines.Count -eq 0) {
        $detailLines = @($processOutput | Where-Object { $_ -notmatch '^\s*[\|/\\-]?\s*\d+%\s' } | Select-Object -Last 12)
    }
    if ($detailLines.Count -eq 0) { $detailLines = @($lines | Select-Object -Last 12) }
    [pscustomobject]@{ Kind = 'Result'; ExitCode = [int]$code; Output = ($detailLines -join [Environment]::NewLine).Trim(); FullOutput = ($lines -join [Environment]::NewLine) }
}

# Invoked in a background job to extract one archive set. The worker saves its complete output to a unique log.
$script:ExtractionWorker = {
    param([string]$Executable, [string]$ArchivePath, [string]$TempFolder, [string]$WorkerLogPath)
    $ErrorActionPreference = 'Continue'
    $outputTail = New-Object 'System.Collections.Generic.Queue[string]'
    $outputLineCount = 0
    $lastProgressKey = $null
    $exitCode = 999
    $startedAt = Get-Date
    $process = $null
    function ConvertTo-WindowsArgument {
        param([string]$Value)
        $builder = New-Object System.Text.StringBuilder
        [void]$builder.Append('"')
        $backslashes = 0
        foreach ($character in $Value.ToCharArray()) {
            if ($character -eq '\') {
                $backslashes++
                continue
            }
            if ($character -eq '"') {
                for ($index = 0; $index -lt (2 * $backslashes + 1); $index++) { [void]$builder.Append('\') }
                [void]$builder.Append('"')
            } else {
                for ($index = 0; $index -lt $backslashes; $index++) { [void]$builder.Append('\') }
                [void]$builder.Append($character)
            }
            $backslashes = 0
        }
        for ($index = 0; $index -lt (2 * $backslashes); $index++) { [void]$builder.Append('\') }
        [void]$builder.Append('"')
        return $builder.ToString()
    }
    function Add-WorkerLogLine {
        param([string]$Text)
        $line = '[{0}] {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'), $Text
        try { [System.IO.File]::AppendAllText($WorkerLogPath, $line + [Environment]::NewLine, [System.Text.Encoding]::UTF8) } catch {}
    }
    try {
        Add-WorkerLogLine "Starting extraction worker. PowerShell=$($PSVersionTable.PSVersion)"
        Add-WorkerLogLine "Executable='$Executable'"
        Add-WorkerLogLine "Archive='$ArchivePath'"
        Add-WorkerLogLine "TempFolder='$TempFolder'"
        if (-not (Test-Path -LiteralPath $Executable -PathType Leaf)) { throw "WinRAR executable not found: $Executable" }
        if (-not (Test-Path -LiteralPath $ArchivePath -PathType Leaf)) { throw "Archive file not found: $ArchivePath" }
        if (-not (Test-Path -LiteralPath $TempFolder -PathType Container)) { throw "Temporary extraction folder not found: $TempFolder" }
        $targetPath = $TempFolder.TrimEnd('\') + '\'
        $commandDescription = "& `"$Executable`" x -y `"$ArchivePath`" `"$targetPath`""
        Add-WorkerLogLine "Command=$commandDescription"
        Add-WorkerLogLine 'WinRAR output begins:'
        # Read characters rather than pipeline lines: WinRAR updates progress with
        # carriage returns/backspaces and may not emit a newline until a volume ends.
        $startInfo = New-Object System.Diagnostics.ProcessStartInfo
        $startInfo.FileName = $Executable
        $arguments = @('x', '-y', $ArchivePath, $targetPath)
        $startInfo.Arguments = (($arguments | ForEach-Object { ConvertTo-WindowsArgument -Value $_ }) -join ' ')
        $startInfo.UseShellExecute = $false
        $startInfo.CreateNoWindow = $true
        $startInfo.RedirectStandardOutput = $true
        $startInfo.RedirectStandardError = $true
        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $startInfo
        if (-not $process.Start()) { throw 'Windows could not start the WinRAR process.' }
        $stderrTask = $process.StandardError.ReadToEndAsync()
        $currentOutput = New-Object System.Text.StringBuilder
        while (($character = $process.StandardOutput.Read()) -ge 0) {
            $character = [char]$character
            if ($character -eq "`r" -or $character -eq "`n") {
                $line = $currentOutput.ToString()
                if (-not [string]::IsNullOrWhiteSpace($line)) {
                    Add-WorkerLogLine $line
                    $outputLineCount++
                    $outputTail.Enqueue($line)
                    if ($outputTail.Count -gt 25) { [void]$outputTail.Dequeue() }
                }
                [void]$currentOutput.Clear()
                continue
            }
            [void]$currentOutput.Append($character)
            if ($character -eq '%') {
                $progressMatches = [regex]::Matches($currentOutput.ToString(), '(?<!\d)(?<percent>\d{1,3})%')
                if ($progressMatches.Count -gt 0) {
                    $match = $progressMatches[$progressMatches.Count - 1]
                    $percent = [Math]::Min(100, [int]$match.Groups['percent'].Value)
                    $progressKey = [string]$percent
                    if ($progressKey -ne $lastProgressKey) {
                        [pscustomobject]@{
                            Kind = 'Progress'
                            Percent = $percent
                            Step = 'Extracting archive data'
                        }
                        $lastProgressKey = $progressKey
                    }
                }
            }
        }
        $line = $currentOutput.ToString()
        if (-not [string]::IsNullOrWhiteSpace($line)) {
            Add-WorkerLogLine $line
            $outputLineCount++
            $outputTail.Enqueue($line)
            if ($outputTail.Count -gt 25) { [void]$outputTail.Dequeue() }
        }
        $process.WaitForExit()
        $stderrText = $stderrTask.GetAwaiter().GetResult()
        if (-not [string]::IsNullOrWhiteSpace($stderrText)) {
            foreach ($line in ($stderrText -split "`r?`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                Add-WorkerLogLine "STDERR: $line"
                $outputLineCount++
                $outputTail.Enqueue("STDERR: $line")
                if ($outputTail.Count -gt 25) { [void]$outputTail.Dequeue() }
            }
        }
        $exitCode = [int]$process.ExitCode
        $process.Dispose()
        Add-WorkerLogLine "WinRAR exit code=$exitCode"
        Add-WorkerLogLine ("Elapsed seconds={0:N2}" -f ((Get-Date) - $startedAt).TotalSeconds)
    } catch {
        $exitCode = 998
        Add-WorkerLogLine "EXCEPTION invoking extraction: $($_.Exception.Message)"
        Add-WorkerLogLine ($_ | Out-String)
        if ($_.ScriptStackTrace) { Add-WorkerLogLine "StackTrace: $($_.ScriptStackTrace)" }
        if ($process) { $process.Dispose() }
    }
    $messageLines = @($outputTail.ToArray() | Select-Object -Last 25)
    [pscustomobject]@{
        Kind = 'Result'
        ExitCode = [int]$exitCode
        Message = ($messageLines -join ' | ')
        LogPath = $WorkerLogPath
        Started = $startedAt.ToString('s')
        Finished = (Get-Date).ToString('s')
        OutputLineCount = $outputLineCount
    }
}

function Start-NextExtraction {
    Update-DiskReserveStatus -ForceRefresh
    if ($script:StopProcessing -or $script:DiskReserveBlocked -or $script:DiskReserveUnavailable) { return }
    if ($script:NextArchiveIndex -ge $script:Archives.Count) { return }
    $archive = $script:Archives[$script:NextArchiveIndex]
    $queueNumber = $script:NextArchiveIndex + 1
    $script:NextArchiveIndex++
    $tempFolder = Join-Path $script:OutputFolder (".ExtractParts_ArchiveExtract_{0}" -f ([guid]::NewGuid().ToString('N')))
    $safeName = [regex]::Replace([System.IO.Path]::GetFileNameWithoutExtension($archive.Name), '[^A-Za-z0-9._-]', '_')
    $workerLogPath = Join-Path $script:RunLogFolder ("RAR_{0:D3}_{1}_{2}.log" -f $queueNumber, $safeName, ([guid]::NewGuid().ToString('N').Substring(0,8)))
    Set-DashboardWork -Text "Starting extraction of $($archive.Name)" -ClearOperationProgress
    try {
        New-TemporaryWorkFolder -Path $tempFolder
        Set-Content -LiteralPath $workerLogPath -Value "[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff')] RAR worker log for '$($archive.Name)'" -Encoding UTF8
        $job = Start-Job -ScriptBlock $script:ExtractionWorker -ArgumentList $script:RarExe, $archive.FullName, $tempFolder, $workerLogPath -ErrorAction Stop
        $script:Active += [pscustomobject]@{
            Archive = $archive
            TempFolder = $tempFolder
            Job = $job
            LogPath = $workerLogPath
            QueueNumber = $queueNumber
            StartedAt = Get-Date
            ProgressPercent = $null
            ProgressStep = 'Starting extraction'
            ProgressBaselineAt = $null
            ProgressBaselinePercent = 0
            EstimatedSecondsRemaining = $null
            LastProgressAt = Get-Date
            WorkerResult = $null
        }
        Write-Host "Started extraction [$queueNumber/$($script:TotalArchives)]: $($archive.Name) (job $($job.Id))" -ForegroundColor Cyan
        $capacityDescription = if ($script:AdaptiveConcurrencyEnabled) {
            "adaptiveCapacity=$($script:EffectiveMaxParallel); no fixed job ceiling"
        } else {
            "fixedLimit=$MaxParallel"
        }
        Write-Log -Message "Started parallel extraction queue #$queueNumber/$($script:TotalArchives): archive='$($archive.FullName)'; jobId=$($job.Id); destinationVolumeStaging='$tempFolder'; workerLog='$workerLogPath'. Active=$($script:Active.Count); $capacityDescription" -Level INFO
    } catch {
        Remove-TemporaryFolder -Path $tempFolder -Purpose 'extraction'
        $script:ErrorCount++
        $script:CompletedArchives++
        Write-Host "Could not start extraction for $($archive.Name): $($_.Exception.Message)" -ForegroundColor Red
        Write-Log -Message "Could not start extraction queue #$queueNumber for '$($archive.FullName)': $($_ | Out-String)" -Level ERROR
    }
}

function Remove-ExtractionTemp {
    param([string]$Folder, [System.Management.Automation.Job]$Job)
    Remove-BackgroundJob -Job $Job
    if ($Folder) { Remove-TemporaryFolder -Path $Folder -Purpose 'extraction' }
}

function Get-PendingArchiveItems {
    param([int]$Limit = $QueuePreviewLimit)
    $items = @()
    $end = [Math]::Min($script:Archives.Count, $script:NextArchiveIndex + $Limit)
    for ($i = $script:NextArchiveIndex; $i -lt $end; $i++) { $items += $script:Archives[$i] }
    return $items
}

function Get-PendingStandaloneItems {
    param([int]$Limit = $QueuePreviewLimit)
    $items = @()
    $end = [Math]::Min($script:StandaloneQueue.Count, $script:StandaloneNextIndex + $Limit)
    for ($i = $script:StandaloneNextIndex; $i -lt $end; $i++) { $items += $script:StandaloneQueue[$i] }
    return $items
}

function Get-ElapsedLabel {
    param([Parameter(Mandatory)][datetime]$StartedAt)
    $elapsed = (Get-Date) - $StartedAt
    return '{0:00}:{1:00}:{2:00}' -f [Math]::Floor($elapsed.TotalHours), $elapsed.Minutes, $elapsed.Seconds
}

function Set-DashboardWork {
    param(
        [Parameter(Mandatory)][string]$Text,
        [switch]$ClearOperationProgress
    )
    $script:DashboardState.CurrentWork = $Text
    if ($ClearOperationProgress) {
        $script:DashboardState.OperationPercent = $null
        $script:DashboardState.OperationLabel = $null
    }
}

function Set-DashboardOperationProgress {
    param(
        [Parameter(Mandatory)][int]$Percent,
        [Parameter(Mandatory)][string]$Label
    )
    $script:DashboardState.OperationPercent = [Math]::Min(100, [Math]::Max(0, $Percent))
    $script:DashboardState.OperationLabel = $Label
}

function Get-FileHashWithDashboard {
    param([Parameter(Mandatory)][string]$LiteralPath)
    $file = Get-Item -LiteralPath $LiteralPath -ErrorAction Stop
    $stream = $null
    $hasher = $null
    $buffer = New-Object byte[] (4MB)
    $bytesReadTotal = [long]0
    $lastPercent = -1
    try {
        $stream = [System.IO.File]::OpenRead($file.FullName)
        $hasher = [System.Security.Cryptography.SHA256]::Create()
        while (($bytesRead = $stream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            [void]$hasher.TransformBlock($buffer, 0, $bytesRead, $buffer, 0)
            $bytesReadTotal += $bytesRead
            $percent = if ($file.Length -gt 0) { [int][Math]::Floor(100 * $bytesReadTotal / $file.Length) } else { 100 }
            if ($percent -ne $lastPercent) {
                Set-DashboardWork -Text "Verifying SHA-256: $($file.Name) ($percent%)"
                Set-DashboardOperationProgress -Percent $percent -Label "SHA-256 $($file.Name)"
                $lastPercent = $percent
            }
            Update-ProgressDashboard -FileOperation
        }
        [void]$hasher.TransformFinalBlock((New-Object byte[] 0), 0, 0)
        return ([BitConverter]::ToString($hasher.Hash)).Replace('-', '')
    } finally {
        if ($hasher) { $hasher.Dispose() }
        if ($stream) { $stream.Dispose() }
        Set-DashboardWork -Text $script:DashboardState.CurrentWork -ClearOperationProgress
    }
}

function Copy-FileWithDashboard {
    param(
        [Parameter(Mandatory)][string]$SourcePath,
        [Parameter(Mandatory)][string]$DestinationPath,
        [switch]$Overwrite
    )
    $source = Get-Item -LiteralPath $SourcePath -ErrorAction Stop
    $inputStream = $null
    $outputStream = $null
    $buffer = New-Object byte[] (4MB)
    $bytesCopied = [long]0
    $lastPercent = -1
    try {
        $inputStream = [System.IO.File]::OpenRead($source.FullName)
        $mode = if ($Overwrite) { [System.IO.FileMode]::Create } else { [System.IO.FileMode]::CreateNew }
        $outputStream = [System.IO.FileStream]::new($DestinationPath, $mode, [System.IO.FileAccess]::Write, [System.IO.FileShare]::None, $buffer.Length, [System.IO.FileOptions]::SequentialScan)
        while (($bytesRead = $inputStream.Read($buffer, 0, $buffer.Length)) -gt 0) {
            $outputStream.Write($buffer, 0, $bytesRead)
            $bytesCopied += $bytesRead
            $percent = if ($source.Length -gt 0) { [int][Math]::Floor(100 * $bytesCopied / $source.Length) } else { 100 }
            if ($percent -ne $lastPercent) {
                Set-DashboardWork -Text "Copying $($source.Name): $percent%"
                Set-DashboardOperationProgress -Percent $percent -Label "Copy $($source.Name)"
                $lastPercent = $percent
            }
            Update-ProgressDashboard -FileOperation
        }
        $outputStream.Flush()
    } finally {
        if ($outputStream) { $outputStream.Dispose() }
        if ($inputStream) { $inputStream.Dispose() }
        Set-DashboardWork -Text $script:DashboardState.CurrentWork -ClearOperationProgress
    }
}

function Get-RemainingLabel {
    param([Parameter(Mandatory)][double]$Seconds)
    $remaining = [TimeSpan]::FromSeconds([Math]::Max(0, $Seconds))
    return '{0:00}:{1:00}:{2:00}' -f [Math]::Floor($remaining.TotalHours), $remaining.Minutes, $remaining.Seconds
}

function Set-TaskProgress {
    param(
        [Parameter(Mandatory)]$Task,
        [Parameter(Mandatory)][int]$Percent,
        [Parameter(Mandatory)][string]$Step
    )
    $now = Get-Date
    $stepChanged = $Task.ProgressStep -ne $Step
    if ($null -eq $Task.ProgressBaselineAt -or $Percent -lt $Task.ProgressBaselinePercent -or $stepChanged) {
        $Task.ProgressBaselineAt = $now
        $Task.ProgressBaselinePercent = $Percent
        $Task.EstimatedSecondsRemaining = $null
    } elseif ($Percent -gt $Task.ProgressBaselinePercent) {
        $baselineSeconds = ($now - $Task.ProgressBaselineAt).TotalSeconds
        $progressedPercent = $Percent - $Task.ProgressBaselinePercent
        if ($baselineSeconds -ge 3 -and $progressedPercent -gt 0 -and $Percent -lt 100) {
            $percentPerSecond = $progressedPercent / $baselineSeconds
            $Task.EstimatedSecondsRemaining = (100 - $Percent) / $percentPerSecond
        } elseif ($Percent -eq 100) {
            $Task.EstimatedSecondsRemaining = 0
        }
        if ($baselineSeconds -ge 3) {
            $Task.ProgressBaselineAt = $now
            $Task.ProgressBaselinePercent = $Percent
        }
    }
    $Task.ProgressPercent = $Percent
    $Task.ProgressStep = $Step
    $Task.LastProgressAt = $now
}

function Receive-ExtractionJobUpdates {
    param([object[]]$Entries = $script:Active)
    foreach ($entry in $Entries) {
        $updates = @(Receive-Job -Job $entry.Job -ErrorAction SilentlyContinue)
        foreach ($update in $updates) {
            if ($null -eq $update -or -not $update.PSObject.Properties['Kind']) { continue }
            if ($update.Kind -eq 'Progress') {
                Set-TaskProgress -Task $entry -Percent ([int]$update.Percent) -Step ([string]$update.Step)
            } elseif ($update.Kind -eq 'Result') {
                $entry.WorkerResult = $update
            }
        }
    }
}

function Get-ProgressJobSummary {
    $summary = [ordered]@{
        Running = 0
        Finished = 0
        Failed = 0
        Stopped = 0
        Names = @()
    }

    foreach ($entry in $script:Active) {
        $state = [string]$entry.Job.State
        switch ($state) {
            'Running' { $summary.Running++ }
            'Completed' { $summary.Finished++ }
            'Failed' { $summary.Failed++ }
            'Stopped' { $summary.Stopped++ }
            default { $summary.Running++ }
        }
        $displayState = switch ($state) {
            'Running' { 'Extracting' }
            'Completed' { 'Extracted; awaiting handling' }
            'Failed' { 'Extraction failed' }
            'Stopped' { 'Stopped' }
            default { 'Starting' }
        }
        if ($state -eq 'Running' -and $null -ne $entry.ProgressPercent) {
            $displayState += " $($entry.ProgressPercent)%"
            if ($null -ne $entry.EstimatedSecondsRemaining) {
                $displayState += " | item ETA ~$(Get-RemainingLabel -Seconds $entry.EstimatedSecondsRemaining)"
            }
        }
        if ($state -eq 'Running' -and $entry.LastProgressAt) {
            $noProgress = (Get-Date) - $entry.LastProgressAt
            if ($noProgress.TotalMinutes -ge 5) {
                $displayState += " | no tool progress report for $(Get-ElapsedLabel -StartedAt $entry.LastProgressAt)"
            }
        }
        if ($entry.StartedAt) { $displayState += " | $(Get-ElapsedLabel -StartedAt $entry.StartedAt)" }
        $summary.Names += "$displayState | $($entry.Archive.Name)"
    }

    return $summary
}

function Get-DashboardRow {
    param(
        [Parameter(Mandatory)][string]$Text,
        [Parameter(Mandatory)][int]$Width
    )
    $Width = [Math]::Max(6, $Width)
    $innerWidth = $Width - 4
    if ($Text.Length -gt $innerWidth) {
        if ($innerWidth -gt 3) {
            $Text = $Text.Substring(0, $innerWidth - 3) + '...'
        } else {
            $Text = $Text.Substring(0, $innerWidth)
        }
    }
    return '| ' + $Text.PadRight($innerWidth) + ' |'
}

function Write-ProgressDashboard {
    param(
        [Parameter(Mandatory)][object[]]$Rows,
        [Parameter(Mandatory)][int]$Width
    )
    $border = '+' + ('-' * ($Width - 2)) + '+'
    $rendered = @($border)
    foreach ($row in $Rows) { $rendered += $row.Text }
    $rendered += $border

    if ($script:ProgressDashboardInteractive) {
        try {
            $rawUI = $script:ProgressDashboardRawUI
            $cursorBefore = $rawUI.CursorPosition
            $windowTop = $rawUI.WindowPosition.Y
            $visibleBottom = [Math]::Min($rawUI.BufferSize.Height - 1, $windowTop + $rawUI.WindowSize.Height - 1)
            $requiredLines = [Math]::Max($script:ProgressDashboardLineCount, $rendered.Count)
            $maxTop = $visibleBottom - $requiredLines - 1
            if ($maxTop -lt $windowTop) {
                throw 'The terminal window is too short to render the live dashboard without scrolling.'
            }
            if ($script:ProgressDashboardTop -lt $windowTop -or $script:ProgressDashboardTop -gt $maxTop) {
                $script:ProgressDashboardTop = [Math]::Max($windowTop, [Math]::Min($cursorBefore.Y, $maxTop))
            }

            $clearLine = ' ' * [Math]::Max(1, [Math]::Min($Width, $rawUI.WindowSize.Width) - 1)
            for ($lineIndex = 0; $lineIndex -lt $script:ProgressDashboardLineCount; $lineIndex++) {
                $rawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $script:ProgressDashboardTop + $lineIndex)
                Write-Host $clearLine -NoNewline
            }

            $rawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $script:ProgressDashboardTop)
            Write-Host $rendered[0] -ForegroundColor DarkCyan
            for ($lineIndex = 0; $lineIndex -lt $Rows.Count; $lineIndex++) {
                Write-Host $Rows[$lineIndex].Text -ForegroundColor $Rows[$lineIndex].Color
            }
            Write-Host $rendered[$rendered.Count - 1] -ForegroundColor DarkCyan

            $script:ProgressDashboardLineCount = $rendered.Count
            $restoreY = [Math]::Max($cursorBefore.Y, $script:ProgressDashboardTop + $script:ProgressDashboardLineCount)
            if ($restoreY -gt $visibleBottom) { $restoreY = $visibleBottom }
            $rawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $restoreY)
        } catch {
            $script:ProgressDashboardInteractive = $false
            $script:ProgressDashboardCursorFailed = $true
            $script:ProgressDashboardTop = $null
            $script:ProgressDashboardLineCount = 0
            Write-Log -Message "Live dashboard cursor rendering is unavailable; using text output instead: $($_.Exception.Message)" -Level WARN
            for ($lineIndex = 0; $lineIndex -lt $rendered.Count; $lineIndex++) {
                $color = if ($lineIndex -eq 0 -or $lineIndex -eq ($rendered.Count - 1)) { 'DarkCyan' } else { $Rows[$lineIndex - 1].Color }
                Write-Host $rendered[$lineIndex] -ForegroundColor $color
            }
        }
    } else {
        for ($lineIndex = 0; $lineIndex -lt $rendered.Count; $lineIndex++) {
            if ($lineIndex -eq 0 -or $lineIndex -eq ($rendered.Count - 1)) {
                Write-Host $rendered[$lineIndex] -ForegroundColor DarkCyan
            } else {
                Write-Host $rendered[$lineIndex] -ForegroundColor $Rows[$lineIndex - 1].Color
            }
        }
    }
    try { [Console]::Out.Flush() } catch {}
}

function Update-ProgressDashboard {
    param([switch]$FileOperation)
    $now = Get-Date
    if ($script:DashboardState.LastRenderedAt -and
        ($now - $script:DashboardState.LastRenderedAt).TotalSeconds -lt 1) {
        return
    }
    $pendingItems = @()
    $pendingCount = 0
    $rows = @()
    $windowWidth = 96
    $interactive = $false
    $rawUI = $null
    try {
        $rawUI = $Host.UI.RawUI
        $windowWidth = [Math]::Max(6, [Math]::Min(96, $rawUI.WindowSize.Width - 2))
        $interactive = -not [Console]::IsOutputRedirected -and
            $rawUI.WindowSize.Width -gt 0 -and $rawUI.BufferSize.Height -gt 0
    } catch {
        $interactive = $false
    }
    if ($interactive -and -not $script:ProgressDashboardInteractive -and -not $script:ProgressDashboardCursorFailed) {
        $script:ProgressDashboardTop = $rawUI.CursorPosition.Y
        $script:ProgressDashboardRawUI = $rawUI
        $script:ProgressDashboardInteractive = $true
    }

    if ($script:ProgressStage -eq 'Archives') {
        $pendingItems = @(Get-PendingArchiveItems)
        $pendingCount = [Math]::Max(0, $script:Archives.Count - $script:NextArchiveIndex)
        $jobSummary = Get-ProgressJobSummary
        $percent = if ($script:TotalArchives -gt 0) { [int](100 * $script:CompletedArchives / $script:TotalArchives) } else { 100 }
        $phase = 'ARCHIVES'
        $progressCount = "$($script:CompletedArchives)/$($script:TotalArchives) finished (success or failure)"
        $jobStatus = "Extraction: $($jobSummary.Running) running, $($jobSummary.Finished) ready for handling"
        if ($jobSummary.Failed -gt 0 -or $jobSummary.Stopped -gt 0) {
            $jobStatus += ", $($jobSummary.Failed) failed, $($jobSummary.Stopped) stopped"
        }
    } else {
        $pendingItems = @(Get-PendingStandaloneItems)
        $pendingCount = [Math]::Max(0, $script:StandaloneQueue.Count - $script:StandaloneNextIndex)
        $percent = if ($script:TotalStandalone -gt 0) { [int](100 * $script:StandaloneProcessedCount / $script:TotalStandalone) } else { 100 }
        $phase = 'LOOSE FILES'
        $progressCount = "$($script:StandaloneProcessedCount)/$($script:TotalStandalone) handled"
        $jobStatus = "Files waiting in queue: $pendingCount"
    }

    $elapsedText = Get-ElapsedLabel -StartedAt $script:RunStartedAt
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text (" Archive Extract Convert  |  {0}  |  Elapsed {1}" -f $phase, $elapsedText) -Width $windowWidth; Color = 'Cyan' }
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text (" Destination: {0}" -f $script:OutputFolder) -Width $windowWidth; Color = 'DarkGray' }
    $barWidth = 20
    $filled = [int][Math]::Floor($barWidth * $percent / 100)
    $bar = ('#' * $filled) + ('.' * ($barWidth - $filled))
    $progressDetails = $progressCount
    if ($script:ActiveConversions.Count -gt 0) {
        $activeProgress = @($script:ActiveConversions | Where-Object { $null -ne $_.ProgressPercent } | Select-Object -First 1)
        if ($activeProgress.Count -gt 0) {
            $progressDetails += " | NSZ $($activeProgress[0].ProgressPercent)% $($activeProgress[0].ProgressStep)"
        } else {
            $progressDetails += " | $($script:ActiveConversions.Count) NSZ conversion(s) in progress"
        }
    }
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text (" Handled  [{0}]  {1,3}%  |  {2}" -f $bar, $percent, $progressDetails) -Width $windowWidth; Color = 'Green' }
    $currentStep = if ([string]::IsNullOrWhiteSpace($script:DashboardState.CurrentWork)) { 'Waiting for the next item' } else { $script:DashboardState.CurrentWork }
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text (" Current   {0}" -f $currentStep) -Width $windowWidth; Color = 'Yellow' }
    $itemProgressPercent = $script:DashboardState.OperationPercent
    $itemProgressLabel = $script:DashboardState.OperationLabel
    if ($null -eq $itemProgressPercent -and $script:ProgressStage -eq 'Archives') {
        $activeItem = @($script:Active | Where-Object { $null -ne $_.ProgressPercent } | Select-Object -First 1)
        if ($activeItem.Count -gt 0) {
            $itemProgressPercent = [int]$activeItem[0].ProgressPercent
            $itemProgressLabel = "RAR $($activeItem[0].Archive.Name)"
        }
    }
    if ($null -eq $itemProgressPercent -and $script:ActiveConversions.Count -gt 0) {
        $activeItem = @($script:ActiveConversions | Where-Object { $null -ne $_.ProgressPercent } | Select-Object -First 1)
        if ($activeItem.Count -gt 0) {
            $itemProgressPercent = [int]$activeItem[0].ProgressPercent
            $itemProgressLabel = "NSZ $($activeItem[0].File.Name)"
        }
    }
    if ($null -ne $itemProgressPercent) {
        $itemFilled = [int][Math]::Floor($barWidth * $itemProgressPercent / 100)
        $itemBar = ('#' * $itemFilled) + ('.' * ($barWidth - $itemFilled))
        $itemProgressText = " Item     [$itemBar] $($itemProgressPercent.ToString().PadLeft(3))% | $itemProgressLabel"
    } else {
        $itemProgressText = " Item     [$('.' * $barWidth)]  -- | Waiting for tool progress"
    }
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text $itemProgressText -Width $windowWidth; Color = 'Cyan' }
    if ($script:StopProcessing) {
        $stopStatus = if ($script:StopReason -eq 'user request') { 'Stopping: active jobs will finish; original sources will be kept' } else { "Stopping: $($script:StopReason); original sources will be kept" }
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text $stopStatus -Width $windowWidth; Color = 'Yellow' }
    } elseif ($script:DiskReserveUnavailable) {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text 'Disk reserve unavailable; waiting until free space can be checked' -Width $windowWidth; Color = 'Yellow' }
    } elseif ($script:DiskReserveBlocked) {
        $spaceDetails = @($script:DiskReserveVolumes | Where-Object { $_.BelowReserve } | ForEach-Object { "$($_.DeviceID): $([Math]::Round($_.FreeBytes / 1GB, 1)) GB free (reserve $([Math]::Round($_.ReserveBytes / 1GB, 1)) GB)" }) -join '; '
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text "Low disk reserve; waiting for free space: $spaceDetails" -Width $windowWidth; Color = 'Yellow' }
    }
    if (-not $script:StopProcessing -and -not $script:NonInteractive -and -not $script:CancellationKeyUnavailable) {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text 'Press Q to stop starting new work; active jobs will finish and source files will be kept.' -Width $windowWidth; Color = 'DarkGray' }
    } elseif (-not $script:StopProcessing -and -not $script:NonInteractive -and $script:CancellationKeyUnavailable) {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text 'Q-to-stop is unavailable in this host; Ctrl+C interrupts and cleanup is attempted.' -Width $windowWidth; Color = 'DarkGray' }
    }
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text $jobStatus -Width $windowWidth; Color = 'White' }
    if ($script:AdaptiveConcurrencyEnabled) {
        if ($script:AdaptiveMetrics) {
            $resourceSummary = 'Auto/no cap | CPU {0:N0}% | RAM {1:N0}% free | disk {2:N0}% | free {3:N1}% | RAR {4} NSZ {5}' -f `
                $script:AdaptiveMetrics.CpuPercent, $script:AdaptiveMetrics.MemoryFreePercent, $script:AdaptiveMetrics.DiskBusyPercent, `
                $script:AdaptiveMetrics.MinimumDiskFreePercent, $script:EffectiveMaxParallel, $script:EffectiveMaxParallelConversions
        } else {
            $resourceSummary = "Auto/no cap | metrics unavailable | capacity RAR $($script:EffectiveMaxParallel), NSZ $($script:EffectiveMaxParallelConversions)"
        }
        $statusColor = if ($script:AdaptiveMetricsUnavailable -or $script:AdaptiveLowSpace -or $script:AdaptiveStatus -match 'pressure|moderate|unavailable|Low disk') { 'Yellow' } else { 'Cyan' }
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text $resourceSummary -Width $windowWidth; Color = $statusColor }
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("Resource status: {0}" -f $script:AdaptiveStatus) -Width $windowWidth; Color = $statusColor }
    } else {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("Fixed limits | slots RAR {0}/{1}, NSZ {2}/{3}" -f $script:EffectiveMaxParallel, $MaxParallel, $script:EffectiveMaxParallelConversions, $MaxParallelConversions) -Width $windowWidth; Color = 'Cyan' }
    }

    if ($script:ProgressStage -eq 'Archives') {
        foreach ($name in $jobSummary.Names) {
            $color = if ($name -match '\[Extraction failed\]') { 'Red' } elseif ($name -match '\[Extracted; awaiting handling\]') { 'Green' } else { 'Gray' }
            $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("   RAR  {0}" -f $name) -Width $windowWidth; Color = $color }
        }
    }
    if ($script:ActiveConversions.Count -gt 0) {
        $conversionCapacity = if ($script:AdaptiveConcurrencyEnabled) { "adaptive capacity $($script:EffectiveMaxParallelConversions)" } else { "fixed limit $($script:EffectiveMaxParallelConversions)" }
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("NSZ conversion jobs: {0} active; {1}" -f $script:ActiveConversions.Count, $conversionCapacity) -Width $windowWidth; Color = 'White' }
        foreach ($conversion in $script:ActiveConversions) {
            $conversionState = if ($conversion.Job.State -eq 'Completed') { 'Result ready' } elseif ($conversion.Job.State -eq 'Running') { 'Converting' } else { [string]$conversion.Job.State }
            if ($null -ne $conversion.ProgressPercent) {
                $conversionState = "$($conversion.ProgressPercent)% $($conversion.ProgressStep)"
                if ($null -ne $conversion.EstimatedSecondsRemaining) {
                    $conversionState += " | ETA ~$(Get-RemainingLabel -Seconds $conversion.EstimatedSecondsRemaining)"
                }
            }
            if ($conversion.LastProgressAt) {
                $noProgress = (Get-Date) - $conversion.LastProgressAt
                if ($noProgress.TotalMinutes -ge 5) {
                    $conversionState += " | no tool progress report for $(Get-ElapsedLabel -StartedAt $conversion.LastProgressAt)"
                }
            }
            $conversionState += " | $($conversion.File.Name)"
            $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("   NSZ  {0}" -f $conversionState) -Width $windowWidth; Color = 'Magenta' }
        }
    }

    if ($script:ProgressStage -eq 'Archives' -and $pendingCount -gt 0) {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text 'Waiting in queue:' -Width $windowWidth; Color = 'DarkCyan' }
    } elseif ($script:ProgressStage -eq 'Standalone' -and $pendingCount -gt 0) {
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text 'Next files:' -Width $windowWidth; Color = 'DarkCyan' }
    }
    $visibleQueueCount = [Math]::Min($QueuePreviewLimit, $pendingItems.Count)
    for ($slot = 0; $slot -lt $visibleQueueCount; $slot++) {
        if ($script:ProgressStage -eq 'Archives') {
            $queueOrdinal = $script:NextArchiveIndex + $slot + 1
            $queueLabel = '{0,2}. ' -f $queueOrdinal
        } else {
            $queueOrdinal = $script:StandaloneNextIndex + $slot + 1
            $queueLabel = '{0,2}. ' -f $queueOrdinal
        }

        $details = $pendingItems[$slot].Name
        if ($pendingItems[$slot].PSObject.Properties['Length']) {
            $details = "$details - {0:N1} MB" -f ([double]$pendingItems[$slot].Length / 1MB)
        }

        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("   {0}{1}" -f $queueLabel, $details) -Width $windowWidth; Color = 'DarkGray' }
    }
    if ($pendingCount -gt $QueuePreviewLimit) {
        $overflowCount = $pendingCount - $QueuePreviewLimit
        $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text ("   ... and {0} more queued" -f $overflowCount) -Width $windowWidth; Color = 'DarkGray' }
    }

    $resultColor = if ($script:ErrorCount -gt 0 -or $script:ConversionErrorCount -gt 0 -or $script:FailedCopyCount -gt 0) { 'Red' } else { 'Green' }
    $summary = "Results: {0} placed | {1} identical skipped | {2} conversions passed | {3} errors | {4} warnings" -f $script:CopiedCount, $script:SkippedCount, $script:ConversionSuccessCount, ($script:ErrorCount + $script:ConversionErrorCount + $script:FailedCopyCount), $script:WarningCount
    $rows += [pscustomobject]@{ Text = Get-DashboardRow -Text $summary -Width $windowWidth; Color = $resultColor }

    $script:DashboardState.Rows = $rows
    Write-ProgressDashboard -Rows $script:DashboardState.Rows -Width $windowWidth
    $script:DashboardState.LastRenderedAt = Get-Date

    # Render the current state before polling jobs or resource providers so a slow
    # host/provider cannot prevent the user from seeing that work has started.
    if (-not $FileOperation) {
        Update-AdaptiveConcurrency
        Update-DiskReserveStatus
        if ($script:ProgressStage -eq 'Archives' -and $script:Active.Count -gt 0) {
            Receive-ExtractionJobUpdates
        }
    }
    if (-not $script:CancellationKeyUnavailable -and -not $script:NonInteractive) {
        try {
            for ($keyRead = 0; $keyRead -lt 16 -and [Console]::KeyAvailable; $keyRead++) {
                $key = [Console]::ReadKey($true)
                if ($key.KeyChar -in @('q', 'Q') -and -not $script:StopProcessing) {
                    $script:StopProcessing = $true
                    $script:StopReason = 'user request'
                    Write-Host 'Stop requested. No new work will start; active jobs will finish and source files will be kept.' -ForegroundColor Yellow
                    Write-Log -Message 'User requested a graceful stop with Q; no new jobs will start and source files will be preserved.' -Level WARN
                }
            }
        } catch {
            $script:CancellationKeyUnavailable = $true
            Write-Log -Message "Interactive Q-to-stop input is unavailable in this host: $($_.Exception.Message)" -Level WARN
        }
    }
}

function Clear-ProgressDashboard {
    if ($script:ProgressDashboardInteractive -and $script:ProgressDashboardLineCount -gt 0) {
        try {
            $rawUI = $script:ProgressDashboardRawUI
            $cursorBefore = $rawUI.CursorPosition
            $rawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $script:ProgressDashboardTop)
            $clearLine = ' ' * [Math]::Max(1, [Math]::Min($rawUI.WindowSize.Width, 96) - 1)
            for ($lineIndex = 0; $lineIndex -lt $script:ProgressDashboardLineCount; $lineIndex++) {
                Write-Host $clearLine -NoNewline
                $rawUI.CursorPosition = [System.Management.Automation.Host.Coordinates]::new(0, $script:ProgressDashboardTop + $lineIndex + 1)
            }
            $rawUI.CursorPosition = $cursorBefore
        } catch {
            $script:ProgressDashboardCursorFailed = $true
            Write-Log -Message "Could not clear the live dashboard: $($_.Exception.Message)" -Level WARN
        }
    }
    $script:DashboardState.LastRenderedAt = $null
    $script:ProgressDashboardTop = $null
    $script:ProgressDashboardLineCount = 0
    $script:ProgressDashboardRawUI = $null
    $script:ProgressDashboardInteractive = $false
}

function Start-NSZConversion {
    param([Parameter(Mandatory)][System.IO.FileInfo]$File)

    Update-DiskReserveStatus -ForceRefresh
    if ($script:StopProcessing -or $script:DiskReserveBlocked -or $script:DiskReserveUnavailable) { return $null }
    $workFolder = Join-Path $script:OutputFolder (".ExtractParts_NszOutput_{0}" -f ([guid]::NewGuid().ToString('N')))
    $expectedExtension = if ($File.Extension -ieq '.nsp') { '.nsz' } else { '.xcz' }
    Set-DashboardWork -Text "Starting $($expectedExtension.ToUpperInvariant()) conversion for $($File.Name)" -ClearOperationProgress
    if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
    if ($script:StopProcessing -or $script:DiskReserveBlocked -or $script:DiskReserveUnavailable) { return $null }
    try {
        New-TemporaryWorkFolder -Path $workFolder
        $verificationMode = if ($script:QuickVerify) { 'quick (NCA hashes only)' } else { 'full (NSP SHA256 and NCA hashes)' }
        Write-Log -Message "Starting NSZ conversion: executable='$script:NszExe'; source='$($File.FullName)'; targetExtension='$expectedExtension'; keys='$script:KeysFile'; verification='$verificationMode'; compressionLevel=$CompressionLevel; workFolder='$workFolder'." -Level INFO
        $job = Start-Job -ScriptBlock $script:NszWorker -ArgumentList $script:NszExe, $File.FullName, $workFolder, $script:KeysFile, $script:QuickVerify, $CompressionLevel -ErrorAction Stop
        $task = [pscustomobject]@{
            File = $File
            Job = $job
            WorkFolder = $workFolder
            ProgressPercent = $null
            ProgressStep = 'Starting conversion'
            StartedAt = Get-Date
            ProgressBaselineAt = $null
            ProgressBaselinePercent = 0
            EstimatedSecondsRemaining = $null
            LastProgressAt = Get-Date
        }
        $script:ActiveConversions += $task
        Set-DashboardWork -Text "Converting $($File.Name)" -ClearOperationProgress
        if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
        return $task
    } catch {
        $script:ConversionErrorCount++
        Write-Log -Message "Could not start NSZ conversion for '$($File.FullName)': $($_ | Out-String)" -Level ERROR
        Write-Host "Could not start compression for '$($File.Name)': $($_.Exception.Message)" -ForegroundColor Red
        Remove-TemporaryFolder -Path $workFolder -Purpose 'conversion'
        return $null
    }
}

function Start-PendingConversions {
    Update-DiskReserveStatus
    if ($script:StopProcessing -or $script:DiskReserveBlocked -or $script:DiskReserveUnavailable) { return }
    while ($script:PendingConversionNextIndex -lt $script:PendingConversionRecords.Count -and
        $script:ActiveConversions.Count -lt $script:EffectiveMaxParallelConversions) {
        $record = $script:PendingConversionRecords[$script:PendingConversionNextIndex]
        $record.Task = Start-NSZConversion -File $record.File
        if (-not $record.Task -and ($script:StopProcessing -or $script:DiskReserveBlocked -or $script:DiskReserveUnavailable)) { return }
        $script:PendingConversionNextIndex++
        $script:PendingConversions = $script:PendingConversionRecords.Count - $script:PendingConversionNextIndex
    }
}

function Complete-NSZConversion {
    param([Parameter(Mandatory)]$Task)

    $file = $Task.File
    if (-not $Task.Job) {
        return [pscustomobject]@{ Success = $false; OutputFile = $null; WorkFolder = $null; Message = 'The NSZ background job could not be started.' }
    }
    $expectedExtension = if ($file.Extension -ieq '.nsp') { '.nsz' } else { '.xcz' }
    $job = $Task.Job
    Set-DashboardWork -Text "Waiting for NSZ conversion to finish: $($file.Name)" -ClearOperationProgress
    if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
    try {
        $result = $null
        while ($job.State -in @('NotStarted', 'Running', 'Blocked', 'Disconnected')) {
            $updates = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
            foreach ($update in $updates) {
                if ($null -eq $update -or -not $update.PSObject.Properties['Kind']) { continue }
                if ($update.Kind -eq 'Progress') {
                    Set-TaskProgress -Task $Task -Percent ([int]$update.Percent) -Step ([string]$update.Step)
                } elseif ($update.Kind -eq 'Result') {
                    $result = $update
                }
            }
            Update-ProgressDashboard
            Start-PendingConversions
            # Handle completed extractions while a conversion waits, so their files can use any idle NSZ slot.
            if ($script:ProgressStage -eq 'Archives' -and
                -not $script:StopProcessing -and
                -not $script:ProcessingFinishedArchiveEntries -and
                $script:ActiveConversions.Count -lt $script:EffectiveMaxParallelConversions) {
                $finishedEntries = @($script:Active | Where-Object { $_.Job.State -in @('Completed', 'Failed', 'Stopped') })
                if ($finishedEntries.Count -gt 0) {
                    $finishedIds = @($finishedEntries | ForEach-Object { $_.Job.Id })
                    $script:Active = @($script:Active | Where-Object { $finishedIds -notcontains $_.Job.Id })
                    $savedPendingRecords = $script:PendingConversionRecords
                    $savedPendingNextIndex = $script:PendingConversionNextIndex
                    $savedPendingCount = $script:PendingConversions
                    $script:ProcessingFinishedArchiveEntries = $true
                    try {
                        if (-not $script:StopProcessing) {
                            while (-not $script:DiskReserveBlocked -and -not $script:DiskReserveUnavailable -and
                                $script:Active.Count -lt $script:EffectiveMaxParallel -and $script:NextArchiveIndex -lt $script:Archives.Count) {
                                Start-NextExtraction
                            }
                        }
                        Process-FinishedArchives -Entries $finishedEntries
                    } finally {
                        $script:PendingConversionRecords = $savedPendingRecords
                        $script:PendingConversionNextIndex = $savedPendingNextIndex
                        $script:PendingConversions = $savedPendingCount
                        $script:ProcessingFinishedArchiveEntries = $false
                        Set-DashboardWork -Text "Waiting for NSZ conversion to finish: $($file.Name)"
                    }
                }
            }
            Start-Sleep -Milliseconds 400
        }
        $updates = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
        foreach ($update in $updates) {
            if ($null -eq $update -or -not $update.PSObject.Properties['Kind']) { continue }
            if ($update.Kind -eq 'Progress') {
                Set-TaskProgress -Task $Task -Percent ([int]$update.Percent) -Step ([string]$update.Step)
            } elseif ($update.Kind -eq 'Result') {
                $result = $update
            }
        }
        if ($null -eq $result) {
            $jobDiagnostics = @("NSZ background job state: $($job.State)")
            $nszChildJobs = @($job.ChildJobs)
            if ($nszChildJobs.Count -gt 0) {
                foreach ($jobError in @($nszChildJobs[0].Error)) { $jobDiagnostics += ('Job error: ' + [string]$jobError) }
                if ($nszChildJobs[0].JobStateInfo.Reason) { $jobDiagnostics += ('Job reason: ' + $nszChildJobs[0].JobStateInfo.Reason.ToString()) }
            } else {
                $jobDiagnostics += 'The NSZ background job has no child-job object.'
            }
            throw "The NSZ background job returned no result. $($jobDiagnostics -join ' | ')"
        }
        Write-Log -Message "NSZ raw output for '$($file.FullName)': $([string]$result.FullOutput)" -Level DEBUG
        if ([int]$result.ExitCode -ne 0) {
            $detail = [string]$result.Output
            if ([string]::IsNullOrWhiteSpace($detail)) { $detail = 'NSZ did not provide diagnostic output.' }
            throw "NSZ exited with code $($result.ExitCode). $detail"
        }

        $outputs = @(Get-ChildItem -LiteralPath $Task.WorkFolder -File -ErrorAction Stop | Where-Object { $_.Extension -ieq $expectedExtension })
        if ($outputs.Count -ne 1) {
            throw "NSZ exited successfully but produced $($outputs.Count) '$expectedExtension' output files; expected exactly one."
        }
        $outputFile = Get-Item -LiteralPath $outputs[0].FullName -ErrorAction Stop
        if ($outputFile.Length -le 0) { throw 'The compressed output file is empty.' }
        $script:ConversionSuccessCount++
        Write-Log -Message "NSZ conversion passed: source='$($file.FullName)'; output='$($outputFile.FullName)'; bytes=$($outputFile.Length)." -Level INFO
        return [pscustomobject]@{ Success = $true; OutputFile = $outputFile; WorkFolder = $Task.WorkFolder; Message = $null }
    } catch {
        $script:ConversionErrorCount++
        Write-Log -Message "NSZ conversion failed for '$($file.FullName)': $($_ | Out-String)" -Level ERROR
        Write-Host "Conversion failed for '$($file.Name)': $($_.Exception.Message)" -ForegroundColor Red
        if ($_.Exception.Message -match '(?i)(no space left|not enough space|disk full)') {
            Write-Host "Free space on the output drive and the system temporary drive before retrying. Source files and the RAR archive will be kept." -ForegroundColor Yellow
        }
        if ($script:ProgressStage -eq 'Archives') {
            Write-Host 'The source RAR archive will be kept because conversion did not complete successfully.' -ForegroundColor Yellow
        }
        if ($Task.WorkFolder -and (Test-Path -LiteralPath $Task.WorkFolder)) {
            Remove-TemporaryFolder -Path $Task.WorkFolder -Purpose 'conversion'
        }
        return [pscustomobject]@{ Success = $false; OutputFile = $null; WorkFolder = $null; Message = $_.Exception.Message }
    } finally {
        Remove-BackgroundJob -Job $job
        $script:ActiveConversions = @($script:ActiveConversions | Where-Object { $_.Job.Id -ne $job.Id })
    }
}

function Process-ConvertedOutput {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo]$File,
        [Parameter(Mandatory)]$Conversion,
        [bool]$IsStandalone = $false
    )
    if (-not $Conversion.Success) { return 'Failed' }

    try {
        Set-DashboardWork -Text "Placing compressed output: $($Conversion.OutputFile.Name)" -ClearOperationProgress
        if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
        $copyResult = Copy-FlatFile -File $Conversion.OutputFile -MoveSource
        if ($copyResult -in @('Copy', 'Move', 'Overwrite', 'Identical')) {
            if ($IsStandalone -and $script:DeleteStandaloneSources -and (Test-DeleteAllowed) -and -not $script:StopProcessing) {
                $destinationPath = $script:LastCopyDestinationPath
                $requiredOutput = [pscustomobject]@{ Path = $destinationPath; Hash = $script:LastCopyDestinationHash }
                if (-not (Test-DeleteCandidate -SourcePath $File.FullName -RequiredOutputs @($requiredOutput))) {
                    Write-Host "Compressed output missing or invalid; original standalone file kept: $($File.Name)" -ForegroundColor Yellow
                    Write-Log -Message "Delete skipped for standalone '$($File.FullName)' because verification failed for output '$destinationPath'." -Level WARN
                } else {
                    try {
                        Update-ProgressDashboard
                        if ($script:StopProcessing) {
                            Write-Host "Compressed output saved; original kept because a stop was requested: $($File.Name)" -ForegroundColor Yellow
                        } else {
                            Remove-Item -LiteralPath $File.FullName -Force -ErrorAction Stop
                            Write-Host "Compressed and deleted original: $($File.Name)" -ForegroundColor Green
                        }
                    } catch {
                        Write-Host "Compressed output saved, but original could not be deleted ('$($File.Name)'): $($_.Exception.Message)" -ForegroundColor Yellow
                    }
                }
            } elseif ($IsStandalone -and $script:DeleteStandaloneSources) {
                Write-Host "Original standalone file kept: $($File.Name) (stop or safety mode is active)." -ForegroundColor Yellow
            }
        }
        return $copyResult
    } finally {
        if ($Conversion.WorkFolder -and (Test-Path -LiteralPath $Conversion.WorkFolder)) {
            Remove-TemporaryFolder -Path $Conversion.WorkFolder -Purpose 'conversion'
        }
    }
}

function Process-FileBatch {
    param(
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [bool]$IsStandalone = $false,
        [switch]$MoveUncompressedSources
    )
    $conversionRecords = @()
    foreach ($file in $Files) {
        if ($script:CompressionEnabled -and $file.Extension -in @('.nsp', '.xci')) {
            $conversionRecords += [pscustomobject]@{ File = $file; Task = $null; QueueIndex = $conversionRecords.Count }
        }
    }

    $script:PendingConversionRecords = $conversionRecords
    $script:PendingConversionNextIndex = 0
    $script:PendingConversions = $conversionRecords.Count
    Start-PendingConversions

    $outcomes = @()
    foreach ($file in $Files) {
        $script:LastCopyDestinationPath = $null
        if ($script:StopProcessing) {
            $outcomes += [pscustomobject]@{ File = $file; Result = 'Quit'; DestinationPath = $null }
            continue
        }

        $conversionRecord = @($conversionRecords | Where-Object { $_.File.FullName -eq $file.FullName } | Select-Object -First 1)
        if ($conversionRecord.Count -gt 0) {
            $record = $conversionRecord[0]
            while (-not $script:StopProcessing -and -not $record.Task -and $record.QueueIndex -ge $script:PendingConversionNextIndex) {
                Set-DashboardWork -Text "Waiting for a conversion slot: $($file.Name)" -ClearOperationProgress
                Update-ProgressDashboard
                Start-PendingConversions
                if (-not $record.Task -and $record.QueueIndex -ge $script:PendingConversionNextIndex) {
                    Start-Sleep -Milliseconds 400
                }
            }
            if ($script:StopProcessing) {
                $outcomes += [pscustomobject]@{ File = $file; Result = 'Quit'; DestinationPath = $null; DestinationHash = $null }
                continue
            }
            if ($record.Task) {
                $conversion = Complete-NSZConversion -Task $record.Task
            } else {
                $conversion = [pscustomobject]@{ Success = $false; OutputFile = $null; WorkFolder = $null; Message = 'The NSZ background job could not be started.' }
            }
            Start-PendingConversions
            $result = Process-ConvertedOutput -File $file -Conversion $conversion -IsStandalone:$IsStandalone
        } else {
            Set-DashboardWork -Text "Copying extracted file: $($file.Name)" -ClearOperationProgress
            if ($script:ProgressStage -in @('Archives', 'Standalone')) { Update-ProgressDashboard }
            if ($MoveUncompressedSources) {
                $result = Copy-FlatFile -File $file -MoveSource
            } else {
                $result = Copy-FlatFile -File $file
            }
        }
        $outcomes += [pscustomobject]@{
            File = $file
            Result = $result
            DestinationPath = $script:LastCopyDestinationPath
            DestinationHash = $script:LastCopyDestinationHash
        }
    }

    foreach ($task in @($script:ActiveConversions)) {
        $conversion = Complete-NSZConversion -Task $task
        if ($conversion.WorkFolder -and (Test-Path -LiteralPath $conversion.WorkFolder)) {
            Remove-TemporaryFolder -Path $conversion.WorkFolder -Purpose 'conversion'
        }
    }
    $script:PendingConversions = 0
    $script:PendingConversionRecords = @()
    $script:PendingConversionNextIndex = 0
    return $outcomes
}

function Complete-ArchiveFileHandling {
    param(
        [Parameter(Mandatory)]$Entry,
        [Parameter(Mandatory)][System.IO.FileInfo[]]$Files,
        [Parameter(Mandatory)]$Outcomes
    )

    $allFilesHandled = $true
    $requiredOutputPaths = @()
    foreach ($file in $Files) {
        $matchedOutcomes = @($Outcomes | Where-Object { $_.File.FullName -eq $file.FullName } | Select-Object -First 1)
        if ($matchedOutcomes.Count -eq 0) {
            $allFilesHandled = $false
            $script:ErrorCount++
            Write-Log -Message "No processing outcome was returned for extracted file '$($file.FullName)'." -Level ERROR
            continue
        }
        $outcome = $matchedOutcomes[0]
        Write-Log -Message "Handled extracted file '$($file.Name)': result='$($outcome.Result)'." -Level INFO
        if ($outcome.Result -notin @('Copy', 'Move', 'Overwrite', 'Identical') -or [string]::IsNullOrWhiteSpace($outcome.DestinationPath)) {
            $allFilesHandled = $false
        } else {
            $requiredOutputPaths += [pscustomobject]@{ Path = $outcome.DestinationPath; Hash = $outcome.DestinationHash }
        }
    }

    if ($DeleteArchivesOnSuccess -and $allFilesHandled -and -not $script:StopProcessing -and (Test-DeleteAllowed)) {
        Set-DashboardWork -Text "Verifying outputs and cleaning up archive: $($Entry.Archive.Name)" -ClearOperationProgress
        if ($script:ProgressStage -eq 'Archives') { Update-ProgressDashboard }
        $volumeFiles = @(Get-ArchiveFilesToDelete -Archive $Entry.Archive)
        $deleteFailure = $false

        Write-Log -Message "Archive content handling succeeded for '$($Entry.Archive.Name)'; attempting to delete source volumes: $($volumeFiles.Name -join ', ')." -Level INFO
        $outputsVerified = $false
        if ($volumeFiles.Count -gt 0) {
            Set-DashboardWork -Text "Verifying destination before deleting archive set: $($Entry.Archive.Name)" -ClearOperationProgress
            if ($script:ProgressStage -eq 'Archives') { Update-ProgressDashboard }
            $outputsVerified = Test-DeleteCandidate -SourcePath $volumeFiles[0].FullName -RequiredOutputs $requiredOutputPaths
            if (-not $outputsVerified) {
                $deleteFailure = $true
                Write-Host "Source archive kept because destination verification failed: $($Entry.Archive.Name)" -ForegroundColor Yellow
                Write-Log -Message "Archive set deletion blocked because destination verification failed for '$($Entry.Archive.Name)'." -Level WARN
            }
        } else {
            $deleteFailure = $true
            Write-Host "Source archive kept because no archive volumes were found: $($Entry.Archive.Name)" -ForegroundColor Yellow
            Write-Log -Message "Archive set deletion skipped because no source volumes were found for '$($Entry.Archive.Name)'." -Level WARN
        }
        foreach ($volume in $volumeFiles) {
            Update-ProgressDashboard
            if ($script:StopProcessing) {
                $deleteFailure = $true
                Write-Host "Source archive kept because a stop was requested: $($volume.Name)" -ForegroundColor Yellow
                Write-Log -Message "Archive volume deletion cancelled by user request: '$($volume.FullName)'." -Level WARN
                continue
            }
            if (-not $outputsVerified) { continue }
            if (-not (Test-DeleteCandidate -SourcePath $volume.FullName -SkipRequiredOutputVerification)) {
                $deleteFailure = $true
                Write-Host "Delete skipped for archive volume '$($volume.Name)' because required output verification failed." -ForegroundColor Yellow
                Write-Log -Message "Archive volume delete skipped because verification failed for '$($volume.FullName)'." -Level WARN
                continue
            }
            try {
                Remove-Item -LiteralPath $volume.FullName -Force -ErrorAction Stop
                $script:DeletedArchiveCount++
                Write-Log -Message "Deleted source volume '$($volume.FullName)'." -Level INFO
            } catch {
                $deleteFailure = $true
                Write-Host "Could not delete source archive file '$($volume.Name)': $($_.Exception.Message)" -ForegroundColor Red
                Write-Log -Message "Source archive deletion failed: '$($volume.FullName)': $($_ | Out-String)" -Level ERROR
            }
        }
        if (-not $deleteFailure) {
            Write-Host "Extraction, conversion/copy succeeded; deleted source archive set: $($Entry.Archive.Name)" -ForegroundColor Green
        } else {
            $script:WarningCount++
            Write-Host "Output succeeded, but one or more source volume files could not be deleted for: $($Entry.Archive.Name)" -ForegroundColor Yellow
        }
    } else {
        Set-DashboardWork -Text "Keeping archive after output handling: $($Entry.Archive.Name)" -ClearOperationProgress
        if ($script:ProgressStage -eq 'Archives') { Update-ProgressDashboard }
        if (-not (Test-DeleteAllowed)) {
            Write-Host "Keeping source archive: $($Entry.Archive.Name) (dry-run/no-delete mode is enabled)." -ForegroundColor Yellow
            Write-Log -Message "Source archive kept because dry-run/no-delete mode is enabled: '$($Entry.Archive.Name)'" -Level WARN
        } else {
            Write-Host "Keeping source archive: $($Entry.Archive.Name) (one or more files failed, were skipped, or processing was stopped)." -ForegroundColor Yellow
            Write-Log -Message "Source archive kept because not all eligible files were copied/converted without skips or errors: '$($Entry.Archive.Name)'" -Level WARN
        }
    }

    Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
    $script:CompletedArchives++
}

function Process-FinishedArchives {
    param([Parameter(Mandatory)]$Entries)

    Receive-ExtractionJobUpdates -Entries $Entries
    $contexts = @()
    foreach ($Entry in $Entries) {
        Set-DashboardWork -Text "Checking extraction result: $($Entry.Archive.Name)" -ClearOperationProgress
        if ($script:ProgressStage -eq 'Archives') { Update-ProgressDashboard }
        $exitCode = 999
        $detail = ''
        $jobState = [string]$Entry.Job.State
        $jobResult = $null
        $childJob = $null
        $childJobs = @($Entry.Job.ChildJobs)
        if ($childJobs.Count -gt 0) { $childJob = $childJobs[0] }
        $diagnosticLines = @()
        try {
            if ($Entry.WorkerResult) {
                $jobResult = $Entry.WorkerResult
            } elseif ($jobState -eq 'Completed') {
                $received = @(Receive-Job -Job $Entry.Job -ErrorAction Continue)
                $jobResult = @($received | Where-Object { $null -ne $_ -and $_.PSObject.Properties['Kind'] -and $_.Kind -eq 'Result' }) | Select-Object -Last 1
                if ($jobResult) {
                    $Entry.WorkerResult = $jobResult
                } else {
                    $diagnosticLines += 'Job reached Completed state but returned no object containing ExitCode.'
                    foreach ($item in $received) { $diagnosticLines += ('Received output type={0}; value={1}' -f $item.GetType().FullName, [string]$item) }
                }
            } else {
                $diagnosticLines += "Background job state was '$jobState', not Completed."
                if ($Entry.Job.JobStateInfo -and $Entry.Job.JobStateInfo.Reason) { $diagnosticLines += ('Job reason: ' + $Entry.Job.JobStateInfo.Reason.ToString()) }
            }
        } catch {
            $diagnosticLines += "Receive-Job exception: $($_.Exception.Message)"
            $diagnosticLines += ($_ | Out-String)
        }

        if ($jobResult) {
            $exitCode = [int]$jobResult.ExitCode
            $detail = [string]$jobResult.Message
            if ($jobResult.LogPath) { $Entry.LogPath = [string]$jobResult.LogPath }
        }

        if ($childJob) {
            try {
                if ($childJob.JobStateInfo -and $childJob.JobStateInfo.Reason) { $diagnosticLines += ('Child job reason: ' + $childJob.JobStateInfo.Reason.ToString()) }
                foreach ($jobError in @($childJob.Error)) {
                    $diagnosticLines += ('Job error stream: ' + [string]$jobError)
                    if ($jobError.ScriptStackTrace) { $diagnosticLines += ('Job error stack: ' + $jobError.ScriptStackTrace) }
                }
                foreach ($jobWarning in @($childJob.Warning)) { $diagnosticLines += ('Job warning stream: ' + [string]$jobWarning) }
                foreach ($jobInfo in @($childJob.Information)) { $diagnosticLines += ('Job information stream: ' + [string]$jobInfo.MessageData) }
            } catch { $diagnosticLines += "Could not inspect child job streams: $($_.Exception.Message)" }
        }
        if ($Entry.LogPath -and (Test-Path -LiteralPath $Entry.LogPath -PathType Leaf)) {
            $diagnosticLines += "Full per-archive WinRAR log: $($Entry.LogPath)"
        }
        if ($diagnosticLines.Count -gt 0) {
            foreach ($line in $diagnosticLines) { Write-Log -Message "Archive='$($Entry.Archive.Name)' diagnostic: $line" -Level ERROR }
        }
        Write-Log -Message "Finished extraction job. Archive='$($Entry.Archive.FullName)'; jobId=$($Entry.Job.Id); jobState=$jobState; childState=$(if($childJob){$childJob.State}else{'none'}); exitCode=$exitCode; workerLog='$($Entry.LogPath)'; resultReturned=$([bool]$jobResult)" -Level INFO

        $archiveSucceeded = ($jobState -eq 'Completed' -and $null -ne $jobResult -and $exitCode -eq 0)
        if (-not $archiveSucceeded) {
            if ($jobResult -and $exitCode -eq 1) {
                $script:WarningCount++
                Write-Host "Warning: $($Entry.Archive.Name) (WinRAR exit code 1). Source archive kept." -ForegroundColor Yellow
            } elseif ($jobResult -and $exitCode -notin @(0,1,998,999)) {
                $script:ErrorCount++
                Write-Host "Extraction failed: $($Entry.Archive.Name) (WinRAR exit code $exitCode). Source archive kept." -ForegroundColor Red
            } elseif ($jobResult -and $exitCode -eq 998) {
                $script:ErrorCount++
                Write-Host "Extraction could not start correctly for: $($Entry.Archive.Name). Source archive kept." -ForegroundColor Red
            } else {
                $script:ErrorCount++
                Write-Host "Extraction worker failed before a reliable WinRAR exit code was returned: $($Entry.Archive.Name) (job state '$jobState'). Source archive kept." -ForegroundColor Red
            }
            if (-not [string]::IsNullOrWhiteSpace($detail)) { Write-Host "  Worker detail: $detail" -ForegroundColor DarkRed }
            foreach ($line in $diagnosticLines) { Write-Host "  Diagnostic: $line" -ForegroundColor DarkYellow }
            if ($Entry.LogPath) { Write-Host "  Worker log: $($Entry.LogPath)" -ForegroundColor Yellow }
            if ($jobResult -and $exitCode -notin @(0, 998, 999)) { Write-Log -Message "WinRAR reported nonzero exit code $exitCode for '$($Entry.Archive.Name)'; detail='$detail'" -Level ERROR }
            Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
            $script:CompletedArchives++
            continue
        }

        if ($script:StopProcessing) {
            Write-Host "Stopped before processing extracted contents: $($Entry.Archive.Name). Source archive kept." -ForegroundColor Yellow
            Write-Log -Message "Completed archive extraction but user stopped before consuming outputs: '$($Entry.Archive.Name)'." -Level WARN
            Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
            $script:CompletedArchives++
            continue
        }

        Set-DashboardWork -Text "Scanning extracted files: $($Entry.Archive.Name)" -ClearOperationProgress
        Update-ProgressDashboard
        $eligibleFiles = @(Get-EligibleFiles -Folder $Entry.TempFolder)
        Write-Log -Message "Archive extraction contents scanned: '$($Entry.Archive.Name)'; eligible files=$($eligibleFiles.Count); temp='$($Entry.TempFolder)'." -Level INFO
        if ($eligibleFiles.Count -eq 0) {
            Write-Host "No eligible files found in $($Entry.Archive.Name); source archive kept." -ForegroundColor Yellow
            $script:WarningCount++
            Remove-ExtractionTemp -Folder $Entry.TempFolder -Job $Entry.Job
            $script:CompletedArchives++
            continue
        }
        foreach ($file in $eligibleFiles) {
            Write-Log -Message "Handling extracted file '$($file.FullName)' ($($file.Length) bytes). CompressionEnabled=$($script:CompressionEnabled)." -Level INFO
        }
        $contexts += [pscustomobject]@{ Entry = $Entry; Files = $eligibleFiles }
    }

    if ($contexts.Count -eq 0) { return }
    $batchFiles = @()
    foreach ($context in $contexts) { $batchFiles += $context.Files }
    $batchOutcomes = @(Process-FileBatch -Files $batchFiles -MoveUncompressedSources)
    foreach ($context in $contexts) {
        Complete-ArchiveFileHandling -Entry $context.Entry -Files $context.Files -Outcomes $batchOutcomes
    }
}

try {
# ------------------------ Input and destination ------------------------
Write-Host 'Step: Select source and destination folders.' -ForegroundColor Cyan
try { Add-Type -AssemblyName System.Windows.Forms } catch {
    Write-Host 'Unable to load Windows Forms for the folder/file pickers.' -ForegroundColor Red
    Write-Log -Message 'Unable to load System.Windows.Forms. Exiting.' -Level ERROR
    if (-not $script:NonInteractive) { Read-Host 'Press Enter to exit' }
    Stop-RunLogging
    exit 1
}

if ([string]::IsNullOrWhiteSpace($SourceFolder) -or -not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
    if ($script:NonInteractive -and -not $script:Force) {
        Write-Host 'A valid -SourceFolder is required in non-interactive mode.' -ForegroundColor Red
        Write-Log -Message 'Non-interactive run requires a valid source folder; exiting.' -Level ERROR
        Stop-RunLogging
        exit 1
    }
    $sourcePromptInitial = if ([string]::IsNullOrWhiteSpace($SourceFolder)) { $env:USERPROFILE } else { $SourceFolder }
    $sourceChoice = Select-Folder -Description 'Choose the folder that contains your RAR archive sets or loose .nsp/.xci files.' -InitialFolder $sourcePromptInitial
    if ([string]::IsNullOrWhiteSpace($sourceChoice)) {
        Write-Host 'Cancelled. No files were processed.' -ForegroundColor Yellow
        Write-Log -Message 'User cancelled source-folder selection; no files processed.' -Level WARN
        Stop-RunLogging
        exit 0
    }
    $SourceFolder = $sourceChoice
    $DefaultOutputFolder = Join-Path $SourceFolder 'Extracted'
}

if (-not (Test-Path -LiteralPath $SourceFolder -PathType Container)) {
    Write-Host "Source folder not found: $SourceFolder" -ForegroundColor Red
    Write-Log -Message "Source folder not found: '$SourceFolder'. Exiting." -Level ERROR
    if (-not $script:NonInteractive) { Read-Host 'Press Enter to exit' }
    Stop-RunLogging
    exit 1
}

Write-Host 'Step: Scanning source folder for RAR archive sets.' -ForegroundColor Cyan
$detectedArchives = @(Get-ArchiveStarts -Folder $SourceFolder)
Write-Log -Message "Detected $($detectedArchives.Count) archive set(s) directly in source folder '$SourceFolder'." -Level INFO
if (-not $script:DryRun -and [string]::IsNullOrWhiteSpace($OutputFolder)) {
    New-Item -ItemType Directory -Path $DefaultOutputFolder -Force | Out-Null
}
$outputChoice = $OutputFolder
if ([string]::IsNullOrWhiteSpace($outputChoice)) {
    if ($script:NonInteractive -and -not $script:Force) {
        $outputChoice = $DefaultOutputFolder
    } else {
        $outputPickerInitialFolder = $DefaultOutputFolder
        if ($script:DryRun -and -not (Test-Path -LiteralPath $outputPickerInitialFolder -PathType Container)) {
            $outputPickerInitialFolder = $SourceFolder
        }
        $outputChoice = Select-Folder -Description 'Choose the destination for extracted and compressed files. Eligible files are flattened into this folder.' -InitialFolder $outputPickerInitialFolder
    }
}
if ([string]::IsNullOrWhiteSpace($outputChoice)) {
    Write-Host 'Cancelled. No files were processed.' -ForegroundColor Yellow
    Write-Log -Message 'User cancelled destination selection; no files processed.' -Level WARN
    Stop-RunLogging
    exit 0
}
$script:OutputFolder = $outputChoice
if (-not $script:DryRun) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    Clear-StaleTemporaryFolders -Roots @($env:TEMP, $script:OutputFolder)
}

Write-Host 'Step: Scanning folders for standalone NSP/XCI files.' -ForegroundColor Cyan
$looseScanFolders = @($SourceFolder, $script:OutputFolder)
if (Read-YesNo -Prompt 'Would you like to scan another folder for previously extracted .nsp/.xci files to compress?') {
    $additionalFolder = Select-Folder -Description 'Choose the additional folder to scan for .nsp/.xci files.' -InitialFolder $script:OutputFolder
    if (-not [string]::IsNullOrWhiteSpace($additionalFolder)) { $looseScanFolders += $additionalFolder }
}
# De-duplicate folder paths without changing the user's selected paths.
$uniqueFolders = @()
$seenFolders = @{}
foreach ($folder in $looseScanFolders) {
    $normalized = [System.IO.Path]::GetFullPath($folder).TrimEnd('\').ToLowerInvariant()
    if (-not $seenFolders.ContainsKey($normalized)) { $seenFolders[$normalized] = $true; $uniqueFolders += $folder }
}
$looseCandidates = @(Get-LooseGameFiles -Folders $uniqueFolders)
Write-Log -Message "Scanned loose-file folders: $($uniqueFolders -join '; '); standalone NSP/XCI candidates=$($looseCandidates.Count)." -Level INFO

if ($script:DryRun) {
    Write-Host ''
    Write-Host '===================== Dry run preview =====================' -ForegroundColor Cyan
    Write-Host "Source folder:            $SourceFolder"
    Write-Host "Destination:              $script:OutputFolder"
    Write-Host "RAR archive sets found:   $($detectedArchives.Count)"
    Write-Host "Loose NSP/XCI files found: $($looseCandidates.Count)"
    Write-Host 'No dependencies will be installed, and no files or destination folders will be created, extracted, converted, copied, or deleted.' -ForegroundColor Yellow
    Write-Host 'Diagnostic logs may still be written to the local app data logs directory.' -ForegroundColor Yellow
    Write-Host '==========================================================' -ForegroundColor Cyan
    Write-Host ''
    Write-Log -Message "Dry run preview completed. Archives=$($detectedArchives.Count); standalone files=$($looseCandidates.Count); no processing performed." -Level INFO
    Stop-RunLogging
    exit 0
}

if ($NoCompression) {
    $script:CompressionRequested = $false
    Write-Host 'Compression disabled by -NoCompression. NSP/XCI files will be copied uncompressed.' -ForegroundColor Yellow
    Write-Log -Message 'Compression explicitly disabled by -NoCompression.' -Level INFO
} elseif (($detectedArchives.Count -gt 0 -or $looseCandidates.Count -gt 0) -and
    -not ($script:NonInteractive -and -not $script:Force)) {
    $script:CompressionRequested = Read-YesNo -Prompt 'Compress NSP/XCI files to NSZ/XCZ? Choose No to copy them uncompressed.' -DefaultNo:$false
    if ($script:CompressionRequested) {
        Write-Host 'Compression enabled for this run.' -ForegroundColor Cyan
        Write-Log -Message 'Compression enabled by the user.' -Level INFO
    } else {
        Write-Host 'Compression disabled for this run. NSP/XCI files will be copied uncompressed.' -ForegroundColor Yellow
        Write-Log -Message 'Compression disabled by the user.' -Level INFO
    }
}

# -------------------------- Dependency checks --------------------------
Write-Host 'Step: Checking extraction and compression tools and keys.' -ForegroundColor Cyan
$archiveProcessingEnabled = ($detectedArchives.Count -gt 0)
if ($archiveProcessingEnabled) {
    $script:RarExe = Resolve-RarExecutable
    if (-not $script:RarExe) { Write-Log -Message 'WinRAR command-line executable not found; asking to install it.' -Level WARN; $script:RarExe = Install-WinRAR }
    if ($script:RarExe) { Write-Log -Message "Using WinRAR executable '$script:RarExe'." -Level INFO }
    if (-not $script:RarExe) {
        $archiveProcessingEnabled = $false
        $script:ErrorCount++
        Write-Host 'RAR extraction is disabled because WinRAR could not be located or installed. Source RAR archives will not be touched.' -ForegroundColor Yellow
    }
}

$needsCompression = $script:CompressionRequested -and (($archiveProcessingEnabled -and $detectedArchives.Count -gt 0) -or $looseCandidates.Count -gt 0)
if ($needsCompression) {
    $script:NszExe = Resolve-NszExecutable
    if (-not $script:NszExe) { Write-Log -Message 'No working NSZ CLI was detected; asking to install NSZ into an isolated Python environment.' -Level WARN; $script:NszExe = Install-Nsz }
    if ($script:NszExe) { Write-Log -Message "Using NSZ executable '$script:NszExe'." -Level INFO }
    if ($script:NszExe) {
        $script:KeysFile = Find-KeysFile
        if (-not $script:KeysFile) { $script:KeysFile = Select-KeysFile }
        if ($script:KeysFile) {
            $script:CompressionEnabled = $true
            Write-Log -Message "Compression enabled with NSZ='$script:NszExe' and keys='$script:KeysFile'." -Level INFO
        } else {
            $script:ConversionErrorCount++
            Write-Host 'NSP/XCI compression is disabled because no compatible keys file was selected. Archives can still be extracted, with NSP/XCI files copied uncompressed.' -ForegroundColor Yellow
        }
    } else {
        $script:ConversionErrorCount++
        Write-Host 'NSP/XCI compression is disabled because NSZ is unavailable or installation was declined. Archive contents, if extracted, will be copied without compression.' -ForegroundColor Yellow
    }
}

if ($archiveProcessingEnabled) { $script:Archives = $detectedArchives } else { $script:Archives = @() }
$script:TotalArchives = $script:Archives.Count
if ($script:CompressionEnabled) {
    $script:StandaloneQueue = $looseCandidates
    $script:TotalStandalone = $script:StandaloneQueue.Count
}

if ($script:CompressionEnabled -and $script:TotalStandalone -gt 0) {
    Write-Host "Found $($script:TotalStandalone) existing .nsp/.xci file(s) to convert." -ForegroundColor Cyan
    if ($script:DryRun -or $script:NoDelete) {
        $script:DeleteStandaloneSources = $false
        Write-Host 'Dry-run/no-delete mode is enabled; standalone source files will be preserved.' -ForegroundColor Yellow
    } elseif ($DeleteStandalone) {
        $script:DeleteStandaloneSources = $true
        Write-Host 'CLI override enabled: original standalone NSP/XCI files will be deleted only after verification succeeds.' -ForegroundColor Yellow
    } else {
        $script:DeleteStandaloneSources = Read-YesNo -Prompt 'After successful verification, delete the original standalone .nsp/.xci files to save space?'
        if ($script:DeleteStandaloneSources) {
            Write-Host 'Original standalone NSP/XCI files will be deleted only after the compressed result is verified and successfully copied.' -ForegroundColor Yellow
        } else {
            Write-Host 'Standalone source files will be preserved after conversion.' -ForegroundColor Yellow
        }
    }
}

if ($script:Archives.Count -gt 0 -and -not $script:DryRun -and -not $script:NoDelete) {
    if ($DeleteArchives) {
        $script:DeleteArchivesConfirmed = $true
        $DeleteArchivesOnSuccess = $true
        Write-Host 'CLI override enabled: original RAR archives will be removed after successful verification.' -ForegroundColor Yellow
    } else {
        $script:DeleteArchivesConfirmed = Read-YesNo -Prompt 'After successful verification, delete the original RAR archive sets? This only happens after each eligible file is handled.' -DefaultNo:$true
        $DeleteArchivesOnSuccess = $script:DeleteArchivesConfirmed
        if ($script:DeleteArchivesConfirmed) {
            Write-Host 'RAR archives will be deleted only after every eligible file has been successfully handled.' -ForegroundColor Yellow
        } else {
            Write-Host 'RAR archives will be preserved after processing.' -ForegroundColor Yellow
        }
    }
} else {
    $DeleteArchivesOnSuccess = $false
    if ($DeleteArchives -and ($script:DryRun -or $script:NoDelete)) {
        Write-Host 'Archive deletion is disabled by safety mode or explicit no-delete flag.' -ForegroundColor Yellow
    }
}

if ($script:DryRun -or $script:NoDelete) {
    Write-Host 'Safety mode active: files will not be deleted after processing.' -ForegroundColor Yellow
}

if ($DeleteArchivesOnSuccess) {
    Write-Host 'Archive delete confirmation: enabled' -ForegroundColor Cyan
} else {
    Write-Host 'Archive delete confirmation: disabled' -ForegroundColor Yellow
}

if ($script:Archives.Count -eq 0 -and $script:StandaloneQueue.Count -eq 0) {
    if ($detectedArchives.Count -eq 0 -and $looseCandidates.Count -eq 0) {
        Write-Host "No .rar archives or standalone .nsp/.xci files were found in the selected scan folders." -ForegroundColor Yellow
    } elseif (-not $script:CompressionRequested -and $looseCandidates.Count -gt 0) {
        Write-Host 'Compression was intentionally disabled. Existing standalone NSP/XCI files were left unchanged.' -ForegroundColor Yellow
    } elseif (-not $script:CompressionEnabled -and $looseCandidates.Count -gt 0) {
        Write-Host 'No standalone files were converted. Install/enable NSZ and select a valid keys file, then run the script again.' -ForegroundColor Yellow
    } else {
        Write-Host 'No tasks are enabled. Files and source archives were left untouched.' -ForegroundColor Yellow
    }
    Write-Host "Diagnostic log: $script:DiagnosticsLogPath"
    Write-Host "Transcript:     $script:TranscriptPath"
    if ($script:ErrorCount -gt 0 -or $script:ConversionErrorCount -gt 0 -or $script:FailedCopyCount -gt 0) {
        $script:RunExitCode = 1
        Write-Host 'No work could be completed because one or more required operations failed.' -ForegroundColor Red
        Write-Log -Message "Run ended without enabled work after operational errors; exitCode=$script:RunExitCode." -Level ERROR
    }
    if (-not $script:NonInteractive) { Read-Host 'Press Enter to exit' }
    Stop-RunLogging
    exit $script:RunExitCode
}

Write-Host ''
Write-Host "Source folder:            $SourceFolder"
Write-Host "Destination:              $script:OutputFolder"
Write-Host "RAR archive sets:         $($script:Archives.Count)"
$concurrencyMode = if ($script:AdaptiveConcurrencyEnabled) { 'AUTO (resource-aware)' } else { 'FIXED' }
Write-Host "Concurrency mode:         $concurrencyMode"
$parallelStatus = if ($script:Archives.Count -gt 0) {
    if ($script:AdaptiveConcurrencyEnabled) { "ON (starting at $($script:EffectiveMaxParallel), no fixed job ceiling)" } else { "ON (fixed limit $MaxParallel WinRAR jobs)" }
} else { 'Not needed' }
$compressionStatus = if ($script:CompressionEnabled) { 'ON (NSP -> NSZ, XCI -> XCZ)' } else { 'OFF' }
Write-Host "Parallel extraction:      $parallelStatus"
Write-Host "NSP/XCI compression:      $compressionStatus"
if ($script:CompressionEnabled) {
    if ($script:AdaptiveConcurrencyEnabled) { Write-Host "Parallel compression:     ON (starting at $($script:EffectiveMaxParallelConversions), no fixed job ceiling)" }
    else { Write-Host "Parallel compression:     ON (fixed limit $MaxParallelConversions NSZ jobs)" }
}
Write-Host "Loose files queued:       $($script:StandaloneQueue.Count)"
Write-Host "Log folder:               $script:RunLogFolder"
if ($script:CompressionEnabled) { Write-Host "NSZ executable:           $script:NszExe"; Write-Host "Keys file:                $script:KeysFile" }
if ($DeleteArchivesOnSuccess) { Write-Host 'RAR archives are deleted only when every eligible extracted file has been successfully converted/copied.' -ForegroundColor Yellow }
if (-not $DeleteArchivesOnSuccess) { Write-Host 'RAR archives will be kept because delete confirmation was not granted or safety mode is active.' -ForegroundColor Yellow }
Write-Host ''

# ---------------------- Parallel archive extraction ----------------------
$script:ProgressStage = 'Archives'
Set-DashboardWork -Text 'Waiting for archive extraction to finish' -ClearOperationProgress
Write-Host 'Step: Extracting archives and handling their output files.' -ForegroundColor Cyan
while (-not $script:StopProcessing -and -not $script:DiskReserveBlocked -and -not $script:DiskReserveUnavailable -and
    $script:Active.Count -lt $script:EffectiveMaxParallel -and $script:NextArchiveIndex -lt $script:Archives.Count) { Start-NextExtraction }

while ((-not $script:StopProcessing -and $script:NextArchiveIndex -lt $script:Archives.Count) -or $script:Active.Count -gt 0) {
    # Refill any failed-to-start slots too; otherwise a Start-Job startup error could leave the queue stuck.
    if (-not $script:StopProcessing) {
        while (-not $script:DiskReserveBlocked -and -not $script:DiskReserveUnavailable -and
            $script:Active.Count -lt $script:EffectiveMaxParallel -and $script:NextArchiveIndex -lt $script:Archives.Count) { Start-NextExtraction }
    }
    $finished = @($script:Active | Where-Object { $_.Job.State -in @('Completed', 'Failed', 'Stopped') })

    if ($finished.Count -gt 0) {
        $finishedIds = @($finished | ForEach-Object { $_.Job.Id })
        $script:Active = @($script:Active | Where-Object { $finishedIds -notcontains $_.Job.Id })

        # Refill queue slots immediately: this is the rolling parallel queue, not sequential batches.
        if (-not $script:StopProcessing) {
            while (-not $script:DiskReserveBlocked -and -not $script:DiskReserveUnavailable -and
                $script:Active.Count -lt $script:EffectiveMaxParallel -and $script:NextArchiveIndex -lt $script:Archives.Count) { Start-NextExtraction }
        }

        Process-FinishedArchives -Entries $finished
    } elseif ($script:Active.Count -gt 0) {
        Set-DashboardWork -Text "Waiting for progress from $($script:Active.Count) extraction job(s)"
    }

    Update-ProgressDashboard
    if ($finished.Count -eq 0) { Start-Sleep -Milliseconds 300 }
}

# If Q was selected, any active extraction jobs are allowed to finish, but their files are not processed and their archives are kept.
foreach ($entry in @($script:Active)) {
    try { if ($entry.Job.State -eq 'Running') { Stop-Job -Job $entry.Job -ErrorAction SilentlyContinue } } catch {}
    Remove-ExtractionTemp -Folder $entry.TempFolder -Job $entry.Job
}
$script:Active = @()

# ---------------------- Previously extracted files ----------------------
if ($script:CompressionEnabled -and $script:StandaloneQueue.Count -gt 0 -and -not $script:StopProcessing) {
    $script:ProgressStage = 'Standalone'
    Set-DashboardWork -Text 'Starting standalone NSP/XCI conversions' -ClearOperationProgress
    while ($script:StandaloneNextIndex -lt $script:StandaloneQueue.Count -and -not $script:StopProcessing) {
        $remainingFiles = $script:StandaloneQueue.Count - $script:StandaloneNextIndex
        $batchCount = if ($script:AdaptiveConcurrencyEnabled) { $remainingFiles } else { [Math]::Min($MaxParallelConversions, $remainingFiles) }
        $batch = @($script:StandaloneQueue[$script:StandaloneNextIndex..($script:StandaloneNextIndex + $batchCount - 1)])
        $script:StandaloneNextIndex += $batchCount
        Set-DashboardWork -Text "Converting next $batchCount standalone file(s)" -ClearOperationProgress
        Update-ProgressDashboard
        $outcomes = @(Process-FileBatch -Files $batch -IsStandalone:$true)
        foreach ($outcome in $outcomes) {
            if ($outcome.Result -eq 'Quit') { break }
            $script:StandaloneProcessedCount++
        }
        Update-ProgressDashboard
    }
}

Clear-ProgressDashboard
Write-Host ''
Write-Host '==================== Run summary ====================' -ForegroundColor Cyan
Write-Host "Source folder:             $SourceFolder"
Write-Host "Output folder:            $script:OutputFolder"
Write-Host "Archives detected:        $($detectedArchives.Count)"
Write-Host "Archives processed:       $($script:CompletedArchives)"
Write-Host "Archives deleted:         $($script:DeletedArchiveCount)"
Write-Host "Loose files processed:    $($script:StandaloneProcessedCount)"
Write-Host "Copied to output:         $($script:CopiedCount)"
Write-Host "Skipped:                  $($script:SkippedCount)"
Write-Host "Copy failures:            $($script:FailedCopyCount)"
Write-Host "Extraction errors:        $($script:ErrorCount)"
Write-Host "Warnings:                 $($script:WarningCount)"
Write-Host "Compression successes:   $($script:ConversionSuccessCount)"
Write-Host "Compression failures:     $($script:ConversionErrorCount)"
Write-Host "Diagnostic log:           $script:DiagnosticsLogPath"
Write-Host "Transcript log:           $script:TranscriptPath"
if ($script:StopProcessing) {
    $stopMessage = if ($script:StopReason -eq 'user request') {
        'Processing stopped by user. Active jobs were allowed to finish; unprocessed sources were left untouched.'
    } elseif ($script:StopReason) {
        "Processing stopped because of $($script:StopReason). Unprocessed sources were left untouched."
    } else {
        'Processing stopped. Unprocessed sources were left untouched.'
    }
    Write-Host $stopMessage -ForegroundColor Yellow
}
if (-not $script:CompressionEnabled -and ($detectedArchives.Count -gt 0 -or $looseCandidates.Count -gt 0)) {
    Write-Host 'Compression was disabled; existing loose NSP/XCI files were left untouched, and archive contents were copied uncompressed.' -ForegroundColor Yellow
}
Write-Host '====================================================' -ForegroundColor Cyan
Write-Log -Message "Run summary: archivesDetected=$($detectedArchives.Count); archivesProcessed=$($script:CompletedArchives); archivesDeleted=$($script:DeletedArchiveCount); copied=$($script:CopiedCount); skipped=$($script:SkippedCount); copyFailures=$($script:FailedCopyCount); extractionErrors=$($script:ErrorCount); extractionWarnings=$($script:WarningCount); conversionsPassed=$($script:ConversionSuccessCount); conversionErrors=$($script:ConversionErrorCount); standaloneProcessed=$($script:StandaloneProcessedCount)/$($script:TotalStandalone)." -Level INFO
Write-Log -Message "Run finished. Transcript='$script:TranscriptPath'; diagnosticLog='$script:DiagnosticsLogPath'." -Level INFO
if ($script:ErrorCount -gt 0 -or $script:ConversionErrorCount -gt 0 -or $script:FailedCopyCount -gt 0) {
    $script:RunExitCode = 1
    Write-Host "Run completed with errors. Exit code: $script:RunExitCode" -ForegroundColor Red
    Write-Log -Message "Run completed with operational errors; exitCode=$script:RunExitCode." -Level ERROR
} else {
    Write-Host 'Run completed successfully. Exit code: 0' -ForegroundColor Green
}
if (-not $script:NonInteractive) { Read-Host 'Press Enter to exit' }
exit $script:RunExitCode
} finally {
    Stop-ActiveWork
    Stop-RunLogging
}