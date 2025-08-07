function Get-CrystalDiskInfoStatus {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$false)]
        [string]$CrystalDiskInfoPath = "C:\Program Files\CrystalDiskInfo\DiskInfo64.exe",
        [Parameter(Mandatory=$false)]
        [string]$OutputPath = "$env:TEMP\DiskInfo.txt"
    )

    function ConvertTo-AsciiSafe {
        param([string]$InputString)
        if ([string]::IsNullOrWhiteSpace($InputString)) { return "" }
        $normalized = $InputString.Normalize([System.Text.NormalizationForm]::FormD)
        $builder = New-Object System.Text.StringBuilder
        foreach ($char in $normalized.ToCharArray()) {
            $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
            if ($cat -ne [System.Globalization.UnicodeCategory]::NonSpacingMark -and [int]$char -lt 128) {
                [void]$builder.Append($char)
            }
        }
        return ($builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)).Trim()
    }

    if (-not (Test-Path $CrystalDiskInfoPath)) {
        Write-Error "CrystalDiskInfo not found: $CrystalDiskInfoPath" *>&1 | Out-Null
        return $null
    }

    $OutputDir = Split-Path $OutputPath
    if (-not (Test-Path $OutputDir)) {
        New-Item -Path $OutputDir -ItemType Directory -Force | Out-Null
    }

    try {
        Start-Process -FilePath $CrystalDiskInfoPath -ArgumentList "/CopyExit" -Wait -WindowStyle Hidden -ErrorAction Stop | Out-Null
        $defaultOutput = Join-Path (Split-Path $CrystalDiskInfoPath) "DiskInfo.txt"
        if (Test-Path $defaultOutput) {
            Move-Item -Path $defaultOutput -Destination $OutputPath -Force | Out-Null
        } else {
            Write-Warning "DiskInfo.txt não encontrado após execução" *>&1 | Out-Null
            return $null
        }
    } catch {
        Write-Error "Erro ao rodar CrystalDiskInfo: $($_.Exception.Message)" *>&1 | Out-Null
        return $null
    }

    if (-not (Test-Path $OutputPath)) {
        Write-Error "Arquivo de saída não encontrado: $OutputPath" *>&1 | Out-Null
        return $null
    }

    $DiskInfoContent = Get-Content -Path $OutputPath
    $Disks = @()
    $currentDisk = $null

    foreach ($line in $DiskInfoContent) {
        $line = $line.Trim()
        if ($line -match "^-{20,}$" -or $line -match "^-{20,} \((\d+)\)") {
            if ($currentDisk) { $Disks += $currentDisk }
            $currentDisk = [PSCustomObject]@{
                Model          = ""
                SerialNumber   = ""
                HealthStatus   = ""
                HealthPercent  = $null
                TemperatureC   = ""
                TemperatureF   = ""
            }
            continue
        }

        if ($line -match "^Model\s+:\s+(.*)") {
            $currentDisk.Model = ConvertTo-AsciiSafe -InputString $Matches[1]
        } elseif ($line -match "^Serial Number\s+:\s+(.*)") {
            $currentDisk.SerialNumber = ConvertTo-AsciiSafe -InputString $Matches[1]
        } elseif ($line -match "^Health Status\s+:\s+(.*)") {
            $statusText = ConvertTo-AsciiSafe -InputString $Matches[1]
            $currentDisk.HealthStatus = $statusText
            if ($statusText -match "\((\d+)\s*%\)") {
                $currentDisk.HealthPercent = [int]$Matches[1]
            }
        } elseif ($line -match "^Temperature\s+:\s+(\d+)\s+C\s+\((\d+)\s+F\)") {
            $currentDisk.TemperatureC = [int]$Matches[1]
            $currentDisk.TemperatureF = [int]$Matches[2]
        }
    }

    if ($currentDisk) {
        $Disks += $currentDisk
    }

    Remove-Item $OutputPath -ErrorAction SilentlyContinue

    $telegrafMetrics = @()
    foreach ($disk in $Disks) {
        if ([string]::IsNullOrWhiteSpace($disk.Model)) { continue }

        $metric = [PSCustomObject]@{
            "Model"         = $disk.Model
            "SerialNumber"  = $disk.SerialNumber
            #"HealthStatus"  = $disk.HealthStatus  # ← Remova o comentário se quiser manter o texto
            "HealthPercent" = $disk.HealthPercent
            "TemperatureC"  = $disk.TemperatureC
            "TemperatureF"  = $disk.TemperatureF
            "host"          = "$env:COMPUTERNAME"
        }
        $telegrafMetrics += $metric
    }

    @{ "disks" = $telegrafMetrics } | ConvertTo-Json -Depth 5 -Compress
}

Get-CrystalDiskInfoStatus
