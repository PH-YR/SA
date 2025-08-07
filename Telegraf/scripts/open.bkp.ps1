function Get-OHMDataFromConsole {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OhmReporterPath = "C:\Program Files\OpenHardwareMonitorReport\OpenHardwareMonitorReport.exe"
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

    # Removido: DEBUG: Verificando executavel do OHM Reporter...
    if (-not (Test-Path $OhmReporterPath)) {
        Write-Error "OpenHardwareMonitorReport.exe not found: $OhmReporterPath" *>&1 | Out-Null
        return $null
    }
    # Removido: DEBUG: Executavel do OHM Reporter encontrado.

    # Removido: DEBUG: Executando OpenHardwareMonitorReport.exe reporttoconsole e capturando saida...
    try {
        $processInfo = New-Object System.Diagnostics.ProcessStartInfo
        $processInfo.FileName = $OhmReporterPath
        $processInfo.Arguments = "reporttoconsole"
        $processInfo.RedirectStandardOutput = $true
        $processInfo.UseShellExecute = $false
        $processInfo.CreateNoWindow = $true

        $process = New-Object System.Diagnostics.Process
        $process.StartInfo = $processInfo
        $process.Start() | Out-Null

        $ReportContentRaw = ($process.StandardOutput.ReadToEnd() | Out-String).Replace("`r`n", "`n").Replace("`r", "`n")
        $process.WaitForExit()

        # Removido: DEBUG: Saida bruta do OHM Reporter (primeiras 500 chars):
        # Removido: Write-Host "  $($ReportContentRaw.Substring(0, [System.Math]::Min(500, $ReportContentRaw.Length)))"
        # Removido: DEBUG: Fim do preview da saida bruta.

        $ReportContent = $ReportContentRaw -split "`n" | ForEach-Object { ConvertTo-AsciiSafe -InputString $_ } | Where-Object { $_.Trim() -ne "" }
        
        # Removido: DEBUG: Conteudo do relatorio apos split e AsciiSafe (primeiras 20 linhas):
        # Removido: $ReportContent | Select-Object -First 20 | ForEach-Object { Write-Host "  $_" }
        # Removido: DEBUG: Fim do preview do conteudo processado.

    } catch {
        Write-Error "Falha ao executar OpenHardwareMonitorReport.exe: $($_.Exception.Message)" *>&1 | Out-Null
        return $null
    }
    
    if (-not $ReportContent) {
        Write-Error "Relatorio vazio ou nao gerado pelo OpenHardwareMonitorReport.exe." *>&1 | Out-Null
        return $null
    }


    $metrics = @()
    $currentHardware = ""
    $currentPath = ""

    foreach ($line in $ReportContent) {
        if ($line -match "^\+\-\s*(.*)\s*\((.*)\)$") {
            $currentHardware = ConvertTo-AsciiSafe -InputString $Matches[1]
            $currentPath = $Matches[2].Trim()
            continue
        }

        if ($line -match "^\s*\|\s+\+\-\s*(.+?)\s*:\s*([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s*\((.+?)\)$") {
            $sensorName = ConvertTo-AsciiSafe -InputString $Matches[1]
            $currentValue = [double]$Matches[2]
            $minValue = [double]$Matches[3]
            $maxValue = [double]$Matches[4]
            $sensorPath = $Matches[5].Trim()

            $sensorType = "other"
            $measurementName = ""

            if ($sensorPath -like "*/temperature/*") {
                if ($sensorPath -like "/amdcpu/*/temperature/*") {
                    $sensorType = "CPU_Temp"
                } elseif ($sensorPath -like "/nvidiagpu/*/temperature/*") {
                    $sensorType = "GPU_Temp"
                } elseif ($sensorPath -like "/hdd/*/temperature/*") {
                    $sensorType = "HDD_Temp"
                } else {
                    $sensorType = "Mainboard_Temp"
                }
                $measurementName = "hardware_temp"

                $metrics += [PSCustomObject]@{
                    "__measurement__" = $measurementName
                    "__tags__"        = @{
                        "host"        = "$env:COMPUTERNAME"
                        "component"   = $currentHardware
                        "sensor_name" = $sensorName
                        "sensor_type" = $sensorType
                        "sensor_path" = $sensorPath
                    }
                    "value_c"       = $currentValue
                    "min_c"           = $minValue
                    "max_c"           = $maxValue
                }
            }
            elseif ($sensorPath -like "*/load/*") {
                if ($sensorPath -like "/amdcpu/*/load/*") {
                    $sensorType = "CPU_Load"
                } elseif ($sensorPath -like "/nvidiagpu/*/load/*") {
                    $sensorType = "GPU_Load"
                } elseif ($sensorPath -like "/ram/load/*") {
                    $sensorType = "RAM_Load"
                } elseif ($sensorPath -like "/hdd/*/load/*") {
                    $sensorType = "HDD_Load"
                } else {
                    $sensorType = "Other_Load"
                }
                $measurementName = "hardware_load"

                $metrics += [PSCustomObject]@{
                    "__measurement__" = $measurementName
                    "__tags__"        = @{
                        "host"        = "$env:COMPUTERNAME"
                        "component"   = $currentHardware
                        "sensor_name" = $sensorName
                        "sensor_type" = $sensorType
                    }
                    "value_percent" = $currentValue
                }
            }
        }
    }

    # Removido: DEBUG: Total de metricas parseadas: X
    $jsonOutput = @()
    foreach ($m in $metrics) {
        $obj = [PSCustomObject]@{
            "__measurement__" = $m."__measurement__"
            "__tags__"        = $m."__tags__"
            "value_c"         = $m.value_c
            "min_c"           = $m.min_c
            "max_c"           = $m.max_c
            "value_percent"   = $m.value_percent
        }
        $obj.PSObject.Properties | Where-Object { $_.Value -eq $null } | ForEach-Object { $obj.PSObject.Properties.Remove($_) }

        $jsonOutput += $obj
    }
    
    # Removido: DEBUG: Gerando JSON final...
    $jsonOutput | ConvertTo-Json -Depth 5 -Compress
}

# --- Chamada Principal do Script ---
Get-OHMDataFromConsole -OhmReporterPath "C:\Program Files\OpenHardwareMonitorReport\OpenHardwareMonitorReport.exe"