function Get-OHMFullHardwareInfo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$OhmReporterPath = "C:\Program Files\OpenHardwareMonitorReport\OpenHardwareMonitorReport.exe" # Caminho para o executável do OHM Reporter
    )

    function ConvertTo-AsciiSafe {
        param([string]$InputString)
        if ([string]::IsNullOrWhiteSpace($InputString)) { return "" }
        $normalized = $InputString.Normalize([System.Text.NormalizationForm]::FormD)
        $builder = New-Object System.Text.StringBuilder
        foreach ($char in $normalized.ToCharArray()) {
            $cat = [System.Globalization.CharUnicodeInfo]::GetUnicodeCategory($char)
            if ($cat -ne [System.Globalization.CharUnicodeInfo]::NonSpacingMark -and [int]$char -lt 128) {
                [void]$builder.Append($char)
            }
        }
        return ($builder.ToString().Normalize([System.Text.NormalizationForm]::FormC)).Trim()
    }

    if (-not (Test-Path $OhmReporterPath)) {
        Write-Error "OpenHardwareMonitorReport.exe not found: $OhmReporterPath" *>&1 | Out-Null
        return $null
    }

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
        
        $ReportContent = $ReportContentRaw -split "`n" | ForEach-Object { ConvertTo-AsciiSafe -InputString $_ } | Where-Object { $_.Trim() -ne "" }
        
    } catch {
        Write-Error "Falha ao executar OpenHardwareMonitorReport.exe: $($_.Exception.Message)" *>&1 | Out-Null
        return $null
    }
    
    if (-not $ReportContent) {
        Write-Error "Relatorio vazio ou nao gerado pelo OpenHardwareMonitorReport.exe." *>&1 | Out-Null
        return $null
    }


    $metrics = @()
    $hostname = ConvertTo-AsciiSafe -InputString $env:COMPUTERNAME

    # Variáveis de estado para o parser
    $currentSection = ""
    $currentHardwareName = "" # Para "AMD A78FX" ou "AMD FX-8320"
    $currentHardwarePath = "" # Para "/mainboard" ou "/amdcpu/0"
    $currentMemoryDeviceIndex = -1 # Para rastrear dispositivos de memória

    foreach ($line in $ReportContent) {
        # --- Seções Principais ---
        if ($line -eq "Sensors") { $currentSection = "Sensors"; continue }
        if ($line -eq "Parameters") { $currentSection = "Parameters"; continue }
        if ($line -eq "Mainboard") { $currentSection = "Mainboard"; continue }
        if ($line -eq "Processor") { $currentSection = "Processor"; continue }
        if ($line -eq "Generic Memory") { $currentSection = "Generic Memory"; continue }
        if ($line -eq "NVIDIA GPU") { $currentSection = "NVIDIA GPU"; continue }
        if ($line -eq "GenericHarddisk") { $currentSection = "GenericHarddisk"; continue }
        # Adicione outras seções principais conforme necessário (ex: AMD CPU, etc.)

        # --- Linhas de Divisão ou Rodapé ---
        if ($line -match "^-{20,}$" -or $line -match "^CPUID$") {
            $currentSection = ""; # Reseta a seção para ignorar o que vem depois
            continue
        }

        # --- Parsing dentro das Seções ---

        if ($currentSection -eq "Sensors") {
            # Extrair nomes de hardware como "AMD A78FX" ou "AMD FX-8320" e seus paths
            if ($line -match "^\+\-\s*(.*)\s*\((.*)\)$") {
                $currentHardwareName = ConvertTo-AsciiSafe -InputString $Matches[1]
                $currentHardwarePath = $Matches[2].Trim()
                continue
            }

            # Identifica sensores de temperatura ou carga
            if ($line -match "^\s*\|\s+\+\-\s*(.+?)\s*:\s*([0-9\.\-]+)\s+([0-9\.\-]+)\s+([0-9\.\-]+)\s*\((.+?)\)$") {
                $sensorName = ConvertTo-AsciiSafe -InputString $Matches[1]
                $currentValue = [double]$Matches[2]
                $minValue = [double]$Matches[3]
                $maxValue = [double]$Matches[4]
                $sensorPath = $Matches[5].Trim()

                $sensorType = "other"
                $measurementName = ""

                if ($sensorPath -like "*/temperature/*") {
                    if ($sensorPath -like "/amdcpu/*/temperature/*") { $sensorType = "CPU_Temp" }
                    elseif ($sensorPath -like "/nvidiagpu/*/temperature/*") { $sensorType = "GPU_Temp" }
                    elseif ($sensorPath -like "/hdd/*/temperature/*") { $sensorType = "HDD_Temp" }
                    else { $sensorType = "Mainboard_Temp" } # Para as Temperature #1 a #6
                    $measurementName = "hardware_temp"

                    $metrics += [PSCustomObject]@{
                        "__measurement__" = $measurementName
                        "__tags__"        = @{
                            "host"        = $hostname
                            "component"   = $currentHardwareName # Agora usando o nome mais genérico da seção
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
                    if ($sensorPath -like "/amdcpu/*/load/*") { $sensorType = "CPU_Load" }
                    elseif ($sensorPath -like "/nvidiagpu/*/load/*") { $sensorType = "GPU_Load" }
                    elseif ($sensorPath -like "/ram/load/*") { $sensorType = "RAM_Load" }
                    elseif ($sensorPath -like "/hdd/*/load/*") { $sensorType = "HDD_Load" }
                    else { $sensorType = "Other_Load" }
                    $measurementName = "hardware_load"

                    $metrics += [PSCustomObject]@{
                        "__measurement__" = $measurementName
                        "__tags__"        = @{
                            "host"        = $hostname
                            "component"   = $currentHardwareName
                            "sensor_name" = $sensorName
                            "sensor_type" = $sensorType
                        }
                        "value_percent" = $currentValue
                    }
                }
            }
        }
        elseif ($currentSection -eq "Mainboard") {
            if ($line -match "^BIOS Vendor:\s+(.*)$") {
                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "Motherboard"
                        "info_type"   = "BIOS"
                    }
                    "bios_vendor"   = ConvertTo-AsciiSafe -InputString $Matches[1].Trim()
                    "bios_version"  = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "BIOS Version:" | Select-Object -First 1).ToString() -replace "BIOS Version:\s+", "").Trim()
                }
            } elseif ($line -match "^Mainboard Manufacturer:\s+(.*)$") {
                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "Motherboard"
                        "info_type"   = "Board_Details"
                        "manufacturer"= ConvertTo-AsciiSafe -InputString $Matches[1].Trim()
                        "name"        = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Mainboard Name:" | Select-Object -First 1).ToString() -replace "Mainboard Name:\s+", "").Trim()
                        "version"     = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Mainboard Version:" | Select-Object -First 1).ToString() -replace "Mainboard Version:\s+", "").Trim()
                    }
                }
            }
            # Note: Serial Number da Mainboard não está diretamente no relatório OHM, precisaria de WMI separado.
        }
        elseif ($currentSection -eq "Processor") {
            if ($line -match "^Processor Vendor:\s+(.*)$") {
                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "CPU_Details"
                        "manufacturer"= ConvertTo-AsciiSafe -InputString $Matches[1].Trim()
                        "name"        = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Processor Brand:" | Select-Object -First 1).ToString() -replace "Processor Brand:\s+", "").Trim()
                    }
                    "cores"           = [int](($ReportContent | Select-String "Processor Core Count:" | Select-Object -First 1).ToString() -replace "Processor Core Count:\s+", "").Trim()
                    "threads"         = [int](($ReportContent | Select-String "Processor Thread Count:" | Select-Object -First 1).ToString() -replace "Processor Thread Count:\s+", "").Trim()
                }
            }
        }
        elseif ($currentSection -eq "Generic Memory") {
            # Captura a capacidade total de RAM se tiver uma linha "Memory:" na seção de sensores
            # Isso é para pegar o total, pois "Used/Available Memory" são "load/data" sensors
            if ($line -match "^Memory Device \[(\d+)\] Manufacturer:\s*(.*)$") {
                $currentMemoryDeviceIndex = [int]$Matches[1]
                $manufacturer = ConvertTo-AsciiSafe -InputString $Matches[2].Trim()
                $partNumber = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Memory Device \[$currentMemoryDeviceIndex\] Part Number:" | Select-Object -First 1).ToString() -replace "Memory Device \[$currentMemoryDeviceIndex\] Part Number:\s+", "").Trim()
                $speed = [int](($ReportContent | Select-String "Memory Device \[$currentMemoryDeviceIndex\] Speed:" | Select-Object -First 1).ToString() -replace "Memory Device \[$currentMemoryDeviceIndex\] Speed:\s+", "").Trim()

                # Capacidade da RAM nao e diretamente aqui, geralmente esta em Used/Available Memory nos sensores
                # Mas o WMI original pega o 'Capacity'. OHM reporta 'Used Memory' e 'Available Memory'
                # Para uma estimativa da capacidade total, podemos somar os 'Used' e 'Available' do ultimo sensor
                
                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "RAM_Stick"
                        "manufacturer"= $manufacturer
                        "part_number" = $partNumber
                        "index"       = "$currentMemoryDeviceIndex"
                        # Serial Number da RAM nao esta disponivel no relatorio OHM
                    }
                    "speed_mhz"       = $speed
                    # 'capacity_gb' pode ser adicionado se encontrado em outra linha ou calculado a partir de 'Used/Available'
                }
            }
        }
        elseif ($currentSection -eq "NVIDIA GPU") {
            if ($line -match "^Name:\s+(.*)$") {
                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "GPU_Details"
                        "name"        = ConvertTo-AsciiSafe -InputString $Matches[1].Trim()
                        "driver_version" = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Driver Version:" | Select-Object -First 1).ToString() -replace "Driver Version:\s+", "").Trim()
                    }
                }
            }
        }
        elseif ($currentSection -eq "GenericHarddisk") {
            # Se for a linha de "Drive name:" para pegar o modelo do disco
            if ($line -match "^Drive name:\s+(.*)$") {
                $driveName = ConvertTo-AsciiSafe -InputString $Matches[1].Trim()
                $firmware = ConvertTo-AsciiSafe -InputString (($ReportContent | Select-String "Firmware version:" | Select-Object -First 1).ToString() -replace "Firmware version:\s+", "").Trim()
                
                # Serial number do disco já está vindo do CrystalDiskInfo.
                # Se não usasse CrystalDiskInfo, poderia tentar extrair o serial daqui
                # mas o OHM reporta "ID Description Raw Value ...", não o Serial Number direto.

                $metrics += [PSCustomObject]@{
                    "__measurement__" = "system_hardware_info"
                    "__tags__"        = @{
                        "host"        = $hostname
                        "component"   = "HDD_Details"
                        "drive_name"  = $driveName
                        "firmware"    = $firmware
                    }
                }
            }
        }
    }

    # Final: Saída JSON
    $jsonOutput = @()
    foreach ($m in $metrics) {
        # Remova campos null para não aparecerem no JSON
        $m.PSObject.Properties | Where-Object { $_.Value -eq $null } | ForEach-Object { $m.PSObject.Properties.Remove($_) }
        $jsonOutput += $m
    }
    $jsonOutput | ConvertTo-Json -Depth 5 -Compress
}

# ---