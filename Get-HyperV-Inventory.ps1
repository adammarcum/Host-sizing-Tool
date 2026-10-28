<#
.SYNOPSIS
    Hyper-V Inventory Collector for Infrastructure Sizing
    Outputs: 2 CSV files (HyperV_vInfo.csv, HyperV_vHost.csv)
    
.DESCRIPTION
    Scans local Host or Cluster Nodes to gather VM entitlements and Host hardware specs.
    designed to mimic RVTools headers for easy ingestion into sizing tools.
    
.NOTES
    Requires: Run as Administrator
    Permissions: Local Admin on target Hyper-V hosts
#>

$ErrorActionPreference = "SilentlyContinue"
$ProgressPreference = "SilentlyContinue"

# --- 1. DETECT SCOPE ---
Write-Host "🔍 Detecting Environment..." -ForegroundColor Cyan
$hostsToScan = @()
$clusterName = "Standalone"

try {
    # Try to get cluster nodes if the module exists
    if (Get-Module -ListAvailable FailoverClusters) {
        $nodes = Get-ClusterNode
        if ($nodes) {
            $clusterName = (Get-Cluster).Name
            $hostsToScan = $nodes.Name
            Write-Host "   Found Cluster: $clusterName ($($hostsToScan.Count) nodes)" -ForegroundColor Green
        }
    }
} catch {
    # Fallback to local machine if cluster check fails
}

# If no cluster found, scan local machine
if ($hostsToScan.Count -eq 0) {
    $hostsToScan = @($env:COMPUTERNAME)
    Write-Host "   No Cluster detected. Scanning local host: $env:COMPUTERNAME" -ForegroundColor Yellow
}

# --- 2. DATA COLLECTIONS ---
$vInfoData = @()
$vHostData = @()

foreach ($computer in $hostsToScan) {
    Write-Host "   ...Scanning $computer" -NoNewline
    
    try {
        # --- HOST HARDWARE (WMI/CIM) ---
        $procStats = Get-CimInstance -ClassName Win32_Processor -ComputerName $computer
        $memStats = Get-CimInstance -ClassName Win32_PhysicalMemory -ComputerName $computer | Measure-Object -Property Capacity -Sum
        
        # Calculate Sockets & Cores
        $sockets = @($procStats).Count
        $coresPerSocket = $procStats[0].NumberOfCores
        $totalRamMB = [math]::Round($memStats.Sum / 1MB, 0)
        
        # Add to Host Report
        $vHostData += [PSCustomObject]@{
            'Host'          = $computer
            'Cluster'       = $clusterName
            '# CPU'         = $sockets
            'Cores per CPU' = $coresPerSocket
            '# Memory'      = $totalRamMB
            'Model'         = (Get-CimInstance -ClassName Win32_ComputerSystem -ComputerName $computer).Model
        }

        # --- VM INVENTORY ---
        $vms = Get-VM -ComputerName $computer
        
        foreach ($vm in $vms) {
            # Calculate Storage (Sum of all VHDX files)
            $storageBytes = 0
            $disks = $vm | Get-VMHardDiskDrive
            foreach ($disk in $disks) {
                if ($disk.Path -and (Test-Path $disk.Path)) {
                    $fileItem = Get-Item $disk.Path
                    $storageBytes += $fileItem.Length
                }
            }
            $storageMB = [math]::Round($storageBytes / 1MB, 0)

            # Add to VM Report
            $vInfoData += [PSCustomObject]@{
                'VM'             = $vm.Name
                'Powerstate'     = if ($vm.State -eq 'Running') { "poweredOn" } else { "poweredOff" }
                'CPUs'           = $vm.ProcessorCount
                'Memory'         = [math]::Round($vm.MemoryAssigned / 1MB, 0)
                'Provisioned MB' = $storageMB
                'In Use MB'      = $storageMB # Hyper-V doesn't easily expose "Thin provisioned used" vs "allocated" without deeper inspection
                'Cluster'        = $clusterName
                'Host'           = $computer
                'OS'             = "Windows/Linux" # Generic placeholder as getting guest OS requires integration services
            }
        }
        Write-Host " [OK]" -ForegroundColor Green

    } catch {
        Write-Host " [FAILED] - Check Permissions/RPC" -ForegroundColor Red
    }
}

# --- 3. EXPORT ---
$pathInfo = "$PSScriptRoot\HyperV_vInfo.csv"
$pathHost = "$PSScriptRoot\HyperV_vHost.csv"

$vInfoData | Export-Csv -Path $pathInfo -NoTypeInformation
$vHostData | Export-Csv -Path $pathHost -NoTypeInformation

Write-Host "`n✅ Collection Complete!" -ForegroundColor Cyan
Write-Host "Files generated:"
Write-Host "1. $pathInfo"
Write-Host "2. $pathHost"
Write-Host "`nPlease zip these two files and send them for analysis." -ForegroundColor Yellow