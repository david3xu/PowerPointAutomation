# Increase Process Memory Limits for PowerPoint Automation
#
# This script modifies Windows Registry settings to allow processes
# to use more memory and handles, which helps with PowerPoint Automation
#
# Run this script as administrator before using PowerPoint Automation
# with large presentations or complex diagrams

param (
    [switch]$RestoreDefaults = $false
)

# Define registry paths
$windowsRegPath = "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Windows"
$memoryManagerRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management"

# Default values to restore if needed
$defaultHandleQuota = 10000
$defaultPoolUsageMaximum = 80

# New values to set
$newHandleQuota = 18000
$newPoolUsageMaximum = 60  # Lower percentage reserves more memory for process

function EnsureAdministrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal(
        [Security.Principal.WindowsIdentity]::GetCurrent())
    $isAdmin = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    
    if (-not $isAdmin) {
        Write-Host "This script requires administrative privileges." -ForegroundColor Red
        Write-Host "Please run PowerShell as administrator and try again." -ForegroundColor Red
        exit 1
    }
}

function SetRegistryValue($path, $name, $value) {
    try {
        if (Test-Path $path) {
            Set-ItemProperty -Path $path -Name $name -Value $value -Type DWORD -Force
            Write-Host "Successfully set $name to $value" -ForegroundColor Green
        } else {
            Write-Host "Registry path $path not found" -ForegroundColor Red
            return $false
        }
        return $true
    } catch {
        Write-Host "Error setting registry value: $_" -ForegroundColor Red
        return $false
    }
}

function GetRegistryValue($path, $name) {
    try {
        if (Test-Path $path) {
            $value = (Get-ItemProperty -Path $path -Name $name -ErrorAction SilentlyContinue).$name
            if ($null -eq $value) {
                Write-Host "Registry value $name not found in $path" -ForegroundColor Yellow
                return $null
            }
            return $value
        } else {
            Write-Host "Registry path $path not found" -ForegroundColor Red
            return $null
        }
    } catch {
        Write-Host "Error getting registry value: $_" -ForegroundColor Red
        return $null
    }
}

function IncreaseProcessMemory {
    Write-Host "Increasing process memory limits for PowerPoint Automation..." -ForegroundColor Cyan
    
    # Backup current values
    $currentHandleQuota = GetRegistryValue $windowsRegPath "USERProcessHandleQuota"
    $currentPoolUsageMaximum = GetRegistryValue $memoryManagerRegPath "PoolUsageMaximum"
    
    Write-Host "Current settings:" -ForegroundColor White
    Write-Host "  - USERProcessHandleQuota: $currentHandleQuota" -ForegroundColor White
    Write-Host "  - PoolUsageMaximum: $currentPoolUsageMaximum" -ForegroundColor White
    
    # Store current values in temporary files for restoration
    if ($currentHandleQuota) {
        $currentHandleQuota | Out-File -FilePath "$env:TEMP\PPT_HandleQuota.txt" -Force
    }
    if ($currentPoolUsageMaximum) {
        $currentPoolUsageMaximum | Out-File -FilePath "$env:TEMP\PPT_PoolUsage.txt" -Force
    }
    
    # Set new values
    $success1 = SetRegistryValue $windowsRegPath "USERProcessHandleQuota" $newHandleQuota
    $success2 = SetRegistryValue $memoryManagerRegPath "PoolUsageMaximum" $newPoolUsageMaximum
    
    if ($success1 -and $success2) {
        Write-Host "`nProcess memory limits have been increased successfully." -ForegroundColor Green
        Write-Host "You should restart your computer for these changes to take effect." -ForegroundColor Yellow
        
        $restart = Read-Host "Would you like to restart your computer now? (y/n)"
        if ($restart -eq "y" -or $restart -eq "Y") {
            Restart-Computer -Force
        } else {
            Write-Host "Please restart your computer before running PowerPoint Automation." -ForegroundColor Yellow
        }
    } else {
        Write-Host "`nFailed to set all registry values. Some changes may not have been applied." -ForegroundColor Red
    }
}

function RestoreProcessMemory {
    Write-Host "Restoring default process memory limits..." -ForegroundColor Cyan
    
    $storedHandleQuota = $defaultHandleQuota
    $storedPoolUsageMaximum = $defaultPoolUsageMaximum
    
    # Try to read backup values
    $handleQuotaFile = "$env:TEMP\PPT_HandleQuota.txt"
    $poolUsageFile = "$env:TEMP\PPT_PoolUsage.txt"
    
    if (Test-Path $handleQuotaFile) {
        $storedHandleQuota = [int](Get-Content $handleQuotaFile -Raw)
    }
    
    if (Test-Path $poolUsageFile) {
        $storedPoolUsageMaximum = [int](Get-Content $poolUsageFile -Raw)
    }
    
    Write-Host "Restoring to:" -ForegroundColor White
    Write-Host "  - USERProcessHandleQuota: $storedHandleQuota" -ForegroundColor White
    Write-Host "  - PoolUsageMaximum: $storedPoolUsageMaximum" -ForegroundColor White
    
    # Set original values
    $success1 = SetRegistryValue $windowsRegPath "USERProcessHandleQuota" $storedHandleQuota
    $success2 = SetRegistryValue $memoryManagerRegPath "PoolUsageMaximum" $storedPoolUsageMaximum
    
    if ($success1 -and $success2) {
        Write-Host "`nDefault process memory limits have been restored successfully." -ForegroundColor Green
        Write-Host "You should restart your computer for these changes to take effect." -ForegroundColor Yellow
        
        $restart = Read-Host "Would you like to restart your computer now? (y/n)"
        if ($restart -eq "y" -or $restart -eq "Y") {
            Restart-Computer -Force
        }
    } else {
        Write-Host "`nFailed to restore all registry values. Some changes may not have been applied." -ForegroundColor Red
    }
    
    # Clean up temporary files
    if (Test-Path $handleQuotaFile) { Remove-Item $handleQuotaFile -Force }
    if (Test-Path $poolUsageFile) { Remove-Item $poolUsageFile -Force }
}

# Main script execution
EnsureAdministrator

if ($RestoreDefaults) {
    RestoreProcessMemory
} else {
    IncreaseProcessMemory
} 