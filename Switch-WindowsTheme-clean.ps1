<#
.SYNOPSIS
    Switch Windows theme mode and accent colors manually or on a schedule.

.DESCRIPTION
    This script allows you to switch between Windows Light and Dark themes,
    change accent colors, and set up automatic theme switching on a schedule.

.PARAMETER Mode
    Theme mode to set: Light, Dark, or Toggle

.PARAMETER AccentColor
    Accent color to set (Red, Orange, Yellow, Green, Cyan, Blue, Purple, Pink, Default)

.PARAMETER Wallpaper
    Path to wallpaper image to set

.PARAMETER SetupSchedule
    Create scheduled tasks based on configuration file

.PARAMETER RemoveSchedule
    Remove all WindowsThemeSwitcher scheduled tasks

.PARAMETER ConfigFile
    Path to configuration file (default: .\Switch-WindowsTheme.json)

.EXAMPLE
    .\Switch-WindowsTheme.ps1 -Mode Toggle
    Toggles between Light and Dark mode

.EXAMPLE
    .\Switch-WindowsTheme.ps1 -SetupSchedule
    Sets up automatic theme switching based on configuration

.EXAMPLE
    .\Switch-WindowsTheme.ps1 -Mode Dark -AccentColor Purple
    Sets dark mode with purple accent color
#>

param(
    [Parameter(Mandatory=$false)]
    [ValidateSet("Light", "Dark", "Toggle")]
    [string]$Mode,
    
    [Parameter(Mandatory=$false)]
    [ValidateSet("Red", "Orange", "Yellow", "Green", "Cyan", "Blue", "Purple", "Pink", "Default")]
    [string]$AccentColor,
    
    [Parameter(Mandatory=$false)]
    [string]$Wallpaper,
    
    [Parameter(Mandatory=$false)]
    [switch]$SetupSchedule,
    
    [Parameter(Mandatory=$false)]
    [switch]$RemoveSchedule,
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\Switch-WindowsTheme.json"
)

# Default configuration
$DefaultConfig = @{
    schedules = @(
        @{
            name = "Morning Light"
            time = "07:00"
            mode = "Light"
            accentColor = "Blue"
            wallpaper = ""
        },
        @{
            name = "Evening Dark"
            time = "19:00"
            mode = "Dark"
            accentColor = "Purple"
            wallpaper = ""
        }
    )
}

# Add required Windows API types
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public class WinAPI {
    [DllImport("user32.dll", SetLastError = true)]
    public static extern int SendMessageTimeout(
        IntPtr hWnd,
        uint Msg,
        IntPtr wParam,
        IntPtr lParam,
        uint fuFlags,
        uint uTimeout,
        out IntPtr lpdwResult
    );

    [DllImport("user32.dll", SetLastError = true)]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

    [DllImport("shell32.dll")]
    public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);

    [DllImport("user32.dll")]
    public static extern bool SetSysColors(int cElements, int[] lpaElements, int[] lpaRgbValues);

    public const int HWND_BROADCAST = 0xFFFF;
    public const int WM_SETTINGCHANGE = 0x001A;
    public const int SMTO_ABORTIFHUNG = 0x0002;
    public const int SHCNE_ASSOCCHANGED = 0x08000000;
    public const int SHCNF_IDLIST = 0x0000;
}
"@

# Function to load configuration
function Load-Config {
    if (Test-Path $ConfigFile) {
        try {
            $content = Get-Content $ConfigFile -Raw | ConvertFrom-Json
            return $content
        } catch {
            Write-Warning "Failed to load config file: $($_.Exception.Message)"
            Write-Host "Using default configuration..."
            return $DefaultConfig
        }
    } else {
        Write-Host "Config file not found. Using default configuration..."
        return $DefaultConfig
    }
}

# Function to get current theme
function Get-CurrentTheme {
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        $lightMode = Get-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -ErrorAction SilentlyContinue
        if ($lightMode -and $lightMode.AppsUseLightTheme -eq 1) {
            return "Light"
        } else {
            return "Dark"
        }
    } catch {
        Write-Warning "Could not determine current theme: $($_.Exception.Message)"
        return "Dark"
    }
}

# Function to set Windows theme
function Set-WindowsTheme {
    param([string]$ThemeMode)
    
    try {
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        switch ($ThemeMode) {
            "Light" {
                Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 1
                Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 1
                Write-Host "Theme set to Light mode"
            }
            "Dark" {
                Set-ItemProperty -Path $regPath -Name "AppsUseLightTheme" -Value 0
                Set-ItemProperty -Path $regPath -Name "SystemUsesLightTheme" -Value 0
                Write-Host "Theme set to Dark mode"
            }
        }
        
        return $true
    } catch {
        Write-Error "Failed to set theme: $($_.Exception.Message)"
        return $false
    }
}

# Function to set accent color
function Set-AccentColor {
    param([string]$ColorName)
    
    $colorMap = @{
        "Red" = 0xFF0000
        "Orange" = 0xFF8000
        "Yellow" = 0xFFFF00
        "Green" = 0x00FF00
        "Cyan" = 0x00FFFF
        "Blue" = 0x0078D4
        "Purple" = 0x881798
        "Pink" = 0xFF69B4
        "Default" = 0x0078D4
    }
    
    if (-not $colorMap.ContainsKey($ColorName)) {
        Write-Warning "Unknown color: $ColorName. Using default blue."
        $ColorName = "Default"
    }
    
    try {
        $colorValue = $colorMap[$ColorName]
        $regPath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Accent"
        
        if (-not (Test-Path $regPath)) {
            New-Item -Path $regPath -Force | Out-Null
        }
        
        Set-ItemProperty -Path $regPath -Name "AccentColor" -Value $colorValue
        Write-Host "Accent color set to $ColorName"
        return $true
    } catch {
        Write-Error "Failed to set accent color: $($_.Exception.Message)"
        return $false
    }
}

# Function to set wallpaper
function Set-Wallpaper {
    param([string]$WallpaperPath)
    
    if (-not (Test-Path $WallpaperPath)) {
        Write-Warning "Wallpaper file not found: $WallpaperPath"
        return $false
    }
    
    try {
        $regPath = "HKCU:\Control Panel\Desktop"
        Set-ItemProperty -Path $regPath -Name "Wallpaper" -Value $WallpaperPath
        Write-Host "Wallpaper set to: $WallpaperPath"
        return $true
    } catch {
        Write-Error "Failed to set wallpaper: $($_.Exception.Message)"
        return $false
    }
}

# Function to close File Explorer windows
function Close-ExplorerWindows {
    try {
        $shell = New-Object -ComObject Shell.Application
        $windows = $shell.Windows()
        
        for ($i = $windows.Count - 1; $i -ge 0; $i--) {
            $window = $windows.Item($i)
            if ($window.Name -eq "File Explorer" -or $window.Name -eq "Windows Explorer") {
                $window.Quit()
            }
        }
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shell) | Out-Null
        Write-Host "File Explorer windows closed"
        return $true
    } catch {
        Write-Warning "Could not close Explorer windows: $($_.Exception.Message)"
        return $false
    }
}

# Function to restart Explorer
function Restart-Explorer {
    try {
        Write-Host "Closing File Explorer windows..."
        Close-ExplorerWindows
        Start-Sleep -Seconds 1
        
        Write-Host "Restarting Windows Explorer process..."
        Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process "explorer.exe"
        Start-Sleep -Seconds 3
        
        Write-Host "Explorer restarted successfully"
        return $true
    } catch {
        Write-Error "Failed to restart Explorer: $($_.Exception.Message)"
        return $false
    }
}

# Function to apply theme changes using Windows API
function Apply-ThemeChanges {
    try {
        Write-Host "Applying theme changes..."
        
        # Notify system of setting changes
        $result = [IntPtr]::Zero
        [WinAPI]::SendMessageTimeout([WinAPI]::HWND_BROADCAST, [WinAPI]::WM_SETTINGCHANGE, [IntPtr]::Zero, [IntPtr]::Zero, [WinAPI]::SMTO_ABORTIFHUNG, 5000, [ref]$result)
        
        # Notify shell of association changes
        [WinAPI]::SHChangeNotify([WinAPI]::SHCNE_ASSOCCHANGED, [WinAPI]::SHCNF_IDLIST, [IntPtr]::Zero, [IntPtr]::Zero)
        
        Start-Sleep -Seconds 1
        
        # If theme changes are not visible, restart Explorer as fallback
        Write-Host "Checking if theme change is applied..."
        Start-Sleep -Seconds 2
        
        # For reliability, we'll do a controlled Explorer restart
        if (Restart-Explorer) {
            Write-Host "Theme changes applied successfully"
            return $true
        } else {
            Write-Warning "Theme changes may not be fully applied"
            return $false
        }
    } catch {
        Write-Error "Failed to apply theme changes: $($_.Exception.Message)"
        return $false
    }
}

# Function to create scheduled tasks
function Setup-ScheduledTasks {
    param($Config)
    
    try {
        # Remove existing tasks first
        Remove-ScheduledTasks
        
        foreach ($schedule in $Config.schedules) {
            $taskName = "WindowsThemeSwitcher_$($schedule.name -replace '\s+', '_')"
            $scriptPath = $PSCommandPath
            
            # Build arguments for the scheduled task
            $arguments = "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`" -Mode $($schedule.mode)"
            
            if ($schedule.accentColor) {
                $arguments += " -AccentColor $($schedule.accentColor)"
            }
            
            if ($schedule.wallpaper -and $schedule.wallpaper.Trim() -ne "") {
                $arguments += " -Wallpaper `"$($schedule.wallpaper)`""
            }
            
            # Create the scheduled task
            $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument $arguments
            $trigger = New-ScheduledTaskTrigger -Daily -At $schedule.time
            $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
            $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive
            
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Description "Automatically switch Windows theme to $($schedule.mode) mode at $($schedule.time)" -Force
            
            Write-Host "Created scheduled task: $taskName"
        }
        
        Write-Host "All scheduled tasks created successfully"
        return $true
    } catch {
        Write-Error "Failed to create scheduled tasks: $($_.Exception.Message)"
        return $false
    }
}

# Function to remove scheduled tasks
function Remove-ScheduledTasks {
    try {
        $tasks = Get-ScheduledTask | Where-Object { $_.TaskName -like "WindowsThemeSwitcher_*" }
        
        foreach ($task in $tasks) {
            Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false
            Write-Host "Removed scheduled task: $($task.TaskName)"
        }
        
        if ($tasks.Count -eq 0) {
            Write-Host "No WindowsThemeSwitcher scheduled tasks found"
        } else {
            Write-Host "Removed $($tasks.Count) scheduled task(s)"
        }
        
        return $true
    } catch {
        Write-Error "Failed to remove scheduled tasks: $($_.Exception.Message)"
        return $false
    }
}

# Main execution logic
try {
    if ($RemoveSchedule) {
        Remove-ScheduledTasks
        exit 0
    }
    
    if ($SetupSchedule) {
        $config = Load-Config
        Setup-ScheduledTasks -Config $config
        exit 0
    }
    
    # Handle theme mode switching
    if ($Mode) {
        $currentTheme = Get-CurrentTheme
        
        if ($Mode -eq "Toggle") {
            $Mode = if ($currentTheme -eq "Light") { "Dark" } else { "Light" }
            Write-Host "Toggling from $currentTheme to $Mode"
        }
        
        if (Set-WindowsTheme -ThemeMode $Mode) {
            $themeChanged = $true
        }
    }
    
    # Handle accent color
    if ($AccentColor) {
        if (Set-AccentColor -ColorName $AccentColor) {
            $themeChanged = $true
        }
    }
    
    # Handle wallpaper
    if ($Wallpaper) {
        if (Set-Wallpaper -WallpaperPath $Wallpaper) {
            $themeChanged = $true
        }
    }
    
    # Apply changes if any were made
    if ($themeChanged) {
        Apply-ThemeChanges
    }
    
    if (-not $Mode -and -not $AccentColor -and -not $Wallpaper) {
        Write-Host "Windows Theme Switcher"
        Write-Host "Current theme: $(Get-CurrentTheme)"
        Write-Host ""
        Write-Host "Usage examples:"
        Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Toggle"
        Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Dark -AccentColor Purple"
        Write-Host "  .\Switch-WindowsTheme.ps1 -SetupSchedule"
        Write-Host "  .\Switch-WindowsTheme.ps1 -RemoveSchedule"
    }
    
} catch {
    Write-Error "Script execution failed: $($_.Exception.Message)"
    exit 1
}
