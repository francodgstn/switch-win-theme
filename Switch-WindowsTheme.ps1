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
    [switch]$RestartExplorer,
    
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\Switch-WindowsTheme.json"
)

# Default configuration
$DefaultConfig = @{
    "DefaultDayTheme" = "Light"
    "DefaultNightTheme" = "Dark"
    "DefaultDayAccent" = "Blue"
    "DefaultNightAccent" = "Blue"
    "DefaultLightWallpaper" = ""
    "DefaultDarkWallpaper" = ""
    "Schedules" = @(
        @{
            "Name" = "Morning"
            "Time" = "07:00"
            "Theme" = "Light"
            "AccentColor" = "Blue"
            "Wallpaper" = ""
            "Enabled" = $true
        },
        @{
            "Name" = "Evening"
            "Time" = "19:00"
            "Theme" = "Dark"
            "AccentColor" = "Blue"
            "Wallpaper" = ""
            "Enabled" = $true
        }
    )
}

# Registry paths for theme settings
$PersonalizePath = "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize"
$AccentPath = "HKCU:\SOFTWARE\Microsoft\Windows\DWM"

# Accent color values (ABGR format)
$AccentColors = @{
    "Red"     = 0xFF0000FF
    "Orange"  = 0xFF0080FF
    "Yellow"  = 0xFF00FFFF
    "Green"   = 0xFF00FF00
    "Cyan"    = 0xFFFFFF00
    "Blue"    = 0xFFFF0000
    "Purple"  = 0xFFFF00FF
    "Pink"    = 0xFF8080FF
    "Default" = 0xFFD77800
}

function Load-Config {
    param([string]$ConfigPath)
    
    if (Test-Path $ConfigPath) {
        try {
            $configContent = Get-Content $ConfigPath -Raw | ConvertFrom-Json
            Write-Host "✓ Configuration loaded from $ConfigPath" -ForegroundColor Green
            return $configContent
        }
        catch {
            Write-Warning "Failed to load config file. Using default configuration."
            return $DefaultConfig
        }
    }
    else {
        Write-Host "Config file not found. Using built-in default configuration." -ForegroundColor Yellow
        Write-Host "To customize, copy 'Switch-WindowsTheme.json.sample' to 'Switch-WindowsTheme.json' and edit it." -ForegroundColor Yellow
        return $DefaultConfig
    }
}

function Save-Config {
    param($Config, [string]$ConfigPath)
    
    try {
        $Config | ConvertTo-Json -Depth 10 | Set-Content $ConfigPath
        Write-Host "✓ Configuration saved to $ConfigPath" -ForegroundColor Green
    }
    catch {
        Write-Error "Failed to save configuration: $($_.Exception.Message)"
    }
}

function Get-CurrentTheme {
    try {
        $appsTheme = Get-ItemProperty -Path $PersonalizePath -Name "AppsUseLightTheme" -ErrorAction Stop
        $systemTheme = Get-ItemProperty -Path $PersonalizePath -Name "SystemUsesLightTheme" -ErrorAction Stop
        
        if ($appsTheme.AppsUseLightTheme -eq 1) {
            return "Light"
        } else {
            return "Dark"
        }
    }
    catch {
        Write-Warning "Could not read current theme. Registry keys may not exist."
        return "Unknown"
    }
}

function Set-WindowsTheme {
    param([string]$ThemeMode)
    
    try {
        if ($ThemeMode -eq "Light") {
            Set-ItemProperty -Path $PersonalizePath -Name "AppsUseLightTheme" -Value 1
            Set-ItemProperty -Path $PersonalizePath -Name "SystemUsesLightTheme" -Value 1
            Write-Host "✓ Theme changed to Light mode" -ForegroundColor Green
        }
        elseif ($ThemeMode -eq "Dark") {
            Set-ItemProperty -Path $PersonalizePath -Name "AppsUseLightTheme" -Value 0
            Set-ItemProperty -Path $PersonalizePath -Name "SystemUsesLightTheme" -Value 0
            Write-Host "✓ Theme changed to Dark mode" -ForegroundColor Green
        }
    }
    catch {
        Write-Error "Failed to change theme: $($_.Exception.Message)"
    }
}

function Set-Wallpaper {
    param([string]$ImagePath)
    
    if (-not $ImagePath) {
        Write-Host "No wallpaper path specified, skipping wallpaper change" -ForegroundColor Gray
        return
    }
    
    if (-not (Test-Path $ImagePath)) {
        Write-Error "Wallpaper file not found: $ImagePath"
        return
    }
    
    try {
        # Get the absolute path
        $fullPath = (Resolve-Path $ImagePath).Path
        
        # Set wallpaper using SystemParametersInfo
        $signature = @'
[DllImport("user32.dll", CharSet = CharSet.Auto)]
public static extern int SystemParametersInfo(int uAction, int uParam, string lpvParam, int fuWinIni);
'@
        
        $setWallpaper = Add-Type -MemberDefinition $signature -Name "Win32SetWallpaper" -Namespace Win32Functions -PassThru
        
        # SPI_SETDESKWALLPAPER = 0x0014, SPIF_UPDATEINIFILE | SPIF_SENDCHANGE = 0x03
        $result = $setWallpaper::SystemParametersInfo(0x0014, 0, $fullPath, 0x03)
        
        if ($result -ne 0) {
            Write-Host "✓ Wallpaper changed to: $ImagePath" -ForegroundColor Green
        } else {
            Write-Warning "Failed to set wallpaper. The image may not be supported or accessible."
        }
    }
    catch {
        Write-Error "Failed to set wallpaper: $($_.Exception.Message)"
    }
}

function Get-DefaultWallpaper {
    param([string]$Theme, $Config)
    
    if ($Theme -eq "Light" -and $Config.DefaultLightWallpaper) {
        return $Config.DefaultLightWallpaper
    }
    elseif ($Theme -eq "Dark" -and $Config.DefaultDarkWallpaper) {
        return $Config.DefaultDarkWallpaper
    }
    
    return $null
}

function Set-AccentColor {
    param([string]$ColorName)
    
    if ($AccentColors.ContainsKey($ColorName)) {
        try {
            $colorValue = $AccentColors[$ColorName]
            
            Set-ItemProperty -Path $AccentPath -Name "AccentColor" -Value $colorValue
            Set-ItemProperty -Path $AccentPath -Name "ColorizationColor" -Value $colorValue
            Set-ItemProperty -Path $PersonalizePath -Name "ColorPrevalence" -Value 0
            
            Write-Host "✓ Accent color changed to $ColorName" -ForegroundColor Green
        }
        catch {
            Write-Error "Failed to set accent color: $($_.Exception.Message)"
        }
    }
    else {
        Write-Error "Invalid color name. Available colors: $($AccentColors.Keys -join ', ')"
    }
}

function Setup-ScheduledTasks {
    param($Config, [string]$ScriptPath)
    
    Write-Host "`nSetting up scheduled tasks..." -ForegroundColor Cyan
    
    # Remove existing tasks first
    Remove-ScheduledTasks
    
    foreach ($schedule in $Config.Schedules) {
        if ($schedule.Enabled) {
            $taskName = "WindowsThemeSwitcher_$($schedule.Name)"
            $time = [DateTime]::ParseExact($schedule.Time, "HH:mm", $null)
            
            try {
                # Create the action
                $arguments = "-Mode $($schedule.Theme)"
                if ($schedule.AccentColor) {
                    $arguments += " -AccentColor $($schedule.AccentColor)"
                }
                if ($schedule.Wallpaper) {
                    $arguments += " -Wallpaper `"$($schedule.Wallpaper)`""
                }
                if ($ConfigFile -ne ".\Switch-WindowsTheme.json") {
                    $arguments += " -ConfigFile `"$ConfigFile`""
                }
                
                $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-WindowStyle Hidden -ExecutionPolicy Bypass -File `"$ScriptPath`" $arguments"
                
                # Create the trigger
                $trigger = New-ScheduledTaskTrigger -Daily -At $time
                
                # Create task settings
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable
                
                # Register the task
                Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger -Settings $settings -Description "Automatically switch Windows theme to $($schedule.Theme) at $($schedule.Time)" -Force | Out-Null
                
                Write-Host "✓ Created task: $taskName ($($schedule.Time) - $($schedule.Theme))" -ForegroundColor Green
            }
            catch {
                Write-Error "Failed to create scheduled task '$taskName': $($_.Exception.Message)"
            }
        }
    }
    
    Write-Host "`n✓ Scheduled tasks setup complete!" -ForegroundColor Green
    Write-Host "Tasks will run daily at the specified times." -ForegroundColor Yellow
}

function Remove-ScheduledTasks {
    Write-Host "Removing existing WindowsThemeSwitcher scheduled tasks..." -ForegroundColor Yellow
    
    try {
        $existingTasks = Get-ScheduledTask -TaskName "WindowsThemeSwitcher_*" -ErrorAction SilentlyContinue
        foreach ($task in $existingTasks) {
            Unregister-ScheduledTask -TaskName $task.TaskName -Confirm:$false
            Write-Host "✓ Removed task: $($task.TaskName)" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "No existing WindowsThemeSwitcher tasks found to remove." -ForegroundColor Gray
    }
}

function Show-CurrentSchedule {
    param($Config)
    
    Write-Host "`nCurrent Schedule Configuration:" -ForegroundColor Cyan
    Write-Host "================================" -ForegroundColor Cyan
    
    foreach ($schedule in $Config.Schedules) {
        $status = if ($schedule.Enabled) { "✓" } else { "✗" }
        $accent = if ($schedule.AccentColor) { " ($($schedule.AccentColor))" } else { "" }
        $wallpaper = if ($schedule.Wallpaper) { " [Wallpaper: $(Split-Path $schedule.Wallpaper -Leaf)]" } else { "" }
        Write-Host "$status $($schedule.Name): $($schedule.Time) - $($schedule.Theme)$accent$wallpaper"
    }
    
    Write-Host "`nScheduled Tasks Status:" -ForegroundColor Cyan
    try {
        $tasks = Get-ScheduledTask -TaskName "WindowsThemeSwitcher_*" -ErrorAction SilentlyContinue
        if ($tasks) {
            foreach ($task in $tasks) {
                $state = $task.State
                $nextRun = (Get-ScheduledTask -TaskName $task.TaskName | Get-ScheduledTaskInfo).NextRunTime
                Write-Host "✓ $($task.TaskName): $state (Next: $nextRun)"
            }
        } else {
            Write-Host "No scheduled tasks found." -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "Could not retrieve scheduled task information." -ForegroundColor Yellow
    }
}

function Show-Usage {
    Write-Host "`nWindows Color Mode and Accent Color Switcher with Scheduling" -ForegroundColor Cyan
    Write-Host "=============================================================" -ForegroundColor Cyan
    Write-Host "`nUsage:"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Mode `<Light|Dark|Toggle`>"
    Write-Host "  .\Switch-WindowsTheme.ps1 -AccentColor `<ColorName`>"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Wallpaper `<ImagePath`>"
    Write-Host "  .\Switch-WindowsTheme.ps1 -SetupSchedule"
    Write-Host "  .\Switch-WindowsTheme.ps1 -RemoveSchedule"
    Write-Host "  .\Switch-WindowsTheme.ps1 -RestartExplorer"
    Write-Host "  .\Switch-WindowsTheme.ps1 -ConfigFile `<Path`>"
    Write-Host "`nOptions:"
    Write-Host "  -RestartExplorer     Force Explorer restart for complete theme refresh"
    Write-Host "`nScheduling:"
    Write-Host "  -SetupSchedule    Create scheduled tasks based on config"
    Write-Host "  -RemoveSchedule   Remove all WindowsThemeSwitcher scheduled tasks"
    Write-Host "`nAvailable Colors:"
    Write-Host "  $($AccentColors.Keys -join ', ')"
    Write-Host "`nExamples:"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Toggle"
    Write-Host "  .\Switch-WindowsTheme.ps1 -SetupSchedule"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Dark -AccentColor Purple"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Light -RestartExplorer"
    Write-Host "  .\Switch-WindowsTheme.ps1 -Wallpaper `"C:\Images\MyWallpaper.jpg`""
    Write-Host "  .\Switch-WindowsTheme.ps1 -Mode Light -Wallpaper `".\light-bg.png`""
    Write-Host "  .\Switch-WindowsTheme.ps1 -ConfigFile `"C:\MyConfig.json`""
    Write-Host "`nNote:"
    Write-Host "  By default, uses API calls for fast theme switching."
    Write-Host "  Use -RestartExplorer for complete refresh if taskbar doesn't update."
    Write-Host "  Supported wallpaper formats: JPG, PNG, BMP"
    Write-Host "`nConfiguration File:"
    Write-Host "  Edit Switch-WindowsTheme.json to customize schedules and defaults."
    Write-Host "  The file will be created automatically with default settings."
}

function Apply-ThemeChanges {
    param([bool]$ForceExplorerRestart = $false)
    
    Write-Host "Applying theme changes..." -ForegroundColor Yellow
    
    if ($ForceExplorerRestart) {
        Write-Host "Using Explorer restart method for complete refresh..." -ForegroundColor Cyan
        
        try {
            # Kill all Explorer processes
            Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            
            # Wait for processes to fully close
            Start-Sleep -Seconds 2
            
            # Start Explorer again
            Start-Process -FilePath "explorer.exe"
            
            # Wait for Explorer to fully start
            Start-Sleep -Seconds 3
            
            # Close any file explorer windows that may have opened, but keep the shell
            try {
                Add-Type -AssemblyName Microsoft.VisualBasic
                $shellApp = New-Object -ComObject Shell.Application
                $windows = $shellApp.Windows()
                
                foreach ($window in $windows) {
                    # Check if it's a file explorer window (not desktop)
                    if ($window.Name -eq "Windows Explorer" -or $window.Name -eq "File Explorer") {
                        try {
                            $window.Quit()
                            Write-Host "✓ Closed file explorer window" -ForegroundColor Green
                        }
                        catch {
                            # Ignore errors when closing windows
                        }
                    }
                }
                
                # Release COM object
                [System.Runtime.Interopservices.Marshal]::ReleaseComObject($shellApp) | Out-Null
            }
            catch {
                # If COM approach fails, try alternative method
                Get-Process -Name "explorer" -ErrorAction SilentlyContinue | Where-Object { 
                    $_.MainWindowTitle -ne "" -and $_.ProcessName -eq "explorer" 
                } | ForEach-Object {
                    try {
                        $_.CloseMainWindow()
                        Write-Host "✓ Closed explorer window: $($_.MainWindowTitle)" -ForegroundColor Green
                    }
                    catch {
                        # Ignore errors
                    }
                }
            }
            
            Write-Host "✓ Theme changes applied with Explorer restart" -ForegroundColor Green
        }
        catch {
            Write-Warning "Failed to restart Explorer. Please manually restart Explorer or log out/in to see all changes."
        }
    }
    else {
        try {
            # Define comprehensive Windows API signatures
            $signature = @'
[DllImport("user32.dll", SetLastError = true)]
public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);

[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

[DllImport("user32.dll", SetLastError = true)]
public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

[DllImport("shell32.dll", SetLastError = true)]
public static extern void SHChangeNotify(int wEventId, int uFlags, IntPtr dwItem1, IntPtr dwItem2);
'@
            
            $winAPI = Add-Type -MemberDefinition $signature -Name "Win32ThemeAPI" -Namespace Win32Functions -PassThru
            
            # Constants
            $HWND_BROADCAST = [IntPtr]0xFFFF
            $WM_SETTINGCHANGE = 0x001A
            $WM_THEMECHANGED = 0x031A
            $SPIF_UPDATEINIFILE = 0x01
            $SPIF_SENDCHANGE = 0x02
            $SPI_SETDESKWALLPAPER = 0x0014
            $SHCNE_ASSOCCHANGED = 0x08000000
            $SHCNF_IDLIST = 0x0000
            
            Write-Host "Refreshing system settings..." -ForegroundColor Cyan
            
            # 1. Refresh desktop wallpaper
            $winAPI::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, [IntPtr]::Zero, $SPIF_UPDATEINIFILE -bor $SPIF_SENDCHANGE) | Out-Null
            
            # 2. Notify all windows of theme change
            $winAPI::PostMessage($HWND_BROADCAST, $WM_THEMECHANGED, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            
            # 3. Notify of setting changes (for personalization)
            $personalizePtr = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni("ImmersiveColorSet")
            $winAPI::PostMessage($HWND_BROADCAST, $WM_SETTINGCHANGE, [IntPtr]::Zero, $personalizePtr) | Out-Null
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($personalizePtr)
            
            # 4. Notify shell of changes
            $winAPI::SHChangeNotify($SHCNE_ASSOCCHANGED, $SHCNF_IDLIST, [IntPtr]::Zero, [IntPtr]::Zero)
            
            # 5. Additional setting change notifications
            $settingsToNotify = @("WindowsThemeElement", "Environment", "Policy")
            foreach ($setting in $settingsToNotify) {
                $settingPtr = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni($setting)
                $winAPI::PostMessage($HWND_BROADCAST, $WM_SETTINGCHANGE, [IntPtr]::Zero, $settingPtr) | Out-Null
                [System.Runtime.InteropServices.Marshal]::FreeHGlobal($settingPtr)
            }
            
            # 6. Find and refresh taskbar/explorer windows specifically
            $taskbarHwnd = $winAPI::FindWindow("Shell_TrayWnd", $null)
            if ($taskbarHwnd -ne [IntPtr]::Zero) {
                $winAPI::SendMessage($taskbarHwnd, $WM_THEMECHANGED, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
                $winAPI::PostMessage($taskbarHwnd, $WM_SETTINGCHANGE, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            }
            
            # 7. Additional taskbar refresh attempts
            $explorerHwnd = $winAPI::FindWindow("Progman", "Program Manager")
            if ($explorerHwnd -ne [IntPtr]::Zero) {
                $winAPI::SendMessage($explorerHwnd, $WM_THEMECHANGED, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null
            }
            
            # 8. Try to refresh system tray area
            $trayHwnd = $winAPI::FindWindow("Shell_TrayWnd", $null)
            if ($trayHwnd -ne [IntPtr]::Zero) {
                # Send invalidate messages to force redraw
                $winAPI::SendMessage($trayHwnd, 0x000F, [IntPtr]::Zero, [IntPtr]::Zero) | Out-Null  # WM_PAINT
            }
            
            # 9. Force desktop refresh
            $winAPI::SystemParametersInfo(0x0073, 0, [IntPtr]::Zero, $SPIF_SENDCHANGE) | Out-Null  # SPI_SETWORKAREA
            
            # 10. Additional Windows 10/11 specific notifications
            $immersivePtr = [System.Runtime.InteropServices.Marshal]::StringToHGlobalUni("intl")
            $winAPI::PostMessage($HWND_BROADCAST, $WM_SETTINGCHANGE, [IntPtr]::Zero, $immersivePtr) | Out-Null
            [System.Runtime.InteropServices.Marshal]::FreeHGlobal($immersivePtr)
            
            Write-Host "✓ Theme changes applied using API calls" -ForegroundColor Green
            Write-Host "Note: If taskbar doesn't refresh, try running with -RestartExplorer parameter" -ForegroundColor Yellow
        }
        catch {
            Write-Warning "Failed to fully apply theme changes using API calls. Try using -RestartExplorer parameter for complete refresh."
        }
    }
}

# Main execution
Write-Host "Windows Color Switcher with Scheduling" -ForegroundColor Cyan
Write-Host "=======================================" -ForegroundColor Cyan

# Load configuration
$config = Load-Config -ConfigPath $ConfigFile

# Handle scheduled task operations
if ($SetupSchedule) {
    $scriptPath = $MyInvocation.MyCommand.Path
    if (-not $scriptPath) {
        Write-Error "Cannot determine script path. Please run the script from a file."
        exit 1
    }
    Setup-ScheduledTasks -Config $config -ScriptPath $scriptPath
    Show-CurrentSchedule -Config $config
    exit 0
}

if ($RemoveSchedule) {
    Remove-ScheduledTasks
    exit 0
}

# Get current theme
$currentTheme = Get-CurrentTheme
Write-Host "Current theme: $currentTheme"

# Handle Mode parameter
if ($Mode) {
    if ($Mode -eq "Toggle") {
        $newMode = if ($currentTheme -eq "Light") { "Dark" } else { "Light" }
        Write-Host "Toggling from $currentTheme to $newMode"
        Set-WindowsTheme -ThemeMode $newMode
        
        # Check for default wallpaper for the new mode
        $defaultWallpaper = Get-DefaultWallpaper -Theme $newMode -Config $config
        if ($defaultWallpaper -and (Test-Path $defaultWallpaper)) {
            Set-Wallpaper -ImagePath $defaultWallpaper
        }
    }
    else {
        Set-WindowsTheme -ThemeMode $Mode
        
        # Check for default wallpaper for the specified mode
        $defaultWallpaper = Get-DefaultWallpaper -Theme $Mode -Config $config
        if ($defaultWallpaper -and (Test-Path $defaultWallpaper)) {
            Set-Wallpaper -ImagePath $defaultWallpaper
        }
    }
}

# Handle AccentColor parameter
if ($AccentColor) {
    Set-AccentColor -ColorName $AccentColor
}

# Handle Wallpaper parameter
if ($Wallpaper) {
    Set-Wallpaper -ImagePath $Wallpaper
}

# Show usage if no parameters provided
if (-not $Mode -and -not $AccentColor -and -not $Wallpaper -and -not $SetupSchedule -and -not $RemoveSchedule) {
    Show-Usage
    Show-CurrentSchedule -Config $config
}
else {
    # Apply theme changes if changes were made
    if ($Mode -or $AccentColor -or $Wallpaper) {
        Apply-ThemeChanges -ForceExplorerRestart:$RestartExplorer
    }
}