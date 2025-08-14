# Windows Theme Switcher

A PowerShell script to automatically switch between Windows Light/Dark themes, change accent colors, and set wallpapers on a schedule.


## Requirements

- Windows 10/11
- PowerShell 7 or later ([Installing PowerShell ](https://learn.microsoft.com/en-us/powershell/scripting/install/installing-powershell-on-windows))
- Administrator privileges for scheduled tasks


## Quick Start

```powershell
# Toggle between light and dark mode
.\Switch-WindowsTheme.ps1 -Mode Toggle

# Set dark mode with purple accent
.\Switch-WindowsTheme.ps1 -Mode Dark -AccentColor Purple

# Force restart explorer 
.\Switch-WindowsTheme.ps1 -Mode Dark -RestartExplorer

# Setup automatic theme switching
.\Switch-WindowsTheme.ps1 -SetupSchedule
```

## Features

- **Theme Switching**: Light/Dark mode toggle
- **Accent Colors**: 9 built-in colors (Red, Orange, Yellow, Green, Cyan, Blue, Purple, Pink, Default)
- **Wallpaper Management**: Automatic wallpaper switching with themes
- **Scheduled Tasks**: Automatic theme switching at specified times 
- **Fast Refresh**: API-based theme switching (no Explorer restart by default)

## Parameters

| Parameter | Description |
|-----------|-------------|
| `-Mode` | Light, Dark, or Toggle |
| `-AccentColor` | Color name from available list |
| `-Wallpaper` | Path to wallpaper image |
| `-SetupSchedule` | Create scheduled tasks |
| `-RemoveSchedule` | Remove all scheduled tasks |
| `-RestartExplorer` | Force Explorer restart for complete refresh |
| `-ConfigFile` | Custom config file path |

## Configuration

Edit `Switch-WindowsTheme.json` to customize:

```json
{
  "DefaultDayTheme": "Light",
  "DefaultNightTheme": "Dark",
  "DefaultDayAccent": "Blue",
  "DefaultNightAccent": "Blue",
  "DefaultLightWallpaper": "C:\\path\\to\\light-bg.jpg",
  "DefaultDarkWallpaper": "C:\\path\\to\\dark-bg.jpg",
  "Schedules": [
    {
      "Name": "Day",
      "Time": "07:00",
      "Theme": "Light",
      "AccentColor": "Blue",
      "Wallpaper": "C:\\path\\to\\light-bg.jpg",
      "Enabled": true
    },
    {
      "Name": "Night",
      "Time": "20:00",
      "Theme": "Dark",
      "AccentColor": "Blue",
      "Wallpaper": "C:\\path\\to\\dark-bg.jpg",
      "Enabled": true
    }
  ]
}

```

## Examples

```powershell
# Basic usage
.\Switch-WindowsTheme.ps1 -Mode Dark
.\Switch-WindowsTheme.ps1 -AccentColor Purple
.\Switch-WindowsTheme.ps1 -Wallpaper "C:\bg.jpg"

# With complete refresh (if taskbar doesn't update)
.\Switch-WindowsTheme.ps1 -Mode Toggle -RestartExplorer

# Schedule management
.\Switch-WindowsTheme.ps1 -SetupSchedule    # Create tasks
.\Switch-WindowsTheme.ps1 -RemoveSchedule   # Remove tasks
```

## Notes & Troubleshooting

- ❌ Taskbar doesn't refresh properly 👉 Use `-RestartExplorer` 

- ❌ Errors like `statement is missing catch` or `Missing closing '}'` 👉 Update powershell 

- Scheduled tasks run with hidden windows

## TODO

Clean up unnecessary code. 