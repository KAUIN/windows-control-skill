---
name: windows-control-skill
description: Control Windows desktop remotely via PowerShell for GUI automation. Use for screenshots, mouse/keyboard control, window management, and application automation from WSL/Linux. Requires PowerShell on Windows with execution policy set to Bypass.
---

# Windows Control Skill (windows-control-skill)

Control Windows desktop remotely through PowerShell for full GUI automation.

## Quick Start

**Note:** The script file is `.txt` format for ClawHub compatibility. 
**Before use, rename it to `.ps1`:**
```bash
# Copy script to Windows and rename to .ps1
cp ~/.openclaw/skills/win-control/scripts/win_control.txt /mnt/c/temp/
ren C:\temp\win_control.txt win_control.ps1

# Screenshot
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.ps1 screenshot C:\temp\screen.png

# Click at position (x, y)
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.txt click 500 300

# Type text
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.txt type "Hello World"

# Press key
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.txt key enter
```

## Commands

### Screenshot
```bash
powershell.exe -File win_control.txt screenshot C:\temp\screen.png
```

### Mouse Control
```bash
# Click (left/right/middle, single/double)
powershell.exe -File win_control.txt click 500 300
powershell.exe -File win_control.txt click 500 300 right
powershell.exe -File win_control.txt click 500 300 left double

# Move cursor
powershell.exe -File win_control.txt move 500 300

# Drag
powershell.exe -File win_control.txt drag 100 100 500 500

# Scroll (up/down)
powershell.exe -File win_control.txt scroll down 5
```

### Keyboard Control
```bash
# Type text
powershell.exe -File win_control.txt type "Hello World"

# Press special key
powershell.exe -File win_control.txt key enter
powershell.exe -File win_control.txt key escape
powershell.exe -File win_control.txt key tab

# Key combinations
powershell.exe -File win_control.txt combo "ctrl+c"
powershell.exe -File win_control.txt combo "alt+tab"
```

### Window Management
```bash
# List windows
powershell.exe -File win_control.txt list

# Activate window
powershell.exe -File win_control.txt activate "Notepad"

# Close window
powershell.exe -File win_control.txt close "Calculator"

# Start application
powershell.exe -File win_control.txt start notepad
```

### System Info
```bash
# Screen resolution
powershell.exe -File win_control.txt info

# Cursor position
powershell.exe -File win_control.txt getcursor
```

## Real-time Screen Capture

```bash
# Continuous screenshot stream (5 FPS for 30 seconds)
python3 ~/.openclaw/skills/win-control/scripts/realtime_screen.py --fps 5 --duration 30

# Get latest frame
python3 -c "
from realtime_screen import WindowsScreenCapture
cap = WindowsScreenCapture()
print(cap.get_latest_frame())
"
```

## Prerequisites

### Windows (Target Machine)
- PowerShell (built-in on Windows 7+)
- Set execution policy: `Set-ExecutionPolicy Bypass -Scope Process`
- Required for: screenshot, mouse/keyboard control, window management

### WSL/Linux (Control Machine)
- WSL with access to Windows drives (`/mnt/c/`)
- Or native Windows with Python
- Python 3.7+ (for realtime_screen.py)
- Pillow library: `pip install Pillow`

### Setup Steps
1. Copy `scripts/win_control.txt` to Windows (e.g., `C:\temp\`)
2. Set PowerShell execution policy (one-time):
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process
   ```
3. Test connection:
   ```bash
   powershell.exe -File C:\temp\win_control.txt info
   ```

## Available Keys

- **Navigation**: up, down, left, right, home, end, pageup, pagedown
- **Editing**: enter, return, tab, backspace, delete, escape
- **Modifiers**: Use with combo - ctrl, alt, shift

## Troubleshooting

### Execution Policy Error
```powershell
Set-ExecutionPolicy Bypass -Scope Process
```

### Path Issues
Use full paths or copy script to `C:\temp\`

## References

- [PowerShell Advanced Usage](references/powershell.md)
