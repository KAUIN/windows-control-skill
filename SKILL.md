---
name: windows-control-skill
description: Control Windows desktop remotely via PowerShell for GUI automation. Use for screenshots, mouse/keyboard control, window management, and application automation from WSL/Linux. Requires PowerShell on Windows with execution policy set to Bypass.
---

# Windows Control Skill (windows-control-skill)

Control Windows desktop remotely through PowerShell for full GUI automation.

## File Versions

This skill includes two versions of the PowerShell script:

- **`win_control.ps1`** - PowerShell script (for direct use)
- **`win_control.txt`** - Text version (for ClawHub compatibility)

**AI Instructions:** When using this skill on Windows, copy the `.txt` file and rename it to `.ps1` before execution.

## Quick Start

```bash
# Step 1: Copy script to Windows (use .txt version from skill)
cp ~/.openclaw/skills/win-control/scripts/win_control.txt /mnt/c/temp/

# Step 2: Rename to .ps1 (REQUIRED for PowerShell execution)
ren C:\temp\win_control.txt win_control.ps1

# Step 3: Execute commands
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.ps1 screenshot C:\temp\screen.png
```

## Commands

### Screenshot
```bash
powershell.exe -File win_control.ps1 screenshot C:\temp\screen.png
```

### Mouse Control
```bash
# Click (left/right/middle, single/double)
powershell.exe -File win_control.ps1 click 500 300
powershell.exe -File win_control.ps1 click 500 300 right
powershell.exe -File win_control.ps1 click 500 300 left double

# Move cursor
powershell.exe -File win_control.ps1 move 500 300

# Drag
powershell.exe -File win_control.ps1 drag 100 100 500 500

# Scroll (up/down)
powershell.exe -File win_control.ps1 scroll down 5
```

### Keyboard Control
```bash
# Type text
powershell.exe -File win_control.ps1 type "Hello World"

# Press special key
powershell.exe -File win_control.ps1 key enter
powershell.exe -File win_control.ps1 key escape
powershell.exe -File win_control.ps1 key tab

# Key combinations
powershell.exe -File win_control.ps1 combo "ctrl+c"
powershell.exe -File win_control.ps1 combo "alt+tab"
```

### Window Management
```bash
# List windows
powershell.exe -File win_control.ps1 list

# Activate window
powershell.exe -File win_control.ps1 activate "Notepad"

# Close window
powershell.exe -File win_control.ps1 close "Calculator"

# Start application
powershell.exe -File win_control.ps1 start notepad
```

### System Info
```bash
# Screen resolution
powershell.exe -File win_control.ps1 info

# Cursor position
powershell.exe -File win_control.ps1 getcursor
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
2. **Rename to `.ps1`**: `ren C:\temp\win_control.txt win_control.ps1`
3. Set PowerShell execution policy (one-time):
   ```powershell
   Set-ExecutionPolicy Bypass -Scope Process
   ```
4. Test connection:
   ```bash
   powershell.exe -File C:\temp\win_control.ps1 info
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
