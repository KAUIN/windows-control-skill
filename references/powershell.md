# PowerShell Windows Control Reference

Advanced PowerShell usage for Windows desktop automation.

## Execution Policy

PowerShell requires execution policy to be set for scripts:

```powershell
# Current session only (recommended)
Set-ExecutionPolicy Bypass -Scope Process

# Current user
Set-ExecutionPolicy Bypass -Scope CurrentUser

# All users (requires admin)
Set-ExecutionPolicy Bypass -Scope LocalMachine
```

## Mouse Control

### Coordinates
- Origin (0,0) is top-left corner
- X increases to the right
- Y increases downward
- Get current resolution: `powershell.exe -File win_control.ps1 info`

### Click Types
- `left` - Left click (default)
- `right` - Right click
- `middle` - Middle click
- `double` - Double click

### Examples
```bash
# Left click at center of screen (960, 540 on 1920x1080)
powershell.exe -File win_control.ps1 click 960 540

# Right click
powershell.exe -File win_control.ps1 click 500 300 right

# Double click
powershell.exe -File win_control.ps1 click 500 300 left double
```

## Keyboard Control

### Special Keys
Use these names with the `key` command:
- Navigation: `up`, `down`, `left`, `right`
- Editing: `enter`, `return`, `tab`, `backspace`, `delete`, `escape`
- Modifiers: `ctrl`, `alt`, `shift` (use with `combo`)

### Key Combinations
Format: `modifier+key`

```bash
# Copy
powershell.exe -File win_control.ps1 combo "ctrl+c"

# Paste
powershell.exe -File win_control.ps1 combo "ctrl+v"

# Alt+Tab
powershell.exe -File win_control.ps1 combo "alt+tab"

# Ctrl+Alt+Delete
powershell.exe -File win_control.ps1 combo "ctrl+alt+delete"

# Ctrl+Shift+Esc (Task Manager)
powershell.exe -File win_control.ps1 combo "ctrl+shift+esc"
```

### Typing Special Characters
Some characters need escaping in PowerShell:

```bash
# Type with quotes
powershell.exe -File win_control.ps1 type '"Hello World"'

# Type path with backslashes
powershell.exe -File win_control.ps1 type 'C:\\Users\\Name'
```

## Window Management

### Finding Window Titles
```bash
powershell.exe -File win_control.ps1 list
```

Output example:
```
ProcessName    MainWindowTitle                    Id
-----------    ---------------                    --
notepad        Untitled - Notepad               1234
chrome         Google Chrome                    5678
```

### Activating Windows
Use exact or partial title match:

```bash
# Activate Notepad
powershell.exe -File win_control.ps1 activate "Notepad"

# Activate Chrome
powershell.exe -File win_control.ps1 activate "Chrome"
```

### Starting Applications
```bash
# Start Notepad
powershell.exe -File win_control.ps1 start notepad

# Start Calculator
powershell.exe -File win_control.ps1 start calc

# Start with full path
powershell.exe -File win_control.ps1 start "C:\Program Files\App\app.exe"

# Start with arguments
powershell.exe -Command "Start-Process notepad -ArgumentList 'C:\\temp\\file.txt'"
```

## Real-time Capture

### Using Python Script
```bash
# Start capture stream
python3 realtime_screen.py --fps 5 --duration 60

# Get latest frame
ls -t /tmp/screen_stream/*.png | head -1
```

### Manual Loop
```bash
while true; do
    powershell.exe -File win_control.ps1 screenshot /tmp/frame_$(date +%s).png
    sleep 0.5
done
```

## Advanced Examples

### Automated Login
```bash
# Click username field
powershell.exe -File win_control.ps1 click 500 300

# Type username
powershell.exe -File win_control.ps1 type "username"

# Press Tab to move to password
powershell.exe -File win_control.ps1 key tab

# Type password
powershell.exe -File win_control.ps1 type "password"

# Press Enter
powershell.exe -File win_control.ps1 key enter
```

### Drag and Drop File
```bash
# Drag from file at (100, 100) to folder at (500, 500)
powershell.exe -File win_control.ps1 drag 100 100 500 500
```

### Scroll Web Page
```bash
# Scroll down 5 lines at position (500, 500)
powershell.exe -File win_control.ps1 scroll down 5
```

### Close All Notepad Windows
```bash
powershell.exe -Command "Get-Process notepad | Stop-Process"
```

## Troubleshooting

### "Execution Policy" Error
```powershell
Set-ExecutionPolicy Bypass -Scope Process
```

### "File not found" Error
Use full path or ensure file exists:
```bash
ls /mnt/c/temp/win_control.ps1
```

### "Access Denied" Error
Run PowerShell as Administrator or check permissions.

### Characters Not Typing Correctly
Some special characters may need escaping. Use single quotes or escape with backtick.

### Window Not Found
Use `list` command to see exact window titles. Titles are case-sensitive.

## Performance Tips

1. **Batch commands**: Combine multiple operations in one script call
2. **Use sleep**: Add delays between actions for UI to respond
3. **Check coordinates**: Use `getcursor` to find positions
4. **Screenshot verification**: Take screenshots to verify state

## Security Considerations

- Scripts run with current user privileges
- Can interact with any visible window
- Be careful with `start` command
- Avoid typing passwords in plain text
- Consider using Windows Credential Manager for sensitive data
