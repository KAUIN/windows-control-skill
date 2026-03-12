# Windows Desktop Control via PowerShell
# Usage: powershell.exe -File win_control.ps1 <command> [args]

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet("screenshot", "click", "move", "drag", "scroll", 
                 "type", "key", "combo", "getcursor", "info", 
                 "start", "list", "activate", "close")]
    [string]$Command,
    
    [Parameter(ValueFromRemainingArguments=$true)]
    [string[]]$Args
)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Mouse event constants
$MOUSEEVENTF_LEFTDOWN = 0x0002
$MOUSEEVENTF_LEFTUP = 0x0004
$MOUSEEVENTF_RIGHTDOWN = 0x0008
$MOUSEEVENTF_RIGHTUP = 0x0010
$MOUSEEVENTF_MIDDLEDOWN = 0x0020
$MOUSEEVENTF_MIDDLEUP = 0x0040
$MOUSEEVENTF_WHEEL = 0x0800

# Screenshot
function Take-Screenshot {
    param([string]$OutputPath)
    
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
    $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
    $graphics.CopyFromScreen($screen.Location, [System.Drawing.Point]::Empty, $screen.Size)
    $bitmap.Save($OutputPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $graphics.Dispose()
    $bitmap.Dispose()
    Write-Output "Screenshot saved: $OutputPath"
}

# Mouse Move
function Move-Cursor {
    param([int]$X, [int]$Y)
    [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($X, $Y)
    Write-Output "Cursor moved to ($X, $Y)"
}

# Mouse Click
function Click-Mouse {
    param(
        [int]$X,
        [int]$Y,
        [ValidateSet("left", "right", "middle")]
        [string]$Button = "left",
        [switch]$Double
    )
    
    Move-Cursor -X $X -Y $Y
    Start-Sleep -Milliseconds 50
    
    $signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
'@
    
    $SendMouse = Add-Type -MemberDefinition $signature -Name "Win32Mouse" -Namespace Win32Functions -PassThru
    
    $down = switch ($Button) {
        "left" { $MOUSEEVENTF_LEFTDOWN }
        "right" { $MOUSEEVENTF_RIGHTDOWN }
        "middle" { $MOUSEEVENTF_MIDDLEDOWN }
    }
    $up = switch ($Button) {
        "left" { $MOUSEEVENTF_LEFTUP }
        "right" { $MOUSEEVENTF_RIGHTUP }
        "middle" { $MOUSEEVENTF_MIDDLEUP }
    }
    
    $SendMouse::mouse_event($down, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 50
    $SendMouse::mouse_event($up, 0, 0, 0, 0)
    
    if ($Double) {
        Start-Sleep -Milliseconds 100
        $SendMouse::mouse_event($down, 0, 0, 0, 0)
        Start-Sleep -Milliseconds 50
        $SendMouse::mouse_event($up, 0, 0, 0, 0)
    }
    
    Write-Output "Clicked $Button at ($X, $Y)"
}

# Mouse Drag
function Drag-Mouse {
    param(
        [int]$StartX, [int]$StartY,
        [int]$EndX, [int]$EndY,
        [ValidateSet("left", "right", "middle")]
        [string]$Button = "left"
    )
    
    $signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
'@
    
    $SendMouse = Add-Type -MemberDefinition $signature -Name "Win32MouseDrag" -Namespace Win32Functions -PassThru
    
    $down = switch ($Button) {
        "left" { $MOUSEEVENTF_LEFTDOWN }
        "right" { $MOUSEEVENTF_RIGHTDOWN }
        "middle" { $MOUSEEVENTF_MIDDLEDOWN }
    }
    $up = switch ($Button) {
        "left" { $MOUSEEVENTF_LEFTUP }
        "right" { $MOUSEEVENTF_RIGHTUP }
        "middle" { $MOUSEEVENTF_MIDDLEUP }
    }
    
    Move-Cursor -X $StartX -Y $StartY
    Start-Sleep -Milliseconds 100
    $SendMouse::mouse_event($down, 0, 0, 0, 0)
    Start-Sleep -Milliseconds 100
    Move-Cursor -X $EndX -Y $EndY
    Start-Sleep -Milliseconds 100
    $SendMouse::mouse_event($up, 0, 0, 0, 0)
    
    Write-Output "Dragged from ($StartX, $StartY) to ($EndX, $EndY)"
}

# Scroll
function Scroll-Mouse {
    param(
        [ValidateSet("up", "down")]
        [string]$Direction,
        [int]$Amount = 3
    )
    
    $signature = @'
    [DllImport("user32.dll", CharSet=CharSet.Auto, CallingConvention=CallingConvention.StdCall)]
    public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, int dwExtraInfo);
'@
    
    $SendMouse = Add-Type -MemberDefinition $signature -Name "Win32Scroll" -Namespace Win32Functions -PassThru
    
    $WHEEL_DELTA = 120
    $delta = if ($Direction -eq "up") { $WHEEL_DELTA } else { -$WHEEL_DELTA }
    
    for ($i = 0; $i -lt $Amount; $i++) {
        $SendMouse::mouse_event($MOUSEEVENTF_WHEEL, 0, 0, $delta, 0)
        Start-Sleep -Milliseconds 50
    }
    
    Write-Output "Scrolled $Direction $Amount times"
}

# Type text
function Send-Text {
    param([string]$Text)
    [System.Windows.Forms.SendKeys]::SendWait($Text)
    Write-Output "Typed: $Text"
}

# Press key
function Send-Key {
    param([string]$Key)
    [System.Windows.Forms.SendKeys]::SendWait("{$Key}")
    Write-Output "Pressed key: $Key"
}

# Key combo
function Send-Combo {
    param([string]$Combo)
    $keySeq = $Combo -replace 'ctrl\+', '^' -replace 'alt\+', '%' -replace 'shift\+', '+'
    [System.Windows.Forms.SendKeys]::SendWait($keySeq)
    Write-Output "Sent combo: $Combo"
}

# Get cursor position
function Get-CursorPosition {
    $pos = [System.Windows.Forms.Cursor]::Position
    Write-Output "Cursor: ($($pos.X), $($pos.Y))"
}

# Get screen info
function Get-ScreenInfo {
    $screen = [System.Windows.Forms.Screen]::PrimaryScreen
    Write-Output "Screen: $($screen.Bounds.Width)x$($screen.Bounds.Height)"
    Write-Output "Working: $($screen.WorkingArea.Width)x$($screen.WorkingArea.Height)"
}

# Start application
function Start-Application {
    param([string]$App)
    Start-Process $App
    Write-Output "Started: $App"
}

# List windows
function List-Windows {
    Get-Process | Where-Object {$_.MainWindowTitle -ne ""} | 
        Select-Object ProcessName, MainWindowTitle, Id | 
        Format-Table -AutoSize
}

# Activate window
function Activate-Window {
    param([string]$Title)
    $signature = @'
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool SetForegroundWindow(IntPtr hWnd);
    
    [DllImport("user32.dll")]
    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    
    $User32 = Add-Type -MemberDefinition $signature -Name "Win32Window" -Namespace Win32Functions -PassThru
    
    $hwnd = $User32::FindWindow($null, $Title)
    if ($hwnd -ne 0) {
        $User32::ShowWindow($hwnd, 9) | Out-Null
        $User32::SetForegroundWindow($hwnd) | Out-Null
        Write-Output "Activated: $Title"
    } else {
        Write-Output "Window not found: $Title"
    }
}

# Close window
function Close-Window {
    param([string]$Title)
    $signature = @'
    [DllImport("user32.dll")]
    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
    
    [DllImport("user32.dll")]
    public static extern bool PostMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);
'@
    
    $User32 = Add-Type -MemberDefinition $signature -Name "Win32Close" -Namespace Win32Functions -PassThru
    
    $hwnd = $User32::FindWindow($null, $Title)
    if ($hwnd -ne 0) {
        $User32::PostMessage($hwnd, 0x10, 0, 0) | Out-Null
        Write-Output "Closed: $Title"
    } else {
        Write-Output "Window not found: $Title"
    }
}

# Main execution
switch ($Command) {
    "screenshot" {
        $path = if ($Args[0]) { $Args[0] } else { "C:\temp\screenshot.png" }
        Take-Screenshot -OutputPath $path
    }
    "click" {
        $x = [int]$Args[0]
        $y = [int]$Args[1]
        $button = if ($Args[2]) { $Args[2] } else { "left" }
        $double = $Args[3] -eq "double"
        Click-Mouse -X $x -Y $y -Button $button -Double:$double
    }
    "move" {
        $x = [int]$Args[0]
        $y = [int]$Args[1]
        Move-Cursor -X $x -Y $y
    }
    "drag" {
        $sx = [int]$Args[0]
        $sy = [int]$Args[1]
        $ex = [int]$Args[2]
        $ey = [int]$Args[3]
        $btn = if ($Args[4]) { $Args[4] } else { "left" }
        Drag-Mouse -StartX $sx -StartY $sy -EndX $ex -EndY $ey -Button $btn
    }
    "scroll" {
        $dir = $Args[0]
        $amount = if ($Args[1]) { [int]$Args[1] } else { 3 }
        Scroll-Mouse -Direction $dir -Amount $amount
    }
    "type" {
        Send-Text -Text ($Args -join " ")
    }
    "key" {
        Send-Key -Key $Args[0]
    }
    "combo" {
        Send-Combo -Combo $Args[0]
    }
    "getcursor" {
        Get-CursorPosition
    }
    "info" {
        Get-ScreenInfo
    }
    "start" {
        Start-Application -App $Args[0]
    }
    "list" {
        List-Windows
    }
    "activate" {
        Activate-Window -Title $Args[0]
    }
    "close" {
        Close-Window -Title $Args[0]
    }
}
