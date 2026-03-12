# Windows Control Skill

[English](#english) | [中文](#中文)

---

<a name="english"></a>
## English

### Overview

**Windows Control Skill** is an OpenClaw skill that enables remote control of Windows desktop via PowerShell. It provides full GUI automation capabilities including screenshots, mouse/keyboard control, window management, and application automation.

### Features

- 📸 **Screenshot** - Capture full screen or specific regions
- 🖱️ **Mouse Control** - Click, move, drag, scroll
- ⌨️ **Keyboard Control** - Type text, press keys, send combinations
- 🪟 **Window Management** - List, activate, close windows
- 🚀 **Application Control** - Start and manage applications
- 📹 **Real-time Capture** - Continuous screen recording

### Requirements

#### Windows (Target Machine)
- Windows 7 or later with PowerShell
- Execution policy set to Bypass:
  ```powershell
  Set-ExecutionPolicy Bypass -Scope Process
  ```

#### WSL/Linux (Control Machine)
- WSL with Windows drive access (`/mnt/c/`)
- Or native Windows with Python
- Python 3.7+ (for real-time capture)
- Pillow library: `pip install Pillow`

### Quick Start

```bash
# 1. Copy script to Windows
cp ~/.openclaw/skills/win-control/scripts/win_control.ps1 /mnt/c/temp/

# 2. Execute commands
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.ps1 screenshot C:\temp\screen.png
```

### Commands

#### Screenshot
```bash
powershell.exe -File win_control.ps1 screenshot C:\temp\screen.png
```

#### Mouse Control
```bash
# Click at position (x, y)
powershell.exe -File win_control.ps1 click 500 300

# Right click
powershell.exe -File win_control.ps1 click 500 300 right

# Double click
powershell.exe -File win_control.ps1 click 500 300 left double

# Move cursor
powershell.exe -File win_control.ps1 move 500 300

# Drag from (100, 100) to (500, 500)
powershell.exe -File win_control.ps1 drag 100 100 500 500

# Scroll down 5 lines
powershell.exe -File win_control.ps1 scroll down 5
```

#### Keyboard Control
```bash
# Type text
powershell.exe -File win_control.ps1 type "Hello World"

# Press key
powershell.exe -File win_control.ps1 key enter
powershell.exe -File win_control.ps1 key escape
powershell.exe -File win_control.ps1 key tab

# Key combinations
powershell.exe -File win_control.ps1 combo "ctrl+c"
powershell.exe -File win_control.ps1 combo "alt+tab"
powershell.exe -File win_control.ps1 combo "ctrl+alt+delete"
```

#### Window Management
```bash
# List all windows
powershell.exe -File win_control.ps1 list

# Activate window by title
powershell.exe -File win_control.ps1 activate "Notepad"

# Close window by title
powershell.exe -File win_control.ps1 close "Calculator"

# Start application
powershell.exe -File win_control.ps1 start notepad
powershell.exe -File win_control.ps1 start chrome
```

#### System Info
```bash
# Get screen resolution
powershell.exe -File win_control.ps1 info

# Get cursor position
powershell.exe -File win_control.ps1 getcursor
```

### Real-time Screen Capture

```bash
# Capture at 5 FPS for 30 seconds
python3 ~/.openclaw/skills/win-control/scripts/realtime_screen.py --fps 5 --duration 30

# Get latest frame
python3 -c "
from realtime_screen import WindowsScreenCapture
cap = WindowsScreenCapture()
print(cap.get_latest_frame())
"
```

### Available Keys

- **Navigation**: up, down, left, right, home, end, pageup, pagedown
- **Editing**: enter, return, tab, backspace, delete, escape
- **Modifiers**: Use with combo - ctrl, alt, shift

### Troubleshooting

**Execution Policy Error:**
```powershell
Set-ExecutionPolicy Bypass -Scope Process
```

**Path Issues:**
Use full paths or copy script to `C:\temp\`

### License

MIT License - Free to use, modify, and distribute.

---

<a name="中文"></a>
## 中文

### 概述

**Windows Control Skill** 是一个 OpenClaw 技能，用于通过 PowerShell 远程控制 Windows 桌面。提供完整的 GUI 自动化功能，包括截图、鼠标/键盘控制、窗口管理和应用程序自动化。

### 功能特性

- 📸 **截图** - 捕获全屏或特定区域
- 🖱️ **鼠标控制** - 点击、移动、拖拽、滚动
- ⌨️ **键盘控制** - 输入文字、按键、发送组合键
- 🪟 **窗口管理** - 列出、激活、关闭窗口
- 🚀 **应用程序控制** - 启动和管理应用程序
- 📹 **实时捕获** - 连续屏幕录制

### 系统要求

#### Windows (目标机器)
- Windows 7 或更高版本，带 PowerShell
- 执行策略设置为 Bypass：
  ```powershell
  Set-ExecutionPolicy Bypass -Scope Process
  ```

#### WSL/Linux (控制机器)
- WSL 带 Windows 驱动器访问 (`/mnt/c/`)
- 或带 Python 的原生 Windows
- Python 3.7+ (用于实时捕获)
- Pillow 库：`pip install Pillow`

### 快速开始

```bash
# 1. 复制脚本到 Windows
cp ~/.openclaw/skills/win-control/scripts/win_control.ps1 /mnt/c/temp/

# 2. 执行命令
powershell.exe -ExecutionPolicy Bypass -File C:\temp\win_control.ps1 screenshot C:\temp\screen.png
```

### 命令说明

#### 截图
```bash
powershell.exe -File win_control.ps1 screenshot C:\temp\screen.png
```

#### 鼠标控制
```bash
# 在位置 (x, y) 点击
powershell.exe -File win_control.ps1 click 500 300

# 右键点击
powershell.exe -File win_control.ps1 click 500 300 right

# 双击
powershell.exe -File win_control.ps1 click 500 300 left double

# 移动光标
powershell.exe -File win_control.ps1 move 500 300

# 从 (100, 100) 拖拽到 (500, 500)
powershell.exe -File win_control.ps1 drag 100 100 500 500

# 向下滚动 5 行
powershell.exe -File win_control.ps1 scroll down 5
```

#### 键盘控制
```bash
# 输入文字
powershell.exe -File win_control.ps1 type "Hello World"

# 按键
powershell.exe -File win_control.ps1 key enter
powershell.exe -File win_control.ps1 key escape
powershell.exe -File win_control.ps1 key tab

# 组合键
powershell.exe -File win_control.ps1 combo "ctrl+c"
powershell.exe -File win_control.ps1 combo "alt+tab"
powershell.exe -File win_control.ps1 combo "ctrl+alt+delete"
```

#### 窗口管理
```bash
# 列出所有窗口
powershell.exe -File win_control.ps1 list

# 按标题激活窗口
powershell.exe -File win_control.ps1 activate "Notepad"

# 按标题关闭窗口
powershell.exe -File win_control.ps1 close "Calculator"

# 启动应用程序
powershell.exe -File win_control.ps1 start notepad
powershell.exe -File win_control.ps1 start chrome
```

#### 系统信息
```bash
# 获取屏幕分辨率
powershell.exe -File win_control.ps1 info

# 获取光标位置
powershell.exe -File win_control.ps1 getcursor
```

### 实时屏幕捕获

```bash
# 以 5 FPS 捕获 30 秒
python3 ~/.openclaw/skills/win-control/scripts/realtime_screen.py --fps 5 --duration 30

# 获取最新帧
python3 -c "
from realtime_screen import WindowsScreenCapture
cap = WindowsScreenCapture()
print(cap.get_latest_frame())
"
```

### 可用按键

- **导航**：up, down, left, right, home, end, pageup, pagedown
- **编辑**：enter, return, tab, backspace, delete, escape
- **修饰键**：与 combo 一起使用 - ctrl, alt, shift

### 故障排除

**执行策略错误：**
```powershell
Set-ExecutionPolicy Bypass -Scope Process
```

**路径问题：**
使用完整路径或将脚本复制到 `C:\temp\`

### 许可证

MIT 许可证 - 可自由使用、修改和分发。

---

## Links

- GitHub: https://github.com/KAUIN/windows-control-skill
- OpenClaw: https://openclaw.ai
