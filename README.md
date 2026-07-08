# tile-terminals

A PowerShell utility that automatically tiles all open terminal windows (Windows Terminal, ConEmu, Alacritty, Hyper, mintty, Tabby, WezTerm, and more) into a grid across your screen's work area.

## How it works

- `tile-terminals.ps1` finds every visible terminal window, arranges them in a grid sized to fit the screen, and remembers window→slot assignments across runs (via a cache in `%TEMP%`) so windows keep their relative positions when the set of open terminals doesn't change shape.
- The window that had focus when you triggered tiling is re-raised on top afterward, so you're not left staring at a random terminal.
- Tiling needs to reposition windows owned by other processes (potentially elevated ones), so it must run elevated. `install_permanent_elevation.ps1` sets that up once so you're never prompted by UAC again:
  - Compiles a small C# console app (`TileTerminals.exe`) that does the same window-tiling work as the PowerShell script, for fast, low-overhead execution.
  - Registers a scheduled task (`TileTerminals_Elevated`) that runs the exe elevated under your account with no UAC prompt.
  - Creates a desktop shortcut that triggers the scheduled task.
  - Running the install script again uninstalls everything (task, shortcut, exe).
- Once `TileTerminals.exe` exists, `tile-terminals.ps1` delegates to it directly (fast path). Without it, the script falls back to tiling terminals itself in pure PowerShell, self-elevating via UAC if needed (slow path).

## Usage

1. Run `install_permanent_elevation.ps1` once to compile the exe and install the scheduled task + shortcut.
2. Trigger tiling by running `tile-terminals.ps1`, using the desktop shortcut, or binding either to a hotkey.

### Parameters (`tile-terminals.ps1`)

| Parameter | Description |
|---|---|
| `-Cols` | Force a specific number of grid columns (default: auto, based on `sqrt(n)`). |
| `-CallerHwnd` | Window handle to re-raise after tiling (auto-detected from the foreground window if omitted). |
| `-Elevated` | Internal flag used when the script self-relaunches with admin rights. |
| `-TerminalHost` | Restrict tiling to a single terminal host, matched case-insensitively against the window's process name (e.g. `WindowsTerminal`, `wezterm-gui`, `alacritty`) or window class (e.g. `ConsoleWindowClass`). Default: empty, meaning all recognized terminal hosts are tiled. Example: `.\tile-terminals.ps1 -TerminalHost WindowsTerminal`. |

Recognized terminal hosts by default (process name or window class): Windows Terminal, ConEmu/cmder, Alacritty, Hyper, mintty, FluentTerminal, Tabby, Terminus, WezTerm, and any raw console host window (`cmd.exe`, `powershell.exe`, etc.).

## Requirements

- Windows with PowerShell 5.1+
- Admin rights for the initial install (to register the scheduled task)
