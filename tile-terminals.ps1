#Requires -Version 5.1
<#
.SYNOPSIS
    Auto-elevates and tiles all open terminal windows.
    Keeps each terminal in the same grid slot across runs (cache keyed by HWND).
    The calling terminal is excluded from the grid and raised on top last.
#>
param(
    [int]   $Cols       = 0,
    [long]  $CallerHwnd = 0,
    [switch]$Elevated
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# PRE-ELEVATION: capture calling window HWND, then re-launch elevated.
# ---------------------------------------------------------------------------
$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

if (-not $isAdmin) {
    if (-not ([System.Management.Automation.PSTypeName]'TTilerPre').Type) {
        Add-Type -Name TTilerPre -Namespace 'W32' -MemberDefinition @'
            [System.Runtime.InteropServices.DllImport("user32.dll")]
            public static extern System.IntPtr GetForegroundWindow();
'@
    }
    $hwndLong = [W32.TTilerPre]::GetForegroundWindow().ToInt64()

    $argStr = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $argStr += " -Elevated -CallerHwnd $hwndLong"
    if ($Cols -gt 0) { $argStr += " -Cols $Cols" }
    Start-Process powershell -Verb RunAs -ArgumentList $argStr
    exit
}

# ---------------------------------------------------------------------------
# Win32 helpers (bump type name whenever C# changes to avoid stale-session)
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'TTiler4').Type) {
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;

public static class TTiler4 {

    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")] public static extern bool   EnumWindows(EnumWindowsProc fn, IntPtr lp);
    [DllImport("user32.dll")] public static extern bool   IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] public static extern int    GetWindowText(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int    GetClassName(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern uint   GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] public static extern bool   SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int w, int h, uint flags);
    [DllImport("user32.dll")] public static extern bool   ShowWindow(IntPtr hWnd, int cmd);
    [DllImport("user32.dll")] public static extern bool   SystemParametersInfo(uint a, uint b, ref RECT r, uint c);
    [DllImport("user32.dll")] public static extern IntPtr GetShellWindow();
    [DllImport("user32.dll")] public static extern IntPtr GetDesktopWindow();

    public const uint SPI_GETWORKAREA  = 0x0030;
    public const uint SWP_NOSIZE       = 0x0001;
    public const uint SWP_NOMOVE       = 0x0002;
    public const uint SWP_NOZORDER     = 0x0004;
    public const uint SWP_SHOWWINDOW   = 0x0040;
    public const uint SWP_FRAMECHANGED = 0x0020;
    public const int  SW_RESTORE       = 9;

    public static readonly IntPtr HWND_TOPMOST   = new IntPtr(-1);
    public static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);

    [StructLayout(LayoutKind.Sequential)]
    public struct RECT { public int Left, Top, Right, Bottom; }

    private static readonly HashSet<string> TClasses = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "ConsoleWindowClass", "CASCADIA_HOSTING_WINDOW_CLASS", "VirtualConsoleClass"
    };

    private static readonly HashSet<string> TProcs = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
        "WindowsTerminal", "alacritty", "hyper", "mintty",
        "ConEmu", "ConEmu64", "ConEmuC", "ConEmuC64", "cmder",
        "FluentTerminal", "tabby", "terminus", "wezterm-gui"
    };

    // Returns list of long[2]{hwnd, pid}
    public static List<long[]> FindTerminals(int[] denyPids) {
        var deny  = new HashSet<int>(denyPids);
        var found = new List<long[]>();
        var shell = GetShellWindow();
        var desk  = GetDesktopWindow();
        var names = new Dictionary<int, string>();

        EnumWindows((hWnd, _) => {
            if (!IsWindowVisible(hWnd)) return true;
            if (hWnd == shell || hWnd == desk) return true;
            var title = new StringBuilder(256);
            if (GetWindowText(hWnd, title, 256) == 0) return true;

            uint u; GetWindowThreadProcessId(hWnd, out u);
            int pid = (int)u;
            if (deny.Contains(pid)) return true;

            var cls = new StringBuilder(128);
            GetClassName(hWnd, cls, 128);
            if (TClasses.Contains(cls.ToString())) {
                found.Add(new long[] { hWnd.ToInt64(), pid });
                return true;
            }
            if (!names.ContainsKey(pid)) {
                try   { names[pid] = System.Diagnostics.Process.GetProcessById(pid).ProcessName; }
                catch { names[pid] = ""; }
            }
            if (TProcs.Contains(names[pid]))
                found.Add(new long[] { hWnd.ToInt64(), pid });
            return true;
        }, IntPtr.Zero);

        return found;
    }

    public static void RaiseWindow(IntPtr hWnd) {
        uint f = SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW;
        SetWindowPos(hWnd, HWND_TOPMOST,   0, 0, 0, 0, f);
        SetWindowPos(hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, f);
    }

    public static RECT GetWorkArea() {
        var r = new RECT();
        SystemParametersInfo(SPI_GETWORKAREA, 0, ref r, 0);
        return r;
    }
}
'@
}

Write-Host "Caller HWND received: $CallerHwnd" -ForegroundColor DarkGray

# ---------------------------------------------------------------------------
# Discover terminal windows, exclude caller
# ---------------------------------------------------------------------------
$allFound = [TTiler4]::FindTerminals([int[]]@([int]$PID))

$callerEntry = $null
$found = New-Object 'System.Collections.Generic.List[object]'
foreach ($e in $allFound) {
    if ($CallerHwnd -ne 0 -and [long]$e[0] -eq $CallerHwnd) {
        $callerEntry = $e
        Write-Host "Excluded caller (hwnd=$($e[0]), pid=$($e[1])) from grid." -ForegroundColor DarkGray
    } else {
        $found.Add($e)
    }
}

if ($found.Count -eq 0) {
    Write-Host "No terminal windows found to tile." -ForegroundColor Yellow
    if ($callerEntry) { [TTiler4]::RaiseWindow([IntPtr][long]$callerEntry[0]) }
    Start-Sleep -Seconds 2
    exit 0
}

# ---------------------------------------------------------------------------
# Grid dimensions
# ---------------------------------------------------------------------------
$n       = $found.Count
$numCols = if ($Cols -gt 0) { [Math]::Min($Cols, $n) }
           else             { [int][Math]::Ceiling([Math]::Sqrt([double]$n)) }
$numRows = [int][Math]::Ceiling($n / $numCols)

$work    = [TTiler4]::GetWorkArea()
$screenW = $work.Right  - $work.Left
$screenH = $work.Bottom - $work.Top
$cellW   = [int]($screenW / $numCols)
$cellH   = [int]($screenH / $numRows)

Write-Host "Grid: ${numCols}x${numRows} for $n window(s)." -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# HWND position cache  (PID is unreliable: Windows Terminal = 1 PID / N windows)
# ---------------------------------------------------------------------------
$cacheFile = Join-Path $env:TEMP 'TileTerminals_cache.json'

# Load cache -> plain hashtable { "hwndString" = slotInt }
$slotTable = @{}
$cachedCols = 0
$cachedRows = 0
if (Test-Path $cacheFile) {
    try {
        $raw = Get-Content $cacheFile -Raw | ConvertFrom-Json
        if ($raw.PSObject.Properties['cols']) { $cachedCols = [int]$raw.cols }
        if ($raw.PSObject.Properties['rows']) { $cachedRows = [int]$raw.rows }
        if ($raw.PSObject.Properties['slots']) {
            foreach ($p in $raw.slots.PSObject.Properties) {
                $slotTable[$p.Name] = [int]$p.Value
            }
        }
    } catch { Write-Host "Cache read failed: $_" -ForegroundColor DarkYellow }
}

Write-Host "Cache loaded: $($slotTable.Count) entries, previous grid ${cachedCols}x${cachedRows}." -ForegroundColor DarkGray

$gridChanged = ($slotTable.Count -eq 0) -or ($cachedCols -ne $numCols) -or ($cachedRows -ne $numRows)

$slotMap     = New-Object 'object[]' $n    # slot -> long[]{hwnd,pid}
$hits = 0; $misses = 0

if ($gridChanged) {
    Write-Host "Grid changed or no cache -> fresh layout." -ForegroundColor DarkGray
    for ($i = 0; $i -lt $n; $i++) { $slotMap[$i] = $found[$i] }
} else {
    $usedSlots   = New-Object 'System.Collections.Generic.HashSet[int]'
    $unplacedIdx = New-Object 'System.Collections.Generic.List[int]'

    # Pass 1: restore cached slots (key by HWND)
    for ($i = 0; $i -lt $n; $i++) {
        $key = "$([long]$found[$i][0])"
        if ($slotTable.ContainsKey($key)) {
            $slot = $slotTable[$key]
            if ($slot -lt $n -and -not $usedSlots.Contains($slot)) {
                $slotMap[$slot] = $found[$i]
                $null = $usedSlots.Add($slot)
                $hits++
                continue
            }
        }
        $null = $unplacedIdx.Add($i)
        $misses++
    }
    # Pass 2: fill empty slots with new windows
    $emptySlots = @(0..($n-1) | Where-Object { $null -eq $slotMap[$_] })
    for ($ei = 0; $ei -lt $unplacedIdx.Count; $ei++) {
        $slotMap[$emptySlots[$ei]] = $found[$unplacedIdx[$ei]]
    }
    Write-Host "Cache hits: $hits   new: $misses" -ForegroundColor DarkGray
}

# Save cache (key by HWND)
$slotsObj = [ordered]@{}
for ($i = 0; $i -lt $n; $i++) {
    $slotsObj["$([long]$slotMap[$i][0])"] = $i
}
@{ cols = $numCols; rows = $numRows; slots = $slotsObj } |
    ConvertTo-Json -Depth 4 | Set-Content $cacheFile -Encoding UTF8

# ---------------------------------------------------------------------------
# Tile + raise
# ---------------------------------------------------------------------------
$posFlags = [TTiler4]::SWP_NOZORDER -bor [TTiler4]::SWP_SHOWWINDOW -bor [TTiler4]::SWP_FRAMECHANGED

for ($i = 0; $i -lt $n; $i++) {
    $hwnd = [IntPtr][long]$slotMap[$i][0]
    $x    = $work.Left + (($i % $numCols) * $cellW)
    $y    = $work.Top  + ([int]($i / $numCols) * $cellH)

    [void][TTiler4]::ShowWindow($hwnd, [TTiler4]::SW_RESTORE)
    [void][TTiler4]::SetWindowPos($hwnd, [IntPtr]::Zero, $x, $y, $cellW, $cellH, $posFlags)
    [TTiler4]::RaiseWindow($hwnd)
}

# Raise caller last (it stays in its own position)
if ($callerEntry) {
    Start-Sleep -Milliseconds 80
    [TTiler4]::RaiseWindow([IntPtr][long]$callerEntry[0])
}

Write-Host ("Done. {0} tiled in {1}x{2}. Cache: {3}" -f $n, $numCols, $numRows, $cacheFile) -ForegroundColor Green