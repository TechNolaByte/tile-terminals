#Requires -Version 5.1
<#
.SYNOPSIS
    Tiles all open terminal windows.
    Fast path: if TileTerminals.exe exists (compiled by install script), uses it directly.
    Slow path: pure PowerShell fallback when exe is not present.
#>
param(
    [int]   $Cols       = 0,
    [long]  $CallerHwnd = 0,
    [switch]$Elevated
)

$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Always capture foreground HWND first, before anything steals focus
# ---------------------------------------------------------------------------
if (-not ([System.Management.Automation.PSTypeName]'W32.TTilerPre').Type) {
    Add-Type -Name TTilerPre -Namespace 'W32' -MemberDefinition @'
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern System.IntPtr GetForegroundWindow();
'@
}
$fgHwnd = 0
try { $fgHwnd = [W32.TTilerPre]::GetForegroundWindow().ToInt64() } catch {}

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)

# ---------------------------------------------------------------------------
# FAST PATH: delegate to compiled exe if available
# ---------------------------------------------------------------------------
$exePath   = Join-Path $PSScriptRoot 'TileTerminals.exe'
$taskName  = 'TileTerminals_Elevated'
$handoff   = Join-Path $env:TEMP 'TileTerminals_handoff.tmp'

if (Test-Path $exePath) {
    # Write handoff: callerHwnd and optional cols
    "$fgHwnd`n$Cols" | Set-Content $handoff -Encoding ASCII

    if ($isAdmin) {
        & $exePath
    } else {
        $task = Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue
        if ($task) {
            schtasks /run /tn $taskName | Out-Null
        } else {
            Start-Process $exePath -Verb RunAs
        }
    }
    exit
}

# ---------------------------------------------------------------------------
# SLOW PATH: pure PowerShell (no exe compiled yet)
# ---------------------------------------------------------------------------
if (-not $isAdmin) {
    if ($CallerHwnd -eq 0) { $CallerHwnd = $fgHwnd }
    $argStr = "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    $argStr += " -Elevated -CallerHwnd $CallerHwnd"
    if ($Cols -gt 0) { $argStr += " -Cols $Cols" }
    Start-Process powershell -Verb RunAs -ArgumentList $argStr
    exit
}
if ($CallerHwnd -eq 0) { $CallerHwnd = $fgHwnd }

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
    public const uint SPI_GETWORKAREA=0x0030,SWP_NOSIZE=1,SWP_NOMOVE=2,SWP_NOZORDER=4,SWP_SHOWWINDOW=0x40,SWP_FRAMECHANGED=0x20;
    public const int SW_RESTORE=9;
    public static readonly IntPtr HWND_TOPMOST=new IntPtr(-1),HWND_NOTOPMOST=new IntPtr(-2);
    [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
    public struct RECT { public int Left,Top,Right,Bottom; }
    static readonly HashSet<string> TC=new HashSet<string>(StringComparer.OrdinalIgnoreCase){"ConsoleWindowClass","CASCADIA_HOSTING_WINDOW_CLASS","VirtualConsoleClass"};
    static readonly HashSet<string> TP=new HashSet<string>(StringComparer.OrdinalIgnoreCase){"WindowsTerminal","alacritty","hyper","mintty","ConEmu","ConEmu64","ConEmuC","ConEmuC64","cmder","FluentTerminal","tabby","terminus","wezterm-gui"};
    public static List<long[]> FindTerminals(int selfPid,long exHwnd){
        var deny=new HashSet<int>{selfPid};var found=new List<long[]>();
        var shell=GetShellWindow();var desk=GetDesktopWindow();var names=new Dictionary<int,string>();
        EnumWindows((h,_)=>{
            if(!IsWindowVisible(h)||h==shell||h==desk)return true;
            var t=new StringBuilder(256);if(GetWindowText(h,t,256)==0)return true;
            uint u;GetWindowThreadProcessId(h,out u);int pid=(int)u;
            if(deny.Contains(pid))return true;
            long hl=h.ToInt64();if(exHwnd!=0&&hl==exHwnd)return true;
            var c=new StringBuilder(128);GetClassName(h,c,128);
            if(TC.Contains(c.ToString())){found.Add(new long[]{hl,pid});return true;}
            if(!names.ContainsKey(pid)){try{names[pid]=System.Diagnostics.Process.GetProcessById(pid).ProcessName;}catch{names[pid]="";}}
            if(TP.Contains(names[pid]))found.Add(new long[]{hl,pid});
            return true;
        },IntPtr.Zero);return found;
    }
    public static void RaiseWindow(IntPtr h){uint f=SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW;SetWindowPos(h,HWND_TOPMOST,0,0,0,0,f);SetWindowPos(h,HWND_NOTOPMOST,0,0,0,0,f);}
    public static RECT GetWorkArea(){var r=new RECT();SystemParametersInfo(SPI_GETWORKAREA,0,ref r,0);return r;}
}
'@
}

Write-Host ("Caller HWND = {0}" -f $CallerHwnd) -ForegroundColor DarkGray
$allFound = [TTiler4]::FindTerminals([int]$PID, $CallerHwnd)
if ($allFound.Count -eq 0) { Write-Host "No terminals to tile." -ForegroundColor Yellow; exit 0 }

$n       = $allFound.Count
$numCols = if ($Cols -gt 0) { [Math]::Min($Cols,$n) } else { [int][Math]::Ceiling([Math]::Sqrt([double]$n)) }
$numRows = [int][Math]::Ceiling($n / $numCols)
$work    = [TTiler4]::GetWorkArea()
$cellW   = [int](($work.Right-$work.Left) / $numCols)
$cellH   = [int](($work.Bottom-$work.Top) / $numRows)

$cacheFile  = Join-Path $env:TEMP 'TileTerminals_cache.json'
$slotTable  = @{}
$cachedCols = 0; $cachedRows = 0
if (Test-Path $cacheFile) {
    try {
        $raw = Get-Content $cacheFile -Raw | ConvertFrom-Json
        if ($raw.PSObject.Properties['cols']) { $cachedCols=[int]$raw.cols }
        if ($raw.PSObject.Properties['rows']) { $cachedRows=[int]$raw.rows }
        if ($raw.PSObject.Properties['slots']) {
            foreach ($p in $raw.slots.PSObject.Properties) { $slotTable[$p.Name]=[int]$p.Value }
        }
    } catch {}
}

$gridChanged = ($slotTable.Count -eq 0) -or ($cachedCols -ne $numCols) -or ($cachedRows -ne $numRows)
$slotMap = New-Object 'object[]' $n

if ($gridChanged) {
    for ($i=0;$i -lt $n;$i++) { $slotMap[$i]=$allFound[$i] }
} else {
    $used=New-Object 'System.Collections.Generic.HashSet[int]'
    $unplaced=New-Object 'System.Collections.Generic.List[int]'
    for ($i=0;$i -lt $n;$i++) {
        $key="$([long]$allFound[$i][0])"
        if ($slotTable.ContainsKey($key)) {
            $slot=$slotTable[$key]
            if ($slot -lt $n -and -not $used.Contains($slot)) { $slotMap[$slot]=$allFound[$i];$null=$used.Add($slot);continue }
        }
        $null=$unplaced.Add($i)
    }
    $empty=@(0..($n-1)|Where-Object{$null -eq $slotMap[$_]})
    for ($ei=0;$ei -lt $unplaced.Count;$ei++) { $slotMap[$empty[$ei]]=$allFound[$unplaced[$ei]] }
}

$slotsObj=[ordered]@{}
for ($i=0;$i -lt $n;$i++) { $slotsObj["$([long]$slotMap[$i][0])"]=$i }
@{cols=$numCols;rows=$numRows;slots=$slotsObj}|ConvertTo-Json -Depth 4|Set-Content $cacheFile -Encoding UTF8

$posFlags=[TTiler4]::SWP_NOZORDER -bor [TTiler4]::SWP_SHOWWINDOW -bor [TTiler4]::SWP_FRAMECHANGED
for ($i=0;$i -lt $n;$i++) {
    $hwnd=[IntPtr][long]$slotMap[$i][0]
    $x=$work.Left+(($i%$numCols)*$cellW); $y=$work.Top+(([int]($i/$numCols))*$cellH)
    [void][TTiler4]::ShowWindow($hwnd,[TTiler4]::SW_RESTORE)
    [void][TTiler4]::SetWindowPos($hwnd,[IntPtr]::Zero,$x,$y,$cellW,$cellH,$posFlags)
    [TTiler4]::RaiseWindow($hwnd)
}
if ($CallerHwnd -ne 0) { Start-Sleep -Milliseconds 80; [TTiler4]::RaiseWindow([IntPtr]$CallerHwnd) }
Write-Host ("Done. {0} tiled in {1}x{2}." -f $n,$numCols,$numRows) -ForegroundColor Green