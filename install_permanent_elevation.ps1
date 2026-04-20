#Requires -Version 5.1
<#
.SYNOPSIS
    First run  -> compiles TileTerminals.exe, installs scheduled task + shortcut.
    Second run -> uninstalls everything (task, shortcut, exe).
#>

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Start-Process powershell -Verb RunAs `
        -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`""
    exit
}

$TaskName     = 'TileTerminals_Elevated'
$ScriptDir    = $PSScriptRoot
$ExePath      = Join-Path $ScriptDir 'TileTerminals.exe'
$ShortcutPath = Join-Path ([Environment]::GetFolderPath('Desktop')) 'TileTerminals.lnk'
$Desc         = 'Tile terminal windows (elevated, no UAC)'

# ---------------------------------------------------------------------------
# TOGGLE: second run = uninstall
# ---------------------------------------------------------------------------
$taskExists = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
if ($taskExists) {
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
    Write-Host "Scheduled task '$TaskName' removed." -ForegroundColor Yellow
    if (Test-Path $ShortcutPath) { Remove-Item $ShortcutPath -Force; Write-Host "Shortcut removed." -ForegroundColor Yellow }
    if (Test-Path $ExePath)      { Remove-Item $ExePath -Force;      Write-Host "TileTerminals.exe removed." -ForegroundColor Yellow }
    Write-Host "Uninstalled." -ForegroundColor Green
    exit 0
}

# ---------------------------------------------------------------------------
# INSTALL: compile C# -> TileTerminals.exe
# ---------------------------------------------------------------------------
Write-Host "Compiling TileTerminals.exe..." -ForegroundColor Cyan

$csharp = @'
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

class TileTerminals {
    delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] static extern bool   EnumWindows(EnumWindowsProc fn, IntPtr lp);
    [DllImport("user32.dll")] static extern bool   IsWindowVisible(IntPtr hWnd);
    [DllImport("user32.dll")] static extern int    GetWindowText(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern int    GetClassName(IntPtr hWnd, StringBuilder s, int n);
    [DllImport("user32.dll")] static extern uint   GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
    [DllImport("user32.dll")] static extern bool   SetWindowPos(IntPtr hWnd, IntPtr after, int x, int y, int w, int h, uint flags);
    [DllImport("user32.dll")] static extern bool   ShowWindow(IntPtr hWnd, int cmd);
    [DllImport("user32.dll")] static extern bool   SystemParametersInfo(uint a, uint b, ref RECT r, uint c);
    [DllImport("user32.dll")] static extern IntPtr GetShellWindow();
    [DllImport("user32.dll")] static extern IntPtr GetDesktopWindow();
    [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

    const uint SPI_GETWORKAREA=0x0030,SWP_NOSIZE=1,SWP_NOMOVE=2,SWP_NOZORDER=4,SWP_SHOWWINDOW=0x40,SWP_FRAMECHANGED=0x20;
    const int SW_RESTORE=9;
    static readonly IntPtr HWND_TOPMOST=new IntPtr(-1),HWND_NOTOPMOST=new IntPtr(-2);

    [StructLayout(LayoutKind.Sequential)]
    struct RECT { public int Left,Top,Right,Bottom; }

    static readonly HashSet<string> TC=new HashSet<string>(StringComparer.OrdinalIgnoreCase){
        "ConsoleWindowClass","CASCADIA_HOSTING_WINDOW_CLASS","VirtualConsoleClass"};
    static readonly HashSet<string> TP=new HashSet<string>(StringComparer.OrdinalIgnoreCase){
        "WindowsTerminal","alacritty","hyper","mintty",
        "ConEmu","ConEmu64","ConEmuC","ConEmuC64","cmder",
        "FluentTerminal","tabby","terminus","wezterm-gui"};

    static List<long> FindTerminals(int selfPid, long exHwnd) {
        var deny=new HashSet<int>{selfPid};
        var found=new List<long>();
        var shell=GetShellWindow(); var desk=GetDesktopWindow();
        var names=new Dictionary<int,string>();
        EnumWindows((h,lp)=>{
            if(!IsWindowVisible(h)||h==shell||h==desk) return true;
            var t=new StringBuilder(256); if(GetWindowText(h,t,256)==0) return true;
            uint u; GetWindowThreadProcessId(h,out u); int pid=(int)u;
            if(deny.Contains(pid)) return true;
            long hl=h.ToInt64(); if(exHwnd!=0&&hl==exHwnd) return true;
            var c=new StringBuilder(128); GetClassName(h,c,128);
            if(TC.Contains(c.ToString())){ found.Add(hl); return true; }
            if(!names.ContainsKey(pid)){
                try{ names[pid]=Process.GetProcessById(pid).ProcessName; }catch{ names[pid]=""; }
            }
            if(TP.Contains(names[pid])) found.Add(hl);
            return true;
        },IntPtr.Zero);
        return found;
    }

    static void RaiseWindow(IntPtr h){
        uint f=SWP_NOMOVE|SWP_NOSIZE|SWP_SHOWWINDOW;
        SetWindowPos(h,HWND_TOPMOST,0,0,0,0,f);
        SetWindowPos(h,HWND_NOTOPMOST,0,0,0,0,f);
    }

    static RECT GetWorkArea(){ var r=new RECT(); SystemParametersInfo(SPI_GETWORKAREA,0,ref r,0); return r; }

    // Cache: simple JSON {"cols":N,"rows":N,"slots":{"hwnd":slot,...}}
    static Dictionary<long,int> LoadCache(string path, out int cols, out int rows){
        cols=0; rows=0; var t=new Dictionary<long,int>();
        if(!File.Exists(path)) return t;
        try{
            var s=File.ReadAllText(path);
            cols=ParseInt(s,"\"cols\":"); rows=ParseInt(s,"\"rows\":");
            int si=s.IndexOf("\"slots\":"); if(si<0) return t;
            int lb=s.IndexOf('{',si+8); int rb=s.IndexOf('}',lb); if(lb<0||rb<0) return t;
            foreach(var e in s.Substring(lb+1,rb-lb-1).Split(',')){
                var p=e.Split(':'); if(p.Length!=2) continue;
                long k; int v;
                if(long.TryParse(p[0].Trim().Trim('"'),out k)&&int.TryParse(p[1].Trim(),out v)) t[k]=v;
            }
        }catch{}
        return t;
    }
    static int ParseInt(string s, string key){
        int i=s.IndexOf(key); if(i<0) return 0; i+=key.Length;
        while(i<s.Length&&!char.IsDigit(s[i]))i++;
        var b=new StringBuilder(); while(i<s.Length&&char.IsDigit(s[i]))b.Append(s[i++]);
        int v; return int.TryParse(b.ToString(),out v)?v:0;
    }
    static void SaveCache(string path, int cols, int rows, List<long> slotMap){
        var b=new StringBuilder();
        b.Append("{\"cols\":"); b.Append(cols);
        b.Append(",\"rows\":"); b.Append(rows);
        b.Append(",\"slots\":{");
        for(int i=0;i<slotMap.Count;i++){
            if(i>0)b.Append(',');
            b.Append('"'); b.Append(slotMap[i]); b.Append("\":"); b.Append(i);
        }
        b.Append("}}");
        File.WriteAllText(path,b.ToString(),System.Text.Encoding.UTF8);
    }

    static void Main(string[] args) {
        // Read handoff file written by tile-terminals.ps1 (callerHwnd + cols)
        long callerHwnd = 0;
        int cols = 0;
        string handoff = Path.Combine(Path.GetTempPath(), "TileTerminals_handoff.tmp");
        if(File.Exists(handoff)){
            try{
                var lines=File.ReadAllLines(handoff);
                if(lines.Length>0) long.TryParse(lines[0].Trim(),out callerHwnd);
                if(lines.Length>1) int.TryParse(lines[1].Trim(),out cols);
                File.Delete(handoff);
            }catch{}
        }
        // If run directly (e.g. from shortcut/task with no handoff), use foreground window
        if(callerHwnd==0) callerHwnd=GetForegroundWindow().ToInt64();

        int selfPid=Process.GetCurrentProcess().Id;
        var found=FindTerminals(selfPid,callerHwnd);
        if(found.Count==0) return;

        int n=found.Count;
        int numCols=cols>0?Math.Min(cols,n):(int)Math.Ceiling(Math.Sqrt((double)n));
        int numRows=(int)Math.Ceiling((double)n/numCols);
        var work=GetWorkArea();
        int cellW=(work.Right-work.Left)/numCols;
        int cellH=(work.Bottom-work.Top)/numRows;

        string cacheFile=Path.Combine(Path.GetTempPath(),"TileTerminals_cache.json");
        int cachedCols,cachedRows;
        var slotTable=LoadCache(cacheFile,out cachedCols,out cachedRows);
        bool gridChanged=slotTable.Count==0||cachedCols!=numCols||cachedRows!=numRows;

        // slotMap[slot]=hwnd
        var slotMap=new List<long>(new long[n]);
        var placed=new bool[n];

        if(!gridChanged){
            var used=new HashSet<int>();
            var unplaced=new List<int>();
            for(int i=0;i<n;i++){
                int slot;
                if(slotTable.TryGetValue(found[i],out slot)&&slot<n&&!used.Contains(slot)){
                    slotMap[slot]=found[i]; placed[slot]=true; used.Add(slot);
                } else { unplaced.Add(i); }
            }
            int ei=0;
            for(int s=0;s<n;s++){
                if(!placed[s]&&ei<unplaced.Count){ slotMap[s]=found[unplaced[ei++]]; placed[s]=true; }
            }
        } else {
            for(int i=0;i<n;i++) slotMap[i]=found[i];
        }

        SaveCache(cacheFile,numCols,numRows,slotMap);

        uint pf=SWP_NOZORDER|SWP_SHOWWINDOW|SWP_FRAMECHANGED;
        for(int i=0;i<n;i++){
            var hwnd=new IntPtr(slotMap[i]);
            int x=work.Left+(i%numCols)*cellW;
            int y=work.Top+(i/numCols)*cellH;
            ShowWindow(hwnd,SW_RESTORE);
            SetWindowPos(hwnd,IntPtr.Zero,x,y,cellW,cellH,pf);
            RaiseWindow(hwnd);
        }
        if(callerHwnd!=0){
            Thread.Sleep(80);
            RaiseWindow(new IntPtr(callerHwnd));
        }
    }
}
'@

try {
    Add-Type -TypeDefinition $csharp `
        -OutputAssembly $ExePath `
        -OutputType ConsoleApplication `
        -ReferencedAssemblies 'System','System.Core'
    Write-Host "Compiled: $ExePath" -ForegroundColor Green
} catch {
    Write-Error "Compilation failed: $_"
    exit 1
}

# ---------------------------------------------------------------------------
# Register scheduled task pointing at the compiled exe
# ---------------------------------------------------------------------------
$currentUser = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name

$action = New-ScheduledTaskAction `
    -Execute  $ExePath `
    -WorkingDirectory $ScriptDir

$principal = New-ScheduledTaskPrincipal `
    -UserId    $currentUser `
    -LogonType Interactive `
    -RunLevel  Highest

$settings = New-ScheduledTaskSettingsSet `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit ([TimeSpan]::Zero) `
    -MultipleInstances IgnoreNew

Register-ScheduledTask `
    -TaskName    $TaskName `
    -Description $Desc `
    -Action      $action `
    -Principal   $principal `
    -Settings    $settings `
    -Force | Out-Null

Write-Host "Scheduled task '$TaskName' registered." -ForegroundColor Green

# Desktop shortcut
$shell    = New-Object -ComObject WScript.Shell
$shortcut = $shell.CreateShortcut($ShortcutPath)
$shortcut.TargetPath        = "$env:SystemRoot\System32\schtasks.exe"
$shortcut.Arguments         = "/run /tn `"$TaskName`""
$shortcut.Description       = $Desc
$shortcut.IconLocation      = "$ExePath,0"
$shortcut.WindowStyle       = 7  # minimized (no console flash)
$shortcut.Save()

Write-Host "Shortcut created:  $ShortcutPath" -ForegroundColor Green
Write-Host ""
Write-Host "Run this script again to uninstall." -ForegroundColor DarkGray