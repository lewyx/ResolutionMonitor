# ResolutionMonitor.ps1 - Resolution Monitor system tray tool

$applicationVersion = "1.0"
$applicationDeveloper = "PanSoft"
$applicationShortName = "Resolution Monitor"
$applicationFullName = "$applicationDeveloper $applicationShortName v$applicationVersion"

# Registry path for settings
$script:SettingsRegPath = "HKCU:\Software\$applicationDeveloper\$applicationShortName"

# Default target settings
$script:TargetResolutions = @(
    @{Width = 1024; Height = 768},
    @{Width = 1152; Height = 864},
    @{Width = 1280; Height = 720},
    @{Width = 1280; Height = 768},
    @{Width = 1280; Height = 800},
    @{Width = 1280; Height = 1024},
    @{Width = 1360; Height = 768},
    @{Width = 1366; Height = 768},
    @{Width = 1440; Height = 900},
    @{Width = 1600; Height = 900},
    @{Width = 1600; Height = 1200},
    @{Width = 1680; Height = 1050},
    @{Width = 1920; Height = 1080},
    @{Width = 1920; Height = 1200},
    @{Width = 2048; Height = 1152},
    @{Width = 2560; Height = 1440},
    @{Width = 2560; Height = 1600},
    @{Width = 3200; Height = 1800},
    @{Width = 3840; Height = 2160},
    @{Width = 5120; Height = 2880},
    @{Width = 7680; Height = 4320}
)
$script:TargetScalings = @(100, 125, 150, 175, 200)

function ReadSetting($name, $default) {
    try {
        $value = Get-ItemProperty -Path $script:SettingsRegPath -Name $name -ErrorAction SilentlyContinue
        if ($value -and $value.$name) { return $value.$name }
    } catch {}
    return $default
}

function ReadSettings($prefix) {
    return @{ Width = ReadSetting "${prefix}Width" 1080; Height = ReadSetting "${prefix}Height" 1920; Scaling = ReadSetting "${prefix}Scaling" 100 } #FullHD in portrait mode
}

function WriteSetting($name, $value) {
    Set-ItemProperty -Path $script:SettingsRegPath -Name $name -Value $value
}

function WriteSettings($prefix, $value) {
    if (-not (Test-Path $script:SettingsRegPath)) {
        New-Item -Path $script:SettingsRegPath -Force | Out-Null
    }
    WriteSetting "${prefix}Width" ([int]$value.Width)
    WriteSetting "${prefix}Height" ([int]$value.Height)
    WriteSetting "${prefix}Scaling" ([int]$value.Scaling)
}

function ReloadTargetState {
    $script:TargetState = @{ Width = 1920; Height = 1080; Scaling = 100 }

    # Load TargetState from registry
    $loaded = ReadSettings "Target"
    if ($loaded) {
        $script:TargetState = $loaded
    }
    $script:menuApplyTarget.Text = "Apply $($script:TargetState.Scaling)% $($script:TargetState.Width)x$($script:TargetState.Height)"

    # Load custom tag values from registry
    $customSetting = ReadSettings "Custom0"

    # Update resolution menu checkmarks
    $found = $false
    foreach ($item in $script:menuTargetRes.DropDownItems) {
        if ($item -isnot [System.Windows.Forms.ToolStripMenuItem]) { break }
        $item.Checked = ($item.Tag.Width -eq $script:TargetState.Width -and $item.Tag.Height -eq $script:TargetState.Height)
        if ($item.Checked) { $found = $true }
    }

    # If no standard resolution matched, check Custom and update its display
    if ($script:menuCustomRes.Checked = -not $found) {
        $customSetting.Width = $script:TargetState.Width
        $customSetting.Height = $script:TargetState.Height
    }
    $script:menuCustomRes.Tag = @{Width = $customSetting.Width; Height = $customSetting.Height}
    $script:menuCustomRes.Text = "Custom: $($script:TargetState.Width)x$($script:TargetState.Height)"

    # Update scaling menu checkmarks
    $found = $false
    foreach ($item in $script:menuTargetScale.DropDownItems) {
        if ($item -isnot [System.Windows.Forms.ToolStripMenuItem]) { break }
        $item.Checked = ($item.Tag -eq $script:TargetState.Scaling)
        if ($item.Checked) { $found = $true }
    }

    # If no standard scaling matched, check Custom and update its display
    if ($script:menuCustomScale.Checked = -not $found) {
        $script:menuCustomScale.Tag = $script:TargetState.Scaling
    }
    $script:menuCustomScale.Tag = $customSetting.Scaling
    $script:menuCustomScale.Text = "Custom: $($customSetting.Scaling)%"

    Refresh
}

function ChangeTargetState {
    WriteSettings "Target" $script:TargetState
    WriteSettings "Custom0" ($script:menuCustomRes.Tag + @{Scaling = $script:menuCustomScale.Tag})
    ReloadTargetState
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

Add-Type -TypeDefinition @"
using System;
using System.Drawing;
using System.Runtime.InteropServices;

public class DisplayHelper
{
    // ---- Structs ----

    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
    public struct DEVMODE
    {
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmDeviceName;
        public short dmSpecVersion;
        public short dmDriverVersion;
        public short dmSize;
        public short dmDriverExtra;
        public int dmFields;
        public int dmPositionX;
        public int dmPositionY;
        public int dmDisplayOrientation;
        public int dmDisplayFixedOutput;
        public short dmColor;
        public short dmDuplex;
        public short dmYResolution;
        public short dmTTOption;
        public short dmCollate;
        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]
        public string dmFormName;
        public short dmLogPixels;
        public int dmBitsPerPel;
        public int dmPelsWidth;
        public int dmPelsHeight;
        public int dmDisplayFlags;
        public int dmDisplayFrequency;
        public int dmICMMethod;
        public int dmICMIntent;
        public int dmMediaType;
        public int dmDitherType;
        public int dmReserved1;
        public int dmReserved2;
        public int dmPanningWidth;
        public int dmPanningHeight;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct LUID { public uint LowPart; public int HighPart; }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_RATIONAL { public uint Numerator; public uint Denominator; }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_SOURCE_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_TARGET_INFO
    {
        public LUID adapterId;
        public uint id;
        public uint modeInfoIdx;
        public int outputTechnology;
        public int rotation;
        public int scaling;
        public DISPLAYCONFIG_RATIONAL refreshRate;
        public int scanLineOrdering;
        public int targetAvailable;
        public uint statusFlags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_PATH_INFO
    {
        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;
        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;
        public uint flags;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_MODE_INFO
    {
        public int infoType;
        public uint id;
        public LUID adapterId;
        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 48)]
        public byte[] modeData;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_DEVICE_INFO_HEADER
    {
        public int type;
        public int size;
        public LUID adapterId;
        public uint id;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_GET
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int minScaleRel;
        public int curScaleRel;
        public int maxScaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_SET
    {
        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;
        public int scaleRel;
    }

    [StructLayout(LayoutKind.Sequential)]
    public struct POINT { public int x, y; }

    // ---- Imports ----

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);

    [DllImport("user32.dll", CharSet = CharSet.Auto)]
    public static extern int ChangeDisplaySettingsEx(string deviceName, ref DEVMODE devMode, IntPtr hwnd, uint flags, IntPtr lParam);

    [DllImport("shcore.dll")]
    public static extern int SetProcessDpiAwareness(int awareness);

    [DllImport("shcore.dll")]
    public static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);

    [DllImport("user32.dll")]
    public static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);

    [DllImport("user32.dll")]
    public static extern IntPtr GetDC(IntPtr hwnd);

    [DllImport("user32.dll")]
    public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);

    [DllImport("gdi32.dll")]
    public static extern int GetDeviceCaps(IntPtr hdc, int index);

    [DllImport("shell32.dll", CharSet = CharSet.Auto)]
    public static extern IntPtr ExtractIcon(IntPtr hInst, string file, int index);

    [DllImport("user32.dll")]
    public static extern bool DestroyIcon(IntPtr hIcon);

    [DllImport("user32.dll")]
    public static extern int GetDisplayConfigBufferSizes(uint flags, out uint pathCount, out uint modeCount);

    [DllImport("user32.dll")]
    public static extern int QueryDisplayConfig(uint flags,
        ref uint pathCount, [Out] DISPLAYCONFIG_PATH_INFO[] paths,
        ref uint modeCount, [Out] DISPLAYCONFIG_MODE_INFO[] modes,
        IntPtr topologyId);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_GET info);

    [DllImport("user32.dll")]
    public static extern int DisplayConfigSetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_SET info);

    // ---- Constants ----

    public const int ENUM_CURRENT_SETTINGS = -1;
    public const uint QDC_ONLY_ACTIVE_PATHS = 2;
    public const int DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE = -3;
    public const int DISPLAYCONFIG_DEVICE_INFO_SET_DPI_SCALE = -4;
    public const int DM_PELSWIDTH = 0x80000;
    public const int DM_PELSHEIGHT = 0x100000;
    public const uint MONITOR_DEFAULTTOPRIMARY = 1;
    public const int MDT_EFFECTIVE_DPI = 0;

    public static readonly int[] DpiSteps = { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };

    // ---- Public API ----

    public static void GetResolution(out int width, out int height)
    {
        var dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        width = dm.dmPelsWidth;
        height = dm.dmPelsHeight;
    }

    public static int GetScaling()
    {
        // Try per-monitor DPI via GetDpiForMonitor
        var pt = new POINT { x = 0, y = 0 };
        IntPtr hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);
        uint dpiX, dpiY;
        if (GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI, out dpiX, out dpiY) == 0 && dpiX > 0)
            return (int)Math.Round(dpiX / 96.0 * 100);

        // Fallback: system DPI via GDI
        IntPtr hdc = GetDC(IntPtr.Zero);
        int logPixels = GetDeviceCaps(hdc, 88); // LOGPIXELSX
        ReleaseDC(IntPtr.Zero, hdc);
        return logPixels > 0 ? (int)Math.Round(logPixels / 96.0 * 100) : 100;
    }

    public static bool SetResolution(int width, int height)
    {
        var dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        dm.dmPelsWidth = width;
        dm.dmPelsHeight = height;
        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
        return ChangeDisplaySettingsEx(null, ref dm, IntPtr.Zero, 0x01, IntPtr.Zero) == 0;
    }

    public static bool SetScaling(int targetPercent)
    {
        // Get active display paths
        uint pathCount, modeCount;
        if (GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out pathCount, out modeCount) != 0)
            return false;
        var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
        var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];
        if (QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero) != 0)
            return false;
        if (pathCount == 0) return false;

        // Find built-in display, fall back to first path
        int idx = 0;
        for (int i = 0; i < (int)pathCount; i++)
        {
            int tech = paths[i].targetInfo.outputTechnology;
            if (tech == unchecked((int)0x80000000) || tech == 11) // INTERNAL or DP_EMBEDDED
            { idx = i; break; }
        }
        var path = paths[idx];

        // Get current DPI scale info
        var get = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        get.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE;
        get.header.size = Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        get.header.adapterId = path.sourceInfo.adapterId;
        get.header.id = path.sourceInfo.id;
        if (DisplayConfigGetDeviceInfo(ref get) != 0) return false;

        // Compute target relative scale
        int curPercent = GetScaling();
        int curAbsIdx = FindClosest(curPercent);
        int recAbsIdx = curAbsIdx - get.curScaleRel;
        int targetAbsIdx = FindClosest(targetPercent);
        int targetRel = targetAbsIdx - recAbsIdx;

        if (targetRel < get.minScaleRel || targetRel > get.maxScaleRel)
            return false;

        // Apply
        var set = new DISPLAYCONFIG_SOURCE_DPI_SCALE_SET();
        set.header.type = DISPLAYCONFIG_DEVICE_INFO_SET_DPI_SCALE;
        set.header.size = Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_SET));
        set.header.adapterId = path.sourceInfo.adapterId;
        set.header.id = path.sourceInfo.id;
        set.scaleRel = targetRel;
        return DisplayConfigSetDeviceInfo(ref set) == 0;
    }

    public static Icon GetShellIcon(int index)
    {
        IntPtr h = ExtractIcon(IntPtr.Zero, "shell32.dll", index);
        if (h == IntPtr.Zero || h == (IntPtr)1)
            return SystemIcons.Application;
        Icon ico = (Icon)Icon.FromHandle(h).Clone();
        DestroyIcon(h);
        return ico;
    }

    private static int FindClosest(int percent)
    {
        int best = 0, bestDiff = Math.Abs(DpiSteps[0] - percent);
        for (int i = 1; i < DpiSteps.Length; i++)
        {
            int diff = Math.Abs(DpiSteps[i] - percent);
            if (diff < bestDiff) { best = i; bestDiff = diff; }
        }
        return best;
    }
}
"@ -ReferencedAssemblies System.Drawing

# Make process per-monitor DPI aware (best effort)
try { [DisplayHelper]::SetProcessDpiAwareness(2) } catch {}

# ---- Icons ----
$iconMonitor = [DisplayHelper]::GetShellIcon(15)   # computer/monitor
$iconWarning = [DisplayHelper]::GetShellIcon(77)   # yellow warning triangle

# ---- UI Setup ----
$form = New-Object System.Windows.Forms.Form
$form.ShowInTaskbar = $false
$form.WindowState = 'Minimized'
$form.Visible = $false
$form.FormBorderStyle = 'None'
$form.Size = New-Object System.Drawing.Size(0, 0)
$form.Opacity = 0
$form.Add_FormClosing({ param($s, $e) $e.Cancel = $true })

# Hook into display change messages
Add-Type -TypeDefinition @"
using System;
using System.Windows.Forms;

public class MessageWindow : NativeWindow {
    public event EventHandler DisplayChanged;

    protected override void WndProc(ref Message m) {
        const int WM_DISPLAYCHANGE = 0x007E;
        const int WM_DPICHANGED = 0x02E0;
        const int WM_SETTINGCHANGE = 0x001A;

        base.WndProc(ref m);
        if ((DisplayChanged != null) && (m.Msg == WM_DISPLAYCHANGE || m.Msg == WM_DPICHANGED || m.Msg == WM_SETTINGCHANGE)) {
            DisplayChanged(this, EventArgs.Empty);
        }
    }
}
"@ -ReferencedAssemblies System.Windows.Forms

$msgWindow = New-Object MessageWindow
$msgWindow.CreateHandle((New-Object System.Windows.Forms.CreateParams))
$msgWindow.Add_DisplayChanged({ Refresh })

$notifyIcon = New-Object System.Windows.Forms.NotifyIcon
$notifyIcon.Visible = $true
$notifyIcon.Text = $applicationFullName

$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip
$script:menuApplyTarget = $contextMenu.Items.Add("Apply Target")
$menuRefresh     = $contextMenu.Items.Add("Refresh")
$contextMenu.Items.Add("-")  # separator

# Options submenu
$menuOptions = New-Object System.Windows.Forms.ToolStripMenuItem("Options")

# Target resolution submenu
$script:menuTargetRes = New-Object System.Windows.Forms.ToolStripMenuItem("Target resolution")
foreach ($res in $script:TargetResolutions) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem("$($res.Width) x $($res.Height)")
    $item.Tag = $res
    $item.CheckOnClick = $true
    $item.Add_Click({
        param($sender, $e)
        $script:TargetState = $sender.Tag + @{Scaling = $script:TargetState.Scaling}
        ChangeTargetState
    })
    $script:menuTargetRes.DropDownItems.Add($item) | Out-Null
}

$script:menuTargetRes.DropDownItems.Add("-") | Out-Null
# Custom resolution menu item
Add-Type -AssemblyName Microsoft.VisualBasic
$script:menuCustomRes = New-Object System.Windows.Forms.ToolStripMenuItem("Custom")
$script:menuCustomRes.CheckOnClick = $true
$script:menuCustomRes.Add_Click({
    param($sender, $e)
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter resolution (e.g. 1920x1080)", "Custom Resolution", "$($sender.Tag.Width) x $($sender.Tag.Height)")
    if ($input) {
        $numbers = ($input -replace '[^\d]', ' ').Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
        if ($numbers.Count -ge 2) {
            $sender.Tag = @{Width = [int]$numbers[0]; Height = [int]$numbers[-1]}
            $script:TargetState = $sender.Tag + @{Scaling = $script:TargetState.Scaling}
            ChangeTargetState
        }
    }
})
$script:menuTargetRes.DropDownItems.Add($script:menuCustomRes) | Out-Null
$menuOptions.DropDownItems.Add($script:menuTargetRes) | Out-Null

# Target scaling submenu
$script:menuTargetScale = New-Object System.Windows.Forms.ToolStripMenuItem("Target scaling")
foreach ($scale in $script:TargetScalings) {
    $item = New-Object System.Windows.Forms.ToolStripMenuItem("$scale%")
    $item.Tag = $scale
    $item.CheckOnClick = $true
    $item.Add_Click({
        param($sender, $e)
        $script:TargetState.Scaling = $sender.Tag
        ChangeTargetState
    })
    $script:menuTargetScale.DropDownItems.Add($item) | Out-Null
}

$script:menuTargetScale.DropDownItems.Add("-") | Out-Null #separator
# Custom scaling menu item
$script:menuCustomScale = New-Object System.Windows.Forms.ToolStripMenuItem("Custom")
$script:menuCustomScale.Tag = 200
$script:menuCustomScale.CheckOnClick = $true
$script:menuCustomScale.Add_Click({
    param($sender, $e)
    $input = [Microsoft.VisualBasic.Interaction]::InputBox("Enter scaling percentage (e.g. 100)", "Custom Scaling", "$($sender.Tag)")
    if ($input) {
        $numbers = ($input -replace '[^\d]', ' ').Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)
        if ($numbers.Count -ge 1) {
            $sender.Tag = [int]$numbers[0]
            $script:TargetState.Scaling = $sender.Tag
            ChangeTargetState
        }
    }
})
$script:menuTargetScale.DropDownItems.Add($script:menuCustomScale) | Out-Null
$menuOptions.DropDownItems.Add($script:menuTargetScale) | Out-Null

$menuOptions.DropDownItems.Add("-")  # separator
# Auto start option
$menuAutoStart = New-Object System.Windows.Forms.ToolStripMenuItem("Auto start")
$menuAutoStart.CheckOnClick = $true
$menuOptions.DropDownItems.Add($menuAutoStart) | Out-Null

$contextMenu.Items.Add($menuOptions) | Out-Null
$contextMenu.Items.Add("-")  # separator
$menuExit = $contextMenu.Items.Add("Exit")
$notifyIcon.ContextMenuStrip = $contextMenu

# ---- Auto-start Logic ----
$AutoStartRegPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Run"
$AutoStartValueName = $applicationFullName
$VbsPath = Join-Path (Split-Path $PSCommandPath -Parent) "ResolutionMonitor.vbs"

# Update auto-start checkbox before menu opens
$contextMenu.Add_Opening({
    try {
        $menuAutoStart.Checked = $null -ne (Get-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -ErrorAction SilentlyContinue)
    } catch {
        $menuAutoStart.Checked = $false
    }
})

# ---- Refresh Logic ----
function Refresh {
    $w = 0; $h = 0
    [DisplayHelper]::GetResolution([ref]$w, [ref]$h)
    $s = [DisplayHelper]::GetScaling()

    if ($w -eq $script:TargetState.Width -and $h -eq $script:TargetState.Height -and $s -eq $script:TargetState.Scaling) {
        $notifyIcon.Icon = $iconMonitor
    } else {
        $notifyIcon.Icon = $iconWarning
    }

    $old = $notifyIcon.Text
    $notifyIcon.Text = "${s}% @ ${w}x${h}"
    if (($old -ne $applicationFullName) -and ($old -ne $notifyIcon.Text)) {
        $notifyIcon.ShowBalloonTip(5000, "Resolution Monitor", $notifyIcon.Text, [System.Windows.Forms.ToolTipIcon]::Info)
    }
}

# ---- Event Handlers ----
$script:menuApplyTarget.Add_Click({
    [DisplayHelper]::SetResolution($script:TargetState.Width, $script:TargetState.Height) | Out-Null
    [DisplayHelper]::SetScaling($script:TargetState.Scaling) | Out-Null
    Refresh
})

$menuRefresh.Add_Click({
    Refresh
})

$menuAutoStart.Add_Click({
    if ($menuAutoStart.Checked) {
        try {
            Set-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -Value "wscript.exe `"$VbsPath`""
        } catch {
            $menuAutoStart.Checked = $false
        }
    } else {
        try {
            Remove-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -ErrorAction SilentlyContinue
        } catch {
            $menuAutoStart.Checked = $true
        }
    }
})

$menuExit.Add_Click({
    $timer.Stop()
    $notifyIcon.Visible = $false
    $notifyIcon.Dispose()
    [System.Environment]::Exit(0)
})

$notifyIcon.Add_Click({
    param($sender, $e)
    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $form.Activate()
        $contextMenu.Show([System.Windows.Forms.Cursor]::Position)
    }
})

# ---- Timer ----
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 30000
$timer.Add_Tick({ Refresh })

# ---- Start ----
ReloadTargetState
$timer.Start()
[System.Windows.Forms.Application]::Run($form)
