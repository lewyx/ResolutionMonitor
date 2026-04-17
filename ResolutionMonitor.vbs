Dim fso, tmp, f
Set fso = CreateObject("Scripting.FileSystemObject")
tmp = fso.BuildPath(fso.GetSpecialFolder(2), "ResMonitor_" & fso.GetTempName & ".ps1")
Set f = fso.CreateTextFile(tmp, True)
f.Write "# ResolutionMonitor.ps1 - Resolution Monitor system tray tool" & vbCrLf & _
    "" & vbCrLf & _
    "$applicationVersion = ""1.0""" & vbCrLf & _
    "$applicationDeveloper = ""PanSoft""" & vbCrLf & _
    "$applicationShortName = ""Resolution Monitor""" & vbCrLf & _
    "$applicationFullName = ""$applicationDeveloper $applicationShortName v$applicationVersion""" & vbCrLf & _
    "" & vbCrLf & _
    "# Registry path for settings" & vbCrLf & _
    "$script:SettingsRegPath = ""HKCU:\Software\$applicationDeveloper\$applicationShortName""" & vbCrLf & _
    "" & vbCrLf & _
    "# Default target settings" & vbCrLf & _
    "$script:TargetResolutions = @(" & vbCrLf & _
    "    @{Width = 1024; Height = 768}," & vbCrLf & _
    "    @{Width = 1152; Height = 864}," & vbCrLf & _
    "    @{Width = 1280; Height = 720}," & vbCrLf & _
    "    @{Width = 1280; Height = 768}," & vbCrLf & _
    "    @{Width = 1280; Height = 800}," & vbCrLf & _
    "    @{Width = 1280; Height = 1024}," & vbCrLf & _
    "    @{Width = 1360; Height = 768}," & vbCrLf & _
    "    @{Width = 1366; Height = 768}," & vbCrLf & _
    "    @{Width = 1440; Height = 900}," & vbCrLf & _
    "    @{Width = 1600; Height = 900}," & vbCrLf & _
    "    @{Width = 1600; Height = 1200}," & vbCrLf & _
    "    @{Width = 1680; Height = 1050}," & vbCrLf & _
    "    @{Width = 1920; Height = 1080}," & vbCrLf & _
    "    @{Width = 1920; Height = 1200}," & vbCrLf & _
    "    @{Width = 2048; Height = 1152}," & vbCrLf & _
    "    @{Width = 2560; Height = 1440}," & vbCrLf & _
    "    @{Width = 2560; Height = 1600}," & vbCrLf & _
    "    @{Width = 3200; Height = 1800}," & vbCrLf & _
    "    @{Width = 3840; Height = 2160}," & vbCrLf & _
    "    @{Width = 5120; Height = 2880}," & vbCrLf & _
    "    @{Width = 7680; Height = 4320}" & vbCrLf & _
    ")" & vbCrLf & _
    "$script:TargetScalings = @(100, 125, 150, 175, 200)" & vbCrLf & _
    "" & vbCrLf & _
    "function ReadSetting($name, $default) {" & vbCrLf & _
    "    try {" & vbCrLf & _
    "        $value = Get-ItemProperty -Path $script:SettingsRegPath -Name $name -ErrorAction SilentlyContinue" & vbCrLf & _
    "        if ($value -and $value.$name) { return $value.$name }" & vbCrLf & _
    "    } catch {}" & vbCrLf & _
    "    return $default" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "function ReadSettings($prefix) {" & vbCrLf & _
    "    return @{ Width = ReadSetting ""${prefix}Width"" 1080; Height = ReadSetting ""${prefix}Height"" 1920; Scaling = ReadSetting ""${prefix}Scaling"" 100 } #FullHD in portrait mode" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "function WriteSetting($name, $value) {" & vbCrLf & _
    "    Set-ItemProperty -Path $script:SettingsRegPath -Name $name -Value $value" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "function WriteSettings($prefix, $value) {" & vbCrLf & _
    "    if (-not (Test-Path $script:SettingsRegPath)) {" & vbCrLf & _
    "        New-Item -Path $script:SettingsRegPath -Force | Out-Null" & vbCrLf & _
    "    }" & vbCrLf & _
    "    WriteSetting ""${prefix}Width"" ([int]$value.Width)" & vbCrLf & _
    "    WriteSetting ""${prefix}Height"" ([int]$value.Height)" & vbCrLf & _
    "    WriteSetting ""${prefix}Scaling"" ([int]$value.Scaling)" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "function ReloadTargetState {" & vbCrLf & _
    "    $script:TargetState = @{ Width = 1920; Height = 1080; Scaling = 100 }" & vbCrLf & _
    "" & vbCrLf & _
    "    # Load TargetState from registry" & vbCrLf & _
    "    $loaded = ReadSettings ""Target""" & vbCrLf & _
    "    if ($loaded) {" & vbCrLf & _
    "        $script:TargetState = $loaded" & vbCrLf & _
    "    }" & vbCrLf & _
    "    $script:menuApplyTarget.Text = ""Apply $($script:TargetState.Scaling)% $($script:TargetState.Width)x$($script:TargetState.Height)""" & vbCrLf & _
    "" & vbCrLf & _
    "    # Load custom tag values from registry" & vbCrLf & _
    "    $customSetting = ReadSettings ""Custom0""" & vbCrLf & _
    "" & vbCrLf & _
    "    # Update resolution menu checkmarks" & vbCrLf & _
    "    $found = $false" & vbCrLf & _
    "    foreach ($item in $script:menuTargetRes.DropDownItems) {" & vbCrLf & _
    "        if ($item -isnot [System.Windows.Forms.ToolStripMenuItem]) { break }" & vbCrLf & _
    "        $item.Checked = ($item.Tag.Width -eq $script:TargetState.Width -and $item.Tag.Height -eq $script:TargetState.Height)" & vbCrLf & _
    "        if ($item.Checked) { $found = $true }" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    # If no standard resolution matched, check Custom and update its display" & vbCrLf & _
    "    if ($script:menuCustomRes.Checked = -not $found) { # intentional assignment" & vbCrLf & _
    "        $customSetting.Width = $script:TargetState.Width" & vbCrLf & _
    "        $customSetting.Height = $script:TargetState.Height" & vbCrLf & _
    "    }" & vbCrLf & _
    "    $script:menuCustomRes.Tag = @{Width = $customSetting.Width; Height = $customSetting.Height}" & vbCrLf & _
    "    $script:menuCustomRes.Text = ""Custom: $($script:TargetState.Width)x$($script:TargetState.Height)""" & vbCrLf & _
    "" & vbCrLf & _
    "    # Update scaling menu checkmarks" & vbCrLf & _
    "    $found = $false" & vbCrLf & _
    "    foreach ($item in $script:menuTargetScale.DropDownItems) {" & vbCrLf & _
    "        if ($item -isnot [System.Windows.Forms.ToolStripMenuItem]) { break }" & vbCrLf & _
    "        $item.Checked = ($item.Tag -eq $script:TargetState.Scaling)" & vbCrLf & _
    "        if ($item.Checked) { $found = $true }" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    # If no standard scaling matched, check Custom and update its display" & vbCrLf & _
    "    if ($script:menuCustomScale.Checked = -not $found) { # intentional assignment" & vbCrLf & _
    "        $customSetting.Scaling = $script:TargetState.Scaling" & vbCrLf & _
    "    }" & vbCrLf & _
    "    $script:menuCustomScale.Tag = $customSetting.Scaling" & vbCrLf & _
    "    $script:menuCustomScale.Text = ""Custom: $($customSetting.Scaling)%""" & vbCrLf & _
    "" & vbCrLf & _
    "    Refresh" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "function ChangeTargetState {" & vbCrLf & _
    "    WriteSettings ""Target"" $script:TargetState" & vbCrLf & _
    "    WriteSettings ""Custom0"" ($script:menuCustomRes.Tag + @{Scaling = $script:menuCustomScale.Tag})" & vbCrLf & _
    "    ReloadTargetState" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "Add-Type -AssemblyName System.Windows.Forms" & vbCrLf & _
    "Add-Type -AssemblyName System.Drawing" & vbCrLf & _
    "" & vbCrLf & _
    "Add-Type -TypeDefinition @'" & vbCrLf & _
    "using System;" & vbCrLf & _
    "using System.Drawing;" & vbCrLf & _
    "using System.Runtime.InteropServices;" & vbCrLf & _
    "" & vbCrLf & _
    "public class DisplayHelper" & vbCrLf & _
    "{" & vbCrLf & _
    "    // ---- Structs ----" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]" & vbCrLf & _
    "    public struct DEVMODE" & vbCrLf & _
    "    {" & vbCrLf & _
    "        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]" & vbCrLf & _
    "        public string dmDeviceName;" & vbCrLf & _
    "        public short dmSpecVersion;" & vbCrLf & _
    "        public short dmDriverVersion;" & vbCrLf & _
    "        public short dmSize;" & vbCrLf & _
    "        public short dmDriverExtra;" & vbCrLf & _
    "        public int dmFields;" & vbCrLf & _
    "        public int dmPositionX;" & vbCrLf & _
    "        public int dmPositionY;" & vbCrLf & _
    "        public int dmDisplayOrientation;" & vbCrLf & _
    "        public int dmDisplayFixedOutput;" & vbCrLf & _
    "        public short dmColor;" & vbCrLf & _
    "        public short dmDuplex;" & vbCrLf & _
    "        public short dmYResolution;" & vbCrLf & _
    "        public short dmTTOption;" & vbCrLf & _
    "        public short dmCollate;" & vbCrLf & _
    "        [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)]" & vbCrLf & _
    "        public string dmFormName;" & vbCrLf & _
    "        public short dmLogPixels;" & vbCrLf & _
    "        public int dmBitsPerPel;" & vbCrLf & _
    "        public int dmPelsWidth;" & vbCrLf & _
    "        public int dmPelsHeight;" & vbCrLf & _
    "        public int dmDisplayFlags;" & vbCrLf & _
    "        public int dmDisplayFrequency;" & vbCrLf & _
    "        public int dmICMMethod;" & vbCrLf & _
    "        public int dmICMIntent;" & vbCrLf & _
    "        public int dmMediaType;" & vbCrLf & _
    "        public int dmDitherType;" & vbCrLf & _
    "        public int dmReserved1;" & vbCrLf & _
    "        public int dmReserved2;" & vbCrLf & _
    "        public int dmPanningWidth;" & vbCrLf & _
    "        public int dmPanningHeight;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct LUID { public uint LowPart; public int HighPart; }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_RATIONAL { public uint Numerator; public uint Denominator; }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_PATH_SOURCE_INFO" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public LUID adapterId;" & vbCrLf & _
    "        public uint id;" & vbCrLf & _
    "        public uint modeInfoIdx;" & vbCrLf & _
    "        public uint statusFlags;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_PATH_TARGET_INFO" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public LUID adapterId;" & vbCrLf & _
    "        public uint id;" & vbCrLf & _
    "        public uint modeInfoIdx;" & vbCrLf & _
    "        public int outputTechnology;" & vbCrLf & _
    "        public int rotation;" & vbCrLf & _
    "        public int scaling;" & vbCrLf & _
    "        public DISPLAYCONFIG_RATIONAL refreshRate;" & vbCrLf & _
    "        public int scanLineOrdering;" & vbCrLf & _
    "        public int targetAvailable;" & vbCrLf & _
    "        public uint statusFlags;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_PATH_INFO" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public DISPLAYCONFIG_PATH_SOURCE_INFO sourceInfo;" & vbCrLf & _
    "        public DISPLAYCONFIG_PATH_TARGET_INFO targetInfo;" & vbCrLf & _
    "        public uint flags;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_MODE_INFO" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public int infoType;" & vbCrLf & _
    "        public uint id;" & vbCrLf & _
    "        public LUID adapterId;" & vbCrLf & _
    "        [MarshalAs(UnmanagedType.ByValArray, SizeConst = 48)]" & vbCrLf & _
    "        public byte[] modeData;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_DEVICE_INFO_HEADER" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public int type;" & vbCrLf & _
    "        public int size;" & vbCrLf & _
    "        public LUID adapterId;" & vbCrLf & _
    "        public uint id;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_GET" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;" & vbCrLf & _
    "        public int minScaleRel;" & vbCrLf & _
    "        public int curScaleRel;" & vbCrLf & _
    "        public int maxScaleRel;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct DISPLAYCONFIG_SOURCE_DPI_SCALE_SET" & vbCrLf & _
    "    {" & vbCrLf & _
    "        public DISPLAYCONFIG_DEVICE_INFO_HEADER header;" & vbCrLf & _
    "        public int scaleRel;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    [StructLayout(LayoutKind.Sequential)]" & vbCrLf & _
    "    public struct POINT { public int x, y; }" & vbCrLf & _
    "" & vbCrLf & _
    "    // ---- Imports ----" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"", CharSet = CharSet.Auto)]" & vbCrLf & _
    "    public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"", CharSet = CharSet.Auto)]" & vbCrLf & _
    "    public static extern int ChangeDisplaySettingsEx(string deviceName, ref DEVMODE devMode, IntPtr hwnd, uint flags, IntPtr lParam);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""shcore.dll"")]" & vbCrLf & _
    "    public static extern int SetProcessDpiAwareness(int awareness);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""shcore.dll"")]" & vbCrLf & _
    "    public static extern int GetDpiForMonitor(IntPtr hmonitor, int dpiType, out uint dpiX, out uint dpiY);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern IntPtr MonitorFromPoint(POINT pt, uint dwFlags);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern IntPtr GetDC(IntPtr hwnd);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern int ReleaseDC(IntPtr hwnd, IntPtr hdc);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""gdi32.dll"")]" & vbCrLf & _
    "    public static extern int GetDeviceCaps(IntPtr hdc, int index);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""shell32.dll"", CharSet = CharSet.Auto)]" & vbCrLf & _
    "    public static extern IntPtr ExtractIcon(IntPtr hInst, string file, int index);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern bool DestroyIcon(IntPtr hIcon);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern int GetDisplayConfigBufferSizes(uint flags, out uint pathCount, out uint modeCount);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern int QueryDisplayConfig(uint flags," & vbCrLf & _
    "        ref uint pathCount, [Out] DISPLAYCONFIG_PATH_INFO[] paths," & vbCrLf & _
    "        ref uint modeCount, [Out] DISPLAYCONFIG_MODE_INFO[] modes," & vbCrLf & _
    "        IntPtr topologyId);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern int DisplayConfigGetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_GET info);" & vbCrLf & _
    "" & vbCrLf & _
    "    [DllImport(""user32.dll"")]" & vbCrLf & _
    "    public static extern int DisplayConfigSetDeviceInfo(ref DISPLAYCONFIG_SOURCE_DPI_SCALE_SET info);" & vbCrLf & _
    "" & vbCrLf & _
    "    // ---- Constants ----" & vbCrLf & _
    "" & vbCrLf & _
    "    public const int ENUM_CURRENT_SETTINGS = -1;" & vbCrLf & _
    "    public const uint QDC_ONLY_ACTIVE_PATHS = 2;" & vbCrLf & _
    "    public const int DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE = -3;" & vbCrLf & _
    "    public const int DISPLAYCONFIG_DEVICE_INFO_SET_DPI_SCALE = -4;" & vbCrLf & _
    "    public const int DM_PELSWIDTH = 0x80000;" & vbCrLf & _
    "    public const int DM_PELSHEIGHT = 0x100000;" & vbCrLf & _
    "    public const uint MONITOR_DEFAULTTOPRIMARY = 1;" & vbCrLf & _
    "    public const int MDT_EFFECTIVE_DPI = 0;" & vbCrLf & _
    "" & vbCrLf & _
    "    public static readonly int[] DpiSteps = { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };" & vbCrLf & _
    "" & vbCrLf & _
    "    // ---- Public API ----" & vbCrLf & _
    "" & vbCrLf & _
    "    public static void GetResolution(out int width, out int height)" & vbCrLf & _
    "    {" & vbCrLf & _
    "        var dm = new DEVMODE();" & vbCrLf & _
    "        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));" & vbCrLf & _
    "        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);" & vbCrLf & _
    "        width = dm.dmPelsWidth;" & vbCrLf & _
    "        height = dm.dmPelsHeight;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    public static int GetScaling()" & vbCrLf & _
    "    {" & vbCrLf & _
    "        // Try per-monitor DPI via GetDpiForMonitor" & vbCrLf & _
    "        var pt = new POINT { x = 0, y = 0 };" & vbCrLf & _
    "        IntPtr hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);" & vbCrLf & _
    "        uint dpiX, dpiY;" & vbCrLf & _
    "        if (GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI, out dpiX, out dpiY) == 0 && dpiX > 0)" & vbCrLf & _
    "            return (int)Math.Round(dpiX / 96.0 * 100);" & vbCrLf & _
    "" & vbCrLf & _
    "        // Fallback: system DPI via GDI" & vbCrLf & _
    "        IntPtr hdc = GetDC(IntPtr.Zero);" & vbCrLf & _
    "        int logPixels = GetDeviceCaps(hdc, 88); // LOGPIXELSX" & vbCrLf & _
    "        ReleaseDC(IntPtr.Zero, hdc);" & vbCrLf & _
    "        return logPixels > 0 ? (int)Math.Round(logPixels / 96.0 * 100) : 100;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    public static bool SetResolution(int width, int height)" & vbCrLf & _
    "    {" & vbCrLf & _
    "        var dm = new DEVMODE();" & vbCrLf & _
    "        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));" & vbCrLf & _
    "        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);" & vbCrLf & _
    "        dm.dmPelsWidth = width;" & vbCrLf & _
    "        dm.dmPelsHeight = height;" & vbCrLf & _
    "        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;" & vbCrLf & _
    "        return ChangeDisplaySettingsEx(null, ref dm, IntPtr.Zero, 0x01, IntPtr.Zero) == 0;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    public static bool SetScaling(int targetPercent)" & vbCrLf & _
    "    {" & vbCrLf & _
    "        // Get active display paths" & vbCrLf & _
    "        uint pathCount, modeCount;" & vbCrLf & _
    "        if (GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out pathCount, out modeCount) != 0)" & vbCrLf & _
    "            return false;" & vbCrLf & _
    "        var paths = new DISPLAYCONFIG_PATH_INFO[pathCount];" & vbCrLf & _
    "        var modes = new DISPLAYCONFIG_MODE_INFO[modeCount];" & vbCrLf & _
    "        if (QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero) != 0)" & vbCrLf & _
    "            return false;" & vbCrLf & _
    "        if (pathCount == 0) return false;" & vbCrLf & _
    "" & vbCrLf & _
    "        // Find built-in display, fall back to first path" & vbCrLf & _
    "        int idx = 0;" & vbCrLf & _
    "        for (int i = 0; i < (int)pathCount; i++)" & vbCrLf & _
    "        {" & vbCrLf & _
    "            int tech = paths[i].targetInfo.outputTechnology;" & vbCrLf & _
    "            if (tech == unchecked((int)0x80000000) || tech == 11) // INTERNAL or DP_EMBEDDED" & vbCrLf & _
    "            { idx = i; break; }" & vbCrLf & _
    "        }" & vbCrLf & _
    "        var path = paths[idx];" & vbCrLf & _
    "" & vbCrLf & _
    "        // Get current DPI scale info" & vbCrLf & _
    "        var get = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();" & vbCrLf & _
    "        get.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE;" & vbCrLf & _
    "        get.header.size = Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));" & vbCrLf & _
    "        get.header.adapterId = path.sourceInfo.adapterId;" & vbCrLf & _
    "        get.header.id = path.sourceInfo.id;" & vbCrLf & _
    "        if (DisplayConfigGetDeviceInfo(ref get) != 0) return false;" & vbCrLf & _
    "" & vbCrLf & _
    "        // Compute target relative scale" & vbCrLf & _
    "        int curPercent = GetScaling();" & vbCrLf & _
    "        int curAbsIdx = FindClosest(curPercent);" & vbCrLf & _
    "        int recAbsIdx = curAbsIdx - get.curScaleRel;" & vbCrLf & _
    "        int targetAbsIdx = FindClosest(targetPercent);" & vbCrLf & _
    "        int targetRel = targetAbsIdx - recAbsIdx;" & vbCrLf & _
    "" & vbCrLf & _
    "        if (targetRel < get.minScaleRel || targetRel > get.maxScaleRel)" & vbCrLf & _
    "            return false;" & vbCrLf & _
    "" & vbCrLf & _
    "        // Apply" & vbCrLf & _
    "        var set = new DISPLAYCONFIG_SOURCE_DPI_SCALE_SET();" & vbCrLf & _
    "        set.header.type = DISPLAYCONFIG_DEVICE_INFO_SET_DPI_SCALE;" & vbCrLf & _
    "        set.header.size = Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_SET));" & vbCrLf & _
    "        set.header.adapterId = path.sourceInfo.adapterId;" & vbCrLf & _
    "        set.header.id = path.sourceInfo.id;" & vbCrLf & _
    "        set.scaleRel = targetRel;" & vbCrLf & _
    "        return DisplayConfigSetDeviceInfo(ref set) == 0;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    public static Icon GetShellIcon(int index)" & vbCrLf & _
    "    {" & vbCrLf & _
    "        IntPtr h = ExtractIcon(IntPtr.Zero, ""shell32.dll"", index);" & vbCrLf & _
    "        if (h == IntPtr.Zero || h == (IntPtr)1)" & vbCrLf & _
    "            return SystemIcons.Application;" & vbCrLf & _
    "        Icon ico = (Icon)Icon.FromHandle(h).Clone();" & vbCrLf & _
    "        DestroyIcon(h);" & vbCrLf & _
    "        return ico;" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    private static int FindClosest(int percent)" & vbCrLf & _
    "    {" & vbCrLf & _
    "        int best = 0, bestDiff = Math.Abs(DpiSteps[0] - percent);" & vbCrLf & _
    "        for (int i = 1; i < DpiSteps.Length; i++)" & vbCrLf & _
    "        {" & vbCrLf & _
    "            int diff = Math.Abs(DpiSteps[i] - percent);" & vbCrLf & _
    "            if (diff < bestDiff) { best = i; bestDiff = diff; }" & vbCrLf & _
    "        }" & vbCrLf & _
    "        return best;" & vbCrLf & _
    "    }" & vbCrLf & _
    "}" & vbCrLf & _
    "'@ -ReferencedAssemblies System.Drawing" & vbCrLf & _
    "" & vbCrLf & _
    "# Make process per-monitor DPI aware (best effort)" & vbCrLf & _
    "try { [DisplayHelper]::SetProcessDpiAwareness(2) } catch {}" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Icons ----" & vbCrLf & _
    "$iconMonitor = [DisplayHelper]::GetShellIcon(15)   # computer/monitor" & vbCrLf & _
    "$iconWarning = [DisplayHelper]::GetShellIcon(77)   # yellow warning triangle" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- UI Setup ----" & vbCrLf & _
    "$appContext = New-Object System.Windows.Forms.ApplicationContext" & vbCrLf & _
    "" & vbCrLf & _
    "# Hook into display change messages" & vbCrLf & _
    "Add-Type -TypeDefinition @'" & vbCrLf & _
    "using System;" & vbCrLf & _
    "using System.Windows.Forms;" & vbCrLf & _
    "" & vbCrLf & _
    "public class MessageWindow : NativeWindow {" & vbCrLf & _
    "    public event EventHandler DisplayChanged;" & vbCrLf & _
    "" & vbCrLf & _
    "    protected override void WndProc(ref Message m) {" & vbCrLf & _
    "        const int WM_DISPLAYCHANGE = 0x007E;" & vbCrLf & _
    "        const int WM_DPICHANGED = 0x02E0;" & vbCrLf & _
    "        const int WM_SETTINGCHANGE = 0x001A;" & vbCrLf & _
    "" & vbCrLf & _
    "        base.WndProc(ref m);" & vbCrLf & _
    "        if ((DisplayChanged != null) && (m.Msg == WM_DISPLAYCHANGE || m.Msg == WM_DPICHANGED || m.Msg == WM_SETTINGCHANGE)) {" & vbCrLf & _
    "            DisplayChanged(this, EventArgs.Empty);" & vbCrLf & _
    "        }" & vbCrLf & _
    "    }" & vbCrLf & _
    "}" & vbCrLf & _
    "'@ -ReferencedAssemblies System.Windows.Forms" & vbCrLf & _
    "" & vbCrLf & _
    "$msgWindow = New-Object MessageWindow" & vbCrLf & _
    "$msgWindow.CreateHandle((New-Object System.Windows.Forms.CreateParams))" & vbCrLf & _
    "$msgWindow.Add_DisplayChanged({ Refresh })" & vbCrLf & _
    "" & vbCrLf & _
    "$notifyIcon = New-Object System.Windows.Forms.NotifyIcon" & vbCrLf & _
    "$notifyIcon.Visible = $true" & vbCrLf & _
    "$notifyIcon.Text = $applicationFullName" & vbCrLf & _
    "" & vbCrLf & _
    "$contextMenu = New-Object System.Windows.Forms.ContextMenuStrip" & vbCrLf & _
    "$script:menuApplyTarget = $contextMenu.Items.Add(""Apply Target"")" & vbCrLf & _
    "$menuRefresh     = $contextMenu.Items.Add(""Refresh"")" & vbCrLf & _
    "$contextMenu.Items.Add(""-"")  # separator" & vbCrLf & _
    "" & vbCrLf & _
    "# Options submenu" & vbCrLf & _
    "$menuOptions = New-Object System.Windows.Forms.ToolStripMenuItem(""Options"")" & vbCrLf & _
    "" & vbCrLf & _
    "# Target resolution submenu" & vbCrLf & _
    "$script:menuTargetRes = New-Object System.Windows.Forms.ToolStripMenuItem(""Target resolution"")" & vbCrLf & _
    "foreach ($res in $script:TargetResolutions) {" & vbCrLf & _
    "    $item = New-Object System.Windows.Forms.ToolStripMenuItem(""$($res.Width) x $($res.Height)"")" & vbCrLf & _
    "    $item.Tag = $res" & vbCrLf & _
    "    $item.CheckOnClick = $true" & vbCrLf & _
    "    $item.Add_Click({" & vbCrLf & _
    "        param($sender, $e)" & vbCrLf & _
    "        $script:TargetState = $sender.Tag + @{Scaling = $script:TargetState.Scaling}" & vbCrLf & _
    "        ChangeTargetState" & vbCrLf & _
    "    })" & vbCrLf & _
    "    $script:menuTargetRes.DropDownItems.Add($item) | Out-Null" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "$script:menuTargetRes.DropDownItems.Add(""-"") | Out-Null" & vbCrLf & _
    "# Custom resolution menu item" & vbCrLf & _
    "Add-Type -AssemblyName Microsoft.VisualBasic" & vbCrLf & _
    "$script:menuCustomRes = New-Object System.Windows.Forms.ToolStripMenuItem(""Custom"")" & vbCrLf & _
    "$script:menuCustomRes.CheckOnClick = $true" & vbCrLf & _
    "$script:menuCustomRes.Add_Click({" & vbCrLf & _
    "    param($sender, $e)" & vbCrLf & _
    "    $input = [Microsoft.VisualBasic.Interaction]::InputBox(""Enter resolution (e.g. 1920x1080)"", ""Custom Resolution"", ""$($sender.Tag.Width) x $($sender.Tag.Height)"")" & vbCrLf & _
    "    if ($input) {" & vbCrLf & _
    "        $numbers = ($input -replace '[^\d]', ' ').Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)" & vbCrLf & _
    "        if ($numbers.Count -ge 2) {" & vbCrLf & _
    "            $sender.Tag = @{Width = [int]$numbers[0]; Height = [int]$numbers[-1]}" & vbCrLf & _
    "            $script:TargetState = $sender.Tag + @{Scaling = $script:TargetState.Scaling}" & vbCrLf & _
    "            ChangeTargetState" & vbCrLf & _
    "        }" & vbCrLf & _
    "    }" & vbCrLf & _
    "})" & vbCrLf & _
    "$script:menuTargetRes.DropDownItems.Add($script:menuCustomRes) | Out-Null" & vbCrLf & _
    "$menuOptions.DropDownItems.Add($script:menuTargetRes) | Out-Null" & vbCrLf & _
    "" & vbCrLf & _
    "# Target scaling submenu" & vbCrLf & _
    "$script:menuTargetScale = New-Object System.Windows.Forms.ToolStripMenuItem(""Target scaling"")" & vbCrLf & _
    "foreach ($scale in $script:TargetScalings) {" & vbCrLf & _
    "    $item = New-Object System.Windows.Forms.ToolStripMenuItem(""$scale%"")" & vbCrLf & _
    "    $item.Tag = $scale" & vbCrLf & _
    "    $item.CheckOnClick = $true" & vbCrLf & _
    "    $item.Add_Click({" & vbCrLf & _
    "        param($sender, $e)" & vbCrLf & _
    "        $script:TargetState.Scaling = $sender.Tag" & vbCrLf & _
    "        ChangeTargetState" & vbCrLf & _
    "    })" & vbCrLf & _
    "    $script:menuTargetScale.DropDownItems.Add($item) | Out-Null" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "$script:menuTargetScale.DropDownItems.Add(""-"") | Out-Null #separator" & vbCrLf & _
    "# Custom scaling menu item" & vbCrLf & _
    "$script:menuCustomScale = New-Object System.Windows.Forms.ToolStripMenuItem(""Custom"")" & vbCrLf & _
    "$script:menuCustomScale.Tag = 200" & vbCrLf & _
    "$script:menuCustomScale.CheckOnClick = $true" & vbCrLf & _
    "$script:menuCustomScale.Add_Click({" & vbCrLf & _
    "    param($sender, $e)" & vbCrLf & _
    "    $input = [Microsoft.VisualBasic.Interaction]::InputBox(""Enter scaling percentage (e.g. 100)"", ""Custom Scaling"", ""$($sender.Tag)"")" & vbCrLf & _
    "    if ($input) {" & vbCrLf & _
    "        $numbers = ($input -replace '[^\d]', ' ').Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries)" & vbCrLf & _
    "        if ($numbers.Count -ge 1) {" & vbCrLf & _
    "            $sender.Tag = [int]$numbers[0]" & vbCrLf & _
    "            $script:TargetState.Scaling = $sender.Tag" & vbCrLf & _
    "            ChangeTargetState" & vbCrLf & _
    "        }" & vbCrLf & _
    "    }" & vbCrLf & _
    "})" & vbCrLf & _
    "$script:menuTargetScale.DropDownItems.Add($script:menuCustomScale) | Out-Null" & vbCrLf & _
    "$menuOptions.DropDownItems.Add($script:menuTargetScale) | Out-Null" & vbCrLf & _
    "" & vbCrLf & _
    "$menuOptions.DropDownItems.Add(""-"")  # separator" & vbCrLf & _
    "# Auto start option" & vbCrLf & _
    "$menuAutoStart = New-Object System.Windows.Forms.ToolStripMenuItem(""Auto start"")" & vbCrLf & _
    "$menuAutoStart.CheckOnClick = $true" & vbCrLf & _
    "$menuOptions.DropDownItems.Add($menuAutoStart) | Out-Null" & vbCrLf & _
    "" & vbCrLf & _
    "$contextMenu.Items.Add($menuOptions) | Out-Null" & vbCrLf & _
    "$contextMenu.Items.Add(""-"")  # separator" & vbCrLf & _
    "$menuExit = $contextMenu.Items.Add(""Exit"")" & vbCrLf & _
    "$notifyIcon.ContextMenuStrip = $contextMenu" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Auto-start Logic ----" & vbCrLf & _
    "$AutoStartRegPath = ""HKCU:\Software\Microsoft\Windows\CurrentVersion\Run""" & vbCrLf & _
    "$AutoStartValueName = $applicationFullName" & vbCrLf & _
    "$VbsPath = Join-Path (Split-Path $PSCommandPath -Parent) ""ResolutionMonitor.vbs""" & vbCrLf & _
    "" & vbCrLf & _
    "# Update auto-start checkbox before menu opens" & vbCrLf & _
    "$contextMenu.Add_Opening({" & vbCrLf & _
    "    try {" & vbCrLf & _
    "        $menuAutoStart.Checked = $null -ne (Get-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -ErrorAction SilentlyContinue)" & vbCrLf & _
    "    } catch {" & vbCrLf & _
    "        $menuAutoStart.Checked = $false" & vbCrLf & _
    "    }" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Refresh Logic ----" & vbCrLf & _
    "function Refresh {" & vbCrLf & _
    "    $w = 0; $h = 0" & vbCrLf & _
    "    [DisplayHelper]::GetResolution([ref]$w, [ref]$h)" & vbCrLf & _
    "    $s = [DisplayHelper]::GetScaling()" & vbCrLf & _
    "" & vbCrLf & _
    "    if ($w -eq $script:TargetState.Width -and $h -eq $script:TargetState.Height -and $s -eq $script:TargetState.Scaling) {" & vbCrLf & _
    "        $notifyIcon.Icon = $iconMonitor" & vbCrLf & _
    "    } else {" & vbCrLf & _
    "        $notifyIcon.Icon = $iconWarning" & vbCrLf & _
    "    }" & vbCrLf & _
    "" & vbCrLf & _
    "    $old = $notifyIcon.Text" & vbCrLf & _
    "    $notifyIcon.Text = ""${s}% @ ${w}x${h}""" & vbCrLf & _
    "    if (($old -ne $applicationFullName) -and ($old -ne $notifyIcon.Text)) {" & vbCrLf & _
    "        $notifyIcon.ShowBalloonTip(5000, ""Resolution Monitor"", $notifyIcon.Text, [System.Windows.Forms.ToolTipIcon]::Info)" & vbCrLf & _
    "    }" & vbCrLf & _
    "}" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Event Handlers ----" & vbCrLf & _
    "$script:menuApplyTarget.Add_Click({" & vbCrLf & _
    "    [DisplayHelper]::SetResolution($script:TargetState.Width, $script:TargetState.Height) | Out-Null" & vbCrLf & _
    "    [DisplayHelper]::SetScaling($script:TargetState.Scaling) | Out-Null" & vbCrLf & _
    "    Refresh" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "$menuRefresh.Add_Click({" & vbCrLf & _
    "    Refresh" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "$menuAutoStart.Add_Click({" & vbCrLf & _
    "    if ($menuAutoStart.Checked) {" & vbCrLf & _
    "        try {" & vbCrLf & _
    "            Set-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -Value ""wscript.exe `""$VbsPath`""""" & vbCrLf & _
    "        } catch {" & vbCrLf & _
    "            $menuAutoStart.Checked = $false" & vbCrLf & _
    "        }" & vbCrLf & _
    "    } else {" & vbCrLf & _
    "        try {" & vbCrLf & _
    "            Remove-ItemProperty -Path $AutoStartRegPath -Name $AutoStartValueName -ErrorAction SilentlyContinue" & vbCrLf & _
    "        } catch {" & vbCrLf & _
    "            $menuAutoStart.Checked = $true" & vbCrLf & _
    "        }" & vbCrLf & _
    "    }" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "$menuExit.Add_Click({" & vbCrLf & _
    "    $timer.Stop()" & vbCrLf & _
    "    $msgWindow.DestroyHandle()" & vbCrLf & _
    "    $notifyIcon.Visible = $false" & vbCrLf & _
    "    $notifyIcon.Dispose()" & vbCrLf & _
    "    [System.Environment]::Exit(0)" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "$notifyIcon.Add_Click({" & vbCrLf & _
    "    param($sender, $e)" & vbCrLf & _
    "    if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) {" & vbCrLf & _
    "        $contextMenu.Show([System.Windows.Forms.Cursor]::Position)" & vbCrLf & _
    "    }" & vbCrLf & _
    "})" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Timer ----" & vbCrLf & _
    "$timer = New-Object System.Windows.Forms.Timer" & vbCrLf & _
    "$timer.Interval = 30000" & vbCrLf & _
    "$timer.Add_Tick({ Refresh })" & vbCrLf & _
    "" & vbCrLf & _
    "# ---- Start ----" & vbCrLf & _
    "ReloadTargetState" & vbCrLf & _
    "$timer.Start()" & vbCrLf & _
    "[System.Windows.Forms.Application]::Run($appContext)" & vbCrLf & _
    ""
f.Close
CreateObject("Wscript.Shell").Run "powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & tmp & """", 0, False
