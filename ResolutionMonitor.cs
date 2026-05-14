// ResolutionMonitor.cs - Resolution Monitor system tray tool
// Single-file C# application targeting .NET Framework 4.x (AnyCPU) and .NET 8 (x64)
// Build with: csc.exe /target:winexe /optimize /r:System.Windows.Forms.dll /r:System.Drawing.dll ResolutionMonitor.cs

using System;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Win32;

// ---- P/Invoke Helper ----

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
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        width = dm.dmPelsWidth;
        height = dm.dmPelsHeight;
    }

    public static int GetScaling()
    {
        POINT pt = new POINT();
        pt.x = 0;
        pt.y = 0;
        IntPtr hmon = MonitorFromPoint(pt, MONITOR_DEFAULTTOPRIMARY);
        uint dpiX, dpiY;
        if (GetDpiForMonitor(hmon, MDT_EFFECTIVE_DPI, out dpiX, out dpiY) == 0 && dpiX > 0)
            return (int)Math.Round(dpiX / 96.0 * 100);

        IntPtr hdc = GetDC(IntPtr.Zero);
        int logPixels = GetDeviceCaps(hdc, 88);
        ReleaseDC(IntPtr.Zero, hdc);
        return logPixels > 0 ? (int)Math.Round(logPixels / 96.0 * 100) : 100;
    }

    public static bool SetResolution(int width, int height)
    {
        DEVMODE dm = new DEVMODE();
        dm.dmSize = (short)Marshal.SizeOf(typeof(DEVMODE));
        EnumDisplaySettings(null, ENUM_CURRENT_SETTINGS, ref dm);
        dm.dmPelsWidth = width;
        dm.dmPelsHeight = height;
        dm.dmFields = DM_PELSWIDTH | DM_PELSHEIGHT;
        return ChangeDisplaySettingsEx(null, ref dm, IntPtr.Zero, 0x01, IntPtr.Zero) == 0;
    }

    public static bool SetScaling(int targetPercent)
    {
        uint pathCount, modeCount;
        if (GetDisplayConfigBufferSizes(QDC_ONLY_ACTIVE_PATHS, out pathCount, out modeCount) != 0)
            return false;
        DISPLAYCONFIG_PATH_INFO[] paths = new DISPLAYCONFIG_PATH_INFO[pathCount];
        DISPLAYCONFIG_MODE_INFO[] modes = new DISPLAYCONFIG_MODE_INFO[modeCount];
        if (QueryDisplayConfig(QDC_ONLY_ACTIVE_PATHS, ref pathCount, paths, ref modeCount, modes, IntPtr.Zero) != 0)
            return false;
        if (pathCount == 0) return false;

        int idx = 0;
        for (int i = 0; i < (int)pathCount; i++)
        {
            int tech = paths[i].targetInfo.outputTechnology;
            if (tech == unchecked((int)0x80000000) || tech == 11)
            { idx = i; break; }
        }
        DISPLAYCONFIG_PATH_INFO path = paths[idx];

        DISPLAYCONFIG_SOURCE_DPI_SCALE_GET get = new DISPLAYCONFIG_SOURCE_DPI_SCALE_GET();
        get.header.type = DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE;
        get.header.size = Marshal.SizeOf(typeof(DISPLAYCONFIG_SOURCE_DPI_SCALE_GET));
        get.header.adapterId = path.sourceInfo.adapterId;
        get.header.id = path.sourceInfo.id;
        if (DisplayConfigGetDeviceInfo(ref get) != 0) return false;

        int curPercent = GetScaling();
        int curAbsIdx = FindClosest(curPercent);
        int recAbsIdx = curAbsIdx - get.curScaleRel;
        int targetAbsIdx = FindClosest(targetPercent);
        int targetRel = targetAbsIdx - recAbsIdx;

        if (targetRel < get.minScaleRel || targetRel > get.maxScaleRel)
            return false;

        DISPLAYCONFIG_SOURCE_DPI_SCALE_SET set = new DISPLAYCONFIG_SOURCE_DPI_SCALE_SET();
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
        int best = 0;
        int bestDiff = Math.Abs(DpiSteps[0] - percent);
        for (int i = 1; i < DpiSteps.Length; i++)
        {
            int diff = Math.Abs(DpiSteps[i] - percent);
            if (diff < bestDiff) { best = i; bestDiff = diff; }
        }
        return best;
    }
}

// ---- Display Change Message Window ----

public class MessageWindow : NativeWindow
{
    public event EventHandler DisplayChanged;

    protected override void WndProc(ref Message m)
    {
        const int WM_DISPLAYCHANGE = 0x007E;
        const int WM_DPICHANGED = 0x02E0;
        const int WM_SETTINGCHANGE = 0x001A;

        base.WndProc(ref m);
        if (DisplayChanged != null &&
            (m.Msg == WM_DISPLAYCHANGE || m.Msg == WM_DPICHANGED || m.Msg == WM_SETTINGCHANGE))
        {
            DisplayChanged(this, EventArgs.Empty);
        }
    }
}

// ---- Input Dialog (replaces VB InputBox) ----

public static class InputDialog
{
    public static string Show(string prompt, string title, string defaultValue)
    {
        Form form = new Form();
        Label label = new Label();
        TextBox textBox = new TextBox();
        Button okButton = new Button();
        Button cancelButton = new Button();

        form.Text = title;
        label.Text = prompt;
        textBox.Text = defaultValue;

        okButton.Text = "OK";
        cancelButton.Text = "Cancel";
        okButton.DialogResult = DialogResult.OK;
        cancelButton.DialogResult = DialogResult.Cancel;

        label.SetBounds(9, 10, 280, 20);
        textBox.SetBounds(12, 36, 280, 20);
        okButton.SetBounds(131, 72, 75, 23);
        cancelButton.SetBounds(212, 72, 75, 23);

        form.ClientSize = new Size(300, 107);
        form.Controls.AddRange(new Control[] { label, textBox, okButton, cancelButton });
        form.FormBorderStyle = FormBorderStyle.FixedDialog;
        form.StartPosition = FormStartPosition.CenterScreen;
        form.MinimizeBox = false;
        form.MaximizeBox = false;
        form.AcceptButton = okButton;
        form.CancelButton = cancelButton;

        DialogResult result = form.ShowDialog();
        string value = textBox.Text;
        form.Dispose();

        if (result == DialogResult.OK)
            return value;
        return null;
    }
}

// ---- Main Application ----

public class ResolutionMonitorApp
{
    const string AppVersion = "1.0";
    const string AppDeveloper = "PanSoft";
    const string AppShortName = "Resolution Monitor";
    static readonly string AppFullName = AppDeveloper + " " + AppShortName + " v" + AppVersion;
    const string SettingsRegPath = @"Software\PanSoft\Resolution Monitor";
    const string AutoStartRegPath = @"Software\Microsoft\Windows\CurrentVersion\Run";

    static readonly int[][] TargetResolutions = new int[][] {
        new int[] {1024, 768},
        new int[] {1152, 864},
        new int[] {1280, 720},
        new int[] {1280, 768},
        new int[] {1280, 800},
        new int[] {1280, 1024},
        new int[] {1360, 768},
        new int[] {1366, 768},
        new int[] {1440, 900},
        new int[] {1600, 900},
        new int[] {1600, 1200},
        new int[] {1680, 1050},
        new int[] {1920, 1080},
        new int[] {1920, 1200},
        new int[] {2048, 1152},
        new int[] {2560, 1440},
        new int[] {2560, 1600},
        new int[] {3200, 1800},
        new int[] {3840, 2160},
        new int[] {5120, 2880},
        new int[] {7680, 4320}
    };

    static readonly int[] TargetScalings = new int[] { 100, 125, 150, 175, 200 };

    // Target state
    int targetWidth = 1920;
    int targetHeight = 1080;
    int targetScaling = 100;

    // Custom menu item tag values
    int customResWidth = 1080;
    int customResHeight = 1920;
    int customScaling = 100;

    // UI elements
    NotifyIcon notifyIcon;
    ContextMenuStrip contextMenu;
    ToolStripMenuItem menuApplyTarget;
    ToolStripMenuItem menuTargetRes;
    ToolStripMenuItem menuTargetScale;
    ToolStripMenuItem menuCustomRes;
    ToolStripMenuItem menuCustomScale;
    ToolStripMenuItem menuAutoStart;
    Icon iconMonitor;
    Icon iconWarning;
    Timer timer;
    MessageWindow msgWindow;

    public ResolutionMonitorApp()
    {
        // Set per-monitor DPI awareness
        try { DisplayHelper.SetProcessDpiAwareness(2); } catch { }

        // Load icons
        iconMonitor = DisplayHelper.GetShellIcon(15);
        iconWarning = DisplayHelper.GetShellIcon(77);

        // Create message window for display change notifications
        msgWindow = new MessageWindow();
        msgWindow.CreateHandle(new CreateParams());
        msgWindow.DisplayChanged += delegate { Refresh(); };

        // Create notify icon
        notifyIcon = new NotifyIcon();
        notifyIcon.Visible = true;
        notifyIcon.Text = AppFullName;

        // Build context menu
        BuildMenu();
        notifyIcon.ContextMenuStrip = contextMenu;

        // Left-click shows context menu
        notifyIcon.MouseClick += delegate(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                contextMenu.Show(Cursor.Position);
        };

        // Polling timer (30 seconds)
        timer = new Timer();
        timer.Interval = 30000;
        timer.Tick += delegate { Refresh(); };

        // Load state and start
        ReloadTargetState();
        timer.Start();
    }

    void BuildMenu()
    {
        contextMenu = new ContextMenuStrip();

        // Apply Target
        menuApplyTarget = new ToolStripMenuItem("Apply Target");
        menuApplyTarget.Click += delegate
        {
            DisplayHelper.SetResolution(targetWidth, targetHeight);
            DisplayHelper.SetScaling(targetScaling);
            Refresh();
        };
        contextMenu.Items.Add(menuApplyTarget);

        // Refresh
        ToolStripMenuItem menuRefresh = new ToolStripMenuItem("Refresh");
        menuRefresh.Click += delegate { Refresh(); };
        contextMenu.Items.Add(menuRefresh);

        contextMenu.Items.Add(new ToolStripSeparator());

        // Options submenu
        ToolStripMenuItem menuOptions = new ToolStripMenuItem("Options");

        // Target resolution submenu
        menuTargetRes = new ToolStripMenuItem("Target resolution");
        for (int i = 0; i < TargetResolutions.Length; i++)
        {
            int w = TargetResolutions[i][0];
            int h = TargetResolutions[i][1];
            ToolStripMenuItem item = new ToolStripMenuItem(w + " x " + h);
            item.Tag = new int[] { w, h };
            item.CheckOnClick = true;
            item.Click += delegate(object sender, EventArgs e)
            {
                int[] res = (int[])((ToolStripMenuItem)sender).Tag;
                targetWidth = res[0];
                targetHeight = res[1];
                ChangeTargetState();
            };
            menuTargetRes.DropDownItems.Add(item);
        }

        menuTargetRes.DropDownItems.Add(new ToolStripSeparator());

        // Custom resolution
        menuCustomRes = new ToolStripMenuItem("Custom");
        menuCustomRes.CheckOnClick = true;
        menuCustomRes.Click += delegate(object sender, EventArgs e)
        {
            string defVal = customResWidth + " x " + customResHeight;
            string input = InputDialog.Show("Enter resolution (e.g. 1920x1080)", "Custom Resolution", defVal);
            if (input != null)
            {
                // Extract numbers from input
                string digits = "";
                foreach (char c in input)
                    digits += char.IsDigit(c) ? c : ' ';
                string[] parts = digits.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 2)
                {
                    customResWidth = int.Parse(parts[0]);
                    customResHeight = int.Parse(parts[parts.Length - 1]);
                    targetWidth = customResWidth;
                    targetHeight = customResHeight;
                    ChangeTargetState();
                    return;
                }
            }
            // If cancelled or invalid, reload to fix checkmarks
            ReloadTargetState();
        };
        menuTargetRes.DropDownItems.Add(menuCustomRes);
        menuOptions.DropDownItems.Add(menuTargetRes);

        // Target scaling submenu
        menuTargetScale = new ToolStripMenuItem("Target scaling");
        for (int i = 0; i < TargetScalings.Length; i++)
        {
            int s = TargetScalings[i];
            ToolStripMenuItem item = new ToolStripMenuItem(s + "%");
            item.Tag = s;
            item.CheckOnClick = true;
            item.Click += delegate(object sender, EventArgs e)
            {
                targetScaling = (int)((ToolStripMenuItem)sender).Tag;
                ChangeTargetState();
            };
            menuTargetScale.DropDownItems.Add(item);
        }

        menuTargetScale.DropDownItems.Add(new ToolStripSeparator());

        // Custom scaling
        menuCustomScale = new ToolStripMenuItem("Custom");
        menuCustomScale.CheckOnClick = true;
        menuCustomScale.Click += delegate(object sender, EventArgs e)
        {
            string input = InputDialog.Show("Enter scaling percentage (e.g. 100)", "Custom Scaling", customScaling.ToString());
            if (input != null)
            {
                string digits = "";
                foreach (char c in input)
                    digits += char.IsDigit(c) ? c : ' ';
                string[] parts = digits.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length >= 1)
                {
                    customScaling = int.Parse(parts[0]);
                    targetScaling = customScaling;
                    ChangeTargetState();
                    return;
                }
            }
            ReloadTargetState();
        };
        menuTargetScale.DropDownItems.Add(menuCustomScale);
        menuOptions.DropDownItems.Add(menuTargetScale);

        menuOptions.DropDownItems.Add(new ToolStripSeparator());

        // Auto start
        menuAutoStart = new ToolStripMenuItem("Auto start");
        menuAutoStart.CheckOnClick = true;
        menuAutoStart.Click += delegate
        {
            if (menuAutoStart.Checked)
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(AutoStartRegPath))
                    {
                        key.SetValue(AppFullName, Application.ExecutablePath);
                    }
                }
                catch
                {
                    menuAutoStart.Checked = false;
                }
            }
            else
            {
                try
                {
                    using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AutoStartRegPath, true))
                    {
                        if (key != null)
                            key.DeleteValue(AppFullName, false);
                    }
                }
                catch
                {
                    menuAutoStart.Checked = true;
                }
            }
        };
        menuOptions.DropDownItems.Add(menuAutoStart);

        contextMenu.Items.Add(menuOptions);
        contextMenu.Items.Add(new ToolStripSeparator());

        // Exit
        ToolStripMenuItem menuExit = new ToolStripMenuItem("Exit");
        menuExit.Click += delegate
        {
            timer.Stop();
            msgWindow.DestroyHandle();
            notifyIcon.Visible = false;
            notifyIcon.Dispose();
            Environment.Exit(0);
        };
        contextMenu.Items.Add(menuExit);

        // Update auto-start checkbox before menu opens
        contextMenu.Opening += delegate
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(AutoStartRegPath))
                {
                    menuAutoStart.Checked = key != null && key.GetValue(AppFullName) != null;
                }
            }
            catch
            {
                menuAutoStart.Checked = false;
            }
        };
    }

    string ReadSetting(string name, string defaultValue)
    {
        try
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(SettingsRegPath))
            {
                if (key != null)
                {
                    object val = key.GetValue(name);
                    if (val != null) return val.ToString();
                }
            }
        }
        catch { }
        return defaultValue;
    }

    void ReadSettings(string prefix, out int width, out int height, out int scaling)
    {
        width = int.Parse(ReadSetting(prefix + "Width", "1080"));
        height = int.Parse(ReadSetting(prefix + "Height", "1920"));
        scaling = int.Parse(ReadSetting(prefix + "Scaling", "100"));
    }

    void WriteSetting(string name, int value)
    {
        using (RegistryKey key = Registry.CurrentUser.CreateSubKey(SettingsRegPath))
        {
            key.SetValue(name, value, RegistryValueKind.DWord);
        }
    }

    void WriteSettings(string prefix, int width, int height, int scaling)
    {
        WriteSetting(prefix + "Width", width);
        WriteSetting(prefix + "Height", height);
        WriteSetting(prefix + "Scaling", scaling);
    }

    void ReloadTargetState()
    {
        targetWidth = 1920;
        targetHeight = 1080;
        targetScaling = 100;

        // Load target from registry
        int loadedW, loadedH, loadedS;
        ReadSettings("Target", out loadedW, out loadedH, out loadedS);
        targetWidth = loadedW;
        targetHeight = loadedH;
        targetScaling = loadedS;

        menuApplyTarget.Text = "Apply " + targetScaling + "% " + targetWidth + "x" + targetHeight;

        // Load custom tag values from registry
        int custW, custH, custS;
        ReadSettings("Custom0", out custW, out custH, out custS);
        customResWidth = custW;
        customResHeight = custH;
        customScaling = custS;

        // Update resolution menu checkmarks
        bool found = false;
        foreach (ToolStripItem tsi in menuTargetRes.DropDownItems)
        {
            ToolStripMenuItem item = tsi as ToolStripMenuItem;
            if (item == null) break; // hit separator = end of standard items
            int[] res = item.Tag as int[];
            if (res != null)
            {
                item.Checked = (res[0] == targetWidth && res[1] == targetHeight);
                if (item.Checked) found = true;
            }
        }

        // If no standard resolution matched, check Custom
        if (!found)
        {
            customResWidth = targetWidth;
            customResHeight = targetHeight;
        }
        menuCustomRes.Checked = !found;
        menuCustomRes.Text = "Custom: " + customResWidth + "x" + customResHeight;

        // Update scaling menu checkmarks
        found = false;
        foreach (ToolStripItem tsi in menuTargetScale.DropDownItems)
        {
            ToolStripMenuItem item = tsi as ToolStripMenuItem;
            if (item == null) break;
            if (item.Tag is int)
            {
                item.Checked = ((int)item.Tag == targetScaling);
                if (item.Checked) found = true;
            }
        }

        if (!found)
        {
            customScaling = targetScaling;
        }
        menuCustomScale.Checked = !found;
        menuCustomScale.Text = "Custom: " + customScaling + "%";

        Refresh();
    }

    void ChangeTargetState()
    {
        WriteSettings("Target", targetWidth, targetHeight, targetScaling);
        WriteSettings("Custom0", customResWidth, customResHeight, customScaling);
        ReloadTargetState();
    }

    void Refresh()
    {
        int w, h;
        DisplayHelper.GetResolution(out w, out h);
        int s = DisplayHelper.GetScaling();

        if (w == targetWidth && h == targetHeight && s == targetScaling)
            notifyIcon.Icon = iconMonitor;
        else
            notifyIcon.Icon = iconWarning;

        string old = notifyIcon.Text;
        string current = s + "% @ " + w + "x" + h;
        notifyIcon.Text = current;

        if (old != AppFullName && old != current)
        {
            notifyIcon.ShowBalloonTip(5000, "Resolution Monitor", current, ToolTipIcon.Info);
        }
    }
}

static class Program
{
    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        new ResolutionMonitorApp();
        Application.Run(new ApplicationContext());
    }
}
