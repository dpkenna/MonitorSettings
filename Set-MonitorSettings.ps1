<#
.SYNOPSIS
    Changes the primary monitor, resolution, orientation and scaling of one monitor
    while keeping your custom monitor arrangement.

.DESCRIPTION
    Pick monitors by name, as shown by -List, instead of by number.
    - The primary-monitor change moves every display by the same offset, so their
      positions relative to each other stay the same. The whole layout is applied
      in one change, the same way Settings > Display does it.
    - Scaling uses the same per-monitor setting as Settings > Display > Scale.

.EXAMPLE
    .\Set-MonitorSettings.ps1 -List
    Lists the connected monitors with their names, positions and scaling.

.EXAMPLE
    .\Set-MonitorSettings.ps1 -Monitor 'HP' -Scaling 100 -SetPrimary
    Makes the monitor whose name contains "HP" the primary monitor at 100% scaling.
#>
[CmdletBinding(DefaultParameterSetName = 'Set')]
param(
    [Parameter(ParameterSetName = 'List', Mandatory)]
    [switch]$List,

    # A monitor name (or part of one) from -List, a device such as \\.\DISPLAY2, or just 2.
    [Parameter(ParameterSetName = 'Set', Mandatory, Position = 0)]
    [string]$Monitor,

    [Parameter(ParameterSetName = 'Set')]
    [ValidatePattern('^\d+x\d+$')]
    [string]$Resolution,

    [Parameter(ParameterSetName = 'Set')]
    [ValidateSet('landscape', 'portrait', 'landscape-flipped', 'portrait-flipped')]
    [string]$Orientation,

    [Parameter(ParameterSetName = 'Set')]
    [ValidateSet(100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500)]
    [int]$Scaling,

    [Parameter(ParameterSetName = 'Set')]
    [switch]$SetPrimary
)

$ErrorActionPreference = 'Stop'

if (-not ('MonitorTools.Display' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace MonitorTools
{
    public class MonitorInfo
    {
        public string Name;
        public string Device;
        public bool Primary;
        public int X;
        public int Y;
        public int Width;
        public int Height;
        public int RefreshHz;
        public string Orientation;
        public int Scaling;
    }

    internal static class Native
    {
        public const uint QDC_ONLY_ACTIVE_PATHS = 0x2;
        public const int ERROR_INSUFFICIENT_BUFFER = 122;
        public const int ENUM_CURRENT_SETTINGS = -1;

        public const int DM_POSITION = 0x00000020;
        public const int DM_DISPLAYORIENTATION = 0x00000080;
        public const int DM_PELSWIDTH = 0x00080000;
        public const int DM_PELSHEIGHT = 0x00100000;

        public const uint CDS_UPDATEREGISTRY = 0x00000001;

        public const uint SDC_USE_SUPPLIED_DISPLAY_CONFIG = 0x00000020;
        public const uint SDC_APPLY = 0x00000080;
        public const uint SDC_SAVE_TO_DATABASE = 0x00000200;
        public const uint SDC_ALLOW_CHANGES = 0x00000400;

        public const uint MODE_INFO_TYPE_SOURCE = 1;

        [StructLayout(LayoutKind.Sequential)]
        public struct LUID { public uint LowPart; public int HighPart; }

        [StructLayout(LayoutKind.Sequential)]
        public struct PATH_SOURCE_INFO
        {
            public LUID adapterId;
            public uint id;
            public uint modeInfoIdx;
            public uint statusFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct RATIONAL { public uint Numerator; public uint Denominator; }

        [StructLayout(LayoutKind.Sequential)]
        public struct PATH_TARGET_INFO
        {
            public LUID adapterId;
            public uint id;
            public uint modeInfoIdx;
            public uint outputTechnology;
            public uint rotation;
            public uint scaling;
            public RATIONAL refreshRate;
            public uint scanLineOrdering;
            public int targetAvailable;
            public uint statusFlags;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct PATH_INFO
        {
            public PATH_SOURCE_INFO sourceInfo;
            public PATH_TARGET_INFO targetInfo;
            public uint flags;
        }

        // DISPLAYCONFIG_MODE_INFO (64 bytes). Only the source-mode fields of the union are
        // declared; the other bytes are kept as they are and passed back unchanged.
        [StructLayout(LayoutKind.Explicit, Size = 64)]
        public struct MODE_INFO
        {
            [FieldOffset(0)] public uint infoType;
            [FieldOffset(4)] public uint id;
            [FieldOffset(8)] public LUID adapterId;
            [FieldOffset(16)] public uint sourceWidth;
            [FieldOffset(20)] public uint sourceHeight;
            [FieldOffset(24)] public uint sourcePixelFormat;
            [FieldOffset(28)] public int sourceX;
            [FieldOffset(32)] public int sourceY;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HEADER { public int type; public uint size; public LUID adapterId; public uint id; }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SOURCE_NAME
        {
            public HEADER header;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string viewGdiDeviceName;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct TARGET_NAME
        {
            public HEADER header;
            public uint flags;
            public uint outputTechnology;
            public ushort edidManufactureId;
            public ushort edidProductCodeId;
            public uint connectorInstance;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)] public string monitorFriendlyDeviceName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)] public string monitorDevicePath;
        }

        // Undocumented but widely used (it's what Settings > Display > Scale uses).
        [StructLayout(LayoutKind.Sequential)]
        public struct DPI_GET { public HEADER header; public int minScaleRel; public int curScaleRel; public int maxScaleRel; }

        [StructLayout(LayoutKind.Sequential)]
        public struct DPI_SET { public HEADER header; public int scaleRel; }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct DEVMODE
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string dmDeviceName;
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
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public string dmFormName;
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

        [DllImport("user32.dll")]
        public static extern int GetDisplayConfigBufferSizes(uint flags, out uint numPaths, out uint numModes);

        [DllImport("user32.dll")]
        public static extern int QueryDisplayConfig(uint flags, ref uint numPaths, [Out] PATH_INFO[] paths,
            ref uint numModes, [Out] MODE_INFO[] modes, IntPtr currentTopologyId);

        [DllImport("user32.dll")]
        public static extern int SetDisplayConfig(uint numPaths, [In] PATH_INFO[] paths,
            uint numModes, [In] MODE_INFO[] modes, uint flags);

        [DllImport("user32.dll", EntryPoint = "DisplayConfigGetDeviceInfo")]
        public static extern int GetSourceName(ref SOURCE_NAME info);

        [DllImport("user32.dll", EntryPoint = "DisplayConfigGetDeviceInfo")]
        public static extern int GetTargetName(ref TARGET_NAME info);

        [DllImport("user32.dll", EntryPoint = "DisplayConfigGetDeviceInfo")]
        public static extern int GetDpi(ref DPI_GET info);

        [DllImport("user32.dll", EntryPoint = "DisplayConfigSetDeviceInfo")]
        public static extern int SetDpi(ref DPI_SET info);

        [DllImport("user32.dll", EntryPoint = "EnumDisplaySettingsW", CharSet = CharSet.Unicode)]
        public static extern bool EnumDisplaySettings(string deviceName, int modeNum, ref DEVMODE devMode);

        [DllImport("user32.dll", EntryPoint = "ChangeDisplaySettingsExW", CharSet = CharSet.Unicode)]
        public static extern int ChangeDisplaySettingsEx(string deviceName, ref DEVMODE devMode, IntPtr hwnd, uint flags, IntPtr lParam);
    }

    public static class Display
    {
        // The scaling steps Windows offers, in order. Index |minScaleRel| is the "recommended" one.
        static readonly int[] Steps = { 100, 125, 150, 175, 200, 225, 250, 300, 350, 400, 450, 500 };
        static readonly string[] OrientationNames = { "landscape", "portrait", "landscape-flipped", "portrait-flipped" };

        class Entry
        {
            public MonitorInfo Info;
            public Native.LUID Adapter;
            public uint SourceId;
        }

        static void QueryActive(out Native.PATH_INFO[] paths, out Native.MODE_INFO[] modes)
        {
            for (int attempt = 0; attempt < 5; attempt++)
            {
                uint numPaths, numModes;
                int err = Native.GetDisplayConfigBufferSizes(Native.QDC_ONLY_ACTIVE_PATHS, out numPaths, out numModes);
                if (err != 0) throw new InvalidOperationException("GetDisplayConfigBufferSizes failed: " + err);

                paths = new Native.PATH_INFO[numPaths];
                modes = new Native.MODE_INFO[numModes];
                err = Native.QueryDisplayConfig(Native.QDC_ONLY_ACTIVE_PATHS, ref numPaths, paths, ref numModes, modes, IntPtr.Zero);
                if (err == Native.ERROR_INSUFFICIENT_BUFFER) continue; // config changed between calls; retry
                if (err != 0) throw new InvalidOperationException("QueryDisplayConfig failed: " + err);

                Array.Resize(ref paths, (int)numPaths);
                Array.Resize(ref modes, (int)numModes);
                return;
            }
            throw new InvalidOperationException("QueryDisplayConfig kept changing; try again.");
        }

        static Native.PATH_INFO[] QueryActivePaths()
        {
            Native.PATH_INFO[] paths;
            Native.MODE_INFO[] modes;
            QueryActive(out paths, out modes);
            return paths;
        }

        static Native.DEVMODE CurrentMode(string device)
        {
            var dm = new Native.DEVMODE();
            dm.dmSize = (short)Marshal.SizeOf(typeof(Native.DEVMODE));
            if (!Native.EnumDisplaySettings(device, Native.ENUM_CURRENT_SETTINGS, ref dm))
                throw new InvalidOperationException("Could not read current settings for " + device);
            return dm;
        }

        static bool TryGetDpi(Native.LUID adapter, uint sourceId, out Native.DPI_GET dpi)
        {
            dpi = new Native.DPI_GET();
            dpi.header.type = -3; // DISPLAYCONFIG_DEVICE_INFO_GET_DPI_SCALE
            dpi.header.size = (uint)Marshal.SizeOf(typeof(Native.DPI_GET));
            dpi.header.adapterId = adapter;
            dpi.header.id = sourceId;
            return Native.GetDpi(ref dpi) == 0;
        }

        static List<Entry> Enumerate()
        {
            var result = new List<Entry>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var p in QueryActivePaths())
            {
                var src = new Native.SOURCE_NAME();
                src.header.type = 1; // GET_SOURCE_NAME
                src.header.size = (uint)Marshal.SizeOf(typeof(Native.SOURCE_NAME));
                src.header.adapterId = p.sourceInfo.adapterId;
                src.header.id = p.sourceInfo.id;
                if (Native.GetSourceName(ref src) != 0) continue;
                if (!seen.Add(src.viewGdiDeviceName)) continue; // duplicated (mirrored) display

                var tgt = new Native.TARGET_NAME();
                tgt.header.type = 2; // GET_TARGET_NAME
                tgt.header.size = (uint)Marshal.SizeOf(typeof(Native.TARGET_NAME));
                tgt.header.adapterId = p.targetInfo.adapterId;
                tgt.header.id = p.targetInfo.id;
                string name = (Native.GetTargetName(ref tgt) == 0 && !string.IsNullOrEmpty(tgt.monitorFriendlyDeviceName))
                    ? tgt.monitorFriendlyDeviceName : "(no name)";

                var dm = CurrentMode(src.viewGdiDeviceName);

                int scaling = 0;
                Native.DPI_GET dpi;
                if (TryGetDpi(p.sourceInfo.adapterId, p.sourceInfo.id, out dpi))
                {
                    int idx = Math.Abs(dpi.minScaleRel) + dpi.curScaleRel;
                    if (idx >= 0 && idx < Steps.Length) scaling = Steps[idx];
                }

                result.Add(new Entry
                {
                    Adapter = p.sourceInfo.adapterId,
                    SourceId = p.sourceInfo.id,
                    Info = new MonitorInfo
                    {
                        Name = name,
                        Device = src.viewGdiDeviceName,
                        X = dm.dmPositionX,
                        Y = dm.dmPositionY,
                        Primary = dm.dmPositionX == 0 && dm.dmPositionY == 0,
                        Width = dm.dmPelsWidth,
                        Height = dm.dmPelsHeight,
                        RefreshHz = dm.dmDisplayFrequency,
                        Orientation = (dm.dmDisplayOrientation >= 0 && dm.dmDisplayOrientation < 4)
                            ? OrientationNames[dm.dmDisplayOrientation] : dm.dmDisplayOrientation.ToString(),
                        Scaling = scaling
                    }
                });
            }

            result.Sort((a, b) => a.Info.X != b.Info.X ? a.Info.X.CompareTo(b.Info.X) : a.Info.Y.CompareTo(b.Info.Y));
            return result;
        }

        static Entry Find(List<Entry> all, string device)
        {
            foreach (var e in all)
                if (string.Equals(e.Info.Device, device, StringComparison.OrdinalIgnoreCase)) return e;
            throw new ArgumentException("Monitor not found: " + device);
        }

        static string DispChangeText(int code)
        {
            switch (code)
            {
                case 1: return "a restart is required";
                case -1: return "the display driver failed the change";
                case -2: return "that mode isn't supported by this monitor";
                case -3: return "the settings couldn't be saved";
                case -4: return "invalid flags";
                case -5: return "invalid parameter";
                case -6: return "invalid DualView setting";
                default: return "error " + code;
            }
        }

        public static MonitorInfo[] GetMonitors()
        {
            var all = Enumerate();
            var list = new MonitorInfo[all.Count];
            for (int i = 0; i < all.Count; i++) list[i] = all[i].Info;
            return list;
        }

        static bool SameAdapter(Native.LUID a, Native.LUID b)
        {
            return a.LowPart == b.LowPart && a.HighPart == b.HighPart;
        }

        /// Makes the monitor primary by moving every monitor by the same offset, so the
        /// chosen one lands at (0,0) and the arrangement is kept. Uses SetDisplayConfig,
        /// which applies the whole layout in one change (as Settings > Display does).
        public static void SetPrimary(string device)
        {
            var target = Find(Enumerate(), device);
            if (target.Info.Primary) return;

            Native.PATH_INFO[] paths;
            Native.MODE_INFO[] modes;
            QueryActive(out paths, out modes);

            int idx = -1;
            for (int i = 0; i < modes.Length; i++)
            {
                if (modes[i].infoType == Native.MODE_INFO_TYPE_SOURCE &&
                    modes[i].id == target.SourceId &&
                    SameAdapter(modes[i].adapterId, target.Adapter)) { idx = i; break; }
            }
            if (idx < 0) throw new InvalidOperationException("Could not find the display layout entry for " + target.Info.Name);

            int dx = modes[idx].sourceX;
            int dy = modes[idx].sourceY;
            for (int i = 0; i < modes.Length; i++)
            {
                if (modes[i].infoType != Native.MODE_INFO_TYPE_SOURCE) continue;
                modes[i].sourceX -= dx;
                modes[i].sourceY -= dy;
            }

            uint flags = Native.SDC_APPLY | Native.SDC_USE_SUPPLIED_DISPLAY_CONFIG |
                         Native.SDC_SAVE_TO_DATABASE | Native.SDC_ALLOW_CHANGES;
            int err = Native.SetDisplayConfig((uint)paths.Length, paths, (uint)modes.Length, modes, flags);
            if (err != 0)
                throw new InvalidOperationException("Changing the primary monitor failed: " +
                    new System.ComponentModel.Win32Exception(err).Message + " (error " + err + ")");
        }

        /// Changes one monitor's resolution and/or orientation, keeping its position.
        /// width/height 0 = keep; orientation -1 = keep.
        public static void SetMode(string device, int width, int height, int orientation)
        {
            var target = Find(Enumerate(), device);
            var dm = CurrentMode(target.Info.Device);
            dm.dmFields = Native.DM_POSITION; // include position so Windows doesn't re-place the display

            bool swap = false;
            if (orientation >= 0 && orientation != dm.dmDisplayOrientation)
            {
                swap = (orientation % 2) != (dm.dmDisplayOrientation % 2);
                dm.dmDisplayOrientation = orientation;
                dm.dmFields |= Native.DM_DISPLAYORIENTATION;
            }
            if (width > 0 && height > 0)
            {
                dm.dmPelsWidth = width;
                dm.dmPelsHeight = height;
                dm.dmFields |= Native.DM_PELSWIDTH | Native.DM_PELSHEIGHT;
            }
            else if (swap)
            {
                int w = dm.dmPelsWidth; dm.dmPelsWidth = dm.dmPelsHeight; dm.dmPelsHeight = w;
                dm.dmFields |= Native.DM_PELSWIDTH | Native.DM_PELSHEIGHT;
            }

            int r = Native.ChangeDisplaySettingsEx(target.Info.Device, ref dm, IntPtr.Zero, Native.CDS_UPDATEREGISTRY, IntPtr.Zero);
            if (r != 0)
                throw new InvalidOperationException("Changing " + target.Info.Name + " (" + target.Info.Device + ") failed: " + DispChangeText(r));
        }

        public static void SetScaling(string device, int percent)
        {
            var target = Find(Enumerate(), device);
            Native.DPI_GET dpi;
            if (!TryGetDpi(target.Adapter, target.SourceId, out dpi))
                throw new InvalidOperationException("Could not read scaling for " + target.Info.Name);

            int recommended = Math.Abs(dpi.minScaleRel);
            int idx = Array.IndexOf(Steps, percent);
            int rel = idx - recommended;
            if (idx < 0 || rel < dpi.minScaleRel || rel > dpi.maxScaleRel)
            {
                int hi = Math.Min(recommended + dpi.maxScaleRel, Steps.Length - 1);
                throw new ArgumentException(target.Info.Name + " supports scaling from 100% to " + Steps[hi] + "%.");
            }

            var set = new Native.DPI_SET();
            set.header.type = -4; // DISPLAYCONFIG_DEVICE_INFO_SET_DPI_SCALE
            set.header.size = (uint)Marshal.SizeOf(typeof(Native.DPI_SET));
            set.header.adapterId = target.Adapter;
            set.header.id = target.SourceId;
            set.scaleRel = rel;
            int err = Native.SetDpi(ref set);
            if (err != 0) throw new InvalidOperationException("Setting scaling failed: error " + err);
        }
    }
}
'@
}

$monitors = [MonitorTools.Display]::GetMonitors()

if ($List) {
    $monitors | Format-Table Name, Device, Primary, X, Y, Width, Height, RefreshHz, Orientation, Scaling -AutoSize
    return
}

# Work out which monitor is meant: exact device or name first, then a partial name.
$key = if ($Monitor -match '^\d+$') { "\\.\DISPLAY$Monitor" } else { $Monitor }
$found = @($monitors | Where-Object { $_.Device -eq $key -or $_.Name -eq $key })
if ($found.Count -eq 0) { $found = @($monitors | Where-Object { $_.Name -like "*$key*" }) }
if ($found.Count -ne 1) {
    $why = if ($found.Count -eq 0) { 'No monitor matches' } else { 'More than one monitor matches' }
    $names = ($monitors | ForEach-Object { "  $($_.Name)  ($($_.Device))" }) -join "`n"
    throw "$why '$Monitor'. Connected monitors:`n$names"
}
$target = $found[0]

$width = 0; $height = 0
if ($Resolution) { $width, $height = [int[]]($Resolution -split 'x') }

$orientationValue = -1
if ($Orientation) {
    $orientationValue = @{ 'landscape' = 0; 'portrait' = 1; 'landscape-flipped' = 2; 'portrait-flipped' = 3 }[$Orientation]
}

# Only change what actually differs (avoids needless flicker).
if ($SetPrimary -and -not $target.Primary) {
    [MonitorTools.Display]::SetPrimary($target.Device)
}

$needsMode = ($Resolution -and ($width -ne $target.Width -or $height -ne $target.Height)) -or
             ($Orientation -and $Orientation -ne $target.Orientation)
if ($needsMode) {
    [MonitorTools.Display]::SetMode($target.Device, $width, $height, $orientationValue)
}

if ($Scaling -and $Scaling -ne $target.Scaling) {
    [MonitorTools.Display]::SetScaling($target.Device, $Scaling)
}

Write-Host "Updated $($target.Name) ($($target.Device))."
