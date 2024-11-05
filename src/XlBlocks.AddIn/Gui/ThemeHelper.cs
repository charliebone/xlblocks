namespace XlBlocks.AddIn.Gui;

using System;
using System.Management;
using System.Security.Principal;
using Microsoft.Win32;
using NLog;
using ReaLTaiizor.Colors;
using ReaLTaiizor.Manager;
using ReaLTaiizor.Util;

internal class ThemeHelper
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(GuiHelper).FullName);

    private ManagementEventWatcher? _watcher;

    public static int GetOfficeVersion()
    {
        try
        {
            var rk = Registry.ClassesRoot.OpenSubKey(@"Excel.Application\\CurVer");

            var officeVersion = rk?.GetValue("")?.ToString();
            var officeNumberVersion = officeVersion?.Split('.')[officeVersion.Split('.').GetUpperBound(0)] ?? "0";
            return Int32.Parse(officeNumberVersion);
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Unable to determine current office version");
            return 0;
        }
    }

    public static OfficeTheme QueryCurrentOfficeTheme()
    {
        var rk = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{GetOfficeVersion():F1}\Common");
        if (rk == null) return OfficeTheme.Colorful;

        return ToOfficeTheme(rk.GetValue("UI Theme", OfficeTheme.Colorful));
    }

    public static SystemTheme QueryCurrentSystemTheme()
    {
        var rk = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize");
        if (rk == null) return SystemTheme.Light;

        return ToSystemTheme(rk.GetValue("AppsUseLightTheme", SystemTheme.Light));
    }

    public void StartWatching(Action<EventArrivedEventArgs> eventHandler)
    {
        var currentUser = WindowsIdentity.GetCurrent();

        var queryStr = $@"SELECT * FROM RegistryValueChangeEvent 
                           WHERE Hive = 'HKEY_USERS' 
                           AND KeyPath = '{currentUser.User?.Value}\\Software\\Microsoft\\Office\\{GetOfficeVersion():F1}\\Common' 
                           AND ValueName = 'UI Theme'";

        var query = new WqlEventQuery(queryStr);

        _watcher = new ManagementEventWatcher(query);
        _watcher.EventArrived += (sender, eventArgs) => eventHandler(eventArgs);

        try
        {
            _watcher.Start();
            _logger.Debug("Theme watcher has started");
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Theme watcher failed to start");
        }
    }

    public void StopWatching()
    {
        _watcher?.Stop();
    }

    public static OfficeTheme ToOfficeTheme(object? value)
    {
        return value?.ToString() switch
        {
            "3" => OfficeTheme.DarkGrey,
            "4" => OfficeTheme.Black,
            "5" => OfficeTheme.White,
            "6" => OfficeTheme.System,
            _ => OfficeTheme.Colorful
        };
    }

    public static SystemTheme ToSystemTheme(object? value)
    {
        return value?.ToString() switch
        {
            "0" => SystemTheme.Dark,
            _ => SystemTheme.Light
        };
    }

    public enum OfficeTheme
    {
        Colorful = 0,
        DarkGrey = 3,
        Black = 4,
        White = 5,
        System = 6
    }

    public enum SystemTheme
    {
        Dark = 0,
        Light = 1
    }

    public enum XlBlocksColor
    {
        Blue = 0x015989,
        Green = 0x158901
    }

    public static MaterialColorScheme GetMaterialColorScheme(OfficeTheme officeTheme, out MaterialSkinManager.Themes theme)
    {
        if (officeTheme == OfficeTheme.System)
        {
            // use system theme to determine color scheme: Light => Colorful, Dark => Black
            var systemTheme = QueryCurrentSystemTheme();
            officeTheme = systemTheme == SystemTheme.Dark ? OfficeTheme.Black : OfficeTheme.Colorful;
        }

        theme = (officeTheme == OfficeTheme.DarkGrey || officeTheme == OfficeTheme.Black) ?
            MaterialSkinManager.Themes.DARK :
            MaterialSkinManager.Themes.LIGHT;

        var colorScheme = new MaterialColorScheme(
            (MaterialPrimary)XlBlocksColor.Blue,
            MaterialPrimary.BlueGrey800,
            MaterialPrimary.BlueGrey200,
            MaterialAccent.Blue200,
            MaterialTextShade.LIGHT);

        return colorScheme;
    }
}
