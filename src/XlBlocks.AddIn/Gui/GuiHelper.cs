namespace XlBlocks.AddIn.Gui;

using System;
using System.Management;
using System.Reactive.Linq;
using System.Windows.Forms;
using ExcelDna.Integration.CustomUI;
using NLog;
using ReaLTaiizor.Forms;
using ReaLTaiizor.Manager;
using XlBlocks.AddIn.Cache;
using XlBlocks.AddIn.Dna;

internal static class GuiHelper
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(GuiHelper).FullName);

    private static IRibbonUI? _ribbonUI;

    private static SynchronizationContext? _syncContext;
    private static readonly object _syncContextLock = new();

    private static readonly ThemeHelper _themeWatcher = new();
    private static readonly MaterialSkinManager _materialSkinManager = MaterialSkinManager.Instance;

    private static AboutForm? _aboutForm;
    private static CacheViewerForm? _cacheViewerForm;
    private static ErrorLogViewerForm? _errorLogViewerForm;

    private static readonly List<MaterialForm> _currentForms = new();

    public static void InitializeWinForms()
    {
        // this method is called by AutoOpen and therefore should be on the main thread
        _syncContext = SynchronizationContext.Current ?? new WindowsFormsSynchronizationContext();

        Application.EnableVisualStyles();

        _materialSkinManager.EnforceBackcolorOnAllComponents = true;

        // start theme watcher and set current theme
        _themeWatcher.StartWatching(OnThemeChanged);
        SetMaterialSkinColorTheme();
    }

    public static void RegisterRibbon(IRibbonUI ribbonUI)
    {
        _ribbonUI = ribbonUI;
    }

    private static void ShowForm(Form form)
    {
        lock (_syncContextLock)
        {
            if (form?.Visible ?? true)
                return;

            _syncContext?.Post(state => form?.Show(), null);
        }
    }

    private static void DoWithForm(Form form, Action<Form> action)
    {
        if (form == null || form.IsDisposed)
            return;

        lock (_syncContextLock)
            _syncContext?.Post(state => action(form), null);
    }

    private static void DoWithRibbon(Action<IRibbonUI> action)
    {
        if (_ribbonUI == null)
            return;

        lock (_syncContextLock)
            _syncContext?.Post(state => action(_ribbonUI), null);
    }

    public static void InvalidateRibbon()
    {
        DoWithRibbon(ribbon => ribbon.Invalidate());
    }

    public static void InvalidateRibbonControl(string controlId)
    {
        DoWithRibbon(ribbon => ribbon.InvalidateControl(controlId));
    }

    public static void ShowAboutForm()
    {
        if (_aboutForm is null)
        {
            _aboutForm = new AboutForm();
            _materialSkinManager.AddFormToManage(_aboutForm);
            _currentForms.Add(_aboutForm);
            _aboutForm.FormClosing += (sender, e) =>
            {
                _currentForms.Remove(_aboutForm);
                _aboutForm = null;
            };
            ShowForm(_aboutForm);
        }
        else
        {
            DoWithForm(_aboutForm, form => form.BringToFront());
        }
    }

    public static void ShowCacheViewerForm()
    {
        if (_cacheViewerForm is null)
        {
            if (_syncContext is null)
                return;

            _cacheViewerForm = new CacheViewerForm();
            _materialSkinManager.AddFormToManage(_cacheViewerForm);
            _currentForms.Add(_cacheViewerForm);

            var token = Observable.FromEvent<CacheInvalidatedEventHandler, EventArgs>(
                    handler => (sender, e) => handler(e),
                    h => XlBlocksAddIn.Instance.Cache.CacheInvalidated += h,
                    h => XlBlocksAddIn.Instance.Cache.CacheInvalidated -= h)
                .Sample(TimeSpan.FromMilliseconds(500))
                .ObserveOn(_syncContext)
                .Subscribe(x => _cacheViewerForm?.NotifyCacheInvalidated());

            _cacheViewerForm.FormClosing += (sender, e) =>
            {
                _currentForms.Remove(_cacheViewerForm);
                _materialSkinManager.RemoveFormToManage(_cacheViewerForm);

                token.Dispose();
                _cacheViewerForm = null;
            };
            ShowForm(_cacheViewerForm);
        }
        else
        {
            DoWithForm(_cacheViewerForm, form => form.BringToFront());
        }
    }

    public static void ShowErrorLogViewerForm()
    {
        if (_errorLogViewerForm is null)
        {
            if (_syncContext is null)
                return;

            _errorLogViewerForm = new ErrorLogViewerForm();
            _materialSkinManager.AddFormToManage(_errorLogViewerForm);
            _currentForms.Add(_errorLogViewerForm);

            var token = Observable.FromEvent<ExecutionExceptionEventHandler, ExecutionExceptionEventArgs>(
                    handler => (sender, e) => handler(e),
                    h => XlBlocksExecutionHandler.Handler.ExecutionException += h,
                    h => XlBlocksExecutionHandler.Handler.ExecutionException -= h)
                .Buffer(TimeSpan.FromMilliseconds(500), 100)
                .ObserveOn(_syncContext)
                .Subscribe(x =>
                {
                    _errorLogViewerForm?.LogExecutionExceptions(x);
                });

            _errorLogViewerForm.FormClosing += (sender, e) =>
            {
                _currentForms.Remove(_errorLogViewerForm);
                _materialSkinManager.RemoveFormToManage(_errorLogViewerForm);

                token.Dispose();
                _errorLogViewerForm = null;
            };
            ShowForm(_errorLogViewerForm);
        }
        else
        {
            DoWithForm(_errorLogViewerForm, form => form.BringToFront());
        }
    }

    public static void ShowCacheClearVerify()
    {
        var message = "Are you sure you want to clear the XlBlocks cache? This is usually never necessary!";
        var caption = "Clear Cache";
        var buttons = MessageBoxButtons.YesNo;

        if (MessageBox.Show(message, caption, buttons) == DialogResult.Yes)
            XlBlocksAddIn.Instance.Cache.Clear();
    }

    private static void SetMaterialSkinColorTheme()
    {
        var officeTheme = ThemeHelper.QueryCurrentOfficeTheme();
        var colorScheme = ThemeHelper.GetMaterialColorScheme(officeTheme, out var theme);

        _materialSkinManager.Theme = theme;
        _materialSkinManager.ColorScheme = colorScheme;

        _currentForms.ForEach(x => DoWithForm(x, form => form.Refresh()));
    }

    private static void OnThemeChanged(EventArrivedEventArgs eventArgs)
    {
        _logger.Debug("Watcher detected Excel theme change", eventArgs);

        SetMaterialSkinColorTheme();
    }
}
