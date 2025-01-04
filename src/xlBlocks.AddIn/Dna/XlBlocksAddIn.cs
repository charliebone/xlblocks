namespace XlBlocks.AddIn.Dna;

using System;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using ExcelDna.Integration;
using ExcelDna.IntelliSense;
using ExcelDna.Registration;
using NLog;
using XlBlocks.AddIn.Cache;
using XlBlocks.AddIn.Gui;

internal class XlBlocksAddIn : IExcelAddIn
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(XlBlocksAddIn).FullName);

    // a bit hacky but we know that the ExcelAgent addin will be loaded when running integration tests...
    private static readonly Regex _excelAgentXllRegex = new(@"^ExcelAgent(64)?\.xll$", RegexOptions.Compiled);
    internal static bool ShouldRegisterTestFunctions => Debugger.IsAttached || Environment.GetCommandLineArgs().Any(x => _excelAgentXllRegex.IsMatch(x));

    private readonly Lazy<ObjectCache> _objectCache = new();
    public ObjectCache Cache
    {
        get => _objectCache.Value;
    }

    private static XlBlocksAddIn? _instance;
    public static XlBlocksAddIn Instance
    {
        get => _instance ?? throw new Exception("Cannot access XlBlocksAddIn.Instance prior to AutoOpen()");
    }

    private static string? _productVersion;
    public static string ProductVersion
    {
        get
        {
            if (_productVersion == null)
            {
                var assemblyInfoVersion = Assembly.GetExecutingAssembly()
                    .GetCustomAttributes<AssemblyInformationalVersionAttribute>()
                    .Select(x => x?.InformationalVersion)
                    .FirstOrDefault();
                _productVersion = assemblyInfoVersion ?? Version?.ToString() ?? string.Empty;
            }
            return _productVersion;
        }
    }
    public static string? Version => Assembly.GetExecutingAssembly().GetName().Version?.ToString();

    public static bool PrintExceptions
    {
        get => XlBlocksExecutionHandler.PrintExceptions;
        set
        {
            XlBlocksExecutionHandler.PrintExceptions = value;
            GuiHelper.InvalidateRibbon();
        }
    }

    public void AutoOpen()
    {
        _logger.Debug("XlBlocks AddIn AutoOpen");
        _instance = this;

        // register unhandled exception logger
        ExcelIntegration.RegisterUnhandledExceptionHandler(XlBlocksExecutionHandler.OnUnhandledExceptionHandler);

        // register excel functions with parameter conversions and function handlers
        _logger.Debug($"Register test functions: {ShouldRegisterTestFunctions}");
        RegistrationUtilities.GetExcelFunctions(ShouldRegisterTestFunctions)
            .ProcessParameterConversions(GetInputParameterConversions())
            .ProcessCacheAwareParamsRegistrations()
            .ProcessParameterConversions(GetReturnParameterConversions())
            .ProcessFunctionExecutionHandlers(GetFunctionExecutionHandlerConfig())
            .RegisterFunctions();

        // start intellisense server
        IntelliSenseServer.Install();

        // initialize winforms
        GuiHelper.InitializeWinForms();
    }

    public void AutoClose()
    {
        _logger.Debug("XlBlocks AddIn AutoClose");

        // stop intellisense server
        IntelliSenseServer.Uninstall();
    }

    private static ParameterConversionConfiguration GetInputParameterConversions()
    {
        return new ParameterConversionConfiguration()
            .AddParameterConversion(CachedParameterConversion.GetCachedInputParameterConversion())
            .AddParameterConversion(BaseParameterConversion.GetBaseInputParameterConversion());
    }

    private static ParameterConversionConfiguration GetReturnParameterConversions()
    {
        return new ParameterConversionConfiguration()
            .AddReturnConversion(CachedParameterConversion.GetCachedReturnParameterConversion())
            .AddReturnConversion(ArrayReturnConversion.GetArrayReturnParameterConversion())
            .AddReturnConversion(BaseParameterConversion.GetBaseReturnParameterConversion());
    }

    private static FunctionExecutionConfiguration GetFunctionExecutionHandlerConfig()
    {
        return new FunctionExecutionConfiguration()
            .AddFunctionExecutionHandler(XlBlocksExecutionHandler.ExceptionHandlerSelector);
    }
}
