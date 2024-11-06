namespace XlBlocks.AddIn.Dna;

using System;
using System.Reactive.Linq;
using ExcelDna.Integration;
using ExcelDna.Registration;
using NLog;
using XlBlocks.AddIn.Cache;

internal delegate void ExecutionExceptionEventHandler(object sender, ExecutionExceptionEventArgs e);

internal class ExecutionExceptionEventArgs : EventArgs
{
    public Exception Exception { get; }
    public DateTime Timestamp { get; }
    public string Function { get; }
    public object[]? Arguments { get; }
    public string CellReference { get; }

    public ExecutionExceptionEventArgs(Exception exception, DateTime timestamp, string function, object[]? arguments, string cellReference)
    {
        Exception = exception;
        Timestamp = timestamp;
        Function = function;
        Arguments = arguments;
        CellReference = cellReference;
    }
}

internal class UnhandledDnaException : Exception
{
    public UnhandledDnaException(object? exceptionObj) :
        base($"Unhandled exception in XlBlocks AddIn: {(exceptionObj as Exception)?.Message ?? exceptionObj}", exceptionObj as Exception)
    {

    }
}

internal class XlBlocksExecutionHandler : FunctionExecutionHandler
{
    private static readonly Logger _logger = LogManager.GetLogger(typeof(XlBlocksExecutionHandler).FullName);

    private static readonly Lazy<XlBlocksExecutionHandler> _handler = new(() => new XlBlocksExecutionHandler());

    public static XlBlocksExecutionHandler Handler => _handler.Value;

    public event ExecutionExceptionEventHandler? ExecutionException;

    public static bool PrintExceptions { get; set; }
    internal static FunctionExecutionHandler ExceptionHandlerSelector(ExcelFunctionRegistration _)
    {
        // every function registration uses the same handler
        return _handler.Value;
    }

    private void OnExecutionException(FunctionExecutionArgs args)
    {
        if (ExecutionException is null)
            return;

        var callingReference = CacheHelper.GetCallingReference(() => "Unknown");
        var eventArgs = new ExecutionExceptionEventArgs(args.Exception, DateTime.Now, args.FunctionName, args.Arguments, callingReference);
        ExecutionException.Invoke(this, eventArgs);
    }

    public override void OnEntry(FunctionExecutionArgs args)
    {
        _logger.Trace($"{args.FunctionName} :: OnEntry - Args: {string.Join(",", args.Arguments.Select(arg => arg.ToString()))}");
    }

    public override void OnSuccess(FunctionExecutionArgs args)
    {
        _logger.Trace($"{args.FunctionName} :: OnSuccess - Result: {args.ReturnValue}");

        // default to #N/A for null value returns instead of #NUM
        args.ReturnValue ??= ExcelError.ExcelErrorNA;
    }

    public override void OnException(FunctionExecutionArgs args)
    {
        _logger.Trace($"{args.FunctionName} :: OnException - Message: {args.Exception}");

        OnExecutionException(args);

        args.ReturnValue = PrintExceptions ? args.Exception.Message : ExcelError.ExcelErrorNA;
        args.FlowBehavior = FlowBehavior.Return;
    }

    public override void OnExit(FunctionExecutionArgs args)
    {
        _logger.Trace($"{args.FunctionName} :: OnExit");
    }

    public static object OnUnhandledExceptionHandler(object obj)
    {
        var unhandledException = new UnhandledDnaException(obj);
        _logger.Error(unhandledException, unhandledException.Message);

        if (Handler.ExecutionException is not null)
        {
            var eventArgs = new ExecutionExceptionEventArgs(unhandledException.InnerException ?? unhandledException,
                DateTime.Now, "(Unhandled Exception)", null, string.Empty);
            Handler.ExecutionException.Invoke(Handler, eventArgs);
        }

        return PrintExceptions ? unhandledException.Message : ExcelError.ExcelErrorNA;
    }
}
