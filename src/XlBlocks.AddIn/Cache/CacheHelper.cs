namespace XlBlocks.AddIn.Cache;

using System.Runtime.CompilerServices;
using ExcelDna.Integration;

internal static class CacheHelper
{
    public static string ExcelReferenceToString(this ExcelReference reference)
    {
        try
        {
            return $"{XlCall.Excel(XlCall.xlSheetNm, reference)}!{ColumnNumberToLetter(reference.ColumnFirst + 1)}{reference.RowFirst + 1:D}";
        }
        catch
        {
            // TODO: investigate what we should do when this happens a bit more, is silently failing okay?
            return $"RefSheet!{ColumnNumberToLetter(reference.ColumnFirst + 1)}{reference.RowFirst + 1:D}";
        }
    }

    public static string GetCallingReference(Func<string> onErrorValue)
    {
        ;
        TryGetCallingReference(out var callingReference, onErrorValue);
        return callingReference;
    }

    public static bool TryGetCallingReference(out string callingReference, Func<string> onErrorValue)
    {
        try
        {
            if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference referenceCaller)
            {
                callingReference = onErrorValue();
                return false;
            }
            else
            {
                callingReference = referenceCaller.ExcelReferenceToString();
                return true;
            }
        }
        catch
        {
            callingReference = onErrorValue();
            return false;
        }
    }

    [MethodImpl(MethodImplOptions.AggressiveInlining)]
    internal static string ColumnNumberToLetter(int columnNumber)
    {
        var columnLetter = string.Empty;
        while (columnNumber > 0)
        {
            columnNumber--;
            columnLetter = (char)('A' + columnNumber % 26) + columnLetter;
            columnNumber /= 26;
        }
        return columnLetter;
    }
}
