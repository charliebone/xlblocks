namespace XlBlocks.AddIn.Gui;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ReaLTaiizor.Forms;
using XlBlocks.AddIn.Dna;

internal partial class ErrorLogViewerForm : MaterialForm
{
    [DllImport("user32.dll")]
    static extern IntPtr SetFocus(IntPtr hWnd);

    public ErrorLogViewerForm()
    {
        InitializeComponent();
    }

    public void LogExecutionExceptions(IList<ExecutionExceptionEventArgs> eventArgsList)
    {
        if (eventArgsList.Count == 0)
            return;

        foreach (var eventArg in eventArgsList)
            materialListViewExceptions.Items.Add(new ListViewItem(new[]
            {
                eventArg.Timestamp.ToString("yyyy-MM-dd HH:mm:ss"),
                eventArg.Function,
                eventArg.CellReference,
                FormatArgList(eventArg.Arguments),
                eventArg.Exception.Message
            }));

        //materialListViewExceptions.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent);
        materialListViewExceptions.Items[^1].EnsureVisible();
        materialListViewExceptions.Refresh();
    }

    private static string FormatArgList(object[]? args)
    {
        if (args is null)
            return "";

        var argStr = string.Join(", ", args.Select(arg =>
        {
            if (arg is ExcelEmpty || arg is ExcelMissing)
                return "(missing)";
            if (arg is string strArg)
                return $"'{strArg}'";
            return arg.ToString();
        }));
        return $"[ {argStr} ]";
    }

    private void ErrorLogViewerForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        try
        {
            // try to set focus back to excel
            SetFocus(ExcelDnaUtil.WindowHandle);
        }
        catch { }
    }

    private void ErrorLogViewerForm_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
            Close();
    }
}
