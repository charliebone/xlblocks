namespace XlBlocks.AddIn.Gui;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ReaLTaiizor.Forms;

internal partial class CachedObjectViewerForm : MaterialForm
{
    [DllImport("user32.dll")]
    static extern IntPtr SetFocus(IntPtr hWnd);

    private readonly object _cachedObject;
    public CachedObjectViewerForm(object cachedObject)
    {
        InitializeComponent();

        _cachedObject = cachedObject;
    }

    private void CachedObjectViewerForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        try
        {
            // try to set focus back to excel
            SetFocus(ExcelDnaUtil.WindowHandle);
        }
        catch { }
    }

    private void CachedObjectViewerForm_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
            Close();
    }
}
