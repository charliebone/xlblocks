namespace XlBlocks.AddIn.Gui;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ReaLTaiizor.Forms;
using XlBlocks.AddIn.Dna;

internal partial class CacheViewerForm : MaterialForm
{
    [DllImport("user32.dll")]
    static extern IntPtr SetFocus(IntPtr hWnd);

    public CacheViewerForm()
    {
        InitializeComponent();

        PopulateListViewCache();
    }

    private void PopulateListViewCache()
    {
        var itemInfoList = XlBlocksAddIn.Instance.Cache.GetItemInfo();
        materialListViewCache.Items.Clear();
        foreach (var info in itemInfoList)
        {
            materialListViewCache.Items.Add(new ListViewItem(new[] { info.Reference, info.HexKey, info.Type.Name }));
        }

        materialListViewCache.Refresh();
    }

    public void NotifyCacheInvalidated()
    {
        if (IsDisposed)
            return;

        PopulateListViewCache();
    }

    private void CacheViewerForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        try
        {
            // try to set focus back to excel
            SetFocus(ExcelDnaUtil.WindowHandle);
        }
        catch { }
    }

    private void CacheViewerForm_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
            Close();
    }
}
