namespace XlBlocks.AddIn.Gui;

using System;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ReaLTaiizor.Forms;
using XlBlocks.AddIn.Dna;

public partial class AboutForm : MaterialForm
{
    [DllImport("user32.dll")]
    static extern IntPtr SetFocus(IntPtr hWnd);

    public AboutForm()
    {
        InitializeComponent();
    }

    private void AboutForm_Load(object sender, EventArgs e)
    {
        materialLabelVersion.Text = $"XlBlocks AddIn v{XlBlocksAddIn.ProductVersion}";
    }

    private void AboutForm_FormClosing(object sender, FormClosingEventArgs e)
    {
        try
        {
            // try to set focus back to excel
            SetFocus(ExcelDnaUtil.WindowHandle);
        }
        catch { }
    }

    private void AboutForm_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
            Close();
    }

    private void materialButtonClose_Click(object sender, EventArgs e)
    {
        Close();
    }
}
