namespace XlBlocks.AddIn.Gui;

partial class CacheViewerForm
{
    /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing && (components != null))
        {
            components.Dispose();
        }
        base.Dispose(disposing);
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(CacheViewerForm));
        panel1 = new Panel();
        materialListViewCache = new ReaLTaiizor.Controls.MaterialListView();
        columnHeaderReference = new ColumnHeader();
        columnHeaderKey = new ColumnHeader();
        columnHeaderType = new ColumnHeader();
        panel1.SuspendLayout();
        SuspendLayout();
        // 
        // panel1
        // 
        panel1.CausesValidation = false;
        panel1.Controls.Add(materialListViewCache);
        panel1.Dock = DockStyle.Fill;
        panel1.Location = new Point(3, 64);
        panel1.Name = "panel1";
        panel1.Size = new Size(794, 383);
        panel1.TabIndex = 0;
        // 
        // materialListViewCache
        // 
        materialListViewCache.AutoSizeTable = false;
        materialListViewCache.BackColor = Color.FromArgb(255, 255, 255);
        materialListViewCache.BorderStyle = BorderStyle.None;
        materialListViewCache.Columns.AddRange(new ColumnHeader[] { columnHeaderReference, columnHeaderKey, columnHeaderType });
        materialListViewCache.Depth = 0;
        materialListViewCache.Dock = DockStyle.Fill;
        materialListViewCache.FullRowSelect = true;
        materialListViewCache.Location = new Point(0, 0);
        materialListViewCache.MinimumSize = new Size(200, 100);
        materialListViewCache.MouseLocation = new Point(-1, -1);
        materialListViewCache.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.OUT;
        materialListViewCache.Name = "materialListViewCache";
        materialListViewCache.OwnerDraw = true;
        materialListViewCache.Size = new Size(794, 383);
        materialListViewCache.TabIndex = 0;
        materialListViewCache.UseCompatibleStateImageBehavior = false;
        materialListViewCache.View = View.Details;
        // 
        // columnHeaderReference
        // 
        columnHeaderReference.Text = "Cell";
        columnHeaderReference.Width = 180;
        // 
        // columnHeaderKey
        // 
        columnHeaderKey.Text = "Key";
        columnHeaderKey.Width = 200;
        // 
        // columnHeaderType
        // 
        columnHeaderType.Text = "Type";
        columnHeaderType.Width = 200;
        // 
        // CacheViewerForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(800, 450);
        Controls.Add(panel1);
        DoubleBuffered = false;
        Icon = XlBlocks.AddIn.Properties.Resources.xlblocks_icon;
        Name = "CacheViewerForm";
        Text = "Cache Viewer";
        FormClosing += CacheViewerForm_FormClosing;
        KeyDown += CacheViewerForm_KeyDown;
        panel1.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Panel panel1;
    private ReaLTaiizor.Controls.MaterialListView materialListViewCache;
    private ColumnHeader columnHeaderReference;
    private ColumnHeader columnHeaderKey;
    private ColumnHeader columnHeaderType;
}
