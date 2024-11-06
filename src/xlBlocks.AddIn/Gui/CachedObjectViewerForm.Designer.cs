namespace XlBlocks.AddIn.Gui;

partial class CachedObjectViewerForm
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
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(CachedObjectViewerForm));
        panel1 = new Panel();
        SuspendLayout();
        // 
        // panel1
        // 
        panel1.CausesValidation = false;
        panel1.Dock = DockStyle.Fill;
        panel1.Location = new Point(3, 28);
        panel1.Name = "panel1";
        panel1.Size = new Size(794, 419);
        panel1.TabIndex = 0;
        // 
        // CachedObjectViewerForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(800, 450);
        Controls.Add(panel1);
        DoubleBuffered = false;
        FormStyle = ReaLTaiizor.Enum.Material.FormStyles.ActionBar_None;
        Icon = XlBlocks.AddIn.Properties.Resources.xlblocks_icon;
        Name = "CachedObjectViewerForm";
        Padding = new Padding(3, 28, 3, 3);
        Text = "CacheViewerForm";
        FormClosing += CachedObjectViewerForm_FormClosing;
        KeyDown += CachedObjectViewerForm_KeyDown;
        ResumeLayout(false);
    }

    #endregion

    private Panel panel1;
}
