namespace XlBlocks.AddIn.Gui;

partial class AboutForm
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
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
        panelAbout = new Panel();
        materialLabelVersion = new ReaLTaiizor.Controls.MaterialLabel();
        materialButtonClose = new ReaLTaiizor.Controls.MaterialButton();
        materialLabelCopyright = new ReaLTaiizor.Controls.MaterialLabel();
        materialLabelInfo = new ReaLTaiizor.Controls.MaterialLabel();
        materialLabelAbout = new ReaLTaiizor.Controls.MaterialLabel();
        pictureBoxIcon = new PictureBox();
        panelAbout.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)pictureBoxIcon).BeginInit();
        SuspendLayout();
        // 
        // panelAbout
        // 
        panelAbout.Controls.Add(materialLabelVersion);
        panelAbout.Controls.Add(materialButtonClose);
        panelAbout.Controls.Add(materialLabelCopyright);
        panelAbout.Controls.Add(materialLabelInfo);
        panelAbout.Controls.Add(materialLabelAbout);
        panelAbout.Controls.Add(pictureBoxIcon);
        panelAbout.Dock = DockStyle.Fill;
        panelAbout.Location = new Point(3, 28);
        panelAbout.Name = "panelAbout";
        panelAbout.Size = new Size(473, 225);
        panelAbout.TabIndex = 0;
        // 
        // materialLabelVersion
        // 
        materialLabelVersion.Anchor = AnchorStyles.None;
        materialLabelVersion.Depth = 0;
        materialLabelVersion.Font = new Font("Roboto", 10F, FontStyle.Regular, GraphicsUnit.Pixel);
        materialLabelVersion.FontType = ReaLTaiizor.Manager.MaterialSkinManager.FontType.Overline;
        materialLabelVersion.Location = new Point(26, 196);
        materialLabelVersion.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelVersion.Name = "materialLabelVersion";
        materialLabelVersion.Size = new Size(130, 23);
        materialLabelVersion.TabIndex = 6;
        materialLabelVersion.Text = "xlBlocks AddIn v1.0.0.1";
        materialLabelVersion.TextAlign = ContentAlignment.MiddleCenter;
        // 
        // materialButtonClose
        // 
        materialButtonClose.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
        materialButtonClose.AutoSize = false;
        materialButtonClose.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        materialButtonClose.Density = ReaLTaiizor.Controls.MaterialButton.MaterialButtonDensity.Default;
        materialButtonClose.Depth = 0;
        materialButtonClose.HighEmphasis = true;
        materialButtonClose.Icon = null;
        materialButtonClose.IconType = ReaLTaiizor.Controls.MaterialButton.MaterialIconType.Rebase;
        materialButtonClose.Location = new Point(371, 187);
        materialButtonClose.Margin = new Padding(4, 6, 4, 6);
        materialButtonClose.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialButtonClose.Name = "materialButtonClose";
        materialButtonClose.NoAccentTextColor = Color.Empty;
        materialButtonClose.Size = new Size(88, 32);
        materialButtonClose.TabIndex = 5;
        materialButtonClose.Text = "Close";
        materialButtonClose.Type = ReaLTaiizor.Controls.MaterialButton.MaterialButtonType.Contained;
        materialButtonClose.UseAccentColor = false;
        materialButtonClose.UseVisualStyleBackColor = true;
        materialButtonClose.Click += materialButtonClose_Click;
        // 
        // materialLabelCopyright
        // 
        materialLabelCopyright.Anchor = AnchorStyles.None;
        materialLabelCopyright.Depth = 0;
        materialLabelCopyright.Font = new Font("Roboto", 12F, FontStyle.Regular, GraphicsUnit.Pixel);
        materialLabelCopyright.FontType = ReaLTaiizor.Manager.MaterialSkinManager.FontType.Caption;
        materialLabelCopyright.Location = new Point(162, 162);
        materialLabelCopyright.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelCopyright.Name = "materialLabelCopyright";
        materialLabelCopyright.Size = new Size(283, 19);
        materialLabelCopyright.TabIndex = 4;
        materialLabelCopyright.Text = "Copyright (C) 2024 by Charlie Friedemann";
        materialLabelCopyright.TextAlign = ContentAlignment.MiddleCenter;
        // 
        // materialLabelInfo
        // 
        materialLabelInfo.Anchor = AnchorStyles.None;
        materialLabelInfo.Depth = 0;
        materialLabelInfo.Font = new Font("Roboto", 14F, FontStyle.Regular, GraphicsUnit.Pixel);
        materialLabelInfo.Location = new Point(162, 76);
        materialLabelInfo.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelInfo.Name = "materialLabelInfo";
        materialLabelInfo.Size = new Size(283, 86);
        materialLabelInfo.TabIndex = 3;
        materialLabelInfo.Text = "xlBlocks is an Excel AddIn that can be used to create functional spreadsheets. To learn more and see examples, read the documentation.";
        materialLabelInfo.TextAlign = ContentAlignment.MiddleCenter;
        // 
        // materialLabelAbout
        // 
        materialLabelAbout.Anchor = AnchorStyles.None;
        materialLabelAbout.AutoSize = true;
        materialLabelAbout.Depth = 0;
        materialLabelAbout.Font = new Font("Roboto", 34F, FontStyle.Bold, GraphicsUnit.Pixel);
        materialLabelAbout.FontType = ReaLTaiizor.Manager.MaterialSkinManager.FontType.H4;
        materialLabelAbout.Location = new Point(205, 35);
        materialLabelAbout.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelAbout.Name = "materialLabelAbout";
        materialLabelAbout.Size = new Size(201, 41);
        materialLabelAbout.TabIndex = 2;
        materialLabelAbout.Text = "About xlBlocks";
        // 
        // pictureBoxIcon
        // 
        pictureBoxIcon.Anchor = AnchorStyles.None;
        pictureBoxIcon.Image = Properties.Resources.xlbrew_icon_full;
        pictureBoxIcon.Location = new Point(26, 47);
        pictureBoxIcon.Name = "pictureBoxIcon";
        pictureBoxIcon.Size = new Size(130, 130);
        pictureBoxIcon.SizeMode = PictureBoxSizeMode.Zoom;
        pictureBoxIcon.TabIndex = 0;
        pictureBoxIcon.TabStop = false;
        // 
        // AboutForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(478, 255);
        Controls.Add(panelAbout);
        FormStyle = ReaLTaiizor.Enum.Material.FormStyles.ActionBar_None;
        Icon = (Icon)resources.GetObject("$this.Icon");
        Margin = new Padding(2);
        Name = "AboutForm";
        Padding = new Padding(3, 28, 2, 2);
        StartPosition = FormStartPosition.CenterScreen;
        Text = "About xlBlocks";
        FormClosing += AboutForm_FormClosing;
        Load += AboutForm_Load;
        KeyDown += AboutForm_KeyDown;
        panelAbout.ResumeLayout(false);
        panelAbout.PerformLayout();
        ((System.ComponentModel.ISupportInitialize)pictureBoxIcon).EndInit();
        ResumeLayout(false);
    }

    #endregion

    private Panel panelAbout;
    private PictureBox pictureBoxIcon;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelAbout;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelInfo;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelCopyright;
    private ReaLTaiizor.Controls.MaterialButton materialButtonClose;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelVersion;
}
