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
        materialLabelInfo = new ReaLTaiizor.Controls.MaterialLabel();
        pictureBoxIcon = new PictureBox();
        panelAbout.SuspendLayout();
        ((System.ComponentModel.ISupportInitialize)pictureBoxIcon).BeginInit();
        SuspendLayout();
        // 
        // panelAbout
        // 
        panelAbout.Controls.Add(materialButtonClose);
        panelAbout.Controls.Add(materialLabelInfo);
        panelAbout.Controls.Add(pictureBoxIcon);
        panelAbout.Controls.Add(materialLabelVersion);
        panelAbout.Dock = DockStyle.Fill;
        panelAbout.Location = new Point(3, 28);
        panelAbout.Name = "panelAbout";
        panelAbout.Size = new Size(451, 262);
        panelAbout.TabIndex = 0;
        // 
        // materialLabelVersion
        // 
        materialLabelVersion.Anchor = AnchorStyles.None;
        materialLabelVersion.Depth = 0;
        materialLabelVersion.Font = new Font("Microsoft Sans Serif", 12F, FontStyle.Regular, GraphicsUnit.Point);
        materialLabelVersion.FontType = ReaLTaiizor.Manager.MaterialSkinManager.FontType.Overline;
        materialLabelVersion.Location = new Point(66, 220);
        materialLabelVersion.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelVersion.Name = "materialLabelVersion";
        materialLabelVersion.Size = new Size(318, 21);
        materialLabelVersion.TabIndex = 6;
        materialLabelVersion.Text = "XlBlocks AddIn v1.0.0.1";
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
        materialButtonClose.Location = new Point(349, 224);
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
        // materialLabelInfo
        // 
        materialLabelInfo.Anchor = AnchorStyles.None;
        materialLabelInfo.Depth = 0;
        materialLabelInfo.Font = new Font("Microsoft Sans Serif", 14F, FontStyle.Regular, GraphicsUnit.Point);
        materialLabelInfo.Location = new Point(66, 134);
        materialLabelInfo.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.HOVER;
        materialLabelInfo.Name = "materialLabelInfo";
        materialLabelInfo.Size = new Size(318, 86);
        materialLabelInfo.TabIndex = 3;
        materialLabelInfo.Text = "XlBlocks is an Excel AddIn that can be used to create functional spreadsheets. To learn more and see examples, visit https://xlblocks.net.";
        materialLabelInfo.TextAlign = ContentAlignment.MiddleCenter;
        // 
        // pictureBoxIcon
        // 
        pictureBoxIcon.Anchor = AnchorStyles.None;
        pictureBoxIcon.Image = Properties.Resources.xlblocks_logo;
        pictureBoxIcon.Location = new Point(66, 1);
        pictureBoxIcon.Name = "pictureBoxIcon";
        pictureBoxIcon.Size = new Size(318, 130);
        pictureBoxIcon.SizeMode = PictureBoxSizeMode.Zoom;
        pictureBoxIcon.TabIndex = 0;
        pictureBoxIcon.TabStop = false;
        // 
        // AboutForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(456, 292);
        Controls.Add(panelAbout);
        FormStyle = ReaLTaiizor.Enum.Material.FormStyles.ActionBar_None;
        Icon = XlBlocks.AddIn.Properties.Resources.xlblocks_icon;
        Margin = new Padding(2);
        Name = "AboutForm";
        Padding = new Padding(3, 28, 2, 2);
        StartPosition = FormStartPosition.CenterScreen;
        Text = "About XlBlocks";
        FormClosing += AboutForm_FormClosing;
        Load += AboutForm_Load;
        KeyDown += AboutForm_KeyDown;
        panelAbout.ResumeLayout(false);
        ((System.ComponentModel.ISupportInitialize)pictureBoxIcon).EndInit();
        ResumeLayout(false);
    }

    #endregion

    private Panel panelAbout;
    private PictureBox pictureBoxIcon;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelInfo;
    private ReaLTaiizor.Controls.MaterialButton materialButtonClose;
    private ReaLTaiizor.Controls.MaterialLabel materialLabelVersion;
}
