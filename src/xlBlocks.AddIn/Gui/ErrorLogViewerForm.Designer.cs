namespace XlBlocks.AddIn.Gui;

partial class ErrorLogViewerForm
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
        var resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorLogViewerForm));
        panel1 = new Panel();
        materialListViewExceptions = new ReaLTaiizor.Controls.MaterialListView();
        columnHeaderTime = new ColumnHeader();
        columnHeaderFunction = new ColumnHeader();
        columnHeaderCell = new ColumnHeader();
        columnHeaderArgs = new ColumnHeader();
        columnHeaderError = new ColumnHeader();
        panel1.SuspendLayout();
        SuspendLayout();
        // 
        // panel1
        // 
        panel1.CausesValidation = false;
        panel1.Controls.Add(materialListViewExceptions);
        panel1.Dock = DockStyle.Fill;
        panel1.Location = new Point(3, 64);
        panel1.Name = "panel1";
        panel1.Size = new Size(911, 376);
        panel1.TabIndex = 0;
        // 
        // materialListViewExceptions
        // 
        materialListViewExceptions.AutoSizeTable = false;
        materialListViewExceptions.BackColor = Color.FromArgb(255, 255, 255);
        materialListViewExceptions.BorderStyle = BorderStyle.None;
        materialListViewExceptions.Columns.AddRange(new ColumnHeader[] { columnHeaderTime, columnHeaderFunction, columnHeaderCell, columnHeaderArgs, columnHeaderError });
        materialListViewExceptions.Depth = 0;
        materialListViewExceptions.Dock = DockStyle.Fill;
        materialListViewExceptions.FullRowSelect = true;
        materialListViewExceptions.Location = new Point(0, 0);
        materialListViewExceptions.MinimumSize = new Size(200, 100);
        materialListViewExceptions.MouseLocation = new Point(-1, -1);
        materialListViewExceptions.MouseState = ReaLTaiizor.Helper.MaterialDrawHelper.MaterialMouseState.OUT;
        materialListViewExceptions.Name = "materialListViewExceptions";
        materialListViewExceptions.OwnerDraw = true;
        materialListViewExceptions.Size = new Size(911, 376);
        materialListViewExceptions.TabIndex = 0;
        materialListViewExceptions.UseCompatibleStateImageBehavior = false;
        materialListViewExceptions.View = View.Details;
        // 
        // columnHeaderTime
        // 
        columnHeaderTime.Text = "Time";
        columnHeaderTime.Width = 150;
        // 
        // columnHeaderFunction
        // 
        columnHeaderFunction.Text = "Function";
        columnHeaderFunction.Width = 150;
        // 
        // columnHeaderCell
        // 
        columnHeaderCell.Text = "Cell";
        columnHeaderCell.Width = 175;
        // 
        // columnHeaderArgs
        // 
        columnHeaderArgs.Text = "Arguments";
        columnHeaderArgs.Width = 250;
        // 
        // columnHeaderError
        // 
        columnHeaderError.Text = "Error";
        columnHeaderError.Width = 400;
        // 
        // ErrorLogViewerForm
        // 
        AutoScaleDimensions = new SizeF(7F, 15F);
        AutoScaleMode = AutoScaleMode.Font;
        ClientSize = new Size(917, 443);
        Controls.Add(panel1);
        DoubleBuffered = false;
        Icon = XlBlocks.AddIn.Properties.Resources.xlblocks_icon;
        Name = "ErrorLogViewerForm";
        Text = "Error Log";
        FormClosing += ErrorLogViewerForm_FormClosing;
        KeyDown += ErrorLogViewerForm_KeyDown;
        panel1.ResumeLayout(false);
        ResumeLayout(false);
    }

    #endregion

    private Panel panel1;
    private ReaLTaiizor.Controls.MaterialListView materialListViewExceptions;
    private ColumnHeader columnHeaderTime;
    private ColumnHeader columnHeaderFunction;
    private ColumnHeader columnHeaderArgs;
    private ColumnHeader columnHeaderError;
    private ColumnHeader columnHeaderCell;
}
