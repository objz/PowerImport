using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Button = System.Windows.Forms.Button;
using TextBox = System.Windows.Forms.TextBox;

namespace PowerImport
{
    public partial class TableDialog : Form
    {
        public bool UseNewSheet => radioNewSheet.Checked;
        public string TargetCellAddress => textCell.Text;

        private RadioButton radioNewSheet;
        private RadioButton radioExistingSheet;
        private TextBox textCell;
        private Button btnPickCell;
        private Button btnOK;
        private Button btnCancel;

        public TableDialog(string tableName)
        {
            InitializeComponent();
            Text = $"Import {tableName}";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            TopMost = true;
            Width = 265;
            Height = 140;

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 3,
                Padding = new Padding(10, 10, 10, 10),
                AutoSize = true
            };

            radioNewSheet = new RadioButton
            {
                Text = "New worksheet",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 4)
            };
            layout.Controls.Add(radioNewSheet, 0, 0);
            layout.SetColumnSpan(radioNewSheet, 3);

            radioExistingSheet = new RadioButton
            {
                Text = "Existing at:",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 0)
            };
            textCell = new TextBox
            {
                Text = "$A$1",
                Enabled = false,
                Width = 95,
                Margin = new Padding(2, 0, 2, 0)
            };
            btnPickCell = new Button
            {
                Text = "...",
                Enabled = false,
                Width = 24,
                Height = 21,
                Margin = new Padding(0)
            };

            layout.Controls.Add(radioExistingSheet, 0, 1);
            layout.Controls.Add(textCell, 1, 1);
            layout.Controls.Add(btnPickCell, 2, 1);

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Width = 75
            };
            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Width = 75
            };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, 10, 0, 0)
            };
            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);
            layout.Controls.Add(buttonPanel, 0, 2);
            layout.SetColumnSpan(buttonPanel, 3);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
            Controls.Add(layout);

            void UpdateCellState(object s, EventArgs e)
            {
                bool enabled = radioExistingSheet.Checked;
                textCell.Enabled = enabled;
                btnPickCell.Enabled = enabled;
            }
            radioNewSheet.CheckedChanged += UpdateCellState;
            radioExistingSheet.CheckedChanged += UpdateCellState;
            btnPickCell.Click += (s, e) =>
            {
                try
                {
                    var xlApp = Globals.ThisAddIn.Application;
                    var range = xlApp.InputBox("Select target cell", "Target Cell", textCell.Text, Type: 8) as Range;
                    if (range != null)
                        textCell.Text = range.get_Address(false, false);
                }
                catch { }
            };
            UpdateCellState(null, null);
        }
    }
}
