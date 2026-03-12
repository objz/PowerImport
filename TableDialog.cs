using System;
using System.Drawing;
using System.Runtime.InteropServices;
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

            float scale;
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
                scale = g.DpiX / 96f;

            Text = $"Import {tableName}";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding((int)(5 * scale));

            int titleTextWidth = TextRenderer.MeasureText(Text, SystemFonts.CaptionFont).Width;
            int titleBarExtra = (int)(80 * scale);
            int minWidth = Math.Max(titleTextWidth + titleBarExtra, (int)(280 * scale));
            MinimumSize = new Size(minWidth, 0);

            var layout = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                RowCount = 3,
                ColumnCount = 3,
                Padding = new Padding((int)(10 * scale)),
            };

            radioNewSheet = new RadioButton
            {
                Text = "New worksheet",
                Checked = true,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, (int)(4 * scale))
            };
            layout.Controls.Add(radioNewSheet, 0, 0);
            layout.SetColumnSpan(radioNewSheet, 3);

            radioExistingSheet = new RadioButton
            {
                Text = "Existing at:",
                AutoSize = true,
                Margin = new Padding(0)
            };
            textCell = new TextBox
            {
                Text = "A1",
                Enabled = false,
                Width = (int)(100 * scale),
                Margin = new Padding((int)(4 * scale), 0, (int)(4 * scale), 0)
            };
            btnPickCell = new Button
            {
                Text = "...",
                Enabled = false,
                Width = (int)(28 * scale),
                Height = (int)(24 * scale),
                Margin = new Padding(0)
            };

            layout.Controls.Add(radioExistingSheet, 0, 1);
            layout.Controls.Add(textCell, 1, 1);
            layout.Controls.Add(btnPickCell, 2, 1);

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Width = (int)(75 * scale)
            };
            btnCancel = new Button
            {
                Text = "Cancel",
                DialogResult = DialogResult.Cancel,
                Width = (int)(75 * scale)
            };

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                Margin = new Padding(0, (int)(10 * scale), 0, 0)
            };
            buttonPanel.Controls.Add(btnOK);
            buttonPanel.Controls.Add(btnCancel);
            layout.Controls.Add(buttonPanel, 0, 2);
            layout.SetColumnSpan(buttonPanel, 3);

            AcceptButton = btnOK;
            CancelButton = btnCancel;
            Controls.Add(layout);

            void UpdateCellState(object s, EventArgs ev)
            {
                bool enabled = radioExistingSheet.Checked;
                textCell.Enabled = enabled;
                btnPickCell.Enabled = enabled;
            }
            radioNewSheet.CheckedChanged += UpdateCellState;
            radioExistingSheet.CheckedChanged += UpdateCellState;

            btnPickCell.Click += PickCell_Click;
            UpdateCellState(null, null);
        }

        private void PickCell_Click(object sender, EventArgs e)
        {
            Range range = null;
            try
            {
                var xlApp = Globals.ThisAddIn.Application;
                var result = xlApp.InputBox("Select target cell", "Target Cell",
                    textCell.Text, Type: 8);

                if (result is bool) return;

                range = result as Range;
                if (range != null)
                    textCell.Text = range.get_Address(false, false);
            }
            catch (COMException)
            {
            }
            catch (Exception ex)
            {
                MessageBox.Show("Could not select cell:\n" + ex.Message,
                    "Cell Picker", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            finally
            {
                if (range != null)
                {
                    try { Marshal.ReleaseComObject(range); }
                    catch { }
                }
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);

            if (DialogResult == DialogResult.OK && radioExistingSheet.Checked)
            {
                string addr = textCell.Text?.Trim();
                if (string.IsNullOrEmpty(addr))
                {
                    MessageBox.Show("Please enter a valid cell address.",
                        "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    e.Cancel = true;
                    return;
                }

                try
                {
                    var ws = Globals.ThisAddIn.Application.ActiveSheet as Worksheet;
                    if (ws != null)
                    {
                        Range test = null;
                        try
                        {
                            test = ws.Range[addr];
                        }
                        catch
                        {
                            MessageBox.Show($"'{addr}' is not a valid cell address.",
                                "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            e.Cancel = true;
                        }
                        finally
                        {
                            if (test != null)
                            {
                                try { Marshal.ReleaseComObject(test); }
                                catch { }
                            }
                        }
                    }
                }
                catch { }
            }
        }
    }
}
