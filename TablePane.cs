using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.AnalysisServices.AdomdClient;

namespace PowerImport
{
    public partial class TablePane : UserControl
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, string lParam);
        private const int EM_SETCUEBANNER = 0x1501;

        private ComboBox instanceSelector;
        private FlowLayoutPanel tablePanel;
        private Label statusLabel;
        private TextBox searchBox;
        private Button refreshButton;
        private List<(int port, string catalog)> runningInstances = new List<(int, string)>();
        private List<string> allTableNames = new List<string>();

        private CancellationTokenSource _cts = new CancellationTokenSource();

        private bool _loading;

        private static readonly float _dpiScale;
        static TablePane()
        {
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
                _dpiScale = g.DpiX / 96f;
        }
        private static int Dpi(int value) => (int)(value * _dpiScale);

        public TablePane()
        {
            InitializeComponent();
            Dock = DockStyle.Fill;
            AutoScroll = true;

            instanceSelector = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = Dpi(320),
                Margin = new Padding(Dpi(10)),
                Enabled = false
            };
            instanceSelector.SelectedIndexChanged += (s, e) => SwitchInstanceAsync();

            statusLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.DarkGray,
                Margin = new Padding(Dpi(10), 0, Dpi(10), 0),
                Text = "Loading..."
            };

            searchBox = new TextBox
            {
                Width = Dpi(320),
                Margin = new Padding(Dpi(10), 0, Dpi(10), Dpi(4)),
                Visible = false
            };
            searchBox.HandleCreated += (s, e) =>
                SendMessage(searchBox.Handle, EM_SETCUEBANNER, (IntPtr)1, "Search tables...");
            searchBox.TextChanged += (s, e) => FilterTables(searchBox.Text);

            tablePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(Dpi(10)),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 4,
                ColumnCount = 1,
                AutoSize = true
            };
            layout.Controls.Add(instanceSelector, 0, 0);
            layout.Controls.Add(statusLabel, 0, 1);
            layout.Controls.Add(searchBox, 0, 2);
            layout.Controls.Add(tablePanel, 0, 3);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            Controls.Add(layout);
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            ReloadInstancesAsync();
        }

        private void ClearTablePanel()
        {
            foreach (Control c in tablePanel.Controls)
                c.Dispose();
            tablePanel.Controls.Clear();
        }

        private string GetPrettyCatalogName(string catalog)
        {
            return catalog;
        }

        private void BuildTableButtons(IEnumerable<string> tableNames)
        {
            ClearTablePanel();
            foreach (var tableName in tableNames)
            {
                var btn = new Button
                {
                    Text = tableName,
                    Width = Dpi(300),
                    Height = Dpi(40),
                    TextAlign = ContentAlignment.MiddleLeft,
                    FlatStyle = FlatStyle.Standard,
                    Tag = tableName
                };
                btn.Click += TableButton_Click;
                tablePanel.Controls.Add(btn);
            }
        }

        private void FilterTables(string filter)
        {
            if (allTableNames.Count == 0) return;

            var filtered = string.IsNullOrWhiteSpace(filter)
                ? allTableNames
                : allTableNames.Where(t => t.IndexOf(filter, StringComparison.OrdinalIgnoreCase) >= 0).ToList();

            BuildTableButtons(filtered);
        }

        private async void ReloadInstancesAsync()
        {
            if (_loading) return;
            _loading = true;

            try
            {
                statusLabel.Text = "Loading Power BI Desktop instances...";
                ClearTablePanel();
                allTableNames.Clear();
                searchBox.Text = "";
                searchBox.Visible = false;
                instanceSelector.Enabled = false;
                instanceSelector.Items.Clear();
                runningInstances.Clear();

                var token = _cts.Token;
                var portFiles = await Task.Run(() => ThisAddIn.FindPowerBIPortFiles(), token);
                token.ThrowIfCancellationRequested();

                if (portFiles.Count == 0)
                {
                    statusLabel.Text = "No running Power BI Desktop instances found.";
                    ClearTablePanel();
                    instanceSelector.Enabled = false;

                    if (refreshButton == null)
                    {
                        refreshButton = new Button
                        {
                            Text = "Refresh",
                            Width = Dpi(120),
                            Height = Dpi(36),
                            Margin = new Padding(Dpi(10))
                        };
                        refreshButton.Click += RefreshButton_Click;
                    }
                    if (!tablePanel.Controls.Contains(refreshButton))
                        tablePanel.Controls.Add(refreshButton);
                }
                else
                {
                    if (refreshButton != null && tablePanel.Controls.Contains(refreshButton))
                        tablePanel.Controls.Remove(refreshButton);

                    foreach (var (port, catalog) in portFiles)
                    {
                        string label = GetPrettyCatalogName(catalog);
                        instanceSelector.Items.Add(label);
                        runningInstances.Add((port, catalog));
                    }

                    instanceSelector.SelectedIndex = 0;
                    instanceSelector.Enabled = true;
                    statusLabel.Text = "";
                }
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                if (!_cts.IsCancellationRequested)
                    statusLabel.Text = "Error loading instances: " + ex.Message;
            }
            finally
            {
                _loading = false;
            }
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            statusLabel.Text = "Retrying...";
            ReloadInstancesAsync();
        }

        private async void SwitchInstanceAsync()
        {
            if (_loading) return;

            int idx = instanceSelector.SelectedIndex;
            if (idx < 0 || idx >= runningInstances.Count)
            {
                statusLabel.Text = "No instance selected.";
                ClearTablePanel();
                allTableNames.Clear();
                searchBox.Text = "";
                searchBox.Visible = false;
                return;
            }

            _loading = true;

            var selected = runningInstances[idx];
            string newConn =
                $"Provider=MSOLAP;Data Source=localhost:{selected.port};" +
                $"Initial Catalog={selected.catalog};Integrated Security=SSPI;" +
                "Impersonation Level=Impersonate;";

            instanceSelector.Enabled = false;
            statusLabel.Text = "Connecting...";
            ClearTablePanel();
            allTableNames.Clear();
            searchBox.Text = "";
            searchBox.Visible = false;

            try
            {
                var token = _cts.Token;

                var conn = await Task.Run(() =>
                {
                    var c = new AdomdConnection(newConn);
                    c.Open();
                    return c;
                }, token);

                token.ThrowIfCancellationRequested();

                Globals.ThisAddIn.ConnectionString = newConn;
                Globals.ThisAddIn.Connection = conn;
                Globals.ThisAddIn.ActiveCatalog = selected.catalog;

                statusLabel.Text = $"Connected to {instanceSelector.SelectedItem}";
                instanceSelector.Enabled = true;
                ReloadTablesAsync();
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                if (!_cts.IsCancellationRequested)
                {
                    statusLabel.Text = "Failed to connect: " + ex.Message;
                    ClearTablePanel();
                    instanceSelector.Enabled = true;
                }
            }
            finally
            {
                _loading = false;
            }
        }

        private async void ReloadTablesAsync()
        {
            ClearTablePanel();
            allTableNames.Clear();
            searchBox.Text = "";
            searchBox.Visible = false;
            statusLabel.Text = "Loading tables...";
            List<string> tables;

            try
            {
                var token = _cts.Token;
                tables = await Task.Run(() => Globals.ThisAddIn.GetAvailableTableNames(), token);
                token.ThrowIfCancellationRequested();
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                if (!_cts.IsCancellationRequested)
                    statusLabel.Text = "Failed to get tables: " + ex.Message;
                return;
            }

            if (tables.Count == 0)
            {
                statusLabel.Text = "No tables available.";
                return;
            }

            allTableNames = tables;
            statusLabel.Text = $"{tables.Count} tables";
            searchBox.Visible = true;
            BuildTableButtons(tables);
        }

        private void TableButton_Click(object sender, EventArgs e)
        {
            var btn = (Button)sender;
            var tableName = (string)btn.Tag;

            using (var dlg = new TableDialog(tableName))
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    int rows = Globals.ThisAddIn.ImportTable(tableName, dlg.UseNewSheet, dlg.TargetCellAddress);
                    if (rows >= 0)
                    {
                        Globals.ThisAddIn.Application.StatusBar =
                            $"Imported '{tableName}' ({rows.ToString("N0")} rows)";
                    }
                }
            }
        }
    }
}
