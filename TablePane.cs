using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerImport
{
    public partial class TablePane : UserControl
    {
        private ComboBox instanceSelector;
        private FlowLayoutPanel tablePanel;
        private Label statusLabel;
        private List<(int port, string catalog)> runningInstances = new List<(int, string)>();
        private Button refreshButton;

        public TablePane()
        {
            InitializeComponent();
            Dock = DockStyle.Fill;
            AutoScroll = true;

            instanceSelector = new ComboBox
            {
                DropDownStyle = ComboBoxStyle.DropDownList,
                Width = 320,
                Margin = new Padding(10),
                Enabled = false
            };
            instanceSelector.SelectedIndexChanged += (s, e) => SwitchInstanceAsync();

            statusLabel = new Label
            {
                AutoSize = true,
                ForeColor = Color.DarkGray,
                Margin = new Padding(10, 0, 10, 0),
                Text = "Loading..."
            };

            tablePanel = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                Padding = new Padding(10),
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false
            };

            var layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                RowCount = 3,
                ColumnCount = 1,
                AutoSize = true
            };
            layout.Controls.Add(instanceSelector, 0, 0);
            layout.Controls.Add(statusLabel, 0, 1);
            layout.Controls.Add(tablePanel, 0, 2);
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            Controls.Add(layout);

            ReloadInstancesAsync();
        }

        private async void ReloadInstancesAsync()
        {
            statusLabel.Text = "Loading Power BI Desktop instances...";
            tablePanel.Controls.Clear();
            instanceSelector.Enabled = false;
            instanceSelector.Items.Clear();
            runningInstances.Clear();

            var portFiles = await Task.Run(() => FindPowerBIPortFiles());

            if (IsHandleCreated)
                Invoke((Action)(() =>
                {
                    if (portFiles.Count == 0)
                    {
                        statusLabel.Text = "No running Power BI Desktop instances found.";
                        tablePanel.Controls.Clear();
                        instanceSelector.Enabled = false;

                        if (refreshButton == null)
                        {
                            refreshButton = new Button
                            {
                                Text = "Refresh",
                                Width = 120,
                                Height = 36,
                                Margin = new Padding(10)
                            };
                            refreshButton.Click += (s, e) =>
                            {
                                statusLabel.Text = "Retrying...";
                                ReloadInstancesAsync();
                            };
                        }
                        if (!Controls.Contains(refreshButton))
                            Controls.Add(refreshButton);
                    }
                    else
                    {
                        if (refreshButton != null && Controls.Contains(refreshButton))
                            Controls.Remove(refreshButton);

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
                }));
        }

        private string GetPrettyCatalogName(string catalog)
        {
            return catalog;
        }

        private async void SwitchInstanceAsync()
        {
            int idx = instanceSelector.SelectedIndex;
            if (idx < 0 || idx >= runningInstances.Count)
            {
                statusLabel.Text = "No instance selected.";
                tablePanel.Controls.Clear();
                return;
            }

            var selected = runningInstances[idx];
            string newConn = $"Provider=MSOLAP;Data Source=localhost:{selected.port};Initial Catalog={selected.catalog};Integrated Security=SSPI;Impersonation Level=Impersonate;Persist Security Info=True;";

            if (Globals.ThisAddIn.HasActiveConnection())
                Globals.ThisAddIn.Connection?.Close();

            instanceSelector.Enabled = false;
            statusLabel.Text = "Connecting...";
            tablePanel.Controls.Clear();

            try
            {
                await Task.Run(() =>
                {
                    Globals.ThisAddIn.ConnectionString = newConn;
                    Globals.ThisAddIn.Connection = new Microsoft.AnalysisServices.AdomdClient.AdomdConnection(newConn);
                    Globals.ThisAddIn.Connection.Open();
                });

                if (IsHandleCreated)
                    Invoke((Action)(() =>
                    {
                        statusLabel.Text = $"Connected to {instanceSelector.SelectedItem}";
                        ReloadTablesAsync();
                        instanceSelector.Enabled = true;
                    }));
            }
            catch (Exception ex)
            {
                if (IsHandleCreated)
                    Invoke((Action)(() =>
                    {
                        statusLabel.Text = "Failed to connect: " + ex.Message;
                        tablePanel.Controls.Clear();
                        instanceSelector.Enabled = true;
                    }));
            }
        }

        private async void ReloadTablesAsync()
        {
            tablePanel.Controls.Clear();
            statusLabel.Text = "Loading tables...";
            var tables = new List<string>();

            try
            {
                tables = await Task.Run(() => Globals.ThisAddIn.GetAvailableTableNames());
            }
            catch (Exception ex)
            {
                statusLabel.Text = "Failed to get tables: " + ex.Message;
                return;
            }

            if (tables.Count == 0)
            {
                statusLabel.Text = "No tables available.";
                return;
            }

            foreach (var tableName in tables)
            {
                var btn = new Button
                {
                    Text = tableName,
                    Width = 300,
                    Height = 40,
                    TextAlign = ContentAlignment.MiddleLeft,
                    FlatStyle = FlatStyle.Standard,
                    Tag = tableName
                };

                btn.Click += (s, e) =>
                {
                    var dlg = new TableDialog(tableName);
                    if (dlg.ShowDialog() == DialogResult.OK)
                    {
                        Globals.ThisAddIn.ImportTable(tableName, dlg.UseNewSheet, dlg.TargetCellAddress);
                    }
                };

                tablePanel.Controls.Add(btn);
            }
            statusLabel.Text = "";
        }

        private List<(int port, string catalog)> FindPowerBIPortFiles()
        {
            var result = new List<(int, string)>();
            string[] baseDirs = new[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "Packages", "Microsoft.MicrosoftPowerBIDesktop_8wekyb3d8bbwe", "LocalCache",
                    "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces")
            };

            foreach (var baseDir in baseDirs)
            {
                if (!Directory.Exists(baseDir)) continue;
                foreach (var dir in Directory.GetDirectories(baseDir))
                {
                    var portPath = Path.Combine(dir, "Data", "msmdsrv.port.txt");
                    if (!File.Exists(portPath)) continue;

                    string raw = File.ReadAllText(portPath, System.Text.Encoding.Unicode);
                    string portText = new string(raw.Where(char.IsDigit).ToArray());

                    if (!int.TryParse(portText, out int port)) continue;

                    try
                    {
                        string connStr = $"Provider=MSOLAP;Data Source=localhost:{port};Integrated Security=SSPI;";
                        using (var conn = new Microsoft.AnalysisServices.AdomdClient.AdomdConnection(connStr))
                        {
                            conn.Open();
                            var catalogs = new List<string>();
                            var cmd = conn.CreateCommand();
                            cmd.CommandText = "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS";
                            using (var reader = cmd.ExecuteReader())
                            {
                                while (reader.Read())
                                {
                                    var name = reader["CATALOG_NAME"].ToString();
                                    if (!string.IsNullOrWhiteSpace(name))
                                        catalogs.Add(name);
                                }
                            }
                            if (catalogs.Count > 0)
                                result.Add((port, catalogs[0]));
                        }
                    }
                    catch { }
                }
            }
            return result;
        }
    }
}
