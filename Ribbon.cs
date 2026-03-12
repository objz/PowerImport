using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

namespace PowerImport
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e) { }

        private void import_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var tablePane = new TablePane();
                Globals.ThisAddIn.ShowImportPane(tablePane);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed:\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void refresh_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var result = Globals.ThisAddIn.RefreshCurrentSheetTables();
                if (result.TablesRefreshed > 0)
                {
                    string msg = $"Refreshed {result.TablesRefreshed} table(s), {result.TotalRows.ToString("N0")} rows.";
                    if (result.Skipped > 0)
                        msg += $" {result.Skipped} skipped.";
                    Globals.ThisAddIn.Application.StatusBar = msg;
                }
                else
                {
                    Globals.ThisAddIn.Application.StatusBar = "No tables to refresh.";
                }
                ClearStatusBarAfterDelay();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Refresh failed:\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void refresh_all_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var result = Globals.ThisAddIn.RefreshAllImportedTables();
                if (result.TablesRefreshed > 0)
                {
                    string msg = $"Refreshed {result.TablesRefreshed} table(s), {result.TotalRows.ToString("N0")} rows.";
                    if (result.Skipped > 0)
                        msg += $" {result.Skipped} skipped.";
                    Globals.ThisAddIn.Application.StatusBar = msg;
                }
                else
                {
                    Globals.ThisAddIn.Application.StatusBar = "No Power BI tables to refresh.";
                }
                ClearStatusBarAfterDelay();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Refresh failed:\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void ClearStatusBarAfterDelay()
        {
            try
            {
                await Task.Delay(5000);
                Globals.ThisAddIn.Application.StatusBar = false;
            }
            catch
            {
            }
        }
    }
}
