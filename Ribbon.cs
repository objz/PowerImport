using System;
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
                MessageBox.Show("Import failed:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void refresh_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var refreshed = Globals.ThisAddIn.RefreshCurrentSheetTables();
                Globals.ThisAddIn.Application.StatusBar = refreshed
                    ? "Refresh successful."
                    : "No tables to refresh.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Refresh failed:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void refresh_all_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var refreshed = Globals.ThisAddIn.RefreshAllImportedTables();
                Globals.ThisAddIn.Application.StatusBar = refreshed
                    ? "Refresh successful."
                    : "No Power BI table to refresh in this sheet.";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Refresh failed:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
