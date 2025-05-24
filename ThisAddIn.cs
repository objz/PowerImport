using System;
using System.Collections.Generic;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.AnalysisServices.AdomdClient;
using System.Linq;

namespace PowerImport
{
    public partial class ThisAddIn
    {
        public string ConnectionString;
        public AdomdConnection Connection;
        public CustomTaskPane ImportPane;

        private void ThisAddIn_Startup(object sender, EventArgs e) { }
        private void ThisAddIn_Shutdown(object sender, EventArgs e) => Connection?.Dispose();

        public bool HasActiveConnection() => Connection?.State == System.Data.ConnectionState.Open;

        public void ShowImportPane(UserControl pane)
        {
            ImportPane?.Dispose();
            ImportPane = CustomTaskPanes.Add(pane, "Import Power BI Tables");
            ImportPane.Width = 350;
            ImportPane.Visible = true;
        }

        public bool RefreshCurrentSheetTables()
        {
            var workbook = Application.ActiveWorkbook;
            var worksheet = Application.ActiveSheet as Worksheet;
            if (worksheet == null) return false;
            Worksheet metaSheet = null;
            try { metaSheet = workbook.Worksheets["PBIDesktop_Metadata"] as Worksheet; } catch { return false; }
            if (metaSheet == null) return false;
            int lastRow = metaSheet.Cells[metaSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
            bool refreshed = false;
            for (int row = 2; row <= lastRow; row++)
            {
                string wsName = metaSheet.Cells[row, 1].Value2 as string;
                string tblName = metaSheet.Cells[row, 2].Value2 as string;
                string pbiName = metaSheet.Cells[row, 3].Value2 as string;
                if (wsName == worksheet.Name)
                {
                    var excelTable = worksheet.ListObjects.Cast<ListObject>().FirstOrDefault(t => t.Name == tblName);
                    if (excelTable != null)
                    {
                        try
                        {
                            UpdateTableFromPowerBI(excelTable, pbiName);
                            refreshed = true;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error refreshing table '{tblName}' on sheet '{wsName}':\n{ex.Message}", "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
            if (!refreshed)
                MessageBox.Show("No Power BI tables to refresh in this sheet.", "Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return refreshed;
        }

        public bool RefreshAllImportedTables()
        {
            var workbook = Application.ActiveWorkbook;
            Worksheet metaSheet = null;
            try { metaSheet = workbook.Worksheets["PBIDesktop_Metadata"] as Worksheet; } catch { return false; }
            if (metaSheet == null) return false;
            int lastRow = metaSheet.Cells[metaSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
            bool refreshed = false;
            for (int row = 2; row <= lastRow; row++)
            {
                string wsName = metaSheet.Cells[row, 1].Value2 as string;
                string tblName = metaSheet.Cells[row, 2].Value2 as string;
                string pbiName = metaSheet.Cells[row, 3].Value2 as string;
                var worksheet = workbook.Worksheets[wsName] as Worksheet;
                if (worksheet == null) continue;
                var excelTable = worksheet.ListObjects.Cast<ListObject>().FirstOrDefault(t => t.Name == tblName);
                if (excelTable != null)
                {
                    try
                    {
                        UpdateTableFromPowerBI(excelTable, pbiName);
                        refreshed = true;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Error refreshing table '{tblName}' on sheet '{wsName}':\n{ex.Message}", "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            return refreshed;
        }

        private bool UpdateTableFromPowerBI(ListObject excelTable, string pbiTable)
        {
            var worksheet = excelTable.Parent as Worksheet;
            var anchorCell = excelTable.Range.Cells[1, 1];
            int dataRows = excelTable.DataBodyRange?.Rows.Count ?? 0;
            if (dataRows > 0)
                excelTable.DataBodyRange.Delete();

            var cmd = Connection.CreateCommand();
            cmd.CommandText = $"EVALUATE {EscapeDaxIdentifier(pbiTable)}";
            var rows = new List<object[]>();
            int colCount = 0;
            string[] headers = null;
            Type[] columnTypes = null;

            using (var reader = cmd.ExecuteReader())
            {
                colCount = reader.FieldCount;
                headers = new string[colCount];
                columnTypes = new Type[colCount];
                for (int i = 0; i < colCount; i++)
                {
                    headers[i] = RemovePrefix(reader.GetName(i), pbiTable);
                    columnTypes[i] = reader.GetFieldType(i);
                }
                System.Diagnostics.Debug.WriteLine("Column types:");
                for (int i = 0; i < colCount; i++)
                    System.Diagnostics.Debug.WriteLine($"[{i}] {headers[i]}: {columnTypes[i].Name}");

                while (reader.Read())
                {
                    var values = new object[colCount];
                    reader.GetValues(values);

                    rows.Add(values);
                }
            }
            if (rows.Count > 0)
            {
                var dataArr = To2DArray(rows);
                var insertStart = anchorCell.Offset[1, 0];
                var insertRange = insertStart.Resize[rows.Count, colCount];
                insertRange.Value2 = dataArr;

                for (int i = 0; i < colCount; i++)
                {
                    if (columnTypes[i] == typeof(DateTime))
                    {
                        Range dateCol = insertStart.Offset[0, i].Resize[rows.Count, 1];
                        dateCol.NumberFormat = "m/d/yyyy";
                    }
                }
            }
            excelTable.Range.Columns.AutoFit();
            return true;
        }


        public List<string> GetAvailableTableNames()
        {
            var cmd = Connection.CreateCommand();
            cmd.CommandText = "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES";
            var result = new List<string>();
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    string tableName = reader["Name"]?.ToString();
                    if (!string.IsNullOrWhiteSpace(tableName))
                        result.Add(tableName);
                }
            }
            return result;
        }

        public void ImportTable(string pbiTable, bool createSheet, string targetCell)
        {
            var workbook = Application.ActiveWorkbook;
            Worksheet worksheet = null;
            Range anchor = null;
            try
            {
                var cmd = Connection.CreateCommand();
                cmd.CommandText = $"EVALUATE {EscapeDaxIdentifier(pbiTable)}";
                using (var reader = cmd.ExecuteReader())
                {
                    int colCount = reader.FieldCount;
                    var headers = new object[colCount];
                    var columnTypes = new Type[colCount];
                    for (int i = 0; i < colCount; i++)
                    {
                        headers[i] = RemovePrefix(reader.GetName(i), pbiTable);
                        columnTypes[i] = reader.GetFieldType(i);
                    }
                    System.Diagnostics.Debug.WriteLine("Column types:");
                    for (int i = 0; i < colCount; i++)
                        System.Diagnostics.Debug.WriteLine($"[{i}] {headers[i]}: {columnTypes[i].Name}");

                    var rows = new List<object[]>();
                    while (reader.Read())
                    {
                        var values = new object[colCount];
                        reader.GetValues(values);
                        rows.Add(values);
                    }
                    worksheet = createSheet ? workbook.Worksheets.Add() : workbook.ActiveSheet;
                    worksheet.Name = createSheet ? (pbiTable.Length > 28 ? pbiTable.Substring(0, 28) : pbiTable) : worksheet.Name;
                    anchor = worksheet.Range[targetCell];
                    anchor.Resize[1, colCount].Value2 = headers;
                    if (rows.Count > 0)
                    {
                        anchor.Offset[1, 0].Resize[rows.Count, colCount].Value2 = To2DArray(rows);

                        for (int i = 0; i < colCount; i++)
                        {
                            if (columnTypes[i] == typeof(DateTime))
                            {
                                Range dateCol = anchor.Offset[1, i].Resize[rows.Count, 1];
                                dateCol.NumberFormat = "m/d/yyyy";
                            }
                        }
                    }

                    var dataRange = worksheet.Range[anchor, anchor.Offset[rows.Count, colCount - 1]];
                    var table = worksheet.ListObjects.Add(XlListObjectSourceType.xlSrcRange, dataRange, Type.Missing, XlYesNoGuess.xlYes);
                    table.Name = $"tbl_{pbiTable}";
                    table.TableStyle = "TableStyleMedium9";
                    dataRange.Columns.AutoFit();
                    WriteOrUpdateMetadata(worksheet.Name, table.Name, pbiTable, anchor.Address[false, false]);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed:\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (createSheet && worksheet != null)
                    worksheet.Delete();
            }
        }


        private void WriteOrUpdateMetadata(string wsName, string excelTable, string pbiTable, string anchorCell)
        {
            var workbook = Application.ActiveWorkbook;
            Worksheet metaSheet = null;
            try { metaSheet = workbook.Worksheets["PBIDesktop_Metadata"] as Worksheet; } catch { }
            if (metaSheet == null)
            {
                metaSheet = workbook.Worksheets.Add();
                metaSheet.Name = "PBIDesktop_Metadata";
                metaSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                metaSheet.Cells[1, 1].Value2 = "Worksheet";
                metaSheet.Cells[1, 2].Value2 = "ExcelTableName";
                metaSheet.Cells[1, 3].Value2 = "PowerBITableName";
                metaSheet.Cells[1, 4].Value2 = "AnchorCell";
            }
            int lastRow = metaSheet.Cells[metaSheet.Rows.Count, 1].End[XlDirection.xlUp].Row;
            bool found = false;
            for (int row = 2; row <= lastRow; row++)
            {
                string sheet = metaSheet.Cells[row, 1].Value2 as string;
                string table = metaSheet.Cells[row, 2].Value2 as string;
                if (sheet == wsName && table == excelTable)
                {
                    metaSheet.Cells[row, 3].Value2 = pbiTable;
                    metaSheet.Cells[row, 4].Value2 = anchorCell;
                    found = true;
                    break;
                }
            }
            if (!found)
            {
                int newRow = lastRow + 1;
                metaSheet.Cells[newRow, 1].Value2 = wsName;
                metaSheet.Cells[newRow, 2].Value2 = excelTable;
                metaSheet.Cells[newRow, 3].Value2 = pbiTable;
                metaSheet.Cells[newRow, 4].Value2 = anchorCell;
            }
        }

        private string RemovePrefix(string column, string table)
        {
            string prefix = table + "[";
            if (column.StartsWith(prefix) && column.EndsWith("]"))
                return column.Substring(prefix.Length, column.Length - prefix.Length - 1);
            return column;
        }

        private object[,] To2DArray(List<object[]> rows)
        {
            int rowCount = rows.Count;
            int colCount = rows[0].Length;
            var array = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
                for (int c = 0; c < colCount; c++)
                    array[r, c] = rows[r][c];
            return array;
        }

        private string EscapeDaxIdentifier(string name) => $"'{name.Replace("'", "''")}'";

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
