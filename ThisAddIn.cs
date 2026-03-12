using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.AnalysisServices.AdomdClient;

namespace PowerImport
{
    public partial class ThisAddIn
    {
        public string ConnectionString;
        public string ActiveCatalog;
        public CustomTaskPane ImportPane;

        private AdomdConnection _connection;
        public AdomdConnection Connection
        {
            get => _connection;
            set
            {
                if (ReferenceEquals(_connection, value))
                    return;
                var old = _connection;
                _connection = value;
                if (old != null)
                {
                    try { old.Dispose(); }
                    catch { }
                }
            }
        }

        public struct RefreshResult
        {
            public int TablesRefreshed;
            public int TotalRows;
            public int Skipped;
        }

        private struct QueryResult
        {
            public string[] Headers;
            public Type[] ColumnTypes;
            public List<object[]> Rows;
            public int ColCount;
        }

        private void ThisAddIn_Startup(object sender, EventArgs e) { }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            Connection = null;
        }

        public bool HasActiveConnection() =>
            _connection != null && _connection.State == System.Data.ConnectionState.Open;

        private static void ReleaseCom(object obj)
        {
            if (obj == null) return;
            try { Marshal.ReleaseComObject(obj); }
            catch { }
        }

        private static string GetCellString(Worksheet sheet, int row, int col)
        {
            Range cell = null;
            try
            {
                cell = (Range)sheet.Cells[row, col];
                return cell.Value2 as string;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private static void SetCellValue(Worksheet sheet, int row, int col, object value)
        {
            Range cell = null;
            try
            {
                cell = (Range)sheet.Cells[row, col];
                cell.Value2 = value;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        public void ShowImportPane(UserControl pane)
        {
            ImportPane?.Dispose();
            ImportPane = CustomTaskPanes.Add(pane, "Import Power BI Tables");
            float scale;
            using (var g = Graphics.FromHwnd(IntPtr.Zero))
                scale = g.DpiX / 96f;
            ImportPane.Width = (int)(350 * scale);
            ImportPane.Visible = true;
        }

        private QueryResult ExecuteTableQuery(string pbiTable, AdomdConnection connection)
        {
            var result = new QueryResult
            {
                Rows = new List<object[]>()
            };

            using (var cmd = connection.CreateCommand())
            {
                cmd.CommandText = $"EVALUATE {EscapeDaxIdentifier(pbiTable)}";

                using (var reader = cmd.ExecuteReader())
                {
                    result.ColCount = reader.FieldCount;
                    if (result.ColCount == 0)
                        return result;

                    result.Headers = new string[result.ColCount];
                    result.ColumnTypes = new Type[result.ColCount];
                    for (int i = 0; i < result.ColCount; i++)
                    {
                        result.Headers[i] = RemovePrefix(reader.GetName(i), pbiTable);
                        result.ColumnTypes[i] = reader.GetFieldType(i);
                    }

                    while (reader.Read())
                    {
                        var values = new object[result.ColCount];
                        reader.GetValues(values);
                        result.Rows.Add(values);
                    }
                }
            }

            return result;
        }

        public RefreshResult RefreshCurrentSheetTables()
        {
            var result = new RefreshResult();
            var workbook = Application.ActiveWorkbook;
            var worksheet = Application.ActiveSheet as Worksheet;
            if (worksheet == null) return result;

            Worksheet metaSheet = GetMetadataSheet(workbook);
            if (metaSheet == null) return result;

            int lastRow = GetMetadataLastRow(metaSheet);
            if (lastRow < 2) return result;

            var runningInstances = FindPowerBIPortFiles();
            var catalogConnections = new Dictionary<string, AdomdConnection>(StringComparer.OrdinalIgnoreCase);

            try
            {
                for (int row = 2; row <= lastRow; row++)
                {
                    string wsName = GetCellString(metaSheet, row, 1);
                    string tblName = GetCellString(metaSheet, row, 2);
                    string pbiName = GetCellString(metaSheet, row, 3);
                    string catalog = GetCellString(metaSheet, row, 5);

                    if (wsName != worksheet.Name) continue;

                    ListObject excelTable = FindListObject(worksheet, tblName);
                    if (excelTable == null) continue;

                    var conn = GetConnectionForCatalog(catalog, runningInstances, catalogConnections);
                    if (conn == null)
                    {
                        result.Skipped++;
                        continue;
                    }

                    try
                    {
                        Application.StatusBar = $"Refreshing '{tblName}'...";
                        int rows = UpdateTableFromPowerBI(excelTable, pbiName, conn);
                        if (rows >= 0)
                        {
                            result.TablesRefreshed++;
                            result.TotalRows += rows;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Error refreshing table '{tblName}' on sheet '{wsName}':\n{ex.Message}",
                            "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            finally
            {
                DisposeTempConnections(catalogConnections);
            }

            if (result.TablesRefreshed == 0 && result.Skipped == 0)
            {
                MessageBox.Show("No Power BI tables to refresh on this sheet.",
                    "Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (result.TablesRefreshed == 0 && result.Skipped > 0)
            {
                MessageBox.Show(
                    $"{result.Skipped} table(s) could not be refreshed because their Power BI Desktop instance is not running.",
                    "Refresh", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            return result;
        }

        public RefreshResult RefreshAllImportedTables()
        {
            var result = new RefreshResult();
            var workbook = Application.ActiveWorkbook;
            Worksheet metaSheet = GetMetadataSheet(workbook);
            if (metaSheet == null) return result;

            int lastRow = GetMetadataLastRow(metaSheet);
            if (lastRow < 2) return result;

            var runningInstances = FindPowerBIPortFiles();
            var catalogConnections = new Dictionary<string, AdomdConnection>(StringComparer.OrdinalIgnoreCase);

            try
            {
                for (int row = 2; row <= lastRow; row++)
                {
                    string wsName = GetCellString(metaSheet, row, 1);
                    string tblName = GetCellString(metaSheet, row, 2);
                    string pbiName = GetCellString(metaSheet, row, 3);
                    string catalog = GetCellString(metaSheet, row, 5);

                    Worksheet ws = null;
                    try { ws = workbook.Worksheets[wsName] as Worksheet; }
                    catch { continue; }
                    if (ws == null) continue;

                    ListObject excelTable = FindListObject(ws, tblName);
                    if (excelTable == null) continue;

                    var conn = GetConnectionForCatalog(catalog, runningInstances, catalogConnections);
                    if (conn == null)
                    {
                        result.Skipped++;
                        continue;
                    }

                    try
                    {
                        Application.StatusBar = $"Refreshing '{tblName}'...";
                        int rows = UpdateTableFromPowerBI(excelTable, pbiName, conn);
                        if (rows >= 0)
                        {
                            result.TablesRefreshed++;
                            result.TotalRows += rows;
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(
                            $"Error refreshing table '{tblName}' on sheet '{wsName}':\n{ex.Message}",
                            "Refresh Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            finally
            {
                DisposeTempConnections(catalogConnections);
            }

            if (result.Skipped > 0)
            {
                MessageBox.Show(
                    $"{result.Skipped} table(s) skipped because their Power BI Desktop instance is not running.",
                    "Refresh All", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            return result;
        }

        private Worksheet GetMetadataSheet(Workbook workbook)
        {
            try { return workbook.Worksheets["PBIDesktop_Metadata"] as Worksheet; }
            catch { return null; }
        }

        private int GetMetadataLastRow(Worksheet metaSheet)
        {
            Range lastCell = null;
            try
            {
                lastCell = metaSheet.Cells[metaSheet.Rows.Count, 1].End[XlDirection.xlUp];
                int row = lastCell.Row;
                if (row == 1)
                {
                    string headerVal = GetCellString(metaSheet, 1, 1);
                    if (string.IsNullOrEmpty(headerVal))
                        return 0;
                }
                return row;
            }
            finally
            {
                ReleaseCom(lastCell);
            }
        }

        private static ListObject FindListObject(Worksheet ws, string tableName)
        {
            if (string.IsNullOrEmpty(tableName)) return null;
            ListObjects listObjects = null;
            try
            {
                listObjects = ws.ListObjects;
                foreach (ListObject lo in listObjects)
                {
                    if (lo.Name == tableName)
                        return lo;
                }
                return null;
            }
            finally
            {
                ReleaseCom(listObjects);
            }
        }

        private AdomdConnection GetConnectionForCatalog(
            string catalog,
            List<(int port, string catalog)> runningInstances,
            Dictionary<string, AdomdConnection> catalogConnections)
        {
            if (!string.IsNullOrEmpty(catalog))
            {
                if (catalogConnections.TryGetValue(catalog, out var existing))
                    return existing;

                var instance = runningInstances.FirstOrDefault(i =>
                    string.Equals(i.catalog, catalog, StringComparison.OrdinalIgnoreCase));

                if (instance.catalog != null)
                {
                    try
                    {
                        string connStr =
                            $"Provider=MSOLAP;Data Source=localhost:{instance.port};" +
                            $"Initial Catalog={instance.catalog};Integrated Security=SSPI;" +
                            "Impersonation Level=Impersonate;";
                        var conn = new AdomdConnection(connStr);
                        conn.Open();
                        catalogConnections[catalog] = conn;
                        return conn;
                    }
                    catch { }
                }
            }

            if (HasActiveConnection())
                return _connection;

            return null;
        }

        private void DisposeTempConnections(Dictionary<string, AdomdConnection> catalogConnections)
        {
            foreach (var conn in catalogConnections.Values)
            {
                if (!ReferenceEquals(conn, _connection))
                {
                    try { conn.Dispose(); }
                    catch { }
                }
            }
        }

        private int UpdateTableFromPowerBI(ListObject excelTable, string pbiTable, AdomdConnection connection)
        {
            var qr = ExecuteTableQuery(pbiTable, connection);
            if (qr.ColCount == 0) return -1;

            Range bodyRange = null;
            Range anchorCell = null;
            Range insertStart = null;
            Range insertRange = null;
            Range colRange = null;

            try
            {
                var worksheet = excelTable.Parent as Worksheet;
                anchorCell = (Range)excelTable.Range.Cells[1, 1];

                bodyRange = excelTable.DataBodyRange;
                if (bodyRange != null)
                {
                    int dataRows = bodyRange.Rows.Count;
                    if (dataRows > 0)
                        bodyRange.Delete();
                    ReleaseCom(bodyRange);
                    bodyRange = null;
                }

                if (qr.Rows.Count > 0)
                {
                    var dataArr = To2DArray(qr.Rows, qr.ColCount);
                    insertStart = anchorCell.Offset[1, 0];
                    insertRange = insertStart.Resize[qr.Rows.Count, qr.ColCount];
                    insertRange.Value2 = dataArr;

                    for (int i = 0; i < qr.ColCount; i++)
                    {
                        if (qr.ColumnTypes[i] == typeof(DateTime))
                        {
                            colRange = insertStart.Offset[0, i].Resize[qr.Rows.Count, 1];
                            colRange.NumberFormat = "yyyy-MM-dd";
                            ReleaseCom(colRange);
                            colRange = null;
                        }
                    }
                }

                Range fitRange = excelTable.Range.Columns;
                fitRange.AutoFit();
                ReleaseCom(fitRange);
            }
            finally
            {
                ReleaseCom(colRange);
                ReleaseCom(insertRange);
                ReleaseCom(insertStart);
                ReleaseCom(anchorCell);
                ReleaseCom(bodyRange);
            }

            return qr.Rows.Count;
        }

        public List<string> GetAvailableTableNames()
        {
            if (!HasActiveConnection())
                return new List<string>();

            var result = new List<string>();

            using (var cmd = _connection.CreateCommand())
            {
                cmd.CommandText = "SELECT * FROM $SYSTEM.TMSCHEMA_TABLES";

                using (var reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        string tableName = reader["Name"]?.ToString();
                        if (!string.IsNullOrWhiteSpace(tableName))
                            result.Add(tableName);
                    }
                }
            }

            return result;
        }

        public int ImportTable(string pbiTable, bool createSheet, string targetCell)
        {
            var workbook = Application.ActiveWorkbook;
            Worksheet worksheet = null;
            Range anchor = null;
            Range headerRange = null;
            Range dataRange = null;
            Range fitRange = null;
            Range dateCol = null;

            try
            {
                var qr = ExecuteTableQuery(pbiTable, _connection);
                if (qr.ColCount == 0)
                {
                    MessageBox.Show("The table has no columns.", "Import",
                        MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return -1;
                }

                if (createSheet)
                {
                    worksheet = workbook.Worksheets.Add() as Worksheet;
                    worksheet.Name = GetUniqueSheetName(workbook, SanitizeSheetName(pbiTable));
                }
                else
                {
                    worksheet = workbook.ActiveSheet as Worksheet;
                }

                anchor = worksheet.Range[targetCell];

                var headerArr = new object[1, qr.ColCount];
                for (int i = 0; i < qr.ColCount; i++)
                    headerArr[0, i] = qr.Headers[i];
                headerRange = anchor.Resize[1, qr.ColCount];
                headerRange.Value2 = headerArr;

                if (qr.Rows.Count > 0)
                {
                    var dataArr = To2DArray(qr.Rows, qr.ColCount);
                    Range insertStart = anchor.Offset[1, 0];
                    Range insertRange = insertStart.Resize[qr.Rows.Count, qr.ColCount];
                    insertRange.Value2 = dataArr;

                    for (int i = 0; i < qr.ColCount; i++)
                    {
                        if (qr.ColumnTypes[i] == typeof(DateTime))
                        {
                            dateCol = insertStart.Offset[0, i].Resize[qr.Rows.Count, 1];
                            dateCol.NumberFormat = "yyyy-MM-dd";
                            ReleaseCom(dateCol);
                            dateCol = null;
                        }
                    }

                    ReleaseCom(insertRange);
                    ReleaseCom(insertStart);
                }

                dataRange = worksheet.Range[anchor, anchor.Offset[qr.Rows.Count, qr.ColCount - 1]];
                var table = worksheet.ListObjects.Add(
                    XlListObjectSourceType.xlSrcRange, dataRange, Type.Missing, XlYesNoGuess.xlYes);
                table.Name = GetUniqueTableName(worksheet, $"tbl_{pbiTable}");

                try { table.TableStyle = "TableStyleMedium9"; }
                catch { }

                fitRange = dataRange.Columns;
                fitRange.AutoFit();

                WriteOrUpdateMetadata(worksheet.Name, table.Name, pbiTable,
                    anchor.Address[false, false], ActiveCatalog);

                return qr.Rows.Count;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Import failed:\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                if (createSheet && worksheet != null)
                {
                    try { worksheet.Delete(); }
                    catch { }
                }
                return -1;
            }
            finally
            {
                ReleaseCom(dateCol);
                ReleaseCom(fitRange);
                ReleaseCom(dataRange);
                ReleaseCom(headerRange);
                ReleaseCom(anchor);
            }
        }

        private static string SanitizeSheetName(string name)
        {
            if (string.IsNullOrEmpty(name)) return "Sheet";
            char[] illegal = { '\\', '/', '?', '*', '[', ']', ':' };
            foreach (char c in illegal)
                name = name.Replace(c.ToString(), "");
            name = name.Trim().Trim('\'');
            if (name.Length > 31) name = name.Substring(0, 31);
            if (string.IsNullOrWhiteSpace(name)) name = "Sheet";
            return name;
        }

        private static string GetUniqueSheetName(Workbook workbook, string baseName)
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (Worksheet ws in workbook.Worksheets)
                existing.Add(ws.Name);

            if (!existing.Contains(baseName))
                return baseName;

            for (int i = 2; i < 1000; i++)
            {
                string suffix = $" ({i})";
                int maxBase = 31 - suffix.Length;
                string candidate = (baseName.Length > maxBase ? baseName.Substring(0, maxBase) : baseName) + suffix;
                if (!existing.Contains(candidate))
                    return candidate;
            }

            return baseName;
        }

        private static string GetUniqueTableName(Worksheet worksheet, string baseName)
        {
            var existing = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            ListObjects los = null;
            try
            {
                los = worksheet.ListObjects;
                foreach (ListObject lo in los)
                    existing.Add(lo.Name);
            }
            finally
            {
                ReleaseCom(los);
            }

            if (!existing.Contains(baseName))
                return baseName;

            for (int i = 2; i < 1000; i++)
            {
                string candidate = $"{baseName}_{i}";
                if (!existing.Contains(candidate))
                    return candidate;
            }

            return baseName;
        }

        private void WriteOrUpdateMetadata(string wsName, string excelTable,
            string pbiTable, string anchorCell, string catalog)
        {
            var workbook = Application.ActiveWorkbook;
            Worksheet metaSheet = GetMetadataSheet(workbook);

            if (metaSheet == null)
            {
                metaSheet = workbook.Worksheets.Add() as Worksheet;
                metaSheet.Name = "PBIDesktop_Metadata";
                metaSheet.Visible = XlSheetVisibility.xlSheetVeryHidden;
                SetCellValue(metaSheet, 1, 1, "Worksheet");
                SetCellValue(metaSheet, 1, 2, "ExcelTableName");
                SetCellValue(metaSheet, 1, 3, "PowerBITableName");
                SetCellValue(metaSheet, 1, 4, "AnchorCell");
                SetCellValue(metaSheet, 1, 5, "Catalog");
            }

            int lastRow = GetMetadataLastRow(metaSheet);
            if (lastRow < 1) lastRow = 1;

            bool found = false;
            for (int row = 2; row <= lastRow; row++)
            {
                string sheet = GetCellString(metaSheet, row, 1);
                string table = GetCellString(metaSheet, row, 2);
                if (sheet == wsName && table == excelTable)
                {
                    SetCellValue(metaSheet, row, 3, pbiTable);
                    SetCellValue(metaSheet, row, 4, anchorCell);
                    SetCellValue(metaSheet, row, 5, catalog);
                    found = true;
                    break;
                }
            }

            if (!found)
            {
                int newRow = lastRow + 1;
                SetCellValue(metaSheet, newRow, 1, wsName);
                SetCellValue(metaSheet, newRow, 2, excelTable);
                SetCellValue(metaSheet, newRow, 3, pbiTable);
                SetCellValue(metaSheet, newRow, 4, anchorCell);
                SetCellValue(metaSheet, newRow, 5, catalog);
            }
        }

        public static List<(int port, string catalog)> FindPowerBIPortFiles()
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

                string[] dirs;
                try { dirs = Directory.GetDirectories(baseDir); }
                catch (IOException) { continue; }
                catch (UnauthorizedAccessException) { continue; }

                foreach (var dir in dirs)
                {
                    var portPath = Path.Combine(dir, "Data", "msmdsrv.port.txt");

                    string raw;
                    try
                    {
                        try
                        {
                            raw = File.ReadAllText(portPath, System.Text.Encoding.Unicode);
                        }
                        catch (IOException)
                        {
                            System.Threading.Thread.Sleep(100);
                            raw = File.ReadAllText(portPath, System.Text.Encoding.Unicode);
                        }
                    }
                    catch { continue; }

                    string portText = new string(raw.Where(char.IsDigit).ToArray());
                    if (!int.TryParse(portText, out int port)) continue;

                    try
                    {
                        string connStr =
                            $"Provider=MSOLAP;Data Source=localhost:{port};Integrated Security=SSPI;";

                        using (var conn = new AdomdConnection(connStr))
                        {
                            conn.Open();
                            var catalogs = new List<string>();

                            using (var cmd = conn.CreateCommand())
                            {
                                cmd.CommandText = "SELECT * FROM $SYSTEM.DBSCHEMA_CATALOGS";
                                using (var reader = cmd.ExecuteReader())
                                {
                                    while (reader.Read())
                                    {
                                        var name = reader["CATALOG_NAME"]?.ToString();
                                        if (!string.IsNullOrWhiteSpace(name))
                                            catalogs.Add(name);
                                    }
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

        private string RemovePrefix(string column, string table)
        {
            string prefix = table + "[";
            if (column.StartsWith(prefix) && column.EndsWith("]"))
                return column.Substring(prefix.Length, column.Length - prefix.Length - 1);
            return column;
        }

        private static object[,] To2DArray(List<object[]> rows, int colCount)
        {
            if (rows == null || rows.Count == 0 || colCount == 0)
                return new object[0, 0];

            int rowCount = rows.Count;
            var array = new object[rowCount, colCount];
            for (int r = 0; r < rowCount; r++)
            {
                var row = rows[r];
                int len = row != null ? Math.Min(row.Length, colCount) : 0;
                for (int c = 0; c < len; c++)
                    array[r, c] = row[c];
            }
            return array;
        }

        private static string EscapeDaxIdentifier(string name)
        {
            if (string.IsNullOrEmpty(name))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(name));
            return $"'{name.Replace("'", "''")}'";
        }

        private void InternalStartup()
        {
            this.Startup += new EventHandler(ThisAddIn_Startup);
            this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
        }
    }
}
