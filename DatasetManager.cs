using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;


namespace PptExcelSync
{
    public class DatasetManager
    {
        public string FileName { get; set; }
        public DataTable LoadExcel(string filePath)
        {
            if (!File.Exists(filePath))
                throw new FileNotFoundException($"Excel file not found: {filePath}");

            var dt = new DataTable();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);
                var firstRow = ws.FirstRowUsed();

                // Add columns
                foreach (var cell in firstRow.CellsUsed())
                    dt.Columns.Add(cell.GetString());

                // Add rows
                foreach (var row in ws.RowsUsed().Skip(1))
                {
                    var dr = dt.NewRow();
                    for (int i = 0; i < dt.Columns.Count; i++)
                        dr[i] = row.Cell(i + 1).GetValue<string>();
                    dt.Rows.Add(dr);
                }
            }

            // 🔹 STEP 2: Load metadata & apply calculated fields
            var metadata = DatasetMetadata.Load(filePath);

            var calcHelper = new PivotHelper();
            foreach (var field in metadata.CalculatedFields)
            {
                calcHelper.AddCalculatedField(dt, field.FieldName, field.Formula);
            }
            return dt;
        }

        private DataTable LoadCsv(string filePath)
        {
            var dt = new DataTable();

            using (var reader = new StreamReader(filePath))
            {
                bool isFirstRow = true;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    var values = line.Split(',');

                    if (isFirstRow)
                    {
                        foreach (var col in values)
                            dt.Columns.Add(col.Trim());
                        isFirstRow = false;
                    }
                    else
                    {
                        dt.Rows.Add(values);
                    }
                }
            }
            return dt;
        }

        public DataTable LoadDataset(string filePath)
        {
            string ext = Path.GetExtension(filePath).ToLower();

            if (ext == ".csv")
                return LoadCsv(filePath);
            else if (ext == ".xlsx" || ext == ".xls")
                return LoadExcel(filePath);
            else
                throw new NotSupportedException($"Unsupported file type: {ext}");
        }

        public DataTable LoadDatasets(List<string> filePaths)
        {
            DataTable merged = null;

            foreach (var filePath in filePaths)
            {
                var dt = LoadDataset(filePath);

                // Ensure we have a Source column
                if (!dt.Columns.Contains("Source"))
                    dt.Columns.Add("Source", typeof(string));

                foreach (DataRow row in dt.Rows)
                    row["Source"] = Path.GetFileNameWithoutExtension(filePath);

                if (merged == null)
                    merged = dt.Clone(); // copy structure
                merged.Merge(dt);
            }

            return merged;
        }

        public DataTable LoadAndMergeDatasets(IEnumerable<string> filePaths)
        {
            // Ensure we have an indexable list
            var fileList = (filePaths ?? Enumerable.Empty<string>()).Where(p => !string.IsNullOrWhiteSpace(p)).ToList();
            var merged = new DataTable();

            // Keep column insertion order and provide case-insensitive uniqueness
            var columnSet = new List<string>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // Hold the loaded tables
            var tables = new List<DataTable>();

            // Load each file using your generic loader (LoadDataset should handle csv/xlsx)
            foreach (var path in fileList)
            {
                if (!File.Exists(path)) continue;

                // Use your existing universal loader. Replace with LoadCsv/LoadExcel if needed.
                DataTable dt = LoadDataset(path);
                if (dt == null || dt.Columns.Count == 0) continue;

                tables.Add(dt);

                // Collect unique column names (case-insensitive)
                foreach (DataColumn c in dt.Columns)
                {
                    var name = (c.ColumnName ?? "").Trim();
                    if (string.IsNullOrEmpty(name)) continue;
                    if (seen.Add(name))
                        columnSet.Add(name);
                }
            }

            // Build merged DataTable schema (union of columns)
            foreach (var colName in columnSet)
                merged.Columns.Add(colName, typeof(string));

            // Add Source column at the end
            merged.Columns.Add("Source", typeof(string));

            // Append rows from each loaded table, mapping to merged columns
            for (int i = 0; i < tables.Count; i++)
            {
                var dt = tables[i];
                var source = Path.GetFileNameWithoutExtension(fileList[i]) ?? $"File{i}";
                foreach (DataRow r in dt.Rows)
                {
                    var nr = merged.NewRow();
                    // copy known columns
                    foreach (DataColumn c in dt.Columns)
                    {
                        var col = c.ColumnName?.Trim();
                        if (string.IsNullOrEmpty(col)) continue;
                        nr[col] = r[c] != null ? r[c].ToString() : "";
                    }
                    nr["Source"] = source;
                    merged.Rows.Add(nr);
                }
            }

            return merged;
        }

   

    }
}
