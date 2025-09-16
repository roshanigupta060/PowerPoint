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
            var merged = new DataTable();

            // Keep track of column order and set
            var columnSet = new List<string>((IEnumerable<string>)StringComparer.OrdinalIgnoreCase);

            // Read each file into a DataTable
            var tables = new List<DataTable>();

            foreach (var path in filePaths)
            {
                if (!File.Exists(path)) continue;

                DataTable dt;
                var ext = Path.GetExtension(path).ToLowerInvariant();
                if (ext == ".csv")
                    dt = LoadCsv(path);
                else // assume Excel
                    dt = LoadExcel(path);

                if (dt == null || dt.Columns.Count == 0) continue;

                tables.Add(dt);

                // union columns
                foreach (DataColumn c in dt.Columns)
                {
                    if (!columnSet.Contains(c.ColumnName, StringComparer.OrdinalIgnoreCase))
                        columnSet.Add(c.ColumnName);
                }
            }

            // Build merged DataTable with union columns plus 'Source'
            foreach (var colName in columnSet)
                merged.Columns.Add(colName, typeof(string));

            // Add Source column at end
            merged.Columns.Add("Source", typeof(string));

            // Append rows from each table
            for (int i = 0; i < tables.Count; i++)
            {
                var dt = tables[i];
                var source = Path.GetFileNameWithoutExtension(filePaths.ElementAt(i));
                foreach (DataRow r in dt.Rows)
                {
                    var nr = merged.NewRow();
                    foreach (DataColumn c in dt.Columns)
                    {
                        nr[c.ColumnName] = r[c]?.ToString() ?? "";
                    }
                    nr["Source"] = source;
                    merged.Rows.Add(nr);
                }
            }

            return merged;
        }

    }
}
