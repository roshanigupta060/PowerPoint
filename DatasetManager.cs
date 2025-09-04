using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using ClosedXML.Excel;


namespace PptExcelSync
{
    public class DatasetManager
    {
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

            return dt;
        }
    }

}
