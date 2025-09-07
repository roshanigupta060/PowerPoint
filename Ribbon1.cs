using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using DataTable = System.Data.DataTable;
using MessageBox = System.Windows.Forms.MessageBox;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Web;

namespace PptExcelSync
{
    public partial class Ribbon1
    {
        private string selectedChartType = "Column";
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            LoadDatasetsIntoDropdown();
        }

        private void btnUploadExcel_Click(object sender, RibbonControlEventArgs e)
        {
            // Open File Dialog
            //OpenFileDialog ofd = new OpenFileDialog();
            //ofd.Filter = "Excel Files|*.xlsx;*.xls;*.csv";
            //if (ofd.ShowDialog() == DialogResult.OK)
            //{
            //    string filePath = ofd.FileName;
            //    // TODO: Upload to Google Cloud
            //    var storage = StorageClient.Create();
            //    var fileStream = File.OpenRead(filePath);
            //    storage.UploadObject("your-bucket", Path.GetFileName(filePath), null, fileStream);
            //}

            using (var ofd = new System.Windows.Forms.OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xlsx;*.xls;*.csv";
                if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        // 1. Define local storage path (e.g., Documents\PptExcelSync\datasets)
                        string datasetsPath = Path.Combine(
                            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                            "PptExcelSync", "datasets");

                        // 2. Ensure folder exists
                        Directory.CreateDirectory(datasetsPath);

                        // 3. Copy the selected file into that folder
                        string destPath = Path.Combine(datasetsPath, Path.GetFileName(ofd.FileName));
                        File.Copy(ofd.FileName, destPath, overwrite: true);

                        // 4. (Optional) Save some metadata alongside it
                        string metaPath = destPath + ".meta.txt";
                        File.WriteAllText(metaPath,
                            $"uploadedBy={Environment.UserName}\r\nuploadedAt={DateTime.UtcNow:o}");
                        //5. Refresh dropdown after upload
                LoadDatasetsIntoDropdown();

                        // 6. Notify user
                        System.Windows.Forms.MessageBox.Show($"File stored locally:\n{destPath}");
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.MessageBox.Show($"Error saving file: {ex.Message}");
                    }
                }
            }
        }

        private void LoadDatasetsIntoDropdown()
        {
            ddlDatasets.Items.Clear();

            string datasetsPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                "PptExcelSync", "datasets");

            if (!Directory.Exists(datasetsPath)) return;

            var files = Directory.GetFiles(datasetsPath, "*.*")
                                 .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                             f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                             f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                                 .ToList();
            var item = this.Factory.CreateRibbonDropDownItem();
            item.Label = "-- select --";
            item.Tag = "select";
            ddlDatasets.Items.Add(item);

            foreach (var file in files)
            {
                var value = this.Factory.CreateRibbonDropDownItem();

                value.Label = Path.GetFileName(file);
                value.Tag = file;
                ddlDatasets.Items.Add(value);
            }
        }

        private void btnListDatasets_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                // 1. Define local storage path
                string datasetsPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                    "PptExcelSync", "datasets");

                // 2. Check if folder exists
                if (!Directory.Exists(datasetsPath))
                {
                    System.Windows.Forms.MessageBox.Show("No datasets folder found yet.");
                    return;
                }

                // 3. Get all Excel/CSV files
                var files = Directory.GetFiles(datasetsPath, "*.*")
                                     .Where(f => f.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                                                 f.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                                 f.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                                     .ToList();

                if (files.Count == 0)
                {
                    System.Windows.Forms.MessageBox.Show("No datasets found.");
                    return;
                }

                // 4. Show file list (simple popup for now)
                string msg = "Available datasets:\n\n" + string.Join("\n", files.Select(Path.GetFileName));
                System.Windows.Forms.MessageBox.Show(msg, "Datasets Found");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error listing datasets: {ex.Message}");
            }
        }

        private void ddlDatasets_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            if (ddlDatasets.SelectedItem == null) return;
        }

        private void InsertChartFromDataset(string filePath, string chartType)
        {

            // choose chart type
            SeriesChartType type = SeriesChartType.Column; // default
            switch (chartType.ToLower())
            {
                case "line": type = SeriesChartType.Line; break;
                case "pie": type = SeriesChartType.Pie; break;
                case "bar": type = SeriesChartType.Bar; break;
            }

            var dt = new DatasetManager().LoadExcel(filePath);

            if (dt.Columns.Count < 2)
            {
                System.Windows.Forms.MessageBox.Show("Need at least 2 columns (labels + values).");
                return;
            }

            string xCol = dt.Columns[0].ColumnName; // first column is labels
            var labels = dt.AsEnumerable().Select(r => r[xCol].ToString()).ToArray();

            var chart = new System.Windows.Forms.DataVisualization.Charting.Chart
            {
                Width = 800,
                Height = 400
            };
            chart.ChartAreas.Add(new ChartArea("MainArea"));

            chart.ChartAreas["MainArea"].AxisX.Title = xCol;
            chart.ChartAreas["MainArea"].AxisX.Interval = 1;
            chart.ChartAreas["MainArea"].AxisX.MajorGrid.LineColor = Color.LightGray;
            chart.ChartAreas["MainArea"].AxisY.MajorGrid.LineColor = Color.LightGray;

            // Loop over remaining columns and add each as a series
            for (int col = 1; col < dt.Columns.Count; col++)
            {
                string yCol = dt.Columns[col].ColumnName;

                // Only try numeric columns
                var values = dt.AsEnumerable()
                               .Select(r =>
                               {
                                   double val;
                                   return double.TryParse(r[yCol].ToString(), out val) ? val : 0;
                               })
                               .ToArray();

                var series = new Series(yCol)
                {
                    ChartType = type,
                    IsValueShownAsLabel = true
                };

                for (int i = 0; i < labels.Length; i++)
                {
                    series.Points.AddXY(labels[i], values[i]);
                }

                chart.Series.Add(series);
            }

            // Save chart as image
            string chartPath = Path.Combine(Path.GetTempPath(), "chart.png");
            chart.SaveImage(chartPath, ChartImageFormat.Png);

            // Insert into PowerPoint
            var app = Globals.ThisAddIn.Application;
            var slide = app.ActivePresentation.Slides.Add(
                app.ActivePresentation.Slides.Count + 1,
                Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

            slide.Shapes.AddPicture(chartPath,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoCTrue,
                100, 100, 600, 300);
        }

        private void InsertTableFromDataset(string filePath, string chartType)
        {
            try
            {
                // Read Excel with ClosedXML
                var dt = new System.Data.DataTable();
                using (var wb = new ClosedXML.Excel.XLWorkbook(filePath))
                {
                    var ws = wb.Worksheet(1);
                    var firstRow = ws.FirstRowUsed();

                    // Columns
                    foreach (var cell in firstRow.CellsUsed())
                        dt.Columns.Add(cell.GetString());

                    // Data
                    foreach (var row in ws.RowsUsed().Skip(1))
                    {
                        var dr = dt.NewRow();
                        for (int i = 0; i < dt.Columns.Count; i++)
                            dr[i] = row.Cell(i + 1).GetValue<string>();
                        dt.Rows.Add(dr);
                    }
                }

                // Insert into current slide
                var app = Globals.ThisAddIn.Application;
                var slide = app.ActivePresentation.Slides.Add(
                    app.ActivePresentation.Slides.Count + 1,
                    Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

                // Remove placeholders if any (safety)
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPlaceholder)
                        shape.Delete();
                }

                int rows = dt.Rows.Count + 1, cols = dt.Columns.Count;
                var table = slide.Shapes.AddTable(rows, cols, 50, 50, 600, 300).Table;

                // headers
                for (int c = 0; c < cols; c++)
                    table.Cell(1, c + 1).Shape.TextFrame.TextRange.Text = dt.Columns[c].ColumnName;

                // data
                for (int r = 0; r < dt.Rows.Count; r++)
                    for (int c = 0; c < cols; c++)
                        table.Cell(r + 2, c + 1).Shape.TextFrame.TextRange.Text = dt.Rows[r][c]?.ToString();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show($"Error inserting dataset: {ex.Message}");
            }
        }

        private void btnInsertChart_Click_Click(object sender, RibbonControlEventArgs e)
        {
            if (ddlDatasets.SelectedItem == null || ddlDatasets.SelectedItem.Label == "-- select --")
            {
                MessageBox.Show("Please select a dataset first.");
                return;
            }

            string filePath = ddlDatasets.SelectedItem.Tag.ToString();
            InsertChartFromDataset(filePath,"column");
        }

        private void btnInsertTable_Click_Click(object sender, RibbonControlEventArgs e)
        {
            if (ddlDatasets.SelectedItem == null || ddlDatasets.SelectedItem.Label == "-- select --")
            {
                MessageBox.Show("Please select a dataset first.");
                return;
            }

            string filePath = ddlDatasets.SelectedItem.Tag.ToString();
            InsertTableFromDataset(filePath, "Table");
        }
       
        private void ddlChartType_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            var dropdown = (RibbonDropDown)sender;
            selectedChartType = dropdown.SelectedItem.Label; 
        }

        private void btnCreateChart_Click_Click(object sender, RibbonControlEventArgs e)
        {
            if (ddlChartType.SelectedItem == null || ddlDatasets.SelectedItem.Label == "-- select --")
            {
                MessageBox.Show("Please select a dropdown option first.");
                return;
            }
            string filePath = ddlDatasets.SelectedItem.Tag.ToString();

            if (!string.IsNullOrEmpty(filePath))
            {
                if (selectedChartType != "Table")
                    InsertChartFromDataset(filePath, selectedChartType);
                else
                    InsertTableFromDataset(filePath, selectedChartType);
            }
        }
        private void btnPivotView_Click(object sender, RibbonControlEventArgs e)
        {
            // Ask user to pick dataset
            string filePath = ddlDatasets.SelectedItem.Tag.ToString(); // reuse your dataset dropdown
            if (string.IsNullOrEmpty(filePath) || ddlDatasets.SelectedItem.Label == "-- select --")
            {
                MessageBox.Show("Please select a dataset first.");
                return;
            }

            var dt = new DatasetManager().LoadExcel(filePath);

            // Show Pivot dialog
            var form = new Pivot(dt);
            if (form.ShowDialog() == DialogResult.OK)
            {
                var pivot = CreatePivot(dt, form.SelectedRowField, form.SelectedValueFields, form.SelectedAggregations);

                // Insert pivot into PowerPoint
                string chartType =  form.SelectedChartType.ToString();

                if(chartType != "0")
                  InsertPivotChartIntoPowerPoint(pivot, form.SelectedChartType);
                else
                  InsertTableIntoPowerPoint(pivot, 25);
            }
        }

        public DataTable CreatePivot(DataTable dt, string rowField, List<string> valueFields, List<string> aggFuncs)
        {
            var grouped = dt.AsEnumerable().GroupBy(r => r[rowField].ToString());

            DataTable pivot = new DataTable();
            pivot.Columns.Add(rowField, typeof(string));

            // Add output columns for each (aggregation + value field) combination
            foreach (var valField in valueFields)
            {
                foreach (var agg in aggFuncs)
                {
                    pivot.Columns.Add($"{agg} of {valField}", typeof(double));
                }
            }

            foreach (var g in grouped)
            {
                var row = pivot.NewRow();
                row[rowField] = g.Key;

                foreach (var valField in valueFields)
                {
                    var numbers = g.Select(r =>
                    {
                        double val;
                        return double.TryParse(r[valField].ToString(), out val) ? val : 0;
                    });

                    foreach (var agg in aggFuncs)
                    {
                        double result = 0;
                        switch (agg.ToLower())
                        {
                            case "sum": result = numbers.Sum(); break;
                            case "average": result = numbers.Any() ? numbers.Average() : 0; break;
                            case "count": result = g.Count(); break;
                            case "max": result = numbers.Any() ? numbers.Max() : 0; break;
                            case "min": result = numbers.Any() ? numbers.Min() : 0; break;
                        }

                        row[$"{agg} of {valField}"] = result;
                    }
                }

                pivot.Rows.Add(row);
            }

            return pivot;
        }

        /// <summary>
        /// Insert a DataTable into PowerPoint slides as a table. Splits into multiple slides if too many rows.
        /// </summary>
        public void InsertTableIntoPowerPoint(DataTable dt, int maxDataRowsPerSlide = 20)
        {
            if (dt == null || dt.Columns.Count == 0)
            {
                MessageBox.Show("No data to insert.");
                return;
            }

            try
            {
                var app = Globals.ThisAddIn.Application;

                // Ensure we have a presentation
                PowerPoint.Presentation pres = null;
                if (app.Presentations.Count == 0)
                    pres = app.Presentations.Add(Office.MsoTriState.msoTrue);
                else
                    pres = app.ActivePresentation;

                int totalRows = dt.Rows.Count;
                int cols = dt.Columns.Count;
                int processedRows = 0;

                // Slide dimensions
                float slideW = (float)pres.PageSetup.SlideWidth;
                float slideH = (float)pres.PageSetup.SlideHeight;
                float marginLeft = 40f;
                float marginTop = 60f;
                float tableWidth = slideW - marginLeft * 2;
                float tableHeight = slideH - marginTop * 2 - 20f; // leave some bottom margin

                while (processedRows < totalRows || (totalRows == 0 && processedRows == 0))
                {
                    int rowsThisSlice = Math.Min(maxDataRowsPerSlide, totalRows - processedRows);
                    // If no data rows (empty dt) still create header-only table once
                    int tableRows = Math.Max(1, rowsThisSlice) + 1; // +1 for header

                    // Create a new blank slide
                    var slide = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                    // Safety: remove placeholders if any
                    foreach (PowerPoint.Shape shp in slide.Shapes)
                    {
                        if (shp.Type == Office.MsoShapeType.msoPlaceholder)
                            shp.Delete();
                    }

                    // Add PPT table
                    var shape = slide.Shapes.AddTable(tableRows, cols, marginLeft, marginTop, tableWidth, tableHeight);
                    var table = shape.Table;

                    // Set approximate column widths
                    float colWidth = tableWidth / cols;
                    try
                    {
                        for (int c = 1; c <= cols; c++)
                        {
                            table.Columns[c].Width = colWidth;
                        }
                    }
                    catch
                    {
                        // Some PowerPoint versions might not allow setting width this way; ignore if it fails.
                    }

                    // Header row formatting
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = table.Cell(1, c + 1); // first row is header

                        // Set header text
                        cell.Shape.TextFrame.TextRange.Text = dt.Columns[c].ColumnName;
                        cell.Shape.TextFrame.TextRange.Font.Bold = Office.MsoTriState.msoTrue;
                        cell.Shape.TextFrame.TextRange.Font.Size = 12;

                        // Header background color
                        cell.Shape.Fill.Visible = Office.MsoTriState.msoTrue;
                        cell.Shape.Fill.BackColor.RGB = ColorTranslator.ToOle(Color.SteelBlue);

                        // Header text color (white)
                        cell.Shape.TextFrame.TextRange.Font.Color.RGB = ColorTranslator.ToOle(Color.White);

                        // Center align header text
                        cell.Shape.TextFrame.VerticalAnchor = Office.MsoVerticalAnchor.msoAnchorMiddle;
                        cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment =
                            PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    }


                    // Fill data rows
                    for (int r = 0; r < rowsThisSlice; r++)
                    {
                        int dataRowIndex = processedRows + r;
                        for (int c = 0; c < cols; c++)
                        {
                            object val = dt.Rows[dataRowIndex][c];
                            string text = val?.ToString() ?? string.Empty;

                            var cell = table.Cell(r + 2, c + 1);
                            cell.Shape.TextFrame.TextRange.Text = text;
                            cell.Shape.TextFrame.TextRange.Font.Size = 10;
                            // left-align non-numeric, right-align numeric
                            double dummy;
                            if (double.TryParse(text, out dummy))
                                cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignRight;
                            else
                                cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                        }
                    }

                    // Move to next slice
                    processedRows += rowsThisSlice;
                    if (totalRows == 0) break; // handled empty table case
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting table into PowerPoint: {ex.Message}");
            }
        }

        public void InsertPivotChartIntoPowerPoint(DataTable pivotTable, Office.XlChartType chartType)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var pres = app.Presentations.Count > 0
                    ? app.ActivePresentation
                    : app.Presentations.Add(Office.MsoTriState.msoTrue);

                var slide = pres.Slides.Add(pres.Slides.Count + 1, PowerPoint.PpSlideLayout.ppLayoutBlank);

                var chartShape = slide.Shapes.AddChart(chartType, 50, 50, 600, 350);
                var chart = chartShape.Chart;

                var workbook = chart.ChartData.Workbook;
                var sheet = workbook.Worksheets[1];
                sheet.Cells.Clear();

                int rows = pivotTable.Rows.Count;
                int cols = pivotTable.Columns.Count;

                // Write headers
                for (int c = 0; c < cols; c++)
                    sheet.Cells[1, c + 1] = pivotTable.Columns[c].ColumnName;

                // Write data
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                    {
                        double val;
                        if (double.TryParse(pivotTable.Rows[r][c].ToString(), out val))
                            sheet.Cells[r + 2, c + 1] = val;
                        else
                            sheet.Cells[r + 2, c + 1] = pivotTable.Rows[r][c]?.ToString() ?? "";
                    }

                // Refresh chart
                chart.ChartData.Activate();
                chart.ChartData.Workbook.Application.CalculateFull();
                chart.Refresh();

                chart.HasLegend = true;

                // Enable title safely
                if (!chart.HasTitle)
                    chart.HasTitle = true;

                chart.ChartTitle.Text = "Pivot Chart";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error inserting chart: {ex.Message}");
            }
        }
    }
}
