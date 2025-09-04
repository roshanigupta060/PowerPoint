using DocumentFormat.OpenXml.Drawing.Charts;
using Google.Cloud.Storage.V1;
using LiveCharts;
using LiveCharts.Wpf;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using DataTable = System.Data.DataTable;
using MessageBox = System.Windows.Forms.MessageBox;
using Series = System.Windows.Forms.DataVisualization.Charting.Series;

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
                var pivot = CreatePivot(dt, form.SelectedRowField, form.SelectedValueField, form.SelectedAggregation);

                // Insert pivot into PowerPoint
               // InsertTableIntoPowerPoint(pivot);
            }
        }

        public DataTable CreatePivot(DataTable dt, string rowField, string valueField, string aggFunc)
        {
            var query = dt.AsEnumerable()
                .GroupBy(r => r[rowField].ToString())
                .Select(g =>
                {
                    double result = 0;
                    switch (aggFunc.ToLower())
                    {
                        case "sum":
                            result = g.Sum(r => double.TryParse(r[valueField].ToString(), out var v) ? v : 0);
                            break;
                        case "average":
                            var nums = g.Select(r => double.TryParse(r[valueField].ToString(), out var v) ? v : 0);
                            result = nums.Any() ? nums.Average() : 0;
                            break;
                        case "count":
                            result = g.Count();
                            break;
                    }
                    return new { Key = g.Key, Value = result };
                });

            DataTable pivot = new DataTable();
            pivot.Columns.Add(rowField);
            pivot.Columns.Add($"{aggFunc} of {valueField}", typeof(double));

            foreach (var item in query)
            {
                pivot.Rows.Add(item.Key, item.Value);
            }

            return pivot;
        }


    }
}
