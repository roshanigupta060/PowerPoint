using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml.Office;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using NCalc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms.VisualStyles;
using ChartArea = System.Windows.Forms.DataVisualization.Charting.ChartArea;
using Color = System.Drawing.Color;
using DataTable = System.Data.DataTable;
using Excel = Microsoft.Office.Interop.Excel;
using MessageBox = System.Windows.Forms.MessageBox;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
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

            var dt = new DatasetManager().LoadDataset(filePath);

            if (dt.Columns.Count < 2)
            {
                System.Windows.Forms.MessageBox.Show("Need at least 2 columns (labels + values).");
                return;
            }

            string xCol = dt.Columns[0].ColumnName; // first column is labels
            var labels = dt.AsEnumerable().Select(r => r[xCol].ToString()).ToArray();

            var chart = new System.Windows.Forms.DataVisualization.Charting.Chart
            {
                Width = 900,
                Height = 400
            };
            chart.ChartAreas.Add(new ChartArea("MainArea"));

            chart.ChartAreas["MainArea"].AxisX.Title = xCol;
            chart.ChartAreas["MainArea"].AxisX.Interval = 1;
            chart.ChartAreas["MainArea"].AxisX.MajorGrid.LineColor = Color.Blue;
            chart.ChartAreas["MainArea"].AxisY.MajorGrid.LineColor = Color.Blue;

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
            //string chartPath = Path.Combine(Path.GetTempPath(), "chart.png");
            //chart.SaveImage(chartPath, ChartImageFormat.Png);

            // Insert into PowerPoint
            var app = Globals.ThisAddIn.Application;
            var slide = app.ActivePresentation.Slides.Add(
                app.ActivePresentation.Slides.Count + 1,
                Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

            //slide.Shapes.AddPicture(chartPath,
            //    Microsoft.Office.Core.MsoTriState.msoFalse,
            //    Microsoft.Office.Core.MsoTriState.msoCTrue,
            //    100, 100, 600, 300);

           // if (config != null)
            {
                slide.Tags.Delete("ChartMakerMeta");
                slide.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(app));
            }
        }

        private void InsertTableFromDataset(string filePath, string chartType)
        {
            try
            {
                // Read Excel with ClosedXML
                //var dt = new System.Data.DataTable();
                var dt = new DatasetManager().LoadDataset(filePath);

  

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
                var shapeTable = slide.Shapes.AddTable(rows, cols, 50, 50, 600, 300);
                var table = shapeTable.Table;

                // headers
                for (int c = 0; c < cols; c++)
                {
                    var cell = table.Cell(1, c + 1);

                    cell.Shape.TextFrame.TextRange.Text = dt.Columns[c].ColumnName;
                }


                // data
                for (int r = 0; r < dt.Rows.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var cell = table.Cell(r + 2, c + 1);
                        cell.Shape.TextFrame.TextRange.Text = dt.Rows[r][c]?.ToString();
                    }
                }

                var config = new PivotConfig
                {
                    DatasetPath = filePath,
                    ChartTypeName    = chartType,  // not really needed for table, but good for consistency
                    RowField = null,
                    ValueFields = null,
                    Aggregations = null
                };

                shapeTable.Tags.Delete("ChartMakerMeta");
                shapeTable.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(config));
                shapeTable.AlternativeText = "ChartMaker|" + config.DatasetPath;

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

            var dt = new DatasetManager().LoadDataset(filePath);

            // Show Pivot dialog
            var form = new Pivot(dt, filePath);
        
            if (form.ShowDialog() == DialogResult.OK)
            {                
                if (form.SelectedValueFields.Count() > 2)
                {
                    MessageBox.Show("Please select only two values.");
                    return;
                }
                var newdt = form._data;
                var filters = form.GetFilters();
                string columnField = form.selectedYOYCompField;
                if (columnField == "-- none --")
                    columnField = null;               

                var dataTable = columnField == null ? dt : newdt;

                var pivot = CreatePivot(dataTable, form.SelectedRowField, form.SelectedValueFields, form.SelectedAggregations, columnField, filters);

                // ⬇️ Get rules from the form
                var rules = form.GetConditionalRules();

                // Insert pivot into PowerPoint
                string chartType =  form.SelectedChartType.ToString();

                string isRepeatedSelected = form.SelectedRepeatByField;

                if(isRepeatedSelected != null)
                CreateRepeatViews(dataTable, form, chartType, rules);

                if (chartType != "0")
                  InsertPivotChartIntoPowerPoint(pivot, form.SelectedChartType, form, rules);
                else
                  InsertTableIntoPowerPoint(pivot, 25, form, rules);
            }
        }


        private Dictionary<string, string> ConvertFilters(Dictionary<string, string> savedFilters)
        {
            var dict = new Dictionary<string, string>();
            if (savedFilters != null)
            {
                foreach (var f in savedFilters)
                {
                    if (!string.IsNullOrWhiteSpace(f.Key) && !string.IsNullOrWhiteSpace(f.Value))
                    {
                        dict[f.Value] = f.Value;
                    }
                }
            }
            return dict;
        }

        public DataTable CreatePivot(      
    DataTable dt, string rowField, List<string> valueFields, List<string> aggFuncs, string columnField = null,
    Dictionary<string, string> filters = null)
        {
            if (dt == null) throw new ArgumentNullException(nameof(dt));
            if (string.IsNullOrEmpty(rowField)) throw new ArgumentNullException(nameof(rowField));

            // --- Step 1: Apply filters if any ---
            IEnumerable<DataRow> rowsQuery = dt.AsEnumerable();
            if (filters != null && filters.Any())
            {
                foreach (var f in filters)
                    rowsQuery = rowsQuery.Where(r => r[f.Key]?.ToString() == f.Value);
            }

           
          

            // If columnField is null or empty => simple pivot (one column per agg/value)
            if (string.IsNullOrEmpty(columnField))
            {
                // --- Step 2: Group by row field ---
                var grouped = rowsQuery.GroupBy(r => r[rowField].ToString());
                var pivot = new DataTable();
                pivot.Columns.Add(rowField, typeof(string));
                foreach (var valField in valueFields)
                    foreach (var agg in aggFuncs)
                        pivot.Columns.Add($"{agg} of {valField}", typeof(double));

                foreach (var g in grouped)
                {
                    var nr = pivot.NewRow();
                    nr[rowField] = g.Key;
                    foreach (var valField in valueFields)
                    {
                        var numbers = g.Select(r => {
                            double v; return double.TryParse(r[valField]?.ToString(), out v) ? v : 0;
                        }).ToList();

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
                            nr[$"{agg} of {valField}"] = result;
                        }
                    }
                    pivot.Rows.Add(nr);
                }
                return pivot;
            }

            if (!dt.Columns.Contains(columnField))  //Source
                throw new ArgumentException($"Column '{columnField}' does not exist in dataset. Available columns: "
                                            + string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => c.ColumnName)));

                                    // --- columnField present (YoY/multi-file) ---
                                    // Determine distinct column groups (for example: "2023","2024" or file names)
                                    var distinctColValues = rowsQuery
                .Select(r => r[columnField]?.ToString() ?? "")
                .Distinct()
                .OrderBy(x => x)
                .ToList();

            // Build pivot columns: RowField + for each columnValue * (for each valueField * agg)
            var pivot2 = new DataTable();
            pivot2.Columns.Add(rowField, typeof(string));

            foreach (var colVal in distinctColValues)
            {
                foreach (var valField in valueFields)
                {
                    foreach (var agg in aggFuncs)
                    {
                        // Include the column value in the header for clarity
                        pivot2.Columns.Add($"{agg} of {valField} [{colVal}]", typeof(double));
                    }
                }
            }

            // Group by row field only, then compute for each distinct columnValue separately
            var groupedRows = rowsQuery.GroupBy(r => r[rowField].ToString());
            foreach (var g in groupedRows)
            {
                var nr = pivot2.NewRow();
                nr[rowField] = g.Key;

                foreach (var colVal in distinctColValues)
                {
                    // rows inside this row group and this column group
                    var rowsInBucket = g.Where(r => (r[columnField]?.ToString() ?? "") == colVal).ToList();

                    foreach (var valField in valueFields)
                    {
                        var numbers = rowsInBucket.Select(r => {
                            double v; return double.TryParse(r[valField]?.ToString(), out v) ? v : 0;
                        }).ToList();

                        foreach (var agg in aggFuncs)
                        {
                            double result = 0;
                            switch (agg.ToLower())
                            {
                                case "sum": result = numbers.Sum(); break;
                                case "average": result = numbers.Any() ? numbers.Average() : 0; break;
                                case "count": result = rowsInBucket.Count; break;
                                case "max": result = numbers.Any() ? numbers.Max() : 0; break;
                                case "min": result = numbers.Any() ? numbers.Min() : 0; break;
                            }
                            nr[$"{agg} of {valField} [{colVal}]"] = result;
                        }
                    }
                }

                pivot2.Rows.Add(nr);
            }

            return pivot2;
        }

        //public DataTable CreatePivot(DataTable dt,string rowField, List<string> valueFields, List<string> aggFuncs,
        //      Dictionary<string, string> filters = null)
        //{
        //    // --- Step 1: Apply filters if provided ---
        //    IEnumerable<DataRow> query = dt.AsEnumerable();
        //    if (filters != null && filters.Any())
        //    {
        //        foreach (var f in filters)
        //        {
        //            string col = f.Key;
        //            string val = f.Value;

        //            query = query.Where(r => r[col]?.ToString() == val);
        //        }
        //    }

        //    // --- Step 2: Group by row field ---
        //    var grouped = query.GroupBy(r => r[rowField].ToString());

        //    // --- Step 3: Build output table ---
        //    DataTable pivot = new DataTable();
        //    pivot.Columns.Add(rowField, typeof(string));

        //    // Add output columns for each (aggregation + value field) combination
        //    foreach (var valField in valueFields)
        //    {
        //        foreach (var agg in aggFuncs)
        //        {
        //            pivot.Columns.Add($"{agg} of {valField}", typeof(double));
        //        }
        //    }

        //    // --- Step 4: Fill data ---
        //    foreach (var g in grouped)
        //    {
        //        var row = pivot.NewRow();
        //        row[rowField] = g.Key;

        //        foreach (var valField in valueFields)
        //        {
        //            var numbers = g.Select(r =>
        //            {
        //                double val;
        //                return double.TryParse(r[valField].ToString(), out val) ? val : 0;
        //            });

        //            foreach (var agg in aggFuncs)
        //            {
        //                double result = 0;
        //                switch (agg.ToLower())
        //                {
        //                    case "sum": result = numbers.Sum(); break;
        //                    case "average": result = numbers.Any() ? numbers.Average() : 0; break;
        //                    case "count": result = g.Count(); break;
        //                    case "max": result = numbers.Any() ? numbers.Max() : 0; break;
        //                    case "min": result = numbers.Any() ? numbers.Min() : 0; break;
        //                }

        //                row[$"{agg} of {valField}"] = result;
        //            }
        //        }

        //        pivot.Rows.Add(row);
        //    }

        //    return pivot;
        //}

        public void InsertTableIntoPowerPoint(DataTable pivotTable, float fontSize, Pivot form, List<ConditionalRule> rules = null)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var pres = app.Presentations.Count > 0
                    ? app.ActivePresentation
                    : app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                var slide = pres.Slides.Add(pres.Slides.Count + 1,
                    Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

                // Insert as table
                int rows = pivotTable.Rows.Count + 1;
                int cols = pivotTable.Columns.Count;
                var tableShape = slide.Shapes.AddTable(rows, cols, 50, 50, 600, 20 * rows);
                var table = tableShape.Table;

                // Write headers
                for (int c = 0; c < cols; c++)
                {
                    table.Cell(1, c + 1).Shape.TextFrame.TextRange.Text = pivotTable.Columns[c].ColumnName;
                }

                // Write data
                for (int r = 0; r < pivotTable.Rows.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var cellText = pivotTable.Rows[r][c].ToString();
                        var cell = table.Cell(r + 2, c + 1);
                        cell.Shape.TextFrame.TextRange.Text = cellText;
                        var abc = double.TryParse(cellText, out var vall);
                        // Conditional formatting
                        if (rules != null)
                        {
                            foreach (var rule in rules)
                            {
                                if (pivotTable.Columns[c].ColumnName.Contains(rule.Field))
                                {
                                    if (Applies(vall, rule))
                                        cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(rule.Color);
                                }
                            }
                        }
                    }
                }

                // ✅ Save metadata (same as chart)
                PivotConfig config = form.GetConfig();
                string json = JsonConvert.SerializeObject(config);
                tableShape.Tags.Add("ChartMakerMeta", json);

                // Optional: visible in PowerPoint’s Alt Text UI
                tableShape.AlternativeText = "ChartMaker|" + config.DatasetPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting table: " + ex.Message);
            }
        }

        public void InsertPivotChartIntoPowerPoint(DataTable pivotTable, Office.XlChartType chartType, Pivot form,List<ConditionalRule> rules = null)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var pres = app.Presentations.Count > 0
                    ? app.ActivePresentation
                    : app.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

                var slide = pres.Slides.Add(pres.Slides.Count + 1,
                    Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank);

                var chartShape = slide.Shapes.AddChart(chartType, 50, 50, 600, 350);
                var chart = chartShape.Chart;

                var workbook = chart.ChartData.Workbook;
                var sheet = workbook.Worksheets[1];
                sheet.Cells.Clear();

                int rows = pivotTable.Rows.Count;
                int cols = pivotTable.Columns.Count;

                // --- Write headers ---
                for (int c = 0; c < cols; c++)
                    sheet.Cells[1, c + 1] = pivotTable.Columns[c].ColumnName;

                // --- Write data ---
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var value = pivotTable.Rows[r][c]?.ToString() ?? "";
                        if (double.TryParse(value, out var val))
                            sheet.Cells[r + 2, c + 1] = val;
                        else
                            sheet.Cells[r + 2, c + 1] = value;
                    }
                }


                // ✅ Build category array (first column)
                Excel.Range categoryRange = sheet.Range[sheet.Cells[2, 1], sheet.Cells[rows + 1, 1]];
                object[,] categories = categoryRange.Value2 as object[,];

                // ✅ Loop through numeric columns → each as a series
                for (int c = 2; c <= cols; c++)
                {
                    Excel.Range valuesRange = sheet.Range[sheet.Cells[2, c], sheet.Cells[rows + 1, c]];
                    object[,] values = valuesRange.Value2 as object[,];
                    string seriesName = pivotTable.Columns[c - 1].ColumnName;

                    if (values != null)
                    {
                        if (c - 1 <= chart.SeriesCollection().Count)
                        {
                            var series = (PowerPoint.Series)chart.SeriesCollection(c - 1);
                            series.Name = seriesName;
                            series.Values = values;
                            series.XValues = categories;
                        }
                        else
                        {
                            chart.SeriesCollection().NewSeries();
                            var series = (PowerPoint.Series)chart.SeriesCollection(chart.SeriesCollection().Count);
                            series.Name = seriesName;
                            series.Values = values;
                            series.XValues = categories;
                        }
                    }
                }

                // Style
                chart.HasLegend = true;
                chart.HasTitle = true;
                chart.ChartTitle.Text = "Pivot Chart";

                // Hide Excel so user doesn’t see embedded sheet
                sheet.Application.Visible = false;

                // --- Apply conditional formatting ---

                if (rules != null)
                {
                    for (int s = 1; s <= chart.SeriesCollection().Count; s++)
                    {
                        var series = chart.SeriesCollection(s);
                        string seriesName = series.Name;

                        for (int p = 1; p <= series.Points().Count; p++)
                        {
                            // pivotTable: first column = category, so data starts at column index 1
                            int dataColIndex = s;            // because s=1 → pivot col[1], s=2 → col[2] ...
                            int dataRowIndex = p - 1;        // chart point index maps directly to pivot row

                            if (dataRowIndex < pivotTable.Rows.Count &&
                                dataColIndex < pivotTable.Columns.Count)
                            {
                                double pointValue;
                                if (double.TryParse(pivotTable.Rows[dataRowIndex][dataColIndex].ToString(), out pointValue))
                                {
                                    foreach (var rule in rules)
                                    {
                                        if (seriesName.Contains(rule.Field) && Applies(pointValue, rule))
                                        {
                                            series.Points(p).Format.Fill.ForeColor.RGB =
                                                ColorTranslator.ToOle(rule.Color);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                // assume `chartShape` is the PowerPoint.Shape you created and `config` is the PivotConfig you used
                PivotConfig config = form.GetConfig();
                string json = JsonConvert.SerializeObject(config);
                chartShape.Tags.Add("ChartMakerMeta", json);

                // Optional: set alt text too (visible in PowerPoint UI)
                chartShape.AlternativeText = "ChartMaker|" + config.DatasetPath;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting chart: " + ex.Message);
            }
        }

        public bool Applies(double value, ConditionalRule rule)
        {
            switch (rule.Operator)
            {
                case ">": return value > rule.Threshold;
                case "<": return value < rule.Threshold;
                case ">=": return value >= rule.Threshold;
                case "<=": return value <= rule.Threshold;
                case "=": return value == rule.Threshold;
                default: return false;
            }
        }

        private async void btnEditWithChartMaker_ClickAsync(object sender, RibbonControlEventArgs e)
        {
            try    
            {
                var app = Globals.ThisAddIn.Application;

                //  Step 1: Ensure user selected a shape
                if ( app.ActiveWindow == null || app.ActiveWindow.Selection == null ||
                    app.ActiveWindow.Selection.Type == PowerPoint.PpSelectionType.ppSelectionNone)
                {
                    MessageBox.Show("Please select a chart or table shape to edit.");
                    return;
                }

                var shape = app.ActiveWindow.Selection.ShapeRange[1];

                //  Step 2: Get config metadata from shape
                string metaJson = shape.Tags["ChartMakerMeta"];
                if (string.IsNullOrEmpty(metaJson))
                {
                    MessageBox.Show("Selected shape is not a ChartMaker object.");
                    return;
                }

                //  Step 3: Deserialize old config
                var oldConfig = JsonConvert.DeserializeObject<PivotConfig>(metaJson);

                if (string.IsNullOrEmpty(oldConfig.DatasetPath) || !File.Exists(oldConfig.DatasetPath))
                {
                    MessageBox.Show("Dataset file not found. Please re-select the Excel file.");
                    return;
                }

                // Step 4: Reload dataset
                // var dt = new DatasetManager().LoadExcel(oldConfig.DatasetPath);

                var dt = await Task.Run(() => DatasetCache.GetOrLoad(oldConfig.DatasetPath));


                //  Step 5: Reapply calculated fields (if missing in DataTable)
                var ph = new PivotHelper();
                if (oldConfig.CalculatedFields != null)
                {
                    foreach (var cf in oldConfig.CalculatedFields)
                    {
                        if (!dt.Columns.Contains(cf.FieldName))
                            ph.AddCalculatedField(dt, cf.FieldName, cf.Formula);
                    }
                }

                //  Step 6: Open Pivot form with pre-filled config
                var form = new Pivot(dt, oldConfig.DatasetPath);
                form.LoadConfig(oldConfig);

                if (form.ShowDialog() == DialogResult.OK)
                {
                    // User updated config → collect latest
                    var newConfig = form.GetConfig();

                    string columnField = form.selectedYOYCompField;
                    if (columnField == "-- none --") columnField = null;

                    var newdt = form._data;
                    var dataTable = columnField == null ? dt : newdt;
                    // Build new pivot
                    var newPivot = CreatePivot(dt,
                        newConfig.RowField, newConfig.ValueFields,
                        newConfig.Aggregations, columnField, newConfig.Filters
                    );

                    var rule = form.GetConditionalRules();

                    //  Step 7: Update existing shape in-place
                    if (shape.Type == Office.MsoShapeType.msoChart)
                    {
                        UpdatePivotChartInPowerPoint(shape, newPivot, newConfig);
                    }
                    else if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        try { UpdatePivotTableInPowerPoint(shape, newPivot, newConfig);}
                        catch(Exception ex)
                        {
                            MessageBox.Show("Edit failed: " + ex.Message);
                        }
                        
                    }
                    try
                    {//  Step 8: Save new config back into shape tag
                        shape.Tags.Delete("ChartMakerMeta");
                        shape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(newConfig));
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Edit failed: " + ex.Message);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Edit failed: " + ex.Message);
            }
        }

        public PowerPoint.Shape UpdatePivotChartInPowerPoint(PowerPoint.Shape chartShape,DataTable pivotTable, PivotConfig config)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            var chart = chartShape.Chart;

            try
            {
                // Obtain workbook & sheet for the embedded chart
                // NOTE: chart.ChartData.Workbook is accessible; we try to use it without showing Excel UI.
                workbook = chart.ChartData.Workbook;
                sheet = (Excel.Worksheet)workbook.Worksheets[1];

                // Clear existing sheet content
                sheet.Cells.Clear();

                int rows = pivotTable.Rows.Count;
                int cols = pivotTable.Columns.Count;

                // Write headers
                for (int c = 0; c < cols; c++)
                    sheet.Cells[1, c + 1] = pivotTable.Columns[c].ColumnName;

                // Write data rows
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        sheet.Cells[r + 2, c + 1] = pivotTable.Rows[r][c] ?? "";
                    }
                }

                // Build Excel range and address string (sheetName!'A1:D5')
                Excel.Range fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows + 1, cols]];
                string addr = fullRange.Address[false, false, Excel.XlReferenceStyle.xlA1]; // e.g. A1:D5
                string sourceString = $"'{sheet.Name}'!{addr}";

                // QUICK TRY: try setting source directly (works in many cases)
                try
                {
                    chart.SetSourceData(sourceString, Excel.XlRowCol.xlColumns);
                }
                catch (COMException)
                {
                    // If PowerPoint refuses, activate the ChartData workbook (this will open Excel UI briefly),
                    // then try again and close workbook afterward.
                    chart.ChartData.Activate();
                    workbook = chart.ChartData.Workbook; // refresh reference after activate
                    sheet = (Excel.Worksheet)workbook.Worksheets[1];

                    // Recompute in case Activate changed anything
                    fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows + 1, cols]];
                    addr = fullRange.Address[false, false, Excel.XlReferenceStyle.xlA1];
                    sourceString = $"'{sheet.Name}'!{addr}";

                    chart.SetSourceData(sourceString, Excel.XlRowCol.xlColumns);
                }

                // Ensure chart plots columns as series
                chart.PlotBy = PowerPoint.XlRowCol.xlColumns;

                // OPTIONAL: rename series to match pivot column headers (skip category column at index 0)
                try
                {
                    for (int s = 2; s <= cols; s++)
                    {
                        int seriesIndex = s - 1; // first series is index 1
                        var series = chart.SeriesCollection(seriesIndex);
                        series.Name = pivotTable.Columns[s - 1].ColumnName;
                    }
                }
                catch { /* ignore if mismatch */ }

                // Refresh chart to commit changes
                chart.Refresh();

                // Apply conditional formatting if provided (map series -> pivot columns)
                if (config?.ConditionalRules != null && config.ConditionalRules.Any())
                {
                    for (int s = 1; s <= chart.SeriesCollection().Count; s++)
                    {
                        var series = chart.SeriesCollection(s);
                        string seriesName = series.Name ?? "";

                        for (int p = 1; p <= series.Points().Count; p++)
                        {
                            int dataRowIndex = p - 1;
                            int dataColIndex = s; // series s corresponds to pivotTable column s (col0 = category)

                            if (dataRowIndex < pivotTable.Rows.Count && dataColIndex < pivotTable.Columns.Count)
                            {
                                if (double.TryParse(pivotTable.Rows[dataRowIndex][dataColIndex]?.ToString(), out var pointVal))
                                {
                                    foreach (var rule in config.ConditionalRules)
                                    {
                                        if (seriesName.Contains(rule.Field) && Applies(pointVal, rule))
                                        {
                                            series.Points(p).Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(rule.Color);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    chart.Refresh();
              
                    chartShape.Tags.Delete("ChartMakerMeta");
                    chartShape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(config));
                    chartShape.AlternativeText = "ChartMaker|" + config.DatasetPath;
                    

                }

                // Close the embedded workbook (non-modal) and release COMs
                try
                {
                    workbook.Close(false);
                    return chartShape;
                }
                catch { /* ignore errors closing embedded workbook */  return chartShape; }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Error updating chart: " + ex.Message);
                return chartShape;
            }
            finally
            {
                // Safe COM clean-up
                if (sheet != null)
                {
                    try { Marshal.ReleaseComObject(sheet); } catch { }
                }
                if (workbook != null)
                {
                    try { Marshal.ReleaseComObject(workbook); } catch { }
                }
                // hint GC for final cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();
                
            }
        }
        public PowerPoint.Shape UpdatePivotTableInPowerPoint( PowerPoint.Shape tableShape,
     DataTable pivotTable,PivotConfig config)
        {
            PowerPoint.Shape newShape = tableShape;

            try
            {
                if (tableShape.HasTable != Office.MsoTriState.msoTrue) return tableShape;

                var slide = (PowerPoint.Slide)tableShape.Parent;
                float left = tableShape.Left, top = tableShape.Top, w = tableShape.Width, h = tableShape.Height;

                int rows = pivotTable.Rows.Count + 1; // header + data rows
                int cols = pivotTable.Columns.Count;

                var table = tableShape.Table;

                // 🔹 Rebuild table if dimension mismatch
                if (table.Rows.Count != rows || table.Columns.Count != cols)
                {
                    tableShape.Delete();
                    newShape = slide.Shapes.AddTable(rows, cols, left, top, w, h);
                    table = newShape.Table;
                }

                float fontSize = config?.Styles?.TableFontSize ?? 12;
                string fontName = config?.Styles?.TableFontName ?? "Calibri";
                int headerColor = config?.Styles?.TableHeaderColor
                                  ?? System.Drawing.Color.LightGray.ToArgb();

                // --- Write headers ---
                for (int c = 0; c < cols; c++)
                {
                    var headerCell = table.Cell(1, c + 1);
                    headerCell.Shape.TextFrame.TextRange.Text = pivotTable.Columns[c].ColumnName;

                    //headerCell.Shape.TextFrame.TextRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
                    //headerCell.Shape.TextFrame.TextRange.Font.Size = fontSize + 2; // headers slightly bigger
                    //headerCell.Shape.TextFrame.TextRange.Font.Name = fontName;
                    //headerCell.Shape.Fill.ForeColor.RGB = headerColor;
                }

                // --- Write values + conditional formatting ---
                for (int r = 0; r < pivotTable.Rows.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string text = pivotTable.Rows[r][c]?.ToString() ?? "";
                        var cell = table.Cell(r + 2, c + 1);
                        cell.Shape.TextFrame.TextRange.Text = text;

                        //cell.Shape.TextFrame.TextRange.Font.Size = fontSize;
                        //cell.Shape.TextFrame.TextRange.Font.Name = fontName;

                        if (config?.ConditionalRules != null && double.TryParse(text, out var val))
                        {
                            foreach (var rule in config.ConditionalRules)
                            {
                                if (pivotTable.Columns[c].ColumnName.Contains(rule.Field) && Applies(val, rule))
                                {
                                    cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(rule.Color);
                                }
                            }
                        }
                    }
                }

                // --- Auto-fit column widths ---
                float totalWidth = newShape.Width;
                float[] colWidths = new float[cols];
                float minColWidth = 40f;

                for (int c = 0; c < cols; c++)
                {
                    int maxLen = pivotTable.Columns[c].ColumnName.Length;
                    foreach (DataRow row in pivotTable.Rows)
                    {
                        int len = row[c]?.ToString()?.Length ?? 0;
                        if (len > maxLen) maxLen = len;
                    }
                    colWidths[c] = Math.Max(minColWidth, maxLen * 7f);
                }

                float scale = totalWidth / colWidths.Sum();
                for (int c = 0; c < cols; c++)
                {
                    table.Columns[c + 1].Width = colWidths[c] * scale;
                }

                // --- Update tag + alt text ---
                newShape.Tags.Delete("ChartMakerMeta");
                newShape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(config));
                newShape.AlternativeText = "ChartMaker|" + config.DatasetPath;

                return newShape;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating table: " + ex.Message);
                return newShape;
            }
        }

        private void btnDrillDown_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                var sel = app.ActiveWindow.Selection;

                if (sel == null)
                {
                    MessageBox.Show("No selection.");
                    return;
                }

                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    PowerPoint.Shape shape = null;

                    try
                    {
                        if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes && sel.ShapeRange != null && sel.ShapeRange.Count >= 1)
                            shape = sel.ShapeRange[1];
                        else if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
                            shape = sel.TextRange.Parent as PowerPoint.Shape ?? sel.ShapeRange?[1];
                    }
                    catch { }

                    if (shape == null)
                    {
                        MessageBox.Show("Please select a chart or table cell.");
                        return;
                    }

                    //  Get PivotConfig from metadata (common for both table and chart)
                    string metaJson = shape.Tags["ChartMakerMeta"];
                    if (string.IsNullOrEmpty(metaJson))
                    {
                        MessageBox.Show("This shape is not managed by ChartMaker.");
                        return;
                    }

                    var cfg = JsonConvert.DeserializeObject<PivotConfig>(metaJson);

                    // --- TABLE CASE ---
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                    {
                        string rowValue = GetSelectedTableFirstColumnValue(sel, shape);
                        if (string.IsNullOrEmpty(rowValue))
                        {
                            MessageBox.Show("Please select inside a table first column value to drill down.");
                            return;
                        }

                        //  Use dynamic RowField
                        ShowDrillDownWindow(rowValue, cfg.RowField);
                        return;
                    }

                    // --- CHART CASE ---
                    if (shape.Type == Office.MsoShapeType.msoChart)
                    {
                        var dt = new DatasetManager().LoadDataset(cfg.DatasetPath);

                        var distinctVals = dt.AsEnumerable()
                            .Select(r => r[cfg.RowField]?.ToString())
                            .Where(x => !string.IsNullOrEmpty(x))
                            .Distinct()
                            .OrderBy(x => x)
                            .ToList();

                        using (var drillForm = new DrillDownForm(cfg.RowField, distinctVals, dt))
                        {
                            drillForm.ShowDialog();
                        }
                        return;
                    }

                    MessageBox.Show("Drill-down works only on ChartMaker tables or charts.");
                }
                else
                {
                    MessageBox.Show("Please select a ChartMaker chart/table shape first.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Drill-down failed: " + ex.Message);
            }
        }

        /// <summary>
        /// Tries to determine the first-column (category) value for the currently selected table cell.
        /// Returns null if cannot determine.
        /// </summary>
        private string GetSelectedTableFirstColumnValue(PowerPoint.Selection sel, PowerPoint.Shape tableShape)
        {
            try
            {
                var tbl = tableShape.Table;
                if (tbl == null || tbl.Rows == null || tbl.Columns == null) return null;

                // 1) If user selected text inside a cell, use that text to find the row
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionText && sel.TextRange != null)
                {
                    string selectedText = sel.TextRange.Text?.Trim();
                    if (!string.IsNullOrEmpty(selectedText))
                    {
                        // Compare with first column (category column) values
                        for (int r = 2; r <= tbl.Rows.Count; r++) // assume row 1 is header
                        {
                            string cellVal = tbl.Cell(r, 1).Shape.TextFrame.TextRange.Text?.Trim();
                            if (string.Equals(cellVal, selectedText, StringComparison.OrdinalIgnoreCase))
                                return cellVal;
                        }

                        // If not found in first column, try to find the exact cell match anywhere in the row and then return the first-col
                        for (int r = 2; r <= tbl.Rows.Count; r++)
                        {
                            for (int c = 1; c <= tbl.Columns.Count; c++)
                            {
                                string cellVal = tbl.Cell(r, c).Shape.TextFrame.TextRange.Text?.Trim();
                                if (string.Equals(cellVal, selectedText, StringComparison.OrdinalIgnoreCase))
                                {
                                    return tbl.Cell(r, 1).Shape.TextFrame.TextRange.Text?.Trim();
                                }
                            }
                        }
                    }
                }

                // 2) If the whole table is selected (shape selected) or we couldn't map above,
                //    fallback: ask user to pick a row via InputBox (simple UX) or return first data row.
                if (tbl.Rows.Count >= 2)
                {
                    // Option A: prompt user for row index (commented)
                    string idxStr = Microsoft.VisualBasic.Interaction.InputBox("Enter row number to drill-down (1 = header):", "Select Row", "2");
                    if (int.TryParse(idxStr, out int userRow) && userRow >= 2 && userRow <= tbl.Rows.Count)
                        return tbl.Cell(userRow, 1).Shape.TextFrame.TextRange.Text.Trim();

                    // Option B: fallback to first data row
                    // return tbl.Cell(2, 1).Shape.TextFrame.TextRange.Text?.Trim();
                }
            }
            catch { /* ignore and return null below */ }

            return null;
        }

      
        private void ShowDrillDownWindow(string rowValue, string rowField)
        {
            // Active file ka dataset nikaalo
            string currentFile = ddlDatasets.SelectedItem.Tag.ToString(); // reuse your dataset dropdown
            if (string.IsNullOrEmpty(currentFile) || ddlDatasets.SelectedItem.Label == "-- select --")
            {
                MessageBox.Show("Dataset not found.");
                return;
            }

            var dt = new DatasetManager().LoadDataset(currentFile);
            var filtered = dt.AsEnumerable()
                .Where(r => r[rowField]?.ToString() == rowValue);

            if (!filtered.Any())
            {
                MessageBox.Show("No data found.");
                return;
            }

            DataTable result = dt.Clone();
            foreach (var r in filtered)
                result.ImportRow(r);

            using (var drillForm = new DrillDownForm(rowField, new List<string> { rowValue }, result))
            {
                drillForm.ShowDialog();
            }
        }

        private void btnRefreshAll_Click(object sender, Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs e)
    {
        try
        {
            var app = Globals.ThisAddIn.Application;
            var pres = app.ActivePresentation;
            if (pres == null)
            {
                MessageBox.Show("No active presentation to refresh.");
                return;
            }

            int updatedCount = 0;
            int skippedCount = 0;
            var errors = new List<string>();

            // Iterate slides and shapes (1-based COM collections)
            for (int s = 1; s <= pres.Slides.Count; s++)
            {
                var slide = pres.Slides[s];
                // iterate forward or backward is fine. we'll go forward.
                for (int i = 1; i <= slide.Shapes.Count; i++)
                {
                    var shape = slide.Shapes[i];
                    string metaJson = null;
                    try
                    {
                        metaJson = shape.Tags["ChartMakerMeta"];
                    }
                    catch { metaJson = null; }

                    if (string.IsNullOrWhiteSpace(metaJson))
                    {
                        // not a ChartMaker shape
                        skippedCount++;
                        continue;
                    }

                    PivotConfig cfg = null;
                    try
                    {
                        cfg = JsonConvert.DeserializeObject<PivotConfig>(metaJson);
                    }
                    catch (Exception ex)
                    {
                        skippedCount++;
                        errors.Add($"Invalid metadata on shape (slide {s}, shape {i}): {ex.Message}");
                        continue;
                    }

                    if (cfg == null || string.IsNullOrEmpty(cfg.DatasetPath))
                    {
                        skippedCount++;
                        continue;
                    }

                    if (!File.Exists(cfg.DatasetPath))
                    {
                        // missing dataset — skip but report
                        skippedCount++;
                        errors.Add($"Dataset not found: {cfg.DatasetPath} (slide {s}, shape {i})");
                        continue;
                    }

                    try
                    {
                        // 1) Load dataset (synchronous; COM must stay on UI thread)
                        var dt = new DatasetManager().LoadDataset(cfg.DatasetPath);
                           // var dt = new DatasetManager().LoadDataset(cfg.DatasetPath, forceReload: true);


                            // 2) Reapply calculated fields if any
                        if (cfg.CalculatedFields != null && cfg.CalculatedFields.Count > 0)
                        {
                            var ph = new PivotHelper();
                            foreach (var cf in cfg.CalculatedFields)
                            {
                                if (!dt.Columns.Contains(cf.FieldName))
                                    ph.AddCalculatedField(dt, cf.FieldName, cf.Formula);
                            }
                        }

                        // 3) Build pivot (uses your existing CreatePivot method)
                        // Assumes CreatePivot signature: CreatePivot(dt, rowField, valueFields, aggFuncs, columnField, filters)
                        

                            if(cfg.RowField != null)
                            {
                                var pivot = Globals.Ribbons.Ribbon1.CreatePivot(
                                  dt, cfg.RowField,
                                  cfg.ValueFields,
                                  cfg.Aggregations,
                                  null,
                                   cfg.Filters
                                 );
                                // 4) Update shape in-place
                                if (shape.Type == MsoShapeType.msoChart)
                                {
                                    shape = UpdatePivotChartInPowerPoint(shape, pivot, cfg);
                                    updatedCount++;
                                }
                                else if (shape.HasTable == MsoTriState.msoTrue)
                                {
                                    shape = UpdatePivotTableInPowerPoint(shape, pivot, cfg);
                                   
                                    updatedCount++;
                                }
                                else
                                {
                                    // unknown shape type; skip
                                    skippedCount++;
                                }
                            }
                       

                      
                    }
                    catch (Exception ex)
                    {
                        skippedCount++;
                        errors.Add($"Error updating shape on slide {s}, shape {i}: {ex.Message}");
                        // don't rethrow—we continue to attempt other shapes
                    }
                }
            }


                //if (shape.HasTable == MsoTriState.msoTrue)
                {
                    string filePath = ddlDatasets.SelectedItem.Tag.ToString();

                    if (filePath == "select")
                    {
                        MessageBox.Show("Select file");
                        return;
                    }
                        
                    InsertTableFromDataset(filePath, "Table");

                }

                var msg = $"Refresh complete.\nUpdated: {updatedCount}\nSkipped: {skippedCount}";
            if (errors.Any()) msg += $"\n\nErrors:\n- {string.Join("\n- ", errors.Take(10))}" + (errors.Count > 10 ? $"\n...({errors.Count - 10} more)" : "");
            MessageBox.Show(msg, "Refresh All");
        }
        catch (Exception ex)
        {
            MessageBox.Show("Refresh All failed: " + ex.Message);
        }
    }

        private void CreateRepeatViews(DataTable dt, Pivot form, string chartType, List<ConditionalRule> rules = null)
        {

           PivotConfig config = form.GetConfig();

            if (string.IsNullOrEmpty(config.RepeatBy))
            {
                // Normal pivot
                if (chartType != "0")
                    InsertPivotChartIntoPowerPoint(dt, form.SelectedChartType,form, rules);
                else
                    InsertTableIntoPowerPoint(dt, 25, form, rules);
                return;
            }

            // Distinct values for repeat column
            var distinctVals = dt.AsEnumerable()
                .Select(r => r[config.RepeatBy]?.ToString())
                .Where(v => !string.IsNullOrEmpty(v))
                .Distinct()
                .ToList();

            foreach (var val in distinctVals)
            {
                var filtered = dt.AsEnumerable()
                    .Where(r => r[config.RepeatBy]?.ToString() == val)
                    .CopyToDataTable();

                // Deep copy config for each slide
                var localConfig = JsonConvert.DeserializeObject<PivotConfig>(
                    JsonConvert.SerializeObject(config));
                localConfig.Title = $"{config.Title} - {config.RepeatBy} = {val}";

                if (chartType != "0")
                    InsertPivotChartIntoPowerPoint(dt, form.SelectedChartType, form, rules);
                else
                    InsertTableIntoPowerPoint(dt, 25, form, rules);
            }
        }


    }
}
