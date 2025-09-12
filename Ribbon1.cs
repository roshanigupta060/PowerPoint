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
using System.Web;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Forms.VisualStyles;
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
            var form = new Pivot(dt, filePath);
        
            if (form.ShowDialog() == DialogResult.OK)
            {                
                if (form.SelectedValueFields.Count() > 2)
                {
                    MessageBox.Show("Please select only two values.");
                    return;
                }

                var filters = form.GetFilters();
                var pivot = CreatePivot(dt, form.SelectedRowField, form.SelectedValueFields, form.SelectedAggregations, filters);

                // ⬇️ Get rules from the form
                var rules = form.GetConditionalRules();

                // Insert pivot into PowerPoint
                string chartType =  form.SelectedChartType.ToString();

                if(chartType != "0")
                  InsertPivotChartIntoPowerPoint(pivot, form.SelectedChartType, rules);
                else
                  InsertTableIntoPowerPoint(pivot, 25, rules);
            }
        }

        private Dictionary<string, string> ConvertFilters(List<FilterRule> savedFilters)
        {
            var dict = new Dictionary<string, string>();
            if (savedFilters != null)
            {
                foreach (var f in savedFilters)
                {
                    if (!string.IsNullOrWhiteSpace(f.Column) && !string.IsNullOrWhiteSpace(f.Value))
                    {
                        dict[f.Column] = f.Value;
                    }
                }
            }
            return dict;
        }


        public DataTable CreatePivot(DataTable dt,string rowField, List<string> valueFields, List<string> aggFuncs,
              Dictionary<string, string> filters = null)
        {
            // --- Step 1: Apply filters if provided ---
            IEnumerable<DataRow> query = dt.AsEnumerable();
            if (filters != null && filters.Any())
            {
                foreach (var f in filters)
                {
                    string col = f.Key;
                    string val = f.Value;

                    query = query.Where(r => r[col]?.ToString() == val);
                }
            }

            // --- Step 2: Group by row field ---
            var grouped = query.GroupBy(r => r[rowField].ToString());

            // --- Step 3: Build output table ---
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

            // --- Step 4: Fill data ---
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

        public void InsertTableIntoPowerPoint(DataTable pivotTable, float fontSize, List<ConditionalRule> rules = null)
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
                var table = slide.Shapes.AddTable(rows, cols, 50, 50, 600, 20 * rows).Table;

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
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error inserting table: " + ex.Message);
            }
        }

        public void InsertPivotChartIntoPowerPoint(DataTable pivotTable, Office.XlChartType chartType,List<ConditionalRule> rules = null)
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

                // assume `chartShape` is the PowerPoint.Shape you created and `config` is the PivotConfig you used
               PivotConfig config = new PivotConfig();
                string json = JsonConvert.SerializeObject(config);
                chartShape.Tags.Add("ChartMakerMeta", json);

                // Optional: set alt text too (visible in PowerPoint UI)
                chartShape.AlternativeText = "ChartMaker|" + config.DatasetPath;

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

        private void btnEditWithChartMaker_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                var app = Globals.ThisAddIn.Application;
                if (app.ActiveWindow == null || app.ActiveWindow.Selection == null ||
                    app.ActiveWindow.Selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    MessageBox.Show("Please select a chart or table shape to edit.");
                    return;
                }

                var shape = app.ActiveWindow.Selection.ShapeRange[1];
                string metaJson = shape.Tags["ChartMakerMeta"];
                if (string.IsNullOrEmpty(metaJson))
                {
                    MessageBox.Show("Selected shape is not a ChartMaker object.");
                    return;
                }

                string filePath = ddlDatasets.SelectedItem.Tag.ToString();
                var dt = new DatasetManager().LoadExcel(filePath);
                // Show Pivot dialog
                var form = new Pivot(dt, filePath);
                var newConfig = form.GetConfig();

                //var config = JsonConvert.DeserializeObject<PivotConfig>(metaJson);

                if (string.IsNullOrEmpty(newConfig.DatasetPath) || !File.Exists(newConfig.DatasetPath))
                {
                    MessageBox.Show("Dataset file not found. Please re-select the Excel file.");
                    return;
                }



                // Load dataset and apply any calculated fields (DatasetManager should handle metadata)
                //var dt = new DatasetManager().LoadExcel(config.DatasetPath);

                // Reapply any calculated fields included directly in the config (optional):
                var ph = new PivotHelper();
                if (newConfig.CalculatedFields != null)
                {
                    foreach (var cf in newConfig.CalculatedFields)
                    {
                        if (!dt.Columns.Contains(cf.FieldName))
                            ph.AddCalculatedField(dt, cf.FieldName, cf.Formula);
                    }
                }

                // Open Pivot form pre-filled
    
                    form.LoadConfig(newConfig); // method you will implement in Pivot form
                    if (form.ShowDialog() == DialogResult.OK)
                    {
                        // user updated config in UI; collect the new config and rebuild
                       // var newConfig = form.GetConfig(); // method returns PivotConfig
                        var newPivot = CreatePivot(dt, newConfig.RowField, newConfig.ValueFields, newConfig.Aggregations, ConvertFilters(newConfig.Filters));

                        // Update the existing shape in-place
                        if (shape.Type == Office.MsoShapeType.msoChart)
                            UpdatePivotChartInPowerPoint(shape, newPivot, newConfig);
                        else if (shape.HasTable == Office.MsoTriState.msoTrue)
                            UpdatePivotTableInPowerPoint(shape, newPivot, newConfig);

                        // Update the shape tag with the new config
                        shape.Tags.Delete("ChartMakerMeta");
                        shape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(newConfig));
                    }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Edit failed: " + ex.Message);
            }
        }

        public void UpdatePivotChartInPowerPoint(PowerPoint.Shape chartShape, DataTable pivotTable, PivotConfig config)
        {
            Excel.Workbook workbook = null;
            Excel.Worksheet sheet = null;
            try
            {
                var chart = chartShape.Chart;
                // Access embedded workbook & worksheet
                workbook = chart.ChartData.Workbook;
                sheet = (Excel.Worksheet)workbook.Worksheets[1];
                sheet.Cells.Clear();

                int rows = pivotTable.Rows.Count;
                int cols = pivotTable.Columns.Count;

                // Write headers
                for (int c = 0; c < cols; c++)
                    sheet.Cells[1, c + 1] = pivotTable.Columns[c].ColumnName;

                // Write data
                for (int r = 0; r < rows; r++)
                    for (int c = 0; c < cols; c++)
                        sheet.Cells[r + 2, c + 1] = pivotTable.Rows[r][c];

                // Build full range and set source
                Excel.Range fullRange = sheet.Range[sheet.Cells[1, 1], sheet.Cells[rows + 1, cols]];
                // Use address approach to be safe
                string addr = sheet.Name + "!" + fullRange.Address[false, false];
                chart.SetSourceData(addr, Excel.XlRowCol.xlColumns);

                chart.PlotBy = PowerPoint.XlRowCol.xlColumns;

                // rename series to pivot headers (skip category col at index 0)
                for (int s = 2; s <= cols; s++)
                {
                    int seriesIndex = s - 1; // series index is 1-based and usually corresponds
                    try
                    {
                        var series = chart.SeriesCollection(seriesIndex);
                        series.Name = pivotTable.Columns[s - 1].ColumnName;
                    }
                    catch { /* ignore if series doesn't exist yet */ }
                }

                // Refresh chart
                // don't call chart.ChartData.Activate(); that shows Excel UI
                workbook.Application.CalculateFull();
                chart.Refresh();

                // Apply conditional formatting rules per point (use pivotTable to map)
                if (config?.ConditionalRules != null && config.ConditionalRules.Any())
                {
                    for (int s = 1; s <= chart.SeriesCollection().Count; s++)
                    {
                        var series = chart.SeriesCollection(s);
                        string seriesName = series.Name;
                        for (int p = 1; p <= series.Points().Count; p++)
                        {
                            int dataRowIndex = p - 1;
                            int dataColIndex = s; // series s maps to pivotTable column s (assuming col0=category)
                            if (dataRowIndex < pivotTable.Rows.Count && dataColIndex < pivotTable.Columns.Count)
                            {
                                if (double.TryParse(pivotTable.Rows[dataRowIndex][dataColIndex].ToString(), out var pointVal))
                                {
                                    foreach (var rule in config.ConditionalRules)
                                    {
                                        if (seriesName.Contains(rule.Field) && Applies(pointVal,rule))
                                        {
                                            series.Points(p).Format.Fill.ForeColor.RGB = ColorTranslator.ToOle(rule.Color);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            finally
            {
                // Close/hide workbook and release COM
                if (workbook != null)
                {
                    try { workbook.Close(false); } catch { }
                    Marshal.ReleaseComObject(sheet);
                    Marshal.ReleaseComObject(workbook);
                }
            }
        }

        public void UpdatePivotTableInPowerPoint(PowerPoint.Shape tableShape, DataTable pivotTable, PivotConfig config)
        {
            try
            {
                if (tableShape.HasTable != Office.MsoTriState.msoTrue) return;
                var table = tableShape.Table;
                int rows = pivotTable.Rows.Count + 1;
                int cols = pivotTable.Columns.Count;

                // Optionally, rebuild if table size mismatches: delete & recreate, or resize shape.Table if possible.
                // Here we'll assume size matches or recreate if necessary:
                if (table.Rows.Count != rows || table.Columns.Count != cols)
                {
                    // remove existing and add new table (simple approach)
                    var slide = tableShape.Parent as PowerPoint.Slide;
                    var left = tableShape.Left; var top = tableShape.Top; var w = tableShape.Width; var h = tableShape.Height;
                    tableShape.Delete();
                    var newShape = slide.Shapes.AddTable(rows, cols, left, top, w, h);
                    table = newShape.Table;
                    tableShape = newShape;
                }

                // Write headers
                for (int c = 0; c < cols; c++)
                    table.Cell(1, c + 1).Shape.TextFrame.TextRange.Text = pivotTable.Columns[c].ColumnName;

                // Write values and conditional formatting
                for (int r = 0; r < pivotTable.Rows.Count; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        var text = pivotTable.Rows[r][c]?.ToString() ?? "";
                        var cell = table.Cell(r + 2, c + 1);
                        cell.Shape.TextFrame.TextRange.Text = text;

                        // conditional format
                        if (config?.ConditionalRules != null)
                        {
                            foreach (var rule in config.ConditionalRules)
                            {
                                if (pivotTable.Columns[c].ColumnName.Contains(rule.Field)
                                    && double.TryParse(text, out var v) && Applies(v, rule))
                                {
                                    cell.Shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(rule.Color);
                                }
                            }
                        }
                    }
                }

                // Update tag on the shape
                tableShape.Tags.Delete("ChartMakerMeta");
                tableShape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(config));
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error updating table: " + ex.Message);
            }
        }

    }
}
