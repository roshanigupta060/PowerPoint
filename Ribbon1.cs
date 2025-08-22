using Google.Cloud.Storage.V1;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace PptExcelSync
{
    public partial class Ribbon1
    {
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

            string filePath = ddlDatasets.SelectedItem.Tag.ToString();

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

    }
}
