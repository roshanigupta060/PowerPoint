using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Office = Microsoft.Office.Core;

namespace PptExcelSync
{
    public partial class Pivot : Form
    {
        public string FilePath = string.Empty;
        private Color selectedColor = Color.Empty;
        public string SelectedRowField => cmbRowField.SelectedItem?.ToString();
        public string SelectedChartTypeField => cmbChartType.SelectedItem?.ToString();
        public List<string> SelectedValueFields =>
            clbValueFields.CheckedItems.Cast<string>().ToList();
        public List<string> SelectedAggregations =>
            clbAggregations.CheckedItems.Cast<string>().ToList();
        public Office.XlChartType SelectedChartType
        {
            get
            {
                switch (cmbChartType.SelectedItem?.ToString())
                {
                    case "Column": return Office.XlChartType.xlColumnClustered;
                    case "Bar": return Office.XlChartType.xlBarClustered;
                    case "Line": return Office.XlChartType.xlLine;
                    case "Pie": return Office.XlChartType.xlPie;
                    default: return 0;
                }
            }
        }

        DataTable _data;
        public Pivot(DataTable data, string filePath)
        {
            _data = data;
            FilePath = filePath;
            InitializeComponent();
            InitializeValueContextMenu();
            PopulateDropdowns(_data);
        }

        private void btnPickColor_Click(object sender, EventArgs e)
        {
            if (colorDialog1.ShowDialog() == DialogResult.OK)
            {
                selectedColor = colorDialog1.Color;
                btnPickColor.BackColor = selectedColor;
            }
        }
        private void btnAddRule_Click(object sender, EventArgs e)
        {
            if (cmbField.SelectedItem == null || cmbOperator.SelectedItem == null || string.IsNullOrWhiteSpace(txtThreshold.Text))
            {
                MessageBox.Show("Please select field, operator, and enter a threshold.");
                return;
            }

            string field = cmbField.SelectedItem.ToString();
            string op = cmbOperator.SelectedItem.ToString();
            string threshold = txtThreshold.Text;
            string color = ColorTranslator.ToHtml(selectedColor);

            string rule = $"{field} {op} {threshold} => {color}";
            lstRules.Items.Add(rule);
        }
        private void btnDeleteRule_Click(object sender, EventArgs e)
        {
            if (lstRules.SelectedItem != null)
                lstRules.Items.Remove(lstRules.SelectedItem);
            else
            {
                MessageBox.Show("Select rule to delete.");
                return;
            }
        }
        private void InitializeValueContextMenu()
        {
            valueContextMenu = new ContextMenuStrip();
            var deleteItem = new ToolStripMenuItem("Delete Calculated Field");
            deleteItem.Click += DeleteCalculatedField_Click;
            valueContextMenu.Items.Add(deleteItem);

            // Attach menu to Value list
            clbValueFields.ContextMenuStrip = valueContextMenu; // if ComboBox
                                                               // OR lstValueField.ContextMenuStrip = valueContextMenu; // if ListBox
        }

        private void DeleteCalculatedField_Click(object sender, EventArgs e)
        {
            if (clbValueFields.SelectedItem == null)
            {
                MessageBox.Show("Please select a field to delete.");
                return;
            }

            string fieldName = clbValueFields.SelectedItem.ToString();

            // Load metadata
            var metadata = DatasetMetadata.Load(FilePath);

            // Check if it's a calculated field
            var calcField = metadata.CalculatedFields
                .FirstOrDefault(f => f.FieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase));

            if (calcField == null)
            {
                MessageBox.Show("This field is not a calculated field and cannot be deleted.");
                return;
            }

            // --- Delete ---
            metadata.CalculatedFields.Remove(calcField);
            metadata.Save(FilePath);

            if (_data.Columns.Contains(fieldName))
                _data.Columns.Remove(fieldName);

            // Refresh dropdowns
            PopulateDropdowns(_data);

            MessageBox.Show($"Calculated field '{fieldName}' deleted successfully.");
        }


        private void PopulateDropdowns(DataTable data)
        {
            cmbRowField.Items.Clear();
            clbValueFields.Items.Clear();
            cmbField.Items.Clear();

            foreach (DataColumn col in data.Columns)
            {
                bool isNumeric = true;

                foreach (DataRow row in data.Rows.Cast<DataRow>().Take(2)) // check first 5 rows
                {
                    var val = row[col.ColumnName]?.ToString();
                    if (string.IsNullOrWhiteSpace(val)) continue;

                    if (!double.TryParse(val, out _))
                    {
                        isNumeric = false;
                        break;
                    }
                }

                if (isNumeric)
                {
                    clbValueFields.Items.Add(col.ColumnName);
                    cmbField.Items.Add(col.ColumnName);
                }
                else
                {
                    cmbRowField.Items.Add(col.ColumnName);
                }
            }

            // Default selections
            if (cmbRowField.Items.Count > 0) cmbRowField.SelectedIndex = 0;
            if (clbValueFields.Items.Count > 0) clbValueFields.SelectedIndex = 0;

            // Aggregations
            clbAggregations.Items.Clear();
            clbAggregations.Items.AddRange(new string[] { "Sum", "Average", "Count", "Max", "Min" });
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {
            if (SelectedRowField == null || SelectedChartTypeField == null || SelectedValueFields.Count == 0 || SelectedAggregations.Count == 0)
            {
                MessageBox.Show("Please select a Row field, at least one Value field, and at least one Aggregation.");
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void cmbChartType_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btnAddField_Click_1(object sender, EventArgs e)
        {
            try
            {
                string fieldName = txtFieldName.Text.Trim();
                string formula = txtFormula.Text.Trim();
                if (string.IsNullOrWhiteSpace(fieldName) || string.IsNullOrWhiteSpace(formula))
                {
                    MessageBox.Show("Please enter both field name and formula.");
                    return;
                }
                string filePath = FilePath;

                // Add calculated field
                var calcHelper = new PivotHelper();
                calcHelper.AddCalculatedField(_data, fieldName, formula); // _data is your DataTable

                // Save into metadata

                
                var metadata = DatasetMetadata.Load(filePath);

                // Avoid duplicates
                if (!metadata.CalculatedFields.Any(f => f.FieldName.Equals(fieldName, StringComparison.OrdinalIgnoreCase)))
                {
                    metadata.CalculatedFields.Add(new CalculatedFieldInfo
                    {
                        FieldName = fieldName,
                        Formula = formula
                    });
                    metadata.Save(filePath);
                }

                // Refresh dropdowns so new field appears in Values
                PopulateDropdowns(_data);

                MessageBox.Show($"Calculated field '{fieldName}' added successfully!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error adding calculated field: " + ex.Message);
            }
        }

        private void clbValueFields_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        public List<ConditionalRule> GetConditionalRules()
        {
            var rules = new List<ConditionalRule>();

            foreach (string item in lstRules.Items)
            {
                // Example format: "Quantity > 50 => #FF0000"
                var parts = item.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                if (parts.Length < 5) continue;

                string field = parts[0];
                string op = parts[1];
                double threshold = double.TryParse(parts[2], out var t) ? t : 0;
                string colorHex = parts[4];
                var color = ColorTranslator.FromHtml(colorHex);

                rules.Add(new ConditionalRule
                {
                    Field = field,
                    Operator = op,
                    Threshold = threshold,
                    Color = color
                });
            }

            return rules;
        }

    }
}
