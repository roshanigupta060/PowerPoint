using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PptExcelSync
{
    public partial class Pivot : Form
    {
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
        public Pivot(DataTable data)
        {
            InitializeComponent();

            PopulateDropdowns(data);
        }

        private void PopulateDropdowns(DataTable data)
        {
            // Row fields string columns
            cmbRowField.Items.Clear();
            //foreach (DataColumn col in data.Columns)
            //{
            //    cmbRowField.Items.Add(col.ColumnName);
            //}

            // Value fields = detect numeric columns by checking first few rows
            //cmbValueField.Items.Clear();
            clbValueFields.Items.Clear();
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
    }
}
