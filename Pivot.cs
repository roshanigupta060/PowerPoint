using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PptExcelSync
{
    public partial class Pivot : Form
    {
        private DataTable _data;

        public string SelectedRowField => cmbRowField.SelectedItem?.ToString();
        public string SelectedValueField => cmbValueField.SelectedItem?.ToString();
        public string SelectedAggregation => cmbAggregation.SelectedItem?.ToString();

        public Pivot(DataTable data)
        {
            InitializeComponent();
            _data = data;

            foreach (DataColumn col in data.Columns)
            {
                // Always add to Row dropdown
                cmbRowField.Items.Add(col.ColumnName);

                // Add to Value dropdown as well
                cmbValueField.Items.Add(col.ColumnName);
            }

            cmbAggregation.Items.AddRange(new string[] { "Sum", "Average", "Count" });
            cmbAggregation.SelectedIndex = 0;
        }

        private void btnGenerate_Click(object sender, EventArgs e)
        {

            if (SelectedRowField == null || SelectedValueField == null)
            {
                MessageBox.Show("Please select Row field and Value field.");
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
