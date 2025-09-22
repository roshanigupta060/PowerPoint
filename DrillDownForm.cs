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
    public partial class DrillDownForm : Form
    {
        private string rowField;
        private DataTable dataset;

        public DrillDownForm(string field, List<string> values, DataTable dt)
        {
            InitializeComponent();
            rowField = field;
            dataset = dt;

            cmbValues.DataSource = values;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            string selectedVal = cmbValues.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(selectedVal)) return;

            var filtered = dataset.AsEnumerable()
                .Where(r => r[rowField]?.ToString() == selectedVal);

            if (!filtered.Any())
            {
                MessageBox.Show("No data found.");
                return;
            }

            DataTable result = dataset.Clone();
            foreach (var r in filtered)
                result.ImportRow(r);
            grid.DataSource = result;

        }
    }

}
