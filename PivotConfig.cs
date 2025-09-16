using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptExcelSync
{
    public class PivotConfig
    {
        public string DatasetPath { get; set; }              // path to the Excel file used
        public string RowField { get; set; }                 // group-by column
        public List<string> ValueFields { get; set; } = new List<string>();
        public List<string> Aggregations { get; set; } = new List<string>();
        public string ChartTypeName { get; set; }            // string representation (e.g., "xlColumnClustered")
        public List<CalculatedFieldInfo> CalculatedFields { get; set; } = new List<CalculatedFieldInfo>();
        public Dictionary<string, string>  Filters { get; set; } = new Dictionary<string, string>();
        public List<ConditionalRule> ConditionalRules { get; set; } = new List<ConditionalRule>();
        public StyleConfig Styles { get; set; } = new StyleConfig();
    }

    public class StyleConfig
    {
        // Table styles
        public float TableFontSize { get; set; } = 12;
        public string TableFontName { get; set; } = "Calibri";
        public int TableHeaderColor { get; set; } = System.Drawing.Color.LightGray.ToArgb();

        // Chart styles
        public bool ShowLegend { get; set; } = true;
        public bool ShowTitle { get; set; } = true;
        public string TitleText { get; set; } = "Pivot Chart";
        public int[] SeriesColors { get; set; } // optional: custom series colors
    }
}
