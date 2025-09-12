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
        public List<FilterRule> Filters { get; set; } = new List<FilterRule>();
        public List<ConditionalRule> ConditionalRules { get; set; } = new List<ConditionalRule>();
    }

    // Reuse the classes you already have (examples):
    //public class CalculatedFieldInfo { public string FieldName { get; set; } public string Formula { get; set; } }
    public class FilterRule { public string Column { get; set; } public string Value { get; set; } }
    //public class ConditionalRule { public string Field { get; set; } public string Operator { get; set; } public double Threshold { get; set; } public Color Color { get; set; } public bool Applies(double v) { switch (Operator) { case ">": return v > Threshold; case "<": return v < Threshold; case ">=": return v >= Threshold; case "<=": return v <= Threshold; case "=": return v == Threshold; default: return false; } } }

}
