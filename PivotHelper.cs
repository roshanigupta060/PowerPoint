using System;
using System.Collections.Generic;
using System.Linq;
using NCalc;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace PptExcelSync
{
    public class PivotHelper
    {
        public void AddCalculatedField(DataTable dt, string fieldName, string formula)
        {
            if (dt.Columns.Contains(fieldName))
                throw new Exception($"Field '{fieldName}' already exists.");

            // Add the new column
            dt.Columns.Add(fieldName, typeof(double));

            foreach (DataRow row in dt.Rows)
            {
                var expr = new Expression(formula);

                // Pass all columns as parameters
                foreach (DataColumn col in dt.Columns)
                {
                    if (col.ColumnName != fieldName) // avoid recursion
                    {
                        double num;
                        if (double.TryParse(row[col].ToString(), out num))
                            expr.Parameters[col.ColumnName] = num;
                        else
                            expr.Parameters[col.ColumnName] = row[col].ToString();
                    }
                }

                // Evaluate safely
                object result = expr.Evaluate();
                if (result is double d)
                    row[fieldName] = d;
                else if (double.TryParse(result?.ToString(), out double parsed))
                    row[fieldName] = parsed;
                else
                    row[fieldName] = 0;
            }
        }
    }

}
