using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PptExcelSync
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shape = Sel.ShapeRange[1];
                    bool hasMeta = !string.IsNullOrEmpty(shape.Tags["ChartMakerMeta"]);
                    Globals.Ribbons.Ribbon1.btnEditWithChartMaker.Enabled = hasMeta; // if your ribbon property is accessible
                }
                else
                {
                    Globals.Ribbons.Ribbon1.btnEditWithChartMaker.Enabled = false;
                }
            }
            catch { }
        }

        #region VSTO generated code

            /// <summary>
            /// Required method for Designer support - do not modify
            /// the contents of this method with the code editor.
            /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
