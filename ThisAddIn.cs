using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptExcelSync
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.Application.WindowBeforeDoubleClick += Application_WindowBeforeRightClick;
            this.Application.WindowBeforeRightClick += Application_WindowBeforeRightClick;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection Sel)
        {
            try
            {
                if (Sel.Type != PowerPoint.PpSelectionType.ppSelectionNone)
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

        private void Application_WindowBeforeRightClick(PowerPoint.Selection Sel, ref bool Cancel)
        {

   

            try
            {
               // if (Sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes)
                {
                    var shape = Sel.ShapeRange[1];
                    bool hasMeta = !string.IsNullOrEmpty(shape.Tags["ChartMakerMeta"]);

                    if (hasMeta)
                    {
                        var commandBars = Globals.ThisAddIn.Application.CommandBars;

                        // Try multiple menus
                        Office.CommandBar contextMenu = null;
                        if (shape.HasTable == Office.MsoTriState.msoTrue)
                            contextMenu = commandBars["Table Context Menu"];
                        else
                            contextMenu = commandBars["Shape"];

                        if (contextMenu == null) return;

                        // Remove old copy
                        foreach (Office.CommandBarControl ctrl in contextMenu.Controls)
                        {
                            if (ctrl.Caption == "Edit with ChartMaker")
                            {
                                ctrl.Delete();
                                break;
                            }
                        }

                        // Add new
                        var button = (Office.CommandBarButton)contextMenu.Controls.Add(
                            Office.MsoControlType.msoControlButton,
                            Temporary: true);

                        button.Caption = "Edit with ChartMaker";
                        button.Click += (Office.CommandBarButton Ctrl, ref bool CancelDefault) =>
                        {
                            EditSelectedShapeWithChartMaker(shape);
                        };
                    }
                }
            }
            catch { }
        }

        private void EditSelectedShapeWithChartMaker(PowerPoint.Shape shape)
        {
            try
            {
                string metaJson = shape.Tags["ChartMakerMeta"];
                if (string.IsNullOrEmpty(metaJson))
                {
                    MessageBox.Show("This shape is not managed by ChartMaker.");
                    return;
                }

                // Reuse your existing btnEditWithChartMaker_Click logic:
                var oldConfig = JsonConvert.DeserializeObject<PivotConfig>(metaJson);
                var dt = new DatasetManager().LoadExcel(oldConfig.DatasetPath);

                var form = new Pivot(dt, oldConfig.DatasetPath);
                form.LoadConfig(oldConfig);

                if (form.ShowDialog() == DialogResult.OK)
                {
                    var newConfig = form.GetConfig();
                    var newPivot = Globals.Ribbons.Ribbon1.CreatePivot(dt, newConfig.RowField, newConfig.ValueFields,
                        newConfig.Aggregations, newConfig.Filters);

                    if (shape.Type == Office.MsoShapeType.msoChart)
                        Globals.Ribbons.Ribbon1.UpdatePivotChartInPowerPoint(shape, newPivot, newConfig);
                    else if (shape.HasTable == Office.MsoTriState.msoTrue)
                        Globals.Ribbons.Ribbon1.UpdatePivotTableInPowerPoint(shape, newPivot, newConfig);

                    shape.Tags.Delete("ChartMakerMeta");
                    shape.Tags.Add("ChartMakerMeta", JsonConvert.SerializeObject(newConfig));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Edit failed: " + ex.Message);
            }
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
