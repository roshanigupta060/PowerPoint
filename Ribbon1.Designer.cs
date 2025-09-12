namespace PptExcelSync
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl2 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl3 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl4 = this.Factory.CreateRibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl5 = this.Factory.CreateRibbonDropDownItem();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnUploadExcel = this.Factory.CreateRibbonButton();
            this.ddlDatasets = this.Factory.CreateRibbonDropDown();
            this.ddlChartType = this.Factory.CreateRibbonDropDown();
            this.btnPivotView = this.Factory.CreateRibbonButton();
            this.btnInsertChart_Click = this.Factory.CreateRibbonButton();
            this.btnInsertTable_Click = this.Factory.CreateRibbonButton();
            this.btnCreateChart_Click = this.Factory.CreateRibbonButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.btnEditWithChartMaker = this.Factory.CreateRibbonButton();
            this.group1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnUploadExcel);
            this.group1.Items.Add(this.ddlDatasets);
            this.group1.Items.Add(this.ddlChartType);
            this.group1.Items.Add(this.btnPivotView);
            this.group1.Items.Add(this.btnInsertChart_Click);
            this.group1.Items.Add(this.btnInsertTable_Click);
            this.group1.Items.Add(this.btnCreateChart_Click);
            this.group1.Items.Add(this.btnEditWithChartMaker);
            this.group1.Name = "group1";
            // 
            // btnUploadExcel
            // 
            this.btnUploadExcel.Image = global::PptExcelSync.Properties.Resources.upload;
            this.btnUploadExcel.Label = "Upload Excel";
            this.btnUploadExcel.Name = "btnUploadExcel";
            this.btnUploadExcel.ShowImage = true;
            this.btnUploadExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUploadExcel_Click);
            // 
            // ddlDatasets
            // 
            this.ddlDatasets.Image = global::PptExcelSync.Properties.Resources.datasets;
            this.ddlDatasets.Label = "Select File";
            this.ddlDatasets.Name = "ddlDatasets";
            this.ddlDatasets.ShowImage = true;
            this.ddlDatasets.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddlDatasets_SelectionChanged);
            // 
            // ddlChartType
            // 
            ribbonDropDownItemImpl1.Label = "Column";
            ribbonDropDownItemImpl2.Label = "Line";
            ribbonDropDownItemImpl3.Label = "Pie";
            ribbonDropDownItemImpl4.Label = "Bar";
            ribbonDropDownItemImpl5.Label = "Table";
            this.ddlChartType.Items.Add(ribbonDropDownItemImpl1);
            this.ddlChartType.Items.Add(ribbonDropDownItemImpl2);
            this.ddlChartType.Items.Add(ribbonDropDownItemImpl3);
            this.ddlChartType.Items.Add(ribbonDropDownItemImpl4);
            this.ddlChartType.Items.Add(ribbonDropDownItemImpl5);
            this.ddlChartType.Label = "Chart Type";
            this.ddlChartType.Name = "ddlChartType";
            this.ddlChartType.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddlChartType_SelectionChanged);
            // 
            // btnPivotView
            // 
            this.btnPivotView.Label = "Pivot View";
            this.btnPivotView.Name = "btnPivotView";
            this.btnPivotView.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPivotView_Click);
            // 
            // btnInsertChart_Click
            // 
            this.btnInsertChart_Click.Label = "Create Chart";
            this.btnInsertChart_Click.Name = "btnInsertChart_Click";
            this.btnInsertChart_Click.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertChart_Click_Click);
            // 
            // btnInsertTable_Click
            // 
            this.btnInsertTable_Click.Label = "Create Table";
            this.btnInsertTable_Click.Name = "btnInsertTable_Click";
            this.btnInsertTable_Click.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInsertTable_Click_Click);
            // 
            // btnCreateChart_Click
            // 
            this.btnCreateChart_Click.Label = "Generate Chart";
            this.btnCreateChart_Click.Name = "btnCreateChart_Click";
            this.btnCreateChart_Click.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCreateChart_Click_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // btnEditWithChartMaker
            // 
            this.btnEditWithChartMaker.Label = "Edit Chart";
            this.btnEditWithChartMaker.Name = "btnEditWithChartMaker";
            this.btnEditWithChartMaker.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditWithChartMaker_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.PowerPoint.Presentation";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUploadExcel;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlDatasets;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertChart_Click;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInsertTable_Click;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddlChartType;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCreateChart_Click;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPivotView;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditWithChartMaker;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
