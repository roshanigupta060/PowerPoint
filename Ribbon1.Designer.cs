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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.ddlDatasets = this.Factory.CreateRibbonDropDown();
            this.btnUploadExcel = this.Factory.CreateRibbonButton();
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1.SuspendLayout();
            this.tab1.SuspendLayout();
            this.SuspendLayout();
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnUploadExcel);
            this.group1.Items.Add(this.ddlDatasets);
            this.group1.Name = "group1";
            // 
            // ddlDatasets
            // 
            this.ddlDatasets.Image = global::PptExcelSync.Properties.Resources.datasets;
            this.ddlDatasets.Label = "Select File";
            this.ddlDatasets.Name = "ddlDatasets";
            this.ddlDatasets.ShowImage = true;
            this.ddlDatasets.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddlDatasets_SelectionChanged);
            // 
            // btnUploadExcel
            // 
            this.btnUploadExcel.Image = global::PptExcelSync.Properties.Resources.upload;
            this.btnUploadExcel.Label = "Upload Excel";
            this.btnUploadExcel.Name = "btnUploadExcel";
            this.btnUploadExcel.ShowImage = true;
            this.btnUploadExcel.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUploadExcel_Click);
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
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
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
