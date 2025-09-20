using System.Windows.Forms;

namespace PptExcelSync
{
    partial class Pivot
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lable1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbRowField = new System.Windows.Forms.ComboBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.clbValueFields = new System.Windows.Forms.CheckedListBox();
            this.clbAggregations = new System.Windows.Forms.CheckedListBox();
            this.cmbChartType = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtFieldName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFormula = new System.Windows.Forms.TextBox();
            this.btnAddField = new System.Windows.Forms.Button();
            this.valueContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.grpConditionalFormatting = new System.Windows.Forms.GroupBox();
            this.lstRules = new System.Windows.Forms.ListBox();
            this.btnDeleteRule = new System.Windows.Forms.Button();
            this.btnAddRule = new System.Windows.Forms.Button();
            this.btnPickColor = new System.Windows.Forms.Button();
            this.txtThreshold = new System.Windows.Forms.TextBox();
            this.cmbOperator = new System.Windows.Forms.ComboBox();
            this.cmbField = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.colorDialog1 = new System.Windows.Forms.ColorDialog();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.grpFilters = new System.Windows.Forms.GroupBox();
            this.btnRemoveFilter = new System.Windows.Forms.Button();
            this.btnAddFilter = new System.Windows.Forms.Button();
            this.lstFilters = new System.Windows.Forms.ListBox();
            this.cmbFilterValue = new System.Windows.Forms.ComboBox();
            this.cmbFilterField = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label10 = new System.Windows.Forms.Label();
            this.cmbColumnField = new System.Windows.Forms.ComboBox();
            this.btnMergeFiles = new System.Windows.Forms.Button();
            this.btnRemoveFile = new System.Windows.Forms.Button();
            this.lstDatasetFiles = new System.Windows.Forms.ListBox();
            this.btnAddFiles = new System.Windows.Forms.Button();
            this.label11 = new System.Windows.Forms.Label();
            this.cmbYoYColumnField = new System.Windows.Forms.ComboBox();
            this.grpConditionalFormatting.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.grpFilters.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // lable1
            // 
            this.lable1.AutoSize = true;
            this.lable1.Location = new System.Drawing.Point(8, 76);
            this.lable1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(34, 16);
            this.lable1.TabIndex = 0;
            this.lable1.Text = "Row";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(60, 132);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(42, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Value";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(237, 132);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(81, 16);
            this.label3.TabIndex = 2;
            this.label3.Text = "Aggregation";
            // 
            // cmbRowField
            // 
            this.cmbRowField.FormattingEnabled = true;
            this.cmbRowField.Location = new System.Drawing.Point(129, 66);
            this.cmbRowField.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbRowField.Name = "cmbRowField";
            this.cmbRowField.Size = new System.Drawing.Size(160, 24);
            this.cmbRowField.TabIndex = 3;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(147, 329);
            this.btnGenerate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(100, 28);
            this.btnGenerate.TabIndex = 6;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // clbValueFields
            // 
            this.clbValueFields.FormattingEnabled = true;
            this.clbValueFields.Location = new System.Drawing.Point(12, 155);
            this.clbValueFields.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.clbValueFields.Name = "clbValueFields";
            this.clbValueFields.Size = new System.Drawing.Size(159, 106);
            this.clbValueFields.TabIndex = 7;
            this.clbValueFields.SelectedIndexChanged += new System.EventHandler(this.clbValueFields_SelectedIndexChanged);
            // 
            // clbAggregations
            // 
            this.clbAggregations.FormattingEnabled = true;
            this.clbAggregations.Location = new System.Drawing.Point(208, 155);
            this.clbAggregations.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.clbAggregations.Name = "clbAggregations";
            this.clbAggregations.Size = new System.Drawing.Size(159, 106);
            this.clbAggregations.TabIndex = 8;
            // 
            // cmbChartType
            // 
            this.cmbChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbChartType.FormattingEnabled = true;
            this.cmbChartType.Items.AddRange(new object[] {
            "Column",
            "Table",
            "Line",
            "Bar",
            "Pie"});
            this.cmbChartType.Location = new System.Drawing.Point(129, 27);
            this.cmbChartType.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbChartType.Name = "cmbChartType";
            this.cmbChartType.Size = new System.Drawing.Size(160, 24);
            this.cmbChartType.TabIndex = 9;
            this.cmbChartType.SelectedIndexChanged += new System.EventHandler(this.cmbChartType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 37);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(73, 16);
            this.label1.TabIndex = 10;
            this.label1.Text = "Chart Type";
            // 
            // txtFieldName
            // 
            this.txtFieldName.Location = new System.Drawing.Point(199, 23);
            this.txtFieldName.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtFieldName.Name = "txtFieldName";
            this.txtFieldName.Size = new System.Drawing.Size(132, 22);
            this.txtFieldName.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(8, 32);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(126, 16);
            this.label5.TabIndex = 13;
            this.label5.Text = "Enter Column Name";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(8, 75);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(90, 16);
            this.label6.TabIndex = 14;
            this.label6.Text = "Enter Formula";
            // 
            // txtFormula
            // 
            this.txtFormula.Location = new System.Drawing.Point(199, 66);
            this.txtFormula.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtFormula.Name = "txtFormula";
            this.txtFormula.Size = new System.Drawing.Size(132, 22);
            this.txtFormula.TabIndex = 15;
            // 
            // btnAddField
            // 
            this.btnAddField.Location = new System.Drawing.Point(199, 119);
            this.btnAddField.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddField.Name = "btnAddField";
            this.btnAddField.Size = new System.Drawing.Size(100, 28);
            this.btnAddField.TabIndex = 16;
            this.btnAddField.Text = "Add Column";
            this.btnAddField.UseVisualStyleBackColor = true;
            this.btnAddField.Click += new System.EventHandler(this.btnAddField_Click_1);
            // 
            // valueContextMenu
            // 
            this.valueContextMenu.ImageScalingSize = new System.Drawing.Size(20, 20);
            this.valueContextMenu.Name = "valueContextMenu";
            this.valueContextMenu.Size = new System.Drawing.Size(61, 4);
            // 
            // grpConditionalFormatting
            // 
            this.grpConditionalFormatting.Controls.Add(this.lstRules);
            this.grpConditionalFormatting.Controls.Add(this.btnDeleteRule);
            this.grpConditionalFormatting.Controls.Add(this.btnAddRule);
            this.grpConditionalFormatting.Controls.Add(this.btnPickColor);
            this.grpConditionalFormatting.Controls.Add(this.txtThreshold);
            this.grpConditionalFormatting.Controls.Add(this.cmbOperator);
            this.grpConditionalFormatting.Controls.Add(this.cmbField);
            this.grpConditionalFormatting.Controls.Add(this.label8);
            this.grpConditionalFormatting.Controls.Add(this.label7);
            this.grpConditionalFormatting.Location = new System.Drawing.Point(16, 193);
            this.grpConditionalFormatting.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grpConditionalFormatting.Name = "grpConditionalFormatting";
            this.grpConditionalFormatting.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grpConditionalFormatting.Size = new System.Drawing.Size(413, 204);
            this.grpConditionalFormatting.TabIndex = 17;
            this.grpConditionalFormatting.TabStop = false;
            this.grpConditionalFormatting.Text = "Conditional Formatting";
            // 
            // lstRules
            // 
            this.lstRules.FormattingEnabled = true;
            this.lstRules.ItemHeight = 16;
            this.lstRules.Location = new System.Drawing.Point(8, 110);
            this.lstRules.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lstRules.Name = "lstRules";
            this.lstRules.Size = new System.Drawing.Size(245, 68);
            this.lstRules.TabIndex = 8;
            // 
            // btnDeleteRule
            // 
            this.btnDeleteRule.Location = new System.Drawing.Point(293, 145);
            this.btnDeleteRule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Size = new System.Drawing.Size(100, 28);
            this.btnDeleteRule.TabIndex = 7;
            this.btnDeleteRule.Text = "Delete Rule";
            this.btnDeleteRule.UseVisualStyleBackColor = true;
            this.btnDeleteRule.Click += new System.EventHandler(this.btnDeleteRule_Click);
            // 
            // btnAddRule
            // 
            this.btnAddRule.Location = new System.Drawing.Point(293, 110);
            this.btnAddRule.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Size = new System.Drawing.Size(100, 28);
            this.btnAddRule.TabIndex = 6;
            this.btnAddRule.Text = "Add Rule";
            this.btnAddRule.UseVisualStyleBackColor = true;
            this.btnAddRule.Click += new System.EventHandler(this.btnAddRule_Click);
            // 
            // btnPickColor
            // 
            this.btnPickColor.Location = new System.Drawing.Point(293, 74);
            this.btnPickColor.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnPickColor.Name = "btnPickColor";
            this.btnPickColor.Size = new System.Drawing.Size(69, 28);
            this.btnPickColor.TabIndex = 5;
            this.btnPickColor.Text = "Color";
            this.btnPickColor.UseVisualStyleBackColor = true;
            this.btnPickColor.Click += new System.EventHandler(this.btnPickColor_Click);
            // 
            // txtThreshold
            // 
            this.txtThreshold.Location = new System.Drawing.Point(187, 76);
            this.txtThreshold.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.txtThreshold.Name = "txtThreshold";
            this.txtThreshold.Size = new System.Drawing.Size(67, 22);
            this.txtThreshold.TabIndex = 4;
            // 
            // cmbOperator
            // 
            this.cmbOperator.FormattingEnabled = true;
            this.cmbOperator.Items.AddRange(new object[] {
            ">",
            "<",
            ">=",
            "<=",
            "="});
            this.cmbOperator.Location = new System.Drawing.Point(93, 74);
            this.cmbOperator.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbOperator.Name = "cmbOperator";
            this.cmbOperator.Size = new System.Drawing.Size(77, 24);
            this.cmbOperator.TabIndex = 3;
            // 
            // cmbField
            // 
            this.cmbField.FormattingEnabled = true;
            this.cmbField.Location = new System.Drawing.Point(93, 32);
            this.cmbField.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbField.Name = "cmbField";
            this.cmbField.Size = new System.Drawing.Size(160, 24);
            this.cmbField.TabIndex = 2;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(5, 84);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(60, 16);
            this.label8.TabIndex = 1;
            this.label8.Text = "Operator";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(5, 42);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(56, 16);
            this.label7.TabIndex = 0;
            this.label7.Text = "Apply to";
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.txtFieldName);
            this.groupBox1.Controls.Add(this.txtFormula);
            this.groupBox1.Controls.Add(this.btnAddField);
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Location = new System.Drawing.Point(16, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox1.Size = new System.Drawing.Size(413, 167);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Calculate Formula";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.cmbYoYColumnField);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.cmbChartType);
            this.groupBox2.Controls.Add(this.lable1);
            this.groupBox2.Controls.Add(this.cmbRowField);
            this.groupBox2.Controls.Add(this.btnGenerate);
            this.groupBox2.Controls.Add(this.clbAggregations);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.clbValueFields);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Location = new System.Drawing.Point(453, 193);
            this.groupBox2.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox2.Size = new System.Drawing.Size(405, 379);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Generate Chart/Table";
            // 
            // grpFilters
            // 
            this.grpFilters.Controls.Add(this.btnRemoveFilter);
            this.grpFilters.Controls.Add(this.btnAddFilter);
            this.grpFilters.Controls.Add(this.lstFilters);
            this.grpFilters.Controls.Add(this.cmbFilterValue);
            this.grpFilters.Controls.Add(this.cmbFilterField);
            this.grpFilters.Controls.Add(this.label9);
            this.grpFilters.Controls.Add(this.label4);
            this.grpFilters.Location = new System.Drawing.Point(16, 414);
            this.grpFilters.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grpFilters.Name = "grpFilters";
            this.grpFilters.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.grpFilters.Size = new System.Drawing.Size(411, 156);
            this.grpFilters.TabIndex = 20;
            this.grpFilters.TabStop = false;
            this.grpFilters.Text = "Filters";
            // 
            // btnRemoveFilter
            // 
            this.btnRemoveFilter.Location = new System.Drawing.Point(120, 121);
            this.btnRemoveFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnRemoveFilter.Name = "btnRemoveFilter";
            this.btnRemoveFilter.Size = new System.Drawing.Size(100, 28);
            this.btnRemoveFilter.TabIndex = 6;
            this.btnRemoveFilter.Text = "Remove";
            this.btnRemoveFilter.UseVisualStyleBackColor = true;
            this.btnRemoveFilter.Click += new System.EventHandler(this.btnRemoveFilter_Click);
            // 
            // btnAddFilter
            // 
            this.btnAddFilter.Location = new System.Drawing.Point(12, 121);
            this.btnAddFilter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddFilter.Name = "btnAddFilter";
            this.btnAddFilter.Size = new System.Drawing.Size(100, 28);
            this.btnAddFilter.TabIndex = 5;
            this.btnAddFilter.Text = "Add Filter";
            this.btnAddFilter.UseVisualStyleBackColor = true;
            this.btnAddFilter.Click += new System.EventHandler(this.btnAddFilter_Click);
            // 
            // lstFilters
            // 
            this.lstFilters.FormattingEnabled = true;
            this.lstFilters.ItemHeight = 16;
            this.lstFilters.Location = new System.Drawing.Point(252, 12);
            this.lstFilters.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lstFilters.Name = "lstFilters";
            this.lstFilters.Size = new System.Drawing.Size(149, 132);
            this.lstFilters.TabIndex = 4;
            // 
            // cmbFilterValue
            // 
            this.cmbFilterValue.FormattingEnabled = true;
            this.cmbFilterValue.Location = new System.Drawing.Point(56, 65);
            this.cmbFilterValue.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbFilterValue.Name = "cmbFilterValue";
            this.cmbFilterValue.Size = new System.Drawing.Size(160, 24);
            this.cmbFilterValue.TabIndex = 3;
            // 
            // cmbFilterField
            // 
            this.cmbFilterField.FormattingEnabled = true;
            this.cmbFilterField.Location = new System.Drawing.Point(55, 27);
            this.cmbFilterField.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbFilterField.Name = "cmbFilterField";
            this.cmbFilterField.Size = new System.Drawing.Size(160, 24);
            this.cmbFilterField.TabIndex = 2;
            this.cmbFilterField.SelectedIndexChanged += new System.EventHandler(this.cmbFilterField_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(7, 76);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(36, 16);
            this.label9.TabIndex = 1;
            this.label9.Text = "View";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(8, 37);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(37, 16);
            this.label4.TabIndex = 0;
            this.label4.Text = "Field";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label10);
            this.groupBox3.Controls.Add(this.cmbColumnField);
            this.groupBox3.Controls.Add(this.btnMergeFiles);
            this.groupBox3.Controls.Add(this.btnRemoveFile);
            this.groupBox3.Controls.Add(this.lstDatasetFiles);
            this.groupBox3.Controls.Add(this.btnAddFiles);
            this.groupBox3.Location = new System.Drawing.Point(453, 15);
            this.groupBox3.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Padding = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.groupBox3.Size = new System.Drawing.Size(405, 167);
            this.groupBox3.TabIndex = 21;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "YoY Comparison";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(8, 66);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(52, 16);
            this.label10.TabIndex = 6;
            this.label10.Text = "Column";
            // 
            // cmbColumnField
            // 
            this.cmbColumnField.FormattingEnabled = true;
            this.cmbColumnField.Location = new System.Drawing.Point(85, 60);
            this.cmbColumnField.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.cmbColumnField.Name = "cmbColumnField";
            this.cmbColumnField.Size = new System.Drawing.Size(160, 24);
            this.cmbColumnField.TabIndex = 5;
            // 
            // btnMergeFiles
            // 
            this.btnMergeFiles.Location = new System.Drawing.Point(267, 83);
            this.btnMergeFiles.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnMergeFiles.Name = "btnMergeFiles";
            this.btnMergeFiles.Size = new System.Drawing.Size(100, 28);
            this.btnMergeFiles.TabIndex = 4;
            this.btnMergeFiles.Text = "Merge Files";
            this.btnMergeFiles.UseVisualStyleBackColor = true;
            this.btnMergeFiles.Click += new System.EventHandler(this.btnMergeFiles_Click);
            // 
            // btnRemoveFile
            // 
            this.btnRemoveFile.Location = new System.Drawing.Point(267, 119);
            this.btnRemoveFile.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnRemoveFile.Name = "btnRemoveFile";
            this.btnRemoveFile.Size = new System.Drawing.Size(100, 28);
            this.btnRemoveFile.TabIndex = 3;
            this.btnRemoveFile.Text = "Remove";
            this.btnRemoveFile.UseVisualStyleBackColor = true;
            this.btnRemoveFile.Click += new System.EventHandler(this.btnRemoveFile_Click);
            // 
            // lstDatasetFiles
            // 
            this.lstDatasetFiles.FormattingEnabled = true;
            this.lstDatasetFiles.ItemHeight = 16;
            this.lstDatasetFiles.Location = new System.Drawing.Point(12, 92);
            this.lstDatasetFiles.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.lstDatasetFiles.Name = "lstDatasetFiles";
            this.lstDatasetFiles.Size = new System.Drawing.Size(233, 68);
            this.lstDatasetFiles.TabIndex = 2;
            // 
            // btnAddFiles
            // 
            this.btnAddFiles.Location = new System.Drawing.Point(12, 25);
            this.btnAddFiles.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnAddFiles.Name = "btnAddFiles";
            this.btnAddFiles.Size = new System.Drawing.Size(100, 28);
            this.btnAddFiles.TabIndex = 1;
            this.btnAddFiles.Text = "Add File";
            this.btnAddFiles.UseVisualStyleBackColor = true;
            this.btnAddFiles.Click += new System.EventHandler(this.btnAddFiles_Click);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(7, 101);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(109, 16);
            this.label11.TabIndex = 11;
            this.label11.Text = "YoY Comparison";
            // 
            // cmbYoYColumnField
            // 
            this.cmbYoYColumnField.FormattingEnabled = true;
            this.cmbYoYColumnField.Location = new System.Drawing.Point(129, 98);
            this.cmbYoYColumnField.Margin = new System.Windows.Forms.Padding(4);
            this.cmbYoYColumnField.Name = "cmbYoYColumnField";
            this.cmbYoYColumnField.Size = new System.Drawing.Size(160, 24);
            this.cmbYoYColumnField.TabIndex = 12;
            // 
            // Pivot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(883, 587);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.grpFilters);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpConditionalFormatting);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "Pivot";
            this.Text = "Pivot";
            this.grpConditionalFormatting.ResumeLayout(false);
            this.grpConditionalFormatting.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.grpFilters.ResumeLayout(false);
            this.grpFilters.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Label lable1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbRowField;
        private System.Windows.Forms.Button btnGenerate;
        private System.Windows.Forms.CheckedListBox clbValueFields;
        private System.Windows.Forms.CheckedListBox clbAggregations;
        private System.Windows.Forms.ComboBox cmbChartType;
        private System.Windows.Forms.Label label1;
        private TextBox txtFieldName;
        private Label label5;
        private Label label6;
        private TextBox txtFormula;
        private Button btnAddField;
        private ContextMenuStrip valueContextMenu;
        private GroupBox grpConditionalFormatting;
        private Label label8;
        private Label label7;
        private TextBox txtThreshold;
        private ComboBox cmbOperator;
        private ComboBox cmbField;
        private Button btnDeleteRule;
        private Button btnAddRule;
        private Button btnPickColor;
        private ListBox lstRules;
        private ColorDialog colorDialog1;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private GroupBox grpFilters;
        private Button btnRemoveFilter;
        private Button btnAddFilter;
        private ListBox lstFilters;
        private ComboBox cmbFilterValue;
        private ComboBox cmbFilterField;
        private Label label9;
        private Label label4;
        private GroupBox groupBox3;
        private ComboBox cmbColumnField;
        private Button btnMergeFiles;
        private Button btnRemoveFile;
        private ListBox lstDatasetFiles;
        private Button btnAddFiles;
        private Label label10;
        private ComboBox cmbYoYColumnField;
        private Label label11;
    }
}