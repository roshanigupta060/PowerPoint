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
            this.grpConditionalFormatting.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // lable1
            // 
            this.lable1.AutoSize = true;
            this.lable1.Location = new System.Drawing.Point(6, 62);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(29, 13);
            this.lable1.TabIndex = 0;
            this.lable1.Text = "Row";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(45, 107);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Value";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(178, 107);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Aggregation";
            // 
            // cmbRowField
            // 
            this.cmbRowField.FormattingEnabled = true;
            this.cmbRowField.Location = new System.Drawing.Point(97, 54);
            this.cmbRowField.Name = "cmbRowField";
            this.cmbRowField.Size = new System.Drawing.Size(121, 21);
            this.cmbRowField.TabIndex = 3;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(107, 234);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 6;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // clbValueFields
            // 
            this.clbValueFields.FormattingEnabled = true;
            this.clbValueFields.Location = new System.Drawing.Point(9, 126);
            this.clbValueFields.Name = "clbValueFields";
            this.clbValueFields.Size = new System.Drawing.Size(120, 94);
            this.clbValueFields.TabIndex = 7;
            this.clbValueFields.SelectedIndexChanged += new System.EventHandler(this.clbValueFields_SelectedIndexChanged);
            // 
            // clbAggregations
            // 
            this.clbAggregations.FormattingEnabled = true;
            this.clbAggregations.Location = new System.Drawing.Point(156, 126);
            this.clbAggregations.Name = "clbAggregations";
            this.clbAggregations.Size = new System.Drawing.Size(120, 94);
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
            this.cmbChartType.Location = new System.Drawing.Point(97, 22);
            this.cmbChartType.Name = "cmbChartType";
            this.cmbChartType.Size = new System.Drawing.Size(121, 21);
            this.cmbChartType.TabIndex = 9;
            this.cmbChartType.SelectedIndexChanged += new System.EventHandler(this.cmbChartType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Chart Type";
            // 
            // txtFieldName
            // 
            this.txtFieldName.Location = new System.Drawing.Point(149, 19);
            this.txtFieldName.Name = "txtFieldName";
            this.txtFieldName.Size = new System.Drawing.Size(100, 20);
            this.txtFieldName.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(6, 26);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(101, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Enter Column Name";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(6, 70);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Enter Formula";
            // 
            // txtFormula
            // 
            this.txtFormula.Location = new System.Drawing.Point(149, 63);
            this.txtFormula.Name = "txtFormula";
            this.txtFormula.Size = new System.Drawing.Size(100, 20);
            this.txtFormula.TabIndex = 15;
            // 
            // btnAddField
            // 
            this.btnAddField.Location = new System.Drawing.Point(70, 98);
            this.btnAddField.Name = "btnAddField";
            this.btnAddField.Size = new System.Drawing.Size(75, 23);
            this.btnAddField.TabIndex = 16;
            this.btnAddField.Text = "Add Column";
            this.btnAddField.UseVisualStyleBackColor = true;
            this.btnAddField.Click += new System.EventHandler(this.btnAddField_Click_1);
            // 
            // valueContextMenu
            // 
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
            this.grpConditionalFormatting.Location = new System.Drawing.Point(319, 158);
            this.grpConditionalFormatting.Name = "grpConditionalFormatting";
            this.grpConditionalFormatting.Size = new System.Drawing.Size(310, 166);
            this.grpConditionalFormatting.TabIndex = 17;
            this.grpConditionalFormatting.TabStop = false;
            this.grpConditionalFormatting.Text = "Conditional Formatting";
            // 
            // lstRules
            // 
            this.lstRules.FormattingEnabled = true;
            this.lstRules.Location = new System.Drawing.Point(6, 89);
            this.lstRules.Name = "lstRules";
            this.lstRules.Size = new System.Drawing.Size(185, 56);
            this.lstRules.TabIndex = 8;
            // 
            // btnDeleteRule
            // 
            this.btnDeleteRule.Location = new System.Drawing.Point(220, 118);
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Size = new System.Drawing.Size(75, 23);
            this.btnDeleteRule.TabIndex = 7;
            this.btnDeleteRule.Text = "Delete Rule";
            this.btnDeleteRule.UseVisualStyleBackColor = true;
            this.btnDeleteRule.Click += new System.EventHandler(this.btnDeleteRule_Click);
            // 
            // btnAddRule
            // 
            this.btnAddRule.Location = new System.Drawing.Point(220, 89);
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Size = new System.Drawing.Size(75, 23);
            this.btnAddRule.TabIndex = 6;
            this.btnAddRule.Text = "Add Rule";
            this.btnAddRule.UseVisualStyleBackColor = true;
            this.btnAddRule.Click += new System.EventHandler(this.btnAddRule_Click);
            // 
            // btnPickColor
            // 
            this.btnPickColor.Location = new System.Drawing.Point(220, 60);
            this.btnPickColor.Name = "btnPickColor";
            this.btnPickColor.Size = new System.Drawing.Size(52, 23);
            this.btnPickColor.TabIndex = 5;
            this.btnPickColor.Text = "Color";
            this.btnPickColor.UseVisualStyleBackColor = true;
            this.btnPickColor.Click += new System.EventHandler(this.btnPickColor_Click);
            // 
            // txtThreshold
            // 
            this.txtThreshold.Location = new System.Drawing.Point(140, 62);
            this.txtThreshold.Name = "txtThreshold";
            this.txtThreshold.Size = new System.Drawing.Size(51, 20);
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
            this.cmbOperator.Location = new System.Drawing.Point(70, 60);
            this.cmbOperator.Name = "cmbOperator";
            this.cmbOperator.Size = new System.Drawing.Size(59, 21);
            this.cmbOperator.TabIndex = 3;
            // 
            // cmbField
            // 
            this.cmbField.FormattingEnabled = true;
            this.cmbField.Location = new System.Drawing.Point(70, 26);
            this.cmbField.Name = "cmbField";
            this.cmbField.Size = new System.Drawing.Size(121, 21);
            this.cmbField.TabIndex = 2;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(4, 68);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(48, 13);
            this.label8.TabIndex = 1;
            this.label8.Text = "Operator";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(4, 34);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(45, 13);
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
            this.groupBox1.Location = new System.Drawing.Point(319, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(310, 127);
            this.groupBox1.TabIndex = 18;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Calculate Formula";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.cmbChartType);
            this.groupBox2.Controls.Add(this.lable1);
            this.groupBox2.Controls.Add(this.cmbRowField);
            this.groupBox2.Controls.Add(this.btnGenerate);
            this.groupBox2.Controls.Add(this.clbAggregations);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.clbValueFields);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Location = new System.Drawing.Point(12, 13);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(282, 268);
            this.groupBox2.TabIndex = 19;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Generate Chart/Table";
            // 
            // Pivot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 365);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.grpConditionalFormatting);
            this.Name = "Pivot";
            this.Text = "Pivot";
            this.grpConditionalFormatting.ResumeLayout(false);
            this.grpConditionalFormatting.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
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
    }
}