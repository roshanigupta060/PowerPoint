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
            this.label4 = new System.Windows.Forms.Label();
            this.txtFieldName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFormula = new System.Windows.Forms.TextBox();
            this.btnAddField = new System.Windows.Forms.Button();
            this.valueContextMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.SuspendLayout();
            // 
            // lable1
            // 
            this.lable1.AutoSize = true;
            this.lable1.Location = new System.Drawing.Point(25, 79);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(29, 13);
            this.lable1.TabIndex = 0;
            this.lable1.Text = "Row";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(25, 117);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Value";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(25, 241);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Aggregation";
            // 
            // cmbRowField
            // 
            this.cmbRowField.FormattingEnabled = true;
            this.cmbRowField.Location = new System.Drawing.Point(97, 79);
            this.cmbRowField.Name = "cmbRowField";
            this.cmbRowField.Size = new System.Drawing.Size(121, 21);
            this.cmbRowField.TabIndex = 3;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(459, 312);
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
            this.clbValueFields.Location = new System.Drawing.Point(97, 117);
            this.clbValueFields.Name = "clbValueFields";
            this.clbValueFields.Size = new System.Drawing.Size(120, 94);
            this.clbValueFields.TabIndex = 7;
            // 
            // clbAggregations
            // 
            this.clbAggregations.FormattingEnabled = true;
            this.clbAggregations.Location = new System.Drawing.Point(98, 241);
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
            this.cmbChartType.Location = new System.Drawing.Point(98, 27);
            this.cmbChartType.Name = "cmbChartType";
            this.cmbChartType.Size = new System.Drawing.Size(121, 21);
            this.cmbChartType.TabIndex = 9;
            this.cmbChartType.SelectedIndexChanged += new System.EventHandler(this.cmbChartType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(25, 27);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Chart Type";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(386, 9);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(82, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Calculated Field";
            // 
            // txtFieldName
            // 
            this.txtFieldName.Location = new System.Drawing.Point(434, 38);
            this.txtFieldName.Name = "txtFieldName";
            this.txtFieldName.Size = new System.Drawing.Size(100, 20);
            this.txtFieldName.TabIndex = 12;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(316, 45);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(101, 13);
            this.label5.TabIndex = 13;
            this.label5.Text = "Enter Column Name";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(316, 86);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Enter Formula";
            // 
            // txtFormula
            // 
            this.txtFormula.Location = new System.Drawing.Point(434, 79);
            this.txtFormula.Name = "txtFormula";
            this.txtFormula.Size = new System.Drawing.Size(100, 20);
            this.txtFormula.TabIndex = 15;
            // 
            // btnAddField
            // 
            this.btnAddField.Location = new System.Drawing.Point(389, 127);
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
            // Pivot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 365);
            this.Controls.Add(this.btnAddField);
            this.Controls.Add(this.txtFormula);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtFieldName);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmbChartType);
            this.Controls.Add(this.clbAggregations);
            this.Controls.Add(this.clbValueFields);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.cmbRowField);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.lable1);
            this.Name = "Pivot";
            this.Text = "Pivot";
            this.ResumeLayout(false);
            this.PerformLayout();

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
        private Label label4;
        private TextBox txtFieldName;
        private Label label5;
        private Label label6;
        private TextBox txtFormula;
        private Button btnAddField;
        private ContextMenuStrip valueContextMenu;
    }
}