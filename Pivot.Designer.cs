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
            this.lable1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbRowField = new System.Windows.Forms.ComboBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.clbValueFields = new System.Windows.Forms.CheckedListBox();
            this.clbAggregations = new System.Windows.Forms.CheckedListBox();
            this.cmbChartType = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lable1
            // 
            this.lable1.AutoSize = true;
            this.lable1.Location = new System.Drawing.Point(71, 30);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(29, 13);
            this.lable1.TabIndex = 0;
            this.lable1.Text = "Row";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(66, 117);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Value";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(368, 117);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Aggregation";
            // 
            // cmbRowField
            // 
            this.cmbRowField.FormattingEnabled = true;
            this.cmbRowField.Location = new System.Drawing.Point(155, 22);
            this.cmbRowField.Name = "cmbRowField";
            this.cmbRowField.Size = new System.Drawing.Size(121, 21);
            this.cmbRowField.TabIndex = 3;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(310, 244);
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
            this.clbValueFields.Location = new System.Drawing.Point(156, 79);
            this.clbValueFields.Name = "clbValueFields";
            this.clbValueFields.Size = new System.Drawing.Size(120, 94);
            this.clbValueFields.TabIndex = 7;
            // 
            // clbAggregations
            // 
            this.clbAggregations.FormattingEnabled = true;
            this.clbAggregations.Location = new System.Drawing.Point(471, 79);
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
            this.cmbChartType.Location = new System.Drawing.Point(471, 22);
            this.cmbChartType.Name = "cmbChartType";
            this.cmbChartType.Size = new System.Drawing.Size(121, 21);
            this.cmbChartType.TabIndex = 9;
            this.cmbChartType.SelectedIndexChanged += new System.EventHandler(this.cmbChartType_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(368, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 10;
            this.label1.Text = "Chart Type";
            // 
            // Pivot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(641, 365);
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
    }
}