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
            this.cmbValueField = new System.Windows.Forms.ComboBox();
            this.cmbAggregation = new System.Windows.Forms.ComboBox();
            this.btnGenerate = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // lable1
            // 
            this.lable1.AutoSize = true;
            this.lable1.Location = new System.Drawing.Point(66, 92);
            this.lable1.Name = "lable1";
            this.lable1.Size = new System.Drawing.Size(29, 13);
            this.lable1.TabIndex = 0;
            this.lable1.Text = "Row";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(66, 137);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(34, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Value";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(66, 189);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "Aggregation";
            // 
            // cmbRowField
            // 
            this.cmbRowField.FormattingEnabled = true;
            this.cmbRowField.Location = new System.Drawing.Point(155, 84);
            this.cmbRowField.Name = "cmbRowField";
            this.cmbRowField.Size = new System.Drawing.Size(121, 21);
            this.cmbRowField.TabIndex = 3;
            // 
            // cmbValueField
            // 
            this.cmbValueField.FormattingEnabled = true;
            this.cmbValueField.Location = new System.Drawing.Point(155, 129);
            this.cmbValueField.Name = "cmbValueField";
            this.cmbValueField.Size = new System.Drawing.Size(121, 21);
            this.cmbValueField.TabIndex = 4;
            // 
            // cmbAggregation
            // 
            this.cmbAggregation.FormattingEnabled = true;
            this.cmbAggregation.Location = new System.Drawing.Point(155, 189);
            this.cmbAggregation.Name = "cmbAggregation";
            this.cmbAggregation.Size = new System.Drawing.Size(121, 21);
            this.cmbAggregation.TabIndex = 5;
            // 
            // btnGenerate
            // 
            this.btnGenerate.Location = new System.Drawing.Point(103, 240);
            this.btnGenerate.Name = "btnGenerate";
            this.btnGenerate.Size = new System.Drawing.Size(75, 23);
            this.btnGenerate.TabIndex = 6;
            this.btnGenerate.Text = "Generate";
            this.btnGenerate.UseVisualStyleBackColor = true;
            this.btnGenerate.Click += new System.EventHandler(this.btnGenerate_Click);
            // 
            // Pivot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnGenerate);
            this.Controls.Add(this.cmbAggregation);
            this.Controls.Add(this.cmbValueField);
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
        private System.Windows.Forms.ComboBox cmbValueField;
        private System.Windows.Forms.ComboBox cmbAggregation;
        private System.Windows.Forms.Button btnGenerate;
    }
}