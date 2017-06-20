namespace AMO.EnPI.AddIn
{
    partial class RegressionControl
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RegressionControl));
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.checkedListBox2 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblRegressionTitle = new System.Windows.Forms.Label();
            this.checkBox1 = new System.Windows.Forms.CheckBox();
            this.chkCo2Emission = new System.Windows.Forms.CheckBox();
            this.label4 = new System.Windows.Forms.Label();
            this.checkedListBox4 = new System.Windows.Forms.CheckedListBox();
            this.label3 = new System.Windows.Forms.Label();
            this.checkedListBox3 = new System.Windows.Forms.CheckedListBox();
            this.lblReportYear = new System.Windows.Forms.Label();
            this.listBox3 = new System.Windows.Forms.ListBox();
            this.SuspendLayout();
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.CheckOnClick = true;
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(7, 229);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(120, 94);
            this.checkedListBox1.TabIndex = 0;
            this.checkedListBox1.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.updateCheckboxes);
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(26, 951);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(75, 23);
            this.btnRun.TabIndex = 3;
            this.btnRun.Text = "Next";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Visible = false;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(7, 683);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(120, 56);
            this.listBox1.TabIndex = 4;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // checkedListBox2
            // 
            this.checkedListBox2.CheckOnClick = true;
            this.checkedListBox2.FormattingEnabled = true;
            this.checkedListBox2.Location = new System.Drawing.Point(7, 348);
            this.checkedListBox2.Name = "checkedListBox2";
            this.checkedListBox2.Size = new System.Drawing.Size(120, 94);
            this.checkedListBox2.TabIndex = 5;
            this.checkedListBox2.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.updateCheckboxes);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 210);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(82, 13);
            this.label1.TabIndex = 6;
            this.label1.Text = "Energy Sources";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(7, 331);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Variables";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(7, 663);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(72, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Baseline Year";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(7, 747);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(61, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "Model Year";
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(7, 764);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(120, 69);
            this.listBox2.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoEllipsis = true;
            this.label7.Location = new System.Drawing.Point(4, 39);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(134, 164);
            this.label7.TabIndex = 15;
            this.label7.Text = resources.GetString("label7.Text");
            // 
            // lblRegressionTitle
            // 
            this.lblRegressionTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblRegressionTitle.Location = new System.Drawing.Point(7, 4);
            this.lblRegressionTitle.Name = "lblRegressionTitle";
            this.lblRegressionTitle.Size = new System.Drawing.Size(120, 35);
            this.lblRegressionTitle.TabIndex = 16;
            this.lblRegressionTitle.Text = "Select Data for Calculations";
            this.lblRegressionTitle.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // checkBox1
            // 
            this.checkBox1.Location = new System.Drawing.Point(7, 976);
            this.checkBox1.Name = "checkBox1";
            this.checkBox1.Size = new System.Drawing.Size(120, 50);
            this.checkBox1.TabIndex = 17;
            this.checkBox1.Text = "Collect cost data and calculate cost savings";
            this.checkBox1.UseVisualStyleBackColor = true;
            this.checkBox1.Visible = false;
            // 
            // chkCo2Emission
            // 
            this.chkCo2Emission.Location = new System.Drawing.Point(7, 977);
            this.chkCo2Emission.Name = "chkCo2Emission";
            this.chkCo2Emission.Size = new System.Drawing.Size(120, 35);
            this.chkCo2Emission.TabIndex = 18;
            this.chkCo2Emission.Text = "Calculate CO2 avoided emissions";
            this.chkCo2Emission.UseVisualStyleBackColor = true;
            this.chkCo2Emission.Visible = false;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(10, 571);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(58, 13);
            this.label4.TabIndex = 22;
            this.label4.Text = "Production";
            this.label4.Visible = false;
            // 
            // checkedListBox4
            // 
            this.checkedListBox4.CheckOnClick = true;
            this.checkedListBox4.FormattingEnabled = true;
            this.checkedListBox4.Location = new System.Drawing.Point(9, 589);
            this.checkedListBox4.Name = "checkedListBox4";
            this.checkedListBox4.Size = new System.Drawing.Size(120, 64);
            this.checkedListBox4.TabIndex = 21;
            this.checkedListBox4.Visible = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(10, 448);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(105, 13);
            this.label3.TabIndex = 20;
            this.label3.Text = "Building Square Feet";
            this.label3.Visible = false;
            // 
            // checkedListBox3
            // 
            this.checkedListBox3.CheckOnClick = true;
            this.checkedListBox3.FormattingEnabled = true;
            this.checkedListBox3.Location = new System.Drawing.Point(9, 466);
            this.checkedListBox3.Name = "checkedListBox3";
            this.checkedListBox3.Size = new System.Drawing.Size(120, 94);
            this.checkedListBox3.TabIndex = 19;
            this.checkedListBox3.Visible = false;
            // 
            // lblReportYear
            // 
            this.lblReportYear.AutoSize = true;
            this.lblReportYear.Location = new System.Drawing.Point(10, 847);
            this.lblReportYear.Name = "lblReportYear";
            this.lblReportYear.Size = new System.Drawing.Size(64, 13);
            this.lblReportYear.TabIndex = 24;
            this.lblReportYear.Text = "Report Year";
            // 
            // listBox3
            // 
            this.listBox3.FormattingEnabled = true;
            this.listBox3.Location = new System.Drawing.Point(10, 864);
            this.listBox3.Name = "listBox3";
            this.listBox3.Size = new System.Drawing.Size(120, 69);
            this.listBox3.TabIndex = 23;
            // 
            // RegressionControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.AutoSize = true;
            this.Controls.Add(this.lblReportYear);
            this.Controls.Add(this.listBox3);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.checkedListBox4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.checkedListBox3);
            this.Controls.Add(this.lblRegressionTitle);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.checkedListBox2);
            this.Controls.Add(this.chkCo2Emission);
            this.Controls.Add(this.checkBox1);
            this.Name = "RegressionControl";
            this.Size = new System.Drawing.Size(141, 1044);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkedListBox1;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.CheckedListBox checkedListBox2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label lblRegressionTitle;
        private System.Windows.Forms.CheckBox checkBox1;
        private System.Windows.Forms.CheckBox chkCo2Emission;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.CheckedListBox checkedListBox4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckedListBox checkedListBox3;
        private System.Windows.Forms.Label lblReportYear;
        private System.Windows.Forms.ListBox listBox3;


    }
}
