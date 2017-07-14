namespace AMO.EnPI.AddIn
{
    partial class WizardControl
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
            this.lblTop = new System.Windows.Forms.Label();
            this.btnNext = new System.Windows.Forms.Button();
            this.lblQuestion = new System.Windows.Forms.Label();
            this.btnBack = new System.Windows.Forms.Button();
            this.lblQuestion2 = new System.Windows.Forms.Label();
            this.btnChangeModel = new System.Windows.Forms.Button();
            this.btnIndVar = new System.Windows.Forms.Button();
            this.btnActualData = new System.Windows.Forms.Button();
            this.btnRegression = new System.Windows.Forms.Button();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.cbEnergy = new System.Windows.Forms.ComboBox();
            this.cbIndVar = new System.Windows.Forms.ComboBox();
            this.btnYear = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnSpecialNext = new System.Windows.Forms.Button();
            this.btnSpecialBack = new System.Windows.Forms.Button();
            this.btnFormatTable = new System.Windows.Forms.Button();
            this.btnAddData = new System.Windows.Forms.Button();
            this.btnHasData = new System.Windows.Forms.Button();
            this.btnRollup = new System.Windows.Forms.Button();
            this.btnConvertUnits = new System.Windows.Forms.Button();
            this.btnAddEnergy = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // lblTop
            // 
            this.lblTop.AutoEllipsis = true;
            this.lblTop.AutoSize = true;
            this.lblTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.lblTop.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTop.Location = new System.Drawing.Point(0, 0);
            this.lblTop.MaximumSize = new System.Drawing.Size(220, 0);
            this.lblTop.Name = "lblTop";
            this.lblTop.Size = new System.Drawing.Size(30, 13);
            this.lblTop.TabIndex = 0;
            this.lblTop.Text = "N/A";
            this.lblTop.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // btnNext
            // 
            this.btnNext.Location = new System.Drawing.Point(107, 500);
            this.btnNext.Name = "btnNext";
            this.btnNext.Size = new System.Drawing.Size(75, 23);
            this.btnNext.TabIndex = 3;
            this.btnNext.Text = "Next";
            this.btnNext.UseVisualStyleBackColor = true;
            this.btnNext.Click += new System.EventHandler(this.btnNext_Click);
            // 
            // lblQuestion
            // 
            this.lblQuestion.AutoSize = true;
            this.lblQuestion.Location = new System.Drawing.Point(3, 50);
            this.lblQuestion.MaximumSize = new System.Drawing.Size(200, 0);
            this.lblQuestion.Name = "lblQuestion";
            this.lblQuestion.Size = new System.Drawing.Size(27, 16);
            this.lblQuestion.TabIndex = 5;
            this.lblQuestion.Text = "N/A";
            // 
            // btnBack
            // 
            this.btnBack.Location = new System.Drawing.Point(15, 500);
            this.btnBack.Name = "btnBack";
            this.btnBack.Size = new System.Drawing.Size(75, 23);
            this.btnBack.TabIndex = 6;
            this.btnBack.Text = "Back";
            this.btnBack.UseVisualStyleBackColor = true;
            this.btnBack.Click += new System.EventHandler(this.btnBack_Click);
            // 
            // lblQuestion2
            // 
            this.lblQuestion2.AutoSize = true;
            this.lblQuestion2.Location = new System.Drawing.Point(3, 243);
            this.lblQuestion2.MaximumSize = new System.Drawing.Size(200, 0);
            this.lblQuestion2.Name = "lblQuestion2";
            this.lblQuestion2.Size = new System.Drawing.Size(27, 13);
            this.lblQuestion2.TabIndex = 7;
            this.lblQuestion2.Text = "N/A";
            // 
            // btnChangeModel
            // 
            this.btnChangeModel.Location = new System.Drawing.Point(21, 256);
            this.btnChangeModel.Name = "btnChangeModel";
            this.btnChangeModel.Size = new System.Drawing.Size(156, 23);
            this.btnChangeModel.TabIndex = 8;
            this.btnChangeModel.Text = "Change Models";
            this.btnChangeModel.UseVisualStyleBackColor = true;
            this.btnChangeModel.Visible = false;
            this.btnChangeModel.Click += new System.EventHandler(this.btnChangeModel_Click);
            // 
            // btnChangeModel
            // 
            this.btnRollup.Location = new System.Drawing.Point(21, 256);
            this.btnRollup.Name = "btnRollup";
            this.btnRollup.Size = new System.Drawing.Size(156, 23);
            this.btnRollup.TabIndex = 8;
            this.btnRollup.Text = "Corporate Roll-up";
            this.btnRollup.UseVisualStyleBackColor = true;
            this.btnRollup.Visible = false;
            this.btnRollup.Click += new System.EventHandler(this.btnRollup_Click);
            // 
            // btnIndVar
            // 
            this.btnIndVar.Location = new System.Drawing.Point(3, 196);
            this.btnIndVar.Name = "btnIndVar";
            this.btnIndVar.Size = new System.Drawing.Size(191, 23);
            this.btnIndVar.TabIndex = 9;
            this.btnIndVar.Text = "Add selected Independent Variable";
            this.btnIndVar.UseVisualStyleBackColor = true;
            this.btnIndVar.Visible = false;
            this.btnIndVar.Click += new System.EventHandler(this.btnIndVar_Click);
            // 
            // btnActualData
            // 
            this.btnActualData.Location = new System.Drawing.Point(44, 137);
            this.btnActualData.Name = "btnActualData";
            this.btnActualData.Size = new System.Drawing.Size(120, 23);
            this.btnActualData.TabIndex = 10;
            this.btnActualData.Text = "Use Actual Data";
            this.btnActualData.UseVisualStyleBackColor = true;
            this.btnActualData.Visible = false;
            this.btnActualData.Click += new System.EventHandler(this.btnActual_Click);
            // 
            // btnRegression
            // 
            this.btnRegression.Location = new System.Drawing.Point(35, 259);
            this.btnRegression.Name = "btnRegression";
            this.btnRegression.Size = new System.Drawing.Size(131, 23);
            this.btnRegression.TabIndex = 11;
            this.btnRegression.Text = "Regression Analysis";
            this.btnRegression.UseVisualStyleBackColor = true;
            this.btnRegression.Visible = false;
            this.btnRegression.Click += new System.EventHandler(this.btnRegression_Click);
            // 
            // btnAnalyze
            // 
            this.btnAnalyze.Location = new System.Drawing.Point(107, 344);
            this.btnAnalyze.Name = "btnAnalyze";
            this.btnAnalyze.Size = new System.Drawing.Size(75, 23);
            this.btnAnalyze.TabIndex = 12;
            this.btnAnalyze.Text = "Analyze";
            this.btnAnalyze.UseVisualStyleBackColor = true;
            this.btnAnalyze.Visible = false;
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(44, 550);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(107, 23);
            this.btnClose.TabIndex = 13;
            this.btnClose.Text = "Close Wizard";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // cbEnergy
            // 
            this.cbEnergy.FormattingEnabled = true;
            this.cbEnergy.Location = new System.Drawing.Point(0, 212);
            int sizeEnergy = System.Enum.GetValues(typeof(Utilities.Constants.EnergySourceTypes)).Length;
            object[] energy = new object[sizeEnergy];
            int countEnergy = 0;
            foreach (Utilities.Constants.EnergySourceTypes typ in System.Enum.GetValues(typeof(Utilities.Constants.EnergySourceTypes)))
            {
                energy.SetValue(AddIn.EnPIRibbon.rsc.GetString(typ.ToString()), countEnergy);
                countEnergy++;
            }
            this.cbEnergy.Items.AddRange(energy);
            this.cbEnergy.Location = new System.Drawing.Point(0, 212);
            this.cbEnergy.Name = "cbEnergy";
            this.cbEnergy.Size = new System.Drawing.Size(197, 21);
            this.cbEnergy.TabIndex = 14;
            this.cbEnergy.Visible = false;
            // 
            // cbIndVar
            // 
            this.cbIndVar.FormattingEnabled = true;
            int sizeVar = System.Enum.GetValues(typeof(Utilities.Constants.VariableTypes)).Length;
            object[] var = new object[sizeVar];
            int countVar = 0;
            foreach (Utilities.Constants.VariableTypes typ in System.Enum.GetValues(typeof(Utilities.Constants.VariableTypes)))
            {
                var.SetValue(AddIn.EnPIRibbon.rsc.GetString(typ.ToString()), countVar);
                countVar++;
            }
            this.cbIndVar.Items.AddRange(var);
            this.cbIndVar.Location = new System.Drawing.Point(0, 169);
            this.cbIndVar.Name = "cbIndVar";
            this.cbIndVar.Size = new System.Drawing.Size(194, 21);
            this.cbIndVar.TabIndex = 15;
            this.cbIndVar.Visible = false;
            // 
            // btnYear
            // 
            this.btnYear.Location = new System.Drawing.Point(21,465);
            this.btnYear.Name = "btnYear";
            this.btnYear.Size = new System.Drawing.Size(150, 23);
            this.btnYear.TabIndex = 16;
            this.btnYear.Text = "Label Reporting Periods";
            this.btnYear.UseVisualStyleBackColor = true;
            this.btnYear.Visible = false;
            this.btnYear.Click += new System.EventHandler(this.btnYear_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.btnAddEnergy);
            this.panel1.Controls.Add(this.btnYear);
            this.panel1.Controls.Add(this.btnSpecialNext);
            this.panel1.Controls.Add(this.btnSpecialBack);
            this.panel1.Controls.Add(this.btnFormatTable);
            this.panel1.Controls.Add(this.btnAddData);
            this.panel1.Controls.Add(this.btnHasData);
            this.panel1.Controls.Add(this.btnChangeModel);
            this.panel1.Controls.Add(this.btnConvertUnits);
            this.panel1.Controls.Add(this.btnBack);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(211, 600);
            this.panel1.TabIndex = 17;
            // 
            // btnSpecialNext
            // 
            this.btnSpecialNext.Location = new System.Drawing.Point(107, 500);
            this.btnSpecialNext.Name = "btnSpecialNext";
            this.btnSpecialNext.Size = new System.Drawing.Size(75, 23);
            this.btnSpecialNext.TabIndex = 5;
            this.btnSpecialNext.Text = "Next";
            this.btnSpecialNext.UseVisualStyleBackColor = true;
            this.btnSpecialNext.Visible = false;
            this.btnSpecialNext.Click += new System.EventHandler(this.btnSpecialNext_Click);
            // 
            // btnSpecialBack
            // 
            this.btnSpecialBack.Location = new System.Drawing.Point(15, 500);
            this.btnSpecialBack.Name = "btnSpecialBack";
            this.btnSpecialBack.Size = new System.Drawing.Size(75, 23);
            this.btnSpecialBack.TabIndex = 4;
            this.btnSpecialBack.Text = "Back";
            this.btnSpecialBack.UseVisualStyleBackColor = true;
            this.btnSpecialBack.Visible = false;
            this.btnSpecialBack.Click += new System.EventHandler(this.btnSpecialBack_Click);
            // 
            // btnFormatTable
            // 
            this.btnFormatTable.Location = new System.Drawing.Point(10, 303);
            this.btnFormatTable.Name = "btnFormatTable";
            this.btnFormatTable.Size = new System.Drawing.Size(200, 23);
            this.btnFormatTable.TabIndex = 3;
            this.btnFormatTable.Text = "Format data as an Excel table";
            this.btnFormatTable.UseVisualStyleBackColor = true;
            this.btnFormatTable.Visible = false;
            this.btnFormatTable.Click += new System.EventHandler(this.btnFormatTable_Click);
            // 
            // btnAddData
            // 
            this.btnAddData.Location = new System.Drawing.Point(15, 166);
            this.btnAddData.Name = "btnAddData";
            this.btnAddData.Size = new System.Drawing.Size(176, 34);
            this.btnAddData.TabIndex = 2;
            this.btnAddData.Text = "I need to enter my energy and variable data";
            this.btnAddData.UseVisualStyleBackColor = true;
            this.btnAddData.Visible = false;
            this.btnAddData.Click += new System.EventHandler(this.btnAddData_Click);
            // 
            // btnHasData
            // 
            this.btnHasData.Location = new System.Drawing.Point(20, 117);
            this.btnHasData.Name = "btnHasData";
            this.btnHasData.Size = new System.Drawing.Size(151, 23);
            this.btnHasData.TabIndex = 1;
            this.btnHasData.Text = "My Data is in the Sheet";
            this.btnHasData.UseVisualStyleBackColor = true;
            this.btnHasData.Visible = false;
            this.btnHasData.Click += new System.EventHandler(this.btnHasData_Click);
            // 
            // btnConvertUnits
            // 
            this.btnConvertUnits.Location = new System.Drawing.Point(46, 225);
            this.btnConvertUnits.Name = "btnConvertUnits";
            this.btnConvertUnits.Size = new System.Drawing.Size(105, 23);
            this.btnConvertUnits.TabIndex = 0;
            this.btnConvertUnits.Text = "Convert Units";
            this.btnConvertUnits.UseVisualStyleBackColor = true;
            this.btnConvertUnits.Visible = false;
            this.btnConvertUnits.Click += new System.EventHandler(this.btnConvertUnits_Click);
            // 
            // btnAddEnergy
            // 
            this.btnAddEnergy.Location = new System.Drawing.Point(15, 259);
            this.btnAddEnergy.Name = "btnAddEnergy";
            this.btnAddEnergy.Size = new System.Drawing.Size(170, 23);
            this.btnAddEnergy.TabIndex = 17;
            this.btnAddEnergy.Text = "Add selected Energy Source";
            this.btnAddEnergy.UseVisualStyleBackColor = true;
            this.btnAddEnergy.Visible = false;
            this.btnAddEnergy.Click += new System.EventHandler(this.btnAddSource_Click);
            // 
            // WizardControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbIndVar);
            this.Controls.Add(this.cbEnergy);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.btnRegression);
            this.Controls.Add(this.btnActualData);
            this.Controls.Add(this.btnIndVar);
            this.Controls.Add(this.lblQuestion2);
            this.Controls.Add(this.lblQuestion);
            this.Controls.Add(this.btnNext);
            this.Controls.Add(this.lblTop);
            this.Controls.Add(this.btnRollup);
            this.Controls.Add(this.btnAnalyze);
            this.Controls.Add(this.panel1);
            this.Name = "WizardControl";
            this.Size = new System.Drawing.Size(211, 434);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

    }


        #endregion

        private System.Windows.Forms.Label lblTop;
        private System.Windows.Forms.Button btnNext;
        private System.Windows.Forms.Label lblQuestion;
        private System.Windows.Forms.Button btnBack;
        private System.Windows.Forms.Label lblQuestion2;
        private System.Windows.Forms.Button btnChangeModel;
        private System.Windows.Forms.Button btnIndVar;
        private System.Windows.Forms.Button btnActualData;
        private System.Windows.Forms.Button btnRegression;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnRollup;
        private System.Windows.Forms.ComboBox cbEnergy;
        private System.Windows.Forms.ComboBox cbIndVar;
        private System.Windows.Forms.Button btnYear;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnConvertUnits;
        private System.Windows.Forms.Button btnHasData;
        private System.Windows.Forms.Button btnAddData;
        private System.Windows.Forms.Button btnFormatTable;
        private System.Windows.Forms.Button btnSpecialNext;
        private System.Windows.Forms.Button btnSpecialBack;
        private System.Windows.Forms.Button btnAddEnergy;
    }
}
