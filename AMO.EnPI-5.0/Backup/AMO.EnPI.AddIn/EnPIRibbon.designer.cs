namespace AMO.EnPI.AddIn
{
    partial class EnPIRibbon
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
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem1 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem2 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem3 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem4 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem5 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem6 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem7 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem8 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem9 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem10 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItem11 = new Microsoft.Office.Tools.Ribbon.RibbonDropDownItem();
            this.tab1 = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.tabEnPI = new Microsoft.Office.Tools.Ribbon.RibbonTab();
            this.group1 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.dropDownYear = new Microsoft.Office.Tools.Ribbon.RibbonDropDown();
            this.button1 = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.grp_EnergySources = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.menu1 = new Microsoft.Office.Tools.Ribbon.RibbonMenu();
            this.grp_Variables = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.menu2 = new Microsoft.Office.Tools.Ribbon.RibbonMenu();
            this.Wizard = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnWizard = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.btnOutputWizard = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.UnitConversion = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnConvertUnits = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.LabelReportingPeriod = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnReportingPeriod = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.grp_ComputeEnPI_Actual = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btn_Actual = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.grp_ComputeEnPI_Regression = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btn_Regression = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Model = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnModelChange = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.Rollup = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnRollup = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.group2 = new Microsoft.Office.Tools.Ribbon.RibbonGroup();
            this.btnAbout = new Microsoft.Office.Tools.Ribbon.RibbonButton();
            this.tab1.SuspendLayout();
            this.tabEnPI.SuspendLayout();
            this.group1.SuspendLayout();
            this.grp_EnergySources.SuspendLayout();
            this.grp_Variables.SuspendLayout();
            this.Wizard.SuspendLayout();
            this.UnitConversion.SuspendLayout();
            this.LabelReportingPeriod.SuspendLayout();
            this.grp_ComputeEnPI_Actual.SuspendLayout();
            this.grp_ComputeEnPI_Regression.SuspendLayout();
            this.Model.SuspendLayout();
            this.Rollup.SuspendLayout();
            this.group2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // tabEnPI
            // 
            this.tabEnPI.Groups.Add(this.group1);
            this.tabEnPI.Groups.Add(this.grp_EnergySources);
            this.tabEnPI.Groups.Add(this.grp_Variables);
            this.tabEnPI.Groups.Add(this.Wizard);
            this.tabEnPI.Groups.Add(this.UnitConversion);
            this.tabEnPI.Groups.Add(this.LabelReportingPeriod);
            this.tabEnPI.Groups.Add(this.grp_ComputeEnPI_Actual);
            this.tabEnPI.Groups.Add(this.grp_ComputeEnPI_Regression);
            this.tabEnPI.Groups.Add(this.Model);
            this.tabEnPI.Groups.Add(this.Rollup);
            this.tabEnPI.Groups.Add(this.group2);
            this.tabEnPI.Label = "EnPI  ";
            this.tabEnPI.Name = "tabEnPI";
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDownYear);
            this.group1.Items.Add(this.button1);
            this.group1.Label = "Set Year for Selection";
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // dropDownYear
            // 
            ribbonDropDownItem2.Label = "Baseline";
            ribbonDropDownItem3.Label = "1";
            ribbonDropDownItem4.Label = "2";
            ribbonDropDownItem5.Label = "3";
            ribbonDropDownItem6.Label = "5";
            ribbonDropDownItem7.Label = "6";
            ribbonDropDownItem8.Label = "7";
            ribbonDropDownItem9.Label = "8";
            ribbonDropDownItem10.Label = "9";
            ribbonDropDownItem11.Label = "10";
            this.dropDownYear.Items.Add(ribbonDropDownItem1);
            this.dropDownYear.Items.Add(ribbonDropDownItem2);
            this.dropDownYear.Items.Add(ribbonDropDownItem3);
            this.dropDownYear.Items.Add(ribbonDropDownItem4);
            this.dropDownYear.Items.Add(ribbonDropDownItem5);
            this.dropDownYear.Items.Add(ribbonDropDownItem6);
            this.dropDownYear.Items.Add(ribbonDropDownItem7);
            this.dropDownYear.Items.Add(ribbonDropDownItem8);
            this.dropDownYear.Items.Add(ribbonDropDownItem9);
            this.dropDownYear.Items.Add(ribbonDropDownItem10);
            this.dropDownYear.Items.Add(ribbonDropDownItem11);
            this.dropDownYear.Label = "Set Year";
            this.dropDownYear.Name = "dropDownYear";
            this.dropDownYear.OfficeImageId = "TableInsertExcel";
            this.dropDownYear.ScreenTip = "Highlight a group of cells and choose a year ";
            this.dropDownYear.ShowImage = true;
            // 
            // button1
            // 
            this.button1.Label = "Auto-set Years";
            this.button1.Name = "button1";
            // 
            // grp_EnergySources
            // 
            this.grp_EnergySources.Items.Add(this.menu1);
            this.grp_EnergySources.Label = "Energy Sources";
            this.grp_EnergySources.Name = "grp_EnergySources";
            this.grp_EnergySources.Visible = false;
            // 
            // menu1
            // 
            this.menu1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Dynamic = true;
            this.menu1.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu1.Label = "menu1";
            this.menu1.Name = "menu1";
            this.menu1.OfficeImageId = "TableInsertExcel";
            this.menu1.ShowImage = true;
            // 
            // grp_Variables
            // 
            this.grp_Variables.Items.Add(this.menu2);
            this.grp_Variables.Label = "Variables";
            this.grp_Variables.Name = "grp_Variables";
            this.grp_Variables.Visible = false;
            // 
            // menu2
            // 
            this.menu2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Dynamic = true;
            this.menu2.ItemSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.menu2.Label = "menu2";
            this.menu2.Name = "menu2";
            this.menu2.OfficeImageId = "TableInsertExcel";
            this.menu2.ShowImage = true;
            // 
            // Wizard
            // 
            this.Wizard.Items.Add(this.btnWizard);
            this.Wizard.Items.Add(this.btnOutputWizard);
            this.Wizard.Label = "Wizard";
            this.Wizard.Name = "Wizard";
            // 
            // btnWizard
            // 
            this.btnWizard.Label = "EnPI Step-by-step Wizard";
            this.btnWizard.Name = "btnWizard";
            this.btnWizard.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnWizard_Click);
            // 
            // btnOutputWizard
            // 
            this.btnOutputWizard.Label = "EnPI Output Wizard";
            this.btnOutputWizard.Name = "btnOutputWizard";
            this.btnOutputWizard.Visible = false;
            this.btnOutputWizard.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnOutputWizard_Click);
            // 
            // UnitConversion
            // 
            this.UnitConversion.Items.Add(this.btnConvertUnits);
            this.UnitConversion.Label = "Unit Conversion";
            this.UnitConversion.Name = "UnitConversion";
            // 
            // btnConvertUnits
            // 
            this.btnConvertUnits.Label = "Convert Units";
            this.btnConvertUnits.Name = "btnConvertUnits";
            this.btnConvertUnits.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnConvertUnits_Click);
            // 
            // LabelReportingPeriod
            // 
            this.LabelReportingPeriod.Items.Add(this.btnReportingPeriod);
            this.LabelReportingPeriod.Label = "Label Reporting Period";
            this.LabelReportingPeriod.Name = "LabelReportingPeriod";
            // 
            // btnReportingPeriod
            // 
            this.btnReportingPeriod.Label = "Label Reporting Period";
            this.btnReportingPeriod.Name = "btnReportingPeriod";
            this.btnReportingPeriod.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnReportingPeriod_Click);
            // 
            // grp_ComputeEnPI_Actual
            // 
            this.grp_ComputeEnPI_Actual.Items.Add(this.btn_Actual);
            this.grp_ComputeEnPI_Actual.Label = "Compute EnPI - Actual";
            this.grp_ComputeEnPI_Actual.Name = "grp_ComputeEnPI_Actual";
            // 
            // btn_Actual
            // 
            this.btn_Actual.Label = "Use Actual Data";
            this.btn_Actual.Name = "btn_Actual";
            this.btn_Actual.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btn_Actual_Click);
            // 
            // grp_ComputeEnPI_Regression
            // 
            this.grp_ComputeEnPI_Regression.Items.Add(this.btn_Regression);
            this.grp_ComputeEnPI_Regression.Label = "Compute EnPI - Regression";
            this.grp_ComputeEnPI_Regression.Name = "grp_ComputeEnPI_Regression";
            // 
            // btn_Regression
            // 
            this.btn_Regression.Label = "Use Regression";
            this.btn_Regression.Name = "btn_Regression";
            this.btn_Regression.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btn_Backcast_Click);
            // 
            // Model
            // 
            this.Model.Items.Add(this.btnModelChange);
            this.Model.Label = "Model";
            this.Model.Name = "Model";
            // 
            // btnModelChange
            // 
            this.btnModelChange.Label = "Change Models";
            this.btnModelChange.Name = "btnModelChange";
            this.btnModelChange.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnModelChange_Click);
            // 
            // Rollup
            // 
            this.Rollup.Items.Add(this.btnRollup);
            this.Rollup.Label = "Roll Up";
            this.Rollup.Name = "Rollup";
            // 
            // btnRollup
            // 
            this.btnRollup.Label = "Corporate Roll Up";
            this.btnRollup.Name = "btnRollup";
            this.btnRollup.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnRollup_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btnAbout);
            this.group2.Label = "About";
            this.group2.Name = "group2";
            // 
            // btnAbout
            // 
            this.btnAbout.Label = "About EnPI";
            this.btnAbout.Name = "btnAbout";
            this.btnAbout.Click += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonControlEventArgs>(this.btnAbout_Click);
            // 
            // EnPIRibbon
            // 
            this.Name = "EnPIRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.tabEnPI);
            this.Load += new System.EventHandler<Microsoft.Office.Tools.Ribbon.RibbonUIEventArgs>(this.EnPIRibbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.tabEnPI.ResumeLayout(false);
            this.tabEnPI.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.grp_EnergySources.ResumeLayout(false);
            this.grp_EnergySources.PerformLayout();
            this.grp_Variables.ResumeLayout(false);
            this.grp_Variables.PerformLayout();
            this.Wizard.ResumeLayout(false);
            this.Wizard.PerformLayout();
            this.UnitConversion.ResumeLayout(false);
            this.UnitConversion.PerformLayout();
            this.LabelReportingPeriod.ResumeLayout(false);
            this.LabelReportingPeriod.PerformLayout();
            this.grp_ComputeEnPI_Actual.ResumeLayout(false);
            this.grp_ComputeEnPI_Actual.PerformLayout();
            this.grp_ComputeEnPI_Regression.ResumeLayout(false);
            this.grp_ComputeEnPI_Regression.PerformLayout();
            this.Model.ResumeLayout(false);
            this.Model.PerformLayout();
            this.Rollup.ResumeLayout(false);
            this.Rollup.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabEnPI;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDownYear;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_EnergySources;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_Variables;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_ComputeEnPI_Actual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Actual;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_Regression;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu1;
        internal Microsoft.Office.Tools.Ribbon.RibbonMenu menu2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Wizard;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnWizard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_ComputeEnPI_Regression;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup UnitConversion;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvertUnits;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Rollup;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRollup;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Model;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModelChange;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOutputWizard;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup LabelReportingPeriod;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReportingPeriod;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAbout;
    }

    partial class ThisRibbonCollection : Microsoft.Office.Tools.Ribbon.RibbonReadOnlyCollection
    {
        internal EnPIRibbon Ribbon1
        {
            get { return this.GetRibbon<EnPIRibbon>(); }
        }
    }
}
