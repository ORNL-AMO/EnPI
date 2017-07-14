using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
//using Microsoft.Office.Tools.Ribbon;
using AMO.EnPI.AddIn.Utilities;
using System.Windows.Forms;

namespace AMO.EnPI.AddIn
{
    public partial class EnPIRibbon : RibbonBase
    {
        public static System.Resources.ResourceManager rsc = 
            new System.Resources.ResourceManager(
            "AMO.EnPI.AddIn.EnPIResources", System.Reflection.Assembly.GetExecutingAssembly());

        public EnPIRibbon() : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }


        private void EnPIRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            this.dropDownYear.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDownYear_SelectionChanged);

            EnPIRibbon_MenuSetup();
        }


        private void EnPIRibbon_MenuSetup()
        {            // add energy source buttons
            this.menu1.Items.Clear();
            this.menu1.Label = rsc.GetString("lbl_AddEnergySources");
            foreach (Constants.EnergySourceTypes typ in System.Enum.GetValues(typeof(Constants.EnergySourceTypes)))
            {
                RibbonButton newButton = this.Factory.CreateRibbonButton();
                newButton.Name = "btn_" + typ.ToString();
                string lbl = (rsc.GetString(typ.ToString() + "_btntext")) ?? typ.ToString();
                newButton.Label = lbl;
                this.menu1.Items.Add(newButton);
                newButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAdd_Click);
            }


            // add independent variable buttons
            this.menu2.Items.Clear();
            this.menu2.Label = rsc.GetString("lbl_AddVariables");
            foreach (Constants.VariableTypes typ in System.Enum.GetValues(typeof(Constants.VariableTypes)))
            {
                RibbonButton newButton = this.Factory.CreateRibbonButton();
                newButton.Name = "btn_" + typ.ToString();
                string lbl = (rsc.GetString(typ.ToString() + "_btntext")) ?? typ.ToString();
                newButton.Label = lbl;

                this.menu2.Items.Add(newButton);
                newButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.buttonAdd_Click);
            }

        }
        
        private void autoyear_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);
            ExcelHelpers.AutoSetYear(thisList, 12);
        }
        private void dropDownYear_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Excel.Range thisRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);
            ExcelHelpers.SetYear(thisList, thisRange, dropDownYear.SelectedItemIndex - 1);
        }

        private void buttonAdd_Click(object sender, RibbonControlEventArgs e)
        {
            add_ListColumn(sender, e, ((RibbonButton)sender).Label);
        }
        private void add_ListColumn(object sender, RibbonControlEventArgs e, string ColName)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);

            ExcelHelpers.AddListColumn(thisList, ColName, 0);
        }


        #region //plot EnPI
        private void runPlot(Constants.EnPITypes method, params Excel.Range[] rngSelected)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Excel.ListObject thisList;
            Excel.Range year1;
            //Excel.Range year2;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            var defRange = System.Type.Missing;
            // prompt user for years to use
            if (rngSelected != null)
                defRange = rngSelected[0].get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
            try
            {
                year1 = prompt_YearRange(rsc.GetString("selectYear1"), rsc.GetString("selectYear"), defRange);
                //year2 = prompt_YearRange(rsc.GetString("selectYear2"), rsc.GetString("selectYear"), System.Type.Missing);
                thisList = year1.ListObject;
            }
            catch (InvalidCastException e)
            {
                return;
            }

            if (method == Constants.EnPITypes.Actual)
                Globals.ThisAddIn.actualEnPI(thisList);// thisSheet,
        }
        #endregion

        public static bool run_TableWizard(object sender)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            if (Globals.ThisAddIn.Application.Dialogs[Excel.XlBuiltInDialog.xlDialogCreateList].Show(System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing))
                return true;
            else
                return false;
        }

        private Excel.Range prompt_YearRange(string prompt, string title, object defRange)
        {
            return (Excel.Range)Globals.ThisAddIn.Application.InputBox(
                prompt, 
                title,
                defRange, 
                System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing,
                8); // 8 = range
        }

        private void btn_Actual_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Globals.ThisAddIn.fromWizard = false;

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchRegressionControl(Constants.EnPITypes.Actual);

            //under dev pop-up
            //Messages m = new Messages();
            //m.Show();
        }

        private void btn_Forecast_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Excel.Range thisRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            runPlot(Constants.EnPITypes.Forecast, thisRange);
        }

        private void btn_Backcast_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Globals.ThisAddIn.fromWizard = false;

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchRegressionControl(Constants.EnPITypes.Backcast);
        }

        private void btn_Chaining_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Range thisRange = Globals.ThisAddIn.Application.Selection as Excel.Range;
            //runPlot(Analytics.EnPITypes.Backcast, thisRange);
        }

        private void btnWizard_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            if (!Globals.ThisAddIn.wizardPane.Visible)
            {
                Globals.ThisAddIn.showWizard();
                Globals.ThisAddIn.LaunchWizardControl();
            }
            else
                Globals.ThisAddIn.hideWizard();
        }

        private void btnConvertUnits_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchUnitConversionControl(false);
        }

        private void btnRollup_Click(object sender, RibbonControlEventArgs e)
        {
            //RollupSheet rollup = new RollupSheet();
            //rollup.Initialize();
            Globals.ThisAddIn.CheckforVaildWizard();
            Globals.ThisAddIn.LaunchRollupControl(false);
        }

        public void btnModelChange_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchChangeModelControl(false);
        }

        private void btnOutputWizard_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            if (!Globals.ThisAddIn.wizardPane.Visible)
            {
                Globals.ThisAddIn.showWizard();
                Globals.ThisAddIn.LaunchWizardControl(8);
            }
            else
                Globals.ThisAddIn.hideWizard();
        }

        private void btnReportingPeriod_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CheckforVaildWizard();
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchReportingPeriodControl(false);
        }

        private void btnAbout_Click(object sender, RibbonControlEventArgs e)
        {
            //this box needs to be updated with every build since it was deemed unnecessary use of time to dynamically get the version number from the setup project
            MessageBox.Show("Version number: 5.0.0009\r\n\r\nBuild date: 11/03/2016 \r\n\r\nDescription: EnPI is a regression analysis based tool developed by the U.S. Department of Energy to help plant and corporate managers establish a normalized baseline of energy consumption, track annual progress of intensity improvements, energy savings, Superior Energy Performance (SEP) EnPIs, and other EnPIs that account for variations due to weather, production, and other variables.");
        }

     

    }
}
