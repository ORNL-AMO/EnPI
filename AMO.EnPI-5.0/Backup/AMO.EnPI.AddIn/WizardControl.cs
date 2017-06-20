using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;

namespace AMO.EnPI.AddIn
{
    public partial class WizardControl : UserControl
    {
        #region Loacl Variables 
        private int cStep = 1;
        private Excel.Range year1;
        #endregion

        public WizardControl()
        {
            InitializeComponent();
            showStep(cStep);
        }

        public WizardControl(int step)
        {
            cStep = step;
            InitializeComponent();
            showStep(cStep);
        }

        #region Button Events

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

        private void btnActual_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.fromWizard = true;

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchRegressionControl(Constants.EnPITypes.Actual);
        }

        private void btnRegression_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.fromWizard = true;

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //  check for list object; if one doesn't exist, run the table wizard          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Globals.ThisAddIn.LaunchRegressionControl(Constants.EnPITypes.Backcast);
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
           
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            if (cStep < 10)
            {
                cStep++;
                showStep(cStep);
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            if (cStep > 1)
            {
                cStep--;
                showStep(cStep);
            }
        }

        private void btnAddSource_Click(object sender, EventArgs e)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;
            if (!cbEnergy.SelectedIndex.Equals(-1))
                add_ListColumn(sender, e, cbEnergy.SelectedItem.ToString(), "");
        }

        private void btnIndVar_Click(object sender, EventArgs e)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;
            if (!cbIndVar.SelectedIndex.Equals(-1))
                add_ListColumn(sender, e, cbIndVar.SelectedItem.ToString(), "");
        }

        private void btnYear_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchReportingPeriodControl(true);
        }

        private void btnModelYear_Click(object sender, EventArgs e)
        {

        }

        private void btnConvertUnits_Click(object sender, EventArgs e)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;

            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);

            Globals.ThisAddIn.LaunchUnitConversionControl(true);
        }

        private void btnRollup_Click(object sender, EventArgs e)
        {
            btnNext.Text = "Next";
            Globals.ThisAddIn.LaunchRollupControl(true);
        }
        

        private void btnHasData_Click(object sender, EventArgs e)
        {

            lblTop.Text = "Step 1.1: Format Data as an Excel Table ";
            //lblQuestion.Text = "Once all your data is entered in the sheet, it must be formatted as an Excel table. If your data is not already formatted as an Excel table, select “Format data as an Excel table” below. Only one row is allowed for the header. Fuel types and units must be listed within the same cell. ";//EnPIResources.WizzardStep3Question1;
            //Modified by Suman TFS Ticket:68840
            lblQuestion.Text = "Once all your data is entered in the sheet, it must be formatted as an Excel table. If your data is not already formatted as an Excel table, first select a cell in the middle of the table and then select “Format data as an Excel table” below. Only one row is allowed for the header. Fuel types and units must be listed within the same cell.";
            lblQuestion2.Text = "";
            btnFormatTable.Visible = true;
            btnSpecialBack.Visible = true;
            btnSpecialNext.Visible = true;
            btnNext.Visible = false;
            btnBack.Visible = false;
            btnHasData.Visible = false;
            btnAddData.Visible = false;
        }

        private void btnAddData_Click(object sender, EventArgs e)
        {
            btnNext_Click(sender, e);
        }

        private void btnFormatTable_Click(object sender, EventArgs e)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;
        }

        private void btnSpecialBack_Click(object sender, EventArgs e)
        {
            showStep(2);
            cStep = 2;
        }

        private void btnSpecialNext_Click(object sender, EventArgs e)
        {
            showStep(5);
            cStep = 5;
        }

        private void btnChangeModel_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchChangeModelControl(true);
        }

        #endregion

        public void showStep(int currentStep)
        {
            

            switch (currentStep)
            {
                //For step 1 -- Welcome Text
                case 1:
                    btnNext.Visible = true;
                    lblTop.Text = EnPIResources.WizardStep1Title;
                    lblQuestion.Text = EnPIResources.WizzardStep1Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep1Question2;
                    btnBack.Visible = false;
                    cbEnergy.Visible = false;
                    btnAddData.Visible = false;
                    btnHasData.Visible = false;
                    btnChangeModel.Visible = false;
                    btnFormatTable.Visible = false;
                    break;
                case 2:
                    lblTop.Text = EnPIResources.WizardStep2Title;
                    lblQuestion.Text = EnPIResources.WizzardStep2Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep2Question2;
                    btnBack.Visible = true;
                    btnAddData.Visible = true;
                    btnHasData.Visible = true;
                    btnNext.Visible = false;
                    btnChangeModel.Visible = false;
                    cbEnergy.Visible = false;
                    btnAddEnergy.Visible = false;
                    btnFormatTable.Visible = false;
                    btnSpecialBack.Visible = false;
                    btnSpecialNext.Visible = false;
                    break;
                //For step 3 -- Add Energy Source
                case 3:
                    lblTop.Text = EnPIResources.WizardStep3Title;
                    lblQuestion.Text = EnPIResources.WizzardStep3Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep3Question2;
                    btnChangeModel.Visible = false;
                    btnNext.Visible = true;
                    btnAddEnergy.Visible = true;
                    cbEnergy.Visible = true;
                    btnAddData.Visible = false;
                    btnHasData.Visible = false;
                    cbIndVar.Visible = false;
                    btnIndVar.Visible = false;
                    btnFormatTable.Visible = false;
                    btnSpecialBack.Visible = false;
                    btnSpecialNext.Visible = false;
                    break;
                //For step 4 -- Add Independent Vars
                case 4:
                    lblTop.Text = EnPIResources.WizardStep4Title;
                    lblQuestion.Text = EnPIResources.WizzardStep4Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep4Question2;
                    btnIndVar.Visible = true;
                    btnNext.Visible = true;
                    btnBack.Visible = true;
                    btnChangeModel.Visible = false;
                    btnAddEnergy.Visible = false;
                    cbEnergy.Visible = false;
                    cbIndVar.Visible = true;
                    btnYear.Visible = false;
                    btnSpecialBack.Visible = false;
                    btnSpecialNext.Visible = false;
                    break;
                //For step 5 -- Add Year Column
                case 5:
                    this.Refresh();
                    lblTop.Text = EnPIResources.WizardStep5Title;
                    lblQuestion.Text = EnPIResources.WizzardStep5Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep5Question2;
                    lblQuestion2.Visible = false;
                    btnRegression.Visible = false;
                    btnActualData.Visible = false;
                    btnIndVar.Visible = false;
                    cbIndVar.Visible = false;
                    cbIndVar.Visible = false;
                    btnConvertUnits.Visible = false;
                    btnChangeModel.Visible = false;
                    btnYear.Visible = true;
                    btnNext.Visible = true;
                    btnBack.Visible = true;
                    btnSpecialBack.Visible = false;
                    btnSpecialNext.Visible = false;
                    btnFormatTable.Visible = false;
                    break;
                //For step 6 -- Unit Conversion
                case 6:
                    lblTop.Text = EnPIResources.WizardStep7Title;
                    lblQuestion.Text = EnPIResources.WizzardStep7Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep7Question2;
                    btnConvertUnits.Visible = true;
                    btnNext.Visible = true;
                    btnIndVar.Visible = false;
                    btnAnalyze.Visible = false;
                    btnYear.Visible = false;
                    btnActualData.Visible = false;
                    btnRegression.Visible = false;
                    btnChangeModel.Visible = false;
                    btnSpecialNext.Visible = false;
                    btnSpecialBack.Visible = false;
                    break;
                //For step 7 -- Regression/Use Actual
                case 7:
                    lblTop.Text = EnPIResources.WizardStep6Title;
                    lblQuestion.Text = EnPIResources.WizzardStep6Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep6Question2;
                    btnActualData.Visible = true;
                    btnRegression.Visible = true;
                    btnConvertUnits.Visible = false;
                    btnNext.Visible = false;
                    btnIndVar.Visible = false;
                    btnAnalyze.Visible = false;
                    btnYear.Visible = false;
                    btnChangeModel.Visible = false;
                    break;
                //For step 8 -- Output wizard - Change model
                case 8:
                    Globals.Ribbons.Ribbon1.btnOutputWizard.Visible = true;
                    lblTop.Text = EnPIResources.WizardStep8Title;
                    lblQuestion.Text = EnPIResources.WizzardStep8Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep8Question2;
                    btnChangeModel.Visible = true;
                    btnNext.Visible = true;
                    btnNext.Text = "Next";
                    btnRollup.Visible = false;
                    btnBack.Visible = false;
                    btnActualData.Visible = false;
                    btnRegression.Visible = false;
                    break;
                 //For step 9 -- Output wizard - Rollup
                case 9:
                    lblTop.Text = EnPIResources.WizardStep9Title;
                    lblQuestion.Text = EnPIResources.WizzardStep9Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep9Question2;
                    btnRollup.Visible = true;
                    btnNext.Visible = false;
                    btnNext.Text = "Skip";
                    btnBack.Visible = true;
                    btnChangeModel.Visible = false;
                    break;
                case 10:
                    lblTop.Text = EnPIResources.WizardStep10Title;
                    lblQuestion.Text = EnPIResources.WizzardStep10Question1;
                    lblQuestion2.Text = EnPIResources.WizzardStep10Question2;
                    btnRollup.Visible = false;
                    btnNext.Visible = false;
                    btnBack.Visible = true;
                    break;
            }
        }

        #region Modified Ribbion Methods
        private void selectYear(params Excel.Range[] rngSelected)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            //  check for list object          
            if (thisSheet.ListObjects.Count == 0)
                if (!run_TableWizard(thisSheet)) return;

            var defRange = System.Type.Missing;
            // compute EnPI with actual data
            if (rngSelected != null)
                defRange = rngSelected[0].get_Address(System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

            try
            {
                year1 = (Excel.Range)Globals.ThisAddIn.Application.InputBox(EnPIResources.selectYear1, EnPIResources.selectYear
                     , defRange, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing
                     , 8); // 8 = range
            }
            catch
            {
                return;
            }
        }

        private void run(Constants.EnPITypes method)
        {

        }

        private bool run_TableWizard(object sender)
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

        private void add_ListColumn(object sender, EventArgs e, string ColName, string ColType)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);

            ExcelHelpers.AddListColumn(thisList, ColName, 0);

        }
        #endregion

        
    }
}
