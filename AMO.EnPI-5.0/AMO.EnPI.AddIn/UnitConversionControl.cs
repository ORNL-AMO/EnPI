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
    public partial class UnitConversionControl : UserControl
    {
        public UnitConversionControl(bool fromWizard)
        {
            this.fromWizard = fromWizard;

            InitializeComponent();

            DataLO = ((Excel.Range)Globals.ThisAddIn.Application.Selection).ListObject;

            if (DataLO == null) DataLO = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[1];

            this.btnBack.Visible = fromWizard;
            this.btnClose.Visible = fromWizard;
            this.btnNext.Visible = fromWizard;
        }

        public UnitConversionControl(bool fromWizard, CheckedListBox c, ComboBox box)
        {
            check = c;
            combo = box;
            this.fromWizard = fromWizard;

            InitializeComponent();

            DataLO = ((Excel.Range)Globals.ThisAddIn.Application.Selection).ListObject;

            if (DataLO == null) DataLO = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[1];

            this.btnBack.Visible = fromWizard;
            this.btnClose.Visible = fromWizard;
            this.btnNext.Visible = fromWizard;
        }

        #region Local Variables

        public Excel.ListObject DataLO;
        public double scfConFactor;
        public double galConFactor;
        public double lbConFactor;
        public double mmbtu;
        public double heatValue;
        private ComboBox combo;
        private CheckedListBox check;
        private bool fromWizard;

        #endregion

        private void UnitConversionControl_Load(object sender, EventArgs e)
        {

        }

        public void Open()
        {
            AddColumnNames();
            PopulateSourceType();

        }

        private void AddColumnNames()
        {
            if (DataLO != null)
            {
                foreach (Excel.ListColumn LC in DataLO.ListColumns)
                {
                    if (LC != Utilities.ExcelHelpers.GetListColumn(DataLO, "Year"))
                    {
                        this.checkedListBox1.Items.Add(LC.Name);
                    }
                }
            }

            if (check != null)
            {
                foreach (string c in this.check.CheckedItems)
                    this.checkedListBox1.SetItemChecked(check.Items.IndexOf(c), true);
                this.checkedListBox1.SelectedItem = this.check.SelectedItem;
            }
        }

        #region Conversion Methods

        private double ConvertToSCF()
        {
            double scf = 0;
            switch (this.comboBox1.SelectedItem.ToString())
            {
                case Constants.UNITS_SCF:
                    scf = Constants.UNITS_SCF_VALUE;
                    break;
                case Constants.UNITS_CCF:
                    scf = Constants.UNITS_CCF_VALUE;
                    break;
                case Constants.UNITS_MCF:
                    scf = Constants.UNITS_MCF_VALUE;
                    break;
                case Constants.UNITS_M3:
                    scf = Constants.UNITS_M3_VALUE;
                    break;
            }
            return scf;
        }

        private double ConvertToGallon()
        {
            double gal = 0;
            switch (this.comboBox1.SelectedItem.ToString())
            {
                case Constants.UNITS_GALLON:
                    gal = Constants.UNITS_GALLON_VALUE;
                    break;
                case Constants.UNITS_BBL:
                    gal = Constants.UNITS_BBL_VALUE;
                    break;
            }
            return gal;
        }

        private double ConvertToLB()
        {
            double lb = 0;
            switch (this.comboBox1.SelectedItem.ToString())
            {
                case Constants.UNITS_LB:
                    lb = Constants.UNITS_LB_VALUE;
                    break;
                case Constants.UNITS_SHORT_TONS:
                    lb = Constants.UNITS_SHORT_TONS_VALUE;
                    break;
                case Constants.UNITS_LONG_TONS:
                    lb = Constants.UNITS_LONG_TONS_VALUE;
                    break;
                case Constants.UNITS_KG:
                    lb = Constants.UNITS_KG_VALUE;
                    break;
                case Constants.UNITS_METRIC_TONS:
                    lb = Constants.UNITS_METRIC_TONS_VALUE;
                    break;
            }
            return lb;
        }

        private double ConvertToMMBTU(double type, double heatingValue)
        {
            double output = type * heatingValue;
            return output;
        }

        #endregion

        #region Population Methods

        private void PopulateSourceType()
        {
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_ELECTRICITY);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_FUEL);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_STEAM);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_CHILLED_WATER_ENGINE);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC);
            this.comboBox3.Items.Add(Constants.SOURCE_PURCHASED_COMPRESSED_AIR);
            this.comboBox3.Items.Add(Constants.SOURCE_ELECTRICITY_SOLD);
            this.comboBox3.Items.Add(Constants.SOURCE_STEAM_SOLD);

            if(combo != null)
                this.comboBox3.SelectedItem = this.combo.SelectedItem;
        }

        private void PopulateSiteToSource(ComboBox sender)
        {
            switch (sender.SelectedItem.ToString())
            {
                case Constants.SOURCE_PURCHASED_ELECTRICITY:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_ELECTRICITY_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_FUEL:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_FUEL_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_STEAM:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_STEAM_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ENGINE:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_CHILLED_WATER_ENGINE_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC_VALUE;
                    break;
                case Constants.SOURCE_PURCHASED_COMPRESSED_AIR:
                    this.textBox2.Text = Constants.SOURCE_PURCHASED_COMPRESSED_AIR_VALUE;
                    break;
                case Constants.SOURCE_ELECTRICITY_SOLD:
                    this.textBox2.Text = Constants.SOURCE_ELECTRICITY_SOLD_VALUE;
                    break;
                case Constants.SOURCE_STEAM_SOLD:
                    this.textBox2.Text = Constants.SOURCE_STEAM_SOLD_VALUE;
                    break;
            }
        }

        private void PopulateAdditionalUnits(ComboBox sender)
        {
            switch (sender.SelectedItem.ToString())
            {
                case Constants.SOURCE_PURCHASED_ELECTRICITY:
                    
                    break;
                case Constants.SOURCE_PURCHASED_STEAM:
                    
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION:
                    this.comboBox1.Items.Add(Constants.UNITS_TON_HOUR);
                    this.comboBox1.Items.Add(Constants.UNITS_GAL_DEG_F);
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ENGINE:
                    this.comboBox1.Items.Add(Constants.UNITS_TON_HOUR);
                    this.comboBox1.Items.Add(Constants.UNITS_GAL_DEG_F);
                    break;
                case Constants.SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC:
                    this.comboBox1.Items.Add(Constants.UNITS_TON_HOUR);
                    this.comboBox1.Items.Add(Constants.UNITS_GAL_DEG_F);
                    break;
                case Constants.SOURCE_PURCHASED_COMPRESSED_AIR:
                    this.comboBox1.Items.Add(Constants.UNITS_FT3);
                    break;
                case Constants.SOURCE_ELECTRICITY_SOLD:
                    
                    break;
                case Constants.SOURCE_STEAM_SOLD:
                    
                    break;
            }
        }

        private void PopulateUnits()
        {
            this.comboBox1.Items.Add(Constants.UNITS_MMBTU);
            this.comboBox1.Items.Add(Constants.UNITS_KWH);
            this.comboBox1.Items.Add(Constants.UNITS_MWH);
            this.comboBox1.Items.Add(Constants.UNITS_GWH);
            this.comboBox1.Items.Add(Constants.UNITS_KJ);
            this.comboBox1.Items.Add(Constants.UNITS_MJ);
            this.comboBox1.Items.Add(Constants.UNITS_GJ);
            this.comboBox1.Items.Add(Constants.UNITS_TJ);
            this.comboBox1.Items.Add(Constants.UNITS_THERMS);
            this.comboBox1.Items.Add(Constants.UNITS_DTH);
            this.comboBox1.Items.Add(Constants.UNITS_KCAL);
            this.comboBox1.Items.Add(Constants.UNITS_GCAL);
            this.comboBox1.Items.Add(Constants.UNITS_OTHER);

            this.comboBox2.Items.Add(Constants.UNITS_MMBTU);
            this.comboBox2.Items.Add(Constants.UNITS_GJ);
            //added until users can select units other than MMBTU to convert to.
            this.comboBox2.SelectedIndex = 0;
        }

        private string PopulateConversionFactor()
        {
            string conversionFactor = "";
            switch (this.comboBox2.SelectedItem.ToString())
            {
                case Constants.UNITS_MMBTU:
                    switch (this.comboBox1.SelectedItem.ToString())
                    {
                        case Constants.UNITS_MMBTU:
                            conversionFactor = Constants.UNITS_MMBTU_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_KWH:
                            conversionFactor = Constants.UNITS_KWH_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_MWH:
                            conversionFactor = Constants.UNITS_MWH_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_GWH:
                            conversionFactor = Constants.UNITS_GWH_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_KJ:
                            conversionFactor = Constants.UNITS_KJ_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_MJ:
                            conversionFactor = Constants.UNITS_MJ_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_GJ:
                            conversionFactor = Constants.UNITS_GJ_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_TJ:
                            conversionFactor = Constants.UNITS_TJ_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_THERMS:
                            conversionFactor = Constants.UNITS_THERMS_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_DTH:
                            conversionFactor = Constants.UNITS_DTH_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_KCAL:
                            conversionFactor = Constants.UNITS_KCAL_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_GCAL:
                            conversionFactor = Constants.UNITS_GCAL_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_TON_HOUR:
                            conversionFactor = Constants.UNITS_TON_HOUR_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_GAL_DEG_F:
                            conversionFactor = Constants.UNITS_GAL_DEG_F_VALUE_MMBTU;
                            break;
                        case Constants.UNITS_FT3:
                            conversionFactor = Constants.UNITS_FT3_VALUE_MMBTU;
                            break;
                    }
                    break;

                case Constants.UNITS_GJ:
                    switch (this.comboBox1.SelectedItem.ToString())
                    {
                        case Constants.UNITS_MMBTU:
                            conversionFactor = Constants.UNITS_MMBTU_VALUE_GJ;
                            break;
                        case Constants.UNITS_KWH:
                            conversionFactor = Constants.UNITS_KWH_VALUE_GJ;
                            break;
                        case Constants.UNITS_MWH:
                            conversionFactor = Constants.UNITS_MWH_VALUE_GJ;
                            break;
                        case Constants.UNITS_GWH:
                            conversionFactor = Constants.UNITS_GWH_VALUE_GJ;
                            break;
                        case Constants.UNITS_KJ:
                            conversionFactor = Constants.UNITS_KJ_VALUE_GJ;
                            break;
                        case Constants.UNITS_MJ:
                            conversionFactor = Constants.UNITS_MJ_VALUE_GJ;
                            break;
                        case Constants.UNITS_GJ:
                            conversionFactor = Constants.UNITS_GJ_VALUE_GJ;
                            break;
                        case Constants.UNITS_TJ:
                            conversionFactor = Constants.UNITS_TJ_VALUE_GJ;
                            break;
                        case Constants.UNITS_THERMS:
                            conversionFactor = Constants.UNITS_THERMS_VALUE_GJ;
                            break;
                        case Constants.UNITS_DTH:
                            conversionFactor = Constants.UNITS_DTH_VALUE_GJ;
                            break;
                        case Constants.UNITS_KCAL:
                            conversionFactor = Constants.UNITS_KCAL_VALUE_GJ;
                            break;
                        case Constants.UNITS_GCAL:
                            conversionFactor = Constants.UNITS_GCAL_VALUE_GJ;
                            break;
                        case Constants.UNITS_TON_HOUR:
                            conversionFactor = Constants.UNITS_TON_HOUR_VALUE_GJ;
                            break;
                        case Constants.UNITS_GAL_DEG_F:
                            conversionFactor = Constants.UNITS_GAL_DEG_F_VALUE_GJ;
                            break;
                        case Constants.UNITS_FT3:
                            conversionFactor = Constants.UNITS_FT3_VALUE_GJ;
                            break;
                    }
                    break;
            }
            return conversionFactor;
        }

        #endregion

        #region Event Handlers

        private void btnRun_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
                Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);

                int orgCount = thisList.ListColumns.Count;

                foreach (string t in this.checkedListBox1.CheckedItems)
                {
                    int lcCount = thisList.ListColumns.Count;
                    int position = this.checkedListBox1.Items.IndexOf(t) + (lcCount - orgCount) + 2;
                    string colName = t + "(" + this.comboBox2.SelectedItem.ToString() + ")";
                    string formula = "=" + ExcelHelpers.CreateValidFormulaName(t) + "*" + this.textBox1.Text + "*" + this.textBox2.Text;
                    string stylename = "Comma";

                    Excel.ListColumn newColumn = thisList.ListColumns.Add(position);
                    newColumn.Name = colName;
                    newColumn.DataBodyRange.Formula = formula;
                    newColumn.DataBodyRange.Style = stylename;

                    //set old column property to not show up in regression/actual options
                }

                //clear and reset columns in pick-list
                this.checkedListBox1.Items.Clear();
                AddColumnNames();
            }
            catch
            {
                //raise error
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.textBox1.Enabled = false;

            if (this.comboBox1.SelectedItem.Equals(Constants.UNITS_OTHER))
            {
                this.textBox1.Enabled = true;
                this.textBox1.Text = Constants.UNITS_OTHER_VALUE;
            }
                
            if (!this.comboBox2.SelectedIndex.Equals(-1))
            {
                this.btnRun.Enabled = true;
                if(!this.comboBox1.SelectedItem.Equals(Constants.UNITS_OTHER))
                    this.textBox1.Text = PopulateConversionFactor();
            }
            
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (!this.comboBox1.SelectedIndex.Equals(-1))
            {
                this.btnRun.Enabled = true;
                this.textBox1.Text = PopulateConversionFactor();
            }
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

            this.comboBox1.Items.Clear();
            this.comboBox2.Items.Clear();

            if (this.comboBox3.SelectedItem.Equals(Constants.SOURCE_PURCHASED_FUEL))
                Globals.ThisAddIn.LaunchFuelUnitConversionControl(this.fromWizard,this.checkedListBox1, this.comboBox3);

            ComboBox s = (ComboBox)sender;
            PopulateSiteToSource(s);
            PopulateAdditionalUnits(s);

            if (this.comboBox1.Items.Count < 1 || this.comboBox2.Items.Count < 1)
                PopulateUnits();
            this.comboBox1.Enabled = true;
            this.comboBox2.Enabled = true;
            this.textBox2.Enabled = true;
        }


        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            if (e.NewValue.Equals(System.Windows.Forms.CheckState.Checked))
            {
                //check if controls are already populated with answers, if so enable contorls
                if (this.comboBox3.SelectedIndex > -1)
                    if (this.comboBox2.SelectedIndex > -1 && this.comboBox1.SelectedIndex > -1)
                        enableControls();
                    else
                    {
                        this.comboBox3.Enabled = true;
                        this.comboBox2.Enabled = true;
                        this.comboBox1.Enabled = true;
                        this.textBox2.Enabled = true;
                    }
                else
                    this.comboBox3.Enabled = true;
            }
            else
            {
                if (this.checkedListBox1.CheckedItems.Count.Equals(1))
                {
                    disableControls();
                }
            }
        }

        #endregion

        private void disableControls()
        {
            //disables all controls except checkedListBox1
            this.comboBox1.Enabled = false;
            this.comboBox2.Enabled = false;
            this.comboBox3.Enabled = false;
            this.textBox1.Enabled = false;
            this.textBox2.Enabled = false;
            this.btnRun.Enabled = false;
        }

        private void enableControls()
        {
            //enables all controls except checkedListBox1 and textBox1
            this.comboBox1.Enabled = true;
            this.comboBox2.Enabled = true;
            this.comboBox3.Enabled = true;
            this.textBox2.Enabled = true;
            this.btnRun.Enabled = true;
        }

        private void btnFuel_Click(object sender, EventArgs e)
        {
            //Globals.ThisAddIn.LaunchFuelUnitConversionControl(this.checkedListBox1);
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(5);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(7);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

    }
}
