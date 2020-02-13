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
    public partial class FuelUnitConversionControl : UserControl
    {

        public Excel.ListObject DataLO;
        private CheckedListBox check;
        private ComboBox combo;
        private double scfConFactor;
        private double galConFactor;
        private double lbConFactor;
        private double mmbtu;
        private double heatValue;
        private string unitDes;
        private bool fromWizard;


        public FuelUnitConversionControl(bool fromWizard, CheckedListBox c, ComboBox box)
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

        public void Open()
        {
            AddColumnNames();
            PopulateSourceType();
            PopulateUnits();
        }

        private void AddColumnNames()
        {
            if (DataLO != null)
            {
                foreach (Excel.ListColumn LC in DataLO.ListColumns)
                {
                    //Commented by Suman:
                    //TFS Ticket 70211 : Not sure why this check is made but as result of this comparison the code is throwing Index out of exception.
                    //if (LC != Utilities.ExcelHelpers.GetListColumn(DataLO, EnPIResources.yearColName))
                    //{
                        this.checkedListBox1.Items.Add(LC.Name);
                    //}
                }
            }

            foreach (string c in this.check.CheckedItems)
                this.checkedListBox1.SetItemChecked(check.Items.IndexOf(c),true);
            this.checkedListBox1.SelectedItem = this.check.SelectedItem;
        }

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

            this.comboBox3.SelectedItem = this.combo.SelectedItem;
        }

        private void PopulateFuelType(int t)
        {
            this.comboBox4.Items.Clear();
            switch (t)
            {
                case 1:
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_NATURAL_GAS);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BLAST_FURNACE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_COKE_OVEN);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LPG);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_PROPANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BUTANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_ISOBUTANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LANDFILL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_OIL_GASSES);
                    break;
                case 2:
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LIQ_PROPANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LIQ_BUTANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LIQ_ISOBUTANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_PENTANE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_ETHYLENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_PROPYLENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BUTENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_PENTENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BENZENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_TOLUENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_XYLENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_METHYL_ALCOHOL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_ETHYL_ALCOHOL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_1_FUEL_OIL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_2_FUEL_OIL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_4_FUEL_OIL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_5_FUEL_OIL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_6_FUEL_OIL_LOW);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_6_FUEL_OIL_HIGH);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_CRUDE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_GASOLINE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_KEROSENE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_GAS_OIL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_LNG);
                    break;
                case 3:
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_COAL);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_COKE);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_PEAT);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_WOOD);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BIOMASS);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_BLACK_LIQUOR);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_SCRAP_TIRES);
                    this.comboBox4.Items.Add(Constants.FUEL_TYPE_SULFUR);
                    break;
            }
        }

        private void PopulateUnits()
        {
            //Fuel specific units
            this.comboBox1.Items.Add(Constants.UNITS_SCF);
            this.comboBox1.Items.Add(Constants.UNITS_CCF);
            this.comboBox1.Items.Add(Constants.UNITS_MCF);
            this.comboBox1.Items.Add(Constants.UNITS_M3);
            this.comboBox1.Items.Add(Constants.UNITS_GALLON);
            this.comboBox1.Items.Add(Constants.UNITS_BBL);
            this.comboBox1.Items.Add(Constants.UNITS_LB);
            this.comboBox1.Items.Add(Constants.UNITS_SHORT_TONS);
            this.comboBox1.Items.Add(Constants.UNITS_LONG_TONS);
            this.comboBox1.Items.Add(Constants.UNITS_KG);
            this.comboBox1.Items.Add(Constants.UNITS_METRIC_TONS);
            //Standard units
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

        private string PopulateConversionFactor(string unitType, double mmbtu)
        {
            string conversionFactor = "";
            switch (unitType)
            {
                case Constants.UNITS_MMBTU:
                    conversionFactor = mmbtu.ToString();
                    break;
                case Constants.UNITS_GJ:
                    conversionFactor = (mmbtu * 1.05505588).ToString();
                    break;
            }
            return conversionFactor;
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!this.comboBox3.SelectedItem.Equals(Constants.SOURCE_PURCHASED_FUEL))
                Globals.ThisAddIn.LaunchUnitConversionControl(this.fromWizard,this.checkedListBox1, this.comboBox3);
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
        }

        private void disableControls()
        {
            //disables all controls except checkedListBox1
            this.comboBox1.Enabled = false;
            this.comboBox2.Enabled = false;
            this.comboBox3.Enabled = false;
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

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.comboBox4.Enabled = true;

            if (!this.comboBox1.SelectedItem.Equals(Constants.UNITS_OTHER))
            {

                if (this.comboBox1.SelectedItem.Equals(Constants.UNITS_SCF) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_CCF) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_MCF) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_M3))
                {
                    this.textBox1.Enabled = true;
                    if (this.comboBox4.SelectedIndex.Equals(-1))
                        this.btnRun.Enabled = false;
                    this.textBox1.Clear();
                    this.label7.Visible = true;
                    scfConFactor = ConvertToSCF();
                    if (this.comboBox4.Items.Count.Equals(0) || !this.comboBox4.Items[0].Equals(Constants.FUEL_TYPE_NATURAL_GAS))
                        PopulateFuelType(1);
                    this.comboBox4.Enabled = true;
                    unitDes = "SCF";

                    if (!this.comboBox2.SelectedIndex.Equals(-1))
                    {
                        this.label8.Visible = true;
                        this.label8.Text = "(" + this.comboBox2.SelectedItem.ToString() + "/" + unitDes + ")";
                    }
                    this.label4.Text = "Heating Value";
                }
                else
                    if (this.comboBox1.SelectedItem.Equals(Constants.UNITS_GALLON) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_BBL))
                    {
                        this.textBox1.Enabled = true;
                        if (this.comboBox4.SelectedIndex.Equals(-1))
                            this.btnRun.Enabled = false;
                        this.textBox1.Clear();
                        this.label7.Visible = true;
                        galConFactor = ConvertToGallon();
                        if (this.comboBox4.Items.Count.Equals(0) || !this.comboBox4.Items[0].Equals(Constants.FUEL_TYPE_LIQ_PROPANE))
                            PopulateFuelType(2);
                        this.comboBox4.Enabled = true;
                        unitDes = "GAL";
                        if (!this.comboBox2.SelectedIndex.Equals(-1))
                        {
                            this.label8.Visible = true;
                            this.label8.Text = "(" + this.comboBox2.SelectedItem.ToString() + "/" + unitDes + ")";
                        }
                        this.label4.Text = "Heating Value";
                    }
                    else
                        if (this.comboBox1.SelectedItem.Equals(Constants.UNITS_LB) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_SHORT_TONS) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_LONG_TONS) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_KG) || this.comboBox1.SelectedItem.Equals(Constants.UNITS_METRIC_TONS))
                        {
                            this.textBox1.Enabled = true;
                            if (this.comboBox4.SelectedIndex.Equals(-1))
                                this.btnRun.Enabled = false;
                            this.textBox1.Clear();
                            this.label7.Visible = true;
                            lbConFactor = ConvertToLB();
                            if (this.comboBox4.Items.Count.Equals(0) || !this.comboBox4.Items[0].Equals(Constants.FUEL_TYPE_COAL))
                                PopulateFuelType(3);
                            this.comboBox4.Enabled = true;
                            unitDes = "LB";
                            if (!this.comboBox2.SelectedIndex.Equals(-1))
                            {
                                this.label8.Visible = true;
                                this.label8.Text = "(" + this.comboBox2.SelectedItem.ToString() + "/" + unitDes + ")";
                            }
                            this.label4.Text = "Heating Value";
                        }
                        else
                        {
                            if (!this.comboBox2.SelectedIndex.Equals(-1))
                                this.textBox1.Text = PopulateConversionFactor();
                            this.comboBox4.Enabled = false;
                            this.comboBox4.Items.Clear();
                            this.textBox1.Enabled = false;
                            this.label8.Visible = false;
                            this.label4.Text = "Conversion Unit";
                        }
            }
            else
            {
                this.label8.Text = "";
                this.label4.Text = "Conversion Unit";
                this.textBox1.Enabled = true;
                this.textBox1.Text = Constants.UNITS_OTHER_VALUE;
                this.comboBox4.Enabled = false;
            }

            if (!this.comboBox2.SelectedIndex.Equals(-1) && !this.comboBox4.SelectedIndex.Equals(-1))
            {
                comboBox4_SelectedIndexChanged(this.comboBox4, null);

                this.btnRun.Enabled = true;
                this.textBox1.Text = PopulateConversionFactor(this.comboBox2.SelectedItem.ToString(), mmbtu);
            }

            if (!this.comboBox2.SelectedIndex.Equals(-1) && (this.textBox1.Text != null && this.textBox1.Text != ""))
            {
                if(this.comboBox4.Enabled == false || !this.comboBox4.SelectedIndex.Equals(-1))
                    this.btnRun.Enabled = true;
            }

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!this.comboBox1.SelectedIndex.Equals(-1) && this.comboBox4.Enabled == false)
            {
                this.label8.Visible = false;
                this.textBox1.Text = PopulateConversionFactor();
            }
            if (!this.comboBox1.SelectedIndex.Equals(-1) && (this.textBox1.Text != null && this.textBox1.Text != ""))
            {
                if (this.comboBox4.Enabled == false || !this.comboBox4.SelectedIndex.Equals(-1))
                    this.btnRun.Enabled = true;
                this.label8.Text = "(" + this.comboBox2.SelectedItem.ToString() + "/" + unitDes + ")";
                this.label8.Visible = true;
            }
            if (!this.comboBox1.SelectedIndex.Equals(-1) && !this.comboBox4.SelectedIndex.Equals(-1))
            {
                this.btnRun.Enabled = true;
                this.textBox1.Text = PopulateConversionFactor(this.comboBox2.SelectedItem.ToString(), mmbtu);
                this.label8.Text = "(" + this.comboBox2.SelectedItem.ToString() + "/" + unitDes + ")";
                this.label8.Visible = true;
            }
        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            //calculate heating value
            if(!this.comboBox1.SelectedIndex.Equals(-1))
                this.textBox1.Text = PopulateConversionFactor(this.comboBox1.SelectedItem.ToString(),1);
            if (!this.comboBox1.SelectedIndex.Equals(-1) && (this.textBox1.Text != null && this.textBox1.Text != ""))
            {
                this.btnRun.Enabled = true;
            }


            ComboBox c = (ComboBox)sender;

            switch (c.SelectedItem.ToString())
            {
                case Constants.FUEL_TYPE_NATURAL_GAS:
                    heatValue = Constants.FUEL_TYPE_NATURAL_GAS_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_BLAST_FURNACE:
                    heatValue = Constants.FUEL_TYPE_BLAST_FURNACE_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_COKE_OVEN:
                    heatValue = Constants.FUEL_TYPE_COKE_OVEN_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_LPG:
                    heatValue = Constants.FUEL_TYPE_LPG_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_PROPANE:
                    heatValue = Constants.FUEL_TYPE_PROPANE_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_BUTANE:
                    heatValue = Constants.FUEL_TYPE_BUTANE_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_ISOBUTANE:
                    heatValue = Constants.FUEL_TYPE_ISOBUTANE_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_LANDFILL:
                    heatValue = Constants.FUEL_TYPE_LANDFILL_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;
                case Constants.FUEL_TYPE_OIL_GASSES:
                    heatValue = Constants.FUEL_TYPE_OIL_GASSES_VALUE;
                    mmbtu = ConvertToMMBTU(scfConFactor,heatValue);
                    break;

                case Constants.FUEL_TYPE_LIQ_PROPANE:
                    heatValue = Constants.FUEL_TYPE_LIQ_PROPANE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_LIQ_BUTANE:
                    heatValue = Constants.FUEL_TYPE_LIQ_BUTANE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_LIQ_ISOBUTANE:
                    heatValue = Constants.FUEL_TYPE_LIQ_ISOBUTANE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_PENTANE:
                    heatValue = Constants.FUEL_TYPE_PENTANE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_ETHYLENE:
                    heatValue = Constants.FUEL_TYPE_ETHYLENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_PROPYLENE:
                    heatValue = Constants.FUEL_TYPE_PROPYLENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_BUTENE:
                    heatValue = Constants.FUEL_TYPE_BUTENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_PENTENE:
                    heatValue = Constants.FUEL_TYPE_PENTENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_BENZENE:
                    heatValue = Constants.FUEL_TYPE_BENZENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_TOLUENE:
                    heatValue = Constants.FUEL_TYPE_TOLUENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_XYLENE:
                    heatValue = Constants.FUEL_TYPE_XYLENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_METHYL_ALCOHOL:
                    heatValue = Constants.FUEL_TYPE_METHYL_ALCOHOL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_ETHYL_ALCOHOL:
                    heatValue = Constants.FUEL_TYPE_ETHYL_ALCOHOL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_1_FUEL_OIL:
                    heatValue = Constants.FUEL_TYPE_1_FUEL_OIL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_2_FUEL_OIL:
                    heatValue = Constants.FUEL_TYPE_2_FUEL_OIL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_4_FUEL_OIL:
                    heatValue = Constants.FUEL_TYPE_4_FUEL_OIL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_5_FUEL_OIL:
                    heatValue = Constants.FUEL_TYPE_5_FUEL_OIL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_6_FUEL_OIL_LOW:
                    heatValue = Constants.FUEL_TYPE_6_FUEL_OIL_LOW_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_6_FUEL_OIL_HIGH:
                    heatValue = Constants.FUEL_TYPE_6_FUEL_OIL_HIGH_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_CRUDE:
                    heatValue = Constants.FUEL_TYPE_CRUDE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_GASOLINE:
                    heatValue = Constants.FUEL_TYPE_GASOLINE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_KEROSENE:
                    heatValue = Constants.FUEL_TYPE_KEROSENE_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_GAS_OIL:
                    heatValue = Constants.FUEL_TYPE_GAS_OIL_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_LNG:
                    heatValue = Constants.FUEL_TYPE_LNG_VALUE;
                    mmbtu = ConvertToMMBTU(galConFactor, heatValue);
                    break;

                case Constants.FUEL_TYPE_COAL:
                    heatValue = Constants.FUEL_TYPE_COAL_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_COKE:
                    heatValue = Constants.FUEL_TYPE_COKE_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_PEAT:
                    heatValue = Constants.FUEL_TYPE_PEAT_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_WOOD:
                    heatValue = Constants.FUEL_TYPE_WOOD_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_BIOMASS:
                    heatValue = Constants.FUEL_TYPE_BIOMASS_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_BLACK_LIQUOR:
                    heatValue = Constants.FUEL_TYPE_BLACK_LIQUOR_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_SCRAP_TIRES:
                    heatValue = Constants.FUEL_TYPE_SCRAP_TIRES_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
                case Constants.FUEL_TYPE_SULFUR:
                    heatValue = Constants.FUEL_TYPE_SULFUR_VALUE;
                    mmbtu = ConvertToMMBTU(lbConFactor, heatValue);
                    break;
            }

            if (!this.comboBox2.SelectedIndex.Equals(-1))
            {
                this.btnRun.Enabled = true;
                this.textBox1.Text = PopulateConversionFactor(this.comboBox2.SelectedItem.ToString(), mmbtu);
            }
        }

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

                this.checkedListBox1.Items.Clear();
                AddColumnNames();
            }
            catch
            {
                //raise error
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(5);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(7);
        }
    }

}
