using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using AMO.EnPI.AddIn.Utilities;
using System.Text.RegularExpressions;

namespace AMO.EnPI.AddIn
{
    public partial class CO2EmissionControl : UserControl
    {
        #region Dynamic Variable Constants
        private const string STR_LBL_PREFIX = "lbl_";
        private const string STR_LBL_ENERGYSOURCE_PREFIX = "lbl_EnergySource_";
        private const string STR_CMB_ENERGYSOURCE_PREFIX = "cmb_EnergySource_";
        private const string STR_LBL_FUELTYPE_PREFIX = "lbl_FuelType_";
        private const string STR_CMB_FUELTYPE_PREFIX = "cmb_FuelType_";
        private const string STR_LBL_CO2EMISSION_PREFIX = "lbl_CO2Emission_";
        private const string STR_TXT_CO2EMISSION_PREFIX = "txt_CO2Emission_";
        #endregion

        #region Variables
        RegressionControl parentControl;
        CheckedListBox parentCheckListBox;
        ControlCollection parentControls;
        //bool isEnergyCostChecked;
        #endregion

        #region Constructors
        
        public CO2EmissionControl(RegressionControl parentControl)
        {
            InitializeComponent();
            //isEnergyCostChecked = blnIsEnergyCostChecked;
            this.parentControl = parentControl;

            if (Globals.ThisAddIn.fromWizard)
            {
                //if (isEnergyCostChecked == false)
                //{

                //    this.label1.Text = "Step 6: CO2 Avoided Emission Data";
                //}
                //else
                //{
                    this.label1.Text = "Step 7: CO2 Avoided Emission Data";
                //}
            }
            else
            {
                this.label1.Text = "CO2 Avoided Emission Data";
            }
        }
        #endregion


        #region Control Events
        
        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.epa.gov/cleanenergy/documents/egridzips/eGRID_9th_edition_V1-0_year_2010_eGRID_subregions.jpg");
        }

        
        public void Open(CheckedListBox clb, System.Windows.Forms.Control.ControlCollection controls)
        {
            parentCheckListBox = clb;
            parentControls = controls;
           Globals.ThisAddIn.Application.ActiveWorkbook.EnableConnections();

           int smallgap = 2;//3;
            int biggap = 10;
            int bottom = label1.Bottom;
            
            foreach (object obj in clb.CheckedItems)
            {
                bottom = AddControls(biggap,smallgap, bottom, obj);

            }
            label2.Top = bottom + (3*biggap);
            label2.Left = 5;
            linkLabel1.Top = label2.Bottom + 1;
            linkLabel1.Left = 5;
            label3.Top = linkLabel1.Bottom + 1;
            label3.Left = 5;
            btnBack.Top = label3.Bottom + biggap;
            btnCalculate.Top = label3.Bottom + biggap;

        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            //if (isEnergyCostChecked == true)
            //{
                Globals.ThisAddIn.LaunchEnergyCostControl(parentCheckListBox, parentControl, parentControls);
            //}
            //else
            //{
            //    Globals.ThisAddIn.LaunchRegressionControl(parentControl.classLevelType);
            //}
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.fromCO2Emission = false;
            bool itemSelected = true;
            //Check whether all the energy sources have the CO2 emission factor and check whether all the factors are valid numbers
            Globals.ThisAddIn.CO2EmissionFactors = new Dictionary<string, string>();
            foreach (object obj in parentCheckListBox.CheckedItems)
            {
                Control[] collection = this.Controls.Find(STR_TXT_CO2EMISSION_PREFIX + obj.ToString(), true);
                if (collection.Length > 0)
                {
                    TextBox txtBox = collection[0] as TextBox;
                    if (txtBox != null)
                    {
                        if (ValidateTextBox(txtBox) == 0)
                        {
                            Globals.ThisAddIn.CO2EmissionFactors.Add(new KeyValuePair<string, string>(obj.ToString(), txtBox.Text));
                        }
                        else
                        {
                            itemSelected = false;
                        }
                    }
                }
            }


            if (parentCheckListBox.CheckedItems.Count == Globals.ThisAddIn.CO2EmissionFactors.Count)
            {
                Globals.ThisAddIn.fromCO2Emission = true;

            }
            if (itemSelected==true)
            {
               parentControl.runFunction(null, null);
            }
            else
            {
                if(Globals.ThisAddIn.CO2EmissionFactors.Count>0)
                MessageBox.Show("Please enter a valid number");
                else
                    parentControl.runFunction(null, null);
            }
                                
                      
        }

       

        private void ComboBoxSelectedIndexChanged(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            if (cmb.Name.Contains(STR_CMB_ENERGYSOURCE_PREFIX))
            {
                string[] strArray = cmb.Name.Split(new string[] { STR_CMB_ENERGYSOURCE_PREFIX }, StringSplitOptions.RemoveEmptyEntries);
                if (strArray.Length > 0)
                {
                    Control[] collection = this.Controls.Find(STR_CMB_FUELTYPE_PREFIX + strArray[0], false);
                    if (collection.Length > 0)
                    {
                        ComboBox cmbBox = collection[0] as ComboBox;
                        LoadComboBox(STR_CMB_FUELTYPE_PREFIX + strArray[0], ComboBoxFillType.FuelType, cmb.GetItemText(cmb.SelectedItem));
                        if (cmb.GetItemText(cmb.SelectedItem) == "Custom")
                        {
                            cmbBox.Enabled = false;
                            FillEmissionFactorTextBox(STR_TXT_CO2EMISSION_PREFIX + strArray[0], cmb.GetItemText(cmb.SelectedItem), string.Empty, true);
                        }
                        else
                        {
                            if (cmbBox.Items.Count > 1)   // Fuel type dropdown is empty so disabling it
                                cmbBox.Enabled = true;
                            else
                                cmbBox.Enabled = false;
                        }
                    }
                }
            }
            else if (cmb.Name.Contains(STR_CMB_FUELTYPE_PREFIX))
            {
                string[] strArray = cmb.Name.Split(new string[] { STR_CMB_FUELTYPE_PREFIX }, StringSplitOptions.RemoveEmptyEntries);
                if (strArray.Length > 0)
                {
                    Control[] collection = this.Controls.Find(STR_CMB_ENERGYSOURCE_PREFIX + strArray[0], false);
                    if (collection.Length > 0)
                    {
                        ComboBox cmbBox = collection[0] as ComboBox;
                        FillEmissionFactorTextBox(STR_TXT_CO2EMISSION_PREFIX + strArray[0], cmbBox.GetItemText(cmbBox.SelectedItem), cmb.GetItemText(cmb.SelectedItem), false);

                    }
                }
            }


        }
        private void ComboBoxLeave(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            tTEmissionData.Hide(cmb);
        }

        private void ComboBoxDropDownClosed(object sender, EventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            tTEmissionData.Hide(cmb);
        }

        private void ComboBoxDrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox cmb = sender as ComboBox;
            if (e.Index < 0) { return; }
            string text = cmb.GetItemText(cmb.Items[e.Index]);
            e.DrawBackground();
            using (SolidBrush br = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(text, e.Font, br, e.Bounds);
            }
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                tTEmissionData.Show(text, cmb, e.Bounds.Right, e.Bounds.Bottom);
            }

            e.DrawFocusRectangle();
        }

      

        #endregion

        #region Helper Methods
        
        private int AddControls(int biggap, int smallgap,int bottom, object obj)
        {

            int right = 115;
            int left = 5;
            Label lbl = new Label();
            lbl.Name = STR_LBL_PREFIX + obj.ToString();
            lbl.Text = obj.ToString();
            lbl.Left = left;
            lbl.Font = new Font(lbl.Font,FontStyle.Bold);
            this.Controls.Add(lbl);
            lbl.AutoSize = true;
            lbl.Top = bottom + biggap;
            bottom = lbl.Bottom;

            Label lblEnergySource = new Label();
            lblEnergySource.Name = STR_LBL_ENERGYSOURCE_PREFIX + obj.ToString();
            lblEnergySource.Text = EnPIResources.co2LblEnergySource;
            lblEnergySource.Left = left;
            this.Controls.Add(lblEnergySource);
            lblEnergySource.AutoSize = true;
            lblEnergySource.Top = bottom + smallgap;//biggap;

            bottom = lblEnergySource.Bottom;
            ComboBox cmbEnergySource = new ComboBox();
            cmbEnergySource.Name = STR_CMB_ENERGYSOURCE_PREFIX + obj.ToString();
            cmbEnergySource.Left = left;
            this.Controls.Add(cmbEnergySource);
            cmbEnergySource.Top = bottom + smallgap;//biggap;
            //cmbEnergySource.Left = right;//lblEnergySource.Right + biggap;
            LoadComboBox(cmbEnergySource.Name, ComboBoxFillType.EnergySource,string.Empty);
            cmbEnergySource.SelectedIndexChanged+=new EventHandler(ComboBoxSelectedIndexChanged);
            cmbEnergySource.DrawMode = DrawMode.OwnerDrawFixed;
            cmbEnergySource.DrawItem += new DrawItemEventHandler(ComboBoxDrawItem);
            cmbEnergySource.DropDownClosed += new EventHandler(ComboBoxDropDownClosed);
            cmbEnergySource.Leave += new EventHandler(ComboBoxLeave);
            cmbEnergySource.Width =230;


            bottom = cmbEnergySource.Bottom;

            Label lblFuelType = new Label();
            lblFuelType.Name = STR_LBL_FUELTYPE_PREFIX + obj.ToString();
            lblFuelType.Text = EnPIResources.co2LblFuelType;
            lblFuelType.Left = left;
            this.Controls.Add(lblFuelType);
            lblFuelType.AutoSize = true;
            lblFuelType.Top = bottom + smallgap;

            bottom = lblFuelType.Bottom;
            
            ComboBox cmbFuelType = new ComboBox();
            cmbFuelType.Name = STR_CMB_FUELTYPE_PREFIX + obj.ToString();
            this.Controls.Add(cmbFuelType);
            cmbFuelType.Top = bottom + smallgap;
            cmbFuelType.Left = left;
            //cmbFuelType.Left = right;//lblFuelType.Right + biggap;
            cmbFuelType.SelectedIndexChanged += new EventHandler(ComboBoxSelectedIndexChanged);
            cmbFuelType.DrawMode = DrawMode.OwnerDrawFixed;
            cmbFuelType.DrawItem += new DrawItemEventHandler(ComboBoxDrawItem);
            cmbFuelType.DropDownClosed += new EventHandler(ComboBoxDropDownClosed);
            cmbFuelType.Leave += new EventHandler(ComboBoxLeave);
            cmbFuelType.Enabled = false;
            cmbFuelType.Width = 230;

            bottom = cmbFuelType.Bottom;

            Label lblCO2Emission = new Label();
            lblCO2Emission.Name = STR_LBL_CO2EMISSION_PREFIX+ obj.ToString();
            lblCO2Emission.Text = EnPIResources.co2LblEmissionFactor;
            lblCO2Emission.Left = left;
            this.Controls.Add(lblCO2Emission);
            lblCO2Emission.AutoSize = true;
            lblCO2Emission.Top = bottom + smallgap;

            bottom = lblCO2Emission.Bottom;

            TextBox txtCO2Emission = new TextBox();
            txtCO2Emission.Name = STR_TXT_CO2EMISSION_PREFIX+ obj.ToString();
            txtCO2Emission.Left = left;
            this.Controls.Add(txtCO2Emission);
            txtCO2Emission.AutoSize = true;
            txtCO2Emission.Top = bottom +smallgap;
            //txtCO2Emission.Left = right;//lblCO2Emission.Right + biggap;
            txtCO2Emission.Width = 230;

            bottom = txtCO2Emission.Bottom + biggap;
            //Label lblUnits = new Label();
            //lblUnits.Text = EnPIResources.co2LblUnits; 
            //this.Controls.Add(lblUnits);
            //lblUnits.AutoSize = true;
            //lblUnits.Top = txtCO2Emission.Bottom + 1;


            //bottom = lblUnits.Bottom + biggap;
                                   
            return bottom;
        }
        
        private void LoadComboBox(string controlName, ComboBoxFillType comboBoxType,string filterOption)
        {
            DataTable dt=CO2EmissionUtils.GetCO2Emissions();
            if (dt != null)
            {   
                Control[] collection = this.Controls.Find(controlName, false);
                if (collection.Length > 0)
                {
                    ComboBox cmbBox = collection[0] as ComboBox;
                    if (comboBoxType == ComboBoxFillType.EnergySource)
                    {
                        cmbBox.DataSource = dt.DefaultView.ToTable(true, "EnergySource");
                        cmbBox.DisplayMember = "EnergySource";
                    }
                    if (comboBoxType == ComboBoxFillType.FuelType)
                    {
                        DataRow[] dr = dt.Select("EnergySource='" + filterOption + "'");
                        DataTable newDt = new DataTable();
                        newDt = dt.Clone();
                        dr.CopyToDataTable<DataRow>(newDt, LoadOption.Upsert);
                        cmbBox.DataSource = newDt;
                        cmbBox.DisplayMember = "FuelType";
                      
                    }
                }

            }
        }

        

        private void FillEmissionFactorTextBox(string txtCtrlName, string energySource, string fuelType,bool isReadOnly)
        {
            DataTable dt = CO2EmissionUtils.GetCO2Emissions();
            if (dt != null)
            {
                Control[] collection = this.Controls.Find(txtCtrlName, false);
                if (collection.Length > 0)
                {
                    TextBox txtBox = collection[0] as TextBox;
                    txtBox.ReadOnly = !isReadOnly;
                    DataRow[] dr = dt.Select("EnergySource='" + energySource + "' AND FuelType='" + fuelType + "'");
                    if (dr.Length > 0)
                    {
                        txtBox.Text = dr[0]["EmissionFactor"].ToString();
                    }
                }

            }
        }


        private int ValidateTextBox(TextBox textBox)
        {
            int status = 0;
            if(string.IsNullOrEmpty(textBox.Text))
            {
                status = 1;
            }
            if (new Regex("^[1-9]\\d*(\\.\\d+)?$").IsMatch(textBox.Text)==false)
            {
                status = 2;
            }
            return status;
        }



     
        #endregion

       

    }
}
