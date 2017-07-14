using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;
using System.Text.RegularExpressions;

namespace AMO.EnPI.AddIn
{
    public partial class RegressionControl : UserControl
    {
        ArrayList myears;
        public Constants.EnPITypes classLevelType;

        public RegressionControl(Constants.EnPITypes type)
        {
            classLevelType = type;

            InitializeComponent();

            switch (type)
            {
                case Constants.EnPITypes.Actual:
                    this.label2.Visible = false;
                    this.checkedListBox2.Visible = false;
                    this.listBox2.Visible = false;
                    this.label6.Visible = false;
                    this.lblReportYear.Visible = false;
                    this.listBox3.Visible = false;
                    this.label3.Visible = true;
                    this.checkedListBox3.Visible = true;
                    this.label4.Visible = true;
                    this.checkedListBox4.Visible = true;
                    selectedType = Constants.EnPITypes.Actual;
                    break;
                case Constants.EnPITypes.Backcast:
                    this.listBox1.Visible = true;
                    this.label5.Visible = true;
                    selectedType = Constants.EnPITypes.Backcast;
                    break;
                default: // this is for regression
                    this.label3.Visible = false;
                    this.checkedListBox3.Visible = false;
                    this.label4.Visible = false;
                    this.checkedListBox4.Visible = false;
                    break;
            }

            //only show partial title for this pane if coming from the wizard - ticket #66438
            if (Globals.ThisAddIn.fromWizard)
            {
                this.lblRegressionTitle.Text = "Step 5: Select Data for Calculations";
            }
            else
            {
                this.lblRegressionTitle.Text = "Select Data for Calculations";
            }

            DataLO = ((Excel.Range)Globals.ThisAddIn.Application.Selection).ListObject;

            if (DataLO == null) DataLO = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[1];

        }

        public Excel.ListObject DataLO;
        private Constants.EnPITypes selectedType;

        private void SetSize()
        {
            int maxwidth = this.label7.Width;

            foreach (Control me in this.Controls)
            {
                if (me.GetType() == new System.Windows.Forms.CheckedListBox().GetType())
                {
                    int sz1 = ((CheckedListBox)me).Items.Count;
                    int sz2 = Convert.ToInt16( ((CheckedListBox)me).GetItemHeight(0) ); //check box height
                    int sz3 = Convert.ToInt16(((CheckedListBox)me).GetItemText(0).Length);
                    int sz4 = Convert.ToInt16(me.Font.SizeInPoints * 0.65);

                    foreach (string str in ((CheckedListBox)me).Items)
                    {
                        sz3 = Math.Max(sz3, str.Length);
                    }

                    me.Width = Math.Max(maxwidth, Math.Max(me.Width, sz3 * sz4 + sz2 + me.Margin.Left + me.Margin.Right));
                    me.Height = Math.Max(sz2, sz1 * sz2 + me.Margin.Top + me.Margin.Bottom);
                }
                else if (me.GetType() == new System.Windows.Forms.ListBox().GetType())
                {
                    int sz1 = ((ListBox)me).Items.Count;
                    int sz2 = Convert.ToInt16(((ListBox)me).GetItemHeight(0)); //check box height
                    int sz3 = Convert.ToInt16(((ListBox)me).GetItemText(0).Length);
                    int sz4 = Convert.ToInt16(me.Font.SizeInPoints * 0.65);

                    foreach (object str in ((ListBox)me).Items)
                    {
                        sz3 = Math.Max(sz3, str.ToString().Length);
                    }

                    me.Width = Math.Max(maxwidth, Math.Max(me.Width, sz3 * sz4 + sz2 + me.Margin.Left + me.Margin.Right));
                    me.Height = Math.Max(sz2, sz1 * sz2 + me.Margin.Top + me.Margin.Bottom);
                }
                maxwidth = Math.Max(maxwidth, me.Width);
            }

            int smallgap = 3;
            int biggap = 10;
            //Modified by Suman : TFS Ticket 69133 
            //TFS Ticket :70388
            this.label2.Top = this.checkedListBox1.Bottom + biggap;
            this.checkedListBox2.Top = this.label2.Bottom + smallgap;
            //Building square feet and production
            this.label3.Top = this.checkedListBox2.Visible ? this.checkedListBox2.Bottom + biggap : this.checkedListBox1.Bottom + biggap;
            this.checkedListBox3.Top = this.label3.Bottom + smallgap;
            this.label4.Top = this.checkedListBox3.Bottom + biggap;
            this.checkedListBox4.Top = this.label4.Bottom + smallgap;
            //Baseline year
            this.label5.Top = this.label3.Visible ? this.checkedListBox4.Bottom + biggap : this.checkedListBox2.Bottom + biggap;
            this.listBox1.Top = this.label5.Bottom + smallgap;
            //Model year
            this.label6.Top = this.listBox1.Bottom + biggap;
            this.listBox2.Top = this.label6.Bottom + smallgap;

            //Report Year             
            this.lblReportYear.Top = this.label6.Visible ? this.listBox2.Bottom +  biggap : this.listBox1.Bottom + biggap;
            this.listBox3.Top = this.lblReportYear.Bottom + smallgap;

            //TFS Ticket:77018
            this.btnRun.Top = this.label6.Visible ? this.listBox3.Bottom + biggap : this.listBox1.Bottom + biggap;

            //this.btnRun.Top = this.listBox3.Bottom + biggap;
            //Tfs Ticket :70388
            //this.label4.Top = this.label2.Visible ? this.checkedListBox2.Bottom + biggap : this.checkedListBox1.Bottom + biggap;
            //this.label5.Top = this.label2.Visible ? this.checkedListBox2.Bottom + biggap : this.checkedListBox1.Bottom + biggap;
            ////this.label3.Top = this.label2.Visible ? this.checkedListBox2.Bottom + biggap : this.checkedListBox1.Bottom + biggap;
            //this.checkedListBox4.Top = this.label4.Bottom + smallgap;
            //this.label3.Top = this.checkedListBox4.Bottom + biggap;
            //this.checkedListBox3.Top = this.label3.Bottom + smallgap;
            ////this.label5.Top = this.checkedListBox3.Bottom + biggap;
            //this.listBox1.Top = this.label5.Bottom + smallgap;
            //this.label6.Top = this.listBox1.Bottom + biggap;
            //this.listBox2.Top = this.label6.Bottom + smallgap;
            //this.checkBox1.Top = this.label6.Visible ? this.listBox2.Bottom + biggap : this.listBox1.Bottom + biggap; //+ listBox1.Height
            //this.chkCo2Emission.Top = this.checkBox1.Bottom + biggap;
            //this.btnRun.Top = this.chkCo2Emission.Bottom + biggap;
            
            Globals.ThisAddIn.wizardPane.Width = ThisAddIn.paneWidth;
            Globals.ThisAddIn.wizardPane.Visible = true;

        }


        public void AddColumnNames()
        {
            var illegalChars = new Regex("[']");

            if (DataLO != null)
            {
                foreach (Excel.ListColumn LC in DataLO.ListColumns)
                {
                    if (!illegalChars.IsMatch(LC.Name))
                    {
                        string nm = LC.Name;
                        int loc = nm.IndexOf(((char)10).ToString());
                        if (loc > 0 && loc != 1 + nm.IndexOf(((char)13).ToString() + ((char)10).ToString()))
                        {
                            nm = nm.Insert(loc, ((char)13).ToString());
                            LC.Name = nm;
                        }

                        if (LC != Utilities.ExcelHelpers.GetListColumn(DataLO, EnPIResources.yearColName))
                        {
                            this.checkedListBox1.Items.Add(LC.Name);
                            this.checkedListBox2.Items.Add(LC.Name);
                            this.checkedListBox3.Items.Add(LC.Name);
                            this.checkedListBox4.Items.Add(LC.Name);
                        }
                    }
                    else
                        MessageBox.Show("The column header for \"" + LC.Name + "\" contains an apostrophe ('), please remove the apostrophe before proceeding.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);

                }

                RestoreValues(this.checkedListBox1, Utilities.Constants.WSPROP_SOURCE);
                RestoreValues(this.checkedListBox2, Utilities.Constants.WSPROP_VARS);
                RestoreValues(this.checkedListBox3, Utilities.Constants.WSPROP_BLDG);
                RestoreValues(this.checkedListBox4, Utilities.Constants.WSPROP_PRODUCTION);
            }


        }

        public void AddYears()
        {
            if (DataLO != null)
            {
                object[] years = Utilities.ExcelHelpers.getYears(DataLO);
                if (years != null)
                {
                    for (int i = 0; i < years.Count(); i++)
                    {
                        this.listBox1.Items.Add(years[i].ToString());
                    }

                    RestoreValues(this.listBox1, Utilities.Constants.WSPROP_BASELINE);
                }
            }
        }
        public void AddReportYears()
        {
           this.listBox3.Items.Clear();
            if (DataLO != null)
            {
                object[] years = Utilities.ExcelHelpers.getYears(DataLO);
                if (years != null)
                {
                    for (int i = this.listBox1.SelectedIndex; i < years.Count(); i++)
                    {
                        if (!this.listBox1.Text.Equals(years[i]))
                        {
                            this.listBox3.Items.Add(years[i].ToString());
                        }
                    }

                    RestoreValues(this.listBox3, Utilities.Constants.WSPROP_REPORTYEAR);
                }
            }
            this.listBox3.Height = this.listBox2.Height;
            this.listBox3.Refresh();
           
        }
       private bool validatenulldata()
        {
            string colName = "";
            int j = 0;
            object[] years = Utilities.ExcelHelpers.getYears(DataLO);
            if (DataLO != null)
            {
                foreach (Excel.ListColumn LC in DataLO.ListColumns)
                {
                    foreach (DataRow dr in Utilities.ExcelHelpers.rangeTable(DataLO).Rows)
                    {
                        string tdata = dr[LC.Name].ToString();
                        if (tdata == "")
                        {
                            j++;
                            colName = LC.Name;
                        }
                    }
                }
                
            }
            if (j == 0)
                return true;
            else
            {
                MessageBox.Show("A cell in column \"" + colName + "\" is blank. Please either delete the row or enter a zero in the blank cell before proceeding.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); //EnPIResources.warningNullValue
                return false;
            }
        }

        private bool validatesplchar()

        {
            int j = 0;
            var cspl = new Regex("[a-zA-Z!@#$%^&*)(]"); 
            //var cspl = new Regex("^[a-zA-Z0-9 ]*$");
            object[] years = Utilities.ExcelHelpers.getYears(DataLO);
            if (DataLO != null)
            {
                foreach (Excel.ListColumn LC in DataLO.ListColumns)
                {
                    //string test = LC.Name;
                    foreach (DataRow dr in Utilities.ExcelHelpers.rangeTable(DataLO).Rows)
                    {
                        string tdata = dr[LC.Name].ToString();
                        if (cspl.IsMatch(tdata))
                        {
                            if(!LC.Name.ToLower().Equals(EnPIResources.yearColName.ToLower()))
                            {
                                //MessageBox.Show("Please relabel the \"" + LC.Name.ToString() + "\" column \"Period\"", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                MessageBox.Show("Column \"" + LC.Name.ToString() + "\" contains a special character or letter. Please either remove the column from the Excel table, or replace the special character with a number before proceeding.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                                return false;
                            }
                        }
                    }
                }


            }
            if (j == 0)
                return true;
            else
            {
                MessageBox.Show(EnPIResources.warningSplChars, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning); 
                return false;
            }
            
        }
        public void AddModelYears()
        {
            myears = new ArrayList();

            string methd;   
            if (DataLO != null)
            {
                object[] years = Utilities.ExcelHelpers.getYears(DataLO);
                if (years != null)
                {
                    for (int i = this.listBox1.SelectedIndex; i < years.Count(); i++)
                    {
                        if (i == this.listBox1.SelectedIndex) methd = " (Forecast)";
                        else if (i == years.Count() - 1) methd = " (Backcast)";
                        else methd = " (Chaining)";

                        myears.Add(new ModelYear(years[i].ToString(), years[i].ToString() + methd) );
                    }

                    RestoreValues(this.listBox2, Utilities.Constants.WSPROP_YEAR);
                }
            }
            this.listBox2.DataSource = myears;
            this.listBox2.ValueMember = "YearName";
            this.listBox2.DisplayMember = "DisplayName";
            this.listBox2.Refresh();

            if (selectedType.Equals(Constants.EnPITypes.Backcast))
            {
                this.listBox2.Height = this.listBox1.Height;
                this.listBox2.Width = this.listBox1.Width;

                this.listBox2.Top = this.label6.Bottom +3;
                //this.btnRun.Top = this.listBox2.Bottom + 10;

                this.label6.Visible = true;
                this.listBox2.Visible = true;
            }
        }

        public void Open()
        {
            AddColumnNames();
            AddYears();
            //AddReportYears();
            SetSize();
            this.checkedListBox1.Visible = true;
            this.btnRun.Visible = true;
        }

        public void Open(RegressionControl rc)
        {
            //AddColumnNames();
            //AddYears();
            //SetSize();
            //this.checkedListBox1.Visible = true;
            //this.btnRun.Visible = true;
            for (int i = 0; i < rc.Controls.Count; i++)
            {
                this.Controls.Add(rc.Controls[i]);
            }
        }

        private void RestoreValues(CheckedListBox lbox, string propname)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string vals = Utilities.ExcelHelpers.getWorksheetCustomProperty(thisSheet, propname);

            if (vals != null)
            {
                System.Xml.XmlReader xvals = System.Xml.XmlReader.Create(new System.IO.StringReader(vals));
                string x;
                while (xvals.Read())
                {
                    if (xvals.NodeType == System.Xml.XmlNodeType.Text)
                    {
                        x = xvals.Value;

                        for (int i = 0; i < lbox.Items.Count; i++)
                        {
                            if (lbox.Items[i].ToString() == x) lbox.SetItemCheckState(i, System.Windows.Forms.CheckState.Checked);
                        }
                    }
                }
            }
        }

        private void RestoreValues(ListBox lbox, string propname)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            string val = (string)Utilities.ExcelHelpers.getWorksheetCustomProperty(thisSheet, propname);

            if (val != null)
            {
                for (int i = 0; i < lbox.Items.Count; i++)
                {
                    if (lbox.Items[i].ToString() == val)
                        lbox.SetSelected(i, true);
                }
            }
        }

        private void SaveValues(List<string> values, string propname)
        {
            System.Text.StringBuilder strvals = new System.Text.StringBuilder();
            System.Xml.XmlWriter xvals = System.Xml.XmlWriter.Create(strvals);
            xvals.WriteStartElement("Values");

            foreach(string s in values)
            {
                xvals.WriteElementString("Value", s);
            }

            xvals.WriteEndElement();
            xvals.Close();
            
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Utilities.ExcelHelpers.addWorksheetCustomProperty(thisSheet, propname, strvals.ToString());
        }

        private void SaveValues(string values, string propname)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            Utilities.ExcelHelpers.addWorksheetCustomProperty(thisSheet, propname, values);
        }

        private bool SetSources()
        {
            Globals.ThisAddIn.SelectedSources = new List<string>();

            int j = 0;
            bool warn = false;

            var chk = new Regex(EnPIResources.sourceUnitCheck.ToUpper());

            for (int i = 0; i < this.checkedListBox1.Items.Count; i++)
            {

                if (this.checkedListBox1.GetItemChecked(i))
                {

                    string nm = this.checkedListBox1.Items[i].ToString();

                    if (!chk.IsMatch(nm.ToUpper()))
                    {
                        warn = true;
                    }
                        nm = Utilities.DataHelper.CreateValidColumnName(nm);
                        Globals.ThisAddIn.SelectedSources.Add(nm);
                        j++;
                }
            }

            
            SaveValues(Globals.ThisAddIn.SelectedSources, Utilities.Constants.WSPROP_SOURCE);

            if (warn)
            {
                DialogResult result = MessageBox.Show(EnPIResources.sourceUnitCheckErr, "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                    return false;
                else
                    return true;
            }
            else
                return true;
        }

        private bool SetVariables()
        {
            Globals.ThisAddIn.SelectedVariables = new List<string>();
            int j = 0;
            for (int i = 0; i < this.checkedListBox2.Items.Count; i++)
            {
                if (this.checkedListBox2.GetItemChecked(i))
                {
                    string nm = this.checkedListBox2.Items[i].ToString();
                    nm = Utilities.DataHelper.CreateValidColumnName(nm);
                    Globals.ThisAddIn.SelectedVariables.Add(nm);
                    j++;
                }
            }

            SaveValues(Globals.ThisAddIn.SelectedVariables, Utilities.Constants.WSPROP_VARS);
            if (j == 0)
            {
                MessageBox.Show("Please select at least one variable.");
                return false;
            }
            else
                return true;
        }

        private void SetBuilding()
        {
            Globals.ThisAddIn.SelectedBuildings = new List<string>();
            
            for (int i = 0; i < this.checkedListBox3.Items.Count; i++)
            {
                if (this.checkedListBox3.GetItemChecked(i))
                {
                    string nm = this.checkedListBox3.Items[i].ToString();
                    nm = Utilities.DataHelper.CreateValidColumnName(nm);
                    Globals.ThisAddIn.SelectedBuildings.Add(nm);
                    
                }
            }

            SaveValues(Globals.ThisAddIn.SelectedBuildings, Utilities.Constants.WSPROP_BLDG);
         
        }



        private void CheckMinDataPt()
        {

            string yearErr = "";
            foreach (ModelYear y in myears)
            {
                int DataPt = Utilities.ExcelHelpers.DataPt(DataLO, y.YearName);
                if (DataPt < Convert.ToInt32(EnPIResources.MinDataPoints))
                {
                    if (yearErr == "")
                    {
                        yearErr = y.YearName;
                    }
                    else
                    {
                        yearErr += "," + y.YearName;
                    }
                }
            }

            if (yearErr != "")
                MessageBox.Show(EnPIResources.MinDataPointErr + " " + yearErr);

        }
        //Revereted back to the above method
        //private string CheckMinDataPt()
        //{

        //    string yearErr = "";
        //    foreach (ModelYear y in myears)
        //    {
        //        int DataPt= Utilities.ExcelHelpers.DataPt(DataLO,y.YearName);
        //        if (DataPt < Convert.ToInt32(EnPIResources.MinDataPoints))
        //        {
        //            if (yearErr == "")
        //            {
        //                yearErr = y.YearName;
        //            }
        //            else
        //            {
        //                yearErr += "," + y.YearName;
        //            }
        //        }
        //    }

        //    //Modified by Suman : TFS Ticket 69146
        //    //if (yearErr != "")
        //    //    //MessageBox.Show(EnPIResources.MinDataPointErr + " " + yearErr); 
        //    //    return false;
        //    //else
        //    //    return true;
        //    return yearErr;

        //}
        private void SetYear()
        {
            Globals.ThisAddIn.BaselineYear = this.listBox1.SelectedItem.ToString();
            Globals.ThisAddIn.SelectedYear = this.listBox2.SelectedValue.ToString();
            Globals.ThisAddIn.ReportYear = this.listBox3.SelectedItem.ToString();

            
            if ((this.listBox2.SelectedItem as ModelYear).DisplayName.Contains("(Forecast)"))
                Globals.ThisAddIn.AdjustmentMethod = "Forecast";
             else if((this.listBox2.SelectedItem as ModelYear).DisplayName.Contains("(Backcast)"))
                Globals.ThisAddIn.AdjustmentMethod = "Backcast";
             else 
                Globals.ThisAddIn.AdjustmentMethod = "Chaining";
            

       
            ArrayList modelyears = new ArrayList();

            foreach (ModelYear y in myears)
            {
                modelyears.Add(y.YearName);
            }

            Globals.ThisAddIn.Years = modelyears;

            SaveValues(Globals.ThisAddIn.BaselineYear, Utilities.Constants.WSPROP_BASELINE);
            SaveValues(Globals.ThisAddIn.SelectedYear, Utilities.Constants.WSPROP_YEAR);
            SaveValues(Globals.ThisAddIn.ReportYear, Utilities.Constants.WSPROP_REPORTYEAR);
        }

        private void SetActualYear()
        {
            Globals.ThisAddIn.BaselineYear = this.listBox1.SelectedItem.ToString();

            ArrayList modelyears = new ArrayList();

            foreach (ModelYear y in myears)
            {
                modelyears.Add(y.YearName);
            }

            Globals.ThisAddIn.Years = modelyears;

            SaveValues(Globals.ThisAddIn.BaselineYear, Utilities.Constants.WSPROP_BASELINE);
        }

        private void SetProduction()
        {
            Globals.ThisAddIn.SelectedProduction = new List<string>();
           
            for (int i = 0; i < this.checkedListBox4.Items.Count; i++)
            {
                if (this.checkedListBox4.GetItemChecked(i))
                {
                    string nm = this.checkedListBox4.Items[i].ToString();
                    nm = Utilities.DataHelper.CreateValidColumnName(nm);
                    Globals.ThisAddIn.SelectedProduction.Add(nm);
                   
                }
            }

            SaveValues(Globals.ThisAddIn.SelectedProduction, Utilities.Constants.WSPROP_PRODUCTION);
          
        }

        private void updateCheckboxes(object sender, ItemCheckEventArgs e)
        {
            
            if (e.NewValue.Equals(System.Windows.Forms.CheckState.Checked))
            {
                //remove any item just checked from all other boxes
                CheckedListBox box = (CheckedListBox)sender;
                string name = box.Name;
                if(box.SelectedItem != null)
                    switch (name)
                    {
                        case "checkedListBox1":
                            this.checkedListBox2.Items.Remove(box.SelectedItem);
                            this.checkedListBox3.Items.Remove(box.SelectedItem);
                            this.checkedListBox4.Items.Remove(box.SelectedItem);
                            break;
                        case "checkedListBox2":
                            this.checkedListBox1.Items.Remove(box.SelectedItem);
                            break;
                        case "checkedListBox3":
                            this.checkedListBox1.Items.Remove(box.SelectedItem);
                            break;
                        case "checkedListBox4":
                            this.checkedListBox1.Items.Remove(box.SelectedItem);
                            break;
                    }

            }
            else
            {
                //add any items just unchecked to all other boxes
                CheckedListBox box = (CheckedListBox)sender;
                string name = box.Name;
                if (box.SelectedItem != null)
                    switch (name)
                    {
                        case "checkedListBox1":
                            if (!checkedListBox2.Items.Contains(box.SelectedItem))
                                this.checkedListBox2.Items.Add(box.SelectedItem);
                            if (!checkedListBox3.Items.Contains(box.SelectedItem))
                                this.checkedListBox3.Items.Add(box.SelectedItem);
                            if (!checkedListBox4.Items.Contains(box.SelectedItem))
                                this.checkedListBox4.Items.Add(box.SelectedItem);
                            break;
                        case "checkedListBox2":
                            if (!checkedListBox1.Items.Contains(box.SelectedItem))
                                this.checkedListBox1.Items.Add(box.SelectedItem);
                            break;
                        case "checkedListBox3":
                            if (!checkedListBox1.Items.Contains(box.SelectedItem))
                                this.checkedListBox1.Items.Add(box.SelectedItem);
                            break;
                        case "checkedListBox4":
                            if (!checkedListBox1.Items.Contains(box.SelectedItem))
                                this.checkedListBox1.Items.Add(box.SelectedItem);
                            break;
                    }
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            string baselineYear = listBox1.Text;
            string modelYear = listBox2.Text;
            modelYear = modelYear.Substring(0, 4);
            string reportYear = listBox3.Text;
            if (listBox2.Text.Contains("Chaining"))
            {
                //The model year and report year , the year might be FY year, so in order to display message the year needs to be validated.
                // Reported By Ashly on 10/11/2015
                if (modelYear.Contains("FY") && reportYear.Contains("FY"))
                {
                    modelYear = modelYear.Replace("FY","");
                    reportYear = reportYear.Replace("FY", "");
                }
                if ((Convert.ToInt32(modelYear) <= Convert.ToInt32(reportYear)))
                {
                    Globals.ThisAddIn.LaunchEnergyCostControl(this.checkedListBox1, this, this.Controls);
                }
                else
                {
                    MessageBox.Show("The model and report year chosen will result in an invalid regression method that will not be accepted by the SEP Program. When using the Chaining regression method, the report year must be after the model year. Please choose an appropriate model and report year, or choose either the Forecast or Backcast regression method before proceeding.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            {
                Globals.ThisAddIn.LaunchEnergyCostControl(this.checkedListBox1, this, this.Controls);
            }
            
            //Globals.ThisAddIn.LaunchEnergyCostControl(this.checkedListBox1, this, this.Controls);
            //TFS Ticket 69337
            #region Commeted 
            //if (checkBox1.Checked == false && chkCo2Emission.Checked == false)
            //{
            //    Globals.ThisAddIn.fromEnergyCost = false;
            //    Globals.ThisAddIn.fromCO2Emission = false;
            //    runFunction(sender, e);
            //}
            //else
            //{
            //    //If Energy Cost is selected then the CO2 emission pane should be shown after the energy cost
            //    if (checkBox1.Checked == true && chkCo2Emission.Checked == false)
            //    {
            //        Globals.ThisAddIn.LaunchEnergyCostControl(this.checkedListBox1, this, this.Controls, chkCo2Emission.Checked);
            //    }
            //    if (checkBox1.Checked == false && chkCo2Emission.Checked == true)
            //    {
            //        Globals.ThisAddIn.LaunchCO2EmissionControl(this.checkedListBox1, this, this.Controls, checkBox1.Checked);
            //    }
            //    if (checkBox1.Checked == true && chkCo2Emission.Checked == true)
            //    {
            //        Globals.ThisAddIn.LaunchEnergyCostControl(this.checkedListBox1, this, this.Controls, chkCo2Emission.Checked);
            //    }
            //}
            #endregion

        }

        public void runFunction(object sender, EventArgs e)
        {   
            //this section was seperated into it's own function to accomidate for the new Energy Cost module
            Globals.ThisAddIn.hasSEPValidationError = false;
            if(Globals.ThisAddIn.lstSEPValidationValues !=null)
               Globals.ThisAddIn.lstSEPValidationValues.Clear();

            Globals.ThisAddIn.SelectedSourcesBestModelFormulas.Clear();

            if (!validatenulldata() || !validatesplchar() || !SetSources())
                return;

            if (selectedType == Constants.EnPITypes.Backcast)
            {
                if (!SetVariables())
                    return;
            }

            if (this.checkedListBox1.CheckedItems.Count == 0)
            {
                MessageBox.Show("Please select at least one energy source in units of MMBtu.");
                return;
            }

            if (selectedType.Equals(Constants.EnPITypes.Actual))
            {
                if (this.checkedListBox3.CheckedItems.Count == 0 && this.checkedListBox4.CheckedItems.Count == 0)
                {
                    MessageBox.Show("Please select a Production variable or Building Square Foot variable.");
                    return;
                }
            }

            if (this.listBox1.SelectedItem == null)
            {
                if(selectedType.Equals(Constants.EnPITypes.Actual))
                    MessageBox.Show("Please select a baseline year.");
                else
                    MessageBox.Show("Please select a baseline and model year.");
                return;
            }
            else
            {
                Globals.ThisAddIn.ModelYearSelected = this.listBox2.SelectedValue.ToString();
            }

            SetProduction();
            SetBuilding();
            Globals.Ribbons.Ribbon1.Model.Visible = true;
            Globals.Ribbons.Ribbon1.Rollup.Visible = true;
            CheckMinDataPt();
            //Code reverted back by suman
            //string yearErr = CheckMinDataPt();
            //if (yearErr == "")
            //{

                if (selectedType.Equals(Constants.EnPITypes.Backcast))
                {

                    SetYear();
                    Globals.ThisAddIn.plotEnPI(DataLO);
                }
                else
                {

                    SetActualYear();
                    Globals.ThisAddIn.actualEnPI(DataLO);
                }

                Globals.ThisAddIn.hideWizard();

                if (Globals.ThisAddIn.fromWizard)
                {
                    Globals.ThisAddIn.showWizard();
                    Globals.ThisAddIn.LaunchWizardControl(8);
                }

                this.Dispose();
            //}
            //else
            //{
            //    MessageBox.Show(EnPIResources.MinDataPointErr + " " + yearErr); 
            //}
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
                AddModelYears();
                AddReportYears();
                
                int smallgap = 3;
                int biggap = 10;

                this.lblReportYear.Top = this.label6.Visible ? this.listBox2.Bottom + biggap : this.listBox1.Bottom + biggap;
                this.listBox3.Top = this.lblReportYear.Bottom + smallgap;
                this.btnRun.Top = this.label6.Visible ? this.listBox3.Bottom + biggap : this.listBox1.Bottom + biggap;    
            //this.btnRun.Top = this.label6.Visible ? this.listBox2.Top + listBox1.Height + biggap : this.listBox1.Bottom + biggap;
               // this.btnRun.Top = this.listBox3.Bottom + biggap;    
            //this.checkBox1.Top = this.label6.Visible ? this.listBox2.Top + listBox1.Height + biggap : this.listBox1.Bottom + biggap;
                //this.chkCo2Emission.Top = this.checkBox1.Bottom + biggap;
                //this.btnRun.Top = this.chkCo2Emission.Bottom + biggap;
        }
     }

    public class ModelYear
    {
        public string YearName { get; set; }
        public string DisplayName { get; set; }

        public ModelYear(string strYear, string strDisplay)
        {
            this.YearName = strYear;
            this.DisplayName = strDisplay;
        }
    }
}
