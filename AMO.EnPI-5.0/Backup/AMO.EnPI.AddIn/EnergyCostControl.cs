using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;

namespace AMO.EnPI.AddIn
{
    public partial class EnergyCostControl : UserControl
    {
        RegressionControl parentControl;
        CheckedListBox parentCheckListBox;
        ControlCollection parentControls;
        bool isCO2EmissionChecked;

        public EnergyCostControl(RegressionControl parentControl)
        {
            InitializeComponent();

            this.parentControl = parentControl;
            

            if (Globals.ThisAddIn.fromWizard)
            {
                this.label1.Text = "Step 6: Energy Cost Data";
            }
            else
            {
                this.label1.Text = "Energy Cost Data";
            }
        }

        public void Open(CheckedListBox clb, System.Windows.Forms.Control.ControlCollection contorls)
        {
            parentCheckListBox = clb;
            parentControls = Controls;
            Globals.ThisAddIn.Application.ActiveWorkbook.EnableConnections();

            int smallgap = 3;
            int biggap = 10;

            int bottom = label2.Bottom;
            int count = 0;

                //initalized to 2 since there can only be a 1 to 1 match for fuel source and it's assciated cost column
                Globals.ThisAddIn.energyCostColumnMatchArray = new string[clb.CheckedItems.Count, 2];

                foreach (object obj in clb.CheckedItems)
                {
                    //add the name of the column for mapping
                    Globals.ThisAddIn.energyCostColumnMatchArray[count, 0] = obj.ToString();
                    Label lbl = new Label();
                    lbl.Text = obj.ToString();
                    this.Controls.Add(lbl);
                    lbl.AutoSize = true;
                    lbl.Top = bottom + biggap;
                    bottom = lbl.Bottom;

                CheckedListBox newCLB = new CheckedListBox();
                newCLB.CheckOnClick = true;
                this.Controls.Add(newCLB);
                newCLB.Top = bottom + smallgap;
                
                for (int i = 0; i < clb.Items.Count; i++)
                {
                    bool notPresentInOtherControls = true;

                    foreach (Control cntrl in contorls)
                    {
                        if (cntrl.GetType() == typeof(CheckedListBox))
                            foreach (object obj2 in ((CheckedListBox)cntrl).CheckedItems)
                            {
                                if(clb.Items[i].Equals(obj2))
                                    notPresentInOtherControls = false;
                            }
                    }

                    if (!clb.Items[i].Equals(obj) && notPresentInOtherControls)
                        newCLB.Items.Add(clb.Items[i]);
                }

                SetSize();
                bottom = newCLB.Bottom;
                count++;
            }

            btnCalculate.Top = bottom + biggap;
            btnBack.Top = bottom + biggap;
                        
            //Globals.ThisAddIn.wizardPane.Visible = false;

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //RestoreValues(thisSheet, checkedListBox1, Utilities.Constants.CB_TABLES);
            //RestoreValues(thisSheet, checkBox1, "NewSheet");

            //this.checkedListBox1.Visible = true;
            //this.btnRun.Visible = true;
        }

        private void SetSize()
        {
            int maxwidth = 0;

            foreach (Control me in this.Controls)
            {
                if (me.GetType() == new System.Windows.Forms.CheckedListBox().GetType())
                {
                    int sz1 = ((CheckedListBox)me).Items.Count;
                    int sz2 = Convert.ToInt16(((CheckedListBox)me).GetItemHeight(0)); //check box height
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

            Globals.ThisAddIn.wizardPane.Width = ThisAddIn.paneWidth;
            Globals.ThisAddIn.wizardPane.Visible = true;

        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
                Globals.ThisAddIn.fromEnergyCost = false;
                //TODO: add check to make sure all energy sources are accounted for
                bool itemSelected = true;
                int count = 0;
                int selectedCount = 0;

                foreach (Control ctrl in this.Controls)
                {
                    if (ctrl.GetType() == new CheckedListBox().GetType())
                    {
                        if (((CheckedListBox)ctrl).CheckedIndices.Count != 1)
                            itemSelected = false;
                        else
                        {
                            //add mapped column to array
                            Globals.ThisAddIn.energyCostColumnMatchArray[count, 1] = ((CheckedListBox)ctrl).CheckedItems[0].ToString();
                            selectedCount++; 
                        }
                        count++;
                    }
                }


                if (count == selectedCount)
                {
                    Globals.ThisAddIn.fromEnergyCost = true;
                }
                if (itemSelected == true)
                {
                   Globals.ThisAddIn.LaunchCO2EmissionControl(parentCheckListBox, parentControl, parentControls);
                }
                else
                {
                    if(selectedCount >0)
                    MessageBox.Show("One and only one selection must be made for each energy source.");
                    else
                        Globals.ThisAddIn.LaunchCO2EmissionControl(parentCheckListBox, parentControl, parentControls);
                                
                }
                
                //Commented TFS Ticket : 69337
                //if (itemSelected)
                //{
                //    Globals.ThisAddIn.fromEnergyCost = true;
                //    //If the CO2 calculations is checked then pass on to next screen else run calculations.
                //    if (isCO2EmissionChecked == true)
                //    {
                //        //The reason why the user is on this screen is user has selected the calculate energy cost check box
                //        //So No harm in sending true here
                //        Globals.ThisAddIn.LaunchCO2EmissionControl(parentCheckListBox, parentControl, parentControls, true);
                //    }
                //    else
                //    {
                //        parentControl.runFunction(null, null);
                //    }
                //}
                //else
                //{
                //    MessageBox.Show("One and only one selection must be made for each energy source.");
                //}
            
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchRegressionControl(parentControl.classLevelType);
        }
    }

}
