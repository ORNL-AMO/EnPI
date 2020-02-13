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
using System.Reflection;

namespace AMO.EnPI.AddIn
{
    public partial class ChangeModelControl : UserControl
    {
        private Excel.ListObject DataLO1;
        private Excel.ListObject DataLO2;
        private SwitchModelCollection sModCol;
        private Excel.Worksheet energySheet;
        private int modelCount;
        IDictionary<string, int> cmbValues = new Dictionary<string, int>();
        public ChangeModelControl(bool fromWizard)
        {
            InitializeComponent();

            energySheet = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet);

            DataLO1 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[1];

            if (((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects.Count > 1)
            {
                DataLO2 = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[2];
            }
            
            this.btnBack.Visible = fromWizard;
            this.btnClose.Visible = fromWizard;
            this.btnNext.Visible = fromWizard;
        }

        public void Open()
        {
            sModCol = new SwitchModelCollection();
            PopulateModels(DataLO1);
            PopulateDropdown(DataLO1);
            if (DataLO2 != null)
            {
                PopulateModels(DataLO2);
                PopulateDropdown(DataLO2);
            }
       }


        private void PopulateModels(Excel.ListObject DataLO)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            //

            if (DataLO != null)
            {
                int bestModelNumber;

                //check to see if model can be switched from current sheet.
                try
                {
                    //TFS Ticket 71242
                    string strBestModel = ((Excel.Range)thisSheet.Cells[2, 1]).Value2.ToString();
                    string[] strArry = strBestModel.Split(new string[] { EnPIResources.bestModel }, StringSplitOptions.RemoveEmptyEntries);
                    bestModelNumber = Convert.ToInt32(strArry[0]);
                    //bestModelNumber = Convert.ToInt32(((Excel.Range)thisSheet.Cells[2, 1]).Value2.ToString().Substring(((Excel.Range)thisSheet.Cells[2, 1]).Value2.ToString().Length - 1));
                }
                catch
                {
                    this.label1.Text = "In order to switch a model, you must first be on the sheet that the model is listed on.";
                    this.comboBox1.Enabled = false;
                    return;
                }

                int currentModelNumber = 0;
                int varCount = 1;
                modelCount = 0;
                bool bestModel = false;

                foreach (Excel.ListRow row in DataLO.ListRows)
                {

                    string[] variables = new string[DataLO.ListRows.Count];
                    double[] VarPvalue= new double[DataLO.ListRows.Count];

                    if (((Excel.Range)row.Range.Cells[1]).Value2 != null)
                    {
                        varCount = 1;
                        modelCount++;
                        //modelCount = row.Range.Count;
                        currentModelNumber = Convert.ToInt32(((Excel.Range)row.Range.Cells[1]).Value2.ToString());
                        if (currentModelNumber.Equals(bestModelNumber))
                            bestModel = true;
                        variables[0] = ((Excel.Range)row.Range.Cells[3]).Value2.ToString();
                        VarPvalue[0] = Convert.ToDouble(((Excel.Range)row.Range.Cells[6]).Value2.ToString());
                        //SwitchModel sm = new SwitchModel(currentModelNumber, variables, Convert.ToDouble(((Excel.Range)row.Range.Cells[7]).Value2.ToString())
                        //                                , ((Excel.Range)row.Range.Cells[13]).Value2.ToString(), bestModel, thisSheet.Name, modelCount
                        //                                , VarPvalue, Convert.ToDouble(((Excel.Range)row.Range.Cells[8]).Value2.ToString())
                        //                                , Convert.ToDouble(((Excel.Range)row.Range.Cells[9]).Value2.ToString()));
                        //Added by suman SEP Changes
                        SwitchModel sm = new SwitchModel(currentModelNumber, variables, Convert.ToDouble(((Excel.Range)row.Range.Cells[8]).Value2.ToString())
                                                      , ((Excel.Range)row.Range.Cells[14]).Value2.ToString(), bestModel, thisSheet.Name, modelCount
                                                      , VarPvalue, Convert.ToDouble(((Excel.Range)row.Range.Cells[9]).Value2.ToString())
                                                      , Convert.ToDouble(((Excel.Range)row.Range.Cells[10]).Value2.ToString()));
                        sModCol.Add(sm);
                        bestModel = false;
                    }
                    else
                    {
                        if (((Excel.Range)row.Range.Cells[3]).Value2 != null)
                        {
                            sModCol.Item(sModCol.Count - 1).VariableNames[varCount] = ((Excel.Range)row.Range.Cells[3]).Value2.ToString();
                            varCount++;
                        }
                        //if (((Excel.Range)row.Range.Cells[6]).Value2 != null)
                        //{
                        //    sModCol.Item(sModCol.Count - 1).VariablePvalues[varCount] =  Convert.ToDouble(((Excel.Range)row.Range.Cells[6]).Value2.ToString());
                        //    varCount++;
                        //}
                        if (((Excel.Range)row.Range.Cells[7]).Value2 != null)
                        {
                            sModCol.Item(sModCol.Count - 1).VariablePvalues[varCount] = Convert.ToDouble(((Excel.Range)row.Range.Cells[7]).Value2.ToString());
                            varCount++;
                        }
                    }
                }
            }

            this.label1.Text = "Select the model you wish to use to calculate the adjusted values on the EnPI Results and SEnPI Results sheets from the drop down below.";// +thisSheet.Name + ":";

            int t = DataLO.ListColumns["Model is Appropriate for SEP"]._Default.Length;
        }

        private void PopulateDropdown(Excel.ListObject DataLO)
        {
            this.comboBox1.Items.Clear();
            cmbValues.Clear();
            if (DataLO != null)
            {
                foreach (SwitchModel SM in sModCol)
                {
                    string vars = "";
                    int count = 0;
                    foreach (string s in SM.VariableNames)
                    {
                        if (s != null && s != "(Intercept)")
                        {
                            if (SM.VariableNames.Length - 1 > count)
                                vars = vars + s + ", ";
                            else
                                vars = vars + s;
                        }
                        count++;
                    }
                    cmbValues.Add(vars.Substring(0, vars.Length - 1) + " (Adj R^2=" + Math.Round(SM.R2, 4) + ")", SM.ModelNumber);
                    this.comboBox1.Items.Add(vars.Substring(0, vars.Length - 1) + " (Adj R^2=" + Math.Round(SM.R2, 4) + ")");            
                }
                
            }
        }
        private Excel.Range BottomCell(Excel.Worksheet WS)
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(WS, 0).ToString();

            return (Excel.Range)WS.get_Range(addr, System.Type.Missing);
        }
        public void populateModelData(Excel.Worksheet thissheet,Excel.Worksheet WS, int ChangedModel)
        {
            
            
                        
            object[,] row1;
            int compareModel = 0;
            string ModelAppr = "";
            int loCount = thissheet.ListObjects.Count;
            bool isSEP = (WS.Name.Contains("SEP") ? true : false);
            foreach (Excel.ListObject ListObj in  thissheet.ListObjects)
            {
                foreach (Excel.ListColumn colm in ListObj.ListColumns)
                {
                    string cname = colm.Name;
                    if (colm.Name == "Model Number")
                    {
                        try
                        {
                            int bestModelNumber = ChangedModel;
                            int rowcount=ListObj.ListRows.Count;
                            string mdlValid = string.Empty;
                            foreach (Excel.ListRow row in ListObj.ListRows)
                            {
                                string[] c1 = new string[ListObj.ListColumns.Count];
                                if (((Excel.Range)row.Range.Cells[1]).Value2 != null)
                                {
                                    compareModel = Convert.ToInt32(((Excel.Range)row.Range.Cells[1]).Value2);
                                }

                                ModelAppr = ((Excel.Range)row.Range.Cells[2]).Value2 != null ? (((Excel.Range)row.Range.Cells[2]).Value2.ToString()) : "";

                                if (bestModelNumber == compareModel )
                                {
                                    int l = ListObj.ListColumns.Count;
                                    Excel.Range target1 = BottomCell(WS);
                                    //start1 = target1.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                                    string start1 = target1.get_Address(2, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

                                    if (ModelAppr.Length > 0)
                                    {
                                        string startadr = "";
                                        string stadr = "";
                                        string endadr = "";
                                        if(isSEP ==false)
                                        startadr = WS.ListObjects[2].Range.AddressLocal[Excel.XlReferenceStyle.xlA1].ToString();
                                        else
                                            startadr = WS.ListObjects[3].Range.AddressLocal[Excel.XlReferenceStyle.xlA1].ToString();
                                        int k = WS.Range[startadr.Substring(0, startadr.IndexOf(":")), start1].Rows.Count;
                                        endadr = start1.ToString();
                                        foreach (Excel.Range LR in WS.Range[startadr.Substring(0, startadr.IndexOf(":")), start1].Rows)
                                        {

                                            string lttxt = LR.Text != null ? LR.Text.ToString() : "";
                                            if (stadr.Length > 0 && LR.Text.ToString().Length > 0)
                                            {
                                                if (stadr.Length > 1)
                                                    break;
                                            }
                                            if (LR.Text.ToString() == thissheet.Name)
                                                stadr = LR.Address.ToString();
                                            if (stadr.Length > 0 && LR.Text.ToString().Length==0 )
                                                endadr = LR.Address.ToString();
                                        }

                                        int rcount = WS.Range[stadr, endadr].Rows.Count;
                                        Excel.Range rng = WS.get_Range(stadr, endadr);
                                        rng.EntireRow.Delete(Excel.XlDirection.xlUp);
                                      
                                    }
                                   

                                     //Added by Suman SEP Changes

                                    if (isSEP == false)
                                    {
                                        row1 = new object[1, 14];
                                        target1 = BottomCell(WS).get_Offset(1, 0).get_Resize(1, 14);
                                        //start1 = target1.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                                        start1 = target1.get_Address(2, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                                        int y = 0;
                                        for (int x = 0; x < ListObj.ListColumns.Count; x++)
                                        {
                                            if (x != 3 && x != 4 && x != 5 && x != 10 && x != 11 && x != 12)
                                            {

                                                if (((Excel.Range)row.Range.Cells[x + 1]).Value2 != null)
                                                {
                                                    if (x != 0)
                                                        row1[0, y] = ((Excel.Range)row.Range.Cells[x + 1]).Value2.ToString();
                                                    else
                                                        row1[0, y] = thissheet.Name;
                                                }
                                                else
                                                {
                                                    row1[0, y] = "";
                                                }
                                                y += 1;
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(row1[0, 1].ToString()))
                                        {
                                            mdlValid = row1[0, 1].ToString();
                                        }
                                        target1.get_Resize(1, 14).Value2 = row1;
                                        target1.Font.Color = mdlValid.ToUpper() == "TRUE" ? 0x00AA00 : 0x0000AA;
                                        target1.Font.Bold = true;
                                    }
                                    else
                                    {
                                        row1 = new object[1, 14];
                                        target1 = BottomCell(WS).get_Offset(1, 0).get_Resize(1, 14);
                                        start1 = target1.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
                                        int y = 0;
                                        for (int x = 0; x < ListObj.ListColumns.Count; x++)
                                        {
                                            if (x != 4 && x != 5 && x != 10 && x != 11 && x != 12)
                                            {

                                                if (((Excel.Range)row.Range.Cells[x + 1]).Value2 != null)
                                                {
                                                    if ((x != 0) && (x != 3))
                                                        row1[0, y] = ((Excel.Range)row.Range.Cells[x + 1]).Value2.ToString();
                                                    else if (x == 3)
                                                        row1[0, y] = GetSEPValue(row1[0, 2].ToString());
                                                 
                                                    else
                                                        row1[0, y] = thissheet.Name;
                                                }
                                                else
                                                {
                                                    row1[0, y] = "";
                                                }
                                                y += 1;
                                            }
                                        }
                                        if (!string.IsNullOrEmpty(row1[0, 1].ToString()))
                                        {
                                            mdlValid = row1[0, 1].ToString();
                                        }
                                        target1.get_Resize(1, 14).Value2 = row1;
                                        target1.Font.Color = mdlValid.ToUpper() == "TRUE" ? 0x00AA00 : 0x0000AA;
                                        //target1.Font.Color = GetSEPValue(row1[0, 2].ToString()) == "TRUE" ? 0x00AA00 : 0x0000AA;
                                        target1.Font.Bold = true;
                                        target1.NumberFormat = "0.0000";
                                    }
                                }

                            }

                        }
                        catch
                        {
                            //do nothing
                        }
                    }
                }
            }
        }

        private void UpdateEnPISheet(SwitchModel sm)
        {
            bool outputPresent = false;
            Excel.Worksheet adjustedDataSheet = new Excel.Worksheet();
            //Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet; //TFS Ticket: 71242
            Excel.Worksheet thisSheet = energySheet;

            //Update the SEP Validation Check values
            string sourceName = energySheet.Name;
            sourceName = sourceName.Substring(1, sourceName.Length - 1).ToString().TrimStart();
            Globals.ThisAddIn.SelectedSourcesBestModelFormulas[sourceName] = sm.Formula;
            string adjustedDatasheetName = string.Empty;
            foreach (GroupSheetCollection gsc in Globals.ThisAddIn.masterGroupCollection)
            {
                bool matchingCollection = false;

                foreach (GroupSheet GS in gsc)
                {
                    string GSname = GS.Name;
                    if (GS.WS.Equals(thisSheet))
                        matchingCollection = true;
                    if (GS.adjustedDataSheet && matchingCollection)
                    {
                        adjustedDataSheet = GS.WS;
                        adjustedDatasheetName = GS.WS.Name;
                        UpdateSEPValidationCheckList(adjustedDatasheetName);
                        outputPresent = true;
                    }
                    if (GS.outputSheet)
                    {                        
                        populateModelData(thisSheet, GS.WS, sm.ModelNumber);
                    }
                }
            }

            if (outputPresent)
            {
                string adjDataSheetName = adjustedDataSheet.Name;
                Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[adjustedDataSheet.Name]).Select(Type.Missing);
                thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
                //Update the latest values on Validation Check Table of Model Sheet
                try
                {
                    
                    for(int i=0;i< Globals.ThisAddIn.lstSEPValidationValues.Count;i++)
                    {
                        Excel.Range sepValidationChk = thisSheet.get_Range("C13").get_Offset(0,i).get_Resize(1,1);
                        sepValidationChk.Value2 = Globals.ThisAddIn.lstSEPValidationValues[i].SEPValidationCheck;
                    
                        sepValidationChk.NumberFormat = "General";
                    }
                    
                }
                catch (Exception ex)
                {

                }



                Excel.ListObject LO = (Excel.ListObject)thisSheet.ListObjects.get_Item(1);


                foreach (Excel.ListColumn col in LO.ListColumns)
                {
                    //------------------------------
                    //modify substring so that regression runs past 9 can switch models
                    string shortColName = col.Name;
                    //since sheet names are limited to 29 characters -- see CreateValidWorksheetName in ExcelHelpers.cs
                    if (col.Name.Length >= 36) 
                        shortColName = col.Name.Substring(0, 35);
                    string shtname = sm.SheetName.Substring(2);
                    if (shortColName.Equals("Modeled " + sm.SheetName.Substring(2)))
                    //------------------------------
                    {
                        col.DataBodyRange.Formula = "=" + sm.Formula;

                        //recalculate all of the List objects so numbers are updated.
                        foreach (Excel.Worksheet WS in WB.Worksheets)
                        {
                            string wsname = WS.Name;

                            foreach (Excel.ListObject ListObj in WS.ListObjects)
                            {
                                if (ListObj.DataBodyRange != null)
                                {
                                    ListObj.DataBodyRange.Dirty();
                                    ListObj.DataBodyRange.Calculate();
                                }
                            }
                        }
                    }
                }
                Excel.Range negativeMessageHeader = thisSheet.get_Range("A2");
                Excel.Range negativeMessageDescription = thisSheet.get_Range("A3");

                if (Globals.ThisAddIn.NegativeCheck(LO, Globals.ThisAddIn.modeledSourceIndex))
                {
                    negativeMessageHeader.EntireRow.Hidden = false;
                    negativeMessageDescription.EntireRow.Hidden = false;
                }
                else
                {
                    negativeMessageHeader.EntireRow.Hidden = true;
                    negativeMessageDescription.EntireRow.Hidden = true;
                }
            }
        }

        public void UpdateSEPValidationCheckList(string adjustedDataSheetName)
        {

            try
            {
                Globals.ThisAddIn.lstSEPValidationValues = new List<SEPValidationValues>();
                Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Sheets[adjustedDataSheetName]).Select(Type.Missing);
                Excel.Worksheet adjustedDataModelSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

                Excel.ListObject LO = (Excel.ListObject)adjustedDataModelSheet.ListObjects.get_Item(2);


                for (int x=1;x<=LO.ListColumns.Count;x++)
                {

                    string col = LO.ListColumns[x].Name;
                    if (!col.StartsWith("Col"))  //Need to remove this hack
                    {
                        object[,] row1 = new object[1, 6];
                        int y = 0;
                        foreach (Excel.ListRow row in LO.ListRows)
                        {

                            row1[0,y]= ((Excel.Range)row.Range.Cells[x]).Value2.ToString();
                            y++;
                        }
                       SEPValidationValues newVal = new SEPValidationValues(col, x, 0, Convert.ToDouble(row1[0,2]),
                                                                                       Convert.ToDouble(row1[0,1]),
                                                                                       Convert.ToDouble(row1[0,0]),
                                                                                       Convert.ToDouble(row1[0, 3]), 
                                                                                       Convert.ToDouble(row1[0, 4]), 
                                                                                       Convert.ToDouble(row1[0, 5]), string.Empty, false);
                        Globals.ThisAddIn.lstSEPValidationValues.Add(newVal);
                    }

                }
                



                foreach (SEPValidationValues sepVals in Globals.ThisAddIn.lstSEPValidationValues)
                {
                    int count = 0;
                    foreach (KeyValuePair<string, string> itm in Globals.ThisAddIn.SelectedSourcesBestModelFormulas)
                    {
                        if (itm.Value.Contains(sepVals.IndependentVariable))
                        {
                            count++;
                        }
                    }

                    if (count > 0)
                    {

                        if ((sepVals.MinModel < sepVals.AvgReportYr && sepVals.AvgReportYr < sepVals.MaxModel) || (sepVals.Minus3DevVal < sepVals.AvgReportYr && sepVals.AvgReportYr < sepVals.Plus3DevVal))
                        {
                            sepVals.SEPValidationCheck = "Pass";
                        }
                        else
                        {
                            sepVals.SEPValidationCheck = "Fail";
                        }
                    }
                    else
                    {
                        sepVals.SEPValidationCheck = "Not Included In Model";
                    }

                }
            }
            catch
            {
                //TODO: Log Implementation
            } 
        }

      
        //Modified by suman TFS Ticket :66435
        private void btnRun_Click(object sender, EventArgs e)
        {
            //int selectedModelNumber = this.comboBox1.SelectedIndex + 1;
            int selectedModelNumber;
            cmbValues.TryGetValue(this.comboBox1.Text,out selectedModelNumber);

           // int ModelCount = 0;
            foreach (SwitchModel sm in sModCol)
            {
                if (sm.ModelNumber == selectedModelNumber)
                {              
                    foreach (Excel.ListObject LO in energySheet.ListObjects)
                    {
                        LO.Range.Interior.Color = 0xFFFFFF;
                        LO.Range.Font.Color = 0x000000;
                        LO.Range.EntireRow.Font.Bold = false;
                        LO.HeaderRowRange.Interior.Color=0x000000;
                        LO.HeaderRowRange.Font.Color = 0xFFFFFF;
                        LO.HeaderRowRange.Font.Bold = true;
                        foreach (Excel.ListRow LR in LO.ListRows)
                        {
                            ((Excel.Range)LR.Range[1]).Font.Color = 0xAA0000;
                            if (Convert.ToInt32(((Excel.Range)LR.Range[1]).Value2) == selectedModelNumber)
                            {
                                
                                string targetAddress = GetRangeForSelectedModel(sm, LR);
                                Excel.Range target =energySheet.get_Range(targetAddress);
                                //If the Model is appropriate for SEP is True then show the row highlighted in Green else Red
                                if (Convert.ToBoolean(((Excel.Range)LR.Range[1, 2]).Value2)==true) 
                                {
                                     //LR.Range.EntireRow.Font.Bold = true;
                                    //LR.Range.Font.Color = 0x00AA00;
                                    //LR.Range.Interior.Color = 0xCDEFC6;
                                    target.EntireRow.Font.Bold = true;
                                    target.Font.Color = 0x00AA00;
                                    target.Interior.Color = 0xCDEFC6;
                                }
                                if (Convert.ToBoolean(((Excel.Range)LR.Range[1, 2]).Value2) == false)
                                {
                                    //LR.Range.EntireRow.Font.Bold = true;
                                    //LR.Range.Font.Color = 0x0000AA;
                                    //LR.Range.Interior.Color = 0xCEC8FF;
                                    target.EntireRow.Font.Bold = true;
                                    target.Font.Color = 0x0000AA;
                                    target.Interior.Color = 0xCEC8FF;
                                }

                                
                            }
                            //else
                            //{
                            //    LR.Range.EntireRow.Font.Bold = false;
                            //    LR.Range.Interior.Color = 0xFFFFFF;
                            //    LR.Range.Font.Color = 0x000000;
                                
                            //}
                        }
                    }
                    UpdateEnPISheet(sm);
                }

                //ModelCount += 1;

                //if (ModelCount.Equals(selectedModelNumber))
                //UpdateEnPISheet(sm);
            }       
        }


        private string GetRangeForSelectedModel(SwitchModel sm,Excel.ListRow LR)
        {
            
            string address = LR.Range.Address;
            int noOfVariables = 0;
            for (int count = 0; count < sm.VariableNames.Length; count++)
            {
                if (!string.IsNullOrEmpty(sm.VariableNames[count]))
                {
                    noOfVariables += 1;
                }
            }

            string[] addr = address.Split(new char[] { '$',':' });
            int destinationAddress = Convert.ToInt32(addr[2]) + noOfVariables;
            string targetAddress = "$" + addr[1] + addr[2] + ":" + "$" + addr[4] + destinationAddress.ToString();
            //Excel.Range target = energySheet.get_Range(targetAddress);
            return targetAddress;
        }
        //Modified by suman TFS Ticket :66435
        private void btnDefault_Click(object sender, EventArgs e)
        {
            foreach (SwitchModel sm in sModCol)
            {
                if (sm.Default)
                {

                    foreach (Excel.ListObject LO in  energySheet.ListObjects)
                    {
                        LO.Range.Interior.Color = 0xFFFFFF;
                        LO.Range.Font.Color = 0x000000;
                        LO.Range.EntireRow.Font.Bold = false;
                        LO.HeaderRowRange.Interior.Color = 0x000000;
                        LO.HeaderRowRange.Font.Color = 0xFFFFFF;
                        LO.HeaderRowRange.Font.Bold = true;
                        foreach (Excel.ListRow LR in LO.ListRows)
                        {
                            if (Convert.ToInt32(((Excel.Range)LR.Range[1]).Value2) == sm.ModelNumber)
                            {
                                string targetAddress = GetRangeForSelectedModel(sm, LR);
                                Excel.Range target = energySheet.get_Range(targetAddress);
                                //If the Model is appropriate for SEP is True then show the row highlighted in Green else Red
                                if (Convert.ToBoolean(((Excel.Range)LR.Range[1, 2]).Value2) == true)
                                {
                                    //LR.Range.EntireRow.Font.Bold = true;
                                    //LR.Range.Font.Color = 0x00AA00;
                                    //LR.Range.Interior.Color = 0xCDEFC6;
                                    target.EntireRow.Font.Bold = true;
                                    target.Font.Color = 0x00AA00;
                                    target.Interior.Color = 0xCDEFC6;
                                }
                                if (Convert.ToBoolean(((Excel.Range)LR.Range[1, 2]).Value2) == false)
                                {
                                    //LR.Range.EntireRow.Font.Bold = true;
                                    //LR.Range.Font.Color = 0x0000AA;
                                    //LR.Range.Interior.Color = 0xCEC8FF;
                                    target.EntireRow.Font.Bold = true;
                                    target.Font.Color = 0x0000AA;
                                    target.Interior.Color = 0xCEC8FF;
                                }
                            }
                            //    LR.Range.EntireRow.Font.Bold = true;
                            //else
                            //    LR.Range.EntireRow.Font.Bold = false;
                        }
                    }
                }

                if (sm.Default)
                    UpdateEnPISheet(sm);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.btnRun.Enabled = true;
            this.btnDefault.Enabled = true;
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(8);
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(9);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

        public string GetSEPValue(string independentVariable)
        {
            string retVal = string.Empty;
            foreach (SEPValidationValues val in Globals.ThisAddIn.lstSEPValidationValues)
            {
                if (val.IndependentVariable.Equals(independentVariable))
                {
                    retVal = val.SEPValidationCheck;
                }
            }
            return retVal;
        }

    }

    public class SwitchModel
    {
        public int ModelNumber { get; set; }
        public string[] VariableNames { get; set; }
        public double R2 { get; set; }
        public string Formula { get; set; }
        public bool Default { get; set; }
        public string SheetName { get; set; }
        public int index { get; set; }
        public double[] VariablePvalues { get; set; }
        public double AdjustedR2 { get; set; }
        public double ModlePvalue { get; set; }
        public SwitchModel()
        {
            ModelNumber = 0;
            VariableNames = null;
            R2 = 0;
            Formula = "";
            Default = false;
            SheetName = "";
            index = 0;
            VariablePvalues= null;
            AdjustedR2= 0;
            ModlePvalue = 0;

        }
        public SwitchModel(int ModelNumber, string[] VariableNames, double R2, string Formula, bool Default, string SheetName, int index, double[] VariablePvalues, double AdjustedR2, double ModlePvalue)
        {
            this.ModelNumber = ModelNumber;
            this.VariableNames = VariableNames;
            this.R2 = R2;
            this.Formula = Formula;
            this.Default = Default;
            this.SheetName = SheetName;
            this.index = index;
            this.VariablePvalues = VariablePvalues;
            this.AdjustedR2 = AdjustedR2 ;
            this.ModlePvalue = ModlePvalue;

        }
    }

    public class SwitchModelCollection : System.Collections.CollectionBase
    {
        public void Add(SwitchModel sModel)
        {
            List.Add(sModel);
        }

        public SwitchModel New()
        {
            SwitchModel sModel = new SwitchModel();
            int i = List.Add(sModel);
            sModel.ModelNumber = i + 1;

            return sModel;
        }

        public void Remove(int index)
        {
            try
            {
                List.RemoveAt(index);
            }
            catch
            {
            }
        }

        public SwitchModel Item(int Index)
        {
            return (SwitchModel)List[Index];
        }

    }
}
