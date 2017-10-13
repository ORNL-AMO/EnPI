using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;

namespace AMO.EnPI.AddIn
{
    class SEPSheet
    {
        public Utilities.EnPIDataSet DS;
        public Excel.Worksheet thisSheet;
        public Excel.Worksheet SourceSheet;
        public Excel.ListObject SourceObject;
        public object[,] SourceData;
        public string[,] AdjustmentMethod;
        public Excel.Worksheet AdjustedDataSheet;
        public Excel.ListObject AdjustedData;
        public Excel.ListObject SummaryData;
        public Excel.ListObject SEPFacilityIndentificationData;
        public Excel.ListObject ModelData;
        public Excel.ListObject WarningData;
        public Excel.ChartObject TotalVsModeledMMBTUChartObj;
        public Excel.ChartObject TrailingTwelveMonthEnPITTMChartObj;
        public Excel.ChartObject PrimaryTwelveMonthEnergyConsumptionChartObj;
        public Excel.ChartObject PrimaryTrailingTwelveMonthsEnergySavingsChartObj;
        public ArrayList Warnings;
        //public bool isSEnPI;
        //public Utilities.Model Model;
        //public ModelSheetCollection ModelSheets;
        //public ModelCollection ModelCollection;
        public int first = 0;
        public string strAdjustmentMethodColName;
      
        public  SEPSheet(Utilities.EnPIDataSet DSIn)
        {

            DS = DSIn;
            string nm = Globals.ThisAddIn.rsc.GetString("senpiTitle");
            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            SourceSheet = Utilities.ExcelHelpers.GetWorksheet(WB, DS.WorksheetName);

            Excel.Worksheet aSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add
                (System.Type.Missing, WB.Sheets.get_Item(WB.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet);
            aSheet.CustomProperties.Add("SheetGUID", System.Guid.NewGuid().ToString());

            aSheet.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(WB, nm, Globals.ThisAddIn.groupSheetCollection.regressionIteration);
            aSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            aSheet.PageSetup.PaperSize = Excel.XlPaperSize.xlPaperA4;
            aSheet.PageSetup.Orientation = Excel.XlPageOrientation.xlLandscape;
            //aSheet.DisplayPageBreaks = true;
            aSheet.Tab.Color = 0x008000;
            aSheet.PageSetup.TopMargin = aSheet.Application.InchesToPoints(1.0);
            aSheet.PageSetup.BottomMargin = aSheet.Application.InchesToPoints(1.0);
            aSheet.PageSetup.RightMargin = aSheet.Application.InchesToPoints(0.5);
            aSheet.PageSetup.LeftMargin = aSheet.Application.InchesToPoints(0.5);

            
            Utilities.ExcelHelpers.addWorksheetCustomProperty(aSheet, Utilities.Constants.WS_ISENPI, "True");
            thisSheet = aSheet;
            Warnings = new ArrayList();

          //  if (!isSEnPI)
            //    ExcelHelpers.addWorksheetCustomProperty(aSheet, Constants.WS_ROLLUP, "True");

        }


        public void Populate()
        {
            Excel.Range rangeTitle = (Excel.Range)thisSheet.get_Range("A1", "H1");
           ((Excel.Range)rangeTitle[1, 1]).Value2 = EnPIResources.senpiSheetTitle;
            ((Excel.Range)rangeTitle[1, 1]).Font.Color = 0x008000;
            ((Excel.Range)rangeTitle[1, 1]).Font.Bold = true;
            ((Excel.Range)rangeTitle[1, 1]).Font.Size = 15;
            rangeTitle.Merge();

            Excel.Range rangeBody = (Excel.Range)thisSheet.get_Range("A2", "H2");
           ((Excel.Range)rangeBody[1, 1]).Value2 = EnPIResources.senpiSheetText;
            rangeBody.Merge();
            rangeBody.WrapText = true;
            rangeBody.EntireRow.RowHeight = 70;

            SourceData = Utilities.DataHelper.dataTableArrayObject(DS.SourceData);
            SourceObject = Utilities.ExcelHelpers.GetListObject(SourceSheet, DS.ListObjectName);

            //As per the requirement on SEP Results
            Excel.Range rangeListHeader = BottomCell().get_Offset(2, 0);
          
            if (Globals.ThisAddIn.AdjustmentMethod != "Chaining")
            {

                rangeListHeader = rangeListHeader.get_Resize(1, 3);
                rangeListHeader.get_Offset(0, 1).get_Resize(1, 1).Value2 = "Baseline";
                rangeListHeader.get_Offset(0, 2).get_Resize(1, 1).Value2 = "Report Year";
                
            }
            else
            {

                if (Globals.ThisAddIn.ReportYear != Globals.ThisAddIn.ModelYearSelected)
                {
                    rangeListHeader = rangeListHeader.get_Resize(1, 4);
                    rangeListHeader.get_Offset(0, 1).get_Resize(1, 1).Value2 = "Baseline Year";
                    rangeListHeader.get_Offset(0, 2).get_Resize(1, 1).Value2 = "Model Year";
                    rangeListHeader.get_Offset(0, 3).get_Resize(1, 1).Value2 = "Report Year";
                }
                else
                {
                    rangeListHeader = rangeListHeader.get_Resize(1, 3);
                    rangeListHeader.get_Offset(0, 1).get_Resize(1, 1).Value2 = "Baseline Year";
                    rangeListHeader.get_Offset(0, 2).get_Resize(1, 1).Value2 = "Report/Model Year";
                    //rangeListHeader.get_Offset(0, 3).get_Resize(1, 1).Value2 = "Report Year";
                }
            }
            rangeListHeader.Cells.Interior.Color = 0x28624F;
            rangeListHeader.Cells.Font.Color = 0xFFFFFF;
            rangeListHeader.Cells.Font.Bold = true;
            
            AddTable();
            
            strAdjustmentMethodColName = EnPIResources.AdjustmentMethodSEPColName;

            AddSubtotalColumns();
            AddSEPFacilityIndentifyingInformation();
          
            FormatSummaryData();
            TotalVsModeledMMBTUChartObj = GenerateChartObj(true);
            if (Globals.ThisAddIn.AdjustmentMethod != EnPIResources.adjustmentBackcast)
            {
                TrailingTwelveMonthEnPITTMChartObj = GenerateChartObj();
                PrimaryTwelveMonthEnergyConsumptionChartObj = GenerateChartObj(true);
                PrimaryTrailingTwelveMonthsEnergySavingsChartObj = GenerateChartObj();

            }
           
            WriteCharts();

            AddVariableWarnings();

            modelInformation();
         
            GroupSheet GS = new GroupSheet(thisSheet, true, false, thisSheet.Name);
            Globals.ThisAddIn.groupSheetCollection.Add(GS);
            thisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            thisSheet.PageSetup.PrintArea = thisSheet.UsedRange.Address;
            thisSheet.Activate();
            
            
        }

        


        internal void AddVariableWarnings()
        {
            bool errorNotPresent = true;

            Excel.Range sumRange = BottomCell().get_Offset(2, 0).get_Resize(1, 12);
            sumRange.Merge();

            ((Excel.Range)sumRange[1, 1]).Value2 = "Warnings";
            ((Excel.Range)sumRange[1, 1]).Font.Color = 0xFFFFFF;
            ((Excel.Range)sumRange[1, 1]).Font.Bold = true;
            ((Excel.Range)sumRange[1, 1]).Interior.Color = 0x008000;

            AMO.EnPI.AddIn.Utilities.Model mdl;



            if (Globals.ThisAddIn.hasSEPValidationError)
                //Warnings.Add("Warning: The cells highlighted in red are out of the allowable range of the model year values. If the model is being evaluated during a period where it is not valid, please use a different model adjustment application method.");
                //Modified by Suman TFS Ticket : 66442
                //Warnings.Add("Warning: The cells highlighted in red on the “Model Data” sheet are out of the allowable range of the model year values. Meaning, the model cannot be used to predict the energy use for the time period shown in red if the variables shown in red are included in the model. It is recommended to select an alternative model which meets the R-squared and p-value requirements and does not include the variable shown in red in the model. If an alternative model cannot be selected with the current model year, try selecting an alternative model year. For more information, see the SEP Measurement and Verification Protocol.");
                Warnings.Add(Globals.ThisAddIn.sepValidationWarningMsg);


            foreach (string st in Warnings)
            {
                Excel.Range rg = BottomCell().get_Offset(1, 0).get_Resize(1, 12);
                rg.Merge();
                //Modified by Suman TFS Ticket : 66442
                rg.WrapText = true;
                rg.RowHeight = 75;
                rg.Value2 = st;
                rg.Font.Color = 0x0000AA;
                //rg.Style = "Bad";
            }

            sumRange.EntireRow.Hidden = errorNotPresent;
        }
       
        private Excel.Range BottomCell()
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(thisSheet, 0).ToString();

            return (Excel.Range)thisSheet.get_Range(addr, System.Type.Missing);
        }
        internal void AddSEPFacilityIndentifyingInformation()
        {
            try
            {
                //TODO:Remove hard coded text and add it to resource file.
                string summaryTableAddress = SummaryData.HeaderRowRange.Address;

                string summartStrCell = summaryTableAddress.Split(new char[] { ':' })[1];

                Excel.Range sepFacilityIndentificationRange = thisSheet.get_Range(summartStrCell).get_Offset(0, 2);
                sepFacilityIndentificationRange = sepFacilityIndentificationRange.get_Resize(1, 2);
                sepFacilityIndentificationRange.get_Offset(0, 0).get_Resize(1, 1).Value2 = "Facility identifying information";
                sepFacilityIndentificationRange.get_Offset(0, 1).get_Resize(1, 1).Value2 = "  ";
                //sepFacilityIndentificationRange.get_Offset(0, 2).get_Resize(1, 1).Value2 = "  ";
                SEPFacilityIndentificationData = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sepFacilityIndentificationRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
                SEPFacilityIndentificationData.Name = "SEPFacilityIdentificationData " + SEPFacilityIndentificationData.Name;
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Company Name");
                SEPFacilityIndentificationData.ListRows[2].Delete();
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Unique Facility Name per SEP Certificate");
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "SEP Enrollment Number");
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Facility zip code");
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Facility boundaries");
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Facility square footage (ft2)");
                AddNewRowToSEPFacilityIdentificationData(SEPFacilityIndentificationData, "Date SEP/ISO 50001 Stage 2 audit started (month/day/year)");
                SEPFacilityIndentificationData.TableStyle = "TableStyleMedium4";
                SEPFacilityIndentificationData.Range.WrapText = true;
                SEPFacilityIndentificationData.ShowAutoFilter = false;

                
                
            }
            catch { }
            
        }

        internal void AddTable()
        {
            Excel.Range sumRange = BottomCell().get_Offset(1, 0);

            sumRange.Value2 = Utilities.ExcelHelpers.GetListColumn(SourceObject, EnPIResources.yearColName).Name;
            //yr();
             
            object[] years = Utilities.ExcelHelpers.getYears(DS.SourceData);
            int ycols = years.Length;

            if (Globals.ThisAddIn.AdjustmentMethod != "Chaining")
            {
                sumRange = sumRange.get_Resize(1, 3);
                sumRange.get_Offset(0, 1).get_Resize(1, 1).Value2 = Globals.ThisAddIn.BaselineYear;
                sumRange.get_Offset(0, 2).get_Resize(1, 1).Value2 = Globals.ThisAddIn.ReportYear;
            }
            else
            {
                if (Globals.ThisAddIn.ReportYear != Globals.ThisAddIn.ModelYearSelected)
                {
                    sumRange = sumRange.get_Resize(1, 4);
                    sumRange.get_Offset(0, 1).get_Resize(1, 1).Value2 = Globals.ThisAddIn.BaselineYear;
                    sumRange.get_Offset(0, 2).get_Resize(1, 1).Value2 = Globals.ThisAddIn.ModelYearSelected;
                    sumRange.get_Offset(0, 3).get_Resize(1, 1).Value2 = Globals.ThisAddIn.ReportYear;
                    //((Excel.Range)sumRange.get_Offset(0, 3)).EntireColumn.Hidden = false;
                }
                else
                {
                    sumRange = sumRange.get_Resize(1, 3);
                    sumRange.get_Offset(0, 1).get_Resize(1, 1).Value2 = Globals.ThisAddIn.BaselineYear;
                    sumRange.get_Offset(0, 2).get_Resize(1, 1).Value2 = Globals.ThisAddIn.ModelYearSelected;
                    
                }

            }


            SummaryData = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sumRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            SummaryData.Name = "Annual" + SummaryData.Name;
            ((Excel.Range)SummaryData.Range[1, 1]).Value2 = " ";
            SummaryData.TableStyle = "TableStyleMedium4";

        }
        internal void AddSubtotalColumns()
        {
            Globals.ThisAddIn.Application.AutoCorrect.AutoFillFormulasInLists = false;

            string stylename = "Comma";
            //Unadjusted Data--------------------------
            //Unadjusted Fuel
            bool firstRow = true;
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                Excel.ListRow newRow = SummaryData.ListRows.Add();
                //for some reason, adding the first row to the ListRows added two rows...
                if (firstRow)
                   SummaryData.ListRows[2].Delete();
                firstRow = false;
                string name = es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                newRow.Range.Value2 = "=" + SubtotalRowFormula("SUMIF", name);
                ((Excel.Range)newRow.Range[1, 1]).Value2 = (EnPIResources.prefixSEPActual + name);
                newRow.Range.Style = stylename;
                ((Excel.Range)newRow.Range).Cells.Interior.Color = 0xBCE4D8;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Color = 0xFFFFFF;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Bold = true;
                newRow.Range.NumberFormat = "###,##0";
            }
            //Unadjusted Total
            if (!Utilities.Constants.MODEL_TOTAL)   // if total wasn't calculated as a source, it needs to be added here
            {
                Excel.ListRow newColumn = SummaryData.ListRows.Add();
                string name3 = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");
                newColumn.Range.Value2 = "=" + SubtotalRowFormula("SUMIF", name3);
                ((Excel.Range)newColumn.Range[1, 1]).Value2 = (Globals.ThisAddIn.rsc.GetString("unadjustedSEPTotalColName"));
                newColumn.Range.Style = stylename;
                newColumn.Range.NumberFormat = "###,##0";
            }
            #region Commented Code
            //Commented as there are no longer production and building data for Regression.
            //TFS Ticket : 71289
            //if(DS.ProductionVariables.Count > 0)
            //{
            //    Excel.ListRow prodOutput = SummaryData.ListRows.Add();
            //    string range = "";
            //    foreach (string prod in DS.ProductionVariables)
            //    {
            //        if (DS.ProductionVariables.IndexOf(prod).Equals(DS.ProductionVariables.Count - 1))
            //            range += SubtotalRowFormula("SUMIF", prod);
            //        else
            //            range += SubtotalRowFormula("SUMIF", prod) + " + ";
            //    }
            //    string name5 = "Total Production Output";
            //    prodOutput.Range.Value2 = "=" + range ;
            //    ((Excel.Range)prodOutput.Range[1, 1]).Value2 = name5;
            //    prodOutput.Range.EntireRow.Hidden = isSEnPI;
            //    prodOutput.Range.NumberFormat = "###,##0";
            //}

            //if (DS.ProductionVariables.Count > 0)
            //{
            //    addUnadjustedEnergyIntensity();
            //}

            //if (DS.BuildingVariables!=null && DS.BuildingVariables.Count > 0)
            //{
            //    addUnadjustedBuildingEnergyInten();
            //}
            #endregion
            Excel.ListRow spacerRow = SummaryData.ListRows.Add(System.Type.Missing);
            ((Excel.Range)spacerRow.Range).Interior.Color = 0x008000;
            spacerRow.Range.Value2 = " ";

            //Adjusted Data---------------------

            Excel.ListRow newRow5 = SummaryData.ListRows.Add(System.Type.Missing);
            ((Excel.Range)newRow5.Range).Cells.Interior.Color = 0xFFFFFF;
            ((Excel.Range)newRow5.Range[1, 1]).Cells.Interior.Color = 0x28624F;

            int modelIndex = 3;
            foreach (Excel.ListColumn LC in SummaryData.ListColumns)
            {
                int index = LC.Index;

                if (Globals.ThisAddIn.ModelYearSelected.Equals(LC.Name))
                    modelIndex = index;
            }
            foreach (Excel.ListColumn LC in SummaryData.ListColumns)
            {
                int index = LC.Index;

                if (modelIndex.Equals(2))
                    ((Excel.Range)newRow5.Range[1, index]).Value2 = Globals.ThisAddIn.rsc.GetString("adjustmentForecast");
                else if (modelIndex.Equals(SummaryData.ListColumns.Count))
                    ((Excel.Range)newRow5.Range[1, index]).Value2 = Globals.ThisAddIn.rsc.GetString("adjustmentBackcast");
                else
                    ((Excel.Range)newRow5.Range[1, index]).Value2 = Globals.ThisAddIn.rsc.GetString("adjustmentChaining");

                ((Excel.Range)newRow5.Range[1, index]).Font.Bold = true;
                newRow5.Range.NumberFormat = "###,##0";
            }

            ((Excel .Range)newRow5.Range[1, modelIndex]).Value2 = Globals.ThisAddIn.rsc.GetString("adjustmentModel");

            ((Excel.Range)newRow5.Range[1, 1]).Value2 = strAdjustmentMethodColName;

            //Adjusted Fuel
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {

                //Modeled Fuel Row
                Excel.ListRow newRow = SummaryData.ListRows.Add();
                string name2 = prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                newRow.Range.Value2 = "=" + SubtotalRowFormula("SUMIF", name2);
                ((Excel.Range)newRow.Range[1, 1]).Value2 = (EnPIResources.prefixSEPAdjusted) + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                newRow.Range.Style = stylename;
                ((Excel.Range)newRow.Range).Cells.Interior.Color = 0xFFFFFF;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                newRow.Range.NumberFormat = "###,##0";

                //Annual Savings Row
                Excel.ListRow newRow2 = SummaryData.ListRows.Add();

                string prefixActual = (EnPIResources.prefixSEPActual);
                string rowName = (EnPIResources.prefixSEPAdjusted) + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                //model = baseline
                if (modelIndex.Equals(2))
                    newRow2.Range.Value2 = "=" + AnnualSavingsRowFormula(1, newRow2, modelIndex, rowName, prefixActual + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));
                //model = last reporting year
                else if (modelIndex.Equals(SummaryData.ListColumns.Count))
                    newRow2.Range.Value2 = "=" + AnnualSavingsRowFormula(3, newRow2, modelIndex, rowName, prefixActual + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));
                //model > baseline & < last reporting year
                else
                    AnnualSavingsRowFormula(2, newRow2, modelIndex, rowName, prefixActual + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));

                string strAnnualSavingsTitle = (EnPIResources.prefixSEPAnnualSavings + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + " Savings");
                ((Excel.Range)newRow2.Range[1, 1]).Value2 = strAnnualSavingsTitle;
                ((Excel.Range)newRow2.Range).Cells.Interior.Color = 0xFFFFFF;
                ((Excel.Range)newRow2.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                newRow2.Range.NumberFormat = "###,##0";

                if (Globals.ThisAddIn.fromEnergyCost)
                {
                    //Estimated Cost savings Row
                    Excel.ListRow newRow3 = SummaryData.ListRows.Add();
                    //TFS Ticket 68851: Modified By Suman 
                    string name3 = "Cost Savings ($): " + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                    newRow3.Range.Value2 = "=" + SubtotalRowFormula("SUMIF", name3);
                    /*
                    //model = baseline
                    if (modelIndex.Equals(2))
                        newRow3.Range.Value2 = "=" + EstimatedCostSavingsRowFormula(1, newRow3, modelIndex, name2, es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));
                    //model = last reporting year
                    else if (modelIndex.Equals(SummaryData.ListColumns.Count))
                        newRow3.Range.Value2 = "=" + EstimatedCostSavingsRowFormula(3, newRow3, modelIndex, name2, es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));
                    //model > baseline & < last reporting year
                    else
                        EstimatedCostSavingsRowFormula(2, newRow3, modelIndex, name2, es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), ""));
                     */
                    ((Excel.Range)newRow3.Range[1, 1]).Value2 = es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + " Estimated Cost Savings ($)"; //TFS Ticket: 68853
                    ((Excel.Range)newRow3.Range).Cells.Interior.Color = 0xFFFFFF;
                    ((Excel.Range)newRow3.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                    newRow3.Range.NumberFormat = "_($* #,##0_);_($* (#,##0);_($* \"-\"??_);@_)";
                }

            }

            Excel.ListRow sumRow = SummaryData.ListRows.Add(System.Type.Missing);
            string name4 = Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName");
            sumRow.Range.Value2 = "=" + SubtotalRowFormula("SUMIF", name4);
            ((Excel.Range)sumRow.Range[1, 1]).Value2 = (Globals.ThisAddIn.rsc.GetString("totalAdjValuesSEPColName")); //"Total Modeled Energy Consumption (MMBTU)");
            sumRow.Range.Style = stylename;
            ((Excel.Range)sumRow.Range).Cells.Interior.Color = 0xFFFFFF;
            ((Excel.Range)sumRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
            sumRow.Range.NumberFormat = "###,##0";

            //setup string array for use with the rest of the calcualtions
            AdjustmentArraySetup(Utilities.ExcelHelpers.getYears(DS.SourceData));

            //calculate SEnPI
            Excel.ListRow senpiRow = SummaryData.ListRows.Add(System.Type.Missing);
            //string senpiName = "SEnPI Cumulative";
            string senpiName = "SEnPI";
            //model = baseline
            if (modelIndex.Equals(2))
                senpiRow.Range.Value2 = "=" + SEnPI(1, senpiRow, modelIndex);
            //model = last reporting year
            else if (modelIndex.Equals(SummaryData.ListColumns.Count))
                senpiRow.Range.Value2 = "=" + SEnPI(3, senpiRow, modelIndex);
            //model > baseline & < last reporting year
            else
                SEnPI(2, senpiRow, modelIndex);
            senpiRow.Range.Style = "Comma";
            senpiRow.Range.NumberFormat = "###,##0.000";
            ((Excel.Range)senpiRow.Range[1, 1]).Value2 = senpiName;
            ((Excel.Range)senpiRow.Range[1, modelIndex]).Value2 = "= 1";
            senpiRow.Range.EntireRow.Hidden = false;


            //calculate Cumulative Improvement
            Excel.ListRow cumulativeImprovRow = SummaryData.ListRows.Add(System.Type.Missing);

            string ciName = EnPIResources.totalImprovementSEPColName;
            if (modelIndex.Equals(2) || modelIndex.Equals(SummaryData.ListColumns.Count))
                cumulativeImprovRow.Range.Value2 = VaryingRowFormula(modelIndex, cumulativeImprovRow, Utilities.Constants.BEFORE_MODEL_CUMULATIVE_IMPROVMENT, Utilities.Constants.AFTER_MODEL_CUMULATIVE_IMPROVMENT);
            else
                VaryingRowFormula(modelIndex, cumulativeImprovRow, Utilities.Constants.BEFORE_MODEL_CUMULATIVE_IMPROVMENT, Utilities.Constants.AFTER_MODEL_CUMULATIVE_IMPROVMENT);
            cumulativeImprovRow.Range.Style = "Percent";
            ((Excel.Range)cumulativeImprovRow.Range[1, 1]).Value2 = ciName;
            ((Excel.Range)cumulativeImprovRow.Range[1, 2]).Value2 = 0;
            cumulativeImprovRow.Range.NumberFormat = "0.00%";

            /*   //calculate Annual Improvement
               if (!isSEnPI)
               {
                   Excel.ListRow annualImprovRow = SummaryData.ListRows.Add(System.Type.Missing);
                   string aiName = "";
                   if (!isSEnPI)
                       aiName = EnPIResources.annualImprovementColName;
                   else
                       aiName = EnPIResources.annualImprovementSEnPIColName;
                   if (modelIndex.Equals(2) || modelIndex.Equals(SummaryData.ListColumns.Count))
                       annualImprovRow.Range.Value2 = VaryingRowFormula(modelIndex, annualImprovRow, Utilities.Constants.BEFORE_MODEL_ANNUAL_IMPROVMENT, Utilities.Constants.AFTER_MODEL_ANNUAL_IMPROVMENT);
                   else
                       VaryingRowFormula(modelIndex, annualImprovRow, Utilities.Constants.BEFORE_MODEL_ANNUAL_IMPROVMENT, Utilities.Constants.AFTER_MODEL_ANNUAL_IMPROVMENT);
                   annualImprovRow.Range.Style = "Percent";
                   ((Excel.Range)annualImprovRow.Range[1, 1]).Value2 = aiName;
                   ((Excel.Range)annualImprovRow.Range[1, 2]).Value2 = 0;
                   annualImprovRow.Range.NumberFormat = "0.00%";
                   annualImprovRow.Range.EntireRow.Hidden = isSEnPI; // TFS Ticket 77017

                   //Caclulate Annual Savings
                   string unadjustedTotalColName = ((isSEnPI == false) ? Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") : Globals.ThisAddIn.rsc.GetString("unadjustedSEPTotalColName"));
                   string before = "OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1,1,1) + ((INDEX(" + SummaryData.Name + ",MATCH(\"" + unadjustedTotalColName + "\"," + SummaryData.Name + "[[ ]],0),COLUMN()-1)-INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + "\"," + SummaryData.Name + "[[ ]],0),COLUMN()-1))-(INDEX(" + SummaryData.Name + ",MATCH(\"" + unadjustedTotalColName + "\"," + SummaryData.Name + "[[ ]],0),)-INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + "\"," + SummaryData.Name + "[[ ]],0),)))";
                   string after = "INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + "\"," + SummaryData.Name + "[[ ]],0),)-INDEX(" + SummaryData.Name + ",MATCH(\"" + unadjustedTotalColName + "\"," + SummaryData.Name + "[[ ]],0),)";
                   Excel.ListRow annualSavingsRow = SummaryData.ListRows.Add(System.Type.Missing);
                   if (modelIndex.Equals(2) || modelIndex.Equals(SummaryData.ListColumns.Count))
                       annualSavingsRow.Range.Value2 = VaryingRowFormula(modelIndex, annualSavingsRow, before, after);
                   else
                       VaryingRowFormula(modelIndex, annualSavingsRow, before, after);
                   annualSavingsRow.Range.Style = "Comma";
                   annualSavingsRow.Range.NumberFormat = "###,##0";
                   if (!isSEnPI)
                       ((Excel.Range)annualSavingsRow.Range[1, 1]).Value2 = "Total Energy Savings since Baseline Year (MMBtu/Year)";
                   else
                       ((Excel.Range)annualSavingsRow.Range[1, 1]).Value2 = "Annual Savings (MMBtu/year)";
                   ((Excel.Range)annualSavingsRow.Range).Cells.Interior.Color = 0xFFFFFF;
                   ((Excel.Range)annualSavingsRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                   ((Excel.Range)annualSavingsRow.Range[1, 2]).Value2 = 0;
                   annualSavingsRow.Range.EntireRow.Hidden = isSEnPI; // TFS Ticket 77017



                   //Calculate Cumulative savings
                   Excel.ListRow cumulativeSavingsRow = SummaryData.ListRows.Add(System.Type.Missing);
                   //cumulativeSavingsRow.Range.Value2 = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1,1,1)";
                   cumulativeSavingsRow.Range.Value2 = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1,1,1) + OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1)";
                   cumulativeSavingsRow.Range.Style = "Comma";
                   cumulativeSavingsRow.Range.NumberFormat = "###,##0";
                   ((Excel.Range)cumulativeSavingsRow.Range[1, 1]).Value2 = "Cumulative Savings (MMBtu)";
                   ((Excel.Range)cumulativeSavingsRow.Range).Cells.Interior.Color = 0xFFFFFF;
                   ((Excel.Range)cumulativeSavingsRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                   ((Excel.Range)cumulativeSavingsRow.Range[1, 2]).Value2 = 0;
                   cumulativeSavingsRow.Range.EntireRow.Hidden = isSEnPI;// true; // TFS Ticket 77017


                   //Calculate 
                   Excel.ListRow newEnergySavingsRow = SummaryData.ListRows.Add(System.Type.Missing);
                   newEnergySavingsRow.Range.Value2 = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-2,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-2,-1,1,1)";
                   newEnergySavingsRow.Range.Style = "Comma";
                   newEnergySavingsRow.Range.NumberFormat = "###,##0";
                   ((Excel.Range)newEnergySavingsRow.Range[1, 1]).Value2 = "New Energy Savings for Current Year (MMBtu/year)";
                   ((Excel.Range)newEnergySavingsRow.Range).Cells.Interior.Color = 0xFFFFFF;
                   ((Excel.Range)newEnergySavingsRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                   ((Excel.Range)newEnergySavingsRow.Range[1, 2]).Value2 = 0;
                   newEnergySavingsRow.Range.EntireRow.Hidden = isSEnPI;

                   //Calculate 
                   Excel.ListRow adjustmentforBaselineRow = SummaryData.ListRows.Add(System.Type.Missing);
                   adjustmentforBaselineRow.Range.Value2 = "=(INDEX(" + SummaryData.Name + ",MATCH(\"" + unadjustedTotalColName + "\"," + SummaryData.Name + "[[ ]],0),) + OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-3,0,1,1)) - (INDEX(" + SummaryData.Name + ",MATCH(\"" + unadjustedTotalColName + "\"," + SummaryData.Name + "[[ ]],0),2))";
                   adjustmentforBaselineRow.Range.Style = "Comma";
                   adjustmentforBaselineRow.Range.NumberFormat = "###,##0";
                   ((Excel.Range)adjustmentforBaselineRow.Range[1, 1]).Value2 = "Adjustment for Baseline Primary Energy Use (MMBtu/year)";
                   ((Excel.Range)adjustmentforBaselineRow.Range).Cells.Interior.Color = 0xFFFFFF;
                   ((Excel.Range)adjustmentforBaselineRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                   ((Excel.Range)adjustmentforBaselineRow.Range[1, 2]).Value2 = 0;
                   adjustmentforBaselineRow.Range.EntireRow.Hidden = isSEnPI;
               }
               else
               {
                   //TODO: This condition is fixing the issue of SEP result sheet not getting generated when the cost and co2 emissions are selected
                   // This empty is causing problems if it is removed , the SEP results is not getting generated when the above two options are not selected.
                   if ((Globals.ThisAddIn.fromEnergyCost == false) && (Globals.ThisAddIn.fromCO2Emission == false))
                   {
                       //Add a blank row
                       Excel.ListRow blankRow = SummaryData.ListRows.Add(System.Type.Missing);
                       ((Excel.Range)blankRow.Range).Cells.Interior.Color = 0xFFFFFF;
                       blankRow.Range.Value2 = string.Empty;
                   }
               }


               //}
               */
            //Added By BJV:TFS Ticket 66432
            //Estimated Annual cost savings
            if (Globals.ThisAddIn.fromEnergyCost)
            {
                string estimatedAnnualCostSavingsFormula = string.Empty;
                foreach (Utilities.EnergySource es in DS.EnergySources)
                {
                    if (!es.Name.Contains("TOTAL")) // Need to find out a way to eliminate total column here 
                        estimatedAnnualCostSavingsFormula += "INDEX(" + SummaryData.Name + ",MATCH(\"" + es.Name + " Estimated Cost Savings ($)" + "\"," + "[[ ]],0),COLUMN())+"; //TFS Ticket: 68853
                }
                estimatedAnnualCostSavingsFormula = estimatedAnnualCostSavingsFormula.Remove(estimatedAnnualCostSavingsFormula.Length - 1, 1);
                AddNewRowToSummaryData(SummaryData, "Estimated Annual Cost Savings", estimatedAnnualCostSavingsFormula, stylename, "_($* #,##0_);_($* (#,##0);_($* " + " -" + "??_);_(@_)");
            }

            //Added By Suman: TFS Ticket 68998
            //CO2 Avoided Emissions
            if (Globals.ThisAddIn.fromCO2Emission)
            {
                string co2EmissionFormula = string.Empty;
                foreach (Utilities.EnergySource es in DS.EnergySources)
                {
                    if (!es.Name.Contains("TOTAL")) // Need to find out a way to eliminate total column here 
                    {
                        string emissionFactor;
                        Globals.ThisAddIn.CO2EmissionFactors.TryGetValue(es.Name, out emissionFactor);
                        //just in case not to break the code 
                        emissionFactor = ((!string.IsNullOrEmpty(emissionFactor) ? emissionFactor : "1"));
                        string strAnnualSavingsTitle = (EnPIResources.prefixSEPAnnualSavings + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + " Savings");
                        co2EmissionFormula += "INDEX(" + SummaryData.Name + ",MATCH(\"" + strAnnualSavingsTitle + "\"," + "[[ ]],0),COLUMN())*" + emissionFactor + "/1000+"; //TFS Ticket: 68853
                    }
                }
                co2EmissionFormula = co2EmissionFormula.Remove(co2EmissionFormula.Length - 1, 1);
                AddNewRowToSummaryData(SummaryData, "Avoided CO2 Emissions (Metric Ton/year)", co2EmissionFormula, stylename, "###,##0"); //TFS Ticket: 70385

            }


        }
        #region Charting
        internal void WriteCharts()
        {
            try
            {
                GenerateTotalVsModeledMMBTUChart();
                if (Globals.ThisAddIn.AdjustmentMethod != EnPIResources.adjustmentBackcast)
                {
                    GenerateTrailingTwelveMonthEnPITTMChart();
                    GeneratePrimaryTwelveMonthEnergyConsumptionChart();
                    GeneratePrimaryTrailingTwelveMonthsEnergySavingsChart();
                }
            }
            catch { }
        }
        internal void GenerateTotalVsModeledMMBTUChart()
        {
            //Modified By Suman TFS Ticket:68735
            TotalVsModeledMMBTUChartObj.Chart.ChartType = Excel.XlChartType.xlLineMarkers;
            TotalVsModeledMMBTUChartObj.Chart.ChartStyle = 5;//2;
            TotalVsModeledMMBTUChartObj.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
            TotalVsModeledMMBTUChartObj.Chart.ChartArea.Width = thisSheet.Application.InchesToPoints(8.5);
            
            Excel.ListObject lo = AdjustedData;
            Excel.ListObject lo2 = SummaryData;
            string tot = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");
            // find baseYear source sum column
            int i = lo.ListColumns.Count;
            int iadj = i;

            foreach (Excel.ListColumn col in lo.ListColumns)
            {
                if (col.Name == tot) i = col.Index;
                //if (col.Name == prefix() + tot) iadj = col.Index;
                string strTotalAdjValues = Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName");
                if (col.Name == strTotalAdjValues) iadj = col.Index;
            }

            lo.ListColumns[i].TotalsCalculation = Excel.
            XlTotalsCalculation.xlTotalsCalculationSum;
            lo.ListColumns[iadj].TotalsCalculation = Excel.
            XlTotalsCalculation.xlTotalsCalculationSum;
            int ct = lo.ListRows.Count;
            object[] xrow = new object[ct];

            int k = 1;
            foreach (Excel.ListRow r in lo.ListRows)
            {
                for (int j = 0; j < lo.ListRows.Count; j++)
                {
                    xrow[j] = k;
                    k++;
                }
                break;
            }

            ((Excel.SeriesCollection)TotalVsModeledMMBTUChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(lo.ListColumns[iadj].Range
            , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrow;


            ((Excel.SeriesCollection)TotalVsModeledMMBTUChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(lo.ListColumns[i].Range
            , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrow;


            //TFS Ticket: 66436 - Added by Suman.
            Excel.SeriesCollection chartObjSeriesColleciton = (Excel.SeriesCollection)TotalVsModeledMMBTUChartObj.Chart.SeriesCollection(System.Type.Missing);
            Excel.Series xSeries = chartObjSeriesColleciton.Item(2);
            xSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DodgerBlue);//  0x4177B8;//12089153; // Calculated from color calculator
            xSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStylePlus;
            //TFS ticket :69332
            xSeries.MarkerBackgroundColorIndex = (Microsoft.Office.Interop.Excel.XlColorIndex)25;
            xSeries.Format.Line.Transparency = 1.0f;

            chartObjSeriesColleciton.Item(1).Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green);
          
                Excel.Series totalSeries = chartObjSeriesColleciton.Item(1);
                totalSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleNone;
                totalSeries.Name = "Total Modeled Energy Consumption (MMBTU)";
                totalSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
                xSeries.Name = "TOTAL (MMBTU)";




             Excel.Axis ChartObjYaxis = (Excel.Axis)TotalVsModeledMMBTUChartObj.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            ChartObjYaxis.HasTitle = true;
            ChartObjYaxis.AxisTitle.Text = "Total Consumption (MMBTU)";


            Excel.Axis ChartObjXaxis = (Excel.Axis)TotalVsModeledMMBTUChartObj.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            ChartObjXaxis.HasTitle = true;

            ChartObjXaxis.AxisTitle.Text =  "Time Period";
            ChartObjXaxis.TickLabelSpacing = 6;
            ChartObjXaxis.TickMarkSpacing = 6;
            int enInten = -1;
            int SEnPI = -1;
            int buildEnInten = -1;
            int anualImprovement = -1;
            int totalImprovment = -1;
            //int year = -1;

            foreach (Excel.ListRow row in lo2.ListRows)
            {
                string test = ((Excel.Range)row.Range[0, 1]).Value2.ToString();

                if (((Excel.Range)row.Range[0, 1]).Value2 != null)
                {
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.unadjustedEnergyIntensColName)) enInten = row.Index - 1;
                    //Uncommented below statement as per TFS Ticket 68850
                    //if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals("SEnPI Cumulative")) SEnPI = row.Index - 1; 
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals("SEnPI")) SEnPI = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.annualImprovementColName) || ((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.annualImprovementSEnPIColName)) anualImprovement = row.Index - 1;
                    //if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.totalImprovementColName) || ((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.totalImprovementSEnPIColName)) totalImprovment = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.totalImprovementColName) || ((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.totalImprovementSEPColName)) totalImprovment = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.unadjustedBuildingColName)) buildEnInten = row.Index - 1;
                }
            }

           
        }
        internal void GenerateTrailingTwelveMonthEnPITTMChart()
        {
            try
            {
                TrailingTwelveMonthEnPITTMChartObj.Chart.ChartType = Excel.XlChartType.xlXYScatter;
                TrailingTwelveMonthEnPITTMChartObj.Chart.ChartStyle = 6;
                TrailingTwelveMonthEnPITTMChartObj.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                TrailingTwelveMonthEnPITTMChartObj.Chart.HasTitle = true;
                TrailingTwelveMonthEnPITTMChartObj.Chart.ChartTitle.Text = string.Empty;
                TrailingTwelveMonthEnPITTMChartObj.Chart.ChartArea.Width = thisSheet.Application.InchesToPoints(8.5);

                Excel.ListObject lo = AdjustedData;
                string sepTTMColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthEnergyPerformanceIndicator");

                int i = lo.ListColumns.Count;

                foreach (Excel.ListColumn col in lo.ListColumns)
                {
                    if (col.Name == sepTTMColName) i = col.Index;

                }

                lo.ListColumns[i].TotalsCalculation = Excel.
                XlTotalsCalculation.xlTotalsCalculationSum;

                Excel.Range rngListColumnVals;
                object[] xrow;

                GetXYValuesForChart(lo, i, out rngListColumnVals, out xrow);



                ((Excel.SeriesCollection)TrailingTwelveMonthEnPITTMChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngListColumnVals
                , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrow;


                //TFS Ticket: 66436 - Added by Suman.
                Excel.SeriesCollection chartObjSeriesColleciton = (Excel.SeriesCollection)TrailingTwelveMonthEnPITTMChartObj.Chart.SeriesCollection(System.Type.Missing);
                Excel.Series totalSeries = chartObjSeriesColleciton.Item(1);
                totalSeries.Name = ((Globals.ThisAddIn.AdjustmentMethod == EnPIResources.adjustmentForecast)? "Energy Performance Indicator (forecast)":"Energy Performance Indicator (chaining)");
                totalSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleCircle;
                totalSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);


                Excel.Axis ChartObjYaxis = (Excel.Axis)TrailingTwelveMonthEnPITTMChartObj.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                ChartObjYaxis.HasTitle = true;
                ChartObjYaxis.HasMajorGridlines = true;
                //ChartObjYaxis.MajorUnit = 0.01;
                ChartObjYaxis.AxisTitle.Text = "Trailing Twelve Month EnPI(TTM)";


                Excel.Axis ChartObjXaxis = (Excel.Axis)TrailingTwelveMonthEnPITTMChartObj.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                ChartObjXaxis.HasTitle = true;
                ChartObjXaxis.HasMajorGridlines = true;
                ChartObjXaxis.MinimumScale = Convert.ToDouble(xrow[0]);

                ChartObjXaxis.AxisTitle.Text = "Input Interval";
                //ChartObjXaxis.TickLabelSpacing = 6;
                //ChartObjXaxis.TickMarkSpacing = 6;
            }
            catch{
               

            }



        }

        
        internal void GeneratePrimaryTwelveMonthEnergyConsumptionChart()
        {

            try
            {
                PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.ChartType = Excel.XlChartType.xlXYScatter;
                PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.ChartStyle = 2;
                PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.ChartArea.Width = thisSheet.Application.InchesToPoints(8.5);

                Excel.ListObject lo = AdjustedData;
                string sepTTMActEngConColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthActualEnergyConsumption");
                string sepTTMActEngConFivPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthActualEnergyConsumptionFivePerTarget");
                string sepTTMActEngConTenPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthActualEnergyConsumptionTenPerTarget");
                string sepTTMActEngConFifPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget");

                int sepTTMActEngConColIndex = lo.ListColumns.Count;
                int sepTTMActEngConFivPerTarColIndex = lo.ListColumns.Count;
                int sepTTMActEngConTenPerTarColIndex = lo.ListColumns.Count;
                int sepTTMActEngConFifPerTarColIndex = lo.ListColumns.Count;

                foreach (Excel.ListColumn col in lo.ListColumns)
                {
                    if (col.Name == sepTTMActEngConColName) sepTTMActEngConColIndex = col.Index;
                    if (col.Name == sepTTMActEngConFivPerTarColName) sepTTMActEngConFivPerTarColIndex = col.Index;
                    if (col.Name == sepTTMActEngConTenPerTarColName) sepTTMActEngConTenPerTarColIndex = col.Index;
                    if (col.Name == sepTTMActEngConFifPerTarColName) sepTTMActEngConFifPerTarColIndex = col.Index;

                }

                lo.ListColumns[sepTTMActEngConColIndex].TotalsCalculation = Excel.
                XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngConFivPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngConTenPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngConFifPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;

                // Series "Actual Energy Consumption (TTM)"
                Excel.Range rngSEPTTMActEngConColVals;
                object[] xrowSEPTTMActEngConCol;
                GetXYValuesForChart(lo, sepTTMActEngConColIndex, out rngSEPTTMActEngConColVals, out xrowSEPTTMActEngConCol);
                ((Excel.SeriesCollection)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngConColVals
                , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngConCol;

                Excel.SeriesCollection chartObjSeriesColleciton = (Excel.SeriesCollection)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.SeriesCollection(System.Type.Missing);
                Excel.Series sepTTMActEngConSeries = chartObjSeriesColleciton.Item(1);
                sepTTMActEngConSeries.Name = "Actual Energy Consumption (TTM)";
                sepTTMActEngConSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleX;
                sepTTMActEngConSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x7D60A0"));//System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Purple);
                
                //Series "Actual Energy Consumption to meet 5% target"
                Excel.Range rngSEPTTMActEngConFivPerTarColVals;
                object[] xrowSEPTTMActEngConFivPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngConFivPerTarColIndex, out rngSEPTTMActEngConFivPerTarColVals, out xrowSEPTTMActEngConFivPerTarCol);
                ((Excel.SeriesCollection)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngConFivPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngConFivPerTarCol;
                Excel.Series sepTTMActEngConFivPerTarSeries = chartObjSeriesColleciton.Item(2);
                sepTTMActEngConFivPerTarSeries.Name = "Actual Energy Consumption to meet 5% target";
                sepTTMActEngConFivPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleDiamond;
                sepTTMActEngConFivPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x4F81BD"));
                sepTTMActEngConFivPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x4F81BD"));

                //Series "Actual Energy Consumption to meet 10% target"
                Excel.Range rngSEPTTMActEngConTenPerTarColVals;
                object[] xrowSEPTTMActEngConTenPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngConTenPerTarColIndex, out rngSEPTTMActEngConTenPerTarColVals, out xrowSEPTTMActEngConTenPerTarCol);
                ((Excel.SeriesCollection)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngConTenPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngConTenPerTarCol;
                Excel.Series sepTTMActEngConTenPerTarSeries = chartObjSeriesColleciton.Item(3);
                sepTTMActEngConTenPerTarSeries.Name = "Actual Energy Consumption to meet 10% target";
                sepTTMActEngConTenPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;
                sepTTMActEngConTenPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0xC0504D"));
                sepTTMActEngConTenPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0xC0504D"));

                //Series "Actual Energy Consumption to meet 15% target"
                Excel.Range rngSEPTTMActEngConFifPerTarColVals;
                object[] xrowSEPTTMActEngConFifPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngConFifPerTarColIndex, out rngSEPTTMActEngConFifPerTarColVals, out xrowSEPTTMActEngConFifPerTarCol);
                ((Excel.SeriesCollection)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngConFifPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngConFifPerTarCol;
                Excel.Series sepTTMActEngConFifPerTarSeries = chartObjSeriesColleciton.Item(4);
                sepTTMActEngConFifPerTarSeries.Name = "Actual Energy Consumption to meet 15% target";
                sepTTMActEngConFifPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle;
                sepTTMActEngConFifPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x9BBB59"));
                sepTTMActEngConFifPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x9BBB59"));

                // Y Axis
                Excel.Axis ChartObjYaxis = (Excel.Axis)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                ChartObjYaxis.HasTitle = true;
                ChartObjYaxis.HasMajorGridlines = true;
                ChartObjYaxis.AxisTitle.Text = "Primary 12 months energy consumption across all energy sources (MMBtu)";


                Excel.Axis ChartObjXaxis = (Excel.Axis)PrimaryTwelveMonthEnergyConsumptionChartObj.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                ChartObjXaxis.HasTitle = true;
                ChartObjXaxis.HasMajorGridlines = true;
                ChartObjXaxis.MinimumScale = Convert.ToDouble(xrowSEPTTMActEngConCol[0]);
                ChartObjXaxis.AxisTitle.Text = "Input Interval";
                //ChartObjXaxis.TickLabelSpacing = 6;
                //ChartObjXaxis.TickMarkSpacing = 6;
            }
            catch { }
        }
        internal void GeneratePrimaryTrailingTwelveMonthsEnergySavingsChart()
        {
            try
            {
                PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.ChartType = Excel.XlChartType.xlXYScatter;
                PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.ChartStyle = 2;
                PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
                PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.ChartArea.Width = thisSheet.Application.InchesToPoints(8.5);

                Excel.ListObject lo = AdjustedData;
                string sepEnrSavTTMColName = Globals.ThisAddIn.rsc.GetString("sepEnergySavingsTTM");
                string sepTTMActEngSavFivPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthEnergySavingsFivePerImprovement");
                string sepTTMActEngSavTenPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthEnergySavingsTenPerImprovement");
                string sepTTMActEngSavFifPerTarColName = Globals.ThisAddIn.rsc.GetString("sepTrailingTwelveMonthEnergySavingsFifteenPerImprovement");

                int sepEnrSavTTMColIndex = lo.ListColumns.Count;
                int sepTTMActEngSavFivPerTarColIndex = lo.ListColumns.Count;
                int sepTTMActEngSavTenPerTarColIndex = lo.ListColumns.Count;
                int sepTTMActEngSavFifPerTarColIndex = lo.ListColumns.Count;

                foreach (Excel.ListColumn col in lo.ListColumns)
                {
                    if (col.Name == sepEnrSavTTMColName) sepEnrSavTTMColIndex = col.Index;
                    if (col.Name == sepTTMActEngSavFivPerTarColName) sepTTMActEngSavFivPerTarColIndex = col.Index;
                    if (col.Name == sepTTMActEngSavTenPerTarColName) sepTTMActEngSavTenPerTarColIndex = col.Index;
                    if (col.Name == sepTTMActEngSavFifPerTarColName) sepTTMActEngSavFifPerTarColIndex = col.Index;

                }

                lo.ListColumns[sepEnrSavTTMColIndex].TotalsCalculation = Excel.
                XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngSavFivPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngSavTenPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;
                lo.ListColumns[sepTTMActEngSavFifPerTarColIndex].TotalsCalculation = Excel.
               XlTotalsCalculation.xlTotalsCalculationSum;



                //Series "Energy savings TTM (MMBTu)"
                Excel.Range rngSEPEnrSavTTMColVals;
                object[] xrowSEPEnrSavTTMCol;
                GetXYValuesForChart(lo, sepEnrSavTTMColIndex, out rngSEPEnrSavTTMColVals, out xrowSEPEnrSavTTMCol);
                ((Excel.SeriesCollection)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPEnrSavTTMColVals
                , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPEnrSavTTMCol;
                Excel.SeriesCollection chartObjSeriesColleciton = (Excel.SeriesCollection)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.SeriesCollection(System.Type.Missing);
                Excel.Series sepEnrSavTTMSeries = chartObjSeriesColleciton.Item(1);
                sepEnrSavTTMSeries.Name = "Energy savings TTM (MMBTu)";
                sepEnrSavTTMSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleX;
                sepEnrSavTTMSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x7D60A0"));
              



                //Series "5% improvement savings"
                Excel.Range rngSEPTTMActEngSavFivPerTarColVals;
                object[] xrowSEPTTMActEngSavFivPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngSavFivPerTarColIndex, out rngSEPTTMActEngSavFivPerTarColVals, out xrowSEPTTMActEngSavFivPerTarCol);
                ((Excel.SeriesCollection)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngSavFivPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngSavFivPerTarCol;
                Excel.Series sepTTMActEngSavFivPerTarSeries = chartObjSeriesColleciton.Item(2);
                sepTTMActEngSavFivPerTarSeries.Name = "5% improvement savings";
                sepTTMActEngSavFivPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleDiamond;
                sepTTMActEngSavFivPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x4F81BD"));
                sepTTMActEngSavFivPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x4F81BD"));

                //Series "10% improvement savings"
                Excel.Range rngSEPTTMActEngSavTenPerTarColVals;
                object[] xrowSEPTTMActEngSavTenPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngSavTenPerTarColIndex, out rngSEPTTMActEngSavTenPerTarColVals, out xrowSEPTTMActEngSavTenPerTarCol);
                ((Excel.SeriesCollection)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngSEPTTMActEngSavTenPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowSEPTTMActEngSavTenPerTarCol;
                Excel.Series sepTTMActEngSavTenPerTarSeries = chartObjSeriesColleciton.Item(3);
                sepTTMActEngSavTenPerTarSeries.Name = "10% improvement savings";
                sepTTMActEngSavTenPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;
                sepTTMActEngSavTenPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0xC0504D"));
                sepTTMActEngSavTenPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0xC0504D"));


                //Series  "15% improvement savings"
                Excel.Range rngsepTTMActEngSavFifPerTarColVals;
                object[] xrowsepTTMActEngSavFifPerTarCol;
                GetXYValuesForChart(lo, sepTTMActEngSavFifPerTarColIndex, out rngsepTTMActEngSavFifPerTarColVals, out xrowsepTTMActEngSavFifPerTarCol);
                ((Excel.SeriesCollection)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(rngsepTTMActEngSavFifPerTarColVals
               , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing).XValues = xrowsepTTMActEngSavFifPerTarCol;
                Excel.Series sepTTMActEngSavFifPerTarSeries = chartObjSeriesColleciton.Item(4);
                sepTTMActEngSavFifPerTarSeries.Name = "15% improvement savings";
                sepTTMActEngSavFifPerTarSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleTriangle;
                sepTTMActEngSavFifPerTarSeries.MarkerForegroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x9BBB59"));
                sepTTMActEngSavFifPerTarSeries.MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.ColorTranslator.FromHtml("0x9BBB59"));



                // Y Axis
                Excel.Axis ChartObjYaxis = (Excel.Axis)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                ChartObjYaxis.HasTitle = true;
                ChartObjYaxis.HasMajorGridlines = true;

                ChartObjYaxis.AxisTitle.Text = "Primary trailing twelve months of energy savings (MMBtu)";


                Excel.Axis ChartObjXaxis = (Excel.Axis)PrimaryTrailingTwelveMonthsEnergySavingsChartObj.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                ChartObjXaxis.HasTitle = true;
                ChartObjXaxis.HasMajorGridlines = true;
                ChartObjXaxis.MinimumScale = Convert.ToDouble(xrowSEPTTMActEngSavFivPerTarCol[0]);
                ChartObjXaxis.AxisTitle.Text = "Input Interval";
            }
            catch { }
        }
       
        internal Excel.ChartObject GenerateChartObj(bool insertPageBreak = false)
        {
            Excel.Range start = BottomCell().get_Offset(2, 0);
            if (insertPageBreak == true)
            {
                thisSheet.HPageBreaks.Add(start);
            }

            double topleft;
            if (!double.TryParse(start.Top.ToString(), out topleft))
                topleft = 0;

            start.EntireRow.RowHeight = Utilities.Constants.CHART_HEIGHT * 1.1;

            Excel.ChartObject CO = ((Excel.ChartObjects)thisSheet.ChartObjects(System.Type.Missing))
                .Add(10, topleft, Utilities.Constants.CHART_WIDTH * 2, Utilities.Constants.CHART_HEIGHT);
            CO.Placement = Excel.XlPlacement.xlMove;

            return CO;
        }
        private void GetXYValuesForChart(Excel.ListObject lo, int i, out Excel.Range rngListColumnVals, out object[] xrow)
        {

            try
            {
                string colRangeAddress = lo.ListColumns[i].DataBodyRange.Address.ToString();


                string endAddress = colRangeAddress.Split(':')[1];
                string startAddress = string.Empty;


                Excel.Range colRange = AdjustedDataSheet.get_Range(colRangeAddress);
                foreach (Excel.Range cell in colRange.Cells)
                {
                    string val = cell.Value.ToString();
                    if (val != "N/A")
                   {
                        // Hack:The reason why the offset is used here is, if the Start address is AC26, the graph series is picking the range from AC27, so in order to include AC26 
                        // added an offset of row -1 and passing the address to the series.
                        startAddress = cell.get_Offset(-1,0).Address;
                        break;
                   }

                }
                rngListColumnVals = AdjustedDataSheet.get_Range(startAddress.Replace("$", "") + ":" + endAddress.Replace("$", ""));
                string colValRangeAddress = rngListColumnVals.Address;
                string[] arrVals = colValRangeAddress.Split(new char[] { '$', ':' }, StringSplitOptions.RemoveEmptyEntries);
                int difference = Convert.ToInt32(arrVals[3]) - Convert.ToInt32(arrVals[1]);
                //difference = difference + 1; // to make to even number

                xrow = new object[difference];
                //int k = 1;
                int k = Convert.ToInt32(arrVals[1]);
                for (int j = 0; j < difference; j++)
                {
                    xrow[j] = k;
                    k++;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion
        #region //computed columns
        internal string SubtotalColumnFormula(string aggFcn, string subCol, string groupCol, string shtnm)
        {
            string formula = aggFcn + "(" + AdjustedData.Name + Utilities.ExcelHelpers.CreateValidFormulaName(subCol) + " " +
                                        Utilities.DataHelper.RowRangebyMatch(AdjustedData.Name, yr(), "[[#This Row]," + groupCol + "]", shtnm)
                                        + ")";
            return formula;
        }

        internal string AdjustedSubtotalRowFormula(string aggFcn, string subRow)
        {

            string formula = aggFcn + "(" + AdjustedData.Name + "[" + EnPIResources.yearColName + "]," + SummaryData.Name + "[#Headers]," + AdjustedData.Name + "[" + prefix() + subRow + "])";
            return formula;
        }

        internal string SubtotalRowFormula(string aggFcn, string subRow)
        {
            string colName = Utilities.ExcelHelpers.CreateValidFormulaName(subRow);
            string formula = aggFcn + "(" + AdjustedData.Name + "[" + EnPIResources.yearColName + "]," + SummaryData.Name + "[#Headers]," + AdjustedData.Name + colName + ")";
            return formula;
        }

        internal string AnnualSavingsRowFormula(int modelPosition, Excel.ListRow row, int modelIndex, string rowNamePrefix, string rowNameRaw)
        {
            string beforeModel = "IFERROR(IFERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1,1,1),0)+((INDEX(" + SummaryData.Name + ",MATCH(\"" + rowNameRaw + "\",[[ ]],0),COLUMN()-1)-INDEX(" + SummaryData.Name + ",MATCH(\"" + rowNamePrefix + "\", [[ ]],0),COLUMN()-1))-(INDEX(" + SummaryData.Name + ",MATCH(\"" + rowNameRaw + "\",[[ ]],0),)-INDEX(" + SummaryData.Name + ",MATCH(\"" + rowNamePrefix + "\",[[ ]],0),))),0)";
            string afterModel = "INDEX(" + SummaryData.Name + ", MATCH(\"" + rowNamePrefix + "\",[[ ]],0),) - INDEX(" + SummaryData.Name + ", MATCH(\"" + rowNameRaw + "\",[[ ]],0),)";

            string formula = "";

            switch (modelPosition)
            {
                //Model = Baseline
                case 1:
                    formula = afterModel;
                    //formula = "0";
                    break;
                //other
                case 2:
                    bool modelSwitch = false;
                    foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                    {
                        if (LC.Index.Equals(modelIndex))
                        {
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                            modelSwitch = true;
                        }
                        //after model
                        else if (modelSwitch)
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + afterModel;
                        //before model
                        else
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                    }
                    break;

                //Model = Last reporting year
                case 3:
                    formula = beforeModel;
                    break;
            }

            return formula;
        }

        internal string EstimatedCostSavingsRowFormula(int modelPosition, Excel.ListRow row, int modelIndex, string rowNamePrefix, string rowNameRaw)
        {
            string beforeModel = "0";
            string afterModel = "1";

            string formula = "";

            switch (modelPosition)
            {
                //Model = Baseline
                case 1:
                    formula = afterModel;
                    break;
                //other
                case 2:
                    bool modelSwitch = false;
                    foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                    {
                        if (LC.Index.Equals(modelIndex))
                        {
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                            modelSwitch = true;
                        }
                        //after model
                        else if (modelSwitch)
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + afterModel;
                        //before model
                        else
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                    }
                    break;

                //Model = Last reporting year
                case 3:
                    formula = beforeModel;
                    break;
            }

            return formula;
        }

        internal void AdjustmentArraySetup(object[] years)
        {
            AdjustmentMethod = new string[SummaryData.ListColumns.Count, 2];
            strAdjustmentMethodColName = (EnPIResources.AdjustmentMethodSEPColName);
            string rangeCheck = "";
            int count = 0;
            int count2 = 0;
            foreach (string yr in years)
            {
                count++;
                foreach (Excel.ListRow LR in SummaryData.ListRows)
                {
                    if (((Excel.Range)LR.Range[1, count]).Value2 != null)
                        if (((Excel.Range)LR.Range[1, count]).Value2.ToString().Equals(strAdjustmentMethodColName))
                        {
                            foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                            {
                                count2++;
                                Excel.Range range = ((Excel.Range)LR.Range[1, count2]);
                                if (range.Value2 != null)
                                {
                                    rangeCheck = range.Value2.ToString();
                                    AdjustmentMethod[count2 - 1, 1] = rangeCheck;
                                    AdjustmentMethod[count2 - 1, 0] = ((Excel.Range)LC.Range[1, count]).Value2.ToString();
                                }
                            }
                        }
                }
            }
        }

        internal string VaryingRowFormula(int modelIndex, Excel.ListRow row, string before, string after)
        {
            string beforeModel = "=" + before;
            string afterModel = "=" + after;
            string formula = "";

            if (AdjustmentMethod[1, 1].Equals(Globals.ThisAddIn.rsc.GetString("adjustmentModel")))
            {
                formula += afterModel;
            }
            else if (AdjustmentMethod[(AdjustmentMethod.Length / 2) - 1, 1].Equals(Globals.ThisAddIn.rsc.GetString("adjustmentModel")))
            {
                formula += beforeModel;
            }
            else
            {
                bool modelSwitch = false;
                foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                {
                    if (LC.Index.Equals(modelIndex))
                    {
                        ((Excel.Range)row.Range[1, LC.Index]).Value2 = beforeModel;
                        modelSwitch = true;
                    }
                    //after model
                    else if (modelSwitch)
                        ((Excel.Range)row.Range[1, LC.Index]).Value2 = afterModel;
                    //before model
                    else
                        ((Excel.Range)row.Range[1, LC.Index]).Value2 = beforeModel;
                }
            }

            return formula;
        }


        internal string SEnPI(int modelPosition, Excel.ListRow row, int modelIndex)
        {

            string unadjustedTotalColName = (Globals.ThisAddIn.rsc.GetString("unadjustedSEPTotalColName"));
            string totalAdjValuesColName = (Globals.ThisAddIn.rsc.GetString("totalAdjValuesSEPColName"));
            string beforeModel = "IFERROR((1" + /*INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),"
                            + modelIndex +
                        ")*/"/INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + unadjustedTotalColName +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),))*((INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + totalAdjValuesColName +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),)/1" +/*INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),"
                            + modelIndex +
                        ")" + ")*/")),1)";
            string afterModel = "IFERROR((INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + unadjustedTotalColName +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),)/1)" + /*INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\"" 
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + 
                        "\"," 
                        + SummaryData.Name + 
                        "[[ ]],0),"
                        + modelIndex +
                        "))*/"*" + /*((INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + 
                        "\"," 
                        + SummaryData.Name +
                        "[[ ]],0),"
                        + modelIndex +
                        ")*/"(1/INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + totalAdjValuesColName +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),)" + "),1)"; ;

            string SEnPI = "";
            switch (modelPosition)
            {
                //Model = Baseline
                case 1:
                    SEnPI = afterModel;
                    break;
                //other
                case 2:
                    bool modelSwitch = false;
                    foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                    {
                        if (LC.Index.Equals(modelIndex))
                        {
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                            modelSwitch = true;
                        }
                        //after model
                        else if (modelSwitch)
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + afterModel + " * OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-" + (LC.Index - 2).ToString() + ",1,1)";
                        //before model
                        else
                            ((Excel.Range)row.Range[1, LC.Index]).Value2 = "=" + beforeModel;
                    }
                    break;

                //Model = Last reporting year
                case 3:
                    SEnPI = beforeModel;
                    break;
            }

            return SEnPI;
        }

        internal string prefix()
        {
            string prefix = Globals.ThisAddIn.rsc.GetString("prefixAdjusted") ?? "Adj.";
            //string strPrefix = ((isSEnPI == false) ? EnPIResources.prefixAdjusted : EnPIResources.prefixSEPAdjusted);
            return prefix;
        }
        internal string yr()
        {
            return Utilities.ExcelHelpers.GetListColumn(SourceObject, EnPIResources.yearColName).Name;
        }
        internal string t()
        {
            return Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName"));
        }
        internal string that()
        {
            return Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName"));
        }
        internal string b()
        {
            return "OFFSET(" + SummaryData.Name + "[[#Headers]," + t() + "],1,0)";
        }
        internal string bhat()
        {
            return "OFFSET(" + SummaryData.Name + "[[#Headers]," + that() + "],1,0)";
        }
        internal string e()
        {
            return SummaryData.Name + "[[#This Row]," + t() + "]";
        }
        internal string ehat()
        {
            return SummaryData.Name + "[[#This Row]," + that() + "]";
        }
        internal string p()
        {
            return "OFFSET(" + SummaryData.Name + "[[#This Row]," + t() + "],-1,0)";
        }
        internal string phat()
        {
            return "OFFSET(" + SummaryData.Name + "[[#This Row]," + that() + "],-1,0)";
        }

        #endregion
        private void AddNewRowToSummaryData(Excel.ListObject SummaryData, string rowName, string rowValue, string stylename, string format)
        {
            Excel.ListRow newRow = SummaryData.ListRows.Add();
            newRow.Range.Value2 = "=" + rowValue;
            ((Excel.Range)newRow.Range[1, 1]).Value2 = rowName;
            newRow.Range.Style = stylename;
            ((Excel.Range)newRow.Range).Cells.Interior.Color = 0xBCE4D8;
            ((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
            ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Color = 0xFFFFFF;
            ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Bold = true;
            newRow.Range.NumberFormat = format;
            ((Excel.Range)newRow.Range[1, 1]).NumberFormat = "General";
        }

        private void AddNewRowToSEPFacilityIdentificationData(Excel.ListObject SEPFacilityIndentificationData, string rowName)
        {

            Excel.ListRow newRow = SEPFacilityIndentificationData.ListRows.Add();
            //newRow.Range.Value2 = string.Empty;
            ((Excel.Range)newRow.Range[1, 1]).Value2 = rowName;
            newRow.Range.Style = "Comma";
            ((Excel.Range)newRow.Range).Cells.Interior.Color = 0xFFFFFF;
            ((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0xBCE4D8;
            //((Excel.Range)newRow.Range).Cells.Interior.Color = 0xBCE4D8;
            ((Excel.Range)newRow.Range).Cells.Borders.Color = 0x000000;
            //((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
            //((Excel.Range)newRow.Range[1, 1]).Cells.Font.Color = 0xFFFFFF;
            ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Bold = true;
            newRow.Range.NumberFormat = "General";
           //((Excel.Range)newRow.Range[1, 1]).NumberFormat = "General";
        }

        public void modelInformation()
        {


            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Range header = BottomCell().get_Offset(2, 0);


            string start = header.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);
            ((Excel.Range)header[1, 1]).Value2 = " ";
            thisSheet.HPageBreaks.Add(((Excel.Range)header[1, 1]));
            ModelData = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, header, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            
            ModelData.Name = "Model" + ModelData.Name;
            ((Excel.Range)ModelData.Range[1, 1]).Value2 = " ";
            ModelData.TableStyle = "TableStyleMedium4";

            foreach (Utilities.Constants.ModelOutputColumns col in System.Enum.GetValues(typeof(Utilities.Constants.ModelOutputColumns)))
            {
                string label = Globals.ThisAddIn.rsc.GetString("label" + col.ToString()) ?? col.ToString();
                if (label != "Coefficients" && label != "Variable Std. Error" && label != "RMSError" && label != "Residual" && label != "AIC")
                {
                    header.Value2 = label;
                    if (label.Equals("Model Number"))
                        header.Value2 = "Energy Use";
                    if (label.Equals("Variables"))
                        header.Value2 = "Relevant Variables";
                    if (label.Equals("Adjusted R2"))
                        header.EntireColumn.Hidden = true;
                

                    header = header.get_Offset(0, 1);
                }
            }

            foreach (Excel.Worksheet WS in WB.Worksheets)
            {
                string sname = WS.Name;
                int IterationCount = 0;
                try
                {
                    IterationCount = Convert.ToInt32(WS.Name.Substring(0, 2).Trim()); //get iteration count from sheet name
                }
                catch
                {
                    //Nothing to catch, exception expected for sheets without iteration count as part of the name.
                }

                //Populate the model inforation only when the sheet iteration count matches
                if (IterationCount.Equals(Globals.ThisAddIn.groupSheetCollection.regressionIteration))
                    populateModelData(WS);

            }
            ModelData.Range.WrapText = true;
        }
        public void populateModelData(Excel.Worksheet WS)
        {
            object[,] row1;
            int compareModel = 0;

            foreach (Excel.ListObject ListObj in WS.ListObjects)
            {

                foreach (Excel.ListColumn colm in ListObj.ListColumns)
                {
                    string cname = colm.Name;
                    if (colm.Name == "Model Number")
                    {
                        try
                        {

                            int bestModelNumber = Convert.ToInt32(((Excel.Range)WS.Cells[2, 1]).Value2.ToString().Substring(((Excel.Range)WS.Cells[2, 1]).Value2.ToString().IndexOf("#") + 1));
                            string mdlValid = string.Empty;
                            foreach (Excel.ListRow row in ListObj.ListRows)
                            {
                                string[] c1 = new string[ListObj.ListColumns.Count];
                                if (((Excel.Range)row.Range.Cells[1]).Value2 != null)
                                {
                                    compareModel = Convert.ToInt32(((Excel.Range)row.Range.Cells[1]).Value2);
                                }

                                int l = ListObj.ListColumns.Count;
                                if (bestModelNumber == compareModel)
                                {
                                    //Modified for SEP Changes, as the model sheet now contains SEP Validation Check column increment the col number here.
                                    //row1 = new object[1, 13];
                                    row1 = new object[1, 14];
                                    Excel.Range target1;

                                    if (first == 0)
                                        //target1 = BottomCell().get_Offset(2, 0).get_Resize(1, 13);
                                        target1 = BottomCell().get_Offset(2, 0).get_Resize(1, 14);
                                    else
                                        //target1 = BottomCell().get_Offset(1, 0).get_Resize(1, 13);
                                         target1 = BottomCell().get_Offset(1, 0).get_Resize(1, 14);

                                    first += 1;
                                    
                                    string baddr = target1.Address.ToString();
                                    int y = 0;
                                    for (int x = 0; x < ListObj.ListColumns.Count; x++)
                                    {
                                        //if (x != 3 && x != 4 && x != 9 && x != 10 && x != 11)
                                        if (x != 4 && x != 5 && x != 10 && x != 11 && x != 12)
                                        {
                                            if (((Excel.Range)row.Range.Cells[x + 1]).Value2 != null)
                                            {
                                                if ((x != 0) && (x != 3))
                                                    row1[0, y] = ((Excel.Range)row.Range.Cells[x + 1]).Value2.ToString();
                                                else if (x == 3)
                                                    row1[0, y] = GetSEPValue(row1[0, 2].ToString());
                                                else
                                                    row1[0, y] = WS.Name;
                                            }

                                            else
                                            {
                                                row1[0, y] = "";
                                            }
                                            y += 1;
                                        }
                                    }

                                    if(!string.IsNullOrEmpty(row1[0,1].ToString()))
                                    {
                                         mdlValid = row1[0,1].ToString();
                                    }
                                    //target1.get_Resize(1, 13).Value2 = row1;
                                    target1.get_Resize(1, 14).Value2 = row1;
                                    //target1.Font.Color = 0x00AA00;
                                    target1.Font.Color = mdlValid.ToUpper() == "TRUE" ? 0x00AA00 : 0x0000AA;
                                    //GetSEPValue(row1[0, 2].ToString()) == "Pass" ? 0x00AA00 : 0x0000AA;
                                    //target1.Style = Globals.ThisAddIn.rsc.GetString("bestModelStyle");
                                    target1.Font.Bold = true;
                                    //target1.WrapText = true;

                                        target1.NumberFormat = "0.0000";
                                    

                                }
                            }
                        }
                        catch
                        {
                            //do nothing
                        }
                    }
                    //if (isSEnPI && this.thisSheet.Name.Contains("SEP"))
                    //{
                    //    if (colm.Name == "Adjusted R2") // TFS Ticket 77015
                    //    {
                    //        colm.Range.EntireColumn.Hidden = true;
                    //    }
                    //}

                }


            }
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

       
        #region //Formatting
        internal void AddConditionalFormatting(Excel.ListObject LO)
        {
            // for each independent variable
            foreach (Excel.ListColumn LC in LO.ListColumns)
            {
                if (DS.IndependentVariables.Contains(LC.Name))
                {
                    double low = 0;
                    double high = 0;
                    double[] values = Utilities.DataHelper.objectTOdblArray(LC.Range.Value2 as object[,]);

                    string fltr = "VariableName='" + LC.Name.Replace("'", "''") + "'";
                    DataRow row = DS.PredictorRange().Select(fltr).First();
                    // get the low value
                    low = Double.TryParse(row[1].ToString(), out low) ? Double.Parse(row[1].ToString()) : 0;
                    // get the high value
                    high = Double.TryParse(row[2].ToString(), out high) ? Double.Parse(row[2].ToString()) : 0;

                    // apply conditional formatting
                    if (low != 0 && high != 0)
                    {
                        Excel.Range rng = Utilities.ExcelHelpers.GetListColumn(LO, LC.Name).DataBodyRange;
                        Excel.FormatCondition fc = (Excel.FormatCondition)rng.FormatConditions.Add(
                                    Excel.XlFormatConditionType.xlCellValue, Excel.XlFormatConditionOperator.xlNotBetween, low, high
                                    , System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
                        fc.Font.ColorIndex = 3;
                    }
                }
            }
        }

        internal void FormatHeaderRow(Excel.ListObject LO)
        {
            Excel.Range header = LO.HeaderRowRange;

            header.Cells.WrapText = true;
            header.Cells.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            header.Cells.UseStandardHeight = true;
            header.Cells.EntireRow.AutoFit();
        }
        internal void FormatSummaryData()
        {
            FormatHeaderRow(SummaryData);

            foreach (Excel.Range col in thisSheet.UsedRange.Columns)
            {
                col.AutoFit();
            }
            SummaryData.Range.WrapText = true;
            
        }
        #endregion
    }
}
