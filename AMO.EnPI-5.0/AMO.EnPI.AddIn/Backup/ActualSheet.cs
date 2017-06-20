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
    public class ActualSheet
    {
        public Utilities.EnPIDataSet DS;
        public Excel.Worksheet thisSheet;
        public Excel.Worksheet SourceSheet;
        public Excel.ListObject SourceObject;
        public object[,] SourceData;
        public string[,] AdjustmentMethod;
        public Excel.Worksheet DetailDataSheet;
        public Excel.ListObject DetailData;
        public Excel.Worksheet AdjustedDataSheet;
        public Excel.ListObject AdjustedData;
        public Excel.ListObject SummaryData;
        public Excel.ListObject ModelData;
        public Excel.ListObject WarningData;
        public Excel.ChartObject ChartObj;
        public Excel.ChartObject ChartObj2;
        public ArrayList Warnings;

        
        public ActualSheet(Utilities.EnPIDataSet DSIn)
        {
            DS = DSIn;
            string nm = Globals.ThisAddIn.rsc.GetString("enpiActualTitle");
            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            SourceSheet = Utilities.ExcelHelpers.GetWorksheet(WB, DS.WorksheetName);

            Excel.Worksheet aSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add
                (System.Type.Missing, WB.Sheets.get_Item(WB.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet);
            aSheet.CustomProperties.Add("SheetGUID", System.Guid.NewGuid().ToString());

            aSheet.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(WB, nm, Globals.ThisAddIn.groupSheetCollection.regressionIteration);
            aSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            aSheet.Tab.Color = 0x008000;
            Utilities.ExcelHelpers.addWorksheetCustomProperty(aSheet, Utilities.Constants.WS_ISENPI, "True");
            ExcelHelpers.addWorksheetCustomProperty(aSheet, Constants.WS_ROLLUP, "True");
            thisSheet = aSheet;
            Warnings = new ArrayList();

        }

        private Excel.Range BottomCell()
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(thisSheet, 0).ToString();

            return (Excel.Range)thisSheet.get_Range(addr, System.Type.Missing);
        }

        public void Populate()
        {
            Excel.Range rangeTitle = (Excel.Range)thisSheet.get_Range("A1", "H1");
            ((Excel.Range)rangeTitle[1, 1]).Value2 = EnPIResources.enpiSheetTitle;
            ((Excel.Range)rangeTitle[1, 1]).Font.Color = 0x008000;
            ((Excel.Range)rangeTitle[1, 1]).Font.Bold = true;
            ((Excel.Range)rangeTitle[1, 1]).Font.Size = 15;
            rangeTitle.Merge();

            //This is added to keep the sheets consistent
            Excel.Range rangeBody = (Excel.Range)thisSheet.get_Range("A2", "H2");
            ((Excel.Range)rangeBody[1, 1]).Value2 = string.Empty;
            rangeBody.Merge();
            rangeBody.WrapText = true;
            rangeBody.EntireRow.Hidden = true;
            
            SourceData = Utilities.DataHelper.dataTableArrayObject(DS.SourceData); 
            SourceObject = Utilities.ExcelHelpers.GetListObject(SourceSheet, DS.ListObjectName);

            AddTable();
            
            AddSubtotalColumns();

            ChartObj = newEnPIChart();
            ChartObj2 = newEnPIChart2();

            FormatSummaryData();

            writeCharts();

            GroupSheet GS = new GroupSheet(thisSheet, true, false,thisSheet.Name);
            Globals.ThisAddIn.groupSheetCollection.Add(GS);
            DetailDataSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            thisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
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

            // write the generic warnings
            foreach (System.Data.DataRow wr in DS.VariableWarnings.Select("VariableName is null"))
            {
                string w = wr[2].ToString();
                if (!Warnings.Contains(w))
                {
                    errorNotPresent = false;
                    Warnings.Add(w);
                }
            }

            // this checks only the variables in the best model for each energy source
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                mdl = es.BestModel();
                foreach (string vr in mdl.VariableNames)
                {
                    string expr = "VariableName = '" + vr.Replace("'","''") + "'";
                    foreach (System.Data.DataRow dr in DS.VariableWarnings.Select(expr))
                    {
                        if (DS.Years.Contains(dr[1].ToString()))
                        {
                            string w = "The average value in the predictor column " + vr + " for the year " + dr[1].ToString() + " is outside of the allowable range.";
                            if (!Warnings.Contains(w))
                            {
                                errorNotPresent = false;
                                Warnings.Add(w);
                            }
                        }
                    }
                }
            }

            if (DS.OutlierCount() != 0)
            {
                errorNotPresent = false;
                Warnings.Add("The predictor data may contain outliers. Look in the detail data table for highlighted values.");
            }

            foreach (string st in Warnings)
            {
                Excel.Range rg = BottomCell().get_Offset(1, 0).get_Resize(1,12);
                rg.Merge();
                rg.Value2 = st;
                rg.Font.Color = 0x0000AA;
                //rg.Style = "Bad";
            }

            sumRange.EntireRow.Hidden = errorNotPresent;
        }
        
        internal void AddTable()
        {
            Excel.Range sumRange = BottomCell().get_Offset(2, 0);

            sumRange.Value2 = yr();

            object[] years = Utilities.ExcelHelpers.getYears(DS.SourceData);
            int ycols = years.Length;

            sumRange = sumRange.get_Resize(1, ycols + 1);
            for (int i = 0; i < ycols; i++)
            {
                if (false)//i.Equals(0))//
                    sumRange.get_Offset(0, i + 1).get_Resize(1, 1).Value2 = years[i] + " (Baseline)";
                else
                    sumRange.get_Offset(0, i + 1).get_Resize(1, 1).Value2 = years[i];
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
                ((Excel.Range)newRow.Range[1, 1]).Value2 = name;
                newRow.Range.Style = stylename;
                ((Excel.Range)newRow.Range).Cells.Interior.Color = 0xBCE4D8;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Color = 0xFFFFFF;
                ((Excel.Range)newRow.Range[1, 1]).Cells.Font.Bold = true;
                newRow.Range.NumberFormat = "###,##0";
         
            }
            //Added By Suman TFS Ticket: 68479
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                string name = es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                if (!name.Contains("TOTAL"))// Need to find out a way to eliminate total column here 
                {
                    AddNewRowToSummaryData(SummaryData, name + " Annual Savings", SubtotalRowFormula("SUMIF", "Energy Savings: " + name), stylename, "###,##0");
                    if (Globals.ThisAddIn.fromEnergyCost)
                    {
                        //TFS Ticket 68851 :Modified By Suman
                        AddNewRowToSummaryData(SummaryData, name + " Estimated Cost Savings ($)", SubtotalRowFormula("SUMIF", "Cost Savings ($): " + name), stylename, "_($* #,##0.00_);_($* (#,##0.00);_($* " + " -" + "??_);_(@_)");
                    }
                }
            }

            if(DS.ProductionVariables.Count > 0)
            {
                Excel.ListRow prodOutput = SummaryData.ListRows.Add();
                string range = "";
                foreach (string prod in DS.ProductionVariables)
                {
                    if (DS.ProductionVariables.IndexOf(prod).Equals(DS.ProductionVariables.Count - 1))
                        range += SubtotalRowFormula("SUMIF", prod);
                    else
                        range += SubtotalRowFormula("SUMIF", prod) + " + ";
                }
                string name5 = "Total Production Output";
                prodOutput.Range.Value2 = "=" + range ;
                ((Excel.Range)prodOutput.Range[1, 1]).Value2 = name5;
                prodOutput.Range.NumberFormat = "###,##0";
            }
            

            if (DS.ProductionVariables.Count > 0)
            {
                addUnadjustedEnergyIntensity();
            }

            if (DS.BuildingVariables.Count > 0)
            {
                addUnadjustedBuildingEnergyInten();
            }

            AdjustmentArraySetup(Utilities.ExcelHelpers.getYears(DS.SourceData));
            
            //calculate Cumulative Improvement
            Excel.ListRow cumulativeImprovRow = SummaryData.ListRows.Add(System.Type.Missing);
            string ciName = EnPIResources.totalImprovementColName;
            if (DS.ProductionVariables.Count > 0)
                cumulativeImprovRow.Range.Value2 = "=(INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedEnergyIntensColName") + "\"," + SummaryData.Name + "[[ ]],0),2)-INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedEnergyIntensColName") + "\"," + SummaryData.Name + "[[ ]],0),COLUMN()))/(INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedEnergyIntensColName") + "\"," + SummaryData.Name + "[[ ]],0),2))";
            else if (DS.BuildingVariables.Count > 0)
                cumulativeImprovRow.Range.Value2 = "=(INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedBuildingColName") + "\"," + SummaryData.Name + "[[ ]],0),2)-INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedBuildingColName") + "\"," + SummaryData.Name + "[[ ]],0),COLUMN()))/(INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedBuildingColName") + "\"," + SummaryData.Name + "[[ ]],0),2))";
            else
                cumulativeImprovRow.Range.EntireRow.Hidden = true;
            cumulativeImprovRow.Range.Style = "Percent";
            ((Excel.Range)cumulativeImprovRow.Range[1, 1]).Value2 = ciName;
            ((Excel.Range)cumulativeImprovRow.Range[1, 2]).Value2 = 0;
            cumulativeImprovRow.Range.NumberFormat = "0.00%";

            //calculate Annual Improvement
            Excel.ListRow annualImprovRow = SummaryData.ListRows.Add(System.Type.Missing);
            string aiName = EnPIResources.annualImprovementColName;
            if (DS.ProductionVariables.Count > 0 || DS.BuildingVariables.Count > 0)
                annualImprovRow.Range.Value2 = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1)-OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,-1,1,1)";
            else
                annualImprovRow.Range.EntireRow.Hidden = true;
            annualImprovRow.Range.Style = "Percent";
            ((Excel.Range)annualImprovRow.Range[1, 1]).Value2 = aiName;
            ((Excel.Range)annualImprovRow.Range[1, 2]).Value2 = 0;
            annualImprovRow.Range.NumberFormat = "0.00%";

            //Calculate Annual Savings
            Excel.ListRow annualSavingsRow = SummaryData.ListRows.Add(System.Type.Missing);
            annualSavingsRow.Range.Value2 = "=(INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + "\"," + SummaryData.Name + "[[ ]],0),2)-INDEX(" + SummaryData.Name + ",MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + "\"," + SummaryData.Name + "[[ ]],0),COLUMN()))";
            annualSavingsRow.Range.Style = "Comma";
            ((Excel.Range)annualSavingsRow.Range[1, 1]).Value2 = "Total Savings Since Baseline Year (MMBtu/Year)";
            ((Excel.Range)annualSavingsRow.Range).Cells.Interior.Color = 0xFFFFFF;
            ((Excel.Range)annualSavingsRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
            ((Excel.Range)annualSavingsRow.Range[1, 2]).Value2 = 0;
            annualSavingsRow.Range.NumberFormat = "###,##0";

            //Calculate Cumulative savings
            Excel.ListRow cumulativeSavingsRow = SummaryData.ListRows.Add(System.Type.Missing);
            cumulativeSavingsRow.Range.Value2 = "=OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1)-OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,-1,1,1)";
            cumulativeSavingsRow.Range.Style = "Comma";
            ((Excel.Range)cumulativeSavingsRow.Range[1, 1]).Value2 = "New Energy Savings for Current Year (MMBtu/year)";
            ((Excel.Range)cumulativeSavingsRow.Range).Cells.Interior.Color = 0xFFFFFF;
            ((Excel.Range)cumulativeSavingsRow.Range[1, 1]).Cells.Interior.Color = 0x28624F;
            ((Excel.Range)cumulativeSavingsRow.Range[1, 2]).Value2 = 0;
            cumulativeSavingsRow.Range.NumberFormat = "###,##0";

            //Added By Suman:TFS Ticket 68479
            //Estimated Annual cost savings
            if (Globals.ThisAddIn.fromEnergyCost)
            {
                string estimatedAnnualCostSavingsFormula =string.Empty;
                foreach(Utilities.EnergySource es in DS.EnergySources)
                {
                    if(!es.Name.Contains("TOTAL")) // Need to find out a way to eliminate total column here 
                    estimatedAnnualCostSavingsFormula += "INDEX("+ SummaryData.Name +",MATCH(\""+es.Name+" Estimated Cost Savings ($)"+ "\"," + "[[ ]],0),COLUMN())+";
                }
                estimatedAnnualCostSavingsFormula=estimatedAnnualCostSavingsFormula.Remove(estimatedAnnualCostSavingsFormula.Length - 1, 1);
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
                        co2EmissionFormula += "INDEX(" + SummaryData.Name + ",MATCH(\"" + es.Name + " Annual Savings" + "\"," + "[[ ]],0),COLUMN())*" + emissionFactor + "/1000+"; //TFS Ticket: 68853
                    }
                }
                co2EmissionFormula = co2EmissionFormula.Remove(co2EmissionFormula.Length - 1, 1);
                AddNewRowToSummaryData(SummaryData, "Avoided CO2 Emissions (Metric Ton/year)", co2EmissionFormula, stylename, "###,##0"); //TFS Ticket: 70385

            }

        }
        //Added By Suman TFS Ticket: 68479
        private void AddNewRowToSummaryData(Excel.ListObject SummaryData, string rowName, string rowValue, string stylename,string format)
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

        internal void addSumColumn(params string[] prepend)
        {
            string colName = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");
            string formula = "=";
            string format = "General";
            string prefix = "";
            if (prepend.Length > 0)
            {
                prefix = prepend[0];
                colName = Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName");
            }

            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                if (es.Name != Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName"))
                {
                    formula += Utilities.ExcelHelpers.CreateValidFormulaName(prefix + es.Name) + "+";

                    if (Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name) != null)
                        format = Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name).DataBodyRange.NumberFormat.ToString();
                }
            }
            
            if (formula.LastIndexOf("+") > 0)
                formula = formula.Substring(0, formula.LastIndexOf("+"));

            Excel.ListColumn newcol = Utilities.ExcelHelpers.AddListColumn(AdjustedData, colName);
            newcol.Range.get_Offset(1, 0).get_Resize(newcol.Range.Rows.Count - 1, 1).Value2 = formula;
            newcol.Range.get_Offset(1, 0).NumberFormat = format ?? "General";

        }

        internal void addAdjustedBuildingEnergyInten()
        {
            if (DS.BuildingVariables != null)
            {
                string formula = "";
                for (int j = 0; j < DS.BuildingVariables.Count(); j++)
                {
                    formula += SubtotalColumnFormula("AVERAGE", DS.BuildingVariables[j].ToString(), yr(), AdjustedDataSheet.Name) + "+";
                }
                if (formula != "")
                {
                    Excel.ListColumn newCol = SummaryData.ListColumns.Add(2);
                    newCol.Name = Globals.ThisAddIn.rsc.GetString("adjustedBuildingName"); ;
                    newCol.DataBodyRange.Value2 = "=" + that() + "/" + formula.Substring(0, formula.Length - 1);
                    newCol.DataBodyRange.Style = "Comma [0]";
                }
            }
        }

        internal void addUnadjustedBuildingEnergyInten()
        {
            Excel.ListRow buildEnIntenRow = SummaryData.ListRows.Add();
            
            string formula = "";
            for (int j = 0; j < DS.BuildingVariables.Count(); j++)
            {
                formula += SubtotalRowFormula("AVERAGEIF", DS.BuildingVariables[j].ToString()) + "+";
            }
            if (formula != "")
            {
                string name = Globals.ThisAddIn.rsc.GetString("unadjustedBuildingColName");
                buildEnIntenRow.Range.Value2 = "=" + "OFFSET(" + SummaryData.Name + "[#Headers], MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + "\"," + SummaryData.Name + "[[#All],[ ]],0)-1,0)" + "/"
                    + formula.Substring(0, formula.Length - 1);
                buildEnIntenRow.Range.Style = "Comma [0]";
                ((Excel.Range)buildEnIntenRow.Range[1, 1]).Value2 = name;
            }

            buildEnIntenRow.Range.NumberFormat = "###,##0.000";
        }

        internal void addAdjustedEnergyIntensity()
        {
            if (DS.ProductionVariables != null)
            {
                string formula = "";
                for (int j = 0; j < DS.ProductionVariables.Count(); j++)
                {
                    formula += SubtotalColumnFormula("SUM", DS.ProductionVariables[j].ToString(), yr(), AdjustedDataSheet.Name) + "+";
                }
                if (formula != "")
                {
                    Excel.ListColumn newCol = SummaryData.ListColumns.Add(2);
                    newCol.Name = Globals.ThisAddIn.rsc.GetString("adjustedEnergyIntensName"); ;
                    newCol.DataBodyRange.Value2 = "=" + that() + "/" + formula.Substring(0, formula.Length - 1);
                    newCol.DataBodyRange.Style = "Comma [0]";
                }

            }
        }

        internal void addUnadjustedEnergyIntensity()
        {
            Excel.ListRow prodEnIntenRow = SummaryData.ListRows.Add();

            string formula = "";
            for (int j = 0; j < DS.ProductionVariables.Count(); j++)
            {
                formula += SubtotalRowFormula("SUMIF", DS.ProductionVariables[j].ToString()) + "+";
            }
            if (formula != "")
            {
                string name = Globals.ThisAddIn.rsc.GetString("unadjustedEnergyIntensColName");
                prodEnIntenRow.Range.Value2 = "=" + "OFFSET(" + SummaryData.Name + "[#Headers], MATCH(\"" + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + "\"," + SummaryData.Name + "[[#All],[ ]],0)-1,0)" + "/"
                    + "OFFSET(" + SummaryData.Name + "[#Headers], MATCH(\"Total Production Output\"," + SummaryData.Name + "[[#All],[ ]],0)-1,0)";
                prodEnIntenRow.Range.Style = "Comma [0]";
                ((Excel.Range)prodEnIntenRow.Range[1, 1]).Value2 = name;
            }
            prodEnIntenRow.Range.NumberFormat = "###,##0.000";
        }

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

            string formula = aggFcn + "(" + AdjustedData.Name + "[" +  EnPIResources.yearColName + "]," + SummaryData.Name + "[#Headers]," + AdjustedData.Name + "[" + prefix() + subRow + "])";
            return formula;
        }

        internal string SubtotalRowFormula(string aggFcn, string subRow)
        {
            string colName = Utilities.ExcelHelpers.CreateValidFormulaName(subRow);
            string formula = aggFcn + "(" + AdjustedData.Name + "["+ EnPIResources.yearColName +"]," + SummaryData.Name + "[#Headers]," + AdjustedData.Name + colName + ")";
            return formula;
        }

        internal void AdjustmentArraySetup(object[] years)
        {
            AdjustmentMethod = new string[SummaryData.ListColumns.Count, 2];

            string rangeCheck = "";
            int count = 0;
            int count2 = 0;
            foreach (string yr in years)
            {
                count++;
                foreach (Excel.ListRow LR in SummaryData.ListRows)
                {
                    if (((Excel.Range)LR.Range[1, count]).Value2 != null)
                        if (((Excel.Range)LR.Range[1, count]).Value2.ToString().Equals("Adjustment Method"))
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
            else if (AdjustmentMethod[(AdjustmentMethod.Length / 2 )-1, 1].Equals(Globals.ThisAddIn.rsc.GetString("adjustmentModel")))
            {
                formula += beforeModel;
            }
            else
            {
                bool modelSwitch = false;
                foreach (Excel.ListColumn LC in SummaryData.ListColumns)
                {
                    if(LC.Index.Equals(modelIndex))
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
            string beforeModel = "IFERROR((INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),"
                            + modelIndex +
                        ")/INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),))*((INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),)/INDEX("
                        + SummaryData.Name +
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") +
                        "\","
                        + SummaryData.Name +
                        "[[ ]],0),"
                            + modelIndex +
                        ")" + ")),1)";;
            string afterModel = "IFERROR((INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\"" 
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + 
                        "\"," 
                        + SummaryData.Name + 
                        "[[ ]],0),)/INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\"" 
                        + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + 
                        "\"," 
                        + SummaryData.Name + 
                        "[[ ]],0),"
                        + modelIndex +
                        "))*((INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\""
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + 
                        "\"," 
                        + SummaryData.Name +
                        "[[ ]],0),"
                        + modelIndex +
                        ")/INDEX(" 
                        + SummaryData.Name + 
                        ",MATCH(\"" 
                        + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName") + 
                        "\"," 
                        + SummaryData.Name + 
                        "[[ ]],0),)" + ")),1)";;

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
            return prefix;
        }
        internal string yr()
        {
            return Utilities.ExcelHelpers.GetListColumn(SourceObject,EnPIResources.yearColName).Name;
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

        #region //Charting

        internal Excel.ChartObject newEnPIChart()
        {
            Excel.Range start = BottomCell().get_Offset(2,0);

            double topleft;
            if (!double.TryParse(start.Top.ToString(), out topleft))
                topleft = 0;

            start.EntireRow.RowHeight = Utilities.Constants.CHART_HEIGHT * 1.1;

            Excel.ChartObject CO = ((Excel.ChartObjects)thisSheet.ChartObjects(System.Type.Missing))
                .Add(10, topleft, Utilities.Constants.CHART_WIDTH, Utilities.Constants.CHART_HEIGHT);
            CO.Placement = Excel.XlPlacement.xlMove;

            return CO;
        }

        internal Excel.ChartObject newEnPIChart2()
        {
            Excel.Range start = BottomCell();

            double topleft;
            if (!double.TryParse(start.Top.ToString(), out topleft))
                topleft = 0;

            Excel.ChartObject CO = ((Excel.ChartObjects)thisSheet.ChartObjects(System.Type.Missing))
                .Add(275, topleft, Utilities.Constants.CHART_WIDTH, Utilities.Constants.CHART_HEIGHT);
            CO.Placement = Excel.XlPlacement.xlMove;

            return CO;
        }

        internal void writeCharts()
        {
            //Modified By Suman TFS Ticket:68735
            
            ChartObj.Chart.ChartType = Excel.XlChartType.xlLineMarkers;
            ChartObj.Chart.ChartStyle = 5;
            ChartObj.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
            ChartObj2.Chart.ChartStyle = 37;
            
            ChartObj2.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;

            Excel.ListObject lo = AdjustedData;
            Excel.ListObject lo2 = SummaryData;

            string tot = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");

            // find baseYear source sum column
            int i = lo.ListColumns.Count;

            foreach (Excel.ListColumn col in lo.ListColumns)
            {
                if (col.Name == tot) i = col.Index;
            }
            lo.ListColumns[i].TotalsCalculation = Excel.XlTotalsCalculation.xlTotalsCalculationSum;

            Excel.Axis ChartObjYaxis = (Excel.Axis)ChartObj.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            ChartObjYaxis.HasTitle = true;
            ChartObjYaxis.AxisTitle.Text = "Total Consumption (MMBtu)";

            Excel.Axis ChartObjXaxis = (Excel.Axis)ChartObj.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            ChartObjXaxis.HasTitle = true;
            ChartObjXaxis.AxisTitle.Text = "Year";

             //plot adjusted values
            Excel.Series xSeries = ((Excel.SeriesCollection)ChartObj.Chart.SeriesCollection(System.Type.Missing)).Add(lo.ListColumns[i].Range
                                , Excel.XlRowCol.xlColumns, true, System.Type.Missing, System.Type.Missing);
            xSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DodgerBlue);
            xSeries.MarkerBackgroundColorIndex = (Microsoft.Office.Interop.Excel.XlColorIndex)25;
            xSeries.Format.Line.Transparency = 1.0f;

            int enInten = -1;
            int buildEnInten = -1;
            int anualImprovement = -1;
            int totalImprovment = -1;
            int year = 0;

            foreach (Excel.ListRow row in lo2.ListRows)
            {
                if (((Excel.Range)row.Range[0, 1]).Value2 != null)
                {
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.unadjustedEnergyIntensColName)) enInten = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.annualImprovementColName)) anualImprovement = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.totalImprovementColName)) totalImprovment = row.Index - 1;
                    if (((Excel.Range)row.Range[0, 1]).Value2.ToString().Equals(EnPIResources.unadjustedBuildingColName)) buildEnInten = row.Index - 1;
                }
            }

            if(!enInten.Equals(-1))
            ((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Add(lo2.ListRows[enInten].Range
                                , Excel.XlRowCol.xlRows, true, System.Type.Missing, System.Type.Missing);
            if (!buildEnInten.Equals(-1))
                ((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Add(lo2.ListRows[buildEnInten].Range
                                    , Excel.XlRowCol.xlRows, true, System.Type.Missing, System.Type.Missing);
            try
            {
                if (!anualImprovement.Equals(-1))
                    ((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Add(lo2.ListRows[anualImprovement].Range
                                        , Excel.XlRowCol.xlRows, true, System.Type.Missing, System.Type.Missing).AxisGroup = Excel.XlAxisGroup.xlSecondary;
            }
            catch (Exception e)
            { }
            try
            {
                if (!totalImprovment.Equals(-1))
                    ((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Add(lo2.ListRows[totalImprovment].Range
                                        , Excel.XlRowCol.xlRows, true, System.Type.Missing, System.Type.Missing).AxisGroup = Excel.XlAxisGroup.xlSecondary;
            }
            catch (Exception e)
            { }

            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count - 1)).ChartType = Excel.XlChartType.xlLineMarkers;
            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count - 1)).Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count - 1)).MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkGreen);
            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count)).ChartType = Excel.XlChartType.xlLineMarkers;
            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count)).Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);
            ((Excel.Series)ChartObj2.Chart.SeriesCollection(((Excel.SeriesCollection)ChartObj2.Chart.SeriesCollection(System.Type.Missing)).Count)).MarkerBackgroundColor = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.YellowGreen);

            Excel.Axis CharObj2Yaxis = (Excel.Axis)ChartObj2.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
            CharObj2Yaxis.HasTitle = true;
            CharObj2Yaxis.AxisTitle.Text = "Energy Intensity (MMBtu/unit)";

            Excel.Axis CharObj2Yaxis2 = (Excel.Axis)ChartObj2.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary);
            CharObj2Yaxis2.HasTitle = true;
            CharObj2Yaxis2.AxisTitle.Text = "Percent Improvement";

            Excel.Axis CharObj2Xaxis = (Excel.Axis)ChartObj2.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
            CharObj2Xaxis.HasTitle = true;
            CharObj2Xaxis.AxisTitle.Text = "Reporting Year";

            ChartObj.Chart.ChartTitle.Text = "";

        }

        internal void modelInformation()
        {
            Excel.Range sumRange = BottomCell().get_Offset(2, 0);

            ((Excel.Range)sumRange[1, 1]).Value2 = " ";

            //sumRange.Value2 = "Info";

            ModelData = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, sumRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            ModelData.Name = "Model" + ModelData.Name;
            ((Excel.Range)ModelData.Range[1, 1]).Value2 = " ";
            ModelData.TableStyle = "TableStyleMedium4";

            ModelData.ListColumns.Add(Type.Missing).Name = "Variables";
            ModelData.ListColumns.Add(Type.Missing).Name = "Model is Appropriate for SEP";
            ModelData.ListColumns.Add(Type.Missing).Name = "Model Validity";
            ModelData.ListColumns.Add(Type.Missing).Name = "Model P-Value";
            ModelData.ListColumns.Add(Type.Missing).Name = "Variable P-Value";
            ModelData.ListColumns.Add(Type.Missing).Name = "Adjusted R2";
            ModelData.ListColumns.Add(Type.Missing).Name = "Model Selected";


            //get all the variables from the currently selected model. 
            
            foreach(String esv in DS.EnergySourceVariables)
            {
                ModelData.ListRows.AddEx(Type.Missing, Type.Missing);
                ModelData.ListRows[DS.EnergySourceVariables.IndexOf(esv) + 1].Range.Value2 = esv;                
            }

        }

        #endregion

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

                    string fltr = "VariableName='" + LC.Name.Replace("'","''") + "'";
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

        }
        #endregion
    }
}