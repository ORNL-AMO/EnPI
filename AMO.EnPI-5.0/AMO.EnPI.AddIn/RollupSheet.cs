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

namespace AMO.EnPI.AddIn
{
    public class RollupSheet
    {
        public Excel.Worksheet WS;
        public string WorksheetName { get; set; }
        public string ListObjectName { get; set; }
        public string SQLStatement { get; set; }
        public bool SaveConnections { get; set; }

        public List<string> SourceTables { get; set; }
        public List<string> ReportYears { get; set; }
        public List<string> EnergySourceNamesMaster { get; set; }
        public ArrayList ReportColumns { get; set; }
        public Excel.ListObject RollupData;
        public int rawIndex = 0;
        private bool pivotRefresh = false;

        static string name_unadj = "TOTAL Primary Energy Consumed (MMBtu/year)";
        static string name_adj = "TOTAL MODELED Primary Energy Consumed (MMBtu/year)";
        static string name_senpi = "Cumulative SEnPI";
        static string name_kpi = "Total Improvement (%)";
        static string name_aimp = "Annual Improvement (%)";
        static string name_bladj = "Total Energy Savings since Baseline Year (MMBtu/year)";
        static string name_savings = "New Energy Savings for Current Year (MMBtu/year)";

        static string col_year = "Period";
        static string col_year_step = "Year Step";
        static string col_year_step2 = "Year Step 2";
        static string col_name = "Name";
        static string col_ba = "After Model Year";
        static string col_blyear = "Baseline Year";
        static string col_mdlyear = "Model Year";
        static string col_unadjbl = "Unadjusted Baseline";
        static string col_ratio = "Ratio of Unadjusted to Adjusted";
        static string col_blratio = "Baseline Year Ratio"; 
        static string col_mdlratio = "Model Year Ratio";
        static string col_adjbl = "Adjusted Baseline";
        static string col_unadjmdl = "Unadjusted Model";
        static string col_adjmdl = "Adjusted Model";
        static string col_unadj = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");
        static string col_adj = Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName");
        static string col_prod = Globals.ThisAddIn.rsc.GetString("productionColName");
        static string col_energyintens = "Energy Intensity";
        static string col_weight = "Baseline Weight";
        static string col_bladj = "Baseline Adj.";
        static string col_svgs = "Savings";
        static string col_svgs_step = "Savings Step";
        static string col_senpi = "Cumulative SEnPI";
        static string col_senpi_step = "SEnPI Step";
        static string col_ci = "Cumulative Improvment";
        static string col_ci_step = "Cumulative Improvment Step";
        static string col_ai = "Annual Improvment";
        static string col_ai_step = "Annual Improvment Step";
        static string col_kpi = "Improvement";
        static string col_aimp_step = "Annual Improvement Step";
        static string col_wimp = "Weighted Improvemt";
        static string col_aimp = "Annual Improvement";

        static ArrayList pivot_src_hdrs = new ArrayList { col_name, col_year, col_year_step, col_year_step2, col_blyear, col_unadjbl, col_adjbl, col_blratio, col_mdlyear, col_ba, col_unadjmdl, col_adjmdl, col_mdlratio, col_unadj, col_adj, col_ratio, col_bladj, col_svgs, col_svgs_step, col_senpi, col_senpi_step, col_ci, col_ci_step, col_ai, col_ai_step, col_kpi, col_aimp_step, col_aimp, col_weight, col_wimp, col_prod, col_energyintens };
       
        static string tab_name = "Rollup Data";
        static string raw_name = "Raw Data";
        // these are the columns that will be added as data fields to the final pivot report
        // note that column captions cannot duplicate an existing column name, so make sure they're unique)
        static string[] pivot_columns = new string[7] { col_unadj, col_adj, col_senpi, col_ai, col_ci, col_bladj, col_svgs };
        static string[] pivot_captions = new string[7] { name_unadj, name_adj, name_senpi, name_aimp, name_kpi, name_savings, name_bladj };
        static string[] pivot_formats = new string[7] { "#,##0", "#,##0", "#0.0", "0.0%", "0.0%", "#,##0", "#,##0" };
        static Excel.XlConsolidationFunction[] pivot_aggs = new Excel.XlConsolidationFunction[7] { Excel.XlConsolidationFunction.xlSum, Excel.XlConsolidationFunction.xlSum, Excel.XlConsolidationFunction.xlMax, Excel.XlConsolidationFunction.xlSum, Excel.XlConsolidationFunction.xlAverage, Excel.XlConsolidationFunction.xlSum, Excel.XlConsolidationFunction.xlMax };


        public DetailTableCollection RollupSources;

        public RollupSheet(Excel.Worksheet RS)
        {
            Excel.Workbook WB = (Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet aSheet = RS;
            if (RS == null)
            {
                try
                {
                    aSheet = (Excel.Worksheet)WB.Worksheets.Add(System.Type.Missing, WB.Sheets[WB.Sheets.Count], 1, Excel.XlSheetType.xlWorksheet);
                }
                catch
                {
                    aSheet = (Excel.Worksheet)WB.Worksheets.Add(System.Type.Missing, System.Type.Missing, 1, Excel.XlSheetType.xlWorksheet);
                }
            }
            aSheet.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(WB, tab_name, 0);
            
            WS = aSheet;
            Excel.Range titleRange = WS.Range["A1", "E1"];
            titleRange.Merge();
            //titleRange.Style = "Heading 1";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 28;
            titleRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone;
            titleRange.Font.Color = 0x008000;
            titleRange.Value2 = "Corporate Roll-up";

            SourceTables = new List<string>();
            ReportYears = new List<string>();
            EnergySourceNamesMaster = new List<string>();
            RollupSources = new DetailTableCollection();
        }

        public void Initialize(int numPlant)
        {
            int masterColCount = 0;
            int headerCount = 1;
            int sourcesCount = RollupSources.Count;
            Array[] masterArray = new Array[sourcesCount];
            string[] baselineYears = new string[sourcesCount];
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet as Excel.Worksheet;
            
            //pull all column names from all imported sheets for processing 
            for (int j = 0; j < sourcesCount; j++)
            {
                int colCount = RollupSources.Item(j).thisTable.ListColumns.Count;
                masterColCount += colCount;
                string[] array = new string[colCount];
                for (int i = 0; i < colCount; i++)
                {
                    //store baseline year
                    if (i.Equals(1))
                        baselineYears[j] = RollupSources.Item(j).thisTable.ListColumns[i + 1].Name;

                    bool emptyColumn = true;

                    //if column is empty exclude it from the tableHeader
                    for (int p = 1; p < RollupSources.Item(j).thisTable.ListRows.Count; p++)//Excel.ListRow lr in RollupSources.Item(j).thisTable.ListRows)
                    {
                        //skip the header row to only check data
                        if (p > 1)
                        {
                            if (((Excel.Range)RollupSources.Item(j).thisTable.ListRows[p].Range[0, i + 1]).Value2 != null)
                                emptyColumn = false;
                        }
                    }

                    //add column if it has contents
                    if (!emptyColumn)
                        array[i] = RollupSources.Item(j).thisTable.ListColumns[i + 1].Name;
                }
                masterArray[j] = array;
            }

            //added per ticket #66443---------------------
            //checks the years of the plant to find mismatches that could cause bad data in the corporate roll-up
            bool baselineFlag = false;
            bool overFive = false;
            string badPlants = "";
            int breakCount = 0;

            for (int bline = 0; bline < baselineYears.Length; bline++)
            {
                //the first plant is used as a frame of reference (BJV thinks this is a bad idea but was vetoed my mgmt)
                if (baselineYears[0] != baselineYears[bline] && breakCount < 5)
                {
                    badPlants += RollupSources.Item(bline).PlantName + ", ";
                    breakCount++;
                    baselineFlag = true;
                }
                else if (baselineFlag && breakCount >= 5)
                {
                    //only display up to five plants in the error message otherwise use a generic message.
                    overFive = true;
                    break;
                }
            }

            //if there is an error in the year comparision display the dialog to the user
            if (baselineFlag)
            {
                //trim the last comma and space from the badPlants list
                badPlants = badPlants.Remove(badPlants.Length - 2, 2);

                DialogResult result = MessageBox.Show(overFive ?
                    "The \"Period\" labels for over 5 sheets do not match. You must use the same period labels in the plant analysis and corporate roll-up; otherwise, the plant level results will not line up. Would you like to proceed with the current period labels?" :
                    "The \"Period\" labels for Plants: " + badPlants + " do not match. You must use the same period labels in the plant analysis and corporate roll-up; otherwise, the plant level results will not line up. Would you like to proceed with the current period labels?"
                    , "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.No)
                {
                    //if the user does not wish to continue, stop the process and hide the roll-up sheet
                    thisSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    return;
                }
            }
            //--------------------------------------------

            //begin construction of header to be added to final report sheet
            string[] tableHeader = new string[masterColCount - masterArray.Length];
            int k = 0;

            foreach (string[] str in masterArray)
            {
                // start from the second column if the first is "Name"
                //int colCheck = 1;
                //if(str[0].Equals("Name"))
                //    colCheck = 2;

                for (int i = 1; i < str.Length; i++)
                {
                    //only add column if it doesn't already exist from another sheet
                    if (!tableHeader.Contains(str[i]))
                        tableHeader[(i - 1) + k] = str[i].ToString();
                }
                k += str.Length - 2;
            }

            thisSheet.get_Range("A2", "A2").Interior.Color = 0x006000;
            int[] baselineYearsIndices = new int[sourcesCount];

            for (int i = 0; i < tableHeader.Length; i++)
            {
                if (tableHeader[i] != null)
                {
                    headerCount += 1;
                    Excel.Range rng = thisSheet.get_Range(GetExcelColumnName(headerCount) + "2", GetExcelColumnName(headerCount) + "2");
                    rng.Value2 = tableHeader[i].ToString();
                    rng.Font.Color = 0xFFFFFF;
                    rng.Interior.Color = 0x006000;
                }
            }

            foreach (DetailTable dt in RollupSources)
            {
                addPlantTemplate(dt, headerCount);
            }

            addCorporateTotals(sourcesCount, headerCount);

            thisSheet.Columns.AutoFit();
        }

        private void addPlantTemplate(DetailTable dt, int headerCount)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet as Excel.Worksheet;
            Excel.Range rngHeader = BottomCell().get_Offset(1, 0).get_Resize(1, headerCount);
            rngHeader.Merge(true);
            rngHeader.Value2 = dt.PlantName;
            rngHeader.Font.Color = 0xFFFFFF;
            rngHeader.Interior.Color = 0x00B000;

            //per ticket #66434 --------------------------------------------
            //this adds the values for every energy type for each plant to the roll-up
            Excel.Range trimmedRange = thisSheet.get_Range(dt.TableRange.Substring(0, dt.TableRange.IndexOf(":")));

            for (int k = 0; k < dt.numOfSources; k++)
            {
                Excel.Range tmpRng = BottomCell().get_Offset(1, 0);
                tmpRng.Formula = "=\'" + dt.TableName + "\'!" + trimmedRange.get_Offset(k + 1).Address.ToString();
                //add the names to a list so we can compare them for the total roll-up
                EnergySourceNamesMaster.Add(tmpRng.Value2.ToString());
                for (int i = 1; i < headerCount; i++)
                {
                    for (int j = 0; j < dt.thisTable.ListColumns.Count; j++)
                    {
                        if (thisSheet.get_Range(GetExcelColumnName(i + 1) + "2", GetExcelColumnName(i + 1) + "2").Value2.ToString().Equals(dt.thisTable.ListColumns[j+1].Name))
                            tmpRng.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + trimmedRange.get_Offset(k + 1, j).Address.ToString();
                    }
                }
                tmpRng.EntireRow.NumberFormat = "#,##0";
            }

            //-------------------------------------------------------------

            Excel.Range rng1 = BottomCell().get_Offset(1, 0);
            rng1.Value2 = name_unadj;
            addValueLinks(dt, headerCount, rng1, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
            Excel.Range rng2 = BottomCell().get_Offset(1, 0);
            rng2.Value2 = name_adj;
            addValueLinks(dt, headerCount, rng2, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
            Excel.Range rng3 = BottomCell().get_Offset(1, 0);
            rng3.Value2 = name_aimp;
            addValueLinks(dt, headerCount, rng3, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
            Excel.Range rng4 = BottomCell().get_Offset(1, 0);
            rng4.Value2 = name_kpi;
            addValueLinks(dt, headerCount, rng4, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
            Excel.Range rng5 = BottomCell().get_Offset(1, 0);
            rng5.Value2 = name_savings;
            addValueLinks(dt, headerCount, rng5, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
            Excel.Range rng6 = BottomCell().get_Offset(1, 0);
            rng6.Value2 = name_bladj;
            addValueLinks(dt, headerCount, rng6, dt.fromActual, dt.hasProd, dt.hasBuildSqFt);
        }

        //private void addValueLinks(DetailTable dt, int headerCount, Excel.Range row, bool fromActual, bool hasProd, bool hasBuildSqFt)
        //{
        //    Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet as Excel.Worksheet;

        //    for (int i = 1; i < headerCount; i++)
        //    {
        //        for (int j = 1; j <= dt.thisTable.ListColumns.Count; j++)
        //        {
        //            //checks to see if the row exists in the imported sheet before adding a formula
        //            if (thisSheet.get_Range(GetExcelColumnName(i + 1) + "2", GetExcelColumnName(i + 1) + "2").Value2.ToString().Equals(dt.thisTable.ListColumns[j].Name))
        //            {
        //                int rowNum = fromActual ? (dt.numOfSources) : ((dt.numOfSources * 2) - 1);
        //                //add formula's to their corresponding rows
        //                switch (row.Value2.ToString())
        //                {
        //                    case "TOTAL Primary Energy Consumed (MMBtu/year)":
        //                        int one = 5;
        //                        if (fromActual)
        //                            one = 4;
        //                        row.get_Offset(0, i).Value2 = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (one + (dt.numOfSources > 0 ? (dt.numOfSources + (fromActual && dt.import? 1 : 0)) : 0)).ToString();//3
        //                        row.EntireRow.NumberFormat = "#,##0";
        //                        break;
        //                    case "TOTAL MODELED Primary Energy Consumed (MMBtu/year)":
        //                        int two = 11 + (dt.numOfSources == 1 && !dt.import && !fromActual ? 1 : 0);
        //                        if (hasProd)
        //                            two += 2;
        //                        if (hasBuildSqFt)
        //                            two++;
        //                        if (dt.fromEnergyCost)
        //                            two += dt.numOfSources;
        //                        if (fromActual)
        //                            two = 5 + (dt.numOfSources == 1 && !dt.import ? 1 : 0);
        //                        row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (two + (dt.numOfSources> 1 ? rowNum + (fromActual && dt.fromEnergyCost ? 1 : 0) : (fromActual && dt.import ? dt.numOfSources + 1 : 0))).ToString();//9
        //                        row.EntireRow.NumberFormat = "#,##0";
        //                        break;
        //                    case "Annual Improvement (%)":
        //                        int three = 14 + (dt.numOfSources == 1 && !dt.import && !fromActual ? 1 : 0);
        //                        if (hasProd)
        //                            three = three + 2;
        //                        if (hasBuildSqFt)
        //                            three++;
        //                        if (dt.fromEnergyCost)
        //                            three += dt.numOfSources;
        //                        if (fromActual)
        //                         three = 9 + (dt.numOfSources == 1 && !dt.import ? 1 : 0) + dt.numOfSources+((hasBuildSqFt)?1:0); //Added by suman TFS Ticket 69473
        //                         row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (three + (dt.numOfSources > 1 ? ((hasBuildSqFt && !hasProd)?0:rowNum) + (fromActual && dt.fromEnergyCost ? (dt.numOfSources * 2 + 1) : 0) : (fromActual && dt.import ? dt.numOfSources + (hasProd ? 1 : 0) + (dt.fromEnergyCost ? (dt.numOfSources * 2) : 0) : 0))).ToString();//12
        //                        row.EntireRow.NumberFormat = "0.0%";
        //                        break;
        //                    case "Total Improvement (%)":
        //                        int four = 13 + (dt.numOfSources == 1 && !dt.import && !fromActual ? 1 : 0);
        //                        if (hasProd)
        //                            four = four + 2;
        //                        if (hasBuildSqFt)
        //                            four++;
        //                        if (dt.fromEnergyCost)
        //                            four += dt.numOfSources;
        //                        if (fromActual)
        //                            four = 8 + (dt.numOfSources == 1 && !dt.import ? 1 : 0) + dt.numOfSources+((hasBuildSqFt)?1:0); //Added by suman TFS Ticket 69473
        //                        row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (four + (dt.numOfSources > 1 ?  ((hasBuildSqFt && !hasProd)?0:rowNum) + (fromActual && dt.fromEnergyCost ? (dt.numOfSources * 2 + 1) : 0) : (fromActual && dt.import ? dt.numOfSources + (hasProd ? 1 : 0) + (dt.fromEnergyCost ? (dt.numOfSources * 2) : 0) : 0))).ToString();//11
        //                        row.EntireRow.NumberFormat = "0.0%";
        //                        break;
        //                    case "New Energy Savings for Current Year (MMBtu/year)":
        //                        int five = 17 + (dt.numOfSources == 1 && !dt.import && !fromActual ? 1 : 0);
        //                        if (hasProd)
        //                            five= five + 2;
        //                        if (hasBuildSqFt)
        //                            five++;
        //                        if (dt.fromEnergyCost)
        //                            five += dt.numOfSources;
        //                        if (fromActual)
        //                            five = 11 + (dt.numOfSources == 1 && !dt.import ? 1 : 0) + dt.numOfSources+((hasBuildSqFt)?1:0);//Added by suman TFS Ticket 69473
        //                        row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (five + (dt.numOfSources > 1 ? ((hasBuildSqFt && !hasProd)?0:rowNum) + (fromActual && dt.fromEnergyCost ? (dt.numOfSources * 2 + 1): 0) : (fromActual && dt.import ? dt.numOfSources + (hasProd ? 1 : 0) + (dt.fromEnergyCost ? (dt.numOfSources * 2) : 0) : 0))).ToString();//15
        //                        row.EntireRow.NumberFormat = "#,##0";
        //                        break;
        //                    case "Total Energy Savings since Baseline Year (MMBtu/year)":
        //                        int six = 15 + (dt.numOfSources == 1 && !dt.import && !fromActual ? 1 : 0);
        //                        if (hasProd)
        //                            six = six + 2;
        //                        if (hasBuildSqFt)
        //                            six++;
        //                        if (dt.fromEnergyCost)
        //                            six += dt.numOfSources;
        //                        if (fromActual)
        //                            six = 10 + (dt.numOfSources == 1 && !dt.import ? 1 : 0) + dt.numOfSources+ ((hasBuildSqFt) ? 1 : 0); //Added by suman TFS Ticket 69473
        //                        row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + (six + (dt.numOfSources > 1 ? ((hasBuildSqFt && !hasProd)?0:rowNum) + (fromActual && dt.fromEnergyCost ? (dt.numOfSources * 2 + 1) : 0) : (fromActual && dt.import ? dt.numOfSources + (hasProd ? 1 : 0) + (dt.fromEnergyCost ? (dt.numOfSources * 2) : 0) : 0))).ToString();//13
        //                        row.EntireRow.NumberFormat = "#,##0";
        //                        break;
        //                }
        //            }
        //        }
        //    }
        //}
        private void addValueLinks(DetailTable dt, int headerCount, Excel.Range row, bool fromActual, bool hasProd, bool hasBuildSqFt)
        {
            Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveWorkbook.ActiveSheet as Excel.Worksheet;

            for (int i = 1; i < headerCount; i++)
            {
                for (int j = 1; j <= dt.thisTable.ListColumns.Count; j++)
                {
                    //checks to see if the row exists in the imported sheet before adding a formula
                    if (thisSheet.get_Range(GetExcelColumnName(i + 1) + "2", GetExcelColumnName(i + 1) + "2").Value2.ToString().Equals(dt.thisTable.ListColumns[j].Name))
                    {
                        //int rowNum = fromActual ? (dt.numOfSources) : ((dt.numOfSources * 2) - 1);
                        //add formula's to their corresponding rows
                        int startRowNum = 4; //Data always starts from 4th row in an given sheet.
                        switch (row.Value2.ToString())
                        {
                            case "TOTAL Primary Energy Consumed (MMBtu/year)":
                                startRowNum = startRowNum + GetRowNumber(EnPIResources.unadjustedTotalColName, dt.thisTable);
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "#,##0";
                                break;
                            case "TOTAL MODELED Primary Energy Consumed (MMBtu/year)":
                                startRowNum = startRowNum + GetRowNumber(fromActual ? EnPIResources.unadjustedTotalColName : EnPIResources.totalAdjValuesColName, dt.thisTable);
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "#,##0";
                                break;
                            case "Annual Improvement (%)":
                                startRowNum = startRowNum + GetRowNumber(EnPIResources.annualImprovementColName, dt.thisTable);
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "0.0%";
                                break;
                            case "Total Improvement (%)":
                                startRowNum = startRowNum + GetRowNumber(EnPIResources.totalImprovementColName , dt.thisTable);
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "0.0%";
                                break;
                            case "New Energy Savings for Current Year (MMBtu/year)":
                                startRowNum = startRowNum + GetRowNumber("New Energy Savings for Current Year (MMBtu/year)", dt.thisTable); //This not defined in Resources files , hence need to hard code
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "#,##0";
                                break;
                            case "Total Energy Savings since Baseline Year (MMBtu/year)":
                                startRowNum = startRowNum + GetRowNumber(fromActual ? "Total Savings Since Baseline Year (MMBtu/Year)" : "Total Energy Savings since Baseline Year (MMBtu/year)", dt.thisTable);//This not defined in Resources files , hence need to hard code
                                row.get_Offset(0, i).Formula = "=\'" + dt.TableName + "\'!" + GetExcelColumnName(j) + startRowNum.ToString();
                                row.EntireRow.NumberFormat = "#,##0";
                                break;
                        }
                    }
                }
            }
        }

        //this method searches for the given searchKeyword and returns the row number
        private int GetRowNumber(string searchKeyword,Excel.ListObject list)
        {
            int rowNumber = 1;
            object[,] firstColumnRange = list.ListColumns[1].DataBodyRange.Value as object[,];
            for (int count=1; count< firstColumnRange.Length;count++)
            {

                if (firstColumnRange[count, 1]!=null)
                {
                    if(firstColumnRange[count,1].ToString().ToUpper().Equals(searchKeyword.ToUpper()))
                    rowNumber = count;
                }
            }
            return rowNumber;
            
            
        }
        private void addCorporateTotals(int plantCount, int headerCount)
        {
            Excel.Range rngHeader = BottomCell().get_Offset(1, 0).get_Resize(1, headerCount);
            int rangeBottom = rngHeader.Row;
            rngHeader.Merge(true);
            rngHeader.Value2 = "Corporate Totals";
            rngHeader.Interior.Color = 0x006000;
            rngHeader.Font.Color = 0xFFFFFF;

            //per ticket #66434 --------------------------------------------
            //compare energy source names and add them to corporate totals
            IEnumerable<string> UniqueEnergySourceNames = EnergySourceNamesMaster.Distinct();

            foreach (string name in UniqueEnergySourceNames)
            {
                Excel.Range tmpRng = BottomCell().get_Offset(1, 0);
                tmpRng.Value2 = name.ToString();
                tmpRng.EntireRow.Font.Bold = true;
                tmpRng.EntireRow.NumberFormat = "#,##0";

                for (int i = 1; i < headerCount; i++)
                {
                    //tmpRng.get_Offset(0, i).Formula = "=SUMIF(A1:A" + Utilities.ExcelHelpers.writeAppendBottomAddress(WS, -2).ToString() + "," + tmpRng.Address.ToString() + "," + GetExcelColumnName(i + 1) + "1:" + GetExcelColumnName(i + 1) + Utilities.ExcelHelpers.writeAppendBottomAddress(WS, -2).ToString() + ")";
                    tmpRng.get_Offset(0, i).FormulaArray = "=SUM(IF(EXACT(A1:A" + rangeBottom + "," + tmpRng.Address.ToString() + ")," + GetExcelColumnName(i + 1) + "1:" + GetExcelColumnName(i + 1) + rangeBottom + "))";
                }
            }
            //--------------------------------------------------------------

            Excel.Range rng1 = BottomCell().get_Offset(1, 0);
            rng1.Value2 = name_unadj;
            addCoporateFormulas(plantCount, headerCount, rng1);
            rng1.EntireRow.Font.Bold = true;
            Excel.Range rng2 = BottomCell().get_Offset(1, 0);
            rng2.Value2 = "Adjustment for Baseline Primary Energy Use (MMBtu/year)";
            addCoporateFormulas(plantCount, headerCount, rng2);
            rng2.EntireRow.Font.Bold = true;
            Excel.Range rng3 = BottomCell().get_Offset(1, 0);
            rng3.Value2 = "Adjusted Baseline Primary Energy Use (MMBtu/year)";
            addCoporateFormulas(plantCount, headerCount, rng3);
            rng3.EntireRow.Font.Bold = true;
            Excel.Range rng4 = BottomCell().get_Offset(1, 0);
            rng4.Value2 = name_aimp;
            addCoporateFormulas(plantCount, headerCount, rng4);
            rng4.EntireRow.Font.Bold = true;
            Excel.Range rng5 = BottomCell().get_Offset(1, 0);
            rng5.Value2 = name_kpi;
            addCoporateFormulas(plantCount, headerCount, rng5);
            rng5.EntireRow.Font.Bold = true;
            Excel.Range rng6 = BottomCell().get_Offset(1, 0);
            rng6.Value2 = "New Energy Savings for Current Year (MMBtu/year)";
            addCoporateFormulas(plantCount, headerCount, rng6);
            rng6.EntireRow.Font.Bold = true;
            Excel.Range rng7 = BottomCell().get_Offset(1, 0);
            rng7.Value2 = name_bladj;
            addCoporateFormulas(plantCount, headerCount, rng7);
            rng7.EntireRow.Font.Bold = true;

        }

        private void addCoporateFormulas(int plantCount, int headerCount, Excel.Range row)
        {
            for (int i = 1; i < headerCount; i++)
            {
                switch (row.Value2.ToString())
                {

                    case "TOTAL Primary Energy Consumed (MMBtu/year)":
                        string output = "SUM(";
                        int rowNum = 4;
                        for (int j = 0; j < plantCount; j++)
                        {
                            rowNum = rowNum + RollupSources.Item(j).numOfSources;

                            if (j.Equals(0))
                            {
                                output += "INDIRECT(ADDRESS(" + rowNum + ",COLUMN()))";
                            }
                            else
                                output += ",INDIRECT(ADDRESS(" + rowNum + ",COLUMN()))";
                            rowNum += 7;
                        }
                        output += ")";
                        row.get_Offset(0, i).Formula = "=" + output;
                        row.EntireRow.NumberFormat = "#,##0";
                        break;
                    case "Adjustment for Baseline Primary Energy Use (MMBtu/year)":
                        row.get_Offset(0, i).Formula = "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),5,0,1,1) + INDIRECT((ADDRESS(ROW()-1,COLUMN()))) - INDIRECT((ADDRESS(ROW()-1,2)))";
                        row.EntireRow.NumberFormat = "#,##0";
                        break;
                    case "Adjusted Baseline Primary Energy Use (MMBtu/year)":
                        row.get_Offset(0, i).Formula = "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1)+OFFSET(INDIRECT(ADDRESS(ROW(),2)),-2,0,1,1)";
                        row.EntireRow.NumberFormat = "#,##0";
                        break;
                    case "Annual Improvement (%)":
                        row.get_Offset(0, i).Formula = "=IF(ISERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1)),0,OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1))";
                        row.EntireRow.NumberFormat = "0.0%";
                        break;
                    case "Total Improvement (%)":
                        string numerator = "(";
                        string denom = "(";
                        int rowNum5 = 5;
                        int rowNum5_2 = 7;
                        for (int j = 0; j < plantCount; j++)
                        {
                            rowNum5 += RollupSources.Item(j).numOfSources;
                            rowNum5_2 += RollupSources.Item(j).numOfSources;

                            if (j.Equals(plantCount - 1))
                            {
                                numerator += "($B$" + rowNum5 + "*(INDIRECT(ADDRESS(" + rowNum5_2 + ",COLUMN()))))";
                                denom += "$B$" + rowNum5;
                            }
                            else
                            {
                                numerator += "($B$" + rowNum5 + "*(INDIRECT(ADDRESS(" + rowNum5_2 + ",COLUMN()))))+";
                                denom += "$B$" + rowNum5 + "+";
                            }

                            rowNum5 += 7;
                            rowNum5_2 += 7;
                        }
                        numerator += ")";
                        denom += ")";
                        row.get_Offset(0, i).Formula = "=" + numerator + "/" + denom;
                        row.EntireRow.NumberFormat = "0.0%";
                        break;
                    case "New Energy Savings for Current Year (MMBtu/year)":
                        row.get_Offset(0, i).Formula = "=IF(ISERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1)),0,OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1))";
                        row.EntireRow.NumberFormat = "#,##0";
                        break;
                    case "Total Energy Savings since Baseline Year (MMBtu/year)":
                        string output7 = "SUM(";
                        int rowNum7 = 9;
                        for (int j = 0; j < plantCount; j++)
                        {
                            //add the number of sources
                            rowNum7 += RollupSources.Item(j).numOfSources;

                            if (j.Equals(0))
                            {
                                output7 += "INDIRECT(ADDRESS(" + rowNum7 + ",COLUMN()))";
                            }
                            else
                                output7 += ",INDIRECT(ADDRESS(" + rowNum7 + ",COLUMN()))";
                            rowNum7 += 7;
                        }
                        output7 += ")";
                        row.get_Offset(0, i).Formula = "=" + output7;
                        row.EntireRow.NumberFormat = "#,##0";
                        break;
                }
            }
        }

        private string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        //    string[] col_array = new string[2];
        //    col_array[0] = col_name;
        //    col_array[1] = "Values";

        //    Excel.Workbook thisWB = ((Excel.Workbook)WS.Parent);
        //    string pivotsource = PivotCacheSource(thisWB);

        //    Excel.PivotCache pivotch = thisWB.PivotCaches().Create(Excel.XlPivotTableSourceType.xlDatabase, pivotsource, System.Type.Missing);

        //    Excel.PivotTable PT = pivotch.CreatePivotTable(WS.get_Range("A2"), System.Type.Missing, System.Type.Missing, System.Type.Missing);

        //    PT.DisplayFieldCaptions = false;
        //    PT.RowGrand = false;
        //    PT.ColumnGrand = false;

        //    PT.TableStyle2 = "PivotStyleMedium4";

        //    PT.AddFields(col_array, col_year, System.Type.Missing, System.Type.Missing);

        //    AddPivotDataFields(PT);

        //    WS.Activate();

        //    //add the total section

        //    Excel.Range sumRangeHeader = BottomCell().get_Offset(1, 0).get_Resize(1, PT.ColumnRange.Count + 1);
        //    sumRangeHeader.Merge();
        //    sumRangeHeader.Interior.Color = 0x3D9375;
        //    sumRangeHeader.Font.Color = 0xFFFFFF;
        //    sumRangeHeader.Value2 = "Corporate Totals";

        //    Excel.Range sumRange = BottomCell().get_Offset(1, 0).get_Resize(7, PT.ColumnRange.Count + 1);
        //    sumRange.Font.Bold = true;

        //    for (int i = 1; i <= PT.ColumnRange.Count; i++)
        //    {
        //        string row1 = "";
        //        string row2 = "";
        //        string row3 = "";
        //        string row4 = "";
        //        string row5 = "";
        //        string row6 = "";
        //        string row7 = "";
        //        string denom = "SUM(";

        //        //format the totals section to match pivot table
        //        ((Excel.Range)sumRange[1, i + 1]).NumberFormat = pivot_formats[0];
        //        ((Excel.Range)sumRange[2, i + 1]).NumberFormat = pivot_formats[1];
        //        ((Excel.Range)sumRange[3, i + 1]).NumberFormat = pivot_formats[5];
        //        ((Excel.Range)sumRange[4, i + 1]).NumberFormat = pivot_formats[3];
        //        ((Excel.Range)sumRange[5, i + 1]).NumberFormat = pivot_formats[4];
        //        ((Excel.Range)sumRange[6, i + 1]).NumberFormat = pivot_formats[5];
        //        ((Excel.Range)sumRange[7, i + 1]).NumberFormat = pivot_formats[5];

        //        Excel.Worksheet pivotSource = (Excel.Worksheet)thisWB.Worksheets[rawIndex];
        //        int rows = BottomCell(pivotSource).Row - 2;
        //        string rawName = "'" + pivotSource.Name +"'!";

        //        string criteriayrcurrent = "SUMIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(ADDRESS(2,COLUMN()))"
        //                    + ")";
        //        string criteriayrbaseline = "SUMIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ")";
        //        string criteriayrspec2 = null;
        //        string criteriayravgspec2 = null;
        //        string criteriayrspec2baseline = null;
        //        string criteriayrbaselinespec = null;
        //        string criteriaavgyr = "AVERAGEIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(ADDRESS(2,COLUMN()))"
        //                    + ")";

        //        //string row2formula = "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),1,0,1,1) - INDIRECT((ADDRESS(ROW()-1,2)))";
        //        string row2formula = "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),5,0,1,1) + INDIRECT((ADDRESS(ROW()-1,COLUMN()))) - INDIRECT((ADDRESS(ROW()-1,2)))";
        //        //tring row3formula ="=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),4,0,1,1)+OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-2,0,1,1)";
        //        string row3formula = "=OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1)+OFFSET(INDIRECT(ADDRESS(ROW(),2)),-2,0,1,1)";

        //        string row4formula = "=IFERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1),0)";

        //        if (numPlant.Equals(1))
        //        {

        //            row1 = "=GETPIVOTDATA(\"" + name_unadj + "\", $A$2, \"Name\", INDIRECT(\"A3\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2)))";
        //            row2 = row2formula;
        //            row3 = row3formula;
        //            row4 = row4formula;
        //            //row4 = "=((" + string.Format(criteriayrcurrent, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")*" + string.Format(criteriayrcurrent, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")/" + string.Format(criteriayrbaseline, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
        //            row5 = "=((" + string.Format(criteriayrbaseline, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")*" + string.Format(criteriaavgyr, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")/" + string.Format(criteriayrbaseline, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
        //            row6 = "=IFERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1),0)";
        //            row7 = "=GETPIVOTDATA(\"" + name_bladj + "\", $A$2, \"Name\", INDIRECT(\"A3\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2)))";
        //        }
        //        else if (numPlant > 1)
        //        {
        //            row1 = "=SUM(";
        //            row2 = row2formula;
        //            row3 = row3formula;
        //            row4 = row4formula;
        //            //row4 = "=SUM(";
        //            row5 = "=SUM(";
        //            row6 = "=IFERROR(OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),1,-1,1,1),0)";
        //            row7 = "=SUM(";

        //            for (int k = 1; k <= numPlant; k++)
        //            {
        //                criteriayrspec2 = "SUMIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(\"A" + (7 * k - 4) + "\")"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(ADDRESS(2,COLUMN()))"
        //                    + ")";
        //                criteriayrspec2baseline = "SUMIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(\"A" + (7 * k - 4) + "\")"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ")";
        //                criteriayravgspec2 = "AVERAGEIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(\"A" + (7 * k - 4) + "\")"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(ADDRESS(2,COLUMN()))"
        //                    + ")";

        //                criteriayrbaselinespec = "SUMIFS(" + rawName + "{0}"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ", INDIRECT(\"A" + (7 * k - 4) + "\")"
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + "," + rawName + pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
        //                    + ")";
                        
        //                if(k != numPlant)
        //                {
        //                    row1 = row1 + "GETPIVOTDATA(\"" + name_unadj + "\", $A$2, \"Name\", INDIRECT(\"A" + (7 * k - 4) + "\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2))),";
        //                    //row4 = row4 + string.Format(criteriayrspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + "*" + string.Format(criteriayrspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ",";
        //                    row5 = row5 + string.Format(criteriayrspec2baseline, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + "*" + string.Format(criteriayravgspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ",";
        //                    denom = denom + string.Format(criteriayrbaselinespec, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ",";
        //                    row7 = row7 + "GETPIVOTDATA(\"" + name_bladj + "\", $A$2, \"Name\", INDIRECT(\"A" + (7 * k - 4) + "\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2))),";
        //                }
        //                else
        //                {
        //                    row1 = row1 + "GETPIVOTDATA(\"" + name_unadj + "\", $A$2, \"Name\", INDIRECT(\"A" + (7 * k - 4) + "\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2))))";
        //                    //row4 = row4 + string.Format(criteriayrspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + "*" + string.Format(criteriayrspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")";
        //                    row5 = row5 + string.Format(criteriayrspec2baseline, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + "*" + string.Format(criteriayravgspec2, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")";
        //                    denom = denom + string.Format(criteriayrbaselinespec, pivotSource.get_Range("A2").get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)) + ")";
        //                    row7 = row7 + "GETPIVOTDATA(\"" + name_bladj + "\", $A$2, \"Name\", INDIRECT(\"A" + (7 * k - 4) + "\"), \"Period\", INDIRECT((CHAR(COLUMN()+64)&2))))";
        //                }
        //            }

        //            //row4 = row4 + "/" + denom;
        //            row5 = row5 + "/" + denom;
        //        }
        //        else
        //        {
        //            row1 = "";
        //            row2 = "";
        //            row3 = "";
        //            row4 = "";
        //            row5 = "";
        //            row6 = "";
        //            row7 = "";
        //        }

        //        //for each 
        //        //equations
        //        ((Excel.Range)sumRange[1, i + 1]).Value2 = row1;
        //        ((Excel.Range)sumRange[2, i + 1]).Value2 = row2;
        //        ((Excel.Range)sumRange[3, i + 1]).Value2 = row3;
        //        ((Excel.Range)sumRange[4, i + 1]).Value2 = row4;
        //        ((Excel.Range)sumRange[5, i + 1]).Value2 = row5;
        //        ((Excel.Range)sumRange[6, i + 1]).Value2 = row6;
        //        ((Excel.Range)sumRange[7, i + 1]).Value2 = row7;
        //    }


        //    //Headers
        //    ((Excel.Range)sumRange[1, 1]).Value2 = name_unadj;
        //    ((Excel.Range)sumRange[2, 1]).Value2 = "Adjustment for Baseline Primary Energy Use (MMBtu/year)";
        //    ((Excel.Range)sumRange[3, 1]).Value2 = "Adjusted Baseline Primary Energy Use (MMBtu/year)";
        //    ((Excel.Range)sumRange[4, 1]).Value2 = name_aimp;
        //    ((Excel.Range)sumRange[5, 1]).Value2 = name_kpi;
        //    ((Excel.Range)sumRange[6, 1]).Value2 = "New Energy Savings for Current Year (MMBtu/year)";
        //    ((Excel.Range)sumRange[7, 1]).Value2 = name_bladj;

        //    AddIn.Globals.ThisAddIn.Application.AfterCalculate += new Excel.AppEvents_AfterCalculateEventHandler(Application_AfterCalculate);

        //    ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Columns.AutoFit();
        //}

        void Application_AfterCalculate()
        {
            if (!pivotRefresh)
            {
                pivotRefresh = true;
                ((Excel.PivotTable)((Excel.Worksheet)AddIn.Globals.ThisAddIn.Application.ActiveSheet).PivotTables(1)).RefreshTable();
                ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).Columns.AutoFit();
                
            }
        }

        private Excel.Range BottomCell()
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(WS, 0).ToString();

            return (Excel.Range)WS.get_Range(addr, System.Type.Missing);
        }

        internal string PivotCacheSource(Excel.Workbook WB)
        {
            Excel.Worksheet pivotSource = (Excel.Worksheet)WB.Worksheets.Add(System.Type.Missing, WB.Worksheets.get_Item(WB.Worksheets.Count), System.Type.Missing, System.Type.Missing);
            pivotSource.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            pivotSource.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(WB, raw_name, 0);

            pivotSource.get_Range("A1").get_Resize(1, pivot_src_hdrs.Count).Value2 = pivot_src_hdrs.ToArray();

            rawIndex = pivotSource.Index;

            foreach (DetailTable dt in RollupSources)
            {
                Excel.Range target = BottomCell(pivotSource).get_Offset(1, 0);
                AddLinkedColumns(dt.thisTable, target);
            }

            int rowct = BottomCell(pivotSource).Row;// -1; removed because the bottom row of the source data was not being included BJV 10/24/2012
            AddComputedColumns(pivotSource.get_Range("A2"), rowct);


            ((Excel.Range)pivotSource.Cells).Dirty();
            ((Excel.Range)pivotSource.Cells).Calculate();
            string datasource = pivotSource.get_Range("A1").get_Resize(rowct, pivot_src_hdrs.Count).get_Address(
                System.Type.Missing, System.Type.Missing, Excel.XlReferenceStyle.xlA1, true, System.Type.Missing);

            return datasource;
        }

        internal void AddLinkedColumns(Excel.ListObject dt, Excel.Range target)
        {
            int rows = dt.ListRows.Count;
            object[,] tmp = (object[,])dt.Range.Value;

            string sheetname = "='" + ((Excel.Worksheet)dt.Parent).Name + "'!{0}";
 
            // NAME Column
            int namendx = (Utilities.ExcelHelpers.GetListColumn(dt, col_name)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, col_name).Index : 0;
            if (namendx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).FormulaArray = LinkedValues(dt, namendx);
            // if no name exists, use the sheet name of the source sheet
            if (namendx == 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).Value = SheetNameFormula(dt);

            // YEAR Column
            int yrndx = (Utilities.ExcelHelpers.GetListColumn(dt, col_year)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, col_year).Index : 0;
            if (yrndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).FormulaArray = LinkedValues(dt, yrndx);

            // BASELINE YEAR Column
            string blycol = Globals.ThisAddIn.rsc.GetString("BaselineYearColName");
            int blyndx = (Utilities.ExcelHelpers.GetListColumn(dt, blycol)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, blycol).Index : 0;
            if (yrndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).FormulaArray = LinkedValues(dt, blyndx);

            // MODEL YEAR Column
            string myrcol = Globals.ThisAddIn.rsc.GetString("ModelYearColName");
            int myrndx = (Utilities.ExcelHelpers.GetListColumn(dt, myrcol)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, myrcol).Index : 0;
            if (myrndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).FormulaArray = LinkedValues(dt, myrndx);

            // BEFOR/AFTER MODEL YEAR Column
            if (myrndx > 0 && yrndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).FormulaArray = ModelYearBeforeAfter(dt, yrndx, myrndx);

            int unadjndx = (Utilities.ExcelHelpers.GetListColumn(dt, col_unadj)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, col_unadj).Index : 0;
            int adjndx = (Utilities.ExcelHelpers.GetListColumn(dt, col_adj)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, col_adj).Index : unadjndx;
            int proddx = (Utilities.ExcelHelpers.GetListColumn(dt, col_prod)) != null ? Utilities.ExcelHelpers.GetListColumn(dt, col_prod).Index : 0;

            //PRODUCTION
            if (proddx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_prod)).FormulaArray = LinkedValues(dt, proddx);//ProductionValues(dt);

            // BASELINE CONSUMPTION
            if (unadjndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).FormulaArray = BaselineValues(dt, yrndx, unadjndx);
            if (adjndx > 0)
                 target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjbl)).FormulaArray = BaselineValues(dt, yrndx, adjndx);

            // MODEL CONSUMPTION
            if (myrndx > 0 && unadjndx > 0)
                 target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjmdl)).FormulaArray = ModelValues(dt, yrndx, myrndx, unadjndx);
            if (myrndx > 0 && adjndx > 0)
                 target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjmdl)).FormulaArray = ModelValues(dt, yrndx, myrndx, adjndx);
             
            // ANNUAL CONSUMPTION
            if (unadjndx > 0)
                 target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).FormulaArray = LinkedValues(dt, unadjndx);
            if (adjndx > 0)
                target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).FormulaArray = LinkedValues(dt, adjndx);
            
            // UNADJ/ADJ RATIO
            // current year ratio

            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("= 0,((");
            arrayformula.Append(string.Format(SumByNameandYearBaseline(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            arrayformula.Append(string.Format(SumByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")/");
            arrayformula.Append(string.Format(SumByNameandYearBaseline(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("),IF(");
            arrayformula.Append(string.Format(AverageByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" < 1,");
            arrayformula.Append("IF(");
            arrayformula.Append(string.Format(AverageByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = 0, 1, ((");
            arrayformula.Append(string.Format(SumByNameandYearModel(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")*(");
            arrayformula.Append(string.Format(SumByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByNameandYearModel(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" ))),((");
            arrayformula.Append(string.Format(SumByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByNameandYearBaseline(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")*(");
            arrayformula.Append(string.Format(SumByNameandYearBaseline(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByNameandYear(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("))))");
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).Formula = arrayformula.ToString(); ;
            
            // baseline year ratio
            arrayformula = new StringBuilder();
            arrayformula.Append("=");
            arrayformula.Append(string.Format(SumByName(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByName(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blratio)).FormulaArray = arrayformula.ToString(); ;

            // model year ratio
            arrayformula = new StringBuilder();
            arrayformula.Append("=");
            arrayformula.Append(string.Format(SumByName(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjmdl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(SumByName(target, rows), target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjmdl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlratio)).FormulaArray = arrayformula.ToString(); ;
        }

        internal void AddComputedColumns(Excel.Range target, int rows)
        {
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_bladj)).Formula = BLAdjustmentFormula(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_senpi)).Formula = SEnPIFormula(target, rows);//Array
            string sdjfks = SEnPIFormulaOffset(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_senpi_step)).Formula = SEnPIFormulaOffset(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_kpi)).Formula = KPIFormula(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_aimp_step)).Formula = AIMPFormulaOffset(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_aimp)).Formula = AIMPFormula(target, rows);//Array
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).Formula = AnnualSavingsFormula(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs_step)).Formula = SavingsFormulaOffset(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step)).Formula = YearFormulaOffset(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).Formula = YearFormulaOffset2(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).Formula = CumulativeImprovFormula(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai)).Formula = AnnualImprovFormula(target, rows);
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci_step)).Formula = CIFormulaOffset(target, rows);//
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai_step)).Formula = AIFormulaOffset(target, rows);//
            target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_energyintens)).Formula = EnergyIntensityFormula(target, rows);
            
        }

        internal string LinkedValues(Excel.ListObject dt, int valuesndx)
        {
            StringBuilder arrayformula = new StringBuilder();
            string linkedrange = "'" + ((Excel.Worksheet)dt.Parent).Name + "'!{0}";

            arrayformula.Append("=");
            arrayformula.Append(string.Format(linkedrange, (dt.ListColumns[valuesndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));

            return arrayformula.ToString();
        }
        internal string ModelYearBeforeAfter(Excel.ListObject dt, int yrndx, int myrndx)
        {
            StringBuilder arrayformula = new StringBuilder();
            string strMatch = "MATCH({0},{1})";

            // compares the index of the current year value to the index of the model year value
            // returns -1 if it's before, 0 if it's equal, and 1 if it's after
            arrayformula.Append("=IFERROR(SIGN(");
            arrayformula.Append(string.Format(strMatch, LinkedValues(dt, yrndx).Substring(1)
                            , LinkedValues(dt, yrndx).Substring(1)));
            arrayformula.Append(" - ");
            arrayformula.Append(string.Format(strMatch, LinkedValues(dt, myrndx).Substring(1)
                            , LinkedValues(dt, yrndx).Substring(1)));
            arrayformula.Append("),0)");

            return arrayformula.ToString();
        }
        internal string SheetNameFormula(Excel.ListObject dt)
        {
            StringBuilder arrayformula = new StringBuilder();
            string sheetname = "'" + ((Excel.Worksheet)dt.Parent).Name + "'!$A$1";
            string formula = "IFERROR(RIGHT(CELL(\"filename\",{0}), LEN(CELL(\"filename\",{0})) - FIND(\"]\",CELL(\"filename\",{0}),1)),\"\")";

            arrayformula.Append("=");
            arrayformula.Append(string.Format(formula, sheetname));

            return arrayformula.ToString();
        }
        internal string BaselineValues(Excel.ListObject dt, int yrndx, int valuesndx)
        {
            StringBuilder arrayformula = new StringBuilder();
            string sheetname = "'" + ((Excel.Worksheet)dt.Parent).Name + "'!{0}";

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[yrndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append("=");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[yrndx].DataBodyRange.get_Resize(1, dt.ListColumns.Count).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append(",");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[valuesndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append(",0)");

            return arrayformula.ToString();
        }

        internal string ModelValues(Excel.ListObject dt, int yrndx, int myrndx, int valuesndx)
        {
            // =IF([year range]=[model year],[values range],0)
            StringBuilder arrayformula = new StringBuilder();
            string sheetname = "'" + ((Excel.Worksheet)dt.Parent).Name + "'!{0}";

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[yrndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append("=");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[myrndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append(",");
            arrayformula.Append(string.Format(sheetname, (dt.ListColumns[valuesndx].DataBodyRange.get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, false, System.Type.Missing))));
            arrayformula.Append(",0)");

            return arrayformula.ToString();
        }

        internal string BLAdjustmentFormula(Excel.Range target, int rows)
        {
            // Column 7
            // =[Adjusted annual baseline] - [Unadjusted annual baseline] + [Unadjusted report year] - [Adjusted report year]

            StringBuilder arrayformula = new StringBuilder();
             //sum group by name (column 1)
            string criteriabl = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            // sum group by name and year (columns 1 and 2)
            string criteriayr = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayr2 = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs_step)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",0, INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,1) - INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,2)),0)");

            return arrayformula.ToString();
        }
        internal string AnnualImprovFormula(Excel.Range target, int rows)
        {
            // Column 7
            // =[Adjusted annual baseline] - [Unadjusted annual baseline] + [Unadjusted report year] - [Adjusted report year]
            
            StringBuilder arrayformula = new StringBuilder();

            string criterianame = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayravg = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            string criteriayravgspec = "AVERAGEIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";
            string criteriayrsumspec = "SUMIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";

            arrayformula.Append("=IF(OR(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci_step)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = 0), ");
            //use actual calc
            arrayformula.Append("INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,1) - INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,2)");

            arrayformula.Append(", IF(");
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" < 1, ");
            //backcast calc
            arrayformula.Append("IFERROR(((1 - ");
            arrayformula.Append(string.Format(criteriayravgspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")-(1 - ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")),0)");
            //forcast calc
            arrayformula.Append(",INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,1) - INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,2)))");

            return arrayformula.ToString();
        }
        internal string SavingsFormula(Excel.Range target, int rows)
        {
            // Column 8
            // = [Unadjusted annual baseline] + [Baseline Adjustment]  - + [Unadjusted report year]

            StringBuilder arrayformula = new StringBuilder();
            //sum group by name
            string criteriabl = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            // sum group by name and year
            string criteriayr = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            
            arrayformula.Append("=0*(");
            arrayformula.Append(string.Format(criteriabl, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("+");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_bladj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("-");
            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")");

            return arrayformula.ToString();
        }
        internal string AnnualSavingsFormula(Excel.Range target, int rows)
        {
            //find if use actual or regression
            //IF(SUMIFS(/model year column/) = 0, *use actual calc*, *regression calc*)

            // year <= model year
            // = sum( [Adjusted report year] - [Unadjusted report year]) - sum([Adjusted baseline] - [Unadjusted baseline])
            // year > model year
            // = sum( [Adjusted report year] - [Unadjusted report year])

            StringBuilder arrayformula = new StringBuilder();

            string criterianame = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayr = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayrspec = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayravg = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            string criteriayravgspec = "AVERAGEIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";

            string criteriabaseline = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = 0, (");
            arrayformula.Append(string.Format(criteriabaseline, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("), IF(");
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" < 1, IFERROR((");
            // previous year savings
            arrayformula.Append(string.Format(criteriayravgspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("+((");
            //EC previous
            arrayformula.Append(string.Format(criteriayrspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            //Mod EC previous
            arrayformula.Append(string.Format(criteriayrspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")-(");
            //EC current
            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            //Mod EC current
            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("))),IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("=");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", 0, INDIRECT(ADDRESS(ROW()-1,COLUMN())))), (");

            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            arrayformula.Append(string.Format(criteriayr, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")))");

            return arrayformula.ToString();
                        
        }
        internal string CumulativeImprovFormula(Excel.Range target, int rows)
        {
            //find if use actual or regression

            // year <= model year
            // = sum( [Adjusted report year] - [Unadjusted report year]) - sum([Adjusted baseline] - [Unadjusted baseline])
            // year > model year
            // = sum( [Adjusted report year] - [Unadjusted report year])

            StringBuilder arrayformula = new StringBuilder();

            string criterianame = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayravg = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            string criteriayravgspec = "AVERAGEIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";
            string criteriayravgbaseline = "AVERAGEIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";
            string criteriayrsumspec = "SUMIFS({0}"
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step2)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                           + ")";

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = 0, (( ");
            //use actual calculation
            arrayformula.Append(string.Format(criteriayravgbaseline, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_energyintens)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("-");
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_energyintens)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")/");
            arrayformula.Append(string.Format(criteriayravgbaseline, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_energyintens)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));

            arrayformula.Append("), IF(");
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" < 1,");
            //backcast calc
            arrayformula.Append("IFERROR(((1 - ");
            arrayformula.Append(string.Format(criteriayravgspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")-(1 - ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(") + ");
            arrayformula.Append(string.Format(criteriayravgspec, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("),IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("=");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", 0, INDIRECT(ADDRESS(ROW()-1,COLUMN())))),(1 - ");
            //forcast calc
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("))");
            arrayformula.Append(")");
            
            return arrayformula.ToString();
            
        }
        internal string KPIFormula(Excel.Range target, int rows)
        {
            // This computes percent improvement from the baseline year
            // = 1 - ([Adjusted annual baseline] / [Unadjusted annual baseline]) * ([Unadjusted report year] / [Adjusted report year])
 
            StringBuilder arrayformula = new StringBuilder();
            //sum group by name 
            string baselinesum = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            // sum group by name and year 
            string reportyearsum = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";


            arrayformula.Append("=1-(");
            arrayformula.Append(string.Format(baselinesum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(baselinesum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")*(");
            arrayformula.Append(string.Format(reportyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(reportyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")");

            return arrayformula.ToString();
        }
        internal string AIMPFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_kpi)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),26)))");

            return arrayformula.ToString();
        }
        internal string SavingsFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_svgs)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),18)))");

            return arrayformula.ToString();
        }
        internal string YearFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),2)))");

            return arrayformula.ToString();
        }
        internal string CIFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ci)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),22)))");

            return arrayformula.ToString();
        }
        internal string AIFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ai)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),24)))");

            return arrayformula.ToString();
        }
        internal string EnergyIntensityFormula(Excel.Range target, int rows)
        {
            string criteriayearsum = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
            string criteriayearavg = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=");
            arrayformula.Append(string.Format(criteriayearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(criteriayearavg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_prod)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));

            return arrayformula.ToString();
        }
        internal string YearFormulaOffset2(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();
            
            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year_step)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",0, INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,2))");

            return arrayformula.ToString();
        }
        internal string AIMPFormula(Excel.Range target, int rows)
        {
            string criterianame = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            string criteriayravg = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = 0, 0, IF(");
            arrayformula.Append(string.Format(criteriayravg, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" < 1,1,");
            arrayformula.Append("IF(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_kpi)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(" = ");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_aimp_step)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",0, INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_kpi)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,1) - INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 2).get_Offset(0, pivot_src_hdrs.IndexOf(col_kpi)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(", ROW()-1,2))");
            arrayformula.Append("))");

            

            return arrayformula.ToString();
        }
        internal string WeightFormula(Excel.Range target, int rows)
        {
            // Column 10
            // =  [Unadjusted annual baseline] / [total unadjusted annual baseline]
            // = sumif([column 1]=[column 1 this row],[column 3],0)/
            //  sum([column 3])

            StringBuilder arrayformula = new StringBuilder();
            //sum group by name 
            string criteriabl = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            arrayformula.Append("=");
            arrayformula.Append(string.Format(criteriabl, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/SUM(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append(")");

            return arrayformula.ToString();
        }
        internal string WeightedImprovement(Excel.Range target, int rows)
        {  
           // Column 11
            // =  [% improvement] * [weight]

            StringBuilder arrayformula = new StringBuilder();
             arrayformula.Append("=");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, 8).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("*");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, 9).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));

            return arrayformula.ToString();

        }
        
        internal string SEnPIFormula(Excel.Range target, int rows)
        {
            // SEnPI
            // = power([Unadjusted report year] / [Adjusted report year], 
            // if [model year] < [report year] then -1 if [model year] > [report year] then 1 if [model year] = [report year] then 0)
            // * 
            // (if [model year] < [report year] then [Adjusted baseline year] / [Unadjusted baseline year] else 1)
            // * 
            // (if [model year] > [report year] then [Unadjusted model year] / [Adjusted model year] else 1)

            StringBuilder arrayformula = new StringBuilder();
            //sum group by name 
            string modelyearsum = SumByName(target, rows);
            // sum group by name and year 
            string reportyearsum = SumByNameandYear(target, rows);
            // before or after model year 

            // current year ratio, inverted or set to 1 based on year position before or after model year
            arrayformula.Append("=POWER((");
            arrayformula.Append(string.Format(reportyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("),");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append(")");
            // for years after the model year, multiply by inverse of the baseline ratio
            arrayformula.Append("*IF(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("=1,POWER(");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",-1),");
            // for years prior to the model year, multiply by inverse of the model ratio
            arrayformula.Append("IF(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("=-1,POWER(");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",-1), 1))");
            
            return arrayformula.ToString();

        }
        internal string SEnPIFormulaOffset(Excel.Range target, int rows)
        {
            StringBuilder arrayformula = new StringBuilder();

            string dsjfklsd = string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));

            arrayformula.Append("=IF(INDIRECT(ADDRESS(ROW(),1))=INDIRECT(ADDRESS(ROW()-1,1)),INDEX(");
            arrayformula.Append(string.Format(null, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(",ROW()-2,1),INDIRECT(ADDRESS(ROW(),20)))");

            return arrayformula.ToString();
        }
        internal string CumulativeImprovementFormula(Excel.Range target, int rows)
        {
            // Cumulative Improvement
            // = if [model year] < [report year] then = 1 - [Report year SEnPI]
            // if [model year] >= [report year] then = [Report year SEnPI] -[Baseline year SEnPI]
            // if [report year] = [baseline year] then = 0

            StringBuilder arrayformula = new StringBuilder();
            //sum group by name 
            string modelyearsum = SumByName(target, rows);
            // sum group by name and year 
            string reportyearsum = SumByNameandYear(target, rows);
            // before or after model year 

            // baseline year is always zero
            arrayformula.Append("=IF(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("=");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append(", 0, 1)");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            // for years after the model year, 1 - [Report year SEnPI]
            arrayformula.Append("*IF(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("=1,1 - ");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append(", 1)");
            // for years prior to or equal to the model year, [Report year SEnPI] -[Baseline year SEnPI]
            arrayformula.Append("*IF(");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ba)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("=1,1, ");
            arrayformula.Append(target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_ratio)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing));
            arrayformula.Append("-(");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjmdl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_adjmdl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append(")*(");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_bladj)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("/");
            arrayformula.Append(string.Format(modelyearsum, target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_unadjbl)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)));
            arrayformula.Append("))");

            return arrayformula.ToString();

        }

        internal string SumByName(Excel.Range target, int rows)
        {
            //assumes the values in the columns are zero if the year is not the baseline year ormodel year
            string criteria = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";
        
            return criteria;            
        }

        internal string SumByNameandYear(Excel.Range target, int rows)
        {
            string criteria = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            return criteria;
        }

        internal string SumByNameandYearModel(Excel.Range target, int rows)
        {
            string criteria = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_mdlyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            return criteria;
        }

        internal string SumByNameandYearBaseline(Excel.Range target, int rows)
        {
            string criteria = "SUMIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_blyear)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            return criteria;
        }

        internal string AverageByNameandYear(Excel.Range target, int rows)
        {
            string criteria = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_year)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            return criteria;
        }

        internal string AverageByName(Excel.Range target, int rows)
        {
            //assumes the values in the columns are zero if the year is not the baseline year ormodel year
            string criteria = "AVERAGEIFS({0}"
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + "," + target.get_Resize(rows, 1).get_Offset(0, pivot_src_hdrs.IndexOf(col_name)).get_Address(true, true, Excel.XlReferenceStyle.xlR1C1, System.Type.Missing, System.Type.Missing)
                            + ")";

            return criteria;
        }

        internal void AddPivotDataFields(Excel.PivotTable PT)
        {
            Excel.PivotFields flds = (Excel.PivotFields)PT.PivotFields();
            ArrayList cols = PivotFieldsArray((Excel.PivotFields)PT.ColumnFields);
            ArrayList hdrs = PivotFieldsArray((Excel.PivotFields)PT.RowFields);

            for (int i = 0; i < pivot_columns.Count(); i++)
            {
                string col = pivot_columns[i];
                string nm = pivot_captions[i];
                string fmt = pivot_formats[i];
                Excel.XlConsolidationFunction cons_fcn = pivot_aggs[i];

                try
                {
                    Excel.PivotField fld = (Excel.PivotField)flds.Item(col);

                    if (!cols.Contains(fld.Name) && !hdrs.Contains(fld.Name))
                    {
                        Excel.PivotField nf = PT.AddDataField(fld
                            , (nm == "" ? System.Type.Missing : nm)
                                , (cons_fcn == Excel.XlConsolidationFunction.xlUnknown ? System.Type.Missing : cons_fcn)  
                                );
                        nf.NumberFormat = fmt;
                    }
                }
                catch
                {
                }
            }

        }

        internal ArrayList PivotFieldsArray(Excel.PivotFields flds)
        {
            ArrayList ret = new ArrayList();
            foreach (Excel.PivotField fld in flds)
            {
                ret.Add(fld.Name);
            }
            return ret;
        }
    
        private Excel.Range BottomCell(Excel.Worksheet ws)
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(ws, 0).ToString();

            return (Excel.Range)ws.get_Range(addr, System.Type.Missing);
        }
}


    public class DetailTable
    {
        public string SourceSheet { get; set; }
        public string FileName { get; set; }
        public string TableName { get; set; }
        public string PlantName { get; set; }
        public string TableRange { get; set; }

        public string YearColumnSelect { get; set; }
        public string TotalColumnSelect { get; set; }
        public string TotalAdjColumnSelect { get; set; }
        public string ProductionColumnSelect { get; set; }
        public string EnPIColumnSelect { get; set; }
        public string BuildingSFColumnSelect { get; set; }
        public string SavingsColumnSelect { get; set; }
        public string IntensityColumnSelect { get; set; }
        public string SQLStatement { get; set; }
        public string DisplayName { get; set; }
        public int numOfSources { get; set; }
        public bool fromActual { get; set; }
        public bool hasProd { get; set; }
        public bool hasBuildSqFt { get; set; }
        public bool import { get; set; }
        public bool fromEnergyCost { get; set; }


        internal Excel.ListObject thisTable;

        public DetailTable(Excel.Worksheet sheet)
        {
            TableName = sheet.Name;
            FileName = ((Excel.Workbook)sheet.Parent).FullName;
            TableRange = "";
 
            thisTable = sheet.ListObjects[1];

            if (thisTable != null)
            {
                YearColumnSelect = ItemSelect("yearColName");
                TotalColumnSelect = ItemSelect("unadjustedTotalColName");
                TotalAdjColumnSelect = ItemSelect("totalAdjValuesColName");

                SetSQLStatement();
            }
        }
        
        public DetailTable(Excel.ListObject list, string plantName, int numOfSources, bool fromActual, bool hasProd, bool hasBuildSqFt, bool import, bool fromEnergyCost)
        {
            thisTable = list;
            PlantName = plantName;
            this.numOfSources = numOfSources;
            this.fromActual = fromActual;
            this.hasProd = hasProd;
            this.hasBuildSqFt = hasBuildSqFt;
            this.import = import;
            this.fromEnergyCost = fromEnergyCost;


            Excel.Worksheet WS = (Excel.Worksheet)list.Parent;
            FileName = ((Excel.Workbook)WS.Parent).FullName;
            DisplayName = plantName ;

            // can't handle table names with spaces or other odd characters
            TableName = WS.CodeName == "" ? WS.Name : WS.CodeName;

            SourceSheet = Utilities.ExcelHelpers.getWorksheetCustomProperty(WS, "SheetGUID");
            TableRange = list.Range.get_Address(false, false, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

            YearColumnSelect = ItemSelect("yearColName");
            TotalColumnSelect = ItemSelect("unadjustedTotalColName");
            TotalAdjColumnSelect = ItemSelect("totalAdjValuesColName");

            SetSQLStatement();      
        }

        public DetailTable(string strSQLstmt, string strDisplayName)
        {
            SQLStatement = strSQLstmt;
            DisplayName = strDisplayName;
        }

        public string ItemColName(string rscName)
        {
            return Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString(rscName)).Replace(".", "");
        }

        internal string ItemSelect(string rscName)
        {
            string itm = Utilities.ExcelHelpers.CreateValidFormulaName(
                (Utilities.ExcelHelpers.GetListColumnName(thisTable, Globals.ThisAddIn.rsc.GetString(rscName)) ?? "")
                    );

            if (itm != Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString(rscName)))
            {
                itm += " AS " + ItemColName(rscName);
            }

            return itm;
        }

        internal void SetSQLStatement()
        {
            SQLStatement = "";
            string nm = PlantName ?? FileName;
            string mt = Utilities.ExcelHelpers.CreateValidFormulaName("");
            string flnm = ("[" + FileName + "].").Replace("[].","");

            if ((YearColumnSelect != mt + " AS " + Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString("yearColName")).Replace(".","")
                && TotalColumnSelect != mt + " AS " + Utilities.ExcelHelpers.CreateValidFormulaName(Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName")).Replace(".", "")
                ))
            {
                SQLStatement = "SELECT \"" + nm + "\" as [Name]"
                    + ", " + YearColumnSelect
                    + ", " + TotalColumnSelect
                    + ", " + TotalAdjColumnSelect
                    //+ ", *"  // no good, because the number of columns needs to match
                    // need some way to add energy sources and give them a common name
                    + " FROM " + flnm + "[" + TableName + "$" + TableRange + "]";
             }

        }
 
        public List<string> ReportYears()
        {
            List<string> years = new List<string>();

            try
            {
                object[] ly = Utilities.ExcelHelpers.getYears(thisTable);

                for (int i = ly.GetLowerBound(0); i <= ly.GetUpperBound(0); i++)
                {
                    years.Add(ly[i].ToString());
                }
            }
            catch
            {
            }

            return years;
        }

        public string DetailTableXML()
        {
            System.Text.StringBuilder str = new System.Text.StringBuilder();
            System.Xml.XmlWriter tXML = System.Xml.XmlWriter.Create(str);
            tXML.WriteStartElement("DetailTable");

            if (this.SQLStatement == "" || this.DisplayName == "")
            {
                tXML.WriteElementString("PlantName", this.PlantName);
                tXML.WriteElementString("FileName", this.FileName);
                tXML.WriteElementString("TableName", this.TableName);
                tXML.WriteElementString("SourceSheet", this.SourceSheet);
                tXML.WriteElementString("FileName", this.FileName);
                tXML.WriteElementString("TableRange ", this.TableRange);
                tXML.WriteElementString("YearColumnSelect", this.YearColumnSelect);
                tXML.WriteElementString("TotalColumnSelect", this.TotalColumnSelect);
                tXML.WriteElementString("ProductionColumnSelect", this.ProductionColumnSelect);
                tXML.WriteElementString("EnPIColumnSelect", this.EnPIColumnSelect);
                tXML.WriteElementString("BuildingSFColumnSelect", this.BuildingSFColumnSelect);
                tXML.WriteElementString("SavingsColumnSelect", this.SavingsColumnSelect);
                tXML.WriteElementString("IntensityColumnSelect", this.IntensityColumnSelect);
            }
            else
            {
                tXML.WriteElementString("SQLStatement", this.SQLStatement);
                tXML.WriteElementString("DisplayName", this.DisplayName);
            }
 
            tXML.WriteEndElement();
            tXML.Close();
            
            return str.ToString();
        }
    }

    public class SourceTable
    {
        public string TableName { get; set; }
        public string SQLStatement { get; set; }
        public string DisplayName { get; set; }

        public SourceTable(string strSQL, string strDisplayname)
        {
            this.SQLStatement = strSQL;
            this.DisplayName = strDisplayname;
        }
    }
    public class SourceFile
    {
        public string FileName {get;set;}
        public string SQLStatement { get; set; }
        public string DisplayName {get; set;}
        public string ShortName { get; set; }
        public bool fromActual { get; set; }
        public int numOfSources { get; set; }
        public bool hasProd { get; set; }
        public bool hasBuildSqFt { get; set; }

        public SourceFile(string strSheetName, string strFileName, int numOfSources, bool fromActual)//, bool hasProd, bool hasBuildSqFt)
        {
            this.FileName = strFileName;
            this.DisplayName = "(" + strFileName + ") " + strSheetName;
            int rows = 18;
            int startRow= 4;
            //if (fromActual)
            //    startRow = 3; // TFS Ticket 71232.
            this.SQLStatement = "SELECT * FROM [" + strFileName + "].[" + strSheetName + "$" + startRow.ToString() + ":" + (rows + (numOfSources * 4)).ToString() + "]";
            this.ShortName = strSheetName;
            this.fromActual = fromActual;
            this.numOfSources = numOfSources;
            //this.hasProd = hasProd;
            //this.hasBuildSqFt = hasBuildSqFt;
       }

        public SourceFile(string strSQL, string strDisplayname, string strFileName, bool fromActual)
        {
            this.FileName = strFileName;
            this.SQLStatement = strSQL;
            this.DisplayName = strDisplayname;
            this.fromActual = fromActual;
        }

        public void WriteFileXML(System.Xml.XmlWriter tXML)
        {
            tXML.WriteStartElement("SourceFile");

            tXML.WriteElementString("FileName", this.FileName);
            tXML.WriteElementString("SQLStatement", this.SQLStatement);
            tXML.WriteElementString("DisplayName", this.DisplayName);

            tXML.WriteEndElement();
        }

        public override Boolean Equals(object sf)
        {
            if (sf.GetType() == System.Type.GetType("System.DBNull"))
                return false;

            if (this.FileName == ((SourceFile)sf).FileName && this.DisplayName == ((SourceFile)sf).DisplayName)
                return true;
            else
                return false;
        }
    }

    public class DetailTableCollection : System.Collections.CollectionBase
    {
        public void Add(Excel.ListObject list)
        {
            List.Add(new DetailTable(list, null, 0, false, false, false, false, false));
        }
        public void Add(DetailTable aTable)
        {
            List.Add(aTable);
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

        public DetailTable Item(int Index)
        {
            return (DetailTable)List[Index];
        }
    }

}
