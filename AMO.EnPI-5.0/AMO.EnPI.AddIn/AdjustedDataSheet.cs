using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms.VisualStyles;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;
using Microsoft.Office.Interop.Excel;

namespace AMO.EnPI.AddIn
{
    class AdjustedDataSheet
    {

        public Utilities.EnPIDataSet DS;
        public Excel.Worksheet thisSheet;
        public Excel.Worksheet SourceSheet;
        public Excel.ListObject SourceObject;
        public object[,] SourceData;
        public string[] ColumnFormatting;
        public Excel.ListObject AdjustedData;
        public Excel.ListObject ValidationTable;
        public ArrayList Warnings;
        internal float lineht = 14; //HACK: this will change if the user's "Normal" font size is different 
        string fmt = "General";
        public IList<SEPValidationValues> lstWarningValidationValues = new List<SEPValidationValues>();

        public AdjustedDataSheet(Utilities.EnPIDataSet DSIn)
        {
            DS = DSIn;
            string nm = Globals.ThisAddIn.rsc.GetString("adjustedDataName");
            if (DS.ModelYear == null)
                nm = Globals.ThisAddIn.rsc.GetString("unadjustedDataName");

            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            SourceSheet = Utilities.ExcelHelpers.GetWorksheet(WB, DS.WorksheetName);

            Excel.Worksheet aSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add
                (System.Type.Missing, WB.Sheets.get_Item(WB.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet);
            aSheet.CustomProperties.Add("SheetGUID", System.Guid.NewGuid().ToString());
            
            aSheet.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(WB, nm, Globals.ThisAddIn.groupSheetCollection.regressionIteration);
            aSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;
            aSheet.Tab.Color = 0x50CC11;
            Utilities.ExcelHelpers.addWorksheetCustomProperty(aSheet, Utilities.Constants.WS_ISENPI, "True");
           // Utilities.ExcelHelpers.addWorksheetCustomProperty(aSheet, Utilities.Constants.WS_ROLLUP, "True");

            thisSheet = aSheet;
            Warnings = new ArrayList();
        }

        private Excel.Range BottomCell()
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(thisSheet, 0).ToString();

            return (Excel.Range)thisSheet.get_Range(addr, System.Type.Missing);
        }

        public void Populate(bool regression)
        {
            System.Text.StringBuilder hdrtxt = new System.Text.StringBuilder();

            hdrtxt.AppendLine(SourceSheet.Name);

            if (DS.ModelYear != null)
            {
                hdrtxt.AppendLine();
                hdrtxt.AppendLine("" + EnPIResources.adjustedDataTitle + "");
                hdrtxt.AppendLine(EnPIResources.adjustedDataText);
            }

            Excel.Shape txtBox1 = thisSheet.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 1, 1, 700, 130);
            txtBox1.Placement = Excel.XlPlacement.xlMove;
            txtBox1.TextFrame.Characters().Text = hdrtxt.ToString();
            txtBox1.Height = lineht * 13; //HACK: ten rows is dependent on the width and font size  
            
            SourceData = Utilities.DataHelper.dataTableArrayObject(DS.SourceData); 
            SourceObject = Utilities.ExcelHelpers.GetListObject(SourceSheet, DS.ListObjectName);
        
            WriteSourceData();
            
            if(Globals.ThisAddIn.SelectedProduction.Count > 0)
                WriteTotalProd();

            WriteBaselineYear(DS.BaselineYear);
            WriteModelYear(DS.ModelYear);

            if (!regression)
            {
                WritePeriodCount();
                WriteBaselineCount();
            }

            if (regression)
            {
                WriteLastYear(DS.Years[DS.Years.Count - 1]);
                WriteAdjustedModel();
            }
            WriteAdjustedData();
            //Added by Suman TFS Ticket: 66429
            if (regression)
            {
                WriteCUSUMHidden();
                WriteCUSUM();
            }
            //ticket #66426
            WriteEnergySavings(regression);
            

            //Commented by Suman: as these columns are required
            // for Actuals as well
            if (Globals.ThisAddIn.fromEnergyCost)  //&& regression) 
            {
                WriteUnitCost();
                WriteCostSavings();
            } 
            FormatAdjustedData();
            //Added By Suman: As per the new SEP Changes
            if (regression)
            {
                WriteSEPColumnsData();
            }

            
            if (regression)
                WriteNegativeWarning();

            Excel.Range negativeMessageHeader = thisSheet.get_Range("A2");
            Excel.Range negativeMessageDescription = thisSheet.get_Range("A3");

            if(regression)
                if (Globals.ThisAddIn.NegativeCheck(AdjustedData, Globals.ThisAddIn.modeledSourceIndex))
                {
                    negativeMessageHeader.EntireRow.Hidden = false;
                    negativeMessageDescription.EntireRow.Hidden = false;
                }
                else
                {
                    negativeMessageHeader.EntireRow.Hidden = true;
                    negativeMessageDescription.EntireRow.Hidden = true;
                }
         
            AddVariableWarnings(txtBox1);

            thisSheet.get_Range("A1").EntireRow.RowHeight = txtBox1.Top + txtBox1.Height + float.Parse(thisSheet.get_Range("A1").Height.ToString());

            GroupSheet GS = new GroupSheet(thisSheet, false, true, thisSheet.Name);
            GS.Name = SourceSheet.Name;
            Globals.ThisAddIn.groupSheetCollection.Add(GS);

            if (regression)
                WriteValidationCheckTable();

            thisSheet.Visible = Excel.XlSheetVisibility.xlSheetVisible;
        }

        internal void AddVariableWarnings(Excel.Shape txtBox)
        {
            System.Text.StringBuilder errs = new System.Text.StringBuilder();
            errs.Append(txtBox.TextFrame.Characters().Text);
            errs.AppendLine();

            AMO.EnPI.AddIn.Utilities.Model mdl;

            foreach (string st in Warnings)
            {
                errs.AppendLine(st);
            }
            errs.AppendLine();

            txtBox.TextFrame.Characters().Text = errs.ToString();

            float scale = 1 + ((Warnings.Count + 1) * lineht) / txtBox.Height; 

            txtBox.ScaleHeight(scale, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoScaleFrom.msoScaleFromTopLeft);

        }

        internal void WriteTotalProd()
        {
            StringBuilder arrayformula = new StringBuilder();

            arrayformula.Append("=SUM(");
            foreach (string str in Globals.ThisAddIn.SelectedProduction)
            {
                arrayformula.Append("SUMIFS(");
                arrayformula.Append(AdjustedData.Name + "[" + str + "],");
                arrayformula.Append(AdjustedData.Name + "[" + EnPIResources.yearColName + "],");
                arrayformula.Append(AdjustedData.Name + "[" + EnPIResources.yearColName + "]");
                arrayformula.Append("),");
            }
            //remove last comma 
            arrayformula.Remove(arrayformula.Length - 1, 1);
            arrayformula.Append(")");

            string colName = Globals.ThisAddIn.rsc.GetString("productionColName");
            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[newcol.Index].Name = colName;

            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = arrayformula.ToString();

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WriteValidationCheckTable()
        {
            Excel.Range titleRange = thisSheet.get_Range("A4");

            titleRange.Value2 = "Validation Check";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;

            Excel.Range descriptionRange = thisSheet.get_Range("A5:G5");
            descriptionRange.Merge();
            descriptionRange.RowHeight = 130;

            descriptionRange.Value2 = "A model must satisfy validity requirements in order to be used for SEP or Better Plants reporting. In addition to the model having acceptable R-squared and p-values, the average of the variables entered into the model must fall within one of the following ranges: \r\n\r\n 1. The range of observed data that went into the model OR \r\n 2. Three standard deviations from the mean of the data that went into the model \r\n\r\n The following table shows these ranges for the data set provided.";

            //Excel is writing extranious data behind the tables while the enpi sheet is being created.
            //I spent numerous hours trying to figure out why and in the end decided just to change the font to white for simplicity.
            Excel.Range extraniousData = thisSheet.get_Range("6:12");
            extraniousData.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

            Excel.Range tableBody = thisSheet.get_Range("A6:B6");
            Excel.Range tableHead = thisSheet.get_Range("A6:B6");

            int cols = tableBody.Columns.Count + DS.IndependentVariables.Count;
            tableBody = tableBody.Resize[1, cols];
            tableHead = tableHead.Resize[1, cols];

            for (int i = 3; i <= tableBody.Columns.Count; i++)
            {
                ((Excel.Range)tableBody[i]).Value2 = DS.IndependentVariables[i - 3].ToString();
            }
            tableBody.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

            // write data
            
            //The below changes are made as per the TFS Ticket : 77019
            //tableBody = tableBody.get_Resize(5, cols);
            //tableBody = tableBody.get_Resize(6, cols);
            // The below changes are made based on the SEP changes 
            tableBody = tableBody.get_Resize(7, cols);

            //Excel.Range range1 = thisSheet.get_Range("A7");
            //Excel.Range range1 = thisSheet.get_Range("A8");
            Excel.Range range1 = thisSheet.get_Range("A9");
            range1.Value2 = "Range 1";

            //Excel.Range range2 = thisSheet.get_Range("A9");
            //Excel.Range range2 = thisSheet.get_Range("A10");
            Excel.Range range2 = thisSheet.get_Range("A11");
            range2.Value2 = "Range 2";

            //Excel.Range label0 = thisSheet.get_Range("B7");
            Excel.Range label0 = thisSheet.get_Range("B7");
            //Excel.Range label1 = thisSheet.get_Range("B7");
            //Excel.Range label1 = thisSheet.get_Range("B8");
            Excel.Range label1 = thisSheet.get_Range("B8");
            //Excel.Range label2 = thisSheet.get_Range("B8");
            Excel.Range label2 = thisSheet.get_Range("B9");
            //Excel.Range label2 = thisSheet.get_Range("B9");
            //Excel.Range label3 = thisSheet.get_Range("B9");
            //Excel.Range label3 = thisSheet.get_Range("B10");
            Excel.Range label3 = thisSheet.get_Range("B10");
            //Excel.Range label4 = thisSheet.get_Range("B10");
            //Excel.Range label4 = thisSheet.get_Range("B11");
            Excel.Range label4 = thisSheet.get_Range("B11");
            Excel.Range label5 = thisSheet.get_Range("B12");
            Excel.Range label6 = thisSheet.get_Range("B13");

            //label0.Value2 = "Mean of Model Variable";
            //label1.Value2 = "Minimum of Model Variable";
            //label2.Value2 = "Maximum of Model Variable";
            //label3.Value2 = "Model Avg -3 Std Dev";
            //label4.Value2 = "Model Avg +3 Std Dev";

            label0.Value2 = "SEP mean baseline year value";
            label1.Value2 = "SEP mean report year value";
            label2.Value2 = "Minimum of Model Variable";
            label3.Value2 = "Maximum of Model Variable";
            label4.Value2 = "Model Avg -3 Std Dev";
            label5.Value2 = "Model Avg +3 Std Dev";
            label6.Value2 = "SEP Validation Check";
            //Excel.Range range1Formatting = thisSheet.get_Range("A7").get_Resize(2,cols);
            //Excel.Range range1Formatting = thisSheet.get_Range("A8").get_Resize(2, cols);
            

            thisSheet.Range["A13"].Style.HorizontalAlignment = HorizontalAlign.Right; // SEP Validation "Fail" or "Pass" aligment
            Excel.Range range1Formatting = thisSheet.get_Range("A9").get_Resize(2, cols);
            range1Formatting.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            //Excel.Range range2Formatting = thisSheet.get_Range("A9").get_Resize(2, cols);
            //Excel.Range range2Formatting = thisSheet.get_Range("A10").get_Resize(2, cols);
            Excel.Range range2Formatting = thisSheet.get_Range("A11").get_Resize(3, cols); 

            range2Formatting.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);

            
            //label4.Cells.EntireColumn.AutoFit();
            label6.Cells.EntireColumn.AutoFit();
            ValidationTable = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tableBody, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            ValidationTable.Name = "Validation" + AdjustedData.Name;
            ValidationTable.Range.Cells.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);
            tableHead.get_Offset(0,2).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            
        }

        internal void WriteSourceData()
        {
            //Excel.Range adjRange = thisSheet.get_Range("A12");
            //Excel.Range adjRange = thisSheet.get_Range("A13");
            //Changes done below are due to the new SEP requirements - Suman.
            Excel.Range adjRange = thisSheet.get_Range("A14");
            int sdrows = DS.SourceData.Rows.Count;
            int sdcols = DS.SourceData.Columns.Count;

            // write headers
            for (int c = 0; c < SourceObject.ListColumns.Count; c++)
            {
                string name = SourceObject.ListColumns[c + 1].Name.Replace(((char)13).ToString(), " ").Replace(((char)10).ToString(), "");
                adjRange.get_Resize(1, 1).get_Offset(0, c).Value2 = name;
            }

            // write data
            adjRange = adjRange.get_Resize(1 + sdrows, sdcols);
            adjRange.get_Offset(1,0).get_Resize(sdrows, sdcols).Value2 = SourceData;

            AdjustedData = thisSheet.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, adjRange, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            AdjustedData.Name = "Detail" + AdjustedData.Name;

            ColumnFormatting = new string[sdcols];

            // copy source object formatting to adjusted data
            Excel.Range tmprg = BottomCell().get_Offset(1, 0);
            for (int j = 0; j < sdcols; j++)
            {
                string sourcecoladdr = SourceObject.DataBodyRange.get_Offset(0, j).get_Resize(1,1).get_Address(true,false,Excel.XlReferenceStyle.xlA1, true, System.Type.Missing);
                tmprg.Formula = "=CELL(\"format\"," + sourcecoladdr + ")";
                tmprg.Copy();
                tmprg.PasteSpecial(Excel.XlPasteType.xlPasteValues, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, System.Type.Missing, System.Type.Missing);
                                
                switch (tmprg.Value2.ToString())
                {
                    case "D1":
                        fmt = "dd-mmm-yy";
                        break;
                    case "D2":
                        fmt = "dd-mmm";
                        break;
                    case "D3":
                        fmt = "mmm-yy";
                        break;
                    case "D4":
                        fmt = "m/d/yyyy";
                        break;
                    case "D5":
                        fmt = "mm/dd";
                        break;
                    default:
                        fmt = "General";
                        break;
                }

                ColumnFormatting[j] = fmt;

                tmprg.Clear();
            }

            addSumColumn();
        }


        internal void WriteAdjustedModel()
        {

            string colName = Globals.ThisAddIn.rsc.GetString("AdjustmentMethodColName");
            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[newcol.Index].Name = colName;
            string formula = "=IF([" + Globals.ThisAddIn.rsc.GetString("BaselineYearColName") + "]=[" + Globals.ThisAddIn.rsc.GetString("ModelYearColName") + "],IF([" + EnPIResources.yearColName + "]=[" + Globals.ThisAddIn.rsc.GetString("ModelYearColName") + "]," + "\"" +
            Globals.ThisAddIn.rsc.GetString("adjustmentModel") + "\"" + "," + "\"" + Globals.ThisAddIn.rsc.GetString("adjustmentForecast") + "\"" + "),IF([" + EnPIResources.yearColName + "]=[" + Globals.ThisAddIn.rsc.GetString("ModelYearColName") + "]," + "\"" + Globals.ThisAddIn.rsc.GetString("adjustmentModel") + "\"" +
            ",IF([" + Globals.ThisAddIn.rsc.GetString("LastYearColName") + "]=[" + Globals.ThisAddIn.rsc.GetString("ModelYearColName") + "]," + "\"" + Globals.ThisAddIn.rsc.GetString("adjustmentBackcast") + "\"" +
            ",IF([" + Globals.ThisAddIn.rsc.GetString("BaselineYearColName") + "]=[" + Globals.ThisAddIn.rsc.GetString("ModelYearColName") + "]," + "\"" + Globals.ThisAddIn.rsc.GetString("adjustmentForecast") + "\"" +
            "," + "\"" + Globals.ThisAddIn.rsc.GetString("adjustmentChaining") + "\"" + "))))";

            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = formula.ToString();
            AdjustedData.ListColumns[newcol.Index].DataBodyRange.ColumnWidth = 20;
          

        }

        internal void WriteModelYear(string ModelYear)
        {
            string colName = Globals.ThisAddIn.rsc.GetString("ModelYearColName");

            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

            AdjustedData.ListColumns[newcol.Index].Name = colName;
            if (ModelYear == null)
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = 0;
            else
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = ModelYear;

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WriteBaselineYear(string BaselineYear)
        {
            string colName = Globals.ThisAddIn.rsc.GetString("BaselineYearColName");

            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

            AdjustedData.ListColumns[newcol.Index].Name = colName;
            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = BaselineYear;

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WritePeriodCount()
        {
            string colName = "Period Count";

            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

            AdjustedData.ListColumns[newcol.Index].Name = colName;
            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=IF(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN([Period]))),-1,0,1,1) = OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN([Period]))),0,0,1,1),OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1) + 1,1)";

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WriteBaselineCount()
        {
            string colName = "Baseline Count";

            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

            AdjustedData.ListColumns[newcol.Index].Name = colName;
            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=IF([Period] = [Baseline Year],IF(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN([Period]))),-1,0,1,1) = OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN([Period]))),0,0,1,1),OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1) + 1,1),0)";

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WriteLastYear(string LastYear)
        {
            string colName = Globals.ThisAddIn.rsc.GetString("LastYearColName");

            Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

            AdjustedData.ListColumns[newcol.Index].Name = colName;
            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = LastYear;

            newcol.Range.EntireColumn.Hidden = true;
        }

        internal void WriteNegativeWarning()
        {
            Excel.Range titleRange = thisSheet.get_Range("A2");

            titleRange.Value2 = "Negative Values";
            titleRange.Font.Bold = true;
            titleRange.Font.Size = 14;

            Excel.Range descriptionRange = thisSheet.get_Range("A3:I3");
            descriptionRange.Merge();
            descriptionRange.RowHeight = 69;

            descriptionRange.Value2 = "One or more of the calculated modeled energy consumption values is negative. The negative modeled energy value(s) is shown in yellow. Consider setting the negative modeled energy value as zero.";
            descriptionRange.WrapText = true;
            descriptionRange.VerticalAlignment = Excel.XlVAlign.xlVAlignTop;
        }

        internal void WriteAdjustedData()
        {
            int startct = AdjustedData.ListColumns.Count;

            Globals.ThisAddIn.modeledSourceIndex = new int[DS.EnergySources.Count];
            int count = 0;

            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                if (es.Models.Count > 0)
                {
                    string nm = prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                    string formula = es.BestModel().Formula();
                    string format = null;

                    if (Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name) != null)
                        format = Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name).DataBodyRange.NumberFormat.ToString();

                    Excel.ListColumn newcol = Utilities.ExcelHelpers.AddListColumn(AdjustedData, nm);

                    AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value2 = "=" + formula;
                    AdjustedData.ListColumns[newcol.Index].DataBodyRange.NumberFormat = format ?? "General";

                    Globals.ThisAddIn.modeledSourceIndex[count] = newcol.Index;
                    count++;
                    //NegativeCheck(AdjustedData, newcol);
                }
            }

            if (AdjustedData.ListColumns.Count > startct)
                addSumColumn(prefix());
        }

        internal void WriteEnergySavings(bool regression)
        {
        
                foreach (Utilities.EnergySource es in DS.EnergySources)
                {
                    Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

                    AdjustedData.ListColumns[newcol.Index].Name = "Energy Savings: " + es.Name;
                    if (regression)
                        AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value2 = "=IF([Baseline Year]=[Model Year],["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]-["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "],IF([Period]=[Model Year],IFERROR((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1)+((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1)-OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1))-([" + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]-["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),0),IF([Last Year]=[Model Year],IFERROR((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1)+((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1)-OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1))-([" + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]-["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),0),IF([Baseline Year]=[Model Year],["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]-["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "],IF([Period]<[Model Year],IFERROR((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1)+((OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1)-OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN(["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))),-1,0,1,1))-([" + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]-["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]))),0),["
                            + prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "] - ["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "])))))";
                    else // Modified By Suman : TFS Ticket :68479
                        AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value2 = "=IF([Period]=[Baseline Year],["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "],IF([Period Count] <= MAX([Baseline Count]),INDEX("
                            + AdjustedData.Name
                            + ",[Period Count],COLUMN(["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "])),["
                            + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")
                            + "]))-[" + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "") + "]";

                    //"=IF([Baseline Year]=[Model Year],[" + prefix() + es.Name + "] - [" + es.Name + "],IF([Period]=[Model Year], 1,IF([Last Year]=[Model Year], 1,IF([Baseline Year]=[Model Year],[" + prefix() + es.Name + "] - [" + es.Name + "],IF([Period]<[Model Year],1,[" + prefix() + es.Name + "] - [" + es.Name + "])))))";

                    //"=[" + prefix() + es.Name + "] - [" + es.Name + "]";

                    //per ticket #68382
                    newcol.DataBodyRange.ColumnWidth = 15.71;
                
            }
            
        }

        //Added by Suman TFS Ticket: 66429
        //When the adjustment method is chaining the values of the model year should be left blank and the summation should continue after the model year rows
        internal void WriteCUSUMHidden()
        {
            
                Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

                AdjustedData.ListColumns[newcol.Index].Name = "CUSUMHidden";//Globals.ThisAddIn.rsc.GetString("adjustedModelCUSUMColName");
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value =
                                                            "=IF([Period]=[Model Year],IF(ISNUMBER(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1))=TRUE,OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1),0),["
                                                            + Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName") + "]-["
                                                            + Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName")
                                                            + "]+IF(ISNUMBER(OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1))=TRUE,OFFSET(INDIRECT(ADDRESS(ROW(),COLUMN())),-1,0,1,1),0))";

                AdjustedData.ListColumns[newcol.Index].Range.EntireColumn.Hidden = true;
            
           
        }

        internal void WriteCUSUM()
        {
            
                Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

                AdjustedData.ListColumns[newcol.Index].Name = Globals.ThisAddIn.rsc.GetString("adjustedModelCUSUMColName");
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=IF([Period]=[Model Year],\"\",[CUSUMHidden])";
            
                                                        
        }


        internal void WriteUnitCost()
        {
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
               
                    Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);

                    //AdjustedData.ListColumns[newcol.Index].Name = "Unit Cost: " + es.Name;
                    AdjustedData.ListColumns[newcol.Index].Name = "Unit Cost in $/MMBtu for column: " + es.Name; //TFS Ticket:68849
                
                    for (int i = 0; i < Globals.ThisAddIn.energyCostColumnMatchArray.Length / 2; i++)
                    {
                        if (Globals.ThisAddIn.energyCostColumnMatchArray[i, 0] == es.Name)
                            AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=[" + Globals.ThisAddIn.energyCostColumnMatchArray[i, 1] + "]/[" + es.Name + "]";
                    }

                    //per ticket #68382
                    newcol.DataBodyRange.ColumnWidth = 11.14;
                    //Modified By Suman TFS Ticket:68737
                    if (es.Name.Contains("TOTAL"))
                    {
                        newcol.Range.EntireColumn.Hidden = true;
                    }
            }
        }

        internal void WriteCostSavings()
        {
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);
                //Modified by Suman :TFS Ticket :68851
                AdjustedData.ListColumns[newcol.Index].Name = "Cost Savings ($): " + es.Name;
                //AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=[Unit Cost: " + es.Name +"]*[Energy Savings: " + es.Name + "]";
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value = "=[Unit Cost in $/MMBtu for column: " + es.Name + "]*[Energy Savings: " + es.Name + "]"; //TFS Ticket:68849 

                //per ticket #68382
                //newcol.DataBodyRange.ColumnWidth = 13.71;
                newcol.DataBodyRange.ColumnWidth = 16.83;
            }
        }

        //Added by suman: As per new SEP requirements
        internal void WriteSEPColumnsData()
        {

            try
            {
                WriteSEPEnergySavingsTTM();
                WriteSEPTrailingTwelveMonthEnergyPerformanceIndicator();
                WriteSEPTrailingTwelveMonthEnergySavings();
                WriteSEPTrailingTwelveMonthActualEnergyConsumption();
                WriteSEPTrailingTwelveMonthActualEnergyConsumptionFivePerTarget();
                WriteSEPTrailingTwelveMonthActualEnergyConsumptionTenPerTarget();
                WriteSEPTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget();
                WriteSEPTrailingTwelveMonthEnergySavingsFivePerImprovement();
                WriteSEPTrailingTwelveMonthEnergySavingsTenPerImprovement();
                WriteSEPTrailingTwelveMonthEnergySavingsFifteenPerImprovement();
            }
            catch
            {  }
        }

       

         
        #region SEP Computed columns
        private void WriteSEPEnergySavingsTTM()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepEnergySavingsTTM;
            //string formula = "=IF([Period]<=[Model Year],\"N/A\",";
            //string formula = "=IF([Period]<=[Model Year],\"N/A\",";
            //foreach (Utilities.EnergySource es in DS.EnergySources)
            //{

            //   formula = formula + "[Energy Savings: " + es.Name + "]+";
            //}
            //formula = formula.Remove(formula.Length - 1, 1);
            //formula = formula + ")";
            /*string formula = "=IF([Period]<=[Model Year],\"N/A\",IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\")," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))," +
                               "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))+" +
                               "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])-SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "])))";*/
            string formula = "=IF([Period]<=[Model Year],\"N/A\"," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-" +
                               "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0)))";
            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0"; 
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        private void WriteSEPTrailingTwelveMonthEnergyPerformanceIndicator()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthEnergyPerformanceIndicator;
            /*string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","
                          + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:"
                          + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))/"
                          + "SUM(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.totalAdjValuesColName+"]]:"
                          + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))))"; */


            string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                             "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))/SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))/SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))*" +
                              "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "]))))";
            AdjustedData.ListColumns[col.Index].DataBodyRange.Formula = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "0.00"; //two decimal places
            col.DataBodyRange.ColumnWidth = 15.71;
        }

        private void WriteSEPTrailingTwelveMonthEnergySavings()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthEnergySavings;
           /* string formula = "=IF([Period]<=[Model Year],\"N/A\","
                        + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:"
                        + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-"
                        + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:"
                        + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0)))"; */

            string formula = "=IF([Period]<=[Model Year],\"N/A\",IF(([Adjustment Method]=\""+ EnPIResources.adjustmentForecast +"\"),"+
                              "SUM(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.totalAdjValuesColName+"]]:OFFSET(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.totalAdjValuesColName+"]],-11,0))-SUM(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.unadjustedTotalColName+"]]:OFFSET(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.unadjustedTotalColName+"]],-11,0)),"+
	                           "SUM(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.totalAdjValuesColName+"]]:OFFSET(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.totalAdjValuesColName+"]],-11,0))-SUM(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.unadjustedTotalColName+"]]:OFFSET(" + AdjustedData.Name + "[[#This Row],["+EnPIResources.unadjustedTotalColName+"]],-11,0))+"+
                               "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])-SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "])))";

            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0"; 
            col.DataBodyRange.ColumnWidth = 15.71;
        }

        private void WriteSEPTrailingTwelveMonthActualEnergyConsumption()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption;
            string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","
                        + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]]:"
                        + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.unadjustedTotalColName + "]],-11,0))"
                        + "))";
            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        private void WriteSEPTrailingTwelveMonthActualEnergyConsumptionFivePerTarget()
        {

            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFivePerTarget;
            /*string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","
                           + "(1-0.05)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]))";*/

            /*string formula = "=IF([Period]<[Model Year],\"N/A\","+
                              "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","+
                               "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.05)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]," +
                                "(1-0.05)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]*SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";*/

            string formula = "=IF([Period]<[Model Year],\"N/A\"," +
                             "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                              "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.05)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))," +
                               "(1-0.05)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))*"+
                               "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";

            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        private void WriteSEPTrailingTwelveMonthActualEnergyConsumptionTenPerTarget()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionTenPerTarget;
            /*string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","
                           + "(1-0.1)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]))";*/

            /*string formula = "=IF([Period]<[Model Year],\"N/A\"," +
                              "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                               "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.1)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]," +
                                "(1-0.1)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]*SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";*/


            string formula = "=IF([Period]<[Model Year],\"N/A\"," +
                             "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                              "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.1)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))," +
                               "(1-0.1)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))*" +
                               "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";
            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }

        private void WriteSEPTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget;
           /* string formula = "=IF([Period]<[Model Year],\"N/A\",IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\","
                           + "(1-0.15)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]))";*/

            /*string formula = "=IF([Period]<[Model Year],\"N/A\"," +
                              "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                               "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.15)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]," +
                                "(1-0.15)*[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumption + "]*SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";*/

            string formula = "=IF([Period]<[Model Year],\"N/A\"," +
                             "IF((AND([Period]=[Model Year],[Period] =OFFSET(" + AdjustedData.Name + "[[#This Row],[Period]],1,0))),\"N/A\"," +
                              "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\"),(1-0.15)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))," +
                               "(1-0.15)*SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))*" +
                               "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "]))))";
            
            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        private void WriteSEPTrailingTwelveMonthEnergySavingsFivePerImprovement()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthEnergySavingsFivePerImprovement;
            /*string formula = "=IF([Period]<=[Model Year],\"N/A\","
                  + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:"
                  + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-"
                  + "[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFivePerTarget + "])";*/

            string formula ="=IF([Period]<=[Model Year],\"N/A\","+
                              "IF(([Adjustment Method]=\""+ EnPIResources.adjustmentForecast +"\"),"+
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFivePerTarget + "]," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFivePerTarget + "]+" +
                              "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear +"\",[" + EnPIResources.totalAdjValuesColName + "])))";


            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        private void WriteSEPTrailingTwelveMonthEnergySavingsTenPerImprovement()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthEnergySavingsTenPerImprovement;
           /* string formula = "=IF([Period]<=[Model Year],\"N/A\","
                 + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:"
                 + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-"
                 + "[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionTenPerTarget + "])";*/

            string formula = "=IF([Period]<=[Model Year],\"N/A\"," +
                              "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\")," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionTenPerTarget + "]," +
                              "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionTenPerTarget + "]+" +
                              "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "])))";


            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }

        private void WriteSEPTrailingTwelveMonthEnergySavingsFifteenPerImprovement()
        {
            Excel.ListColumn col = AdjustedData.ListColumns.Add(System.Type.Missing);
            AdjustedData.ListColumns[col.Index].Name = EnPIResources.sepTrailingTwelveMonthEnergySavingsFifteenPerImprovement;
            /*string formula = "=IF([Period]<=[Model Year],\"N/A\","
                 + "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:"
                 + "OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-"
                 + "[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget + "])";*/
            string formula = "=IF([Period]<=[Model Year],\"N/A\"," +
                             "IF(([Adjustment Method]=\"" + EnPIResources.adjustmentForecast + "\")," +
                             "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget + "]," +
                             "SUM(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]]:OFFSET(" + AdjustedData.Name + "[[#This Row],[" + EnPIResources.totalAdjValuesColName + "]],-11,0))-[" + EnPIResources.sepTrailingTwelveMonthActualEnergyConsumptionFifteenPerTarget + "]+" +
                             "SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.unadjustedTotalColName + "])/SUMIF([Period],\"" + Globals.ThisAddIn.BaselineYear + "\",[" + EnPIResources.totalAdjValuesColName + "])))";

            AdjustedData.ListColumns[col.Index].DataBodyRange.Value = formula;
            AdjustedData.ListColumns[col.Index].DataBodyRange.HorizontalAlignment = XlHAlign.xlHAlignRight;
            AdjustedData.ListColumns[col.Index].DataBodyRange.NumberFormat = "##,##0";
            col.DataBodyRange.ColumnWidth = 15.71;
        }
        #endregion

        internal void addSumColumn(params string[] prepend)
        {

            try
            {
                string colName = Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName");
                StringBuilder formula = new StringBuilder("=");
                string format = "General";
                string prefix = "";
                if (prepend.Length > 0)
                {
                    prefix = prepend[0];
                    colName = Globals.ThisAddIn.rsc.GetString("totalAdjValuesColName");
                }

                object[,] thistmp = (object[,])AdjustedData.HeaderRowRange.Value2;

                foreach (Utilities.EnergySource es in DS.EnergySources)
                {
                    if (es.Name != Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName"))
                    {
                        if (formula.Length > 1) formula.Append("+");
                        formula.Append(Utilities.ExcelHelpers.CreateValidFormulaName(prefix + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "")));

                        if (Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name) != null)
                            format = Utilities.ExcelHelpers.GetListColumn(SourceObject, es.Name).DataBodyRange.NumberFormat.ToString();
                    }
                }
                Excel.ListColumn newcol = AdjustedData.ListColumns.Add(System.Type.Missing);// Utilities.ExcelHelpers.AddListColumn(AdjustedData, colName);

                AdjustedData.ListColumns[newcol.Index].Name = colName;
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.Value2 = formula.ToString();
                AdjustedData.ListColumns[newcol.Index].DataBodyRange.NumberFormat = format ?? "General";
            }
            catch (Exception ex)
            {

            }
 
        }

        #region //computed columns
        

        internal string prefix()
        {
            string prefix = Globals.ThisAddIn.rsc.GetString("prefixAdjusted") ?? "Adj.";
            return prefix;
        }

        #endregion

        #region //Formatting
        internal void AddConditionalFormatting(Excel.ListObject LO)
        {
            string[,] formatArray = new string[DS.IndependentVariables.Count + 1, LO.ListRows.Count];
            int[] breakIndex = new int[DS.Years.Count - 1];
            int varCount = 0;
            int breakCount = 0;
            int modelYearIndex = 0;
            int reportYearIndex = 0;
            int baselineYearIndex = 0;
            int yearColumnIndex = 0;
            bool backcast = false;

            foreach (Excel.ListColumn LC in LO.ListColumns)
            {
                if(LC.Name.ToLower().Equals(EnPIResources.yearColName.ToLower()))
                {
                    yearColumnIndex = LC.Index;
                    //capture column and put into array
                    for (int i = 2; i < LO.ListRows.Count + 2; i++)
                    {
                        if (((Excel.Range)LC.Range[i, 1]).Value2.ToString().Equals(DS.ModelYear))
                            modelYearIndex = i;
                        if (((Excel.Range)LC.Range[i, 1]).Value2.ToString().Equals(DS.ReportYear))
                            reportYearIndex = i;
                        if (((Excel.Range)LC.Range[i, 1]).Value2.ToString().Equals(DS.BaselineYear))
                            baselineYearIndex = i;
                        formatArray[0,i-2] = ((Excel.Range)LC.Range[i, 1]).Value2.ToString();
                    }
                
                    string startText = formatArray[0, 0];

                    //LO.ListRows.Count - 1
                    // - 1 removed for fix for ticket #68384
                    for (int i = 0; i < LO.ListRows.Count; i++)
                    {
                        if (!formatArray[0, i].Equals(startText))
                        {
                            startText = formatArray[0, i];
                            breakIndex[breakCount] = i ;
                            breakCount++;
                        }
                    }
                }

                if (DS.IndependentVariables.Contains(LC.Name))
                {
                    varCount++;
                    //capture column and put into array
                    for (int i = 2; i < LO.ListRows.Count + 2; i++)
                    {
                        try
                        {
                           formatArray[varCount, i - 2] = ((Excel.Range)LC.Range[i, 1]).Value2.ToString();
                        }
                        catch(Exception ex)
                        {
                           // Do Nothing, the reason for this try catch is to not to stop the tool from generating the other sheets.
                            //TFS Ticket : 69719
                        }
                    }                    
                }

            }

            double[,] yearTotals = new double[DS.Years.Count, DS.IndependentVariables.Count];
            double[,] yearAverages = new double[DS.Years.Count, DS.IndependentVariables.Count];
            double[,] yearMin = new double[DS.Years.Count, DS.IndependentVariables.Count];
            double[,] yearMax = new double[DS.Years.Count, DS.IndependentVariables.Count];

            //total the variables and add to array
            int k = 0;
            double total;
            double min;
            double max;
            for (int c = 0; c < DS.IndependentVariables.Count; c++)
            {
                total = 0;
                k = 0;
                min = Convert.ToDouble(formatArray[c + 1, 0]);
                max = 0;

                for(int i = 0; i < LO.ListRows.Count; i++)
                {
                    total += Convert.ToDouble(formatArray[c + 1, i]);

                    if (Convert.ToDouble(formatArray[c + 1, i]) > max)
                        max = Convert.ToDouble(formatArray[c + 1, i]);

                    if (Convert.ToDouble(formatArray[c + 1, i]) < min)
                        min = Convert.ToDouble(formatArray[c + 1, i]);

                    if (k.Equals(breakIndex.Length))
                        k--;
                    if (i.Equals(breakIndex[k] - 1))
                    {
                        yearTotals[k, c] = total;
                        yearMin[k, c] = min;
                        yearMax[k, c] = max;
                        total = 0;
                        //min = Convert.ToDouble(formatArray[c + 1, i]);
                        //TFS - 79523 
                        min = Convert.ToDouble(formatArray[c + 1, breakIndex[k]]); //Increment to next number in the array
                        max = 0;
                        k++;
                    }
                    //For the last reporting year
                    if (i.Equals(LO.ListRows.Count - 1))
                    {
                        yearTotals[k + 1, c] = total;
                        yearMin[k + 1, c] = min;
                        yearMax[k + 1, c] = max;
                        total = 0;
                        min = Convert.ToDouble(formatArray[c + 1, i]);
                        max = 0;
                        k = 0;
                    }
                }
            }

            for (int i = 0; i < breakIndex.Length; i++)
            {
                //for forcast and chaining
                if (modelYearIndex.Equals(breakIndex[i] + 1))
                    modelYearIndex = i;
                if (reportYearIndex.Equals(breakIndex[i] + 1))
                    reportYearIndex = i;
                if (baselineYearIndex.Equals(breakIndex[i] + 1))
                    baselineYearIndex = i;
                //for backcast
                if (modelYearIndex.Equals(LO.ListRows.Count + 1))
                {
                    modelYearIndex = breakIndex.Length - 1;
                    backcast = true;
                }
                if (reportYearIndex.Equals(LO.ListRows.Count + 1))
                {
                    //reportYearIndex = breakIndex.Length - 1;
                    reportYearIndex = breakIndex.Length;
                    
                }


            }

            double[] yearRowCount = new double[breakCount + 1];

            for (int i = 0; i < breakCount + 1; i++)
            {
                int newCount = 0;
                if (i.Equals(0))
                    newCount = breakIndex[i];
                else if (!i.Equals(breakCount))
                    newCount = (breakIndex[i] - breakIndex[i - 1]);
                else
                    newCount = LO.ListRows.Count - breakIndex[i - 1];

                yearRowCount[i] = newCount;
            }

            //average the variables and add to array
            for(int i = 0; i < DS.Years.Count; i++)
            {
                for (int c = 0; c < DS.IndependentVariables.Count; c++)
                {
                    yearAverages[i, c] = (yearTotals[i, c] / yearRowCount[i]);
                }
            }

            bool[,] yearFlag = new bool[DS.Years.Count, DS.IndependentVariables.Count];
            //chanegd to max for backbast with more data points than the model year. BJV - ticket #68756
            //double[,] modelValues = new double[DS.IndependentVariables.Count, Convert.ToInt32(yearRowCount[modelYearIndex])];
            double[,] modelValues = new double[DS.IndependentVariables.Count, Convert.ToInt32(yearRowCount.Max())];
            int count;

            for (int i = 0; i < DS.IndependentVariables.Count; i++)
            {
                count = 0;

                for (int a = 0; a < LO.ListRows.Count; a++)
                {
                    if (a >= (breakIndex[modelYearIndex] - yearRowCount[modelYearIndex]) && a < breakIndex[modelYearIndex] && !backcast)
                    {
                        modelValues[i, count] = Convert.ToDouble(formatArray[i + 1, a]);
                        count++;
                    }
                    else if (backcast && a >= (breakIndex[modelYearIndex]) && a < (breakIndex[modelYearIndex] + yearRowCount[modelYearIndex + 1]))
                    {
                        modelValues[i, count] = Convert.ToDouble(formatArray[i + 1, a]);
                        count++;
                    }
                }

                for (int c = 0; c < DS.Years.Count; c++)
                {
                    yearFlag[c, i] = validationCheck(DS.IndependentVariables[i],yearTotals[c, i], yearAverages[c, i],yearAverages[reportYearIndex,i],yearAverages[baselineYearIndex,i], yearMin[modelYearIndex, i], yearMax[modelYearIndex, i], yearMin[modelYearIndex + 1, i],
                                                   yearMax[modelYearIndex + 1, i],i, c, modelYearIndex, formatArray, modelValues, 
                                                   backcast ? Convert.ToInt32(yearRowCount[modelYearIndex + 1]) : Convert.ToInt32(yearRowCount[modelYearIndex]), backcast);
                }
            }


            //highlight years that cause validation and raise error message to user

            for (int i = 0; i < DS.IndependentVariables.Count; i++)
            {
                for (int c = 0; c < DS.Years.Count; c++)
                {
                    if (yearFlag[c, i].Equals(true))
                    {
                        foreach (Excel.ListColumn LC in LO.ListColumns)
                        {
                            if (DS.IndependentVariables[i].Equals(LC.Name))
                            {
                                foreach (Excel.ListRow LR in LO.ListRows)
                                {
                                    if(DS.Years[c].Equals(((Excel.Range)LR.Range[1,yearColumnIndex]).Value2.ToString()))
                                    {
                                        ((Excel.Range)LR.Range[1, LC.Index]).Font.ColorIndex = 3;
                                        //Globals.ThisAddIn.hasSEPValidationError = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            Globals.ThisAddIn.lstSEPValidationValues = lstWarningValidationValues;

            foreach (SEPValidationValues oSEPVal in lstWarningValidationValues)
            {
                if (oSEPVal.YearFlag == true && !oSEPVal.SEPValidationCheck.Contains("Not"))
                {
                    string minmaxModel = string.Empty;
                    string plusminusSTDDev = string.Empty;
                    string avgCompareYear = string.Empty;

                    minmaxModel = Convert.ToInt32(oSEPVal.MinModel).ToString() + " , " +
                                  Convert.ToInt32(oSEPVal.MaxModel).ToString();
                    plusminusSTDDev = Convert.ToInt32(oSEPVal.Minus3DevVal).ToString() + " , " +
                                      Convert.ToInt32(oSEPVal.Plus3DevVal).ToString();

                    //forecast
                    if (!backcast)
                    {
                        if (!((oSEPVal.MinModel < oSEPVal.AvgReportYr && oSEPVal.AvgReportYr < oSEPVal.MaxModel) ||
                              (oSEPVal.Minus3DevVal < oSEPVal.AvgReportYr &&
                               oSEPVal.AvgReportYr < oSEPVal.Plus3DevVal)) &&
                            ((oSEPVal.MinModel < oSEPVal.AvgBaselineYr && oSEPVal.AvgBaselineYr < oSEPVal.MaxModel) ||
                             (oSEPVal.Minus3DevVal < oSEPVal.AvgBaselineYr &&
                              oSEPVal.AvgBaselineYr < oSEPVal.Plus3DevVal)))

                        {
                            avgCompareYear = Convert.ToInt32(oSEPVal.AvgReportYr).ToString();
                        }

                        //chaining

                        else if (!((oSEPVal.MinModel < oSEPVal.AvgReportYr && oSEPVal.AvgReportYr < oSEPVal.MaxModel) ||
                                   (oSEPVal.Minus3DevVal < oSEPVal.AvgReportYr &&
                                    oSEPVal.AvgReportYr < oSEPVal.Plus3DevVal)) &&
                                 !((oSEPVal.MinModel < oSEPVal.AvgBaselineYr &&
                                    oSEPVal.AvgBaselineYr < oSEPVal.MaxModel) ||
                                   (oSEPVal.Minus3DevVal < oSEPVal.AvgBaselineYr &&
                                    oSEPVal.AvgBaselineYr < oSEPVal.Plus3DevVal)))
                        {
                            avgCompareYear = Convert.ToInt32(oSEPVal.AvgReportYr).ToString() + " and " +
                                             Convert.ToInt32(oSEPVal.AvgBaselineYr).ToString();

                        }

                        else if (((oSEPVal.MinModel < oSEPVal.AvgReportYr && oSEPVal.AvgReportYr < oSEPVal.MaxModel) ||
                                  (oSEPVal.Minus3DevVal < oSEPVal.AvgReportYr &&
                                   oSEPVal.AvgReportYr < oSEPVal.Plus3DevVal)) &&
                                 !((oSEPVal.MinModel < oSEPVal.AvgBaselineYr && oSEPVal.AvgBaselineYr < oSEPVal.MaxModel) ||
                                   (oSEPVal.Minus3DevVal < oSEPVal.AvgBaselineYr &&
                                    oSEPVal.AvgBaselineYr < oSEPVal.Plus3DevVal)))
                        {
                            avgCompareYear = Convert.ToInt32(oSEPVal.AvgBaselineYr).ToString();

                        }
                    }
                    // backcast
                    if (backcast) { 
                         if (((oSEPVal.MinModel < oSEPVal.AvgReportYr && oSEPVal.AvgReportYr < oSEPVal.MaxModel) ||
                                  (oSEPVal.Minus3DevVal < oSEPVal.AvgReportYr &&
                                   oSEPVal.AvgReportYr < oSEPVal.Plus3DevVal)) &&
                                 !((oSEPVal.MinModel < oSEPVal.AvgBaselineYr && oSEPVal.AvgBaselineYr < oSEPVal.MaxModel) ||
                                   (oSEPVal.Minus3DevVal < oSEPVal.AvgBaselineYr &&
                                    oSEPVal.AvgBaselineYr < oSEPVal.Plus3DevVal)))
                        {
                            avgCompareYear = Convert.ToInt32(oSEPVal.AvgBaselineYr).ToString();

                        }
                    }



                    Globals.ThisAddIn.sepValidationWarningMsg = "Warning: The mean of the cells highlighted in red on the “Model Data” sheet (" + avgCompareYear + ") are out of the allowable range of the model year values (min, max = " + minmaxModel + " ; mean +/-3 std. dev. = " + plusminusSTDDev + " ). Meaning, the model cannot be used to predict the energy consumption for the time period shown in red if the variables shown in red are included in the model. It is recommended to select an alternative model which meets the R-squared and p-value requirements and does not include the variable shown in red in the model. If an alternative model cannot be selected with the current model year, try selecting an alternative model year. For more information, see the SEP Measurement and Verification Protocol.";
                    Globals.ThisAddIn.hasSEPValidationError = true;
                    
                }
            }



            //Added by suman SEP Changes.
            if (Globals.ThisAddIn.hasSEPValidationError)
            {
                //text changed per ticket #66441
                //Warnings.Add("Warning: The cells highlighted in red are out of the allowable range of the model year values. Meaning, the model cannot be used to predict the energy consumption for the time period shown in red if the variables shown in red are included in the model. It is recommended to select an alternative model which meets the R-squared and p-value requirements and does not include the variable shown in the model. If an alternative model cannot be selected with the current model year, try selecting an alternative model year. For more information, see the SEP Measurement and Verification Protocol.");
                //Warnings.Add("Warning: The cells highlighted in red are out of the allowable range of the model year values. If the model is being evaluated during a period where it is not valid, please use a different model adjustment application method.");
                // string minmaxModel = string.Empty;
                // string plusminusSTDDev = string.Empty;
                // string avgReportYear = string.Empty;
                //// for (int cnt = 0; cnt < lstWarningValidationValues.Count; cnt++)
                // if(lstWarningValidationValues.Count >0)
                // {
                //     minmaxModel =  Convert.ToInt32(lstWarningValidationValues[0].MinModel).ToString() + " , " + Convert.ToInt32(lstWarningValidationValues[0].MaxModel).ToString();
                //     plusminusSTDDev =  Convert.ToInt32(lstWarningValidationValues[0].Minus3DevVal).ToString() + " , " + Convert.ToInt32(lstWarningValidationValues[0].Plus3DevVal).ToString();
                //     avgReportYear =  Convert.ToInt32(lstWarningValidationValues[0].AvgReportYr).ToString();
                // }
                // Globals.ThisAddIn.sepValidationWarningMsg = "Warning: The mean of the cells highlighted in red on the “Model Data” sheet (" + avgReportYear + ") are out of the allowable range of the model year values (min, max = " + minmaxModel + " ; mean +/-3 std. dev. = " + plusminusSTDDev + " ). Meaning, the model cannot be used to predict the energy consumption for the time period shown in red if the variables shown in red are included in the model. It is recommended to select an alternative model which meets the R-squared and p-value requirements and does not include the variable shown in red in the model. If an alternative model cannot be selected with the current model year, try selecting an alternative model year. For more information, see the SEP Measurement and Verification Protocol.";
                Warnings.Add(Globals.ThisAddIn.sepValidationWarningMsg);
            }
        }

        private bool validationCheck(string independentVariable,double total, double avrg, double avgReportYr,double avgBaselineYr, double minModel, double maxModel, double minModelBackcast, double maxModelBackcast, int varIndex, int yearIndex, int modelYearIndex, string[,] dataSet, double[,] modelVarData, int modelPoints, bool backcast)
        {

            double[] forStdDev = new double[modelPoints];
            bool retVal = false;
            double plus3DevVal=0, minus3DevVal=0;
            string sepValidationCheck = string.Empty;
            for(int i = 0; i < modelPoints; i++)
            {
                forStdDev[i] = modelVarData[varIndex, i];
            }

            if (avrg > (forStdDev.Average() + 3 * ArrayStdDev(forStdDev)))
                retVal = true;
            else if (avrg < (forStdDev.Average() - 3 * ArrayStdDev(forStdDev)))
                retVal = true;

            if (!backcast)
            {
                sepValidationCheck = GetSEPValidationCheckValue(independentVariable, avrg, avgReportYr, avgBaselineYr,
                    minModel, maxModel, (forStdDev.Average() - 3 * ArrayStdDev(forStdDev)),
                    (forStdDev.Average() + 3 * ArrayStdDev(forStdDev)), backcast);
            }
            else
            {
                sepValidationCheck = GetSEPValidationCheckValue(independentVariable, avrg, avgReportYr, avgBaselineYr,
                    minModelBackcast, maxModelBackcast, (forStdDev.Average() - 3 * ArrayStdDev(forStdDev)),
                    (forStdDev.Average() + 3 * ArrayStdDev(forStdDev)), backcast);
            }


            if (yearIndex.Equals(modelYearIndex) && !backcast)
            {
                //add values to the validation table on the sheet
                //These changes are made as per the TFS Ticket: 77019
                //Changes below are done as per the new SEP Changes requirements by Suman. 5/31/2016
                Excel.Range meanOfModel = thisSheet.get_Range("C7").get_Offset(0, varIndex);
                meanOfModel.Value2 = avgBaselineYr.ToString();
                meanOfModel.NumberFormat = "#,###0";
                meanOfModel.EntireRow.Hidden = (Globals.ThisAddIn.AdjustmentMethod =="Chaining"? false: true);

                Excel.Range meanOfModel2 = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                meanOfModel2.Value2 = avgReportYr.ToString();
                meanOfModel2.NumberFormat = "#,###0";
                
                //Excel.Range minOfModel = thisSheet.get_Range("C7").get_Offset(0, varIndex);
                //Excel.Range minOfModel = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                Excel.Range minOfModel = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                minOfModel.Value2 = minModel.ToString();
                minOfModel.NumberFormat = "#,###0";
                //Excel.Range maxOfModel = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                //Excel.Range maxOfModel = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                Excel.Range maxOfModel = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                maxOfModel.Value2 = maxModel.ToString();
                maxOfModel.NumberFormat = "#,###0";

                //Excel.Range minus3Dev = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                //Excel.Range minus3Dev = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                Excel.Range minus3Dev = thisSheet.get_Range("C11").get_Offset(0, varIndex);
                minus3DevVal=(forStdDev.Average() - 3 * ArrayStdDev(forStdDev));
                minus3Dev.Value2 = minus3DevVal.ToString();
                minus3Dev.NumberFormat = "#,###0";
                //Excel.Range plus3Dev = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                //Excel.Range plus3Dev = thisSheet.get_Range("C11").get_Offset(0, varIndex);
                Excel.Range plus3Dev = thisSheet.get_Range("C12").get_Offset(0, varIndex);
                plus3DevVal =(forStdDev.Average() + 3 * ArrayStdDev(forStdDev));
                plus3Dev.Value2 = plus3DevVal.ToString();
                plus3Dev.NumberFormat = "#,###0";

                Excel.Range sepValidationChk = thisSheet.get_Range("C13").get_Offset(0, varIndex);
               
                sepValidationChk.Value2 = sepValidationCheck;
                sepValidationChk.NumberFormat = "General";

                //return false;
            }
            else if (yearIndex.Equals(modelYearIndex + 1) && backcast)
            {
                //add values to the validation table on the sheet
                Excel.Range meanOfModel = thisSheet.get_Range("C7").get_Offset(0, varIndex);
                meanOfModel.Value2 = avgBaselineYr.ToString();
                meanOfModel.NumberFormat = "#,###0";
                
                Excel.Range meanOfModel2 = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                meanOfModel2.Value2 = avgReportYr.ToString();
                meanOfModel2.NumberFormat = "#,###0";
                meanOfModel2.EntireRow.Hidden = true;
                //Excel.Range minOfModel = thisSheet.get_Range("C7").get_Offset(0, varIndex);
                //Excel.Range minOfModel = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                Excel.Range minOfModel = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                minOfModel.Value2 = minModelBackcast.ToString();
                minOfModel.NumberFormat = "#,###0";
                //Excel.Range maxOfModel = thisSheet.get_Range("C8").get_Offset(0, varIndex);
                //Excel.Range maxOfModel = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                Excel.Range maxOfModel = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                maxOfModel.Value2 = maxModelBackcast.ToString();
                maxOfModel.NumberFormat = "#,###0";

                //Excel.Range minus3Dev = thisSheet.get_Range("C9").get_Offset(0, varIndex);
                //Excel.Range minus3Dev = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                Excel.Range minus3Dev = thisSheet.get_Range("C11").get_Offset(0, varIndex);
                minus3DevVal =(forStdDev.Average() - 3 * ArrayStdDev(forStdDev));
                minus3Dev.Value2 = minus3DevVal.ToString();
                minus3Dev.NumberFormat = "#,###0";
                //Excel.Range plus3Dev = thisSheet.get_Range("C10").get_Offset(0, varIndex);
                //Excel.Range plus3Dev = thisSheet.get_Range("C11").get_Offset(0, varIndex);
                Excel.Range plus3Dev = thisSheet.get_Range("C12").get_Offset(0, varIndex);
                plus3DevVal = (forStdDev.Average() + 3 * ArrayStdDev(forStdDev));
                plus3Dev.Value2 = plus3DevVal.ToString();
                plus3Dev.NumberFormat = "#,###0";

                Excel.Range sepValidationChk = thisSheet.get_Range("C13").get_Offset(0, varIndex);
               // sepValidationCheck = GetSEPValidationCheckValue(independentVariable, avrg, avgReportYr, minModel, maxModel, minus3DevVal, plus3DevVal);
                sepValidationChk.Value2 = sepValidationCheck;
                sepValidationChk.NumberFormat = "General";

                //return false;
            }

            //else if (avrg < minModel || avrg > maxModel)
            //{
            //check against 3 * std dev

            //else
            //  return false;
            //}
            //else
            //  return false;
            //if (retVal == true)
            //{
            if (!backcast)
            {
                lstWarningValidationValues.Add(new SEPValidationValues(independentVariable, varIndex, yearIndex, minModel, avgReportYr, avgBaselineYr, maxModel, (forStdDev.Average() - 3 * ArrayStdDev(forStdDev)), (forStdDev.Average() + 3 * ArrayStdDev(forStdDev)), sepValidationCheck, retVal));
            }
            else
            {
                lstWarningValidationValues.Add(new SEPValidationValues(independentVariable, varIndex, yearIndex, minModelBackcast, avgReportYr, avgBaselineYr, maxModelBackcast, (forStdDev.Average() - 3 * ArrayStdDev(forStdDev)), (forStdDev.Average() + 3 * ArrayStdDev(forStdDev)), sepValidationCheck, retVal));
            }
            //}


            return retVal;
        }

        private string GetSEPValidationCheckValue(string independentVariable, double avrg, double avgReportYr, double avgBaselineYr,double minModel, double maxModel, double minus3DevVal, double plus3DevVal,bool backcast)
        {
            string retVal = string.Empty;
            bool includedInModel = false;
            
            foreach (Utilities.EnergySource es in DS.EnergySources)
            {
                if (es.Models.Count > 0)
                {
                    string nm = prefix() + es.Name.Replace(((char)13).ToString(), "").Replace(((char)10).ToString(), "");
                    string formula = es.BestModel().Formula();
                    if (formula.Contains(independentVariable))
                    {
                        includedInModel = true;
                    }
                }
            }

            if (includedInModel == true)
            {
                
                    if (((minModel < avgReportYr && avgReportYr < maxModel) || (minus3DevVal < avgReportYr && avgReportYr < plus3DevVal))
                  &&  ((minModel < avgBaselineYr && avgBaselineYr < maxModel) || (minus3DevVal < avgBaselineYr && avgBaselineYr < plus3DevVal)))
                    {
                        retVal = "Pass";
                    }
                    else
                    {
                        retVal = "Fail";
                    }
               
               
            }
            else
            {
                retVal = "Not Included In Model"; //TODO:Read this hard coded statement from resource files.
            }



            return retVal;

        }

        

        static internal double ArrayStdDev(double[] arrayVals)
        {
            double sum = 0;
            double sumSquare = 0;
            double value;
            int count = 0;

            for (int i = arrayVals.GetLowerBound(0); i <= arrayVals.GetUpperBound(0); i++)
            {
                if (double.TryParse(arrayVals[i].ToString(), out value))
                {
                    count += 1;
                    sum += value;
                    sumSquare += value * value;
                }
            }


            double stdev = 0;

            if (count != 0) stdev = Math.Sqrt((sumSquare - (sum * sum / count)) / (count - 1));

            return stdev;
        }

        internal void FormatHeaderRow(Excel.ListObject LO)
        {
            Excel.Range header = LO.HeaderRowRange;

            header.Cells.WrapText = true;
            header.Cells.HorizontalAlignment = Excel.XlVAlign.xlVAlignCenter;
            header.Cells.UseStandardHeight = true;
            header.Cells.EntireRow.AutoFit();
            header.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
            header.Font.Bold = true;
            header.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Black);

        }

        internal void FormatAdjustedData()
        {
            AdjustedData.ShowTotals = false;

            if (DS.ModelData != null)
            {
                // highlight model year
               
                int yearColIndex = Utilities.ExcelHelpers.GetListColumn(AdjustedData, EnPIResources.yearColName).Index;

                for (int i = 1; i <= AdjustedData.DataBodyRange.Rows.Count; i++)
                {
                    AdjustedData.Range.get_Offset(i, 1).NumberFormat = fmt;
                    AdjustedData.Range.get_Offset(i, 1).NumberFormat = "#,###0";
                    AdjustedData.Range.get_Offset(i, yearColIndex - 1).NumberFormat = "####";
                    AdjustedData.Range.get_Offset(i, yearColIndex).NumberFormat = "#,###0";

                    string year = AdjustedData.Range.get_Offset(i, yearColIndex - 1).get_Resize(1, 1).Value2.ToString();

                    //Highlight the model year
                    if (year == DS.ModelYear)
                    {
                        AdjustedData.Range.get_Offset(i, 0).get_Resize(1, AdjustedData.ListColumns.Count).Font.Color = 0xAA0000;// "Input";
                        AdjustedData.Range.get_Offset(i, 0).get_Resize(1, AdjustedData.ListColumns.Count).Interior.Color = 0xD8E4BC; //System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.GreenYellow);//
                    }
                }

                for (int j = 0; j < ColumnFormatting.Length; j++)
                {
                    if (ColumnFormatting[j] != "General" )
                        AdjustedData.ListColumns[j + 1].DataBodyRange.NumberFormat = ColumnFormatting[j];
                    if (AdjustedData.ListColumns[1].Name != EnPIResources.dateColName)
                        AdjustedData.ListColumns[1].DataBodyRange.NumberFormat = "#,###0";
                }
                
                if(Globals.ThisAddIn.fromRegression)//only if using regression, use actual can't be formatted.
                    AddConditionalFormatting(AdjustedData);
            }

            FormatHeaderRow(AdjustedData);
        }
        #endregion

    }

   public class SEPValidationValues
    {
        
        internal int VarIndex { get; set; }
        internal int YearIndex { get; set; }
        internal double MinModel { get; set; }
        internal double MaxModel { get; set; }
        internal double Minus3DevVal { get; set; }
        internal double Plus3DevVal { get; set; }
        internal double AvgReportYr { get; set; }
        internal string SEPValidationCheck { get; set; }
        internal string IndependentVariable { get; set; }
        internal bool YearFlag { get; set; }
        internal double AvgBaselineYr { get; set; }
       internal SEPValidationValues(string independentVariable,int varIndex, int yearIndex, double minModel, double avgReportYr, double avgBaselineYr,
                                    double maxModel, double minus3DevVal, double plus3DevVal,string validationCheck,
                                    bool yearFlag)
        {
            IndependentVariable = independentVariable;
            VarIndex = varIndex;
            YearIndex = yearIndex;
            MinModel = minModel;
            MaxModel = maxModel;
            Minus3DevVal = minus3DevVal;
            Plus3DevVal = plus3DevVal;
            AvgReportYr = avgReportYr;
            SEPValidationCheck = validationCheck;
            YearFlag = yearFlag;
            AvgBaselineYr = avgBaselineYr;
        }
    }
}
