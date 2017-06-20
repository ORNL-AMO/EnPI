using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;


namespace AMO.EnPI.AddIn
{
    public class ModelSheetCollection : System.Collections.CollectionBase
    {
        
        public ModelSheet Add(Utilities.EnergySource aSource)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveWorkbook.Worksheets.Add
                (System.Type.Missing, Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.get_Item(Globals.ThisAddIn.Application.ActiveWorkbook.Sheets.Count), 1, Excel.XlSheetType.xlWorksheet);
            thisSheet.CustomProperties.Add("SheetGUID", System.Guid.NewGuid().ToString());
            
            thisSheet.Name = Utilities.ExcelHelpers.CreateValidWorksheetName(Globals.ThisAddIn.Application.ActiveWorkbook, aSource.Name, Globals.ThisAddIn.groupSheetCollection.regressionIteration);
            thisSheet.Tab.Color = 0x50CC11;
            thisSheet.Visible = Excel.XlSheetVisibility.xlSheetHidden;

            string sheetString = thisSheet.ToString();

            GroupSheet GS = new GroupSheet(thisSheet, false, false, thisSheet.Name);
            Globals.ThisAddIn.groupSheetCollection.Add(GS);

            ModelSheet aSheet = new ModelSheet(thisSheet);
            aSheet.Source = aSource;

            return aSheet;
        }
    }

    public class ModelSheet 
    {
        public Utilities.EnergySource Source;
        public Excel.Worksheet WS;
        public Utilities.ModelCollection mCol;

        public ModelSheet(Excel.Worksheet sheet)
        {
            WS = sheet;
        }

        private Excel.Range BottomCell()
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(WS, 0).ToString();

            return (Excel.Range)WS.get_Range(addr, System.Type.Missing);
        }

        public void Populate()
        {
            WriteBestModel();
            WriteWarning();
            WriteDescription();

            WriteResultsTable(5, true);
            newChart(5);

            WriteResultsTable(5, false);

        }

        private void WriteBestModel()
        {
            WS.Range["A1"].EntireRow.Hidden = true;

            Utilities.Model bestmodel = Source.BestModel();
            Globals.ThisAddIn.SelectedSourcesBestModelFormulas.Add(new KeyValuePair<string,string>(Source.Name,bestmodel.Formula()));
            
            Excel.Range target = BottomCell().get_Offset(1, 0).get_Resize(1, 5);
            target.Merge(true);

            target.Value2 = Globals.ThisAddIn.rsc.GetString("bestModel") + bestmodel.ModelNumber.ToString();

            if (!bestmodel.Valid())
            {
                target = BottomCell().get_Offset(1, 0).get_Resize(1, 5);
                target.Merge(true);
                target.Value2 = Globals.ThisAddIn.rsc.GetString("noValidModel");
            }

            target.EntireRow.Hidden = true;
        }

        private void WriteWarning()
        {
            int t;
            if (!int.TryParse(Globals.ThisAddIn.rsc.GetString("warningThreshold"), out t)) t = 10;

            if (Source.Ys.Count() < t)
            {
                Excel.Range target = BottomCell().get_Offset(1, 0).get_Resize(1, 7);
                target.Merge(true);
                target.Value2 = Globals.ThisAddIn.rsc.GetString("warningSmallModel") ;
                target.Font.ColorIndex = 3;

            }
        }

        private void WriteDescription()
        {
            Excel.Range title = BottomCell().get_Offset(1, 0).get_Resize(1, 7);
            title.Merge(true);

            ((Excel.Range)title[1, 1]).Value2 = WS.Name + " Models";
            title.Font.Bold = true;
            title.WrapText = true;

            Excel.Range target = BottomCell().get_Offset(1, 0).get_Resize(1, 7);
            target.Merge(true);

            ((Excel.Range)target[1, 1]).Value2 = "The table below shows all possible models for " + WS.Name + " consumption. The model highlighted in green in the table below is the model with the highest Adjusted R2 value. If \"true\" is shown in column B, the model is designated as valid. A model is considered valid if the model p-value is less than 0.10. The model highlighted in green is used to calculate the adjusted data on the EnPI Results, SEnPI Results, and Adjusted Data tabs. If the model is switched, the corresponding data will be updated with the model selected. The models can be switched using the \"Change Models\" icon in the top navigation.";
            target.WrapText = true;
            target.EntireRow.RowHeight = 105;
            target.Name = "test";

            bool noVaildModels = true;

            foreach (Utilities.Model mdl in Source.Models)
            {
                if (mdl.Valid())
                    noVaildModels = false;
            }

            if (noVaildModels)
            {
                Excel.Range text2 = BottomCell().get_Offset(2, 0).get_Resize(1, 7);
                text2.Merge(true);
                ((Excel.Range)text2[1, 1]).Value2 = "None of the models produced have a p-value of less than 0.10, variable p-values of less than 0.20, at least one variable p-value less than 0.10 and an R-squared value of at least 0.5. Recommend selecting an alternate year for the model.";
                text2.Font.Bold = true;
                text2.WrapText = true;
                text2.EntireRow.RowHeight = 50;
            }
        }

        private void WriteResultsTable(int n, bool top)
        {
            // put the model results data into an object array
            //object[,] amodel ;

            Excel.Range header = BottomCell().get_Offset(2, 0);
            string start = header.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

            ArrayList lst = this.Source.Models.ModelSort();
            int first = top ? 0 : Math.Min(n, lst.Count);
            int last = top ? Math.Min(n, lst.Count) : lst.Count;

            
            if (first == last) return;

            // write the column headers
            //{ ModelNo, ModelValid, IVNames,IVSEPValChk, IVCoefficients, IVses, IVpVals, R2, adjR2, pVal, RMSError, Residual, AIC, Formula };
            foreach (Utilities.Constants.ModelOutputColumns col in System.Enum.GetValues(typeof(Utilities.Constants.ModelOutputColumns)))
            {
                string label = Globals.ThisAddIn.rsc.GetString("label" + col.ToString()) ?? col.ToString();
                header.Value2 = label;
                header = header.get_Offset(0, 1);
            }

            object[,] row;
            int rowct = 1;

            for (int j = first; j < last; j++)
            {
                Utilities.Model mdl = this.Source.Models.Item(j);
                int varct = mdl.VariableNames.Length;
                rowct += varct+1;
                //Added SEP Validation Check column here and incremented the columns number.
                //Excel.Range target = BottomCell().get_Offset(1,0).get_Resize(varct + 1, 13);
                Excel.Range target = BottomCell().get_Offset(1, 0).get_Resize(varct + 1, 14);
                
                //row = new object[1, 13];
                row = new object[1, 14];
                row[0, 0] = mdl.ModelNumber;
                row[0, 1] = mdl.Valid();
                //row[0, 6] = mdl.R2();
                //row[0, 7] = mdl.AdjustedR2();
                //row[0, 8] = mdl.ModelPValue();
                //row[0, 9] = mdl.RMSError;
                //row[0, 10] = mdl.ResidualSS();
                //row[0, 11] = mdl.AICFormula();
                //row[0, 12] = mdl.Formula();
                row[0, 7] = mdl.R2();
                row[0, 8] = mdl.AdjustedR2();
                row[0, 9] = mdl.ModelPValue();
                row[0, 10] = mdl.RMSError;
                row[0, 11] = mdl.ResidualSS();
                row[0, 12] = mdl.AICFormula();
                row[0, 13] = mdl.Formula();

                //target.get_Resize(1, 13).Value2 = row;
                target.get_Resize(1, 14).Value2 = row;

                if (top)
                {
                    Excel.Hyperlink hl = (Excel.Hyperlink)WS.Hyperlinks.Add(target.get_Resize(1, 1), "",
                        target.get_Resize(1, 1).get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing), System.Type.Missing, System.Type.Missing);
                    hl.ScreenTip = "Graph model " + mdl.ModelNumber.ToString();
                }

                //row = new object[varct+1,4];
                row = new object[varct + 1, 5];

                for (int i = 0; i < varct; i++)
                {
                    row[i, 0] = mdl.VariableNames[i];
                    //row[i, 1] = mdl.Coefficients[i];
                    //row[i, 2] = mdl.StandardErrors()[i];
                    //row[i, 3] = mdl.PValues()[i];
                    row[i, 1] = mdl.SEPValidationCheck()[i];
                    row[i, 2] = mdl.Coefficients[i];
                    row[i, 3] = mdl.StandardErrors()[i];
                    row[i, 4] = mdl.PValues()[i];
                }
               // row[varct, 0] = "(Intercept)";
                //row[varct, 1] = mdl.Coefficients[varct];
                row[varct, 2] = mdl.Coefficients[varct];

                //target.get_Offset(0, 2).get_Resize(varct+1, 4).Value2 = row;
                target.get_Offset(0, 2).get_Resize(varct + 1, 5).Value2 = row;

                // put a box around 
                target.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, System.Type.Missing);

                // highlight the best model
                if (Source.BestModel() != null)
                {
                    if (mdl.ModelNumber == Source.BestModel().ModelNumber)
                    {

                        if (mdl.Valid())
                        {
                            target.Font.Color = 0x00AA00;
                            target.Interior.Color = 0xCDEFC6;
                            //target.Style = Globals.ThisAddIn.rsc.GetString("bestModelStyle");
                            target.Font.Bold = true;
                        }
                        else
                        {
                            target.Font.Color = 0x0000AA;
                            target.Interior.Color = 0xCEC8FF;
                            //target.Style = "Bad";
                            target.Font.Bold = true;
                        }
                    }
                }
            }

            //Excel.Range tbl = WS.get_Range(start, System.Type.Missing).get_Resize(rowct, 13);
            Excel.Range tbl = WS.get_Range(start, System.Type.Missing).get_Resize(rowct, 14);
            Excel.ListObject LO = WS.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tbl, System.Type.Missing , Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            LO.TableStyle = "TableStyleLight8";
            LO.ShowTableStyleRowStripes = false;

            //((Excel.Range)LO.Range[0, 4]).EntireColumn.Hidden = true;
            //((Excel.Range)LO.Range[0, 5]).EntireColumn.Hidden = true;
            //((Excel.Range)LO.Range[0, 10]).EntireColumn.Hidden = true;
            //((Excel.Range)LO.Range[0, 11]).EntireColumn.Hidden = true;
            //((Excel.Range)LO.Range[0, 12]).EntireColumn.Hidden = true;

            //((Excel.Range)LO.Range[0, 6]).EntireColumn.NumberFormat = "0.0000";
            //((Excel.Range)LO.Range[0, 7]).EntireColumn.NumberFormat = "0.0000";
            //((Excel.Range)LO.Range[0, 8]).EntireColumn.NumberFormat = "0.0000";
            //((Excel.Range)LO.Range[0, 9]).EntireColumn.NumberFormat = "0.0000";

            ((Excel.Range)LO.Range[0, 4]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 5]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 6]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 11]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 12]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 13]).EntireColumn.Hidden = true;

            ((Excel.Range)LO.Range[0, 7]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 8]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 9]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 10]).EntireColumn.NumberFormat = "0.0000";

         }


        public void WriteResultsTable(int n, bool top,bool isRead)
        {
            // put the model results data into an object array
            //object[,] amodel ;

            Excel.Range header = BottomCell().get_Offset(2, 0);
            string start = header.get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing);

            ArrayList lst = this.Source.Models.ModelSort();
            int first = top ? 0 : Math.Min(n, lst.Count);
            int last = top ? Math.Min(n, lst.Count) : lst.Count;


            if (first == last) return;

            // write the column headers
            //{ ModelNo, ModelValid, IVNames, IVCoefficients, IVses, IVpVals, R2, adjR2, pVal, RMSError, Residual, AIC, Formula };
            foreach (Utilities.Constants.ModelOutputColumns col in System.Enum.GetValues(typeof(Utilities.Constants.ModelOutputColumns)))
            {
                string label = Globals.ThisAddIn.rsc.GetString("label" + col.ToString()) ?? col.ToString();
                header.Value2 = label;
                header = header.get_Offset(0, 1);
            }

            object[,] row;
            int rowct = 1;

            for (int j = first; j < last; j++)
            {
                Utilities.Model mdl = this.Source.Models.Item(j);
                int varct = mdl.VariableNames.Length;
                rowct += varct + 1;
                Excel.Range target = BottomCell().get_Offset(1, 0).get_Resize(varct + 1, 13);

                row = new object[1, 13];
                row[0, 0] = mdl.ModelNumber;
                row[0, 1] = mdl.Valid();
                row[0, 6] = mdl.R2();
                row[0, 7] = mdl.AdjustedR2();
                row[0, 8] = mdl.ModelPValue();
                row[0, 9] = mdl.RMSError;
                row[0, 10] = mdl.ResidualSS();
                row[0, 11] = mdl.AICFormula();
                row[0, 12] = mdl.Formula();

                target.get_Resize(1, 13).Value2 = row;

                if (top)
                {
                    Excel.Hyperlink hl = (Excel.Hyperlink)WS.Hyperlinks.Add(target.get_Resize(1, 1), "",
                        target.get_Resize(1, 1).get_Address(1, 1, Excel.XlReferenceStyle.xlA1, System.Type.Missing, System.Type.Missing), System.Type.Missing, System.Type.Missing);
                    hl.ScreenTip = "Graph model " + mdl.ModelNumber.ToString();
                }

                row = new object[varct + 1, 4];

                for (int i = 0; i < varct; i++)
                {
                    row[i, 0] = mdl.VariableNames[i];
                    row[i, 1] = mdl.Coefficients[i];
                    row[i, 2] = mdl.StandardErrors()[i];
                    row[i, 3] = mdl.PValues()[i];
                }
                //row[varct, 0] = "(Intercept)";
                row[varct, 1] = mdl.Coefficients[varct];

                target.get_Offset(0, 2).get_Resize(varct + 1, 4).Value2 = row;

                // put a box around 
                target.BorderAround(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, System.Type.Missing);

                // highlight the best model
                if (Source.BestModel() != null)
                {
                    if (mdl.ModelNumber == Source.BestModel().ModelNumber)
                    {

                        if (mdl.Valid())
                        {
                            target.Font.Color = 0x0000AA;
                            //target.Style = Globals.ThisAddIn.rsc.GetString("bestModelStyle");
                            target.Font.Bold = true;
                        }
                        else
                        {
                            target.Font.Color = 0x0000AA;
                            //target.Style = "Bad";
                            target.Font.Bold = true;
                        }
                    }
                }
            }

            Excel.Range tbl = WS.get_Range(start, System.Type.Missing).get_Resize(rowct, 13);
            Excel.ListObject LO = WS.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, tbl, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);
            LO.TableStyle = "TableStyleLight8";
            LO.ShowTableStyleRowStripes = false;

            ((Excel.Range)LO.Range[0, 4]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 5]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 10]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 11]).EntireColumn.Hidden = true;
            ((Excel.Range)LO.Range[0, 12]).EntireColumn.Hidden = true;

            ((Excel.Range)LO.Range[0, 6]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 7]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 8]).EntireColumn.NumberFormat = "0.0000";
            ((Excel.Range)LO.Range[0, 9]).EntireColumn.NumberFormat = "0.0000";

        }

        #region //Charting

        internal void newChart(int n)
        {
            Excel.Range chartText1 = BottomCell().get_Offset(2, 0).get_Resize(1, 3);
            Excel.Range chartText2 = chartText1.get_Offset(0, 6).get_Resize(1, 8);

            ((Excel.Range)chartText1[1, 1]).Value2 = "The plots below show the actual energy consumption versus the independent variables for the model year.";
            ((Excel.Range)chartText2[1, 1]).Value2 = "The line labeled \"Actuals\"  in the plot below has not been adjusted. This is the original data entered by the user. The line labeled \"Model\" is the predicted energy consumption using the model selected above.";

            chartText1.EntireRow.RowHeight = 75;
            chartText1.EntireRow.Font.Bold = true;
            chartText1.EntireRow.WrapText = true;
            chartText1.Merge();
            chartText2.Merge();

            Excel.Range start = BottomCell().get_Offset(1, 0);

            double topleft;
            if (!double.TryParse(start.Top.ToString(), out topleft))
                topleft = 0;

            start.EntireRow.RowHeight = Utilities.Constants.CHART_HEIGHT * 1.1;

            ArrayList lst = this.Source.Models.ModelSort();
            string[] variables = new string[]{};
            int longest = 0;
            foreach (Utilities.Model model in lst)
            {
                if (model.VariableNames.Length > longest)
                {
                    longest = model.VariableNames.Length;
                    variables = model.VariableNames;
                }
            }

            object[] sources = new object[variables.Length];

            int k = 0;
            for (int i = 0; i < variables.Length; i++)
            {
                double[] temp = new double[Source.knownXs.Rows.Count];

                foreach (System.Data.DataRow row in Source.knownXs.Rows)
                {
                    temp[k] = Convert.ToDouble(row.ItemArray[i]);
                    k++;
                }

                k = 0;
                sources[i] = temp;
            }

            for(int i = 0; i < variables.Length; i++)
            {
                if (i.Equals(0))
                {
                    Excel.ChartObject CO = ((Excel.ChartObjects)WS.ChartObjects(System.Type.Missing))
                    .Add(10, topleft, Utilities.Constants.CHART_WIDTH, Utilities.Constants.CHART_HEIGHT);

                    CO.Chart.HasLegend = false;
                    CO.Chart.ChartType = Excel.XlChartType.xlXYScatter;
                    CO.Chart.HasTitle = false;

                    Excel.Series newSeries = ((Excel.SeriesCollection)CO.Chart.SeriesCollection(System.Type.Missing)).NewSeries();
                    newSeries.Values = Source.Ys;
                    newSeries.XValues = sources[i];

                    Excel.Axis COYaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    COYaxis.TickLabels.NumberFormat = "###,##0";
                    COYaxis.HasTitle = true;
                    COYaxis.AxisTitle.Text = WS.Name;

                    Excel.Axis COXaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    COXaxis.TickLabels.NumberFormat = "###,##0";
                    COXaxis.HasTitle = true;
                    COXaxis.AxisTitle.Text = variables[i];

                }
                else
                {
                    Excel.Range afterFirst = BottomCell().get_Offset(1, 0);
                    afterFirst.EntireRow.RowHeight = Utilities.Constants.CHART_HEIGHT * 1.1;

                    double topleftinner;
                    if (!double.TryParse(afterFirst.Top.ToString(), out topleftinner))
                        topleftinner = 0;

                    Excel.ChartObject CO = ((Excel.ChartObjects)WS.ChartObjects(System.Type.Missing))
                    .Add(10, topleftinner, Utilities.Constants.CHART_WIDTH, Utilities.Constants.CHART_HEIGHT);

                    CO.Chart.HasLegend = false;
                    CO.Chart.ChartType = Excel.XlChartType.xlXYScatter;
                    CO.Chart.HasTitle = false;

                    Excel.Series newSeries = ((Excel.SeriesCollection)CO.Chart.SeriesCollection(System.Type.Missing)).NewSeries();
                    newSeries.Values = this.Source.Ys;
                    newSeries.XValues = sources[i];

                    Excel.Axis COYaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                    COYaxis.TickLabels.NumberFormat = "###,##0";
                    COYaxis.HasTitle = true;
                    COYaxis.AxisTitle.Text = WS.Name;

                    Excel.Axis COXaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                    COXaxis.TickLabels.NumberFormat = "###,##0";
                    COXaxis.HasTitle = true;
                    COXaxis.AxisTitle.Text = variables[i];
                    
                }
            }

            for (int j = 0; j < Math.Min(n, lst.Count); j++)
            {
                Utilities.Model mdl = this.Source.Models.Item(j);
                Excel.ChartObject CO = ((Excel.ChartObjects)WS.ChartObjects(System.Type.Missing))
                    .Add(400, topleft, Utilities.Constants.CHART_WIDTH, Utilities.Constants.CHART_HEIGHT);
                CO.Placement = Excel.XlPlacement.xlMove;
                CO.Name = "Model " + mdl.ModelNumber;
                CO.Visible = false;

                CO.Chart.ChartType = Excel.XlChartType.xlLineMarkers;
                CO.Chart.ChartStyle = 5;
                

                Excel.Axis CharObj2Yaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary);
                CharObj2Yaxis.TickLabels.NumberFormat = "###,##0";
                CharObj2Yaxis.HasTitle = true;
                CharObj2Yaxis.AxisTitle.Text = "Input Interval";

                Excel.Axis CharObj2Xaxis = (Excel.Axis)CO.Chart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlPrimary);
                CharObj2Xaxis.TickLabels.NumberFormat = "###,##0";
                CharObj2Xaxis.HasTitle = true;
                CharObj2Xaxis.AxisTitle.Text = WS.Name;
                //Modified by suman: TFS ticket :69333
                Excel.Series newSeries = ((Excel.SeriesCollection)CO.Chart.SeriesCollection(System.Type.Missing)).NewSeries();            
                newSeries.Values = this.Source.Ys;
                newSeries.Name = "Actuals";

                Excel.Series modelSeries = ((Excel.SeriesCollection)CO.Chart.SeriesCollection(System.Type.Missing)).NewSeries();
                modelSeries.Values = mdl.PredictedYs();
                modelSeries.Name = "Model " + mdl.ModelNumber;

                newSeries.Format.Line.ForeColor.RGB = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.DarkRed);//  0x4177B8;//12089153; // Calculated from color calculator
                newSeries.MarkerStyle = Excel.XlMarkerStyle.xlMarkerStyleSquare;
                newSeries.MarkerBackgroundColorIndex = (Microsoft.Office.Interop.Excel.XlColorIndex)9;
                newSeries.Format.Line.Transparency = 1.0f;

           
                
                if (mdl == this.Source.BestModel()) CO.Visible = true;
                CO.Chart.AutoScaling = true;
                CO.Chart.Refresh();
                CO.Chart.HasTitle = true;
                CO.Chart.ChartTitle.Text = "Comparison of Actual and Model Data";
                CO.Chart.Legend.Position = Excel.XlLegendPosition.xlLegendPositionTop;
            }
        }

       #endregion

    }
}
