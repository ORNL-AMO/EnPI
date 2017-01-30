using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using Microsoft.Office.Tools.Excel.Extensions;
using AMO.EnPI.AddIn.Utilities;
using Integration = System.Windows.Forms.Integration;
//using ITP.Utilities.Data;
//using ITP.Utilities.Business;

namespace AMO.EnPI.AddIn
{
    public partial class ThisAddIn
    {
        public List<string> SelectedSources;
        public List<string> SelectedProduction;
        public List<string> SelectedVariables;
        public List<string> SelectedBuildings;
        public ArrayList Years;
        public int[] modeledSourceIndex;
        public string[,] energyCostColumnMatchArray;
        //Added By suman TFS Ticket:68998
        public IDictionary<string, string> CO2EmissionFactors;
        public bool hasSEPValidationError = false;
        //Added by suman SEP Changes 
        public string sepValidationWarningMsg = string.Empty;
        public IList<SEPValidationValues> lstSEPValidationValues;
        public IDictionary<string, string> SelectedSourcesBestModelFormulas = new Dictionary<string, string>();
     
        public bool fromWizard = false;
        public bool fromEnergyCost = false;
        public bool fromCO2Emission = false;
        public bool fromRegression = true;
        public string BaselineYear;
        public string SelectedYear;
        public string ReportYear;
        public string ModelYearSelected;
        public string AdjustmentMethod;
        string dataSheetName = EnPIResources.stateDataSheetName;
        public EnPIDataSet EnPIData;
        public ModelSheetCollection ModelSheets;
        public ModelCollection ModelCollection;
        public MasterGroupSheetCollection masterGroupCollection;
        public GroupSheetCollection groupSheetCollection;
        public static int paneWidth = 250;

        public Microsoft.Office.Tools.CustomTaskPane wizardPane;
        public System.Resources.ResourceManager rsc = EnPIRibbon.rsc;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            wizardInit();
            
            masterGroupCollection = new MasterGroupSheetCollection();

            Globals.ThisAddIn.Application.SheetFollowHyperlink += new Microsoft.Office.Interop.Excel.AppEvents_SheetFollowHyperlinkEventHandler(refreshChart);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
      
        private void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            Excel.Worksheet ws = ExcelHelpers.GetWorksheet(Wb, dataSheetName);
            if (ws != null)
                loadStateData(ws, Wb);

            if (masterGroupCollection.Count > 0)
            {
                Globals.Ribbons.Ribbon1.Model.Visible = true;
                Globals.Ribbons.Ribbon1.Rollup.Visible = true;
            }

            //this will prevent the "would you like to save" dialog showing even when the user hasn't done anything
            Wb.Saved = true;
        }
        
        private void Application_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            bool EnPIexists = false;

            foreach (GroupSheetCollection gsc in masterGroupCollection)
            {
                if (gsc.WB == Wb)
                {
                    EnPIexists = true;
                    break;
                }
            }

            //only add the hidden sheet if there are EnPI results sheets
            if (EnPIexists)
            {
                Excel.Worksheet ws = ExcelHelpers.GetWorksheet(Wb, dataSheetName);

                if (ws == null)
                //create dataSheet if it doesn't exist
                {
                    
                    ws = (Excel.Worksheet)Wb.Worksheets.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    ws.Name = dataSheetName;
                    ws.Visible = Excel.XlSheetVisibility.xlSheetVeryHidden;//.xlSheetVisible;//
                }
                    
                saveStateData(ws);
                
                
            }
        }

        private void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            this.CustomTaskPanes.Remove(wizardPane);
            wizardPane= this.CustomTaskPanes.SingleOrDefault();
        }
        
        private void removeEntriesFromMasterGroupCollection(string wbName)
        {
            
                MasterGroupSheetCollection tempGroupCollection = new MasterGroupSheetCollection();
                foreach (GroupSheetCollection gsc in this.masterGroupCollection)
                {
                    if (gsc.WBName != wbName)
                    {
                        tempGroupCollection.Add(gsc);
                    }
                }
                masterGroupCollection = tempGroupCollection;
          
        }
        private void loadStateData(Excel.Worksheet ws, Excel.Workbook Wb)
        {
            try
            {
                removeEntriesFromMasterGroupCollection(Wb.Name);
                foreach (Excel.ListObject LO in ws.ListObjects)
                {
                    GroupSheetCollection gsc = new GroupSheetCollection();
                    gsc.WB = Wb;
                    gsc.WBName = Wb.Name;
                    int rowCount = 0;
                    int r;
                    foreach (Excel.ListRow row in LO.ListRows)
                    {
                        rowCount++;
                        string sheetguid = row.Range.get_Resize(1, 1).Value2.ToString();
                        string iter = row.Range.get_Resize(1, 1).Value2.ToString() == "regressionIteration" ? row.Range.get_Offset(0, 1).get_Resize(1, 1).Value2.ToString() : null;
                        string pltnm = row.Range.get_Resize(1, 1).Value2.ToString() == "plantName" ? row.Range.get_Offset(0, 1).get_Resize(1, 1).Value2.ToString() : null;

                        Excel.Worksheet workSheet = ExcelHelpers.GetWorksheetbyGUID(Wb, sheetguid);

                        if (workSheet != null)
                        {
                            GroupSheet gs = new GroupSheet((Excel.Worksheet)workSheet, Convert.ToBoolean(((Excel.Range)row.Range.Cells[1, 3]).Value2.ToString()), Convert.ToBoolean(((Excel.Range)row.Range.Cells[1, 4]).Value2.ToString()), ws.Name);

                            if (gs.adjustedDataSheet && gs.WS.get_Range("A1").Value2 != null) //.ToString()
                                gs.Name = gs.WS.get_Range("A1").Value2.ToString();
                            gsc.Add(gs);
                        }

                        gsc.regressionIteration = int.TryParse(iter, out r) ? r : gsc.regressionIteration;
                        gsc.PlantName = pltnm == null ? gsc.PlantName : pltnm;
                    }
                    if (masterGroupCollection.IndexOf(gsc.WBName, gsc.regressionIteration) < 0)
                        masterGroupCollection.Add(gsc);   // add the workbooks sheet collections to the master collection
                }
                //this will prevent the "would you like to save" dialog showing even when the user hasn't done anything
                Wb.Saved = true;
            }
            catch (Exception ex)
            {
                //throw ex;
                //TODO: Add Exception and Logging module
            }
        }

       private void saveStateData(Excel.Worksheet ws)
       { 
            //clear previous contents
            ws.Rows.Delete();

            int collectionCount = 0;
            foreach (GroupSheetCollection gsc in masterGroupCollection)
            {
                // only write the sheet collections that are part of the workbook
                if (gsc.WB == (Excel.Workbook)ws.Parent)
                {
                    Excel.Range range;

                    if (collectionCount > 0)
                        range = (Excel.Range)ws.get_Range("A1", "D1").get_Offset(0, (collectionCount) * 5);
                    else
                        range = (Excel.Range)ws.get_Range("A1", "D1").get_Offset(0, 0);

                    int rowCount = 0;
                    foreach (GroupSheet gs in gsc)
                    {
                        // only write the sheets that haven't been deleted
                        if (ExcelHelpers.GetWorksheetbyGUID(gsc.WB, gs.WSGUID) != null)
                        {
                            range.get_Resize(1, 1).get_Offset(rowCount, 0).Value2 = gs.WSGUID;
                            range.get_Resize(1, 1).get_Offset(rowCount, 1).Value2 = gs.WS.Name;
                            range.get_Resize(1, 1).get_Offset(rowCount, 2).Value2 = gs.outputSheet.ToString();
                            range.get_Resize(1, 1).get_Offset(rowCount, 3).Value2 = gs.adjustedDataSheet.ToString();
                            rowCount++;
                        }
                    }

                    if (rowCount > 0)
                    {
                        collectionCount++;

                        range.get_Resize(1, 1).get_Offset(rowCount, 0).Value2 = "regressionIteration";
                        range.get_Resize(1, 1).get_Offset(rowCount, 1).Value2 = gsc.regressionIteration;
                        range.get_Resize(1, 1).get_Offset(rowCount+1, 0).Value2 = "plantName";
                        range.get_Resize(1, 1).get_Offset(rowCount+1, 1).Value2 = gsc.PlantName;

                        range = range.get_Resize(rowCount + 2, 4);
                        Excel.ListObject LO = ws.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range, System.Type.Missing, Excel.XlYesNoGuess.xlNo, System.Type.Missing);
                    }
                }
            }
        }

        public void actualEnPI(Excel.ListObject LO)
        {
            groupSheetCollection = new GroupSheetCollection();
            groupSheetCollection.PlantName = ((Excel.Worksheet)LO.Parent).Name;
            groupSheetCollection.WB = (Excel.Workbook)((Excel.Worksheet)LO.Parent).Parent;
            groupSheetCollection.WBName = groupSheetCollection.WB.Name;
            masterGroupCollection.Add(groupSheetCollection);
            groupSheetCollection.regressionIteration = masterGroupCollection.WorkbookNextIteration(groupSheetCollection.WB);

            object[,] sourcerows = ExcelHelpers.getYearArray(LO, Years);
            DataTable dtSourceData = DataHelper.ConvertToDataTable(LO.HeaderRowRange.Value2 as object[,], sourcerows);

            EnPIData = new EnPIDataSet();

            EnPIData.WorksheetName = ((Excel.Worksheet)LO.Parent).Name;
            EnPIData.SourceData = dtSourceData;
            EnPIData.ListObjectName = LO.Name;
            EnPIData.BaselineYear = BaselineYear;

            EnPIData.ProductionVariables = SelectedProduction;
            EnPIData.BuildingVariables = SelectedBuildings;

            SelectedSources.Add(rsc.GetString("unadjustedTotalColName")); 
            EnPIData.EnergySourceVariables = SelectedSources;

            EnPIData.Init();

            fromRegression = false;

            AdjustedDataSheet adjData = new AdjustedDataSheet(EnPIData);
            adjData.Populate(false);

            ActualSheet results = new ActualSheet(EnPIData);
            results.DetailDataSheet = adjData.thisSheet;
            results.AdjustedData = adjData.AdjustedData;
            results.Populate();

            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            // cleanup
            ExcelHelpers.Cleanup(Globals.ThisAddIn.Application.ActiveWorkbook);

        }

        public void plotEnPI(Excel.ListObject LO)
        {
            groupSheetCollection = new GroupSheetCollection();
            groupSheetCollection.PlantName = ((Excel.Worksheet)LO.Parent).Name;
            groupSheetCollection.WB = (Excel.Workbook)((Excel.Worksheet)LO.Parent).Parent;
            groupSheetCollection.WBName = groupSheetCollection.WB.Name;
            masterGroupCollection.Add(groupSheetCollection);
            groupSheetCollection.regressionIteration = masterGroupCollection.WorkbookNextIteration(groupSheetCollection.WB);


            // create a data table with the data to be adjusted
            object[,] sourcerows = ExcelHelpers.getYearArray(LO, Years);
            DataTable dtSourceData = DataHelper.ConvertToDataTable(LO.HeaderRowRange.Value2 as object[,], sourcerows);

            EnPIData = new EnPIDataSet();
            EnPIData.WorksheetName = ((Excel.Worksheet)LO.Parent).Name;
            EnPIData.SourceData = dtSourceData;
            EnPIData.ListObjectName = LO.Name;
            EnPIData.BaselineYear = BaselineYear;
            EnPIData.ModelYear = SelectedYear;
            EnPIData.ReportYear = ReportYear; //Added By Suman for SEP Changes
            EnPIData.Years = Years.Cast<string>().ToList();

            if (Utilities.Constants.MODEL_TOTAL)
            {
                SelectedSources.Add(rsc.GetString("unadjustedTotalColName"));
            }
            EnPIData.EnergySourceVariables = SelectedSources;
            EnPIData.IndependentVariables = SelectedVariables;
            EnPIData.ProductionVariables = SelectedProduction;
            EnPIData.BuildingVariables = SelectedBuildings;

            EnPIData.Init();

            ModelSheets = new ModelSheetCollection();
            foreach (Utilities.EnergySource src in EnPIData.EnergySources)
            {
                ModelSheet aSrc = ModelSheets.Add(src);
                aSrc.Populate();
                aSrc.WS.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            }

            AdjustedDataSheet adjData = new AdjustedDataSheet(EnPIData);
            adjData.Populate(true);

            EnPISheet results = new EnPISheet(EnPIData, false);
            results.AdjustedDataSheet = adjData.thisSheet;
            results.AdjustedData = adjData.AdjustedData;
            results.Populate();

            //EnPISheet senpiresults = new EnPISheet(EnPIData, true);
            //senpiresults.AdjustedDataSheet = adjData.thisSheet;
            //senpiresults.AdjustedData = adjData.AdjustedData;
            //senpiresults.Populate();


            SEPSheet sepResults = new SEPSheet(EnPIData);
            sepResults.AdjustedDataSheet = adjData.thisSheet;
            sepResults.AdjustedData = adjData.AdjustedData;
            sepResults.Populate();

            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            // cleanup
            ExcelHelpers.Cleanup(Globals.ThisAddIn.Application.ActiveWorkbook);

        }

        #region //wizard pane controls
        public void wizardInit()
        {
            WizardControl wControl = new WizardControl();

            wizardPane = this.CustomTaskPanes.Add(wControl, "EnPI Step-by-step Wizard");
        }
        //This method is mostly needed for Office versions above 2010 as the microsoft has changed the design of handling workbook from Mutilple Display interface to Single Display Interface.
        //For reference: http://msdn.microsoft.com/en-us/library/office/dn251093(v=office.15).aspx
        public void CheckforVaildWizard()
        {
            
            try
            {
                if (wizardPane != null)
                {
                    if ((this.wizardPane.Visible == true) || (this.wizardPane.Visible == false))
                    {
                        //Do nothing
                    }
                }
                else
                {
                    wizardInit();
                }
            }
            catch (Exception ex)
            {
                this.CustomTaskPanes.Remove(wizardPane);
                if (this.CustomTaskPanes.Count > 0)
                    wizardPane = this.CustomTaskPanes.FirstOrDefault();
                else
                    wizardInit();
            }
            
        }

        public void wizardInit(int step)
        {
            WizardControl wControl = new WizardControl(step);

            wizardPane = this.CustomTaskPanes.Add(wControl, "EnPI Step-by-step Wizard");
        }

        public void showWizard()
        {
            if (wizardPane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionRight || wizardPane.DockPosition == Office.MsoCTPDockPosition.msoCTPDockPositionLeft)
            {
                wizardPane.Width = paneWidth;
            }
            wizardPane.Visible = true;
            wizardPane.Control.AutoScroll = true;
        }

        public void hideWizard()
        {
            if (wizardPane.Width != paneWidth && wizardPane.Width != 0)
                paneWidth = wizardPane.Width;

            wizardPane.Visible = false;
        }
 
        public void LaunchWizardControl()
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            wizardInit();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchWizardControl(int step)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            wizardInit(step);

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchRegressionControl(Constants.EnPITypes type)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            RegressionControl newControl = new RegressionControl(type);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Select Energy Sources");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchReportingPeriodControl(bool fromWizard)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();
            ReportingPeriodControl newControl = new ReportingPeriodControl(fromWizard);
            
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Select Energy Sources");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();

        }

        public void LaunchUnitConversionControl(bool fromWizard)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            UnitConversionControl newControl = new UnitConversionControl(fromWizard);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Unit Conversion");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchUnitConversionControl(bool fromWizard, CheckedListBox c, ComboBox box)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            UnitConversionControl newControl = new UnitConversionControl(fromWizard,c, box);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Unit Conversion");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchChangeModelControl(bool fromWizard)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            ChangeModelControl newControl = new ChangeModelControl(fromWizard);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Switch Between Models");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchFuelUnitConversionControl(bool fromWizard, CheckedListBox c, ComboBox box)
        {
            
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            FuelUnitConversionControl newControl = new FuelUnitConversionControl(fromWizard,c, box);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Fuel Unit Conversion");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchRollupControl(bool fromWizard)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            RollupControl newControl = new RollupControl(fromWizard);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Rollup");
            newControl.Open();

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchEnergyCostControl(CheckedListBox clb, RegressionControl parentControl, System.Windows.Forms.Control.ControlCollection controls)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            EnergyCostControl newControl = new EnergyCostControl(parentControl);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "Energy Cost");
            newControl.Open(clb, controls);

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }

        public void LaunchCO2EmissionControl(CheckedListBox clb, RegressionControl parentControl, System.Windows.Forms.Control.ControlCollection controls)
        {
            if (Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.hideWizard();

            CO2EmissionControl newControl = new CO2EmissionControl(parentControl);
            Globals.ThisAddIn.CustomTaskPanes.Remove(Globals.ThisAddIn.wizardPane);
            Globals.ThisAddIn.wizardPane = Globals.ThisAddIn.CustomTaskPanes.Add(newControl, "CO2 Avoided Emission Data");
            newControl.Open(clb, controls);

            if (!Globals.ThisAddIn.wizardPane.Visible)
                Globals.ThisAddIn.showWizard();
        }


        #endregion

        public bool CheckForYear(Excel.ListObject LO)
        {
            if (Utilities.ExcelHelpers.GetListColumn(LO, rsc.GetString("yearColName")) == null)
                return false;

            return true;
        }

        public bool NegativeCheck(Excel.ListObject LO, int[] indexList)
        {
            bool hasNegative = false;
            if (indexList != null)
            {
                //check if values are in given indicies of a list object are negative, if they are highlight them in yellow
                for (int j = 0; j < indexList.Length; j++)
                {
                    for (int i = 2; i < LO.Range.Rows.Count; i++)
                    {
                        try
                        {
                            Excel.Range negativeCheck = (Excel.Range)LO.Range[i, indexList[j]];

                            //remove any formatting in case the user is changeing models and the new value is not negative
                            if (System.Convert.ToInt32(negativeCheck.Cells.Interior.Color).Equals(System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow)))
                            {
                                negativeCheck.Cells.Interior.ColorIndex = 0;
                            }

                            //add yellow background if the value is negative
                            if ((double)negativeCheck.Value2 < 0)
                            {
                                hasNegative = true;
                                negativeCheck.Cells.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                            }

                        }
                        catch (Exception ex)
                        {
                            //DO Nothing keep the loop running 
                        }
                    }
                }

            }
            return hasNegative;
        }

        private Excel.Range BottomCell(Excel.Worksheet ws)
        {
            string addr = "A" + Utilities.ExcelHelpers.writeAppendBottomAddress(ws, 0).ToString();

            return (Excel.Range)ws.get_Range(addr, System.Type.Missing);
        }

        internal string addProductionColumn(Excel.Worksheet WS, Excel.ListObject LO, string dataListName)
        {
            string colName = "";
            string formula = "";
            Excel.ListObject dataList = ExcelHelpers.GetListObject(WS, dataListName);
            object[,] headers = dataList.HeaderRowRange.Value2 as object[,];
            string dttbl = dataListName;
            string dtyearcol = ExcelHelpers.GetListColumn(dataList, EnPIResources.yearColName).Name;
            string yearcol = ExcelHelpers.GetListColumn(LO, EnPIResources.yearColName).Name;

            int i = headers.GetLowerBound(0);
            for (int j = headers.GetLowerBound(1); j <= headers.GetUpperBound(1); j++)
            {
                if (DataHelper.isProduction(headers[i, j].ToString()))
                    formula += "SUM(" + dttbl + ExcelHelpers.CreateValidFormulaName(headers[i, j].ToString()) + " " +
                                   DataHelper.RowRangebyMatch(dttbl, dtyearcol, "[[#This Row]," + yearcol + "]", dataSheetName)
                                   + ")" + "+";
            }
            if (formula != "")
            {
                Excel.ListColumn newCol = LO.ListColumns.Add(2);
                colName = "Production";
                newCol.Name = colName;
                newCol.DataBodyRange.Value2 = "=" + formula.Substring(0, formula.Length - 1);
                newCol.DataBodyRange.Style = "Comma [0]";
            }

            return colName;
        }
        
        internal void addEIColumn(Excel.Worksheet WS, Excel.ListObject LO, string dataListName)
        {
            string formula = "";
            string newname = rsc.GetString("buildingEnPIColName");
            string tbl = LO.Name;
            string bsqfcol = rsc.GetString("buildingSQFColName");
            string colhat = rsc.GetString("totalAdjValuesColName");
            string that;    //adjusted current year energy consumption
            string bsqf;    //building square feet

            that = tbl + "[[#This Row]," + ExcelHelpers.CreateValidFormulaName(colhat) + "]";
            bsqf = tbl + "[[#This Row]," + ExcelHelpers.CreateValidFormulaName(bsqfcol) + "]";

            formula = "=IFERROR(" + that + "/" + bsqf + ",0)";

            if (bsqfcol != "")
            {
                Excel.ListColumn newCol = LO.ListColumns.Add(missing);
                newCol.Name = newname;
                newCol.DataBodyRange.Value2 = formula;
                newCol.DataBodyRange.Style = "Comma";
            }
        }

        internal void refreshChart(object sender, Excel.Hyperlink e)
        {
            string mdl = ((Excel.Worksheet)this.Application.ActiveSheet).get_Range(e.Name).Value2.ToString();
            Excel.ChartObjects col = (Excel.ChartObjects)((Excel.Worksheet)this.Application.ActiveSheet).ChartObjects();

            int tmp1 = col.Count;
            for (int i = 1; i <= col.Count; i++ )
            {
                Excel.ChartObject co = (Excel.ChartObject)col.Item(i);
                if (co.Name.Substring(0, 5) == "Model")
                {
                    if (co.Name == "Model " + mdl)
                        co.Visible = true;
                    else
                        co.Visible = false;
                }
            }

        }

      


        
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
            this.Application.WorkbookOpen += new Excel.AppEvents_WorkbookOpenEventHandler(Application_WorkbookOpen);
            this.Application.WorkbookBeforeSave += new Excel.AppEvents_WorkbookBeforeSaveEventHandler(Application_WorkbookBeforeSave);
            this.Application.WorkbookBeforeClose +=new Excel.AppEvents_WorkbookBeforeCloseEventHandler(Application_WorkbookBeforeClose);
            
        }
        
        #endregion

        
    }

    public class GroupSheet
    {
        public Excel.Worksheet WS { get; set; }
        public string WSGUID { get; set; }
        public string Name { get; set; }
        public bool outputSheet { get; set; }
        public bool adjustedDataSheet { get; set; }

        public GroupSheet()
        {
            WS = new Excel.Worksheet();
            WSGUID = WS.CodeName == "" ? ExcelHelpers.getWorksheetCustomProperty(WS, "SheetGUID") : WS.CodeName;
            outputSheet = false;
            adjustedDataSheet = false;
        }
        public GroupSheet(Excel.Worksheet WS, bool outputSheet, bool adjustedDataSheet, string WSName)
        {
            this.WS = WS;
            this.WSGUID = WS.CodeName == "" ? ExcelHelpers.getWorksheetCustomProperty(WS, "SheetGUID") : WS.CodeName;
            this.Name = WSName;
            this.outputSheet = outputSheet;
            this.adjustedDataSheet = adjustedDataSheet;
        }

        public string WSName()
        {
            try
            {
                return WS.Name;
            }
            catch
            {
                return null;
            }
        }
    }

    public class GroupSheetCollection : System.Collections.CollectionBase
    {
        public Excel.Workbook WB { get; set; }
        public string WBName { get; set; }
        public string PlantName { get; set; }
        public int regressionIteration;

        public void Add(GroupSheet sheet)
        {
            List.Add(sheet);
        }

        public GroupSheet New()
        {
            GroupSheet sheet = new GroupSheet();
            List.Add(sheet);

            return sheet;
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

        public GroupSheet Item(int Index)
        {
            return (GroupSheet)List[Index];
        }

    }

    [Serializable()]
    public class MasterGroupSheetCollection : System.Collections.CollectionBase
    {
        public void Add(GroupSheetCollection group)
        {
            List.Add(group);
        }

        public GroupSheetCollection New()
        {
            GroupSheetCollection group = new GroupSheetCollection();
            List.Add(group);

            return group;
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

        public GroupSheetCollection Item(int Index)
        {
            return (GroupSheetCollection)List[Index];
        }

        public int IndexOf(string WBName, int regressionIteration)
        {
            for (int i = 0; i < List.Count; i++)
            {
                GroupSheetCollection gsc = (GroupSheetCollection)List[i];
                if (gsc.WBName == WBName && gsc.regressionIteration == regressionIteration)
                    return i;
            }

            return -1;
        }
        public int WorkbookNextIteration(Excel.Workbook WB)
        {
            int j = 0;
            foreach (GroupSheetCollection gsc in InnerList)
            {
                if (gsc.WBName == WB.Name) j = Math.Max(j, gsc.regressionIteration); 
            }
            
            return ValidateNextRegressionIteration(j+1, WB);
            //return j + 1;
        }
        //Added By suman TFS Ticket: 68832
        //Note: In some cases the regression iteration number is not saved properly in the stateData sheet , as a result of that the work sheet numbers are not getting generated correctly.
        //In order to avoid that, verify the generated iteration number with the all the work sheets in the work book and then if the number still exists increment it recursively and then 
        //return the correct next one.
        public int ValidateNextRegressionIteration(int iteration,Excel.Workbook WB)
        {
            int count = 0;
            foreach (Excel.Worksheet ws in WB.Worksheets)
            {
                if (ws.Name.StartsWith(iteration.ToString()))
                {
                    count++;

                }
            }

            if (count > 0)
            {

               return  ValidateNextRegressionIteration(iteration +1, WB);
            }
            else
            {
                return iteration;
            }
        }

    }
}
