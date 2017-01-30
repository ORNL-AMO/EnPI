using System;
using System.Collections;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace AMO.EnPI.AddIn.Utilities
{
    public class ExcelHelpers
    {
        // AMO.EnPI.Utilities.ExcelHelpers
        //
        // This class contains the helper methods for Excel. The class does depend on the Excel Interop
        // classes, but uses components common to both versions 11 and 12 of the Interop.  
        // 

        private static object missing = System.Type.Missing;

        #region // data table helpers

        // rangeTable: Returns a datatable created from a list object that contains the same rows as the selected range
        //      <param name="LO">The list object.</param>
        //      <param name="srcRange">The range containing the rows to include</param>
        public static DataTable rangeTable(Excel.ListObject LO, Excel.Range srcRange)
        {
            //use the range to select a range with all of the list columns
             string tmp = srcRange.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);
            object[,] tmpo = LO.Range.Value2 as object[,];
            
            // find the first row of the selection, and the number of rows to include
            int startRow = srcRange.Row;
            int ctRows = srcRange.Rows.Count;

            // the start row of the list object is the start row of the range minus the start row of the header 
            int startLO = LO.HeaderRowRange.Row;

            // get the portion of the LO that matches the selection
            Excel.Range partialLO = LO.Range.get_Offset(startRow - startLO, 0).get_Resize(ctRows, LO.ListColumns.Count);

            //convert the range to a data table
            return DataHelper.ConvertToDataTable(LO.HeaderRowRange.Value2 as object[,],
                                    partialLO.Value2 as object[,]);

        }

        public static DataTable rangeTable(Excel.ListObject LO)
        {
            // find the first row of the selection, and the number of rows to include
            int ctRows = LO.Range.Rows.Count - 1;
            if (LO.ShowTotals) ctRows = ctRows - 1;

            // get the data portion of the LO 
            Excel.Range partialLO = LO.Range.get_Offset(1, 0).get_Resize(ctRows, LO.ListColumns.Count);

            //convert the range to a data table
            return DataHelper.ConvertToDataTable(LO.HeaderRowRange.Value2 as object[,],
                                    partialLO.Value2 as object[,]);

        }

        public static Excel.ListObject writeDataTable(string startCell, ref Excel.Worksheet WS, ref DataTable dt)
        {
            // create list objects for the model year and adjusted year data
            string start = startCell;
            int rowOffset = 3;
            Excel.Range range1 = WS.get_Range(start, missing)
                                //.get_Offset(rowOffset, 0)
                                .get_Resize(dt.Rows.Count + 1, dt.Columns.Count);

            range1.get_Resize(1, dt.Columns.Count).Value2 = DataHelper.dataTableHeaders(dt);

            range1.get_Offset(1, 0).get_Resize(dt.Rows.Count, dt.Columns.Count).Value2 = DataHelper.dataTableArray(dt);

            // add in formulae
            foreach (DataColumn dc in dt.Columns)
            {
                if (dc.ExtendedProperties["Formula"] != null)
                {
                    range1.get_Offset(1, dc.Ordinal).get_Resize(dt.Rows.Count, 1).Value2 = dc.ExtendedProperties["formula"];
                }
            }

            Excel.ListObject newObject = WS.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, range1, missing, Excel.XlYesNoGuess.xlYes, missing);

            string tmp = newObject.Range.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);

            return newObject;
        }

        public static string GetColumnName(DataTable DT, String ColName)
        {
            string thisColumn = null;
            string thisName = ColName.Replace("\\", "/");

            foreach (DataColumn i in DT.Columns)
            {
                if (i.ColumnName.ToLower() == thisName.ToLower())
                {
                    thisColumn = i.ColumnName;
                }
            }

            return thisColumn;
        }

        #endregion

        #region //string cleanup for worksheet and column names
        static public string CreateValidFormulaName(string name)
        {
            String stringName;

            stringName = "[" + name.Replace("'", "''")
                            .Replace("#", "'#")
                            .Replace("[", "'[")
                            .Replace("]", "']") + "]";

            if (stringName == "[]") { stringName = "0"; }

            return stringName;
        }

        static public string CreateValidWorksheetName(Excel.Workbook workbook, string name, int regressionIteration)
        {
            // Worksheet name cannot be longer than 31 characters.
            int maxlen = 29; 
            System.Text.StringBuilder escapedString;
            String stringName;
            name = regressionIteration.ToString() + " " == "0 " ? name : regressionIteration.ToString() + " " + name;
            if (name.Length <= maxlen)  //
            {
                escapedString = new System.Text.StringBuilder(name);
            }
            else
            {
                escapedString = new System.Text.StringBuilder(name, 0, maxlen, maxlen);
            }

            for (int i = 0; i < escapedString.Length; i++)
            {
                if (escapedString[i] == ':' ||
                    escapedString[i] == '\\' ||
                    escapedString[i] == '/' ||
                    escapedString[i] == '?' ||
                    escapedString[i] == '*' ||
                    escapedString[i] == '[' ||
                    escapedString[i] == ']')
                {
                    escapedString[i] = '_';
                }
            }

            stringName = escapedString.ToString();
            //stringName = regressionIteration.ToString() + " " == "0 " ? stringName : regressionIteration.ToString() + " " + stringName;

            // perform one last check to ensure the name doesn't already exist
            if (GetWorksheet(workbook, stringName) != null)
                stringName = incrementDupWSName(workbook, stringName);         

            return stringName;
        }

        internal static string incrementDupWSName(Excel.Workbook WB, string strName)
        {
            string i = strName.Contains("_") ? strName.Substring(0,strName.LastIndexOf("_") - 2) : strName;
            int x;

            int j = (int.TryParse(WSNameIterations(WB, i), out x)) ? x + 1 : 1;

            return strName + "_" + j.ToString();
        }

        internal static string WSNameIterations(Excel.Workbook WB, string strName)
        {
            int j = 0;
            foreach (Excel.Worksheet WS in WB.Worksheets)
            {
                string nm = WS.Name.Contains("_") ? WS.Name.Substring(0, WS.Name.LastIndexOf("_")) : WS.Name;
                if (nm == strName)
                {
                    string i = WS.Name.Contains("_") ? WS.Name.Substring(WS.Name.LastIndexOf("_")).Replace("_", "") : "0";
                    j = Math.Max(j, int.Parse(i));
                }
            }
            return j.ToString();
        }
        #endregion

        #region //list object helpers

        static void AddListObject(ref Excel.Worksheet WS)
        {
            Excel.Range firstCell = WS.get_Range(WS.UsedRange.get_Address(0, 0, Excel.XlReferenceStyle.xlA1, missing, missing), missing).get_Offset(1, 0);
            //firstCell.Value2 = EnPIResources.dateColName;

            WS.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, firstCell, missing, Excel.XlYesNoGuess.xlYes, missing);
        }
        public static Excel.ListObject GetListObject(Excel.Worksheet WS, string Name)
        {
            foreach (Excel.ListObject LO in WS.ListObjects)
            {
                if (LO.Name == Name) return LO;
            }
            
            return null;
        }
        public static Excel.ListObject GetListObject(Excel.Worksheet WS)
        {
            if (WS.ListObjects.Count == 0)
            {
                AddListObject(ref WS);
            }

            Excel.ListObject thisList = WS.ListObjects.get_Item(1);
            return thisList;

        }

        public static Excel.ListColumn GetListColumn(Excel.ListObject LO, String ColName)
        {
            Excel.ListColumn thisColumn = null;
            if (ColName == null || LO == null)
                return null;

            string thisName = ColName.Replace("\\", "/");

            foreach (Excel.ListColumn i in LO.ListColumns)
            {
                if (i.Name.ToLower() == thisName.ToLower())
                {
                    thisColumn = i;
                }
            }

            //if (thisColumn == null)
            //{
            //    thisColumn = AddListColumn(LO, thisName, Pos);
            //}

            return thisColumn;

        }

        public static Excel.ListRow GetListRow(Excel.ListObject LO)
        {
            Excel.ListRow thisRow = null;
            if (LO == null)
                return null;

            foreach (Excel.ListRow i in LO.ListRows)
            {
                if(i.Index.Equals(1))
                    thisRow = i;
            }
            return thisRow;
        }

        public static string GetListColumnName(Excel.ListObject LO, String ColName)
        {
            Excel.ListColumn thisColumn = null;

            thisColumn = GetListColumn(LO, ColName);

            if (thisColumn == null) return null;

            return thisColumn.Name;
        }

        static public Excel.ListObject newListObject(ref Excel.Worksheet WS, Excel.Range firstCell, params String[] addColumns)
        {
            string tmp = firstCell.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);
            for (int i = addColumns.GetLowerBound(0); i <= addColumns.GetUpperBound(0); i++)
            {
                firstCell = firstCell.get_Resize(firstCell.Rows.Count, firstCell.Columns.Count + 1);
                firstCell.get_Resize(1,1).get_Offset(0, firstCell.Columns.Count - 1).Value2 = addColumns[i];
            }
            tmp = firstCell.get_Address(missing, missing, Excel.XlReferenceStyle.xlA1, missing, missing);

            Excel.ListObject newLO = WS.ListObjects.Add(Excel.XlListObjectSourceType.xlSrcRange, firstCell, missing, Excel.XlYesNoGuess.xlYes, missing);

            return newLO;
        }

        static public Excel.ListColumn AddListColumn(Excel.ListObject LO, String ColName, params Int32[] Pos)
        {
            Excel.ListColumn thisColumn = null;

            if (Pos.Length > 0 && Pos[0] > 0) thisColumn = LO.ListColumns.Add(Pos);
            if (Pos.Length > 0 && Pos[0] < 0) return thisColumn;
            if (Pos.Length == 0 || Pos[0] == 0) thisColumn = LO.ListColumns.Add(missing);

            if(LO.ListColumns[1].Name == "Column1")
                LO.ListColumns[1].Name = "Date";
            for (int i = 1; i < LO.ListColumns.Count; i++)
            {
               
                if (LO.ListColumns[i].Name == ColName)
                    thisColumn.Name = ColName + "1";
                else
                    thisColumn.Name = ColName;
            }
           
            thisColumn.Range.BorderAround(Excel.XlLineStyle.xlLineStyleNone, Excel.XlBorderWeight.xlHairline, Excel.XlColorIndex.xlColorIndexNone, missing);

            return thisColumn;
        }

        static public Excel.ListObject AddSourceSumColumn(Excel.ListObject LO)
        {
            Excel.ListColumn sum = GetListColumn(LO, EnPIResources.totalSourceColName);
            if (sum == null)
                sum = AddListColumn(LO, EnPIResources.totalSourceColName, 0);

            string strTotal = "";
            foreach (Excel.ListColumn col in LO.ListColumns)
            {
                string colNameclean = ExcelHelpers.CreateValidFormulaName(col.Name);
                if (DataHelper.targetTable((object[,])col.Range.Value2) == Constants.COLUMNTAG_DVS)
                {
                    strTotal += colNameclean + "+";
                }
            }
            sum.Range.Formula = "= " + strTotal.Substring(0, strTotal.Length - 1);
            return LO;
        }

        public static void copyFormatting(Excel.ListObject srcObject, Excel.ListObject tgtObject)
        {
            foreach (Excel.ListColumn col in srcObject.ListColumns)
            {
                string tmp = col.Name;
                string tmp2 = tgtObject.ListColumns[col.Index].Name;
                string tmp3 = col.Range.get_Offset(1, 0).get_Resize(1, 1).NumberFormat.ToString();
                if (tgtObject.ListColumns[col.Index] !=  null)
                    tgtObject.ListColumns[col.Index].Range.get_Offset(1,0).NumberFormat = col.Range.get_Offset(1, 0).get_Resize(1, 1).NumberFormat;
            }
        }

        public static void formatGeneral(ref Excel.ListObject LO, bool overwriteDates)
        {
            foreach (Excel.ListColumn LC in LO.ListColumns)
            {
                string curr = LC.Range.get_Offset(1,0).get_Resize(1,1).NumberFormat.ToString();
                if (overwriteDates || !(curr.Contains("m") || curr.Contains("y")) ) 
                    LC.Range.NumberFormat = "@";
            }
        }

        public static void formatRowsinListObject(Excel.ListObject LO, string frow, string lrow, params string[] format)
        {
            string stylename = "Normal";
            int firstrow = 0;
            int lastrow = 0;

            if (!(int.TryParse(frow, out firstrow) && int.TryParse(lrow, out lastrow)))
                return;

            if (format.Length > 0) stylename = format[0];

            for (int i = Math.Min(firstrow,lastrow); i <= Math.Max(lastrow, firstrow); i++)
            {
                LO.ListRows[i].Range.Style = stylename;
            }
        }

        public static void formatPercent(Excel.Workbook WB)
        {
            // set the "Percent" style to have two decimal places
            foreach (Excel.Style style in WB.Styles)
            {
                string tmp = style.NumberFormat;
                if (style.Name == "Percent" && tmp != "0.00%" )
                    style.NumberFormat = "0.00%";
            }
        }
        #endregion

        #region // year helpers
        public static void SetYear(Excel.ListObject LO, Excel.Range Sel, int YearNo)
        {
            Excel.ListColumn yearCol = GetListColumn(LO, EnPIResources.yearColName);

            if (yearCol == null) yearCol = AddListColumn(LO, EnPIResources.yearColName, 1);

            // the selection
            int firstRow = Sel.Row;
            // if there are header rows in the selection or the selection starts above the header row, return 
            // we don't know where we are on the sheet relative to the list object
            if (Sel.ListHeaderRows > 0 || firstRow < LO.HeaderRowRange.Row)
                return;

            // get the corresponding range in the year column
            Excel.Range yearRows = yearCol.Range.get_Offset(firstRow - LO.HeaderRowRange.Row, 0).get_Resize(Sel.Rows.Count, 1);
            yearRows.Value2 = YearNo;

            yearRows.NumberFormat = "@";
        }

        public static void AutoSetYear(Excel.ListObject LO, int interval)
        {
            Excel.Range yearCol = GetListColumn(LO, EnPIResources.yearColName).Range;

            if (yearCol == null) yearCol = AddListColumn(LO, EnPIResources.yearColName, 1).Range;

            int YearNo;

            for (int i = 0; i < yearCol.Rows.Count; i++)
            {
                if (((Excel.ListRow)yearCol.Rows[i + 1,1]).Range.ListHeaderRows == 0)
                {
                    YearNo = int.Parse(Math.Round(((decimal)(i / interval)), 0).ToString());
                    yearCol[i + 1, 1] = YearNo;
                }
            }

            yearCol.ClearFormats();
            yearCol.NumberFormat = "@";
        }


        public static int DataPt(Excel.ListObject LO, string year)
        { 
            int TotDataPt=0;
            TotDataPt = Analytics.getDataPt(rangeTable(LO), year, EnPIResources.yearColName);
            return TotDataPt;
        }



        public static object[] getYears(Excel.ListObject LO)
        {
            Excel.ListObject tmpLO = LO;
            DataTable tmp = rangeTable(LO);
            string yearname = (GetListColumn(LO, EnPIResources.yearColName) != null) ? GetListColumn(LO, EnPIResources.yearColName).Name : "";

            //Commented below sorting funtion to fix defect #13194
            //try 
            //{   // sort on the year column
            //    tmpLO.Sort.SortFields.Clear();
            //    Excel.ListColumn lc = tmpLO.ListColumns.get_Item(yearname);
            //    tmpLO.Sort.SortFields.Add(lc.Range, System.Type.Missing, System.Type.Missing, System.Type.Missing, System.Type.Missing);
            //    tmpLO.Sort.Apply();
            //}
            //catch
            //{
            //}

            object[] res = (yearname != "") ?  Analytics.getYears(rangeTable(tmpLO), yearname) : null;
            
            return res;
        }

        public static object[] getYears(DataTable DT)
        {
            string yearname = GetColumnName(DT, EnPIResources.yearColName) ?? "";

            object[] res = (yearname != "") ? Analytics.getYears(DT, yearname) : null;

            return res;
        }

        public static Excel.Range getYearRange(Excel.ListObject LO, string year)
        {
            Excel.ListColumn LCyear = GetListColumn(LO, EnPIResources.yearColName);
            int first = 0;
            int last = 0;
            int j = 0;

            if (LCyear != null)
            {
                foreach (Excel.Range row in LCyear.Range.Rows)
                {
                    j += 1;
                    if (row.Value2.ToString() == year && first == 0) first = j;
                    if (row.Value2.ToString() == year) last = j;
                }
            }

            if (first != 0)
                return LO.Range.get_Offset(first - 1, 0).get_Resize(last - (first - 1), LO.ListColumns.Count);
            else
                return null;

        }

        public static object[,] getYearArray(Excel.ListObject LO, ArrayList years)
        {
            Excel.ListColumn LCyear = GetListColumn(LO, EnPIResources.yearColName);
            int j = 0;
            object[,] LOyears = null ;
 
            if (LCyear != null)
            {
                foreach (Excel.Range row in LCyear.Range.Rows)
                {
                    string tmp1 = row.Value2.ToString();
                    if (years.Contains(row.Value2.ToString()))
                    {
                        LOyears = DataHelper.arrayAppend(LOyears, LO.ListRows[j].Range.Value2 as object[,]);
                    }
                    j++;
                }
            }

           return LOyears;

        }
        
        #endregion

        #region //worksheet helpers
        // Creates a worksheet
        static public Excel.Worksheet AddWorksheet(Excel.Workbook workbook, string name, params Excel.Worksheet[] aftersheet)
        {
            string nm = CreateValidWorksheetName(workbook, name, 1);
            Excel.Worksheet after = null;
            Excel.Worksheet newSheet;

            if (aftersheet.Length > 0) after = aftersheet[0];

            if (after != null)
                newSheet = (Excel.Worksheet)workbook.Worksheets.Add(missing, after, 1, Excel.XlSheetType.xlWorksheet);
            else
                newSheet = (Excel.Worksheet)workbook.Worksheets.Add(missing, missing, 1, Excel.XlSheetType.xlWorksheet);
            
            newSheet.Name = nm;

            return newSheet;
        }

        // Gets a worksheet by name
        public static Excel.Worksheet GetWorksheet(Excel.Workbook workbook, string name)
        {
            int i;

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ws.Name == name)
                {
                    i = ws.Index;
                    return workbook.Worksheets[i] as Excel.Worksheet;
                }
            }
            return null;
        }

        public static Excel.Worksheet GetWorksheetbyGUID(Excel.Workbook workbook, string sheetguid)
        {
            int i;

            foreach (Excel.Worksheet ws in workbook.Worksheets)
            {
                if (ExcelHelpers.getWorksheetCustomProperty(ws, "SheetGUID") == sheetguid)
                {
                    i = ws.Index;
                    return workbook.Worksheets[i] as Excel.Worksheet;
                }
            }
            return null;
        }

        // Deletes all worksheets marked "very hidden" from the given workbook
        static public void Cleanup(Excel.Workbook WB)
        {
            Excel.Application thisApp = (Excel.Application)WB.Parent;
            thisApp.DisplayAlerts = false;

            foreach (Excel.Worksheet ws in thisApp.ActiveWorkbook.Worksheets)
            {
                if (ws.Visible == Excel.XlSheetVisibility.xlSheetVeryHidden)
                {
                    ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    ws.Delete();
                }
            }

            thisApp.DisplayAlerts = true;
        }

        // writes a string to the bottom of a worksheet, below [offset] blank rows
        public static void writeAppendBottom(ref Excel.Worksheet WS, string[] vals, params int[] offset)
        {
            int start = writeAppendBottomAddress(ref WS, offset);
            WS.get_Range("A" + start.ToString(), missing).Value2 = vals;
        }
        // gets the address of the bottom of the worksheet,  below [offset] blank rows
        public static int writeAppendBottomAddress(Excel.Worksheet WS, params int[] offset)
        {
            return writeAppendBottomAddress(ref WS, offset);
        }
        public static int writeAppendBottomAddress(ref Excel.Worksheet WS, params int[] offset)
        {
            int off = 0;
            if (offset.Length > 0) off = offset[offset.GetLowerBound(0)];
            Excel.Range rng = WS.UsedRange;
            string used = WS.UsedRange.get_Address(missing, missing, Excel.XlReferenceStyle.xlR1C1, missing, missing);
            if (used.IndexOf(":") >= 0) used = used.Substring(used.IndexOf(":"));
            string start;
            start = used.Substring(used.IndexOf("R") + 1);
            start = start.Substring(0, start.IndexOf("C"));

            return (int.Parse(start) + off);
        }
        
        #endregion

        #region //custom property helpers

        static public string getWorksheetCustomProperty(Excel.Worksheet WS, string Property)
        {
            foreach (Excel.CustomProperty var in WS.CustomProperties)
            {
                if (var.Name == Property)
                    return var.Value.ToString();
            }

            return null;
        }

        static public void addWorksheetCustomProperty(Excel.Worksheet WS, string Property, string value)
        {
            foreach (Excel.CustomProperty prop in WS.CustomProperties)
            {
                if (prop.Name == Property)
                    prop.Delete();
            }

            WS.CustomProperties.Add(Property, value);
        }

        static public void copyWorksheetCustomProperties(Excel.Worksheet WSFrom, Excel.Worksheet WSTo)
        {
            foreach (Excel.CustomProperty prop in WSFrom.CustomProperties)
            {
                addWorksheetCustomProperty(WSTo, prop.Name, prop.Value.ToString());
            }

        }
        
        static public string getWorkbookCustomProperty(Excel.Workbook WB, string Property)
        {
            foreach (Office.DocumentProperty var in (Office.DocumentProperties)WB.CustomDocumentProperties)
            {
                if (var.Name == Property)
                    return var.Value.ToString();
            }

            return null;
        }

        static public void addWorkbookCustomProperty(Excel.Workbook WB, string Property, string value)
        {
            // note -- the msoPropertyTypeString has a 255 character limit
            foreach (Office.DocumentProperty var in (Office.DocumentProperties)WB.CustomDocumentProperties)
            {
                if (var.Name == Property)
                    var.Delete();
            }

            ((Office.DocumentProperties)WB.CustomDocumentProperties).Add(Property, false, Office.MsoDocProperties.msoPropertyTypeString, value, missing);
        }       
        
        #endregion

        // Gets a workbook by name
        public static Excel.Workbook GetWorkbook(Excel.Application app, string name)
        {
            if (name == null)
                return null;

            foreach (Excel.Workbook wb in app.Workbooks)
            {   
                if (wb.Name == name || wb.FullName == name)
                {
                    return wb;
                }
            }

            return app.Workbooks.Open(name, Excel.XlUpdateLinks.xlUpdateLinksNever, true, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
        }



    }
}
