using System;
using System.Data;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AMO.EnPI.AddIn.Utilities
{
    public class DataHelper
    {
        // AMO.EnPI.Utilities.DataHelper
        //
        // This class contains the tools to help with data processing for the EnPI Addin. The class does not depend on any Excel
        // classes, so it is compatible with all versions.
        //
        // Enumerations:
        //      EnPITypes           Available computation methods (forecast, backcast, actual, chaining)
        //      EnergySourceTypes   Names of energy sources to look for
        //      VariableTypes       List of types of independent variables
        //      OutputColumns       The output columns to describe the model that will exist in the model data table
        //
        // Public Methods:
        //      
        //
        // Internal Methods:
        //      
        //      
        //      
        private static System.Resources.ResourceManager rsc = new System.Resources.ResourceManager
         ("AMO.EnPI.AddIn.Utilities.EnPIResources", System.Reflection.Assembly.GetExecutingAssembly());

        private static object missing = System.Type.Missing;

        #region // string helpers
        // identifies whether a column should be a dependent variable, an independent variable or excluded
        static public string targetTable(object[,] col)
        {
            string hcol = col[col.GetLowerBound(0), col.GetLowerBound(1)].ToString();
            string tbl = targetTable(hcol);
            if (tbl == "")
            {
                //check to see if it's numerical data; if it is, treat it as another variable
                if (arrayIsNumeric(col, true) && !hcol.ToLower().Contains("date")
                    && !hcol.ToLower().Contains("month")
                    && !hcol.ToLower().Contains("year")
                    && !hcol.ToLower().Contains("week")
                    && !hcol.ToLower().Contains("quarter"))
                    tbl = Constants.TABLENAME_IVS;
            }


            return tbl;
        }
        static public string targetTable(string hcol)
        {
            string tbl = "";

            foreach (Constants.EnergySourceTypes typ
                in System.Enum.GetValues(typeof(Constants.EnergySourceTypes)))
            {
                if (hcol.ToLower().Contains(rsc.GetString(typ.ToString()).ToLower()))
                {
                    tbl = Constants.COLUMNTAG_DVS;
                    return tbl;
                }
            }

            foreach (Constants.VariableTypes typ
                in System.Enum.GetValues(typeof(Constants.VariableTypes)))
            {
                if (hcol.ToLower().Contains(rsc.GetString(typ.ToString())))
                {
                    tbl = Constants.COLUMNTAG_IVS;
                    return tbl;
                }
            }

            //check to see if it's date data
            if (hcol.ToLower().Contains("date")
                || hcol.ToLower().Contains("month")
                || hcol.ToLower().Contains("year")
                || hcol.ToLower().Contains("week")
                || hcol.ToLower().Contains("quarter"))
                tbl = Constants.COLUMNTAG_DATE;

            return tbl;
        }
        // checks to see if the colum name is something like production
        static public bool isProduction(string colname)
        {
            if (colname.ToLower().Contains("product")) return true;

            return false;
        }
        // checks to see if the colum name is something like square feet
        static public bool isBuildingArea(string colname)
        {
            bool ret = false;

            if (colname.ToLower().Contains("sq") &&
               (colname.ToLower().Contains("ft") || colname.ToLower().Contains("feet")))
                ret = true;

            return ret;
        }
        static internal string IndirectAddress(string strRow, string strCol)
        {
            return  "INDIRECT(ADDRESS(" + strRow + ',' + strCol + ",1,TRUE,\"{shtnm}\"))";
        }

        // selects all columns of a table whose rows include the "match" value in the specified column
        // creates an excel range reference for use in a formula for a cell/column
        // The {tblname}  parameter is the table from which the rows for the range are selected
        // The {colname} parameter is the column to search for the {match} value to identify the rows for the range
        // The {match} parameter can be a column name or a value
        // If the {match} parameter is a string, ensure it includes escaped quotes: "\"my string\""
        // if the parameter {match} is a column name, it must include the [#This Row] qualifier 
        static public string RowRangebyMatch(string tblname, string colname, string match, string sheetnm)
        {
            string Ar = "ROW({table}[#Headers]) + IFERROR(MATCH({match}, {table}[[#Data],[{column}]],0), ROWS({table}) )";
            string Ac = "1";
            string Zr = "ROW({table}[#Headers]) + IFERROR(MATCH(OFFSET({match},1,0), {table}[[#Data],[{column}]],0)-1, ROWS({table}) )";
            string Zc = "COUNTA({table}[#Headers])";

            Ar = Ar.Replace("{table}", tblname).Replace("{column}", colname).Replace("{match}", match);
            Ac = Ac.Replace("{table}", tblname).Replace("{column}", colname).Replace("{match}", match);
            Zr = Zr.Replace("{table}", tblname).Replace("{column}", colname).Replace("{match}", match);
            Zc = Zc.Replace("{table}", tblname).Replace("{column}", colname).Replace("{match}", match);

            return IndirectAddress(Ar, Ac).Replace("{shtnm}", sheetnm) + ":" + IndirectAddress(Zr, Zc).Replace("{shtnm}", sheetnm);
        }
        #endregion
        
        #region // array helpers
        static internal object[] arrayResize(string[] arr, int addRows)
        {
            object[] arr2 = arr as object[];
            return arrayResize( arr2, addRows);
        }
        static internal object[] arrayResize(object[] arr, int addRows)
        {
            object[] newArr = new object[arr.GetLength(0) + addRows];

            for (int i = arr.GetLowerBound(0); i <= arr.GetUpperBound(0); i++)
            {
                newArr[i] = arr[i];
            }

            return newArr;
        }
        static internal object[,] arrayResize(object[,] arr, int addRows, int addCols)
        {
            object[,] newArr = new object[arr.GetLength(0) + addRows, arr.GetLength(1) + addCols];

            for (int i= arr.GetLowerBound(0); i <= arr.GetUpperBound(0); i++)
            {
                for (int j = arr.GetLowerBound(1); j <= arr.GetUpperBound(1); j++)
                {
                    newArr[i,j] = arr[i,j];
                }
            }

            return newArr;
        }
        static public double[,] arrayAddIdentity(double[,] arr, int addRows, int addCols)
        {
            double[,] newArr = new double[arr.GetLength(0) + addRows, arr.GetLength(1) + addCols];

            for (int i = arr.GetLowerBound(0); i <= arr.GetUpperBound(0) + addRows; i++)
            {
                for (int j = arr.GetLowerBound(1); j <= arr.GetUpperBound(1) + addCols; j++)
                {
                    newArr[i, j] = (i > arr.GetUpperBound(0) || j > arr.GetUpperBound(1)) ? 1 : arr[i, j];
                }
            }

            return newArr;
        }
        static internal bool arrayIsNumeric(object[] arr, bool skipFirstRow)
        {
            double x;
            int start = arr.GetLowerBound(0);
            if (skipFirstRow) start += 1;

            for (int i = start; i <= arr.GetUpperBound(0); i++)
            {
                if (!double.TryParse(arr[i].ToString(), out x))
                    return false;
            }

            return true;
        }
        static internal bool arrayIsNumeric(object[,] arr, bool skipFirstRow)
        {
            double x;
            int start = arr.GetLowerBound(0);
            if (skipFirstRow) start += 1;

            for (int i = start; i <= arr.GetUpperBound(0); i++)
            {
                for (int j = arr.GetLowerBound(1); j <= arr.GetUpperBound(1); j++)
                {
                if (arr[i,j] == null || !double.TryParse(arr[i, j].ToString(), out x))
                    return false;
                }
            }

            return true;
        }
        
        static internal object[,] arrayAppend(object[,] arr1, object[,] arr2)
        { 
            if (arr1 == null)
                return arr2;
            if (arr2 == null)
                return arr1;

            // will create an array with the smallest dimensions of the two
            // this array will be zero-based
            object[,] newArr = new object[arr1.GetLength(0)+arr2.GetLength(0), Math.Min(arr1.GetLength(1),arr2.GetLength(1))];

            // Rows
            // the offset is there to handle non-zero-based arrays
            int rowOffset1 = arr1.GetLowerBound(0);
            int rowOffset2 = arr2.GetLowerBound(0);

            // Columns
            // the offset is there to handle non-zero-based arrays
            int colOffset1 = arr1.GetLowerBound(1);
            int colOffset2 = arr2.GetLowerBound(1);

            for (int i = 0; i < newArr.GetLength(1); i++)
            {
                for (int r = 0; r < arr1.GetLength(0); r++)
                {
                    newArr[r,i] = arr1[r + rowOffset1, i + colOffset1];
                }
                for (int r = arr1.GetLength(0); r <= newArr.GetUpperBound(0); r++)
                {
                    newArr[r, i] = arr2[r - arr1.GetLength(0) + rowOffset2, i  + colOffset2];
                }
            }

            return newArr;
        }

        static internal object[,] arrayUnion(object[,] arr1, object[,] arr2)
        {
            if (arr1 == null)
                return arr2;
            if (arr2 == null)
                return arr1;

            // will create an array with the smallest dimensions of the two
            // this array will be zero-based
            object[,] newArr = new object[Math.Min(arr1.GetLength(0), arr2.GetLength(0)), arr1.GetLength(1) + arr2.GetLength(1)];

            // Rows
            // the offset is there to handle non-zero-based arrays
            int rowOffset1 = arr1.GetLowerBound(0);
            int rowOffset2 = arr2.GetLowerBound(0);

            // Columns
            // the offset is there to handle non-zero-based arrays
            int colOffset1 = arr1.GetLowerBound(1);
            int colOffset2 = arr2.GetLowerBound(1);

            for (int r = 0; r < newArr.GetLength(0); r++)
            {
                for (int i = 0; i < arr1.GetLength(1); i++)
                {
                    newArr[r, i] = arr1[r + rowOffset1, i + colOffset1];
                }
                for (int j = arr1.GetLength(1); j <= newArr.GetUpperBound(1); j++)
                {
                    newArr[r, j] = arr2[r + rowOffset2, (j - arr1.GetLength(1)) + colOffset2];
                }
            }

            return newArr;
        }

        static internal double[,] dblarrayUnion(double[,] arr1, double[,] arr2)
        {
            if (arr1 == null)
                return arr2;
            if (arr2 == null)
                return arr1;

            // will create an array with the smallest dimensions of the two
            // this array will be zero-based
            double[,] newArr = new double[Math.Min(arr1.GetLength(0), arr2.GetLength(0)), arr1.GetLength(1) + arr2.GetLength(1)];

            // Rows
            // the offset is there to handle non-zero-based arrays
            int rowOffset1 = arr1.GetLowerBound(0);
            int rowOffset2 = arr2.GetLowerBound(0);

            // Columns
            // the offset is there to handle non-zero-based arrays
            int colOffset1 = arr1.GetLowerBound(1);
            int colOffset2 = arr2.GetLowerBound(1);

            for (int r = 0; r < newArr.GetLength(0); r++)
            {
                for (int i = 0; i < arr1.GetLength(1); i++)
                {
                    newArr[r, i] = arr1[r + rowOffset1, i + colOffset1];
                }
                for (int j = arr1.GetLength(1); j <= newArr.GetUpperBound(1); j++)
                {
                    newArr[r, j] = arr2[r + rowOffset2, (j - arr1.GetLength(1)) + colOffset2];
                }
            }

            return newArr;
        }

        static internal double[,] dbl2DArray(double[] arr1)
        {
            // will create a 2D array from a 1D
            int lenI = arr1.Length;

            double[,] newArr = new double[lenI, 1];
            // Rows
            // the offset is there to handle non-zero-based arrays
            int rowOffset = arr1.GetLowerBound(0);


            for (int i = 0; i < lenI; i++)
            {
                newArr[i, 0] = arr1[i + rowOffset];
            }

            return newArr;
        }

        static public double[] objectTOdblArray(object[,] arr1)
        {
            // will create a 2D array from a 1D
            int lenI = arr1.GetLength(0);
            double[] newArr = new double[lenI];

            // Rows
            // the offset is there to handle non-zero-based arrays
            int rowOffset = arr1.GetLowerBound(0);

            for (int i = 0; i < lenI; i++)
            {
                if (!double.TryParse(arr1[i+rowOffset, arr1.GetLowerBound(1)].ToString(), out newArr[i]))
                    newArr[i] = 0;
            }

            return newArr;
        }
        #endregion

        #region //data table helpers

        public static DataTable ConvertToDataTable(object[,] headers, object[,] data)
        {
            if (headers != null && data != null)
            {
                DataTable dt = CreateDataTable(headers, data);
                string cn;

                // write row data into datatable rows
                if (data.GetLength(0) != 0)
                {
                    for (int r = 0; r < data.GetLength(0); r++)
                    {
                        DataRow dr = dt.NewRow();
                        
                        for (int c = 0; c < headers.GetLength(1); c++)
                        {
                            // TO DO: this is where data points can be excluded. If column should be numeric but has text, don't add the row to the table
                            cn = CreateValidColumnName(headers[headers.GetLowerBound(0), c + headers.GetLowerBound(1)].ToString());
                            dr[cn] = data[r + data.GetLowerBound(0), c + data.GetLowerBound(1)] ?? System.Convert.DBNull;
                        }
                        if (CreateValidColumnName(dr[0].ToString()) != dt.Columns[0].ColumnName)
                            dt.Rows.Add(dr);
                    }
                }

                return dt;
            }
            else
                return null;
        }

        internal static System.Type ColumnDataType(int col, object[,] data)
        {
            bool isint = true;
            bool isdbl = true;
            bool allnull = true;
            int testint;
            double testdbl;

            for (int i = data.GetLowerBound(0); i <= data.GetUpperBound(0); i++)
            {
                if (data[i, col] != null)
                {
                    allnull = false;
                    // TO DO: to exclude data points, will need to look for "n/a" or something here so that the column
                    // datatype doesn't get messed up by the text
                    if (!int.TryParse((data[i, col].ToString()), out testint)) isint = false;
                    if (!double.TryParse((data[i, col].ToString()), out testdbl)) isdbl = false;
                }
            }

            if (allnull)
                return System.Type.GetType("System.String");
            else if (isint)
                return System.Type.GetType("System.Int32");
            else if (isdbl)
                return System.Type.GetType("System.Double");                
            else
                return System.Type.GetType("System.String");
        }

        private static DataTable CreateDataTable(object[,] headers, object[,] data)
        {
            DataTable dt = new DataTable();
            DataColumn dc = null;
            string nm = "";

            for ( int i = 0; i < headers.GetLength(1); i++)
            {
                dc = new DataColumn();
                nm = headers[headers.GetLowerBound(0),i + headers.GetLowerBound(1)].ToString();
                dc.ColumnName = CreateValidColumnName(nm);
                dc.DataType = ColumnDataType(i + data.GetLowerBound(1), data);
                dt.Columns.Add(dc);
            }
            return dt;
        }

        // gets the column name non-case-sensitive
        static public string getColumn(DataTable dt, string colName)
        {
            string ret = "";
            foreach(DataColumn dc in dt.Columns)
            {
                if (dc.ColumnName.ToLower() == colName.ToLower()) ret = dc.ColumnName;
            }
            return ret;
        }
        static public DataTable SelectedColumns(DataTable dt, string[] ColumnNames)
        {
            DataTable nt = dt.Copy();

            foreach( DataColumn dc in dt.Columns)
            {
                if (!ColumnNames.Contains(dc.ColumnName))
                {
                    nt.Columns.Remove(dc.ColumnName);
                }
            }

            return nt;
        }
        static public DataTable AddSumColumn(DataTable dt)
        {
            DataColumn dc;
            dc = dt.Columns[EnPIResources.unadjustedTotalColName]; // if a column with this name already exists, reset the formula
            if (dc == null)
            {
                dc = new DataColumn(EnPIResources.unadjustedTotalColName);
                dt.Columns.Add(dc);
            }

            string strTotal = "";
            foreach (DataColumn col in dt.Columns)
            {
                string colNameclean = CreateValidFormulaName(col.ColumnName);
                if (col != dc) strTotal += colNameclean + "+";
            }
            if (strTotal.IndexOf("+") > 0) strTotal = strTotal.Substring(0, strTotal.LastIndexOf("+"));
            
            dc.Expression = strTotal;

            return dt;
        }

        static public DataTable AddSourceSumColumn(DataTable dt)
        {
            DataColumn dc;
            dc = dt.Columns[EnPIResources.unadjustedTotalColName]; // if a column with this name already exists, reset the formula
            if (dc == null)
            {
                dc = new DataColumn(EnPIResources.unadjustedTotalColName);
                dt.Columns.Add(dc);
            }

            string strTotal = "";
            foreach (DataColumn col in dt.Columns)
            {
                string colNameclean = CreateValidFormulaName(col.ColumnName);
                if (col.ExtendedProperties["Tag"] == Constants.COLUMNTAG_DVS)
                {
                    strTotal += colNameclean + "+";
                }
            }
            dc.Expression = strTotal.Substring(0, strTotal.LastIndexOf("+"));

            dc.ExtendedProperties.Add("Tag", Constants.COLUMNTAG_TOTAL);

            return dt;
        }

        // creates a new datatable using a subset of columns from the given datatable,
        // based on whether or not the column extended property "Tag" matches the provided tag name
        static public DataTable CreateVariableTable(DataTable dt, string TableName, string TagName)
        {
            DataTable nt = new DataTable();
            nt.TableName = TableName;
            foreach (DataColumn dc in dt.Columns)
            {
                if (dc.ExtendedProperties["Tag"] == TagName)
                {
                    DataColumn nc = new DataColumn();
                    nc.ColumnName = dc.ColumnName;
                    nc.DataType = System.Type.GetType("System.Double");
                    nc.ExtendedProperties["Tag"] = dc.ExtendedProperties["Tag"];
                    nt.Columns.Add(nc);
                }
            }

            double x;
            foreach (DataRow dr in dt.Rows)
            {
                DataRow nr = nt.NewRow();
                foreach (DataColumn nc in nt.Columns)
                {
                    if (!double.TryParse(dr[nc.ColumnName].ToString(),out x)) x = 0;
                    nr[nc.ColumnName] = x;
                }
                nt.Rows.Add(nr);
            }
            return nt;
        }

        // converts a data column into a 1D double array
        static public double[] dataColumnArray(DataColumn dc)
        {
            double[] dca = new double[dc.Table.Rows.Count];
            double x = 0;

            for (int i = 0; i < dc.Table.Rows.Count; i++)
            {
                dca[i] = (double.TryParse(dc.Table.Rows[i][dc].ToString(), out x)) ? x : 0;
            }

            return dca;
        }

        static public double[] dataColumnArray(DataTable dt, string col)
        {
            double[] dca = new double[dt.Rows.Count];
            double x = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                dca[i] = (double.TryParse(dt.Rows[i][col].ToString(), out x)) ? x : 0;
            }

            return dca;
        }

        // converts a data table into a 2D object array; includes all data
        static public object[,] dataTableArrayObject(DataTable dt)
        {
            object[,] dca = new object[dt.Rows.Count, dt.Columns.Count];

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                   dca[i, j] = dt.Rows[i][j];
                }
            }

            return dca;
        }
        // converts a data table into a 2D object array; includes only columns specified in parameter
        static public object[,] dataTableArrayObject(DataTable dt, params int[] cols)
        {
            object[,] dca = new object[dt.Rows.Count, cols.Length];

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < cols.Length; j++)
                {
                    dca[i, j] = dt.Rows[i][cols[j]];
                }
            }

            return dca;
        }
        // converts a data table into a 2D double array; includes all data
        static public double[,] dataTableArray(DataTable dt)
        {
            double[,] dca = new double[dt.Rows.Count, dt.Columns.Count];
            double x = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    dca[i, j] = (dt.Rows[i][j] != null && double.TryParse(dt.Rows[i][j].ToString(), out x)) ? x : 0;
                }
            }

            return dca;
        }
        // converts a data table into a 2D double array; includes only columns specified in parameter
        static public double[,] dataTableArray(DataTable dt, string[] ivs)
        {
            double[,] dca = new double[dt.Rows.Count, ivs.Length];
            double x = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < ivs.Length; j++)
                {
                    dca[i, j] = (double.TryParse(dt.Rows[i][ivs[j]].ToString(), out x)) ? x : 0;
                }
            }

            return dca;
        }
        static public double[,] dataTableArray(DataTable dt, int[] ivs)
        {
            double[,] dca = new double[dt.Rows.Count, ivs.Length];
            double x = 0;

            for (int i = 0; i < dt.Rows.Count; i++)
            {
                for (int j = 0; j < ivs.Length; j++)
                {
                    dca[i, j] = (double.TryParse(dt.Rows[i][ivs[j]].ToString(), out x)) ? x : 0;
                }
            }

            return dca;
        }
        // returns a 2D object array that contains the column names in a data table
        static public object[,] dataTableHeadersObject(DataTable dt)
        {
            object[,] dth = new object[1,dt.Columns.Count];
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                dth[0,j] = dt.Columns[j].ColumnName;
            }
            return dth;
        }
        // returns a 1D string array that contains the column names in a data table
        static public string[] dataTableHeaders(DataTable dt)
        {
            string[] dth = new string[dt.Columns.Count];
            for (int j = 0; j < dt.Columns.Count; j++)
            {
                dth[j] = dt.Columns[j].ColumnName;
            }
            return dth;
        }
        static public string[] dataTableHeaders(DataTable dt, int[] ivs)
        {
            string[] dth = new string[ivs.Length];
            for (int j = 0; j < ivs.Length; j++)
            {
                dth[j] = dt.Columns[ivs[j]].ColumnName;
            }
            return dth;
        }

        static public string CreateValidColumnName(string name)
        {
            String stringName;

            stringName = name.Replace("\n", " ")
                            //.Replace("\\", "\\\\")
                            .Replace("\t", " ")
                            //.Replace("[", "\\[")
                            //.Replace("]", "\\]")
                            ;
            return stringName;
        }

        static public string CreateValidFormulaName(string name)
        {
            String stringName;

            //If the column name is enclosed in square brackets then any ']' and '\' characters (but not any other characters) in it must be escaped by prepending them with the backslash ("\") character.
            stringName = CreateValidColumnName(name);
            stringName = stringName.Replace("\\", "\\\\")
                            .Replace("]", "\\]")
                            ;

            stringName = "[" + stringName + "]";

            return stringName;
        }

        #endregion

    }
}
