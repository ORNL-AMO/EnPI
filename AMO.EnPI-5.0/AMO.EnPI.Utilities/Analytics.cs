using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Linq;
using System.Text;
using static alglib;


namespace AMO.EnPI.AddIn.Utilities
{
    public class Analytics
    {
        // AMO.EnPI.Utilities.Analytics
        //
        // This class contains the processing components for the EnPI Addin. The class does not depend on any Excel
        // classes, so it is compatible with all versions.
        //
        // Enumerations:
        //      OutputColumns       The output columns to describe the model that will exist in the model data table
        //
        // Public Methods:
        //      findLowStdDev (overloaded)
        //      findHighStdDev (overloaded)
        //      getYears
        //
        // Internal Methods:
        //      ArrayStdDev
        //      

        #region //enumerations
        //public enum EnPITypes { Actual, Forecast, Backcast };

        //public enum EnergySourceTypes { srcElectricity, srcNaturalGas, srcLightFuelOil, srcHeavyFuelOil, srcCoal, srcCoke, srcFurnaceGas, 
        //    srcWoodWaste, srcOtherGas, srcOtherLiquid, srcOtherSolid, srcOtherEnergySource };

        //public enum VariableTypes { ivProduction, ivHDD, ivCDD, ivTemperature, ivHumidity, ivBuildingSqFt, ivOtherVariable };

        enum OutputColumns { ModelNo, ModelValid, IVNames, IVCoefficients, IVses, IVpVals, R2, adjR2, pVal, RMSError, Residual, AIC, Formula };

        #endregion

        private static System.Resources.ResourceManager rsc = new System.Resources.ResourceManager
         ("AMO.EnPI.AddIn.Utilities.EnPIResources", System.Reflection.Assembly.GetExecutingAssembly());
        private static object missing = System.Type.Missing;

        #region //public methods

        public static double findLowStdDev(DataColumn values)
        {
            double[] i = DataHelper.dataColumnArray(values);
            return findLowStdDev(i);
        }
        
        public static double findLowStdDev(double[] values)
        {
            double result = 0;

            double ibar = values.Average();
            double isd = ArrayStdDev(values);

            result = Math.Min(ibar - ((double)3 * isd), values.Min());
            result = ibar - ((double)3 * isd);

            return result;
        }
        public static double findHighStdDev(DataColumn values)
        {
            double[] i = DataHelper.dataColumnArray(values);
            return findHighStdDev(i);
        }
        public static double findHighStdDev(double[] values)
        {
            double result = 0;

            double ibar = values.Average();
            double isd = ArrayStdDev(values);

            result = Math.Max(ibar + ((double)3 * isd), values.Max());
            result = ibar + ((double)3 * isd);

            return result;
        }


        public static int getDataPt(DataTable rawdata, string year, string yearColName)
        {   // get a list of distinct years


            int dataPt = 0;

            foreach (DataRow drow in rawdata.Rows)
            {
                string yr = drow[yearColName].ToString();
                if (yr == year)
                {
                    dataPt += 1;
                }
            }
            return dataPt;
        }


        public static object[] getYears(DataTable rawdata, string yearColName)
        {   // get a list of distinct years

            object[] stryrs = new object[1];
            int j = 0;

            foreach (DataRow drow in rawdata.Rows)
            {
                string yr = drow[yearColName].ToString();
                if (!stryrs.Contains(yr))
                {
                    if (j > stryrs.GetUpperBound(0)) stryrs = DataHelper.arrayResize(stryrs, 1);
                    stryrs[j] = yr;
                    j += 1;
                }
            }
            return stryrs;
        }

        public static object[] getYearsFromDate(DataTable rawdata, string yearColName)
        {   // get a list of distinct years

            object[] stryrs = new object[1];
            int j = 0;

            foreach (DataRow drow in rawdata.Rows)
            {
                string yr = DateTime.FromOADate(Convert.ToDouble(drow[yearColName])).Year.ToString();
                if (!stryrs.Contains(yr))
                {
                    if (j > stryrs.GetUpperBound(0)) stryrs = DataHelper.arrayResize(stryrs, 1);
                    stryrs[j] = yr;
                    j += 1;
                }
            }
            return stryrs;
        }

        #endregion

        #region //internal methods

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

            if (count != 0) stdev = Math.Sqrt((sumSquare - (sum * sum / count)) / (count-1));

            return stdev;
        }
        
        #endregion
    }

    public class EnPIDataSet
    {
        public DataTable ModelData { get; set; }
        public DataTable knownYs { get; set; }
        public DataTable knownXs { get; set; }
        public string WorksheetName { get; set; }
        public string ListObjectName { get; set; }
        public DataTable SourceData;
        public DataTable VariableWarnings;

        public List<string> EnergySourceVariables { get; set; }
        public List<string> ProductionVariables { get; set; }
        public List<string> IndependentVariables { get; set; }
        public List<string> BuildingVariables { get; set; }
        public List<string> Years { get; set; }
        public string BaselineYear { get; set; }
        public string ModelYear { get; set; } 
        public string ReportYear { get; set; } // Added by Suman for SEP changes.

        public EnergySourceCollection EnergySources;

        public EnPIDataSet()
        {
            IndependentVariables = new List<string>();
            BuildingVariables = new List<string>();
            EnergySources = new EnergySourceCollection();

            VariableWarnings = new DataTable("VariableWarnings");
            VariableWarnings.Columns.Add("VariableName");
            VariableWarnings.Columns.Add(EnPIResources.yearColName);
            VariableWarnings.Columns.Add("Warning");
        }


        public void Init()//bool fromRegression
        {
            if (ModelYear != null)
                ExcludeBlanks();

            // set model data
            string yrcol = EnPIResources.yearColName;
            string fltr = yrcol + "='" + ModelYear + "'";
            ModelData = SourceData.Copy();

            if (ModelYear != null)  //replace all data with just the data for the model year
            {
                ModelData.Clear();

                foreach (DataRow dr in SourceData.Select(fltr))
                {
                    ModelData.ImportRow(dr);
                }

                if (ModelData.Rows.Count < Utilities.Constants.MODEL_MIN_DATAPOINTS)
                {
                    DataRow vr = VariableWarnings.NewRow();
                    vr[2] = "Selected model year contains less than " + Utilities.Constants.MODEL_MIN_DATAPOINTS.ToString() + " data points";
                    VariableWarnings.Rows.Add(vr);
                }
            }

            // set knownXs and knownYs
            knownXs = ModelData.Copy();
            knownYs = ModelData.Copy();

            foreach (DataColumn dc in ModelData.Columns)
            {
                if (!EnergySourceVariables.Contains(dc.ColumnName))
                {
                    knownYs.Columns.Remove(dc.ColumnName);
                }

                if (!IndependentVariables.Contains(dc.ColumnName))
                {
                    knownXs.Columns.Remove(dc.ColumnName);
                }
            }

            if (EnergySourceVariables.Contains(EnPIResources.unadjustedTotalColName))
                knownYs = DataHelper.AddSumColumn(knownYs);

            
            if (EnergySourceVariables != null)
            {
                // this will create the collection of energy sources and the list of IVs
                foreach (string col in EnergySourceVariables)
                {
                    EnergySource aSource = new EnergySource(col);
                    aSource.knownXs = knownXs;
                    try
                    {
                        double dcol = Convert.ToDouble(col);
                        string col1 = dcol.ToString("#,###0");
                        aSource.Ys = Ys(col1);
                    }

                    catch
                    {
                        aSource.Ys = Ys(col);
                    }
                    aSource.Combinations = AllCombinations();
                    aSource.AddModels();
                    EnergySources.Add(aSource);
                }
            }

            WriteVariableWarnings();

        }

        public int OutlierCount()
        {
            int ct = 0;
            //double mn;
            string expr;
            string fltr;
            string col;
            string yrcol = EnPIResources.yearColName;
            string blyr = ModelYear;

            foreach (DataColumn dc in SourceData.Columns)
            {
                if (IndependentVariables.Contains(dc.ColumnName))
                {
                    fltr = yrcol + "='" + blyr + "'";
                    col = "[" + dc.ColumnName.Replace("]", "\\]") + "]";
                    expr = "COUNT(" + col + ")";
                    double basect = double.Parse(SourceData.Compute(expr, fltr).ToString());

                    if (basect > 1) //can't compute stdev with one point
                    {
                        expr = "MIN(" + col + ")";
                        double min = double.Parse(SourceData.Compute(expr, fltr).ToString());

                        expr = "MAX(" + col + ")";
                        double max = double.Parse(SourceData.Compute(expr, fltr).ToString());

                        expr = "AVG(" + col + ")";
                        double mean = double.Parse(SourceData.Compute(expr, fltr).ToString());

                        expr = "STDEV(" + col + ")"; // standard deviation computed using (n-1) method
                        double stdev = double.Parse(SourceData.Compute(expr, fltr).ToString());

                        double low = Math.Min(mean - 3 * stdev, min);
                        double high = Math.Max(mean + 3 * stdev, max);

                        expr = col + " < " + low.ToString() + " OR " + col + " > " + high.ToString();
                        ct += SourceData.Select(expr).Count();
                    }
                }
            }
            return ct;
        }

        internal void ExcludeBlanks()
        {
            bool modelnulls = false;
            bool othernulls = false;
            string fltr;
            string col;
            string yrcol = EnPIResources.yearColName;
            string blyr = ModelYear;

            foreach (DataColumn dc in SourceData.Columns)
            {
                if (IndependentVariables.Contains(dc.ColumnName) || EnergySourceVariables.Contains(dc.ColumnName))
                {
                    col = "[" + dc.ColumnName.Replace("]", "\\]") + "]";
                    // remove rows with missing values from model year
                    fltr = yrcol + "='" + blyr + "' and " + col + " is null";
                    foreach (DataRow dr in SourceData.Select(fltr))
                    {
                        SourceData.Rows.Remove(dr);
                        modelnulls = true;
                    }
                    // remove rows with missing values from other years 
                    fltr = yrcol + "<>'" + blyr + "' and " + col + " is null";
                    foreach (DataRow dr in SourceData.Select(fltr))
                    {
                        SourceData.Rows.Remove(dr);
                        othernulls = true;
                    }

                }
                if (modelnulls)
                {
                    DataRow vr = VariableWarnings.NewRow();
                    vr[2] = "Rows with blank values were excluded from the model";
                    VariableWarnings.Rows.Add(vr);
                }
                if (othernulls)
                {
                    DataRow vr = VariableWarnings.NewRow();
                    vr[2] = "Rows with blank values were excluded from the results";
                    VariableWarnings.Rows.Add(vr);
                }
                    
            }
        }

        internal void WriteVariableWarnings()
        {

            double mn;
            string expr;
            string fltr;
            string col;
            string yrcol = EnPIResources.yearColName;
            string blyr = ModelYear;

            foreach (DataColumn dc in SourceData.Columns)
            {
                if (IndependentVariables.Contains(dc.ColumnName))
                {
                    fltr = yrcol + "='" + blyr + "'";
                    col = "[" + dc.ColumnName.Replace("]", "\\]") + "]";
                    expr = "MIN(" + col + ")";
                    double min = double.Parse( SourceData.Compute(expr, fltr).ToString() );

                    expr = "MAX(" + col + ")";
                    double max = double.Parse( SourceData.Compute(expr, fltr).ToString() );

                    expr = "AVG(" + col + ")";
                    double mean = double.Parse(SourceData.Compute(expr, fltr).ToString());

                    expr = "COUNT(" + col + ")"; 
                    double ct = double.Parse(SourceData.Compute(expr, fltr).ToString());

                    if (ct > 1) //stdev method fails with only one row
                    {
                        expr = "STDEV(" + col + ")"; // standard deviation computed using (n-1) method
                        double stdev = double.Parse(SourceData.Compute(expr, fltr).ToString());

                        double low = Math.Min(mean - 3 * stdev, min);
                        double high = Math.Max(mean + 3 * stdev, max);

                        expr = "AVG(" + col + ")";
                        foreach (string yr in Years)
                        {
                            fltr = "[" + yrcol + "]='" + yr + "'";
                            mn = double.Parse(SourceData.Compute(expr, fltr).ToString());
                            if (mn < low || mn > high)
                            {
                                DataRow dr = VariableWarnings.NewRow();
                                dr[0] = dc.ColumnName;
                                dr[1] = yr;
                                VariableWarnings.Rows.Add(dr);
                            }
                        }
                    }
                }
            }
        }

        public double[,] Xs()
        {
            return DataHelper.dataTableArray(knownXs);
        }

        public double[] Ys(string col)
        {
            if (knownYs != null)
                return DataHelper.dataColumnArray(knownYs, col);
            else
                return null;
        }

        public DataTable PredictorRange()
        {
            DataTable pr = new DataTable("PredicterRange");
            pr.Columns.Add("VariableName", System.Type.GetType("System.String"));
            pr.Columns.Add("LowValue", System.Type.GetType("System.Double"));
            pr.Columns.Add("HighValue", System.Type.GetType("System.Double"));

            foreach (DataColumn dc in knownXs.Columns)
            {
                DataRow dr = pr.NewRow();
                dr[0] = dc.ColumnName;
                dr[1] = Analytics.findLowStdDev(dc);
                dr[2] = Analytics.findHighStdDev(dc);
                pr.Rows.Add(dr);
            }

            return pr;
        }

        public List<string[]> AllCombinations()
        {
            int K = IndependentVariables.Count();
            string[] c; 
            List<string[]> A = new List<string[]>();

            for (int k = 1; k <= K; k++)
            {
                Combination C = new Combination(K, k);
                // loop through all possible combinations of k independent variables
                for (int m = 0; m < (int)Combination.Choose(K, k); m++)
                {
                    int[] ivs = C.Element(m).ToArray() as int[];
                    c = new string[ivs.Count()];

                    for (int i = 0; i < ivs.Count(); i++)
                    {
                        c[i] = IndependentVariables[ivs[i]];
                    }

                    A.Add(c);
                }
            }

            return A;
        }
    }

    public class EnergySource
    {
        public string Name { get; set; }
        public DataTable knownXs { get; set; }
        public double[] Ys { get; set; }

        public List<string[]> Combinations;
        public ModelCollection Models;

        public EnergySource()
        {
            Name = "";
            Models = new ModelCollection();
        }
        public EnergySource(string NameIn)
        {
            Name = NameIn;
            Models = new ModelCollection();
        }
        public EnergySource(string Name, double[,] Xs, double[] Ys, List<string[]> Combinations)
        {
            try
            {
                Models = new ModelCollection();

                foreach (string[] Variables in Combinations)
                {
                    Model model = Models.New();
                    model.Xs = Xs;
                    model.Ys = Ys;
                    model.VariableNames = Variables;
                    model.Run();
                }
            }
            catch
            {
            }
        }   

        public void AddModels()
        {
            try
            {
                Models.Clear();

                foreach (string[] Variables in Combinations)
                {
                    Model model = Models.New();
                    model.Xs = DataHelper.dataTableArray(knownXs, Variables);
                    model.Ys = Ys;
                    model.VariableNames = Variables;
                    model.Run();
                }
                //Models to public ModelCollection
            }
            catch
            {
                throw;
            }
        }

        public Model BestModel()
        {
            double r = 0;
            Model n = null;

            foreach (Model model in Models)
            {
                if (model.Valid())
                {
                    if (model.AdjustedR2() > r)
                    {
                        r = model.AdjustedR2();
                        n = model;
                    }
                }
            }

            // no valid model found - select the model with the highest adjusted R2 from all models
            if (n == null)
            {
                foreach (Model model in Models)
                {
                    if (model.AdjustedR2() > r)
                    {
                        r = model.AdjustedR2();
                        n = model;
                    }
                }
            }

            if (n == null)
            {
                n = Models.Item(0);
            }

            return n;
        }

        public double[,] Xs()
        {
            double[,] res = (knownXs != null) ? DataHelper.dataTableArray(knownXs) : null;
            return res;
        }
    }

    public class EnergySourceCollection : System.Collections.CollectionBase
    {
        public void Add(EnergySource aSource)
        {
            List.Add(aSource);
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

        public EnergySource Item(int Index)
        {
            return (EnergySource)List[Index];
        }

    }

    [Serializable]
    public class Model
    {
        public int ModelNumber { get; set; }
        public double[] Ys { get; set; }
        public double[,] Xs { get; set; }
        public string[] VariableNames { get; set; }

        public double RMSError { get; set; }

        public double[] Coefficients { get; set; }

        public Model()
        {
            ModelNumber = 0;
            Ys = null;
            Xs = null;
            VariableNames = null;
            RMSError = 0;
            Coefficients = null;
        }

        protected bool Equals(Model other)
        {
            return ModelNumber == other.ModelNumber && RMSError.Equals(other.RMSError) && Equals(Coefficients, other.Coefficients);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            return obj.GetType() == typeof(Model) && Equals((Model) obj);
        }

        public override int GetHashCode() => (VariableNames != null ? VariableNames.GetHashCode() : 0);

        public static bool operator ==(Model left, Model right)
        {
            return Equals(left, right);
        }

        public static bool operator !=(Model left, Model right)
        {
            return !Equals(left, right);
        }
        
        public Model(int ModelNumber, double[] Ys, double[,] Xs, string[] VariableNames)
        {
            RMSError = 0;
            Coefficients = null;

            // run LLS
            double[] c;
            lsfitreport rep;
            try
            {
                lsfitlinear(Ys, Xsplusone(), out var info, out c, out rep);
            }
            catch
            {
                throw;
            }

            Coefficients = c;
            RMSError = rep.rmserror;
        }

        public void Run() //double[] Ys, double[,] Xs, string[] VariableNames)
        {
            RMSError = 0;
            Coefficients = null;

            if (Ys != null && Xs != null)
            {
                // run LLS
                double[] c;
                lsfitreport rep;
                try
                {
                    lsfitlinear(Ys, Xsplusone(), out var info, out c, out rep);
                }
                catch
                {
                    throw;
                }

                Coefficients = c;
                RMSError = rep.rmserror;
            }
        }
        
        public int N()
        {
            return Ys.Count();
        }

        public int df()
        {
            return N() - k() - 1;
        }

        public int k()
        {
            return VariableNames.Count();
        }

        public double TotalSS()
        {
            // compute total sum of squares
            var ybar = Ys.Average();
            double sst = 0;
            for (int i = Ys.GetLowerBound(0); i <= Ys.GetUpperBound(0); i++)
            {
                sst += Math.Pow(Ys[i] - ybar, 2);
            }

            return sst;
        }

        public double ResidualSS ()
        {
            return ( N() * Math.Pow( RMSError, 2));
        }

        public double R2()
        {
            return (1 - (ResidualSS() / TotalSS()));
        }

        public double AdjustedR2()
        {
             return (1 - (((1 - R2()) * (N() - 1)) / (N() - k() - 1)));

        }

        public double F()
        {
            return ( (R2() / k()) / ((1 - R2()) / (N() - k() - 1)));
        }

        public double ModelPValue()
        {
            double modelP = 0;
            var modelF = F();
            if (modelF < 0) modelF = 0;

            try
            {
                modelP = fcdistribution(N() - df() - 1, df(), modelF);

            }
            catch (alglibexception e)
            {
            }
            return modelP;
        }

        public bool Valid()
        {
            // Model validity criteria, from the SEP M&V protocol:
            // The model p-value must be less than 0.1
            // All variables must have p-values less than 0.2
            // At least one variable must have a p-value of less than 0.1
            // The R2 value must be greater than 0.5

            var ps = PValues();
            var varsvalid = true;
            var varlowexists = false;

            for (var i = 0; i < ps.Count(); i++)
            {
                if (ps[i] <= Constants.PVALUE_THRESHOLD)
                    varlowexists = true;
                if (ps[i] > Constants.PVALUE_HIGH)
                    varsvalid = false;
            }

            if (!varlowexists)
                return false;

            if (!varsvalid)
                return false;

            if (ModelPValue() > Constants.PVALUE_THRESHOLD)
                return false;

            return !(R2() < Constants.R2VALUE_MIN);
        }

        public string Formula()
        {
            var formula = "";
            var offset = Coefficients.GetLowerBound(0) - VariableNames.GetLowerBound(0);
            for (var i = Coefficients.GetLowerBound(0); i < Coefficients.GetUpperBound(0); i++)
            {
                 formula += "(" + Coefficients[i].ToString("0.000000000000") + " * " + ExcelHelpers.CreateValidFormulaName(VariableNames[i - offset]) + ") + ";
                // formula += "(" + Coefficients[i].ToString() + " * " + ExcelHelpers.CreateValidFormulaName(VariableNames[i - offset]) + ") + ";
            }

            formula += Coefficients[Coefficients.GetUpperBound(0)].ToString("0.00");

            return formula;
        }

        public double[,] Xsplusone()
       {
           return DataHelper.arrayAddIdentity(Xs, 0, 1); // add on a column of ones for the intercept
       }

        public double[] PredictedYs()
       {            // compute the predicted ys
           var yhat = new double[N()];
           var xs = Xsplusone();
           var c = Coefficients;

           for (var i = 0; i < N(); i++)
           {
               yhat[i] = 0;
               for (var j = 0; j < k() + 1; j++)
               {
                   yhat[i] += xs[i, j] * c[j];
               }
           }

           return yhat;
       }

        public double[,] CovarianceMatrix()
       {
           // compute the coefficient covariance matrix
           var twodYs = DataHelper.dbl2DArray(Ys);
           var XYs = DataHelper.dblarrayUnion(Xs, twodYs);
           double[,] cov;
            linearmodel lm;
            lrreport rpt;

           try
           {
               lrbuild(XYs, N(), k(), out var info, out lm, out rpt);
               cov = rpt.c;
           }
           catch
           {
               throw;
           }

           return cov;
       }

        public double[] StandardErrors()
        {
           // compute the x std errors and p-values
           var cov = CovarianceMatrix();
           var se = new double[k()];

           if (cov.GetLength(0) > 0 && cov.GetLength(1) > 0)
           {
               for (var j = 0; j < k(); j++)
               {
                   se[j] = Math.Sqrt(cov[j, j]);
               }
           }

           return se;
       }

        public double[] PValues()
       {
           var c = Coefficients;
           var cov = CovarianceMatrix();
           var se = StandardErrors();
           var pv = new double[k()];

           if (cov.GetLength(0) > 0 && cov.GetLength(1) > 0)
           {
               for (var j = 0; j < k(); j++)
               {
                   se[j] = Math.Sqrt(cov[j, j]);
                   try
                   {
                       pv[j] = 2 * (1 - studenttdistribution(df(), Math.Abs(c[j] / se[j])));
                   }
                   catch
                   {
                       // ignored
                   }
               }
           }

           return pv;
       }
        
        public string AICFormula()
        {
            return "";
        }

        //Added By Suman for SEP Validation changes
        public string[] SEPValidationCheck()
        {
            var sepChk = new string[k()];
            for (var cnt = 0; cnt < sepChk.Length; cnt++)
            {
                if (Valid() == true)
                {
                    sepChk[cnt] = "Pass";
                }
                else
                {
                    sepChk[cnt] = "Fail";
                }
            }
            return sepChk;
        }
    }

    public class ModelCollection : System.Collections.CollectionBase
    {
        public void Add(Model aModel)
        {
            List.Add(aModel);
        }

        public Model New()
        {
            var aModel = new Model();
            var i = List.Add(aModel);
            aModel.ModelNumber = i + 1;

            return aModel;
        }

        public void Remove(int index)
        {
            try
            {
                List.RemoveAt(index);
            }
            catch (Exception)
            {
                // ignored
            }
        }

        public Model Item(int Index)
        {
            return (Model)List[Index];
        }

        public int IndexOf(int ModelNo)
        {
            if (ModelNo <= 0) throw new ArgumentOutOfRangeException(nameof(ModelNo));
            for (var i = 0; i < List.Count; i++)
            {
                if (Item(i).ModelNumber == ModelNo) return i;
            }
            return -1;
        }

        public ArrayList ModelSort()
        {
            var sorter = new R2SortHelper();
            InnerList.Sort(sorter);
            return InnerList;
        }

        private class R2SortHelper : System.Collections.IComparer
        {
            public int Compare(object x, object y)

            {
                if (x != y )
                {
                    var m1 = (x as Model).R2() + (x as Model).Valid().GetHashCode();
                    var m2 = (y as Model).R2() + (y as Model).Valid().GetHashCode();

                    if (m1 > m2)
                        return -1;
                    else if (m1 < m2)
                        return 1;
                    
                }
                return 0;
                
            }
        }

     }


}
