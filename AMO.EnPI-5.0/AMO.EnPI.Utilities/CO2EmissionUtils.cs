using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Reflection;

namespace AMO.EnPI.AddIn.Utilities
{

    public enum ComboBoxFillType
    {
        EnergySource=1,
        FuelType
    };
    public sealed class CO2EmissionUtils
    {
        static string xmlFileName = "AMO.EnPI.AddIn.Utilities.CO2EmissionConstants.xml";

        public static DataTable GetCO2Emissions()
        {
            try
            {
                
                System.IO.Stream xmlStream = Assembly.GetExecutingAssembly().GetManifestResourceStream(xmlFileName);
                xmlStream.Position = 0;
                DataSet ds = new DataSet();
                ds.ReadXml(xmlStream);
                if (ds.Tables.Count > 0)
                    return ds.Tables[0];
                else
                    return null;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
