using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AMO.EnPI.AddIn.Utilities
{
    public sealed class Constants
    {
        #region // these are settings and might be moved elsewhere so that they can be user-adjusted
        public const bool MODEL_TOTAL = false;  // set to true to perform regression modeling on total energy
        public const double PVALUE_THRESHOLD = 0.1;
        public const double PVALUE_HIGH = 0.2;
        public const double R2VALUE_MIN = 0.5;
        public const int MODEL_MIN_DATAPOINTS = 10;

        #endregion

        public enum EnPITypes { Actual, Forecast, Backcast };
        public enum EnergySourceTypes
        {
            srcElectricity, srcNaturalGas, srcLightFuelOil, srcHeavyFuelOil, srcCoal, srcCoke, srcFurnaceGas,
            srcWoodWaste, srcOtherGas, srcOtherLiquid, srcOtherSolid, srcOtherEnergySource
        };

        public enum VariableTypes { ivProduction, ivHDD, ivCDD, ivTemperature, ivHumidity, ivBuildingSqFt, ivOtherVariable };

        public enum ModelOutputColumns { ModelNo, ModelValid, IVNames,IVSEPValChk, IVCoefficients, IVses, IVpVals, R2, adjR2, pVal, RMSError, Residual, AIC, Formula };

        // Connection string used by OLEDB to pull data from other Excel files
        public const string EXCEL_CONNSTRING = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES\";";

        public const string WSPROP_SOURCE = "Sources";
        public const string WSPROP_SQL = "TableSQL";
        public const string WSPROP_FILES = "Files";
        public const string WSPROP_VARS = "Variables";
        public const string WSPROP_BLDG = "BuildingSF";
        public const string WSPROP_BASELINE = "BaselineYear";
        public const string WSPROP_YEAR = "ModelYear";
        public const string WSPROP_REPORTYEAR = "ReportYear";
        public const string WSPROP_PRODUCTION = "Production";
        public const string WS_ISENPI = "HasEnPITables";
        public const string WS_ROLLUP = "IsValidRollupSource";
        public const string WS_SRCFILE = "SourceFileName";
        public const string CB_TABLES = "CheckedTables";
        public const string CB_FILES = "CheckedFiles";
        
        public const int CHART_HEIGHT = 244;
        public const int CHART_WIDTH = 358;
        public const string COLUMNTAG_DATE = "Date";
        public const string COLUMNTAG_DVS = "Sources";
        public const string COLUMNTAG_IVS = "Variables";
        public const string COLUMNTAG_TOTAL = "TOTAL";

        public const string TABLETAG_MODEL = "Model";
        public const string TABLENAME_BESTMODEL = "BestModels";
        public const string TABLENAME_DVS = "DependentVariables";
        public const string TABLENAME_IVS = "IndependentVariables";
        public const string TABLENAME_RAWDATA = "RawData";
        public const string TABLENAME_SUMMARY = "Summary";

        public const string SOURCE_PURCHASED_ELECTRICITY = "Purchased Electricity";
        public const string SOURCE_PURCHASED_FUEL = "Purchased Fuel";
        public const string SOURCE_PURCHASED_STEAM = "Purchased Steam";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION = "Purchased Chilled Water (Absorption Chiller)";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ENGINE = "Purchased Chilled Water (Engine-driven Compressor)";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC = "Purchased Chilled Water (Electric-driven Compressor)";
        public const string SOURCE_PURCHASED_COMPRESSED_AIR = "Purchased Compressed Air";
        public const string SOURCE_ELECTRICITY_SOLD = "Electricity Sold";
        public const string SOURCE_STEAM_SOLD = "Steam Sold";

        public const string UNITS_MMBTU = "MMBTU";
        public const string UNITS_KWH = "kWh";
        public const string UNITS_MWH = "MWh";
        public const string UNITS_GWH = "GWh";
        public const string UNITS_KJ = "kJ";
        public const string UNITS_MJ = "MJ";
        public const string UNITS_GJ = "GJ";
        public const string UNITS_TJ = "TJ";
        public const string UNITS_THERMS = "Therms";
        public const string UNITS_DTH = "DTh";
        public const string UNITS_KCAL = "KCal";
        public const string UNITS_GCAL = "GCal";
        public const string UNITS_OTHER = "Other";
        public const string UNITS_SCF = "SCF";
        public const string UNITS_CCF = "CCF";
        public const string UNITS_MCF = "MCF";
        public const string UNITS_M3 = "M3";
        public const string UNITS_GALLON = "Gallon";
        public const string UNITS_BBL = "BBL";
        public const string UNITS_LB = "Lb";
        public const string UNITS_SHORT_TONS = "Short tons";
        public const string UNITS_LONG_TONS = "Long tons";
        public const string UNITS_KG = "Kg";
        public const string UNITS_METRIC_TONS = "Metric tons";
        public const string UNITS_BDST = "BDST (Biomass)";
        public const string UNITS_TON_HOUR = "ton-hour";
        public const string UNITS_GAL_DEG_F = "gal ᵒF";
        public const string UNITS_FT3 = "ft3 (at 100 psi, motor driven compressor)";

        public const string FUEL_TYPE_NATURAL_GAS = "Natural Gas";
        public const string FUEL_TYPE_BLAST_FURNACE = "Blast-Furnace gas";
        public const string FUEL_TYPE_COKE_OVEN = "Coke-oven gas";
        public const string FUEL_TYPE_LPG = "Liquefied Petroleum Gas (LPG)";
        public const string FUEL_TYPE_PROPANE = "Propane (vapor/gas)";
        public const string FUEL_TYPE_BUTANE = "Butane (vapor/gas)";
        public const string FUEL_TYPE_ISOBUTANE = "Isobutane (vapor/gas)";
        public const string FUEL_TYPE_LANDFILL = "Landfill gas";
        public const string FUEL_TYPE_OIL_GASSES = "Oil Gases";

        public const string FUEL_TYPE_LIQ_PROPANE = "Propane (liquid)";
        public const string FUEL_TYPE_LIQ_BUTANE = "Butane (liquid)";
        public const string FUEL_TYPE_LIQ_ISOBUTANE = "Isobutane (liquid)";
        public const string FUEL_TYPE_PENTANE = "Pentane";
        public const string FUEL_TYPE_ETHYLENE = "Ethylene";
        public const string FUEL_TYPE_PROPYLENE = "Propylene";
        public const string FUEL_TYPE_BUTENE = "Butene";
        public const string FUEL_TYPE_PENTENE = "Pentene";
        public const string FUEL_TYPE_BENZENE = "Benzene";
        public const string FUEL_TYPE_TOLUENE = "Toluene";
        public const string FUEL_TYPE_XYLENE = "Xylene";
        public const string FUEL_TYPE_METHYL_ALCOHOL = "Methyl alcohol";
        public const string FUEL_TYPE_ETHYL_ALCOHOL = "Ethyl alcohol";
        public const string FUEL_TYPE_1_FUEL_OIL = "#1 Fuel Oil";
        public const string FUEL_TYPE_2_FUEL_OIL = "#2 Fuel Oil";
        public const string FUEL_TYPE_4_FUEL_OIL = "#4 Fuel Oil";
        public const string FUEL_TYPE_5_FUEL_OIL = "#5 Fuel Oil";
        public const string FUEL_TYPE_6_FUEL_OIL_LOW = "#6 Fuel Oil (Low sulfur)";
        public const string FUEL_TYPE_6_FUEL_OIL_HIGH = "#6 Fuel Oil (High sulfur)";
        public const string FUEL_TYPE_CRUDE = "Crude petroleum";
        public const string FUEL_TYPE_GASOLINE = "Gasoline";
        public const string FUEL_TYPE_KEROSENE = "Kerosene";
        public const string FUEL_TYPE_GAS_OIL = "Gas oil";
        public const string FUEL_TYPE_LNG = "Liquefied Natural Gas (LNG)";

        public const string FUEL_TYPE_COAL = "Coal";
        public const string FUEL_TYPE_COKE = "Coke";
        public const string FUEL_TYPE_PEAT = "Peat";
        public const string FUEL_TYPE_WOOD = "Wood";
        public const string FUEL_TYPE_BIOMASS = "Biomass";
        public const string FUEL_TYPE_BLACK_LIQUOR = "Black liquor";
        public const string FUEL_TYPE_SCRAP_TIRES = "Scrap tires";
        public const string FUEL_TYPE_SULFUR = "Sulfur";



        //For MMBtu conversion
        public const string UNITS_MMBTU_VALUE_MMBTU = "1";
        public const string UNITS_KWH_VALUE_MMBTU = "0.003412142";
        public const string UNITS_MWH_VALUE_MMBTU = "3.412142";
        public const string UNITS_GWH_VALUE_MMBTU = "3412.142";
        public const string UNITS_KJ_VALUE_MMBTU = "0.0000009478171";
        public const string UNITS_MJ_VALUE_MMBTU = "0.0009478171";
        public const string UNITS_GJ_VALUE_MMBTU = "0.9478171";
        public const string UNITS_TJ_VALUE_MMBTU = "947.8171";
        public const string UNITS_THERMS_VALUE_MMBTU = "0.1";
        public const string UNITS_DTH_VALUE_MMBTU = "1";
        public const string UNITS_KCAL_VALUE_MMBTU = "0.000003968321";
        public const string UNITS_GCAL_VALUE_MMBTU = "3.968321";
        public const string UNITS_BDST_VALUE_MMBTU = "18.00";
        public const string UNITS_TON_HOUR_VALUE_MMBTU = "0.012";
        public const string UNITS_GAL_DEG_F_VALUE_MMBTU = "0.00000824";
        public const string UNITS_FT3_VALUE_MMBTU = "0.0000328";
        //For GJ conversion
        public const string UNITS_MMBTU_VALUE_GJ = "1.05505588";
        public const string UNITS_KWH_VALUE_GJ = "0.0036";
        public const string UNITS_MWH_VALUE_GJ = "3.600000464";
        public const string UNITS_GWH_VALUE_GJ = "3600.000464";
        public const string UNITS_KJ_VALUE_GJ = "0.000001";
        public const string UNITS_MJ_VALUE_GJ = "0.001";
        public const string UNITS_GJ_VALUE_GJ = "1";
        public const string UNITS_TJ_VALUE_GJ = "1000";
        public const string UNITS_THERMS_VALUE_GJ = "0.105505588";
        public const string UNITS_DTH_VALUE_GJ = "1.055055875";
        public const string UNITS_KCAL_VALUE_GJ = "0.000004186800386";
        public const string UNITS_GCAL_VALUE_GJ = "4.186800386";
        public const string UNITS_BDST_VALUE_GJ = "0";//GJ Conversion?
        public const string UNITS_TON_HOUR_VALUE_GJ = "0";//GJ Conversion?
        public const string UNITS_GAL_DEG_F_VALUE_GJ = "0";//GJ Conversion?
        public const string UNITS_FT3_VALUE_GJ = "0";//GJ Conversion?

        public const string UNITS_OTHER_VALUE = "0";
        public const double UNITS_SCF_VALUE = 1;
        public const double UNITS_CCF_VALUE = 100;
        public const double UNITS_MCF_VALUE = 1000;
        public const double UNITS_M3_VALUE = 35.31467;
        public const double UNITS_GALLON_VALUE = 1;
        public const double UNITS_BBL_VALUE = 42;
        public const double UNITS_LB_VALUE = 1;
        public const double UNITS_SHORT_TONS_VALUE = 2000;
        public const double UNITS_LONG_TONS_VALUE = 2240;
        public const double UNITS_KG_VALUE = 2.20462;
        public const double UNITS_METRIC_TONS_VALUE = 2204.62;
        
        public const string SOURCE_PURCHASED_ELECTRICITY_VALUE = "3";
        public const string SOURCE_PURCHASED_FUEL_VALUE = "1";
        public const string SOURCE_PURCHASED_STEAM_VALUE = "1.33";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ABSORPTION_VALUE = "1.25";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ENGINE_VALUE = "0.83";
        public const string SOURCE_PURCHASED_CHILLED_WATER_ELECTRIC_VALUE = "0.24";
        public const string SOURCE_PURCHASED_COMPRESSED_AIR_VALUE = "1";
        public const string SOURCE_ELECTRICITY_SOLD_VALUE = "-1";
        public const string SOURCE_STEAM_SOLD_VALUE = "-1";


        public const double FUEL_TYPE_NATURAL_GAS_VALUE = 0.001027;
        public const double FUEL_TYPE_BLAST_FURNACE_VALUE = 0.00009;
        public const double FUEL_TYPE_COKE_OVEN_VALUE = 0.00059;
        public const double FUEL_TYPE_LPG_VALUE = 0.0027;
        public const double FUEL_TYPE_PROPANE_VALUE = 0.002516;
        public const double FUEL_TYPE_BUTANE_VALUE = 0.00328;
        public const double FUEL_TYPE_ISOBUTANE_VALUE = 0.0031;
        public const double FUEL_TYPE_LANDFILL_VALUE = 0.0006;
        public const double FUEL_TYPE_OIL_GASSES_VALUE = 0.0007;

        public const double FUEL_TYPE_LIQ_PROPANE_VALUE = 0.09169;
        public const double FUEL_TYPE_LIQ_BUTANE_VALUE = 0.102032;
        public const double FUEL_TYPE_LIQ_ISOBUTANE_VALUE = 0.10375;
        public const double FUEL_TYPE_PENTANE_VALUE = 0.1108;
        public const double FUEL_TYPE_ETHYLENE_VALUE = 0.012072;
        public const double FUEL_TYPE_PROPYLENE_VALUE = 0.017473;
        public const double FUEL_TYPE_BUTENE_VALUE = 0.023008;
        public const double FUEL_TYPE_PENTENE_VALUE = 0.028693;
        public const double FUEL_TYPE_BENZENE_VALUE = 0.028057;
        public const double FUEL_TYPE_TOLUENE_VALUE = 0.03354;
        public const double FUEL_TYPE_XYLENE_VALUE = 0.03912;
        public const double FUEL_TYPE_METHYL_ALCOHOL_VALUE = 0.006492;
        public const double FUEL_TYPE_ETHYL_ALCOHOL_VALUE = 0.011968;
        public const double FUEL_TYPE_1_FUEL_OIL_VALUE = 0.1374;
        public const double FUEL_TYPE_2_FUEL_OIL_VALUE = 0.1396;
        public const double FUEL_TYPE_4_FUEL_OIL_VALUE = 0.1451;
        public const double FUEL_TYPE_5_FUEL_OIL_VALUE = 0.1488;
        public const double FUEL_TYPE_6_FUEL_OIL_LOW_VALUE = 0.1524;
        public const double FUEL_TYPE_6_FUEL_OIL_HIGH_VALUE = 0.149;
        public const double FUEL_TYPE_CRUDE_VALUE = 0.1387;
        public const double FUEL_TYPE_GASOLINE_VALUE = 0.1276;
        public const double FUEL_TYPE_KEROSENE_VALUE = 0.1351;
        public const double FUEL_TYPE_GAS_OIL_VALUE = 0.138;
        public const double FUEL_TYPE_LNG_VALUE = 0.086;

        public const double FUEL_TYPE_COAL_VALUE = 0.014;
        public const double FUEL_TYPE_COKE_VALUE = 0.013;
        public const double FUEL_TYPE_PEAT_VALUE = 0.006;
        public const double FUEL_TYPE_WOOD_VALUE = 0.008;
        public const double FUEL_TYPE_BIOMASS_VALUE = 0.009;
        public const double FUEL_TYPE_BLACK_LIQUOR_VALUE = 0.0065;
        public const double FUEL_TYPE_SCRAP_TIRES_VALUE = 0.016;
        public const double FUEL_TYPE_SULFUR_VALUE = 0.004;

        public const string INTERVAL_TYPE_DAILY = "Daily";
        public const string INTERVAL_TYPE_WEEKLY = "Weekly";
        public const string INTERVAL_TYPE_MONTHLY = "Monthly";

      
        public const int INTERVAL_TYPE_WEEK_COUNT = 52;
        public const int  INTERVAL_TYPE_MONTH_COUNT = 12;
        public const int INTERVAL_TYPE_DAYS_COUNT = 364;

        public const string LABEL_FISCAL_YEAR = "Fiscal Year";
        public const string LABEL_CURRENT_YEAR = "Calendar Year";



        public const string BEFORE_MODEL_ANNUAL_IMPROVMENT = "((1 - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-2,-1,1,1) )-(1 - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-2,0,1,1)))";
        public const string BEFORE_MODEL_CUMULATIVE_IMPROVMENT = "((1 - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,-1,1,1) )-(1 - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1))) + OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),0,-1,1,1)";
        public const string AFTER_MODEL_ANNUAL_IMPROVMENT = "OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1) - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,-1,1,1)";
        public const string AFTER_MODEL_CUMULATIVE_IMPROVMENT = "(1 - OFFSET(INDIRECT(ADDRESS(ROW(), COLUMN())),-1,0,1,1))";
    }
}
