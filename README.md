# EnPI
The EnPI V5.0 is a regression analysis based tool developed by the U.S. Department of Energy, Office of Advanced Manufacturing, to help plant and corporate managers establish a normalized baseline of energy consumption, track annual progress of intensity improvements, energy savings, Superior Energy Performance (SEP) EnPIs, and other EnPIs that account for variations due to weather, production, and other variables. The tool is designed to accommodate multiple users including Better Buildings, Better Plants Program and Challenge Partners, SEP participants, other manufacturing firms, and non-manufacturing facilities such as data centers.

Regression analysis is a statistical technique that estimates the dependence of a variable (typically energy consumption for energy use and intensity tracking) on one or more independent variables, such as ambient temperature, while controlling for the influence of other variables at the same time. Regression is commonly used for estimating energy savings through the measurement and verification of energy projects and programs, and has proven to be reliable when the input data covers the full annual variation in operating conditions. A properly used regression analysis can provide a reliable estimate of energy savings resulting from energy improvement strategies and projects by accounting for the effects of variables such as production variation and weather.

In addition to providing a normalized view of energy performance, the EnPI tool calculates metrics specific to the SEP and Better Plants Programs. For the SEP Program, the tool calculates SEnPIs, cumulative improvement, and annual improvement. Metrics required for the Better Plants annual report form are formatted to allow easy entry into the online eCenter annual report form. The tool also allows corporate energy managers to roll plant level energy data and metrics up to a corporate level to determine corporate energy performance.

# Inputs

-	Monthly Energy Consumption Data (preferably separately by type of energy, e.g., electricity, natural gas)
-	Any variables that affect the energy consumption in a facility (e.g.,  heating degree days (HDD), cooling degree days (CDD), dew point temperature, product output, moisture content of the product, shift schedule adjustments, etc.)

# Outputs

The tool identifies key variables affecting facility energy performance and calculates a modeled consumption based on the independent variables selected for regression. The tool outputs metrics required for both the SEP and Better Plants Program. The tool calculates SEnPIs, cumulative improvement, annual improvement, and normalized energy savings for the SEP Program. Other optional outputs include cost savings and avoided CO2 emissions. Metrics for the Better Plants Program are formatted in a manner consistent with the Better Plants Annual Reporting form, which allows for the data to be easily entered into the online reporting form. For the Better Plants Program, the tool calculates the following fields required for the annual report:

- Total Baseline Primary Energy Consumed (MMBtu/year)
-	Total Current Year Primary Energy Consumed (MMBtu/year)
- Adjustment for Baseline Primary Energy use (MMBtu/year)
- Adjusted Baseline of Primary Energy (MMBtu/year)
-	New Energy Savings for Current Year (MMBtu/year)
-	Total Energy Savings since Baseline Year (MMBtu/year)
-	Annual Improvement in Energy Intensity for Current Year (%)
-	Total Improvement in Energy Intensity for Baseline Year (%)

# System Requirements

Many companies have policies that prevent installation of external software components. Use of the EnPI tool requires a download of software to your computer. If you have difficulty downloading the EnPI tool, please send the following description of the EnPI tool (software) to your IT team to request assistance.
The EnPI tool is a standard executable Microsoft Excel COM add-in, which uses Microsoft Office libraries. The tool is downloaded from the Department of Energy (DOE) Energy Resource Center . The eCenter is a secure site and all tools located on the site are compliant with DOE’s security policies.
If you have issues downloading or running the tool, please contact the AMO Help Desk at AMO_ToolHelpDesk@ee.doe.gov.

# Known Issues

1.	If the calculate button fails to run the program, check the variable names to ensure that they are appropriate (e.g., no returns in column headers, etc.).
2.	If cannot switch model:
  	Save your work (in all open Excel Workbooks)
  	Close Excel
  	Open Task Manager
  	Navigate to the “Details” Tab
  	Locate and select EXCEL.EXE and end task
  	NOTE: this will force close all running Excel windows
3. If the tool is run in Office 2013, the wizard may flash once when first opened

