using System;
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

    public partial class ReportingPeriodControl : UserControl
    {
        public Excel.ListObject DataLO;
        Excel.Worksheet thisSheet = Globals.ThisAddIn.Application.ActiveSheet as Excel.Worksheet;
        int dateColPos = 0;
        string dateaddr="";
        private bool fromWizard;

        public ReportingPeriodControl(bool fromWizard)
        {
            InitializeComponent();
            DataLO = ((Excel.Range)Globals.ThisAddIn.Application.Selection).ListObject;

            if (DataLO == null) DataLO = ((Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet).ListObjects[1];

            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);
            bool dateColChk = false;

            foreach (Excel.ListColumn cname in thisList.ListColumns)
            {
                string colName=cname.Name.ToUpper();
                if (colName.Equals("DATE"))
                {
                    dateColChk = true;
                    cname.Name="Date";
                }
            }

            if (!dateColChk)
            {
                this.cbBaselineYear.Enabled = false;
                this.cbInterval.Enabled = false;
                this.cbLabel.Enabled = false;
                this.btnReportingPeriod.Enabled = false;
            }

            this.btnBack.Visible = fromWizard;
            this.btnClose.Visible = fromWizard;
            this.btnNext.Visible = fromWizard;
            
        }

        private void ReportingPeriodControl_Load(object sender, EventArgs e)
        {

        }

        private void btnReportingPeriod_Click(object sender, EventArgs e)
        {


            int rowStart=Convert.ToInt32(this.cbBaselineYear.SelectedIndex.ToString());
            int rowInterval = Convert.ToInt32(this.cbInterval.SelectedIndex.ToString());
            string Interval=this.cbInterval.SelectedItem.ToString();
            int rowLabel = Convert.ToInt32(this.cbLabel.SelectedIndex.ToString());
            string Label = this.cbLabel.SelectedItem.ToString();
            int ccount = 0;
            int position=0;


            Excel.ListObject thisList = ExcelHelpers.GetListObject(thisSheet);
            Excel.ListColumn newColumn;
            newColumn = null;
         
            foreach (Excel.ListColumn cname in thisList.ListColumns)
            {
                if (cname.Name.Equals(EnPIResources.yearColName))
                {
                    position = ccount;
                    newColumn = cname;
                    
                }

                if (cname.Name.ToUpper()=="DATE")
                {
                    dateColPos = ccount;
                     dateaddr = cname.Range.Cells.Address;
                }
                ccount += 1;
            }




            if (position == 0)
            {
                newColumn = thisList.ListColumns.Add(thisList.ListColumns.Count + 1);
                newColumn.Name = EnPIResources.yearColName;
            }
            
            
            string stylename = "Comma";       
            if (Label == Constants.LABEL_FISCAL_YEAR)
                updateFiscialYear(Interval, newColumn,rowStart);
            if (Label == Constants.LABEL_CURRENT_YEAR)
                updateCalenderYear(thisList, newColumn, rowStart);
            
            newColumn.DataBodyRange.NumberFormat="####";       
        }

        private void updateFiscialYear(string interval, Excel.ListColumn LC, int rowStart)
        {
            //string formula = "=";
            //TFS Ticket : 70984
            if (interval == Constants.INTERVAL_TYPE_DAILY)
            {            
                //formula = formula + "\"" + "FY1" + "\"";
                //LC.DataBodyRange.Formula = formula;
                updateRowFYFormula(Constants.INTERVAL_TYPE_DAYS_COUNT, LC, rowStart);
            }
            
            if (interval == Constants.INTERVAL_TYPE_MONTHLY)
            {
                updateRowFYFormula(Constants.INTERVAL_TYPE_MONTH_COUNT,LC,rowStart );
            }

            if (interval == Constants.INTERVAL_TYPE_WEEKLY)
            {
                updateRowFYFormula(Constants.INTERVAL_TYPE_WEEK_COUNT, LC,rowStart);
            } 

        }

        private void updateCalenderYear(Excel.ListObject LO,Excel.ListColumn LC, int rowStart)
        {

            Excel.Range LCDate = ExcelHelpers.GetListColumn(LO, "Date").Range;
            if (LCDate == null) LCDate = ExcelHelpers.AddListColumn(LO, "Date", 1).Range;
            updateRowCYFormula(LO,LC, rowStart);

      
        }


        private void updateRowFYFormula(int Intervals, Excel.ListColumn LC,int rowStart)
        {
            int j = 1, i = 0;
            foreach (Excel.Range row in LC.Range.Rows)
            {

                if (i > rowStart)
                    row.Formula = "=" + "\"" + "FY" + j + "\"";
                if(i > 0 && i < rowStart)
                    row.Formula = "";

                if (i == Intervals + rowStart) 
                {
                    j += 1;
                    i = 0;
                    rowStart = 0;
                }

                i += 1;
            }
        }

        private void updateRowCYFormula(Excel.ListObject LO, Excel.ListColumn LC, int rowStart)
        {
            int i = 0;
            foreach (Excel.Range row in LC.Range.Rows)
            {
                if (i > rowStart)
                {
                    string addrStr = dateaddr.ToString().Substring(1, dateaddr.ToString().IndexOf("$", 1) - 1).ToString() + (i+1);
                    row.Formula = "=YEAR([Date])";
                }
                if (i > 0 && i <= rowStart)
                    row.Formula = "";
                i += 1;
            }
        }

        private int GetDaysInAYear(int year)
        {
            int days = 0;
            for (int i = 1; i <= 12; i++)
            {
                days += DateTime.DaysInMonth(year, i);
            }
            return days;
        }

        public void Open()
        {
            populateInterval();
            populateLabel();
            populateStartDate();
        }

        private void populateInterval()
        {
            this.cbInterval.Items.Add(Constants.INTERVAL_TYPE_DAILY);
            this.cbInterval.Items.Add(Constants.INTERVAL_TYPE_WEEKLY);
            this.cbInterval.Items.Add(Constants.INTERVAL_TYPE_MONTHLY);
        }

        private void populateLabel()
        {
            this.cbLabel.Items.Add(Constants.LABEL_FISCAL_YEAR);
            this.cbLabel.Items.Add(Constants.LABEL_CURRENT_YEAR);
        }

        private void populateStartDate()
        {
            int dateIndex = 1;

            foreach (Excel.ListColumn LC in DataLO.ListColumns)
            {
                if (LC.Name.Equals(EnPIResources.dateColName))
                    dateIndex = LC.Index;
            }

            foreach (Excel.ListRow row in DataLO.ListRows)
            {
                this.cbBaselineYear.Items.Add(Convert.ToString(((Excel.Range)row.Range[1,dateIndex]).Text.ToString()));
            }
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(5);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(6);
        }

    }
}
