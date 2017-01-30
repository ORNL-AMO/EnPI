using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.IO.Packaging;
using System.Xml;
using Excel = Microsoft.Office.Interop.Excel;
using AMO.EnPI.AddIn.Utilities;

namespace AMO.EnPI.AddIn
{
    public partial class RollupControl : UserControl
    {
        public RollupControl(bool fromWizard)
        {
            InitializeComponent();

            this.btnBack.Visible = fromWizard;
            this.btnClose.Visible = fromWizard;
            this.btnNext.Visible = fromWizard;

        }

        private DetailTableCollection sources;
        private BindingList<DetailTable> tables;
        private BindingList<SourceFile> files;
        public int numberOfSources;
        public bool fromActual = false;
        public bool hasProd = false;
        public bool hasBuildSqFt = false;
        public bool fromEnergyCost = false;

        public void AddFiles()
        {
            try
            {
                // loop through the list objects in the open files and add them to the list box
                foreach (GroupSheetCollection gsc in Globals.ThisAddIn.masterGroupCollection)
                {
                    if (gsc.WBName != Globals.ThisAddIn.Application.ActiveWorkbook.Name)
                    {
                        Excel.Workbook gsWB = ExcelHelpers.GetWorkbook(Globals.ThisAddIn.Application, gsc.WBName);
                        bool amisaved = gsWB.Saved;

                        for (int j = 0; j < gsc.Count; j++)
                        {
                            GroupSheet gs = gsc.Item(j);

                            if (ExcelHelpers.GetWorksheetbyGUID(gsWB, gs.WSGUID) != null)
                            {
                                if (gs.adjustedDataSheet)
                                {
                                    string nm = gs.Name == null ? gs.WSName() : gs.Name + " -- " + gs.WSName();
                                    string disp = " (" + gsc.WBName + ") " + nm;

                                    try
                                    {
                                        DetailTable dt = new DetailTable(gs.WS.ListObjects.get_Item(1), nm, numberOfSources, fromActual, hasProd, hasBuildSqFt, true, fromEnergyCost);
                                        if (dt.SQLStatement != "")
                                        {
                                            SourceFile sf = new SourceFile(dt.SQLStatement, disp, gsWB.FullName, false);
                                            if (!files.Contains(sf))
                                            {
                                                files.Add(sf);
                                            }
                                        }
                                    }
                                    catch
                                    {
                                        continue; 
                                    }
                                }
                            }
                        }
                        gsWB.Saved = amisaved;
                    }
                }
                AddFilesToList();
            }
            catch
            {
            }
        }

        public void AddTablesToList()
        {
            if (tables.Count == 0) { return; }

            this.checkedListBox1.DataSource = tables;
            this.checkedListBox1.ValueMember = "SQLStatement";
            this.checkedListBox1.DisplayMember = "DisplayName";
            this.checkedListBox1.Refresh();

            int len = 0;

            foreach (DetailTable st in tables)
            {
                len = Math.Max(st.DisplayName.Length, len);
            }

            int sz1 = Convert.ToInt16(this.checkedListBox1.Font.SizeInPoints * 0.75);
            int sz2 = 11; //check box size
        }
        public void AddFilesToList()
        {
            if (files.Count == 0) { return; }

            this.checkedListBox2.DataSource = files;
            this.checkedListBox2.ValueMember = "SQLStatement";
            this.checkedListBox2.DisplayMember = "DisplayName";

            this.checkedListBox2.Refresh();

            int len = 0;

            foreach (SourceFile sf in files)
            {
                len = Math.Max(sf.DisplayName.Length, len);
            }

            bool addScroll = false;
            if(files.Count >= 6)
                addScroll = true;

            int sz1 = Convert.ToInt16(this.checkedListBox2.Font.SizeInPoints * 0.75);
            int sz2 = 11; //check box size

            this.Width = (sz1 * len) + sz1 + sz2 + this.checkedListBox1.Margin.Left + this.checkedListBox1.Margin.Right;
            int checkWidth = sz1 * len + sz2 + this.checkedListBox2.Margin.Left + this.checkedListBox2.Margin.Right;
            if (checkWidth > 510)
            {
                this.checkedListBox2.HorizontalScrollbar = true;
                this.checkedListBox2.Width = 510;
            }
            else
                this.checkedListBox2.Width = checkWidth;

            if (addScroll)
                this.checkedListBox2.ScrollAlwaysVisible = true;
            //removed and added scroll 
            //this.checkedListBox2.Height = (sz2 + this.checkedListBox2.Margin.Top + this.checkedListBox2.Margin.Bottom)
            //                                * (this.checkedListBox2.Items.Count + 1);
            this.checkedListBox2.Visible = true;
            this.btnImport.Visible = false;

            Globals.ThisAddIn.wizardPane.Width = this.checkedListBox2.Width + this.checkedListBox2.Left + this.Margin.Right;
            Globals.ThisAddIn.wizardPane.Visible = true;
        }

        public void OpenFiles()
        {
            bool caught = false;
            
            // opens the selected files so that the group sheet collections can be read
            for (int i = 0; i < this.openFileDialog1.FileNames.Count(); i++)
            {
                try
                {
                    string fileName = this.openFileDialog1.FileNames.GetValue(i).ToString();
                    addExcelSheetNames(fileName);
                }
                catch
                {
                    caught = true;
                    MessageBox.Show("An error was encountered while attempting to import sheets from workbook " + this.openFileDialog1.FileNames.GetValue(i).ToString() + ". Please re-run the data in a new workbook before including it in the corporate roll-up.");
                }
            }

            if(!caught)
                AddFilesToList();
            
        }

        public void Open()
        {
            Globals.ThisAddIn.wizardPane.Visible = false;
            Globals.ThisAddIn.Application.ActiveWorkbook.EnableConnections();

            this.tables = new BindingList<DetailTable>();
            this.files = new BindingList<SourceFile>();

            addRawDataSheets();
            AddTablesToList();

            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;

            RestoreValues(thisSheet, checkedListBox1, Utilities.Constants.CB_TABLES);
            RestoreValues(thisSheet, checkBox1, "NewSheet");

            this.checkedListBox1.Visible = true;
            this.btnRun.Visible = true;
        }

        private void RestoreFiles(Excel.Worksheet thisSheet, string propname)
        {
            string vals = Utilities.ExcelHelpers.getWorksheetCustomProperty(thisSheet, propname);

            if (vals != null)
            {
                System.Xml.XmlReader xvals = System.Xml.XmlReader.Create(new System.IO.StringReader(vals));
                string x = "";
                string y = "";
                string z = "";
                while (xvals.Read())
                {
                    if (xvals.NodeType == System.Xml.XmlNodeType.Element && xvals.Name == "SourceFile")
                    {
                        x = "";
                        y = "";
                        z = "";
                    }
                    if (xvals.NodeType == System.Xml.XmlNodeType.Element && xvals.Name == "FileName")
                    {
                        z = xvals.ReadElementContentAsString();
                    }
                    if (xvals.NodeType == System.Xml.XmlNodeType.Element && xvals.Name == "SQLStatement")
                    {
                        y = xvals.ReadElementContentAsString();
                    }
                    if (xvals.NodeType == System.Xml.XmlNodeType.Element && xvals.Name == "DisplayName")
                    {
                        x = xvals.ReadElementContentAsString();
                    }
                    if (xvals.NodeType == System.Xml.XmlNodeType.EndElement && xvals.Name == "SourceFile")
                    {
                        if (!files.Contains(new SourceFile(y, x, z, true)))
                        {
                            files.Add(new SourceFile(y, x, z, true));
                        }
                    }
                }
            }
            AddFilesToList();
        }

        private void RestoreValues(Excel.Worksheet thisSheet, CheckedListBox thisList, string propname)
        {
            string vals = Utilities.ExcelHelpers.getWorksheetCustomProperty(thisSheet, propname);

            if (vals != null)
            {
                System.Xml.XmlReader xvals = System.Xml.XmlReader.Create(new System.IO.StringReader(vals));

                while (xvals.Read())
                {
                    if (xvals.NodeType == System.Xml.XmlNodeType.Text)
                    {
                        for (int i = 0; i < thisList.Items.Count; i++)
                        {
                            if (thisList.GetItemText(thisList.Items[i]) == xvals.Value)
                            {
                                thisList.SetItemChecked(i, true);
                            }
                        }
                    }
                }
            }
        }

        private void RestoreValues(Excel.Worksheet thisSheet, CheckBox thisBox, string propname)
        {
            string vals = Utilities.ExcelHelpers.getWorksheetCustomProperty(thisSheet, propname);

            if (vals != null)
            {
                System.Xml.XmlReader xvals = System.Xml.XmlReader.Create(new System.IO.StringReader(vals));

                while (xvals.Read())
                {
                    if (xvals.Name == propname)
                    {
                        Boolean x;
                        if (Boolean.TryParse(xvals.GetAttribute("Checked").ToString(), out x))
                            thisBox.Checked = Boolean.Parse(xvals.GetAttribute("Checked").ToString());
                    }
                }
            }
        }

        private void SaveFiles(Excel.Worksheet thisSheet, BindingList<SourceFile> files, string propname)
        {
            if (files == null || files.Count == 0) { return; }

            System.Text.StringBuilder strvals = new System.Text.StringBuilder();
            System.Xml.XmlWriter xvals = System.Xml.XmlWriter.Create(strvals);
            xvals.WriteStartElement(propname);

            foreach (SourceFile s in files)
            {
                s.WriteFileXML(xvals);
            }

            xvals.WriteEndElement();
            xvals.Close();

            Utilities.ExcelHelpers.addWorksheetCustomProperty(thisSheet, propname, strvals.ToString());
        }

        private void SaveValues(Excel.Worksheet thisSheet, CheckedListBox thisList, string propname)
        {
            if (thisList.CheckedItems.Count == 0) { return; }

            System.Text.StringBuilder strvals = new System.Text.StringBuilder();
            System.Xml.XmlWriter xvals = System.Xml.XmlWriter.Create(strvals);
            xvals.WriteStartElement(propname);

            foreach (var s in thisList.CheckedItems)
            {
                xvals.WriteElementString("Item", thisList.GetItemText(s));
            }

            xvals.WriteEndElement();
            xvals.Close();

            Utilities.ExcelHelpers.addWorksheetCustomProperty(thisSheet, propname, strvals.ToString());
        }

        private void SaveValues(Excel.Worksheet thisSheet, Boolean thisYN, string propname)
        {
            System.Text.StringBuilder strvals = new System.Text.StringBuilder();
            System.Xml.XmlWriter xvals = System.Xml.XmlWriter.Create(strvals);
            xvals.WriteStartElement(propname);
            xvals.WriteAttributeString("Checked", thisYN.ToString());
            xvals.WriteEndElement();
            xvals.Close();

            Utilities.ExcelHelpers.addWorksheetCustomProperty(thisSheet, propname, strvals.ToString());
        }

        private void SetSources()
        {
            sources = new DetailTableCollection();

            for (int j = 0; j < this.checkedListBox1.CheckedItems.Count; j++)
            {
                DetailTable st = (DetailTable)this.checkedListBox1.CheckedItems[j];
                sources.Add(st);
            }
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            Excel.Worksheet thisSheet = (Excel.Worksheet)Globals.ThisAddIn.Application.ActiveSheet;
            SetSources();
            SaveValues(thisSheet, checkedListBox1, Utilities.Constants.CB_TABLES);
            SaveValues(thisSheet, checkBox1.Checked, "NewSheet");

            RollupSheet rs;
            if (this.checkBox1.Checked)
            {
                rs = new RollupSheet(null);
                Utilities.ExcelHelpers.copyWorksheetCustomProperties(thisSheet, rs.WS);
            }
            else
            {
                rs = new RollupSheet(thisSheet);
            }

            rs.RollupSources = sources;
            rs.Initialize(checkedListBox1.CheckedItems.Count);

            Globals.ThisAddIn.hideWizard();
            this.Dispose();
        }

        private void openFileDialog1_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Path == "")
            {
                System.Windows.Forms.MessageBox.Show("Please save your current workbook before linking to other files", "Save File");
                return;
            }
            else
                this.openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            OpenFiles();
        }

        private void openFileDialog2_Click(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook.Path == "")
            {
                System.Windows.Forms.MessageBox.Show("Please save your current workbook before continuing", "Save File");
                return;
            }
            else
                btnRun_Click(sender, e);
        }

        private void addRawDataSheets()
        {
            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;

            foreach (Excel.Worksheet ws in WB.Sheets)
            {
                fromActual = false;
                fromEnergyCost = false;

                string raw = ExcelHelpers.getWorksheetCustomProperty(ws, Utilities.Constants.WS_ROLLUP) ?? "FALSE";

                string flnm = ExcelHelpers.getWorksheetCustomProperty(ws, Utilities.Constants.WS_SRCFILE) ?? "";

                numberOfSources = 0;

                if (bool.Parse(raw) && ws.ListObjects.Count > 0)
                {
                    //parse through ws to populate numberOfSources, fromActual, hasProd and hasBuildSqFt

                    if (ws.Name.Contains("EnPI Actual Results"))
                        fromActual = true;
                    
                    int parseRows = ws.ListObjects[1].ListRows.Count;
                    bool endOfSources = false;

                    for (int i = 5; i < parseRows + 4; i++)
                    {

                        try
                        {
                            string cellValue = ws.Range["A" + i.ToString()].Value2.ToString();

                            if (cellValue.Equals(Globals.ThisAddIn.rsc.GetString("unadjustedTotalColName")))
                                endOfSources = true;

                            if (cellValue.Equals("Total Production Output"))
                                hasProd = true;

                            if (cellValue.Equals(Globals.ThisAddIn.rsc.GetString("unadjustedBuildingColName")))
                                hasBuildSqFt = true;

                            if (cellValue.Contains("Estimated Cost Savings"))
                                fromEnergyCost = true;

                            //This is a work around, as there is a problem in identifying whether the sheet is Actual or from Regression, but the below item is only in actual sheet so 
                            //Considered this is as a differentiator.
                            if (cellValue.Contains("Total Savings Since Baseline Year (MMBtu/Year)"))
                                fromActual = true;

                            if (!endOfSources)
                                numberOfSources++;
                        }
                        catch (Exception e)
                        {

                        }

                    }

                    DetailTable newt = new DetailTable(ws.ListObjects[1], ws.Name, numberOfSources, fromActual, hasProd, hasBuildSqFt, false, fromEnergyCost);
                    newt.DisplayName = flnm == "" ? ws.Name : ws.Name + " (" + flnm + ")";

                    if (newt.SQLStatement != null)
                        tables.Add(newt);
                }
            }
        }

        private void addExcelSheetNames(string excelFile)
        {
            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;\";";
            connString = connString.Replace("{0}",excelFile);

            System.Data.OleDb.OleDbConnection aConn = new System.Data.OleDb.OleDbConnection(connString);
            aConn.Open();

            string cmdText = "SELECT * FROM [" + excelFile + "].[stateData$]";

            System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(cmdText, aConn);
            DataTable pt = new DataTable();
            da.Fill(pt);
            aConn.Close();

            
            //the way data is stored in stateData is horizontal
            int reps = (pt.Columns.Count + 1) / 5;

            //added per ticket# 66444
            //displays error message if tool has not been run in the workbook the user is attempting to import
            if (reps == 0)
                MessageBox.Show("The EnPI tool has not been run in workbook: " + excelFile + ". The tool must be run at the plant level in the workbook before proceeding.");
            
            for (int j = 0; j < reps; j++)
            {

                numberOfSources = 0;
                bool checkModelData = false;
                string actualSheetName = "";
                foreach (DataRow dr in pt.Rows)
                {
                    string sheetName = dr[(5 * (j + 1)) - 4].ToString();

                    if (sheetName.Contains("Model Data"))
                    {
                        //this try block will catch if the sheet name in stateData was deleted from the workbook
                        try
                        {
                            string cmd = "SELECT * FROM [" + excelFile + "].[" + sheetName + "$]";
                            System.Data.OleDb.OleDbConnection aConn3 = new System.Data.OleDb.OleDbConnection(connString);

                            aConn3.Open();
                            System.Data.OleDb.OleDbDataAdapter da3 = new System.Data.OleDb.OleDbDataAdapter(cmd, aConn3);
                            DataTable pt3 = new DataTable();
                            da3.Fill(pt3);
                            aConn3.Close();

                            checkModelData = true;
                        }
                        catch
                        {
                            //ensure that none of the working sheets from the workbook are imported if one of the sheets needed was deleted.
                            files.Clear();
                        }
                    }
                    if (sheetName.Contains("Actual Results"))
                        actualSheetName = dr[(5 * (j + 1)) - 4].ToString();
                    if (!checkModelData)
                        numberOfSources++;
                }

                //if we never hit the model data sheet, then it is a use actual run and we need to determine the number of sources
                if (!checkModelData)
                {
                    numberOfSources = 0;
                    string cmd = "SELECT * FROM [" + excelFile + "].[" + actualSheetName + "$]";
                    System.Data.OleDb.OleDbConnection aConn2 = new System.Data.OleDb.OleDbConnection(connString);

                    aConn2.Open();
                    System.Data.OleDb.OleDbDataAdapter da2 = new System.Data.OleDb.OleDbDataAdapter(cmd, aConn2);
                    DataTable pt2 = new DataTable();
                    da2.Fill(pt2);
                    aConn2.Close();


                    bool checkTotal = false;
                    for(int i = 3; i < pt2.Rows.Count; i++) //2 as an empty row is added in the actual result the number is increased by 1
                    {
                        if(!checkTotal)
                        {
                            if (pt2.Rows[i][0].ToString().Contains("TOTAL"))
                                checkTotal = true;
                            if (!checkTotal)
                                numberOfSources++;
                        }
                    }
                }

                //string plantName="";
                //SourceFile f = new SourceFile("",excelFile);
                //for (int i = 0; i < pt.Rows.Count; i++)
                //{
                //    string sheetName = pt.Rows[i][1 + j * 5].ToString();
                //    string isdetail = pt.Rows[i][3 + j * 5].ToString() == "" ? "false" : pt.Rows[i][3 + j * 5].ToString();
                //    plantName = pt.Rows[i][0 + j * 5].ToString() == "plantName" ? pt.Rows[i][1 + j * 5].ToString() : plantName;
                //    if (bool.Parse(isdetail))
                //        f = new SourceFile(sheetName, excelFile);
                //}
                //f.ShortName = plantName;
                //if (f.ShortName != "")
                //    files.Add(f);
                string plantName = "";
                int enpiIndex = -1;
                SourceFile f = new SourceFile("", excelFile, numberOfSources, fromActual);
                for (int i = 0; i < pt.Rows.Count; i++)
                {
                    fromActual = false;
                    string sheetName = pt.Rows[i][1 + j * 5].ToString();

                    if (sheetName.Contains("EnPI Results"))
                        enpiIndex = i;

                    if (sheetName.Contains("EnPI Actual"))
                    {
                        enpiIndex = i;
                        fromActual = true;
                    }

                    //string isdetail = pt.Rows[i][3 + j * 5].ToString() == "" ? "false" : pt.Rows[i][3 + j * 5].ToString();
                    plantName = pt.Rows[i][0 + j * 5].ToString() == "plantName" ? pt.Rows[i][1 + j * 5].ToString() : plantName;
                    if (i.Equals(enpiIndex))//bool.Parse(isdetail))
                        f = new SourceFile(sheetName, excelFile, numberOfSources, fromActual);
                }

                f.ShortName = plantName;
                if (f.ShortName != "")
                    files.Add(f);
                else
                { }
            }
        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            bool nonechecked = false;

            for(int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (checkedListBox2.GetItemChecked(i))
                {
                    nonechecked = true;
                    break;
                }
            }

            this.btnImport.Visible = nonechecked;
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            Excel.Workbook WB = Globals.ThisAddIn.Application.ActiveWorkbook;
            Excel.Worksheet activeSheet = (Excel.Worksheet)WB.ActiveSheet;
            List<int> deletes = new List<int>();

            string connString = Utilities.Constants.EXCEL_CONNSTRING;

            for (int i = 0; i < checkedListBox2.Items.Count; i++)
            {
                if (checkedListBox2.GetItemChecked(i))
                {
                    SourceFile fl = (SourceFile)checkedListBox2.Items[i];
                    string flnm = fl.FileName.Substring(fl.FileName.LastIndexOf("\\")).Replace("\\","");

                    // create new worksheet
                    string newSheetNm = ExcelHelpers.CreateValidWorksheetName(WB, fl.ShortName, 0);
                    Excel.Worksheet newSheet = (Excel.Worksheet)WB.Sheets.Add(System.Type.Missing, WB.Sheets[WB.Sheets.Count], System.Type.Missing, System.Type.Missing);
                    newSheet.Name = newSheetNm;

                    // add custom properties
                    ExcelHelpers.addWorksheetCustomProperty(newSheet, Utilities.Constants.WS_ROLLUP, "TRUE");
                    ExcelHelpers.addWorksheetCustomProperty(newSheet, Utilities.Constants.WS_SRCFILE, flnm);

                    // import the data
                    DataTable dt = new DataTable();
                    connString = connString.Replace("{0}", fl.FileName);
                    System.Data.OleDb.OleDbConnection aConn = new System.Data.OleDb.OleDbConnection(connString);
                    aConn.Open();
                    System.Data.OleDb.OleDbDataAdapter da = new System.Data.OleDb.OleDbDataAdapter(fl.SQLStatement, aConn);
                    try
                    {
                        da.Fill(dt);
                    }
                    catch
                    {
                        MessageBox.Show("An error was encountered while attempting to import sheets from workbook " + this.openFileDialog1.FileNames.GetValue(i).ToString() + ". Please re-run the data in a new workbook before including it in the corporate roll-up.");
                        break;
                    }
                    aConn.Close();
                    aConn.Dispose();

                    // copy the data to the new sheet
                    Excel.Range rng = newSheet.Rows.get_Resize(dt.Rows.Count + 1, dt.Columns.Count + 1).get_Offset(3, 0);

                    //rng.get_Offset(0, 0).get_Resize(1, 1).Value2 = "Name";

                    for (int c = 0; c < dt.Columns.Count; c++)
                    {
                        rng.get_Offset(0, c).get_Resize(1, 1).Value2 = dt.Columns[c].ColumnName;
                    }

                    rng.get_Offset(1, 0).get_Resize(dt.Rows.Count, dt.Columns.Count).Value2 = DataHelper.dataTableArrayObject(dt);
                    //rng.get_Offset(1, 0).get_Resize(dt.Rows.Count, 1).Formula = fl.ShortName;//"=IFERROR(RIGHT(CELL(\"filename\",$A$1), LEN(CELL(\"filename\",$A$1)) - FIND(\"]\",CELL(\"filename\",$A$1),1)),\"\")";
                    object[,] tmp1 = (object[,])rng.Value2;

                    bool hasProd = false;
                    bool hasBuildSqFt = false;
                    fromEnergyCost = false;

                    foreach (DataRow dr in dt.Rows)
                    {
                        if (dr[0].ToString().Contains("Total Production Output") || dr[0].ToString().Contains("Production Energy Intensity (MMBtu/unit production)"))
                            hasProd = true;
                        if (dr[0].ToString().Contains("Building Energy Intensity"))
                            hasBuildSqFt = true;
                        if (dr[0].ToString().Contains("Estimated Cost Savings"))
                            fromEnergyCost = true;
                    }

                    //-----------------------------------------------
                    newSheet.ListObjects.Add(Microsoft.Office.Interop.Excel.XlListObjectSourceType.xlSrcRange, rng, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing);//.AddEx(Excel.XlListObjectSourceType.xlSrcRange, rng, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing, System.Type.Missing);

                    DetailTable newdt = new DetailTable(newSheet.ListObjects[1], fl.ShortName, fl.numOfSources, fl.fromActual, hasProd, hasBuildSqFt, true, fromEnergyCost);
                    newdt.DisplayName = newSheetNm + " (" + flnm + ")";
                    tables.Add(newdt);

                }

            }

            files.Clear();
            checkedListBox2.Refresh();
            checkedListBox2.Visible = false;
            btnImport.Visible = false;

            // reload the tables box
            AddTablesToList();

            activeSheet.Activate();

            //        // copy the data to the new sheet
            //        int jslkdf = dt.Rows.Count;
            //        Excel.Range rng = newSheet.get_Range("A1").get_Resize(((dt.Rows.Count + 1)/2)+1, dt.Columns.Count + 1);
            //        rng.get_Offset(0, 0).get_Resize(1, 1).Value2 = "Name";

            //        for (int c = 0; c < dt.Columns.Count; c++)
            //        {
            //            rng.get_Offset(0, c + 1).get_Resize(1,1).Value2 = dt.Columns[c].ColumnName;
            //        }
                    
            //        rng.get_Offset(1, 1).get_Resize(dt.Rows.Count, dt.Columns.Count).Value2 = DataHelper.dataTableArrayObject(dt);
            //        rng.get_Offset(1, 0).get_Resize(dt.Rows.Count, 1).Formula = "=IFERROR(RIGHT(CELL(\"filename\",$A$1), LEN(CELL(\"filename\",$A$1)) - FIND(\"]\",CELL(\"filename\",$A$1),1)),\"\")";
            //        object[,] tmp1 = (object[,])rng.Value2;
            //        newSheet.ListObjects.AddEx(Excel.XlListObjectSourceType.xlSrcRange, rng, System.Type.Missing, Excel.XlYesNoGuess.xlYes, System.Type.Missing, System.Type.Missing);

            //        DetailTable newdt = new DetailTable(newSheet.ListObjects[1], fl.ShortName);
            //        newdt.DisplayName = newSheetNm + " (" + flnm + ")";
            //        tables.Add(newdt);

            //    }

            //}

            //files.Clear();
            //checkedListBox2.Refresh();
            //checkedListBox2.Visible = false;
            //btnImport.Visible = false;

            //// reload the tables box
            //AddTablesToList();

            //activeSheet.Activate();
        }

        private void btnBack_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(9);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.hideWizard();
        }

        private void btnNext_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.LaunchWizardControl(10);
        }
    }
}
