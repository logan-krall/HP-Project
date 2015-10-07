using System;
using System.Data;
using System.Data.OleDb;

//Spire XLS
using Spire.Xls;

//using Excel;

//Open XML sdk
using DocumentFormat;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

//Epplus
using OfficeOpenXml;

using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;
using System.Xml.Serialization;


namespace HP_Analytics_Project.Images
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string fullName = (string)(Session["name"]);
            string extension = System.IO.Path.GetExtension(fullName).ToLower();
            Session["radioCount"] = 0;

            List<string> cellTypes = new List<string>();
            
            DataSet myDataSet = new DataSet();
            DataTable myDataTable = new DataTable();

            if (extension == ".xls")
            {
                //Olebdb version
                string connectionString = string.Empty;
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;";
                connectionString += "Data Source='" + fullName + "';";
                connectionString += "Extended Properties='Excel 8.0;HDR=YES;IMEX=1;READONLY=TRUE;';";

                OleDbCommand cmd = new OleDbCommand();
                OleDbConnection conn = new OleDbConnection(connectionString);

                conn.Open();
                cmd.Connection = conn;
                //Get all sheets/tables from the file
                myDataTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                //Loop through all sheets/tables in the file
                foreach (DataRow dr in myDataTable.Rows)
                {

                    string sheetName = dr["TABLE_NAME"].ToString();
                    //Get all rows from the sheet/table
                    cmd.CommandText = "SELECT * FROM [" + sheetName + "]";

                    DataTable dt = new DataTable();
                    dt.TableName = sheetName;

                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);
                    da.Fill(dt);

                    if (dt != null && (dt.Rows.Count > 1 || dt.Columns.Count > 1))
                    {
                        myDataSet.Tables.Add(dt);
                    }
                }
                cmd = null;
                conn.Close();
            }
            else if (extension == ".xlsx")
            {
                //Spire XLS
                Spire.Xls.Workbook wrkbook = new Spire.Xls.Workbook();          //create new workbook
                wrkbook.LoadFromFile(@fullName);                                //load a file
                Spire.Xls.Worksheet wrksheet = wrkbook.Worksheets[0];           //initialize worksheet
                
                myDataTable = wrksheet.ExportDataTable();
                myDataSet.Tables.Add(myDataTable);

                /*
                 Open XML sdk
                 
                DocumentFormat.OpenXml.Packaging.SpreadsheetDocument doc = DocumentFormat.OpenXml.Packaging.SpreadsheetDocument.Open(fullName, false);

                //Creates Workbook Part
                WorkbookPart workbookPart = doc.WorkbookPart;
                //Creates IEnumerable list of sheets from the document, starting from the first
                IEnumerable<Sheet> sheets = doc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                //Extracts the ID of the first sheet to a string
                string relationID = sheets.First().Id.Value;
                //Finds the first worksheet part by the ID
                WorksheetPart worksheetPart = (WorksheetPart)doc.WorkbookPart.GetPartById(relationID);
                //Extracts the first worksheet from the first worksheet part
                Worksheet worksheet = worksheetPart.Worksheet;
                //Extracts the first sheet data from the first worksheet
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                //Creates IEnumerable list of rows from the first sheet data.
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (Cell cell in rows.ElementAt(0))
                {
                    DataColumn col = new DataColumn();
                    myDataTable.Columns.Add(col);
                }

                foreach (Row row in rows)
                {
                    DataRow newRow = myDataTable.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        newRow[i] = row.Descendants<Cell>().ElementAt(i).InnerText;
                    }

                    myDataTable.Rows.Add(newRow);
                }

                doc.Close();
                myDataSet.Tables.Add(myDataTable);
                */
                /*
                 EPPlus
                 * 
                DataTable dt = new DataTable();
                FileInfo fi = new FileInfo(fullName);

                // Check if the file exists
                if (!fi.Exists)
                    throw new Exception("File " + fullName + " Does Not Exists");

                using (ExcelPackage xlPackage = new ExcelPackage(fi))
                {
                    // get the first worksheet in the workbook
                    ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets.First();

                    // Fetch the WorkSheet size
                    ExcelCellAddress startCell = worksheet.Dimension.Start;
                    ExcelCellAddress endCell = worksheet.Dimension.End;

                    // create all the needed DataColumn
                    for (int col = startCell.Column; col <= endCell.Column; col++)
                        dt.Columns.Add(col.ToString());

                    // place all the data into DataTable
                    for (int row = startCell.Row; row <= endCell.Row; row++)
                    {
                        DataRow dr = dt.NewRow();
                        int x = 0;
                        for (int col = startCell.Column; col <= endCell.Column; col++)
                        {
                            dr[x++] = worksheet.Cells[row, col].Value;
                        }
                        dt.Rows.Add(dr);
                    }
                }
                myDataSet.Tables.Add(dt);
                */
                                
            }
            else if (extension == ".csv")
            {

            }

            Header_Initialization();

            var dataTypes = new[] { typeof(Byte), typeof(SByte), typeof(Decimal), typeof(Double), typeof(Single), typeof(Int16), 
                typeof(Int32), typeof(Int64), typeof(UInt16), typeof(UInt32), typeof(UInt64), typeof(Char), typeof(string) };

            ViewState["missing"] = false;

            foreach (DataTable dt in myDataSet.Tables)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dataTypes.Contains(dc.DataType))
                    {
                        double mean = 0, min = 0, max = 0, stdD = 0;
                        string varName = String.Empty, varType = String.Empty, uniqueVals = "*";

                        TableRow tRow = new TableRow();

                        TableCell dependCell = new TableCell();
                        TableCell nameCell = new TableCell();
                        TableCell varTCell = new TableCell();
                        TableCell meanCell = new TableCell();
                        TableCell minCell = new TableCell();
                        TableCell maxCell = new TableCell();
                        TableCell stdCell = new TableCell();
                        TableCell cardCell = new TableCell();

                        /*
                        //div for button format
                        HtmlGenericControl div = new HtmlGenericControl("div");
                        int num = (int)Session["radioCount"];
                        int radioNum1 = num * 100;
                        int radioNum2 = num * 100 + 1;
                        int radioNum3 = num * 100 + 2;

                        string name = "radio" + num.ToString();
                        div.Attributes.Add("id", name);

                        string inputText = "<input type='radio' id='" + radioNum1.ToString() + "' name='" + name + "' value = 'i " + dc.ColumnName.ToString() + "' onclick='javascript:__doPostBack(this.id, this.value)'><label for='" + radioNum1.ToString() + "'>Independent</label>" +
                                           "<input type='radio' id='" + radioNum2.ToString() + "' name='" + name + "' value = 'd " + dc.ColumnName.ToString() + "' onclick='javascript:__doPostBack(this.id, this.value)'><label for='" + radioNum2.ToString() + "'>Dependent</label>" +
                                           "<input type='radio' id='" + radioNum3.ToString() + "' name='" + name + "' value = '0 " + dc.ColumnName.ToString() + "' onclick='javascript:__doPostBack(this.id, this.value)'><label for='" + radioNum3.ToString() + "'>Ignore</label>";

                        div.InnerHtml = inputText;
                        
                                                
                        //depend1.Controls.Add(div);
                        //dependCell.Controls.Add(div);
                        //Session["radioCount"] = ++num;
                         
                         * 
                        */


                        //dependency radio list
                        RadioButtonList depend1 = new RadioButtonList();
                        
                        depend1.ID = dc.ColumnName.ToString();
                        depend1.AutoPostBack = true;
                        depend1.SelectedIndexChanged += new EventHandler((s, e1) => Radio_Changed(s, e1, dc.ColumnName.ToString()));
                        depend1.RepeatDirection = System.Web.UI.WebControls.RepeatDirection.Horizontal;
                        depend1.Font.Size = System.Web.UI.WebControls.FontUnit.XSmall;

                        ListItem ind = new ListItem();
                        ListItem dep = new ListItem();
                        ListItem ign = new ListItem();
                        ind.Text = "  Independent";
                        dep.Text = "  Dependent";
                        ign.Text = "  Ignore";
                        ind.Value = "i";
                        dep.Value = "d";
                        ign.Value = "0";

                        depend1.Items.Add(ind);
                        depend1.Items.Add(dep);
                        depend1.Items.Add(ign);
                        dependCell.Controls.Add(depend1);
                        
                        //Block for calculating Cardinality.
                        DataTable catVals = dt.DefaultView.ToTable(true, dc.ColumnName.ToString());
                        uniqueVals = catVals.Rows.Count.ToString();
                        if (catVals.Rows.Count == 2)
                        {
                            List<string> logits;
                            if (ViewState["logits"] == null)
                            {
                                logits = new List<string>();
                                ViewState["logits"] = logits;
                            }
                            else
                            {
                                logits = (List<string>) ViewState["logits"];
                            }
                            logits.Add(dc.ColumnName.ToString());
                        }

                        if (dc.DataType.ToString() == typeof(char).ToString() || dc.DataType.ToString() == typeof(string).ToString())
                        {
                            meanCell.Text = "*";
                            minCell.Text = "*";
                            maxCell.Text = "*";
                            stdCell.Text = "*";
                            varType = "Nominal";

                            //int indx = dt.Columns.IndexOf(dc);
                            //string type = dt.Rows[1][indx].GetType().ToString();
                            //varType = type;
                        }
                        else
                        {
                            varType = "Numeric";
                        }
                        //Block for calculating Mean.
                        if (dc.DataType.ToString() != typeof(char).ToString() && dc.DataType.ToString() != typeof(string).ToString())
                        {
                            object meanObject;
                            meanObject = dt.Compute("Avg(" + dc.ColumnName.ToString() + ")", string.Empty);
                            mean = Double.Parse(meanObject.ToString());
                            
                            //Block for calculating STD.
                            object stdDobject;
                            stdDobject = dt.Compute("StDev(" + dc.ColumnName.ToString() + ")", string.Empty);
                            stdD = Double.Parse(stdDobject.ToString());

                            //Block for calculating Min.
                            object minObject;
                            minObject = dt.Compute("Min(" + dc.ColumnName.ToString() + ")", string.Empty);
                            min = Double.Parse(minObject.ToString());

                            //Block for calculating Max.
                            object maxObject;
                            maxObject = dt.Compute("Max(" + dc.ColumnName.ToString() + ")", string.Empty);
                            max = Double.Parse(maxObject.ToString());

                            meanCell.Text = mean.ToString("0.#####");
                            minCell.Text = min.ToString("0.#####");
                            maxCell.Text = max.ToString("0.#####");
                            stdCell.Text = stdD.ToString("0.#####");
                        }

                        varName = dc.ColumnName.ToString();
                        nameCell.Text = varName;
                        cardCell.Text = uniqueVals;                            
                        varTCell.Text = varType;

                        dependCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        dependCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        nameCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        nameCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        varTCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        varTCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        meanCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        meanCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        minCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        minCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        maxCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        maxCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        stdCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        stdCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        cardCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        cardCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        tRow.Cells.Add(dependCell);
                        tRow.Cells.Add(nameCell);
                        tRow.Cells.Add(varTCell);
                        tRow.Cells.Add(meanCell);
                        tRow.Cells.Add(minCell);
                        tRow.Cells.Add(maxCell);
                        tRow.Cells.Add(stdCell);
                        tRow.Cells.Add(cardCell);

                        Table1.Rows.Add(tRow);

                        int colNum = 0, missingV = 0;
                        colNum = dt.Columns.IndexOf(dc);

                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr[colNum].ToString().Length == 0) { missingV += 1; }
                        }

                        if (missingV > 0)
                        {
                            ViewState["missing"] = true;
                            if (Table2.Rows[1].Cells[0].Text == "-")
                            {
                                Table2.Rows.Remove(Table2.Rows[1]);
                            }
                            TableRow missRow = new TableRow();
                            TableCell missVCell = new TableCell();
                            TableCell missNCell = new TableCell();
                            missVCell.Text = missingV.ToString();
                            missNCell.Text = nameCell.Text;
                            missNCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            missNCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            missVCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            missVCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            missRow.Cells.Add(missNCell);
                            missRow.Cells.Add(missVCell);

                            Table2.Rows.Add(missRow);
                        }
                    }
                }               
            }
        }

        void Build_Startup_Script()
        {
            //javascript for button format
            ClientScriptManager csm = Page.ClientScript;
            StringBuilder sb = new StringBuilder();
            Type csType = this.GetType();
            string csName = "RadioScript";

            sb.Append("<script>");
            int count = (int)Session["radioCount"];
            for (int i = 0; i < count; i++)
            {
                sb.Append("$(function() {");
                sb.Append("$( '#radio" + i.ToString() + "' ).buttonset()");
                sb.Append("});");
            }
            sb.Append("</script>");
            csm.RegisterStartupScript(csType, csName, sb.ToString());
        }

        void Header_Initialization()
        {
            //missing value header row
            TableRow hRow2 = new TableRow();
            TableRow hRow3 = new TableRow();

            TableCell missNCellH = new TableCell();
            TableCell missVCellH = new TableCell();
            missNCellH.Text = "Name of Column";
            missVCellH.Text = "Rows Missing";
            missNCellH.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
            missNCellH.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
            missVCellH.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
            missVCellH.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
            hRow2.Cells.Add(missNCellH);
            hRow2.Cells.Add(missVCellH);

            TableCell filler1 = new TableCell();
            TableCell filler2 = new TableCell();
            filler1.Text = "-";
            filler2.Text = "-";
            hRow3.Cells.Add(filler1);
            hRow3.Cells.Add(filler2);

            Table2.Rows.Add(hRow2);
            Table2.Rows.Add(hRow3);

            //main table header row
            TableRow hRow = new TableRow();

            TableCell dependH = new TableCell();
            TableCell nameCellH = new TableCell();
            TableCell varTCellH = new TableCell();
            TableCell meanCellH = new TableCell();
            TableCell medCellH = new TableCell();
            TableCell moCellH = new TableCell();
            TableCell stdCellH = new TableCell();
            TableCell cardCellH = new TableCell();

            dependH.Text = "Variable Dependency";
            nameCellH.Text = "Name";
            varTCellH.Text = "Type";
            meanCellH.Text = "Mean";
            medCellH.Text = "Min";
            moCellH.Text = "Max";
            stdCellH.Text = "Std Dev";
            cardCellH.Text = "Cardinality";

            hRow.Cells.Add(dependH); ;
            hRow.Cells.Add(nameCellH);
            hRow.Cells.Add(varTCellH);
            hRow.Cells.Add(meanCellH);
            hRow.Cells.Add(medCellH);
            hRow.Cells.Add(moCellH);
            hRow.Cells.Add(stdCellH);
            hRow.Cells.Add(cardCellH);

            Table1.Rows.Add(hRow);
        }

        void Radio_Changed(object sender, EventArgs e, string col)
        //void Radio_Changed(object sender, string col)
        {
            Dictionary<string, string> depDic = new Dictionary<string, string>();

            if (ViewState["dict"] != null)
            { 
                depDic = (Dictionary<string, string>)ViewState["dict"]; 
            }

            RadioButtonList rb1 = (sender as RadioButtonList);
            if (depDic.ContainsKey(col))
            {   
                depDic[col] = rb1.SelectedItem.Value; 
            }
            else
            {   
                depDic.Add(col, rb1.SelectedItem.Value); 
            }

            List<string> main = new List<string>();
            List<string> dep = new List<string>();

            foreach ( KeyValuePair<string,string> kp in depDic)
            {
                if (kp.Value == "i")
                {   
                    main.Add(kp.Key); 
                }
                else if (kp.Value == "d")
                {   
                    dep.Add(kp.Key); 
                }
            }
            main.Sort();
            dep.Sort();

            foreach ( string v in dep )
            {   
                main.Add(v); 
            }

            TableRow trH = new TableRow();
            TableCell corner = new TableCell();
            corner.Text = "-";
            trH.Cells.Add(corner);
            Table4.Rows.Add(trH);

            foreach ( string v in main )
            {
                TableCell tcH1 = new TableCell();
                tcH1.Text = v;
                if (depDic[v] == "i")
                {   
                    tcH1.Font.Bold = true; 
                }
                tcH1.HorizontalAlign = HorizontalAlign.Center;
                tcH1.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                tcH1.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                trH.Cells.Add(tcH1);                

                TableRow tr = new TableRow();
                TableCell tc1 = new TableCell();
                tc1.Text = v;
                if (depDic[v] == "i")
                { 
                    tc1.Font.Bold = true; 
                }
                tc1.HorizontalAlign = HorizontalAlign.Center;
                tc1.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                tc1.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                tr.Cells.Add(tc1);

                for (int i = 0; i < main.Count; i++)
                {
                    TableCell cell = new TableCell();

                    if (i == main.IndexOf(v))
                    {   
                        cell.Text = "1"; 
                    }
                    else
                    {   
                        cell.Text = "-"; 
                    }

                    cell.HorizontalAlign = HorizontalAlign.Center;
                    cell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                    cell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                    tr.Cells.Add(cell);
                }
                Table4.Rows.Add(tr);
            }
            ViewState["dict"] = depDic;
        }

        public bool Depend_Check()
        {
            Dictionary<string, string> depDic = new Dictionary<string, string>();
            if (ViewState["dict"] != null)
            { depDic = (Dictionary<string, string>) ViewState["dict"]; }
            if ( depDic.Count > 0 && depDic.Values.Contains("i") && depDic.Values.Contains("d") )
            {   
                return true; 
            }
            return false; 
        }

        public bool Multi_Reg_Check()
        {
            Dictionary<string, string> depDic = new Dictionary<string, string>();
            if (ViewState["dict"] != null)
            { depDic = (Dictionary<string, string>)ViewState["dict"]; }
            if (depDic.Count > 0 && depDic.Values.Contains("d") && depDic.Count(D=>D.Value.Contains("i")) >= 2)
            { 
                return true; 
            }
            return false; 
        }

        public bool Logit_Reg_Check()
        {
            Dictionary<string, string> depDic = new Dictionary<string, string>();
            if (ViewState["dict"] != null)
            { depDic = (Dictionary<string, string>)ViewState["dict"]; }
            if (depDic.Count > 0 && depDic.Values.Contains("i") && depDic.Values.Contains("d"))
            {
                foreach (KeyValuePair<string, string> kp in depDic.Where(D => D.Value.Contains("d")))
                {
                    List<string> logits = (List<string>) ViewState["logits"];
                    if (logits != null && logits.Count > 0 && logits.Contains(kp.Key))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public void Model_Options()
        {

        }
    }
}