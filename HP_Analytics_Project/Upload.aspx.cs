using System;
using System.Data;
using System.Data.OleDb;

//Accord.net
using Accord.IO;
using Accord.Math;
using Accord.Statistics;
using Accord.Collections;

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
                //Accord.net
                /*
                ExcelReader reader = new ExcelReader(fullName);                 //Create a new reader, given the filepath
                string[] sheets = reader.GetWorksheetList();                    //Query the file for all worksheets

                for (int i = 0; i < sheets.Length; i++)
                {
                    DataTable temp = reader.GetWorksheet(sheets[i]);
                    if (temp.Rows.Count > 1 || temp.Columns.Count > 1)
                    {
                        myDataSet.Tables.Add(temp);
                    }    
                }
                */

                //Spire XLS
                
                Spire.Xls.Workbook wrkbook = new Spire.Xls.Workbook();          //create new workbook
                wrkbook.LoadFromFile(@fullName);                                //load a file
                //Spire.Xls.Worksheet wrksheet = wrkbook.Worksheet[1];           //initialize worksheet
                //myDataSet.Tables.Add(myDataTable);
                Spire.Xls.Collections.WorksheetsCollection wrksheets = wrkbook.Worksheets;

                for (int i = 0; i < wrksheets.Count; i++)
                {
                    DataTable temp = wrksheets[i].ExportDataTable();
                    if (temp.Rows.Count > 1 || temp.Columns.Count > 1)
                    {
                        myDataSet.Tables.Add(temp);
                    }
                    
                }

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


            //Start Correlational Matrix 
            DataTable tempdt = myDataSet.Tables[0].Copy();

            for (int i = 0; i < tempdt.Columns.Count; i++)
            {
                string type = tempdt.Columns[i].DataType.ToString();
                if ((type.Contains(typeof(string).ToString()) || type.Contains(typeof(char).ToString())))
                {
                    tempdt.Columns.Remove(tempdt.Columns[i]);
                }
            }
            
            double[,] tableMatrix = tempdt.ToMatrix();
            double[,] correlMatrix = Accord.Statistics.Tools.Correlation(tableMatrix);
            string[] names = new string[myDataSet.Tables[0].Columns.Count];
            int h = 0;
            TableRow headRow = new TableRow();

            TableCell cornerCell = new TableCell();
            cornerCell.Text = "-";
            cornerCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
            cornerCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
            headRow.Cells.Add(cornerCell);

            foreach (DataColumn col in myDataSet.Tables[0].Columns)
            {
                string type = col.DataType.ToString();
                if (!(type.Contains(typeof(string).ToString()) || type.Contains(typeof(char).ToString())))
                {
                    names[h] = col.ColumnName.ToString();

                    TableCell cell = new TableCell();
                    cell.Text = names[h];
                    cell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                    cell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                    headRow.Cells.Add(cell);

                    h++;
                }
            }

            CorrTable.Rows.Add(headRow);

            for (int i = 0; i < correlMatrix.Rows(); i++)
            {
                double[] row = correlMatrix.GetRow(i);
                int length = row.Length;

                TableRow tR = new TableRow();

                TableCell rowName = new TableCell();
                rowName.Text = names[i];
                rowName.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                rowName.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                tR.Cells.Add(rowName);
                
                for (int j = 0; j < length; j++)
                {
                    TableCell cell = new TableCell();

                    if (j <= i)
                    {
                        cell.Text = row[j].ToString("0.###");
                    }
                    else
                    {
                        cell.Text = "-";
                    }

                    cell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                    cell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                    tR.Cells.Add(cell);
                }
                CorrTable.Rows.Add(tR);
            }
            //End Correlational Matrix

            foreach (DataTable dt in myDataSet.Tables)
            {
                foreach (DataColumn dc in dt.Columns)
                {
                    if (dataTypes.Contains(dc.DataType))
                    {
                        string varName = String.Empty, varType = String.Empty, uniqueVals = "*";
                        double mean = 0, min = 0, max = 0, stdD = 0;
                        TableRow tRow = new TableRow();
                        TableCell meanCell = new TableCell();
                        TableCell minCell = new TableCell();
                        TableCell maxCell = new TableCell();
                        TableCell stdCell = new TableCell();
                        TableCell dependCell = new TableCell();
                        TableCell nameCell = new TableCell();
                        TableCell varTCell = new TableCell();
                        TableCell cardCell = new TableCell();
                        TableCell levelsCell = new TableCell();

                        //Start Dependency Radio List
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
                        //End Dependency Radio List

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
                                logits = (List<string>)ViewState["logits"];
                            }
                            logits.Add(dc.ColumnName.ToString());
                            ViewState["logits"] = logits;
                        }
                        //End Block for calculating Cardinality.

                        if (dc.DataType.ToString() == typeof(char).ToString() || dc.DataType.ToString() == typeof(string).ToString())
                        {
                            varType = "Nominal";
                            levelsCell.Text = "-";
                        }
                        else
                        {
                            varType = "Numeric";

                            //Block for calculating Mean.
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
                        cardCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                        cardCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                        tRow.Cells.Add(dependCell);
                        tRow.Cells.Add(nameCell);
                        tRow.Cells.Add(varTCell);

                        if (!(dc.DataType.ToString() == typeof(char).ToString() || dc.DataType.ToString() == typeof(string).ToString()))
                        {
                            meanCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            meanCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            minCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            minCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            maxCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            maxCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            stdCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            stdCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            tRow.Cells.Add(meanCell);
                            tRow.Cells.Add(minCell);
                            tRow.Cells.Add(maxCell);
                            tRow.Cells.Add(stdCell);
                        }
                        tRow.Cells.Add(cardCell);

                        if (dc.DataType.ToString() == typeof(char).ToString() || dc.DataType.ToString() == typeof(string).ToString())
                        {
                            levelsCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                            levelsCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                            tRow.Cells.Add(levelsCell);
                            Table3.Rows.Add(tRow);
                            Table3.Visible = true;
                        }
                        else
                        {
                            Table1.Rows.Add(tRow);
                        }

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
                        */

                        int colNum = 0, missingV = 0;
                        colNum = dt.Columns.IndexOf(dc);

                        foreach (DataRow dr in dt.Rows)
                        {
                            if (dr[colNum].ToString().Length == 0)
                            {
                                missingV += 1;
                            }
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

            //Numeric table header row
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

            //Nomical table header row
            TableRow hRow4 = new TableRow();

            TableCell dependH2 = new TableCell();
            TableCell nameCellH2 = new TableCell();
            TableCell varTCellH2 = new TableCell();
            TableCell cardCellH2 = new TableCell();
            TableCell levelsCellH = new TableCell();

            dependH2.Text = "Variable Dependency";
            nameCellH2.Text = "Name";
            varTCellH2.Text = "Type";
            cardCellH2.Text = "Cardinality";
            levelsCellH.Text = "Levels";

            hRow4.Cells.Add(dependH2); ;
            hRow4.Cells.Add(nameCellH2);
            hRow4.Cells.Add(varTCellH2);
            hRow4.Cells.Add(cardCellH2);
            hRow4.Cells.Add(levelsCellH);

            Table3.Rows.Add(hRow4);
        }

        void Radio_Changed(object sender, EventArgs e, string col)
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

            foreach (KeyValuePair<string, string> kp in depDic)
            {
                foreach (TableCell tc in CorrTable.Rows[0].Cells)
                {
                    if (!(kp.Value == "i") && tc.Text == kp.Key)
                    {
                        int index = CorrTable.Rows[0].Cells.GetCellIndex(tc);
                        CorrTable.Rows[0].Cells[index].Font.Bold = false;

                        for (int i = 0; i <= index; i++)
                        {
                            CorrTable.Rows[index].Cells[i].Font.Bold = false;
                        }
                        for (int i = index; i < CorrTable.Rows.Count; i++)
                        {
                            CorrTable.Rows[i].Cells[index].Font.Bold = false;
                        }

                    }
                }
            }
            foreach (KeyValuePair<string, string> kp in depDic)
            {
                foreach (TableCell tc in CorrTable.Rows[0].Cells)
                {
                    if (kp.Value == "i" && tc.Text == kp.Key)
                    {
                            int index = CorrTable.Rows[0].Cells.GetCellIndex(tc);
                            CorrTable.Rows[0].Cells[index].Font.Bold = true;

                            for (int i = 0; i <= index; i++)
                            {
                                CorrTable.Rows[index].Cells[i].Font.Bold = true;
                            }
                            for (int i = index; i < CorrTable.Rows.Count; i++)
                            {
                                CorrTable.Rows[i].Cells[index].Font.Bold = true;
                            }
                        
                    }
                }
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