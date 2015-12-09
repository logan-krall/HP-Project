using System;
using System.Data;
using System.Net;

//SQL Server
using System.Data.SqlClient;

//OleDB
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
            Session["radioCount"] = 0;
            List<string> cellTypes = new List<string>();

            //Loads file into DataSet
            DataSet myDataSet = Load_File(new DataSet());

            //Populate table headers
            Table_Header_Init();

            //Populate Correlational Matrix 
            DataTable tempdt = myDataSet.Tables[0].Copy();               
            Build_Corr_Matrix(tempdt);

            //Initialize missing viewstate variable to use for finding missing values in main loop
            ViewState["missing"] = false;
            ViewState["num_nominal"] = 0;
            ViewState["num_numeric"] = 0;

            //Build a list of all acceptable data types for use in main loop type checking
            var dataTypes = new[] { typeof(Byte), typeof(SByte), typeof(Decimal), typeof(Double), typeof(Single), typeof(Int16), 
                typeof(Int32), typeof(Int64), typeof(UInt16), typeof(UInt32), typeof(UInt64), typeof(Char), typeof(string) };

            //Main loop through columns
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
                            int num = (int)ViewState["num_nominal"];
                            ViewState["num_nominal"] = ++num;

                        }
                        else
                        {
                            varType = "Numeric";

                            int num = (int)ViewState["num_numeric"];
                            ViewState["num_numeric"] = ++num;

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

                        if (varType == "Numeric")
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
                            tRow.Cells.Add(cardCell);
                            Table1.Rows.Add(tRow);
                        }
                        else
                        {
                            tRow.Cells.Add(cardCell);
                            Table3.Rows.Add(tRow);
                            Table3.Visible = true;
                        }

                        //Check for missing values in this column
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

        DataSet Load_File(DataSet myDataSet)
        {
            string fullName = (string)(Session["name"]);
            string extension = System.IO.Path.GetExtension(fullName).ToLower();
            DataTable myDataTable = new DataTable();

            if (extension == ".xls" || extension == ".xlsx")
            {
                //Olebdb version
                string connectionString = string.Empty;

                if (extension == ".xls")
                {
                    connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;";
                    connectionString += "Data Source='" + fullName + "';";
                    connectionString += "Extended Properties='Excel 8.0;HDR=YES;IMEX=1;READONLY=TRUE;';";
                }
                else if (extension == ".xlsx")
                {
                    connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;";
                    connectionString += "Data Source='" + fullName + "';";
                    connectionString += "Extended Properties='Excel 12.0 Xml;HDR=YES';";
                }

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

            return myDataSet;
        }

        void Table_Header_Init()
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

            //Nominal table header row
            TableRow hRow4 = new TableRow();

            TableCell dependH2 = new TableCell();
            TableCell nameCellH2 = new TableCell();
            TableCell varTCellH2 = new TableCell();
            TableCell cardCellH2 = new TableCell();

            dependH2.Text = "Variable Dependency";
            nameCellH2.Text = "Name";
            varTCellH2.Text = "Type";
            cardCellH2.Text = "Cardinality";

            hRow4.Cells.Add(dependH2); ;
            hRow4.Cells.Add(nameCellH2);
            hRow4.Cells.Add(varTCellH2);
            hRow4.Cells.Add(cardCellH2);

            Table3.Rows.Add(hRow4);
        }

        void Build_Corr_Matrix(DataTable tempdt)
        {
            for (int i = 0; i < tempdt.Columns.Count; i++)
            {
                string type = tempdt.Columns[i].DataType.ToString();
                if ((type.Contains(typeof(string).ToString()) || type.Contains(typeof(char).ToString())))
                {
                    tempdt.Columns.Remove(tempdt.Columns[i]);
                }
                else
                {
                    foreach (DataRow dr in tempdt.Rows)
                    {
                        if (dr[i].GetType().ToString().Contains("DBNull"))
                        {
                            dr[i] = 0.0;
                        }
                    }
                }
            }

            double[,] tableMatrix = tempdt.ToMatrix();
            double[,] correlMatrix = Accord.Statistics.Tools.Correlation(tableMatrix);
            string[] names = new string[tempdt.Columns.Count];
            int h = 0;
            TableRow headRow = new TableRow();

            TableCell cornerCell = new TableCell();
            cornerCell.Text = "-";
            cornerCell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
            cornerCell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
            headRow.Cells.Add(cornerCell);

            foreach (DataColumn col in tempdt.Columns)
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

        protected void saveButton_Click(object sender, EventArgs e)
        {
            DataSet myDataSet = Load_File(new DataSet());

            string name = Server.MapPath("/Uploads/");
            string file = "Spreadsheet" + DateTime.Now.ToString("yyyy-MM-dd--hh-mm-ss") + ".xls";
            name += file;

            var fileStr = new FileInfo(name);

            //Build a list of all acceptable data types for use in main loop type checking
            var dataTypes = new[] { typeof(Byte), typeof(SByte), typeof(Decimal), typeof(Double), typeof(Single), typeof(Int16), 
                typeof(Int32), typeof(Int64), typeof(UInt16), typeof(UInt32), typeof(UInt64), typeof(Char), typeof(string) };

            using (ExcelPackage p = new ExcelPackage(fileStr))
            {
                string sheetName = "New Sheet";
                p.Workbook.Worksheets.Add(sheetName);
                ExcelWorksheet ws = p.Workbook.Worksheets[1];
                ws.Name = sheetName; //Setting Sheet's name
                ws.Cells.Style.Font.Size = 11; //Default font size for whole sheet
                ws.Cells.Style.Font.Name = "Calibri"; //Default Font name for whole sheet

                //Main loop through columns
                foreach (DataTable dt in myDataSet.Tables)
                {   
                    //ws.InsertColumn(1, dt.Columns.Count);
                    int colNum = 0;

                    foreach (DataColumn dc in dt.Columns)
                    {
                        colNum++;

                        if (dataTypes.Contains(dc.DataType))
                        {
                            var hcell = ws.Cells[1, colNum];
                            hcell.Value = dc.ColumnName;

                            int rowIndex = 1;

                            if (dc.DataType.ToString() != typeof(char).ToString() && dc.DataType.ToString() != typeof(string).ToString())
                            {
                                //Block for calculating Mean.
                                object meanObject;
                                meanObject = dt.Compute("Avg(" + dc.ColumnName.ToString() + ")", string.Empty);
                                double mean = Double.Parse(meanObject.ToString());

                                //Check for missing values in this column
                                foreach (DataRow dr in dt.Rows)
                                {
                                    rowIndex++;
                                    if (dr[colNum].ToString().Length == 0)
                                    {
                                        dr[colNum] = mean;
                                    }
                                    //Find corresponding cell in worksheet
                                    var cell = ws.Cells[rowIndex, colNum];

                                    //Setting Value in cell
                                    //cell.Value = Convert.ToInt32(dr[dc.ColumnName]);
                                    cell.Value = dr[dt.Columns.IndexOf(dc)];
                                }
                            }
                            else
                            {
                                //Check for missing values in this column
                                foreach (DataRow dr in dt.Rows)
                                {
                                    rowIndex++;
                                    //Find corresponding cell in worksheet
                                    var cell = ws.Cells[rowIndex, colNum];

                                    //Setting Value in cell
                                    cell.Value = dr[dt.Columns.IndexOf(dc)];
                                }
                            }
                            //-------- Now leaving if acceptable data type block
                        }
                        //-------- Now leaving the for each datacolumns block    
                    }
                    //-------- Now leaving the for each datatable block
                }
                p.Save();
                //---------- Now leaving the using statement
            }

            //Create and Save Excel file to client desktop   
            UploadStatusLabel.Text = "File download started.";

            //string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string baseUrl = Request.Url.Scheme + "://" + Request.Url.Authority + Request.ApplicationPath.TrimEnd('/') + "/";

            //Clear the response               
            Response.Clear();
            Response.ClearContent();
            Response.ClearHeaders();
            Response.Cookies.Clear();

            Response.ContentType = "application/vnd.ms-excel";
            Response.AppendHeader("Content-Disposition", "attachment; filename=" + file);
            Response.TransmitFile(name);
            Response.End();            
        }
    }
}