using System;
using System.Data;
using System.Data.OleDb;
using System.Xml.Serialization;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;


namespace HP_Analytics_Project.Images
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            string path = Server.MapPath("~/Uploads/");
            string fullName = path + (string)(Session["name"]);
            string extension = System.IO.Path.GetExtension(fullName).ToLower();
            string connectionString = string.Empty;
            OleDbCommand cmd = new OleDbCommand();
            DataSet myDataSet = new DataSet();
            DataTable myDataTable = new DataTable();

            if (extension == ".xls")
            {
                connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;";
                connectionString += "Data Source='" + fullName + "';";
                connectionString += "Extended Properties='Excel 8.0;HDR=YES;IMEX=1;READONLY=TRUE;';";
            }
            else if (extension == ".xlsx")
            {
                connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;";
                connectionString += "Data Source=" + fullName + ";";
                connectionString += "Extended Properties='Excel 12.0;HDR=YES;IMEX=1;READONLY=TRUE;';";
            }
            else if (extension == ".csv")
            {
                connectionString = "Provider=;";
                connectionString += "Data Source=" + fullName + ";";
            }

            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
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
                //File.Delete(fullName);

                var dataTypes = new[] { typeof(Byte), typeof(SByte), typeof(Decimal), typeof(Double), typeof(Single), typeof(Int16), 
                    typeof(Int32), typeof(Int64), typeof(UInt16), typeof(UInt32), typeof(UInt64), typeof(Char), typeof(string) };

                //missing value header row
                TableRow hRow2 = new TableRow();
                TableRow hRow3 = new TableRow();

                TableCell missNCellH = new TableCell();
                TableCell missVCellH = new TableCell();
                missNCellH.Text = "Name of Column";
                missVCellH.Text = "Rows Missing";
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

                hRow.Cells.Add(dependH);;
                hRow.Cells.Add(nameCellH);
                hRow.Cells.Add(varTCellH);
                hRow.Cells.Add(meanCellH);
                hRow.Cells.Add(medCellH);
                hRow.Cells.Add(moCellH);
                hRow.Cells.Add(stdCellH);
                hRow.Cells.Add(cardCellH);

                Table1.Rows.Add(hRow);

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

                            RadioButton ind1 = new RadioButton();
                            RadioButton dep1 = new RadioButton();
                            RadioButtonList depend1 = new RadioButtonList();
                            
                            //dependency radio list
                            depend1.ID = dc.ColumnName.ToString();
                            depend1.AutoPostBack = true;
                            depend1.SelectedIndexChanged += new EventHandler((s, e1) => Radio_Changed(s, e1, dc.ColumnName.ToString()));
                            depend1.RepeatDirection = System.Web.UI.WebControls.RepeatDirection.Horizontal;
                            depend1.Font.Size = System.Web.UI.WebControls.FontUnit.XSmall;
                            ListItem ind = new ListItem();
                            ListItem dep = new ListItem();
                            ind.Text = "  Independent";
                            dep.Text = "  Dependent";
                            ind.Attributes.Remove("font-weight");
                            ind.Value = "i";
                            dep.Value = "d";
                            depend1.Items.Add(ind);
                            depend1.Items.Add(dep);
                            dependCell.Controls.Add(depend1);

                            DataTable catVals = dt.DefaultView.ToTable(true, dc.ColumnName.ToString());
                            uniqueVals = catVals.Rows.Count.ToString();

                            if (dc.DataType.ToString() == typeof(char).ToString() || dc.DataType.ToString() == typeof(string).ToString())
                            {
                                meanCell.Text = "*";
                                minCell.Text = "*";
                                maxCell.Text = "*";
                                stdCell.Text = "*";
                                varType = "String";
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
                                if (Table2.Rows[1].Cells[0].Text == "-")
                                {
                                    Table2.Rows.Remove(Table2.Rows[1]);
                                }
                                TableRow missRow = new TableRow();
                                TableCell missVCell = new TableCell();
                                TableCell missNCell = new TableCell();
                                missVCell.Text = missingV.ToString();
                                missNCell.Text = nameCell.Text;
                                missRow.Cells.Add(missNCell);
                                missRow.Cells.Add(missVCell);

                                Table2.Rows.Add(missRow);
                            }
                        }
                    }
                }
            }
        }

        void Radio_Changed(object sender, EventArgs e, string col)
        {
            Dictionary<string, string> depDic = new Dictionary<string, string>();

            if (ViewState["dict"] != null)
            { depDic = (Dictionary<string, string>)ViewState["dict"]; }

            RadioButtonList rb1 = (sender as RadioButtonList);
            if (depDic.ContainsKey(col))
            {   depDic[col] = rb1.SelectedItem.Value; }
            else
            {   depDic.Add(col, rb1.SelectedItem.Value); }

            List<string> main = new List<string>();
            List<string> dep = new List<string>();

            foreach ( KeyValuePair<string,string> kp in depDic)
            {
                if (kp.Value == "i")
                {   main.Add(kp.Key); }
                else
                {   dep.Add(kp.Key); }
            }
            main.Sort();
            dep.Sort();

            foreach ( string v in dep )
            {   main.Add(v); }

            TableRow trH = new TableRow();
            TableCell corner = new TableCell();
            corner.Text = "-";
            trH.Cells.Add(corner);
            Table4.Rows.Add(trH);

            foreach ( string v in main )
            {
                TableCell tcH1 = new TableCell();
                tcH1.Text = v;
                tcH1.HorizontalAlign = HorizontalAlign.Center;
                if (depDic[v] == "i")
                {   tcH1.Font.Bold = true; }

                tcH1.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                tcH1.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                trH.Cells.Add(tcH1);                

                TableRow tr = new TableRow();
                TableCell tc1 = new TableCell();
                tc1.Text = v;
                if (depDic[v] == "i")
                { tc1.Font.Bold = true; }
                tc1.HorizontalAlign = HorizontalAlign.Center;
                tc1.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                tc1.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                tr.Cells.Add(tc1);

                for (int i = 0; i < main.Count; i++)
                {
                    TableCell cell = new TableCell();

                    if (i == main.IndexOf(v))
                    {   cell.Text = "1"; }
                    else
                    {   cell.Text = "-"; }

                    cell.HorizontalAlign = HorizontalAlign.Center;
                    cell.BorderStyle = System.Web.UI.WebControls.BorderStyle.Solid;
                    cell.BorderWidth = System.Web.UI.WebControls.Unit.Pixel(1);
                    tr.Cells.Add(cell);
                }
                Table4.Rows.Add(tr);
            }
            ViewState["dict"] = depDic;
        }
    }
}