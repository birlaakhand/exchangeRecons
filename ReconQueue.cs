using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Data.OleDb;
using System.Data.Sql;
using System.Windows;
using Microsoft.VisualBasic.FileIO;

namespace ExchangeRecon
{
    class ReconQueue
    {
        //Class Properties
        public string filepath { get; set; }
        public DataTable QData { get; set; }
        
        //CONSTRUCTORS
        public ReconQueue(DataTable inp, string outFile)
        {
            QData = inp;
            filepath = outFile;
            ToCSV();
        }

        public ReconQueue(string filename) : this(filename, "", filename) { }

        public ReconQueue(string inFile, string sheetName, string outFile)
        {
            if (File.Exists(inFile))
                if (String.Compare(inFile.Substring(inFile.IndexOf('.')), ".csv", true) == 0)
                    QData = ReconQueue.FromCSV(inFile);
                else if (sheetName.Length > 0 && String.Compare(inFile.Substring(inFile.IndexOf('.')), ".xls", true) == 0)
                {
                    OleDbConnection con = new OleDbConnection(
                        "provider=Microsoft.Jet.OLEDB.4.0;data source=" + inFile
                        + ";Extended Properties='Excel 8.0;HDR=Yes;'");
                    //    "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "
                    //    + filepath + "; Extended Properties = Excel 12.0; ");
                    //@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;'"
                    StringBuilder stbQuery = new StringBuilder();
                    stbQuery.Append("SELECT * FROM [" + sheetName + "]");
                    OleDbDataAdapter adp = new OleDbDataAdapter(stbQuery.ToString(), con);

                    DataSet dsXLS = new DataSet();
                    adp.Fill(dsXLS);
                    QData = dsXLS.Tables[0];
                    con.Close();
                }
                else QData = new DataTable();
            else QData = new DataTable();

            filepath = outFile;
            ToCSV();
        }

        //Class Methods
        public void ToCSV()
        {

            if (filepath.Length > 0 && QData != null)
            {
                Console.WriteLine("Converting to CSV : " + filepath);
                File.WriteAllLines(filepath, ReconQueue.ToCSV(QData));
                Console.WriteLine("Converted to CSV : " + filepath);
            }
        }

        private OleDbType GetOleDbType(Type sysType)
        {
            //Console.WriteLine(sysType.ToString());
            if (sysType == typeof(System.String) || sysType == typeof(string))
                return OleDbType.VarChar;
            else if (sysType == typeof(System.Int16) || sysType == typeof(System.Int32) || sysType == typeof(int))
                return OleDbType.Integer;
            else if (sysType == typeof(bool))
                return OleDbType.Boolean;
            else if (sysType == typeof(DateTime))
                return OleDbType.Date;
            else if (sysType == typeof(char))
                return OleDbType.Char;
            else if (sysType == typeof(decimal))
                return OleDbType.Decimal;
            else if (sysType == typeof(Single))
                return OleDbType.Single;
            else if (sysType == typeof(byte))
                return OleDbType.Binary;
            else if (sysType == typeof(Guid))
                return OleDbType.Guid;
            else if (sysType == typeof(System.Double))
                return OleDbType.Double;
            else
                return OleDbType.VarChar;
        }

        //STATIC Methods
        public static DataTable FromCSV(string fileName)
        {
            return CSVtoDataTable(fileName);
            /*
            string[] rows = File.ReadAllLines(fileName);
            DataTable t = new DataTable();
            string[] colHeaders = rows[0].Split(',');
            foreach(string colHeader in colHeaders)
                t.Columns.Add(colHeader);
            
            for(int i=1; i < rows.Length; i++)
                t.Rows.Add(rows[i].Split(','));

            return t;
            */
        }

        public static DataTable CSVtoDataTable(string filepath)
        {
            int count = 0;
            //char fieldSeparator = ',';
            DataTable csvData = new DataTable();

            using (TextFieldParser csvReader = new TextFieldParser(filepath))
            {
                Console.WriteLine("Reading CSV File " + filepath);
                csvReader.HasFieldsEnclosedInQuotes = true;
                while (!csvReader.EndOfData)
                {
                    csvReader.SetDelimiters(new string[] { "," });
                    string[] fieldData = csvReader.ReadFields();
                    if (count == 0)
                    {
                        foreach (string column in fieldData)
                        {
                            DataColumn datecolumn = new DataColumn(column);
                            datecolumn.AllowDBNull = true;
                            csvData.Columns.Add(datecolumn);
                        }
                    }
                    else
                    {
                        //Console.WriteLine(csvData.Columns.Count + "<- Col count, row len -> " + fieldData.Length);
                        csvData.Rows.Add(fieldData);
                    }
                    count++;

                }
            }
            return csvData;

        }

        public static string[] ToCSV(DataTable QData)
        {
            string[] res = new string[QData.Rows.Count + 1];
            //MessageBox.Show("Col Count : " + QData.Columns.Count);
            for (int i = 0; i < QData.Columns.Count; i++)
            {
                res[0] += "\"" + QData.Columns[i].ColumnName + "\"";
                res[0] += i + 1 < QData.Columns.Count ? "," : "";
            }

            int j = 1;
            foreach (DataRow row in QData.Rows)
            {
                for (int i = 0; i < QData.Columns.Count; i++)
                {
                    //if (row[i].GetType() == typeof(System.String))
                        res[j] += "\"" + row[i] + "\"";
                    //else
                    //    res[j] += row[i];
                    res[j] += i+1<QData.Columns.Count? ",":"";
                    //result.Append(row[i].ToString());
                    //result.Append(i == QData.Columns.Count - 1 ? "\n" : ",");
                }
                j++;
            }

            return res;
        }

        //DESTRUCTORS
        ~ReconQueue()
        {
            //adp.DeleteCommand = new OleDbCommand("Delete * FROM ["+sheetname+"]");
            //adp.InsertCommand;
            //adp.Update(QData);
        }

        //OLD CODE : Trying to Implement OleDBConnections for Excel Files
        /*
        //public string sheetname { get; set; }
        //public DataSet dsXLS { get; set; }
        //private OleDbConnection con;
        //public OleDbDataAdapter adp;

        public ReconQueue()
        {
            filepath = @"C:\Users\Akhand\Downloads\DealReport2.xls";
            sheetname = @"Collated$";
            con = new OleDbConnection(
            //    "Provider = Microsoft.ACE.OLEDB.12.0; Data Source = "
            //    + filepath + "; Extended Properties = Excel 12.0; ");
                "provider=Microsoft.Jet.OLEDB.4.0;data source=" + filepath
                + ";Extended Properties='Excel 8.0;HDR=Yes;'");
            //@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;'"
            StringBuilder stbQuery = new StringBuilder();
            stbQuery.Append("SELECT * FROM [" + sheetname + "]");
            adp = new OleDbDataAdapter(stbQuery.ToString(), con);

            dsXLS = new DataSet();
            adp.Fill(dsXLS);


            QData = dsXLS.Tables[0];
            QData.Columns["Pkey"].Unique = true;
            QData.PrimaryKey = new DataColumn[] { QData.Columns["Pkey"] };
            generateParams();
            
            Console.WriteLine("====================================  "+adp.UpdateCommand.ToString());
            
        }
        public void generateParams()
        {
            //UPDATE COMMAND
            OleDbCommand command;
            int i = 0;
            string updateString = "UPDATE ["+sheetname+"] set ";
            string whereString = " WHERE ";
            foreach (DataColumn dc in QData.Columns)
            {
                updateString += i == 0 ? "" : ", ";
                whereString += i == 0 ? "" : " AND ";
                //updateString += "\"" + dc.ColumnName + "\" = ? ";
                //whereString += "\"" + dc.ColumnName + "\" = ? ";
                updateString += "[" + dc.ColumnName + "] = ? ";
                whereString += "[" + dc.ColumnName + "] = ? ";
                i++;
            }
            command = new OleDbCommand(updateString + whereString,con);
            Console.WriteLine(command.CommandText);
            i = 0;
            foreach (DataColumn dc in QData.Columns)
            {
                Console.WriteLine(QData.Columns[i].ColumnName + " : " + QData.Columns[i].DataType.ToString());
                //command.Parameters.Add("@" + dc.ColumnName, GetOleDbType(QData.Columns[i].DataType)).SourceColumn = dc.ColumnName;
                //command.Parameters.Add("@Old" + dc.ColumnName, GetOleDbType(QData.Columns[i].DataType), 255, dc.ColumnName).SourceVersion = DataRowVersion.Original;
                command.Parameters.Add("@" + dc.ColumnName, OleDbType.Char,255).SourceColumn = dc.ColumnName;
                command.Parameters.Add("@Old" + dc.ColumnName, OleDbType.Char, 255, dc.ColumnName).SourceVersion = DataRowVersion.Original;
                command.Parameters.Add("@Old" + dc.ColumnName, OleDbType.Char, 255, dc.ColumnName).SourceVersion = DataRowVersion.Original;
                i++;
            }
            
            adp.UpdateCommand = command;


            //INSERT COMMAND
            i = 0;
            updateString = "INSERT INTO [" + sheetname + "] (";
            whereString = " VALUES( ";
            foreach (DataColumn dc in QData.Columns)
            {
                updateString += i == 0 ? "" : ", ";
                whereString += i == 0 ? "" : " , ";
                //updateString += "\"" + dc.ColumnName + "\" = ? ";
                //whereString += "\"" + dc.ColumnName + "\" = ? ";
                updateString += "[" + dc.ColumnName + "]";
                whereString += " ? ";
                i++;
            }
            command = new OleDbCommand(updateString + ") " + whereString + ")", con);
            Console.WriteLine(command.CommandText);
            i = 0;
            foreach (DataColumn dc in QData.Columns)
            {
                Console.WriteLine(QData.Columns[i].ColumnName + " : " + QData.Columns[i].DataType.ToString());
                //command.Parameters.Add("@" + dc.ColumnName, GetOleDbType(QData.Columns[i].DataType)).SourceColumn = dc.ColumnName;
                command.Parameters.Add("@" + dc.ColumnName, OleDbType.Char, 255).SourceColumn = dc.ColumnName;
                i++;
            }

            adp.InsertCommand = command;

        }
        
        */

    }
}
