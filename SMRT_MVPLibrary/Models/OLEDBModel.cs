using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using OfficeOpenXml;


namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class OLEDBModel : ISMRTDataModel
    {

        private OleDbConnection connGlobal;
        private bool useJetEngine;

        private string connectionStringJet = "provider=Microsoft.Jet.OLEDB.4.0; data source='{0}'; Extended Properties='Excel 8.0;'";
        // The following is supposedly the one to use for 2007+, but the driver doesn't work on 64-bit without a patch.
        // Downloaded the MS Access Database Engine 2010 Redist: http://www.microsoft.com/en-us/download/details.aspx?id=13255
        private string connectionStringACE = "provider=Microsoft.ACE.OLEDB.12.0; data source='{0}'; Extended Properties=Excel 12.0;";

        public OLEDBModel(bool _useJetEngine)
        {
            useJetEngine = _useJetEngine;
        }

        public void Dispose()
        {
            if (connGlobal != null)
            {
                Disconnect();
            }
        }

        private OleDbConnection Connect(string fileName)
        {
            if (connGlobal == null)
            {
                connGlobal = new OleDbConnection(useJetEngine ? connectionStringJet : connectionStringACE);
                connGlobal.Open();
            }
            else if (!connGlobal.ConnectionString.Equals(useJetEngine ? connectionStringJet : connectionStringACE))
            {
                Disconnect();
                connGlobal = new OleDbConnection(useJetEngine ? connectionStringJet : connectionStringACE);
                connGlobal.Open();
            }

            return connGlobal;
        }

        private void Disconnect()
        {
            //conn.Close();
            if (connGlobal != null)
            {
                connGlobal.Dispose();
                connGlobal = null;
            }
            OleDbConnection.ReleaseObjectPool();
        }

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName))
            {
                string connectionString = useJetEngine ? connectionStringJet : connectionStringACE;
                connectionString = String.Format(connectionString, fileName);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    //Connect(filename);

                    //DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                    DataTable schemaTable = conn.GetSchema("Tables");

                    foreach (DataRow row in schemaTable.Rows)
                    {
                        string sheetName = row.Field<string>("TABLE_NAME");
                        if (sheetName.EndsWith("$") || sheetName.EndsWith("$'"))
                        {
                            if (sheetName[0] == '\'' && sheetName.EndsWith("'"))
                                sheetName = sheetName.Substring(1, sheetName.Length - 2);
                            sheetNames.Add(sheetName);
                        }
                    }
                }
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                string connectionString = useJetEngine ? connectionStringJet : connectionStringACE;
                connectionString = String.Format(connectionString, fileName);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    //Connect(fileName);

                    string query = "select * from [" + sheetName + "]";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    foreach (DataColumn column in table.Columns)
                        columnNames.Add(column.ColumnName, column.ColumnName);
                }
            }

            return columnNames;
        }

        public List<KeyValuePair<string, string>> GetColumnValues(string fileName, string sheetName, string columnName, bool ignoreFirstRow)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnName))
            {
                string connectionString = useJetEngine ? connectionStringJet : connectionStringACE;
                connectionString = String.Format(connectionString, fileName);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    //Connect(fileName);

                    string query = "select [" + columnName + "] from [" + sheetName + "]";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                        columnValues.Add(new KeyValuePair<string, string>(row[0].ToString(), row[0].ToString()));
                }
            }

            TestWriting(fileName, sheetName, columnName);

            return columnValues;
        }

        public void TestWriting(string fileName, string sheetName, string columnName)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnName))
            {
                //AddColumnToFile(fileName, sheetName, "NewColumn");
                string connectionString = useJetEngine ? connectionStringJet : connectionStringACE;
                connectionString = String.Format(connectionString, fileName);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    //Connect(fileName);

                    string query = "select top 1 * from [" + sheetName + "]";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                    DataTable table = new DataTable();
                    adapter.Fill(table);

                    foreach (DataRow row in table.Rows)
                        columnValues.Add(new KeyValuePair<string, string>(row[0].ToString(), row[0].ToString()));


                    table.Columns.Add("NewColumn");
                    table.Rows[0][table.Rows[0].ItemArray.Length - 1] = "ABCXYZ";
                    adapter.UpdateCommand = new OleDbCommand("UPDATE [" + sheetName + "] SET NewColumn = ? WHERE " + columnName + " = ?", conn);
                    //adapter.UpdateCommand = new OleDbCommand("UPDATE [" + sheetName + "] SET [F5] = 'xxx' WHERE [Snippet] = ?", conn);
                    adapter.UpdateCommand.Parameters.Add("@NewColumn", OleDbType.VarChar, 255).SourceColumn = "NewColumn";
                    adapter.UpdateCommand.Parameters.Add("@" + columnName, OleDbType.VarChar, 255).SourceColumn = columnName;
                    adapter.Update(table);
                }

            }

        }

 
        private void AddColumnToFile(string fileName, string sheetName, string columnName)
        {
            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnName))
            {
                string connectionString = useJetEngine ? connectionStringJet : connectionStringACE;
                connectionString = String.Format(connectionString, fileName);
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    //Connect(fileName);

                    OleDbCommand cmd = new OleDbCommand();
                    cmd.Connection = conn;
                    cmd.CommandText = "UPDATE [" + sheetName + "G1:G1] SET F1 = '" + columnName + "'";
                    cmd.ExecuteNonQuery();
                }
            }

        }

        public ReturnCode AddColumns(string fileName, string sheetName, string[] columnNames)
        {
            throw new NotImplementedException();
        }


        public bool FileIsValid(string fileName)
        {
            throw new NotImplementedException();
        }


        public ReturnCode WriteColumnValues(string fileName, string sheetName, string[] columnNames, List<MentionPart> newValues, int firstRow)
        {
            throw new NotImplementedException();
        }
    }
}
