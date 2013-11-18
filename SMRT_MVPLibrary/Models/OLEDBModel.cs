using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using TwinArch.SMRT_MVPLibrary.Interfaces;


namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class OLEDBModel : ISMRTDataModel
    {

        private OleDbConnection conn;
        private bool useJetEngine;

        public OLEDBModel(bool _useJetEngine)
        {
            useJetEngine = _useJetEngine;
        }

        public void Dispose()
        {
            if (conn != null)
                conn.Close();
        }

        private OleDbConnection Connect(string fileName)
        {
            string connectionStringJet = "provider=Microsoft.Jet.OLEDB.4.0; data source='" + fileName + "'; Extended Properties='Excel 8.0;'";
            // The following is supposedly the one to use for 2007+, but the driver doesn't work on 64-bit without a patch.
            // Downloaded the MS Access Database Engine 2010 Redist: http://www.microsoft.com/en-us/download/details.aspx?id=13255
            string connectionStringACE = "provider=Microsoft.ACE.OLEDB.12.0; data source=" + fileName + "; Extended Properties=Excel 12.0;";
            if (conn == null)
            {
                conn = new OleDbConnection(useJetEngine ? connectionStringJet : connectionStringACE);
                conn.Open();
            }
            else if (!conn.ConnectionString.Equals(useJetEngine ? connectionStringJet : connectionStringACE))
            {
                conn.Close();
                conn = new OleDbConnection(useJetEngine ? connectionStringJet : connectionStringACE);
                conn.Open();
            }

            return conn;
        }

        public List<string> GetSheetNames(string filename)
        {
            List<string> sheetNames = new List<string>();

            if (!String.IsNullOrEmpty(filename))
            {
                Connect(filename);

                //DataTable schemaTable = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                DataTable schemaTable = conn.GetSchema("Tables");

                foreach (DataRow row in schemaTable.Rows)
                {
                    string sheetName = row.Field<string>("TABLE_NAME");
                    if (sheetName.EndsWith("$") || sheetName.EndsWith("$'"))
                    {
                        if (sheetName[0]=='\'' && sheetName.EndsWith("'"))
                            sheetName = sheetName.Substring(1, sheetName.Length - 2);
                        sheetNames.Add(sheetName);
                    }
                }
                conn.Close();
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                Connect(fileName);

                string query = "select * from [" + sheetName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable table = new DataTable();
                adapter.Fill(table);

                foreach (DataColumn column in table.Columns)
                    columnNames.Add(column.ColumnName, column.ColumnName);
                conn.Close();
            }

            return columnNames;
        }

        public List<KeyValuePair<string, string>> GetColumnValues(string fileName, string sheetName, string columnName)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnName))
            {
                Connect(fileName);

                string query = "select " + columnName + " from [" + sheetName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable table = new DataTable();
                adapter.Fill(table);

                foreach (DataRow row in table.Rows)
                    columnValues.Add(new KeyValuePair<string,string>(row[0].ToString(), row[0].ToString()));

                conn.Close();
            }

            return columnValues;
        }
    }
}
