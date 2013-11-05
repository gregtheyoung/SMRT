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
    public class OLEDBModel : ISMRTModel
    {

        private OleDbConnection conn;

        private OleDbConnection Connect(string fileName)
        {
            string connectionString = "provider=Microsoft.Jet.OLEDB.4.0; data source=" + fileName + "; Extended Properties=Excel 8.0;";
            if (conn == null)
            {
                conn = new OleDbConnection(connectionString);
                conn.Open();
            }
            else if (!conn.ConnectionString.Equals(connectionString))
            {
                conn.Close();
                conn = new OleDbConnection(connectionString);
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
            }
                            
            return sheetNames;
        }

        public List<string> GetColumnNames(string fileName, string sheetName)
        {
            List<string> columnNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                Connect(fileName);

                string query = "select * from [" + sheetName + "]";
                OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
                DataTable table = new DataTable();
                adapter.Fill(table);

                foreach (DataColumn column in table.Columns)
                    columnNames.Add(column.ColumnName);
            }

            return columnNames;
        }
    }
}
