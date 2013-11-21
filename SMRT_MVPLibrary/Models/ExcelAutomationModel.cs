using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class ExcelAutomationModel : ISMRTDataModel
    {

        private Application app;

        private Application ExcelApp
        {
            get
            {
                if (app == null)
                    app = new Application();
                return app;
            }
        }

        public void Dispose()
        {
            if (app != null)
                app.Quit();
        }

        private Workbook OpenWorkbook(string fileName)
        {
            Workbook book = null;

            if (!String.IsNullOrEmpty(fileName))
                book = ExcelApp.Workbooks.Open(fileName);

            return book;
        }

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            Workbook book = OpenWorkbook(fileName);
            if (book != null)
            {
                foreach (Worksheet sheet in book.Sheets)
                    sheetNames.Add(sheet.Name);

                book.Close(false);
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            Workbook book = OpenWorkbook(fileName);
            if ((book != null) && !String.IsNullOrEmpty(sheetName))
            {
                Worksheet sheet = book.Sheets[sheetName];
                Range firstRow = sheet.get_Range("A1", sheet.get_Range("A1").get_End(XlDirection.xlToRight));

                foreach (Range cell in firstRow.Cells)
                    columnNames.Add(cell.get_Address(Type.Missing, false).Substring(0,1), cell.get_Value().ToString());

                book.Close(false);
            }

            return columnNames;
        }


        public List<KeyValuePair<string, string>> GetColumnValues(string fileName, string sheetName, string columnName, bool ignoreFirstRow)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            Workbook book = OpenWorkbook(fileName);

            if ((book != null) && !String.IsNullOrEmpty(sheetName) && !String.IsNullOrEmpty(columnName))
            {
                Worksheet sheet = book.Sheets[sheetName];
                Range range = sheet.get_Range(columnName + ":" + columnName);

                foreach (Range cell in range.Cells)
                {
                    if (cell.Value2 != null)
                        columnValues.Add(new KeyValuePair<string, string>(cell.get_Address(), cell.Value2));
                }

                book.Close(false);
            }

            return columnValues;
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


// Just need somewhere to put this sample code. Ended up being too slow.


//string connectionStringACE = "provider=Microsoft.ACE.OLEDB.12.0; data source='" + fileName + "'; Extended Properties='Excel 12.0;HDR=YES'";
//OleDbConnection conn = new OleDbConnection(connectionStringACE);
//conn.Open();
//string query = "select ID, MentionType, [Domain], PosterID, MentionID from [" + sheetName + "$]";
//OleDbDataAdapter adapter = new OleDbDataAdapter(query, conn);
//System.Data.DataTable table = new System.Data.DataTable();
//adapter.Fill(table);

//foreach (KeyValuePair<string, List<string>> columnSet in newValues)
//{
//    for (int i = 0; i < columnSet.Value.Count; i++)
//    {
//        table.Rows[i][columnSet.Key] = columnSet.Value[i];
//    }
//}

//adapter.UpdateCommand = new OleDbCommand("UPDATE [" + sheetName + "$] SET MentionType = ?, [Domain] = ?, PosterID = ?, MentionID = ? WHERE ID = ?" , conn);
//adapter.UpdateCommand.Parameters.Add("@MentionType", OleDbType.VarChar, 255).SourceColumn = "MentionType";
//adapter.UpdateCommand.Parameters.Add("@Domain", OleDbType.VarChar, 255).SourceColumn = "Domain";
//adapter.UpdateCommand.Parameters.Add("@PosterID", OleDbType.VarChar, 255).SourceColumn = "PosterID";
//adapter.UpdateCommand.Parameters.Add("@MentionID", OleDbType.VarChar, 255).SourceColumn = "MentionID";
//adapter.UpdateCommand.Parameters.Add("@ID", OleDbType.VarChar, 255).SourceColumn = "ID";
//adapter.Update(table);

//OleDbCommand cmd;
//for (int i = 0; i < newValues["MentionType"].Count; i++)
//{
//    string cellAddress = "G" + (i + 2);
//    cmd = new OleDbCommand("UPDATE [" + sheetName + "$" + cellAddress + ":" + cellAddress + "] SET F1='" + newValues["MentionType"][i] + "'", conn);
//    cmd.ExecuteNonQuery();
//}