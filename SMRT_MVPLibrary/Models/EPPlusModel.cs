using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using OfficeOpenXml;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using TweetinCore;
using Tweetinvi;
using TwitterToken;
using System.Configuration;


namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class EPPlusModel : ISMRTDataModel
    {
        string MaxColumnID = "XFD"; // With Excel 2010+ there are 16384 columns available
        string MaxRowID = "1000000"; // With Eccel 2010+ there are a little more than 1 million rows available

        ExcelPackage _pkg = null;
        TweetinCore.Interfaces.TwitterToken.IToken _token = null;

        public void Dispose()
        {
            PackageDispose();
        }

        private ExcelPackage Package(string fileName)
        {
            if ((_pkg == null) || (_pkg.File.FullName != fileName))
            {
                if (_pkg != null)
                    _pkg.Dispose();
                _pkg = new ExcelPackage(new FileInfo(fileName));
            }

            return _pkg;
        }

        private void PackageDispose()
        {
            if (_pkg != null)
            {
                _pkg.Dispose();
                _pkg = null;
            }
        }

        private TweetinCore.Interfaces.TwitterToken.IToken Token
        {
            get
            {
                if (_token == null)
                {
                    _token = new Token(
                        ConfigurationManager.AppSettings["token_AccessToken"],
                        ConfigurationManager.AppSettings["token_AccessTokenSecret"],
                        ConfigurationManager.AppSettings["token_ConsumerKey"],
                        ConfigurationManager.AppSettings["token_ConsumerSecret"]);
                    TokenSingleton.Token = _token;
                }

                return _token;
            }
        }

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName))
            {
                ExcelPackage pkg = Package(fileName);
                foreach (ExcelWorksheet sheet in pkg.Workbook.Worksheets)
                {
                    string sheetName = sheet.Name;
                    sheetNames.Add(sheetName);
                }
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];
                ExcelRange firstRow = sheet.Cells["A1:"+MaxColumnID+"1"];
                foreach (ExcelRangeBase cell in firstRow)
                    columnNames.Add(cell.Address, cell.Text);
                firstRow.Dispose();
            }

            return columnNames;
        }

        public List<KeyValuePair<string, string>> GetColumnValuesForColumnID(string fileName, string sheetName, string columnID, bool firstRowHasHeaders)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnID))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                ExcelRange cells = sheet.Cells[columnID + ":" + columnID[0] + MaxRowID];
                foreach (ExcelRangeBase cell in cells)
                    columnValues.Add(new KeyValuePair<string, string>(cell.Address, cell.Text));
                cells.Dispose();
                if (firstRowHasHeaders)
                    columnValues.RemoveAt(0);
            }

            return columnValues;
        }

        public List<KeyValuePair<string, string>> GetColumnValuesForColumnName(string fileName, string sheetName, string columnName, bool firstRowHasHeaders)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnName))
            {
                Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);
                foreach (KeyValuePair<string, string> column in columnNames)
                {
                    if (column.Value.Equals(columnName))
                    {
                        string columnID = column.Key;
                        columnValues = GetColumnValuesForColumnID(fileName, sheetName, columnID, firstRowHasHeaders);
                    }
                }
            }

            return columnValues;
        }

        public ReturnCode AddColumns(string fileName, string sheetName, string[] columnNames)
        {
            ReturnCode rc = ReturnCode.Failed;

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && (columnNames.Length >= 0))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                ExcelRange firstRow = sheet.Cells["A1:"+MaxColumnID+"1"];
                bool columnAdded = false;
                foreach (string columnName in columnNames)
                {
                    if (firstRow.FirstOrDefault<ExcelRangeBase>(cell => cell.Text.Equals(columnName)) == null)
                    {
                        int columnNumber = firstRow.Count<ExcelRangeBase>() + 1;
                        sheet.Cells[1, columnNumber].Value = columnName;
                        columnAdded = true;
                    }
                }
                firstRow.Dispose();

                if (columnAdded)
                {
                    pkg.Save();
                    PackageDispose();
                }
                rc = ReturnCode.Success;
            }
        

            return rc;
        }


        public bool FileIsValid(string fileName)
        {
            bool isValid = false;

            if (!String.IsNullOrEmpty(fileName))
            {
                try
                {
                    ExcelPackage pkg = Package(fileName);
                    isValid = true;
                }
                catch (Exception e)
                {
                }
            }

            return isValid;
        }



        public ReturnCode WriteColumnValues(string fileName, string sheetName, System.Data.DataTable newValuesTable, bool firstRowHasHeaders)
        {
            string tempSheetName = "SMRT_Work";

            ReturnCode rc = ReturnCode.Failed;

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                // If the temp worksheet already exists, delete it so we can start from scratch.
                // This is rather important because EPPlus is really fast at writing
                // into empty space. But overwriting existing values or writing in new columns
                // where there are already values in the rows is really, really slow.
                if (pkg.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals(tempSheetName)) != null)
                    pkg.Workbook.Worksheets.Delete(tempSheetName);
                ExcelWorksheet tempWorkSheet = pkg.Workbook.Worksheets.Add(tempSheetName);

                // Load from the datatable into the worksheet.
                tempWorkSheet.Cells["A1"].LoadFromDataTable(newValuesTable, true);

                // Save it off.
                pkg.Save();
                PackageDispose();


                // Now we have to use Excel automation in order to copy the data from the working temp sheet
                // to the final destination. This is because all the other methods were far too slow for large files.
                // Talking 10+ mins (didn't wait for it to finish) for files with 250k rows.
                // Tried EPPlus, OLEDB with indiovidual Updates and parameterized from a DataTable, etc.
                Application app = new Application();
                app.Visible = false;
                app.DisplayAlerts = false;
                Workbook book = app.Workbooks.Open(fileName);
                Worksheet sheetTo = book.Sheets[sheetName];
                Worksheet sheetFrom = book.Sheets[tempSheetName];

                Range firstRowSheetFrom = sheetFrom.get_Range("A1", MaxColumnID+"1");
                Range firstRowSheetTo = sheetTo.get_Range("A1", MaxColumnID+"1");
                int numRows = newValuesTable.Rows.Count;

                // Do this one column at a time...
                foreach (DataColumn col in newValuesTable.Columns)
                {

                    // Find the first cell below the header column in the from sheet - there will always be a header because
                    // we put it there when we wrote it out from the datatable.
                    Range fromCell=null, toCell=null;
                    foreach (Range cell in firstRowSheetFrom.Cells)
                        if ((cell.Value2 != null) && (cell.Value2.Equals(col.ColumnName)))
                        {
                            fromCell = cell.Offset[1, 0];
                            break;
                        }
                    // Find the first cell below the header. The header text will be there because we put it there,
                    // but if it wasn't supposed to be there (since the others don't), then overwrite it.
                    foreach (Range cell in firstRowSheetTo.Cells)
                        if ((cell.Value2 != null) && (cell.Value2.Equals(col.ColumnName)))
                        {
                            toCell = cell;
                            if (firstRowHasHeaders) // Then don't overwrite it - start one row down
                                toCell = cell.Offset[1, 0];
                            break;
                        }

                    // Do a simple copy/paste
                    sheetFrom.get_Range(fromCell, fromCell.get_Offset(numRows, 0)).Copy();
                    sheetTo.get_Range(toCell, toCell.get_Offset(numRows, 0)).PasteSpecial(XlPasteType.xlPasteValues);
                }


                // Clean up by deleting the temp working sheet
                book.Sheets[tempSheetName].Delete();

                // Save and exit
                book.Save();
                book.Close();
                app.Quit();

                rc = ReturnCode.Success;

            }

            return rc;
        }


        public ReturnCode GetTwitterUserInfo(string userID, ref TwitterUserInfo userInfo)
        {
            ReturnCode rc = ReturnCode.Failed;

            User user = new User(userID, Token);

            userInfo.Location = user.Location;
            userInfo.Name = user.Name;
            userInfo.NumberOfFollowers = user.FollowersCount.HasValue ? user.FollowersCount.Value : 0;
            userInfo.NumberFollowing = user.FriendsCount.HasValue ? user.FriendsCount.Value : 0;

            return rc;

        }
    }
}



#region oldcode
//public ReturnCode WriteColumnValues(string fileName, string sheetName, Dictionary<string,List<string>> newValues, int firstRowNumber)
//{
//    // Remember that firstRow is 0-based, but in EPPlus rows are 1-based
//    firstRowNumber++;

//    ReturnCode rc = ReturnCode.Failed;

//    if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
//    {
//        ExcelPackage pkg = Package(fileName);
//        ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

//        // For each KeyValuePair... Each KeyValuePair contains a string that is the column name and
//        // a list of strings that are the values.
//        // So we will write the new values one column at a time.
//        foreach (KeyValuePair<string, List<string>> columnSet in newValues)
//        {
//            string columnName = columnSet.Key;
//            List<string> valueList = columnSet.Value;

//            // Find the location of the column in the sheet
//            ExcelRange firstRow = sheet.Cells["A1:" + MaxColumnID + "1"];
//            ExcelRangeBase headerCell = firstRow.FirstOrDefault<ExcelRangeBase>(cell => cell.Text.Equals(columnName));
//            if (headerCell != null)
//            {
//                // We will fill in N cells in the column that had the column name in the top cell.
//                // N will be the number of values in the valueList.
//                // So this range will be one cell wide and N tall.
//                ExcelRangeBase firstCell = headerCell.Offset(1, 0);
//                ExcelRangeBase lastCell = firstCell.Offset(valueList.Count - 1, 0);

//                // Fill them in.
//                // This method seems to have a major performance hit when there are more than 10,000 values.
//                //sheet.Cells[firstCell.Address].LoadFromCollection(from s in valueList select s);

//                // So we need to chunk at 10,000
//                //int start = 0;
//                //int chunkSize = 10000;
//                //while (start < valueList.Count())
//                //{
//                //    List<string> chunkedValueList = new List<string>(valueList.GetRange(start, Math.Min(chunkSize,valueList.Count()-start)));
//                //    sheet.Cells[firstCell.Address].LoadFromCollection(from s in chunkedValueList select s);
//                //    firstCell = firstCell.Offset(chunkSize, 0);
//                //    start += chunkSize;
//                //}

//                // Still too slow after 10000. Lte's try writing them on at a time
//                foreach (string s in valueList)
//                {
//                    sheet.Cells[firstCell.Address].Value = s;
//                    firstCell = firstCell.Offset(1, 0);
//                }

//                rc = ReturnCode.Success;
//            }
//        }

//        pkg.Save();
//    }

//    return rc;
//}

#endregion
