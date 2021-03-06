﻿using System;
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
using Tweetinvi;
using System.Configuration;
using System.Text.RegularExpressions;
using System.Threading;
using Tweetinvi.Logic.TwitterEntities;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class EPPlusModel : ISMRTDataModel
    {
        string MaxColumnID = "XFD"; // With Excel 2010+ there are 16384 columns available
        //string MaxRowID = "1000000"; // With Excel 2010+ there are a little more than 1 million rows available

        ExcelPackage _pkg = null;
        Tweetinvi.Models.ITwitterCredentials _twitterCreds = null;

        // Need a place to cache the user info that was already obtained to we don't ask Twitter
        // for the same info repeatedly - Twitter does have a rate limit on the API use
        Dictionary<string, TwitterUserInfo> cachedUserInfo = new Dictionary<string, TwitterUserInfo>();

        public void Dispose()
        {
            PackageDispose();
        }

        private static readonly Encoding Utf8Encoder = Encoding.GetEncoding(
            "UTF-8",
            new EncoderReplacementFallback(string.Empty),
            new DecoderExceptionFallback()
        );

        private void ConvertSheetToUTF8(string fileName, string sheetName)
        {

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                // Now we have to use Excel automation in order to copy the data from the working temp sheet
                // to the final destination. This is because all the other methods were far too slow for large files.
                // Talking 10+ mins (didn't wait for it to finish) for files with 250k rows.
                // Tried EPPlus, OLEDB with indiovidual Updates and parameterized from a DataTable, etc.
                Application app = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
                Workbook book = app.Workbooks.Open(fileName);
                Worksheet sheetTo = book.Sheets[sheetName];

                Range usedCells = sheetTo.UsedRange;

                int count = 0;
                foreach (Range cell in usedCells)
                {
                    if (cell.Value2 is String)
                    {
                        //var utf8Text = Utf8Encoder.GetString(Utf8Encoder.GetBytes(cell.Value2));
                        //if (utf8Text != cell.Value2)
                        //    cell.Value2 = utf8Text;
                        string s = cell.Value2;
                        s = Regex.Replace(s, @"[^\u0000-\u007F]+", string.Empty);
                        if (s != cell.Value2)
                            cell.Value2 = s;
                    }
                    count++;
                }

                // Save and exit
                try
                {
                    book.Save();
                    book.Close();
                }
                finally
                {
                    app.Quit();
                }


            }
        }

        private void Save(ExcelPackage pkg, string fileName)
        {
            FileInfo fi = new FileInfo(fileName);
            String newFileName = String.Concat(fi.DirectoryName, "\\", Path.GetFileNameWithoutExtension(fi.Name), "-", DateTime.Now.Year, "-", DateTime.Now.Month.ToString("D2"), "-", DateTime.Now.Day.ToString("D2"),
                " ", DateTime.Now.Hour.ToString("D2"), "-", DateTime.Now.Minute.ToString("D2"), "-", DateTime.Now.Second.ToString("D2"), "-", DateTime.Now.Millisecond, fi.Extension);
            File.Copy(fileName, newFileName);

            pkg.Save();
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

        private Tweetinvi.Models.ITwitterCredentials SetupTwitterConnection()
        {
            if (_twitterCreds == null)
            {
                _twitterCreds = new Tweetinvi.Models.TwitterCredentials(
                    ConfigurationManager.AppSettings["token_ConsumerKey"],
                    ConfigurationManager.AppSettings["token_ConsumerSecret"],
                    ConfigurationManager.AppSettings["token_AccessToken"],
                    ConfigurationManager.AppSettings["token_AccessTokenSecret"]);
                Tweetinvi.Auth.SetCredentials(_twitterCreds);

                Tweetinvi.RateLimit.RateLimitTrackerMode = RateLimitTrackerMode.TrackAndAwait;

                //TweetinviEvents.QueryBeforeExecute += TweetinviEvents_QueryBeforeExecute;
            }

            return _twitterCreds;
        }

        private void TweetinviEvents_QueryBeforeExecute(object sender, Tweetinvi.Events.QueryBeforeExecuteEventArgs e)
        {
            var queryRateLimits = RateLimit.GetQueryRateLimit(e.QueryURL);

            if ((queryRateLimits != null) && (queryRateLimits.Remaining == 0))
            {
                Thread.Sleep((int)queryRateLimits.ResetDateTimeInMilliseconds + 1000);
                //e.Cancel = true;
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

        public System.Data.DataTable GetColumnValuesForColumnNames(string fileName, string sheetName, string[] newColumnNames, bool firstRowHasHeaders)
        {
            System.Data.DataTable table = new System.Data.DataTable();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && (newColumnNames.Count()>0))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                Dictionary<string, string> columnNamesInFile = GetColumnNames(fileName, sheetName);

                int rowCount = sheet.Dimension.End.Row;
                int colCount = sheet.Dimension.End.Column;

                foreach (string newColumnName in newColumnNames)
                    table.Columns.Add(newColumnName);

                for (int i = 0; i < rowCount; i++)
                    table.Rows.Add(table.NewRow());

                foreach (string newColumnName in newColumnNames)
                {
                    foreach (KeyValuePair<string, string> columnNameInFile in columnNamesInFile)
                    {
                        if (columnNameInFile.Value.Equals(newColumnName))
                        {
                            string columnID = columnNameInFile.Key;

                            ExcelRange cells = sheet.Cells[columnID + ":" + Regex.Match(columnID, @"[A-Z]+") + Convert.ToString(rowCount)];
                            foreach (ExcelRangeBase cell in cells)
                                table.Rows[cell.Start.Row-1][newColumnName] = cell.Text;
                            cells.Dispose();
                        }
                    }
                }
                if (firstRowHasHeaders && (table.Rows.Count > 0))
                    table.Rows[0].Delete();
            }

            return table;
        }

        public ReturnCode AddColumns(string fileName, string sheetName, string[] columnNames)
        {
            ReturnCode rc = ReturnCode.Failed;

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && (columnNames.Length >= 0))
            {


                // ConvertSheetToUTF8(fileName, sheetName);

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
                    try
                    {
                        Save(pkg, fileName);
                    }
                    finally
                    {
                        PackageDispose();
                    }
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
                ExcelPackage pkg = Package(fileName);
                if (pkg.File.IsReadOnly)
                    isValid = false;
                else
                    isValid = true;
            }

            return isValid;
        }

        
        public ReturnCode WriteColumnValuesToNewSheet(string fileName, string sheetName, System.Data.DataTable newValuesTable, bool firstRowHasHeaders)
        {
            string tempSheetName = "SMRT_" + sheetName;
            // Apparently EPPlus, or perhaps it's Excel, has a limit of 30 chars for a sheet name;
            tempSheetName = tempSheetName.Substring(0, Math.Min(tempSheetName.Length, 30));

            ReturnCode rc = ReturnCode.Failed;

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                ExcelPackage pkg = Package(fileName);
                ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];
                ExcelWorksheet tempWorkSheet;

                // Create a temp spreadsheet and make sure that we don't overwrite any existing ones.
                if (pkg.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals(tempSheetName)) == null)
                    tempWorkSheet = pkg.Workbook.Worksheets.Add(tempSheetName);
                else
                {
                    // Keep incrementing a counter to add to the end of a candidate sheet name until there
                    // isn't a sheet by that name.
                    int i = 1;
                    string suffix = "_" + i;
                    string nameToTry = tempSheetName.Substring(0, Math.Min(tempSheetName.Length, 30 - suffix.Length)) + suffix;
                    while (pkg.Workbook.Worksheets.FirstOrDefault(s => s.Name.Equals(nameToTry)) != null)
                    {
                        i++;
                        suffix = "_" + i;
                        nameToTry = tempSheetName.Substring(0, Math.Min(tempSheetName.Length, 30 - suffix.Length)) + suffix;
                    }
                    tempWorkSheet = pkg.Workbook.Worksheets.Add(nameToTry);
                }

                // Load from the datatable into the worksheet.
                tempWorkSheet.Cells["A1"].LoadFromDataTable(newValuesTable, true);

                // Save it off.
                try
                {
                    Save(pkg, fileName);
                }
                finally
                {
                    PackageDispose();
                }
            }

            return rc;
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
                try
                {
                    Save(pkg, fileName);
                }
                finally
                {
                    PackageDispose();
                }


                // Now we have to use Excel automation in order to copy the data from the working temp sheet
                // to the final destination. This is because all the other methods were far too slow for large files.
                // Talking 10+ mins (didn't wait for it to finish) for files with 250k rows.
                // Tried EPPlus, OLEDB with indiovidual Updates and parameterized from a DataTable, etc.
                Application app = new Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };
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
                try
                {
                    book.Save();
                    book.Close();
                }
                finally
                {
                    app.Quit();
                }

                rc = ReturnCode.Success;

            }

            return rc;
        }


        public ReturnCode GetTwitterUserInfo(string userID, ref TwitterUserInfo userInfo)
        {
            ReturnCode rc = ReturnCode.Failed;

            SetupTwitterConnection();

            if (cachedUserInfo.ContainsKey(userID))
            {
                userInfo = cachedUserInfo[userID];
                rc = ReturnCode.Success;
            }
            //else if (Token.XRateLimitRemaining <= 170)
            //{
            //    userInfo.Name = "<error - rate limit>";
            //    userInfo.NumberOfFollowers = 0;
            //    userInfo.NumberFollowing = 0;
            //    userInfo.Description = "";
            //}
            else
            {
                try
                {
                    var user = User.GetUserFromScreenName(userID);

                    userInfo.Screenname = user.ScreenName;
                    userInfo.Location = user.Location;
                    userInfo.Name = user.Name;
                    userInfo.NumberOfFollowers = user.FollowersCount;
                    userInfo.NumberFollowing = user.FriendsCount;
                    userInfo.Description = user.Description;
                    rc = ReturnCode.Success;
                }
                catch (System.Net.WebException e)
                {
                    userInfo.Screenname = userID;
                    userInfo.Location = "";
                    userInfo.NumberOfFollowers = 0;
                    userInfo.NumberFollowing = 0;
                    userInfo.Description = "";
                    if (e.Message.Contains("403"))
                        userInfo.Name = "<error - forbidden>";
                    else if (e.Message.Contains("429"))
                        userInfo.Name = "<error - rate limit exceeded or other client error>";
                    else if (e.Message.Contains("404"))
                        userInfo.Name = "<error - not found>";
                    else
                        userInfo.Name = "<error - " + e.Message + ">";
                }
                catch (Exception)
                {
                    userInfo.Screenname = userID;
                    userInfo.Location = "";
                    userInfo.Name = "<error - unknown>";
                    userInfo.NumberOfFollowers = 0;
                    userInfo.NumberFollowing = 0;
                    userInfo.Description = "";
                }

                // We should cache this info just in case it is asked for again
                cachedUserInfo.Add(userID, userInfo);
            }


            return rc;

        }

        public ReturnCode GetTwitterConnections(string userName, TwitterConnectionType connectionType, int maxConnections, ref TwitterUserInfo sourceUserInfo, ref List<TwitterUserInfo> connections)
        {
            ReturnCode rc = ReturnCode.Success;

            SetupTwitterConnection();

            RateLimitTrackerMode priorTrackerMode = RateLimit.RateLimitTrackerMode;
            RateLimit.RateLimitTrackerMode = RateLimitTrackerMode.TrackAndAwait;
            var rateLimits = RateLimit.GetCurrentCredentialsRateLimits();

            sourceUserInfo = new TwitterUserInfo();
            rc = GetTwitterUserInfo(userName, ref sourceUserInfo);

            var connectionEnum = connectionType==TwitterConnectionType.Followers ? 
                Tweetinvi.User.GetFollowers(userName, Math.Min(sourceUserInfo.NumberOfFollowers, maxConnections)) : 
                Tweetinvi.User.GetFriends(userName, Math.Min(sourceUserInfo.NumberFollowing, maxConnections));

            foreach (var user in connectionEnum)
            {
                TwitterUserInfo userInfo = new TwitterUserInfo
                {
                    Screenname = user.ScreenName,
                    Location = user.Location,
                    Name = user.Name,
                    NumberOfFollowers = user.FollowersCount,
                    NumberFollowing = user.FriendsCount,
                    Description = user.Description
                };
                connections.Add(userInfo);
            }

            var rateLimitsAfter = RateLimit.GetCurrentCredentialsRateLimits();
            RateLimit.RateLimitTrackerMode = priorTrackerMode;

            return rc;
        }

        public ReturnCode GetTwitterFriendship(string userName1, string userName2, ref bool followedBy, ref bool following)
        {
            ReturnCode rc = ReturnCode.Success;

            SetupTwitterConnection();

            var relationshipDetails = (RelationshipDetails)Tweetinvi.Friendship.GetRelationshipDetailsBetween(userName1, userName2);

            if (relationshipDetails != null)
            {
                followedBy = relationshipDetails.FollowedBy;
                following = relationshipDetails.Following;
            }

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
