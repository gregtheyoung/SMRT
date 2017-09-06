using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;
using System.Text.RegularExpressions;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class DomainModel : ISMRTDomain, IDisposable
    {
        string[] newURLSplitColumns = { "MentionType", "Domain", "PosterID", "MentionID", "NumPosts" };
        enum NewSplitColumnIndex {MentionType, Domain, PosterID, MentionID, NumPosts};

        string[] newTwitterColumns = { "TwitterName", "TwitterLocation", "NumFollowers", "NumFollowing", "Description" };
        enum NewTwitterColumnIndex {Name, Location, NumFollowers, NumFollowing, Description};

        protected ISMRTDataModel dataModel;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataModelToUse">Which data model to use: 0=ExcelAutomation, 1=OLEDB Jet, 2=OLEDB ACE, 3=EPPlus</param>
        public DomainModel(int dataModelToUse)
        {
            switch (dataModelToUse)
            {
                //case (0):
                //{
                //    dataModel = new ExcelAutomationModel();
                //    break;
                //}
                //case (1):
                //{
                //    dataModel = new OLEDBModel(true);
                //    break;
                //}
                //case (2):
                //{
                //    dataModel = new OLEDBModel(false);
                //    break;
                //}
                case (3):
                {
                    dataModel = new EPPlusModel();
                    break;
                }
            }
        }

        public void Dispose()
        {
            dataModel.Dispose();
        }

        /// <summary>
        /// Gets the names of worksheets in an Excel file.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook file</param>
        /// <returns>A list of the names of the columns in the workbook</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        public List<string> GetSheetNames(string filename)
        {
            List<string> sheetNames = null;

            // Check that the file is valid, throw exception if not
            if (dataModel.FileIsValid(filename))
                sheetNames = dataModel.GetSheetNames(filename);
            else
                throw new System.IO.FileNotFoundException();

            return sheetNames;
        }

        /// <summary>
        /// Gets the column identifiers and names from a worksheet in an Excel file
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>A list of the identifiers and names of the columns in the worksheet.</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = null;

            // Check that the file is valid, throw exception if not
            if (dataModel.FileIsValid(fileName))
                columnNames = dataModel.GetColumnNames(fileName, sheetName);
            else
                throw new System.IO.FileNotFoundException();

            return columnNames;
        }

        /// <summary>
        /// Checks that the columns to contain the split values already exist in the sheet or not.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>True if any of the columns already exists in the sheet.</returns>
        private bool DoURLSplitColumnsAlreadyExist(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);

            bool newColumnAlreadyExists = false;

            // For each column to contain the split values...
            foreach (string newColumnName in newURLSplitColumns)
            {
                // If that column already exists in the sheet...
                if (columnNames.ContainsValue(newColumnName))
                    newColumnAlreadyExists = true;
            }

            return newColumnAlreadyExists;
        }

        /// <summary>
        /// When the contents of a cell in the column specified is a URI of a valid mention/post, three new columns will be added to the
        /// sheet and populated with the domain from the URI, the ID of the poster, and the ID of the mention.
        /// 
        /// This is only valid when the URI is determined to be a mention/post on Facebook, Twitter, Tumblr, or Blogspot.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="urlColumnID">The column identifier of the column to be parsed</param>
        /// <param name="overwriteExistingData"></param>
        /// <returns>A return code indicating success of failure</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        public ReturnCode SplitURLs(string fileName, string sheetName, string urlColumnName, bool overwriteExistingData, bool firstRowHasHeaders)
        {
            ReturnCode rc = ReturnCode.Success;

            // We are using this because of an issue with saving data to an existing sheet that already has data in it.
            // It would fail when writing to Conger's server.
            bool writeToNewSheet = true;

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileName))
                throw new System.IO.FileNotFoundException();

            // If writing to a new sheet or the existing split out columns are to be overwritten, or if they don't exist yet...
            if (writeToNewSheet || overwriteExistingData || !DoURLSplitColumnsAlreadyExist(fileName, sheetName))
            {
                // Add the new columns to the sheet...
                if (!writeToNewSheet)
                    rc = dataModel.AddColumns(fileName, sheetName, newURLSplitColumns);

                // If writing to a new sheet or the columns were added successfully
                if (writeToNewSheet || (rc == ReturnCode.Success))
                {
                    // Get the contents of the column. This will provide the row/cell identifier (depends on the
                    // underlying data model implementation) and the value of column for that row/cell.
                    DataTable urlStringTable = dataModel.GetColumnValuesForColumnNames(fileName, sheetName, new string[] {urlColumnName}, firstRowHasHeaders);

                    // Keep track of processing and failure counts for a short-circuit abort if needed
                    int numprocessed = 0;
                    int numFailed = 0;

                    // Keep a list of all the split out info. Ensure that there is exactly one for each value in the column.
                    DataTable newValuesTable = new DataTable();
                    foreach (string splitColumnName in newURLSplitColumns)
                        if (splitColumnName.Equals(newURLSplitColumns[(int)NewSplitColumnIndex.NumPosts]))
                            newValuesTable.Columns.Add(splitColumnName, System.Type.GetType("System.Int16"));
                        else
                            newValuesTable.Columns.Add(splitColumnName);

                    // For each cell/value...
                    foreach (DataRow row in urlStringTable.Rows)
                    {
                        DataRow newRow = newValuesTable.NewRow();
                        try
                        {
                            // Try to parse it as a URI
                            Uri uri = new Uri((string)row[0]);

                            // Get the domain part.
                            string domain = uri.GetComponents(UriComponents.Host, UriFormat.Unescaped);

                            // Get the segments of the URI - these are the slash-delimited pieces after the domain.
                            string[] segments = uri.Segments;
                            NameValueCollection queryCollection = HttpUtility.ParseQueryString(uri.Query);

                            if (domain.EndsWith("twitter.com"))
                            {
                                GetTwitterParts(uri, domain, segments, queryCollection, ref newRow);
                            }
                            else if (domain.EndsWith("facebook.com"))
                            {
                                GetFacebookParts(uri, domain, segments, queryCollection, ref newRow);
                            }
                            else if (domain.EndsWith("blogspot.com"))
                            {
                                GetBloggerParts(uri, domain, segments, queryCollection, ref newRow);
                            }
                            else if (domain.EndsWith("tumblr.com"))
                            {
                                GetTumblrParts(uri, domain, segments, queryCollection, ref newRow);
                            }
                            else
                            {
                                GetUnknownParts(uri, domain, segments, queryCollection, ref newRow);
                            }
                        }

                        catch (UriFormatException e)
                        {
                            newRow[(int)NewSplitColumnIndex.Domain] = null;
                            newRow[(int)NewSplitColumnIndex.MentionID] = null;
                            newRow[(int)NewSplitColumnIndex.MentionType] = "Failed";
                            newRow[(int)NewSplitColumnIndex.NumPosts] = DBNull.Value;
                            newRow[(int)NewSplitColumnIndex.PosterID] = null;
                            numFailed++;
                        }
                        numprocessed++;

                        // If more than 1/2 have failed, then this likely is not a column that contains URLs.
                        // Only check this condition after we have processed some of them - don't want to error just because 
                        // one of the first two failed.
                        if ((numprocessed >= 50) && (numFailed > numprocessed / 2))
                        {
                            rc = ReturnCode.NotURLColumn;
                            break;
                        }

                        newValuesTable.Rows.Add(newRow);

                    }

                    if (rc == ReturnCode.Success)
                    {
                        AddNumPostCounts(newValuesTable);
                        if (writeToNewSheet)
                            dataModel.WriteColumnValuesToNewSheet(fileName, sheetName, newValuesTable, firstRowHasHeaders);
                        else
                            dataModel.WriteColumnValues(fileName, sheetName, newValuesTable, firstRowHasHeaders);
                    }
                }
            }
            else
            {
                rc = ReturnCode.ColumnsAlreadyExist;
            }

            return rc;
        }

        private void AddNumPostCounts(DataTable newValuesTable)
        {
            Dictionary<string, int> postCounts = new Dictionary<string, int>(); ;

            foreach(DataRow row in newValuesTable.Rows)
            {
                string posterID = Convert.ToString(row[(int)NewSplitColumnIndex.PosterID]);
                if ((posterID != null) && !String.IsNullOrEmpty(posterID))
                {
                    if (postCounts.ContainsKey(posterID))
                        postCounts[posterID] = postCounts[posterID] + 1;
                    else
                        postCounts.Add(posterID, 1);
                }
            }

            foreach (DataRow row in newValuesTable.Rows)
            {
                string posterID = Convert.ToString(row[(int)NewSplitColumnIndex.PosterID]);
                if ((posterID != null) && !String.IsNullOrEmpty(posterID))
                    row[(int)NewSplitColumnIndex.NumPosts] = postCounts[(string)row[(int)NewSplitColumnIndex.PosterID]];
            }
            

        }

        private void GetFacebookParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection, ref DataRow newRow)
        {
            // Patterns of URLs
            // http://facebook.com/permalink.php?story_fbid=100127316863475&id=100005986200023
            // http://facebook.com/events/122767261239362/permalink/122767264572695
            // http://facebook.com/media/set/?set=a.10151209676509212.454268.38951299211&type=1
            // http://facebook.com/notes/complex-child-e-magazine/childrens-mental-health-edition/10151458448874231
            // http://facebook.com/105indaklubb/posts/10151141293211990

            try
            {
                if (segments[1].Equals("permalink.php"))
                {
                    newRow[(int)NewSplitColumnIndex.MentionType] = "FacebookPost";
                    newRow[(int)NewSplitColumnIndex.Domain] = domain;
                    newRow[(int)NewSplitColumnIndex.PosterID] = queryCollection["id"];
                    newRow[(int)NewSplitColumnIndex.MentionID] = queryCollection["story_fbid"];
                }
                else if (segments[1].Equals("events/"))
                {
                    newRow[(int)NewSplitColumnIndex.MentionType] = "FacebookEvent";
                    newRow[(int)NewSplitColumnIndex.Domain] = domain;
                    newRow[(int)NewSplitColumnIndex.PosterID] = segments[4].Trim('/');
                    newRow[(int)NewSplitColumnIndex.MentionID] = segments[2].Trim('/');
                }
                else if (segments[1].Equals("media/"))
                {
                    newRow[(int)NewSplitColumnIndex.MentionType] = "FacebookMedia";
                    newRow[(int)NewSplitColumnIndex.Domain] = domain;
                    newRow[(int)NewSplitColumnIndex.PosterID] = "";
                    newRow[(int)NewSplitColumnIndex.MentionID] = queryCollection["set"];
                }
                else if (segments[1].Equals("notes/"))
                {
                    newRow[(int)NewSplitColumnIndex.MentionType] = "FacebookNote";
                    newRow[(int)NewSplitColumnIndex.Domain] = domain;
                    newRow[(int)NewSplitColumnIndex.PosterID] = segments[2].Trim('/');
                    newRow[(int)NewSplitColumnIndex.MentionID] = segments[4].Trim('/');
                }
                else
                {
                    newRow[(int)NewSplitColumnIndex.MentionType] = "FacebookGroupPost";
                    newRow[(int)NewSplitColumnIndex.Domain] = domain;
                    newRow[(int)NewSplitColumnIndex.PosterID] = segments[1].Trim('/');
                    newRow[(int)NewSplitColumnIndex.MentionID] = segments[3].Trim('/');
                }
            }
            catch (Exception e)
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Facebook";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = null;
                newRow[(int)NewSplitColumnIndex.MentionID] = null;
            }
        }

        private void GetTwitterParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection, ref DataRow newRow)
        {
            try
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Twitter";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = segments[1].Trim('/');
                newRow[(int)NewSplitColumnIndex.MentionID] = segments[3].Trim('/');
            }
            catch (Exception e)
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Twitter";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = null;
                newRow[(int)NewSplitColumnIndex.MentionID] = null;
            }
        }

        private void GetBloggerParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection, ref DataRow newRow)
        {
            try
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Blogger";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;

                // Call blogs/byurl with the full URL
                // Then call posts/bypath using the user ID from the first one and the part of the URL after the domain.
                newRow[(int)NewSplitColumnIndex.PosterID] = uri.Host.Replace(".blogspot.com", "");
                newRow[(int)NewSplitColumnIndex.MentionID] = uri.PathAndQuery;
            }
            catch (Exception e)
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Blogger";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = null;
                newRow[(int)NewSplitColumnIndex.MentionID] = null;
            }
        }

        private void GetTumblrParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection, ref DataRow newRow)
        {
            try
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Tumblr";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;

                // Call blogs/byurl with the full URL
                // Then call posts/bypath using the user ID from the first one and the part of the URL after the domain.
                newRow[(int)NewSplitColumnIndex.PosterID] = uri.Host.Replace(".tumblr.com", "");
                newRow[(int)NewSplitColumnIndex.MentionID] = segments[2].Trim('/');
            }
            catch (Exception e)
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Tumblr";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = null;
                newRow[(int)NewSplitColumnIndex.MentionID] = null;
            }
        }

        private void GetUnknownParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection, ref DataRow newRow)
        {
            try
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Unknown";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;

                // Call blogs/byurl with the full URL
                // Then call posts/bypath using the user ID from the first one and the part of the URL after the domain.
                newRow[(int)NewSplitColumnIndex.PosterID] = domain;
                newRow[(int)NewSplitColumnIndex.MentionID] = uri.PathAndQuery;
            }
            catch (Exception e)
            {
                newRow[(int)NewSplitColumnIndex.MentionType] = "Unknown";
                newRow[(int)NewSplitColumnIndex.Domain] = domain;
                newRow[(int)NewSplitColumnIndex.PosterID] = null;
                newRow[(int)NewSplitColumnIndex.MentionID] = null;
            }
        }

        public ReturnCode AddTwitterInfo(string fileName, string sheetName, bool overwriteExistingData, bool firstRowHasHeaders, int numUsersToRetrieve)
        {
            ReturnCode rc = ReturnCode.Success;

            // We are using this because of an issue with saving data to an existing sheet that already has data in it.
            // It would fail when writing to Conger's server.
            bool writeToNewSheet = true;

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileName))
                throw new System.IO.FileNotFoundException();

            // Need to make sure that the SplitURL columns do exist
            if (!DoURLSplitColumnsAlreadyExist(fileName, sheetName))
                rc = ReturnCode.ColumnsMissing;
            else
            {
                // If the existing split out columns are to be overwritten, or if they don't exist yet...
                if (writeToNewSheet || overwriteExistingData || !DoTwitterColumnsAlreadyExist(fileName, sheetName))
                {
                    // Add the new columns to the sheet...
                    if (!writeToNewSheet)
                        rc = dataModel.AddColumns(fileName, sheetName, newTwitterColumns);

                    // If writing to a new sheet or the columns were added successfully...
                    if (writeToNewSheet || (rc == ReturnCode.Success))
                    {
                        // Get the contents of the column. This will provide the row/cell identifier (depends on the
                        // underlying data model implementation) and the value of column for that row/cell.
                        DataTable posterInfoTable = dataModel.GetColumnValuesForColumnNames(fileName, sheetName,
                            newURLSplitColumns,
                            firstRowHasHeaders);

                        List<string> topTwitterPosters = GetTopTwitterPosters(posterInfoTable);

                        // We will only use the top N
                        if (topTwitterPosters.Count > numUsersToRetrieve)
                            topTwitterPosters.RemoveRange(numUsersToRetrieve, topTwitterPosters.Count - numUsersToRetrieve);

                        // Keep a table of all the twitter info. Ensure that there is exactly one for each value in the column.
                        DataTable newValuesTable = new DataTable();
                        foreach (string twitterColumnName in newTwitterColumns)
                            if (twitterColumnName.Equals(newTwitterColumns[(int)NewTwitterColumnIndex.NumFollowers]) ||
                                twitterColumnName.Equals(newTwitterColumns[(int)NewTwitterColumnIndex.NumFollowing]))
                                newValuesTable.Columns.Add(twitterColumnName, System.Type.GetType("System.Int32"));
                            else
                                newValuesTable.Columns.Add(twitterColumnName);

                        // For each cell/value...
                        foreach (DataRow row in posterInfoTable.Rows)
                        {
                            DataRow newRow = newValuesTable.NewRow();

                            string mentionType = (string)row[newURLSplitColumns[(int)NewSplitColumnIndex.MentionType]];
                            string posterID = (string)row[newURLSplitColumns[(int)NewSplitColumnIndex.PosterID]];
                            if (mentionType.Equals("Twitter") && topTwitterPosters.Contains(posterID))
                            {
                                try
                                {
                                    GetTwitterInfo(posterID, ref newRow);
                                }
                                catch
                                {
                                    newRow[(int)NewTwitterColumnIndex.Location] = "";
                                    newRow[(int)NewTwitterColumnIndex.Name] = "";
                                    newRow[(int)NewTwitterColumnIndex.NumFollowers] = DBNull.Value;
                                    newRow[(int)NewTwitterColumnIndex.NumFollowing] = DBNull.Value;
                                }
                            }
                            else
                            {
                                newRow[(int)NewTwitterColumnIndex.Location] = "";
                                newRow[(int)NewTwitterColumnIndex.Name] = "";
                                newRow[(int)NewTwitterColumnIndex.NumFollowers] = DBNull.Value;
                                newRow[(int)NewTwitterColumnIndex.NumFollowing] = DBNull.Value;
                            }
                            newValuesTable.Rows.Add(newRow);
                        }

                        if (writeToNewSheet)
                            dataModel.WriteColumnValuesToNewSheet(fileName, sheetName, newValuesTable, firstRowHasHeaders);
                        else
                            dataModel.WriteColumnValues(fileName, sheetName, newValuesTable, firstRowHasHeaders);
                    }
                }
                else
                    rc = ReturnCode.ColumnsAlreadyExist;
            }

            return rc;
        }

        private List<string> GetTopTwitterPosters(DataTable posterInfoTable)
        {
            // Put all posters into a dictionary along with their count.
            Dictionary<string, int> users = new Dictionary<string, int>();
            foreach (DataRow row in posterInfoTable.Rows)
                // Only do it if there is actually a PosterID and NumPosts
                if (row[(int)NewSplitColumnIndex.MentionType].Equals("Twitter") &&
                    !String.IsNullOrEmpty((string)row[(int)NewSplitColumnIndex.NumPosts]))
                        if (!users.ContainsKey((string)row[(int)NewSplitColumnIndex.PosterID]))
                        users.Add((string)row[(int)NewSplitColumnIndex.PosterID], Convert.ToInt32(row[(int)NewSplitColumnIndex.NumPosts]));

            // Sort descending by count and put into a list.
            List<string> usersInOrder = new List<string>();
            foreach (KeyValuePair<string, int> pair in users.OrderByDescending(value => value.Value))
                usersInOrder.Add(pair.Key);

            return usersInOrder;
        }

        private void GetTwitterInfo(string posterID, ref DataRow newRow)
        {
            TwitterUserInfo userInfo = new TwitterUserInfo();

            ReturnCode rc = dataModel.GetTwitterUserInfo(posterID, ref userInfo);

            newRow[(int)NewTwitterColumnIndex.Location] = userInfo.Location;
            newRow[(int)NewTwitterColumnIndex.Name] = userInfo.Name;
            newRow[(int)NewTwitterColumnIndex.NumFollowers] = userInfo.NumberOfFollowers;
            newRow[(int)NewTwitterColumnIndex.NumFollowing] = userInfo.NumberFollowing;
            newRow[(int)NewTwitterColumnIndex.Description] = userInfo.Description;

        }

        /// <summary>
        /// Checks that the columns to contain the twitter values already exist in the sheet or not.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>True if any of the columns already exists in the sheet.</returns>
        private bool DoTwitterColumnsAlreadyExist(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);

            bool newColumnAlreadyExists = false;

            // For each column to contain the split values...
            foreach (string newColumnName in newTwitterColumns)
            {
                // If that column already exists in the sheet...
                if (columnNames.ContainsValue(newColumnName))
                    newColumnAlreadyExists = true;
            }

            return newColumnAlreadyExists;
        }

        public ReturnCode Autocode(string fileNameMentionFile, string sheetName, string firstColumnName, string secondColumnName, string fileNameAutocodeFile, bool overwriteExistingData, bool firstRowHasHeaders)
        {
            bool useRegex = true;

            ReturnCode rc = ReturnCode.Success;

            // We are going to write the same sheet as that which contains the mention text.
            // There is risk in this because of an issue with saving data to an existing sheet that already has data in it.
            // It would fail when writing to Conger's server.
            bool writeToNewSheet = false;

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileNameMentionFile))
                throw new System.IO.FileNotFoundException();

            // Find the code family names from the code family workbook. The names will be the name of the sheets in the code family workbook.
            List<string> codeFamilyNames = dataModel.GetSheetNames(fileNameAutocodeFile);
            List<string> newAutocodeColumns = new List<string>();
            foreach (string codeFamilyName in codeFamilyNames)
            {
                newAutocodeColumns.Add(codeFamilyName + ": Selected");
                newAutocodeColumns.Add(codeFamilyName + ": Total");
                newAutocodeColumns.Add(codeFamilyName + ": Terms");
            }

            // If writing to a new sheet or the existing split out columns are to be overwritten, or if they don't exist yet...
            if (writeToNewSheet || overwriteExistingData || !DoCodeFamilyColumnsAlreadyExist(fileNameMentionFile, newAutocodeColumns, sheetName))
            {

                // Add the new columns to the sheet...
                if (!writeToNewSheet)
                    rc = dataModel.AddColumns(fileNameMentionFile, sheetName, newAutocodeColumns.ToArray());

                // If writing to a new sheet or the columns were added successfully
                if (writeToNewSheet || (rc == ReturnCode.Success))
                {
                    // Get the contents of the column. This will provide the row/cell identifier (depends on the
                    // underlying data model implementation) and the value of column for that row/cell.
                    DataTable mentionStringTable;
                    if (String.IsNullOrEmpty(secondColumnName))
                        mentionStringTable = dataModel.GetColumnValuesForColumnNames(fileNameMentionFile, sheetName, new string[] { firstColumnName }, firstRowHasHeaders);
                    else
                        mentionStringTable = dataModel.GetColumnValuesForColumnNames(fileNameMentionFile, sheetName, new string[] { firstColumnName, secondColumnName }, firstRowHasHeaders);

                    // Get the code family terms.
                    Dictionary<string, List<string>> codeFamilyTermSets = new Dictionary<string, List<string>>();
                    foreach (string codeFamilyName in codeFamilyNames)
                    {
                        List<string> termSets = new List<string>();

                        // Get the contents of the Terms columns. The column must be named "Terms".
                        DataTable termsTable = dataModel.GetColumnValuesForColumnNames(fileNameAutocodeFile, codeFamilyName, new string[] { "Terms" }, true);

                        foreach (DataRow row in termsTable.Rows)
                        {
                            string termSet = row["Terms"].ToString();
                            if ((termSet != null) && (termSet.Length > 0))
                                termSets.Add(termSet.Trim());
                        }

                        codeFamilyTermSets[codeFamilyName] = termSets;
                    }

                    // Create a data table to hold the new columns.
                    DataTable newValuesTable = new DataTable();
                    foreach (string codeFamilyName in codeFamilyNames)
                    {
                        newValuesTable.Columns.Add(codeFamilyName + ": Selected");
                        newValuesTable.Columns.Add(codeFamilyName + ": Total", System.Type.GetType("System.Int16"));
                        newValuesTable.Columns.Add(codeFamilyName + ": Terms");
                    }

                    // For each row in the mention worksheet...
                    foreach (DataRow row in mentionStringTable.Rows)
                    {
                        // Create a row for the new values - need to match these rows one-to-one with the existing rows.
                        DataRow newRow = newValuesTable.NewRow();
                        string mentionText = row[firstColumnName].ToString();
                        // If the user specified a second column to use, then concat it to the first
                        if (!String.IsNullOrEmpty(secondColumnName))
                            mentionText += " " + row[secondColumnName].ToString();
                        mentionText.Replace("  ", " "); // Just some minor cleanup.

                        // For each code family...
                        foreach (string codeFamilyName in codeFamilyTermSets.Keys)
                        {
                            // This will hold an entry for each term found and its count.
                            Dictionary<string, int> termCounts = new Dictionary<string, int>();

                            int total = 0;

                            // For each term set for the code family...
                            foreach (string termSetString in codeFamilyTermSets[codeFamilyName])
                            {
                                int termCount = 0;
                                string[] splitTerms = termSetString.Split(';');
                                string firstTerm = splitTerms[0];

                                // For each term in the term set...
                                foreach (string term in splitTerms)
                                {
                                    if (useRegex)
                                    {
                                        string rgxTerm = term;
                                        if (!term.StartsWith("*"))  // If it doesn't start with a wildcard, force it to be word-boundary
                                            if ("#".Contains(term[0])) // Special case for hashtag since it is considered a word-boundary, but we want it to be considered part of the tag
                                                rgxTerm = @"\B" + rgxTerm;
                                            else
                                                rgxTerm = @"\b" + rgxTerm;
                                        if (!term.EndsWith("*"))    // If it doesn't end with a wildcard, force it to be a word-boundary
                                            rgxTerm += @"\b";
                                        rgxTerm = rgxTerm.Replace("*", ""); // Get rid of wildcards that aren't at the beginning or end
                                        Regex rgx = new Regex(rgxTerm, RegexOptions.IgnoreCase);
                                        termCount += rgx.Matches(mentionText).Count;
                                    }
                                    // This else code isn't used - just saved in case we need it later
                                    else
                                    {
                                        string[] termArray = new string[] { term.ToLower() };
                                        termCount += mentionText.ToLower().Split(termArray, StringSplitOptions.None).Length - 1;
                                    }
                                }
                                
                                if (termCount > 0)
                                    termCounts[firstTerm] = termCount;
                             
                                total += termCount;
                            }
                            
                            // Sort the dictionary of terms and their counts by the count.
                            var sortedCounts = from entry in termCounts orderby entry.Value descending select entry;
                            
                            // Build the string of terms found and their counts.
                            string termString = "";
                            foreach (KeyValuePair<string, int> pair in sortedCounts)
                            {
                                if (termString.Length > 0)
                                    termString += "; ";
                                termString += pair.Key + ":" + pair.Value.ToString();
                            }

                            newRow[codeFamilyName + ": Selected"] = "";
                            newRow[codeFamilyName + ": Total"] = total;
                            newRow[codeFamilyName + ": Terms"] = termString;
                        }
                        newValuesTable.Rows.Add(newRow);
                    }

                    if (rc == ReturnCode.Success)
                    {
                        if (writeToNewSheet)
                            dataModel.WriteColumnValuesToNewSheet(fileNameMentionFile, sheetName, newValuesTable, firstRowHasHeaders);
                        else
                            dataModel.WriteColumnValues(fileNameMentionFile, sheetName, newValuesTable, firstRowHasHeaders);
                    }
                }
            }
            else
            {
                rc = ReturnCode.ColumnsAlreadyExist;
            }

            return rc;
        }

        private bool DoCodeFamilyColumnsAlreadyExist(string fileName, List<string> codeFamilyNames, string sheetName)
        {
            Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);

            bool newColumnAlreadyExists = false;

            // For each column to contain the split values...
            foreach (string newColumnName in codeFamilyNames)
            {
                // If that column already exists in the sheet...
                if (columnNames.ContainsValue(newColumnName))
                    newColumnAlreadyExists = true;
            }

            return newColumnAlreadyExists;
        }

        public ReturnCode RandomSelect(string fileNameMentionFile, string sheetName, string columnNameAutocodeCounts, string columnNameRandomSelect, int percent, int floor, int ceiling, bool overwriteExistingData, bool firstRowHasHeaders)
        {
            ReturnCode rc = ReturnCode.Success;

            // We are going to write the same sheet as that which contains the mention text.
            // There is risk in this because of an issue with saving data to an existing sheet that already has data in it.
            // It would fail when writing to Conger's server.
            bool writeToNewSheet = false;

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileNameMentionFile))
                throw new System.IO.FileNotFoundException();

            // If writing to a new sheet or the existing split out columns are to be overwritten, or if they don't exist yet...
            if (writeToNewSheet || overwriteExistingData || !DoRandomSelectColumnsAlreadyExist(fileNameMentionFile, new string[] { columnNameAutocodeCounts, columnNameRandomSelect }, sheetName))
            {

                // Add the new columns to the sheet...
                if (!writeToNewSheet)
                    rc = dataModel.AddColumns(fileNameMentionFile, sheetName, new string[] { columnNameAutocodeCounts, columnNameRandomSelect });

                // If writing to a new sheet or the columns were added successfully
                if (writeToNewSheet || (rc == ReturnCode.Success))
                {
                    // Get the contents of the column. This will provide the row/cell identifier (depends on the
                    // underlying data model implementation) and the value of column for that row/cell.
                    DataTable countsAndSelectionsTable = dataModel.GetColumnValuesForColumnNames(fileNameMentionFile, sheetName, new string[] { columnNameAutocodeCounts, columnNameRandomSelect }, firstRowHasHeaders);

                    // This will contain the row number and the selection "Y" value for all rows that are part of the 
                    // candidate set of rows. The selection will contain what the file had, and then will be changed as needed
                    // by the random sampling below.
                    List<Tuple<int, string>> candidateRows = new List<Tuple<int, string>>();
                    
                    int rowNum = 0;
                    int selectedCount = 0;
                    
                    // For each row in the worksheet...
                    foreach (DataRow row in countsAndSelectionsTable.Rows)
                    {
                        // If the count is > 0 or if the selection is already a "Y", then add it as a candidate row.
                        if ((Convert.ToInt32(row[columnNameAutocodeCounts]) > 0) || row[columnNameRandomSelect].ToString().ToUpper().Equals("Y"))
                            candidateRows.Add(Tuple.Create(rowNum, row[columnNameRandomSelect].ToString().ToUpper()));
                        // If the selection is already a "Y", up our already selected count.
                        if (row[columnNameRandomSelect].ToString().ToUpper().Equals("Y"))
                            selectedCount++;
                        rowNum++;
                    }

                    // How many should we select? The target percent, but no less than the floor and no more than the ceiling.
                    int targetCount = Math.Min(ceiling, Math.Max(floor, Convert.ToInt32(Math.Round(candidateRows.Count * percent / 100.0))));

                    // Only if the current count is already less than the target count do we continue, since otherwise 
                    // we already have enough selected and we never reduce the count.
                    if ((selectedCount < targetCount) && (candidateRows.Count > 0))
                    {

                        Random randomGenerator = new Random();

                        // While we have not selected enough, generate another random number.
                        // If that candidate row has not yet been selected, selected it and add to the selected count.
                        while (selectedCount < targetCount)
                        {
                            int r = randomGenerator.Next(candidateRows.Count);
                            if (!candidateRows[r].Item2.Equals("Y"))
                            {
                                candidateRows[r] = Tuple.Create(candidateRows[r].Item1, "Y");
                                selectedCount++;
                            }
                        }

                        // Keep a list of all the new columns info. Ensure that there is exactly one for each value in the column.
                        DataTable newValuesTable = new DataTable();
                        newValuesTable.Columns.Add(columnNameRandomSelect);

                        // Create the new data table but without any selections. We'll populate those next.
                        foreach (DataRow row in countsAndSelectionsTable.Rows)
                        {
                            DataRow newRow = newValuesTable.NewRow();
                            newRow[columnNameRandomSelect] = "";
                            newValuesTable.Rows.Add(newRow);
                        }

                        // For each candidate row that was selected, mark that in the new columns.
                        foreach (Tuple<int, string> t in candidateRows)
                        {
                            if (t.Item2.Equals("Y"))
                                newValuesTable.Rows[t.Item1][columnNameRandomSelect] = "Y";
                        }

                        if (rc == ReturnCode.Success)
                        {
                            if (writeToNewSheet)
                                dataModel.WriteColumnValuesToNewSheet(fileNameMentionFile, sheetName, newValuesTable, firstRowHasHeaders);
                            else
                                dataModel.WriteColumnValues(fileNameMentionFile, sheetName, newValuesTable, firstRowHasHeaders);
                        }
                    }
                }
            }
            else
            {
                rc = ReturnCode.ColumnsAlreadyExist;
            }

            return rc;
        }

        private bool DoRandomSelectColumnsAlreadyExist(string fileName, string[] randomSelectColumnNames, string sheetName)
        {
            Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);

            bool newColumnAlreadyExists = false;

            // For each column to contain the split values...
            foreach (string newColumnName in randomSelectColumnNames)
            {
                // If that column already exists in the sheet...
                if (columnNames.ContainsValue(newColumnName))
                    newColumnAlreadyExists = true;
            }

            return newColumnAlreadyExists;
        }

    }
}
