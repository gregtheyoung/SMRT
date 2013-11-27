using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Data;
using System.Configuration;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class DomainModel : ISMRTDomain, IDisposable
    {
        int MaxTwitterUsersToPull = 10;

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

            if (!String.IsNullOrEmpty(ConfigurationManager.AppSettings["MaxTwitterUsersToPull"]))
                MaxTwitterUsersToPull = Convert.ToInt32(ConfigurationManager.AppSettings["MaxTwitterUsersToPull"]);
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

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileName))
                throw new System.IO.FileNotFoundException();

            // If the existing split out columns are to be overwritten, or if they don't exist yet...
            if (overwriteExistingData || !DoURLSplitColumnsAlreadyExist(fileName, sheetName))
            {
                // Add the new columns to the sheet...
                rc = dataModel.AddColumns(fileName, sheetName, newURLSplitColumns);

                // If they were added successfully...
                if (rc == ReturnCode.Success)
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

        public ReturnCode AddTwitterInfo(string fileName, string sheetName, bool overwriteExistingData, bool firstRowHasHeaders)
        {
            ReturnCode rc = ReturnCode.Success;

            // Check that the file is valid, throw exception if not
            if (!dataModel.FileIsValid(fileName))
                throw new System.IO.FileNotFoundException();

            // Need to make sure that the SplitURL columns do exist
            if (!DoURLSplitColumnsAlreadyExist(fileName, sheetName))
                rc = ReturnCode.ColumnsMissing;
            else
            {
                // If the existing split out columns are to be overwritten, or if they don't exist yet...
                if (overwriteExistingData || !DoTwitterColumnsAlreadyExist(fileName, sheetName))
                {
                    // Add the new columns to the sheet...
                    rc = dataModel.AddColumns(fileName, sheetName, newTwitterColumns);

                    // If they were added successfully...
                    if (rc == ReturnCode.Success)
                    {
                        // Get the contents of the column. This will provide the row/cell identifier (depends on the
                        // underlying data model implementation) and the value of column for that row/cell.
                        DataTable posterInfoTable = dataModel.GetColumnValuesForColumnNames(fileName, sheetName,
                            newURLSplitColumns,
                            firstRowHasHeaders);

                        List<string> topTwitterPosters = GetTopTwitterPosters(posterInfoTable);

                        // We will only use the top N
                        if (topTwitterPosters.Count > MaxTwitterUsersToPull)
                            topTwitterPosters.RemoveRange(MaxTwitterUsersToPull, topTwitterPosters.Count - MaxTwitterUsersToPull);

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

    }
}
