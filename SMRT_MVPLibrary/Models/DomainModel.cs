﻿using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class DomainModel : ISMRTDomain, IDisposable
    {

        string[] newURLSplitColumns = { "MentionType", "Domain", "PosterID", "MentionID" };

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
                    List<KeyValuePair<string, string>> columnValues = dataModel.GetColumnValues(fileName, sheetName, urlColumnName, firstRowHasHeaders);

                    // Keep track of processing and failure counts for a short-circuit abort if needed
                    int numprocessed = 0;
                    int numFailed = 0;

                    // Keep a list of all the split out info. Ensure that there is exactly one for each value in the column.
                    //List<MentionPart> newValues = new List<MentionPart>();
                    //List<KeyValuePair<string, List<string>>> newValues = new List<KeyValuePair<string, List<string>>>();
                    Dictionary<string, List<string>> newValues = new Dictionary<string,List<string>>();
                    foreach (string splitColumnName in newURLSplitColumns)
                        newValues.Add(splitColumnName, new List<string>());
                    List<string> splitOutValues;

                    // For each cell/value...
                    foreach (KeyValuePair<string, string> pair in columnValues)
                    {
                        try
                        {
                            // Try to parse it as a URI
                            Uri uri = new Uri(pair.Value);

                            // Get the domain part.
                            string domain = uri.GetComponents(UriComponents.Host, UriFormat.Unescaped);

                            // Get the segments of the URI - these are the slash-delimited pieces after the domain.
                            string[] segments = uri.Segments;
                            NameValueCollection queryCollection = HttpUtility.ParseQueryString(uri.Query);

                            if (domain.Equals("twitter.com"))
                            {
                                splitOutValues = GetTwitterParts(uri, domain, segments, queryCollection);
                            }
                            else if (domain.Equals("facebook.com"))
                            {
                                splitOutValues = GetFacebookParts(uri, domain, segments, queryCollection);
                            }
                            else if (domain.Contains("blogspot.com"))
                            {
                                splitOutValues = GetBloggerParts(uri, domain, segments, queryCollection);
                            }
                            else
                                splitOutValues = new List<string>() { null, null, null, null };
                        }
                        catch (UriFormatException e)
                        {
                            numFailed++;
                            splitOutValues = new List<string>() {null, null, null, null};
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

                        int columnNumber = 0;
                        foreach (string columnName in newURLSplitColumns)
                        {
                            newValues[columnName].Add(splitOutValues[columnNumber]);
                            columnNumber++;
                        }

                    }

                    if (rc == ReturnCode.Success)
                    {
                        dataModel.WriteColumnValues(fileName, sheetName, newValues, firstRowHasHeaders);
                    }
                }
            }
            else
            {
                rc = ReturnCode.ColumnsAlreadyExist;
            }

            return rc;
        }

        private List<string> GetFacebookParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection)
        {
            // http://facebook.com/105indaklubb/posts/10151141293211990
            // http://facebook.com/events/122767261239362/permalink/122767264572695
            // http://facebook.com/media/set/?set=a.10151209676509212.454268.38951299211&type=1
            // http://facebook.com/notes/complex-child-e-magazine/childrens-mental-health-edition/10151458448874231
            // http://facebook.com/permalink.php?story_fbid=100127316863475&id=100005986200023


            List<string> newvalues = new List<string>();
            try
            {
                if (segments[1].Equals("permalink.php"))
                {
                    newvalues.Add("FacebookPost");
                    newvalues.Add(domain);
                    newvalues.Add(queryCollection["id"]);
                    newvalues.Add(queryCollection["story_fbid"]);
                }
                else if (segments[1].Equals("events/"))
                {
                    newvalues.Add("FacebookEvent");
                    newvalues.Add(domain);
                    newvalues.Add(segments[4].Trim('/'));
                    newvalues.Add(segments[2].Trim('/'));
                }
                else if (segments[1].Equals("media/"))
                {
                    newvalues.Add("FacebookMedia");
                    newvalues.Add(domain);
                    newvalues.Add("");
                    newvalues.Add(queryCollection["set"]);
                }
                else if (segments[1].Equals("notes/"))
                {
                    newvalues.Add("FacebookNote");
                    newvalues.Add(domain);
                    newvalues.Add(segments[2].Trim('/'));
                    newvalues.Add(segments[4].Trim('/'));
                }
                else
                {
                    newvalues.Add("FacebookGroupPost");
                    newvalues.Add(domain);
                    newvalues.Add(segments[1].Trim('/'));
                    newvalues.Add(segments[3].Trim('/'));
                }
            }
            catch (Exception e)
            {
                newvalues.Add("Facebook");
                newvalues.Add(domain);
                newvalues.Add(null);
                newvalues.Add(null);
            }

            return newvalues;
        }

        private List<string> GetTwitterParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection)
        {
            List<string> newvalues = new List<string>();
            newvalues.Add("Twitter");
            newvalues.Add(domain);

            newvalues.Add(segments[1].Trim('/'));
            newvalues.Add(segments[3].Trim('/'));

            return newvalues;
        }

        private List<string> GetBloggerParts(Uri uri, string domain, string[] segments, NameValueCollection queryCollection)
        {
            List<string> newvalues = new List<string>();
            newvalues.Add("Blogger");
            newvalues.Add(domain);

            // Call blogs/byurl with the full URL
            // Then call posts/bypath using the user ID from the first one and the part of the URL after the domain.
            newvalues.Add("");
            newvalues.Add("");

            return newvalues;
        }


    }
}
