using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class DomainModel : ISMRTDoman, IDisposable
    {

        string[] newURLSplitColumns = { "Domain", "PosterID", "MentionID" };

        protected ISMRTDataModel dataModel;
        /// <summary>
        /// 
        /// </summary>
        /// <param name="dataModelToUse">Which data model to use: 0=ExcelAutomation, 1=OLEDB Jet, 2=OLEDB ACE</param>
        public DomainModel(int dataModelToUse)
        {
            switch (dataModelToUse)
            {
                case (0):
                {
                    dataModel = new ExcelAutomationModel();
                    break;
                }
                case (1):
                {
                    dataModel = new OLEDBModel(true);
                    break;
                }
                case (2):
                {
                    dataModel = new OLEDBModel(false);
                    break;
                }
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

        public List<string> GetSheetNames(string filename)
        {
            return dataModel.GetSheetNames(filename);
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            return dataModel.GetColumnNames(fileName, sheetName);
        }

        private bool DoURLSplitColumnsAlreadyExist(string fileName, string sheetName)
        {
            // Check to make sure the new columns don't exist yet.
            Dictionary<string, string> columnNames = GetColumnNames(fileName, sheetName);

            bool newColumnAlreadyExists = false;

            foreach (string newColumnName in newURLSplitColumns)
            {
                if (columnNames.ContainsValue(newColumnName))
                    newColumnAlreadyExists = true;
            }

            return newColumnAlreadyExists;
        }

        public ReturnCode SplitURLs(string fileName, string sheetName, string urlColumnName, bool overwriteExistingData)
        {
            ReturnCode rc = ReturnCode.Success;

            if (!overwriteExistingData && DoURLSplitColumnsAlreadyExist(fileName, sheetName))
                return ReturnCode.ColumnsAlreadyExist;
            else
                AddURLSplitColumns(fileName, sheetName);

            List<KeyValuePair<string, string>> columnValues = dataModel.GetColumnValues(fileName, sheetName, urlColumnName);

            int numprocessed = 0;
            int numFailed = 0;
            foreach (KeyValuePair<string, string> pair in columnValues)
            {
                try
                {
                    Uri uri = new Uri(pair.Value);
                    string domain = uri.GetComponents(UriComponents.Host, UriFormat.Unescaped);
                    string[] segments = uri.Segments;
                    NameValueCollection queryCollection = HttpUtility.ParseQueryString(uri.Query);

                    if (domain.Equals("twitter.com"))
                    {
                        GetTwitterParts(uri);
                    }
                    else if (domain.Equals("facebook.com"))
                    {
                        GetFacebookParts(uri);
                    }
                    else if (domain.Contains("blogspot.com"))
                    {
                        GetBloggerParts(uri);
                    }
                }
                catch (UriFormatException e)
                {
                    numFailed++;
                }
                numprocessed++;

                // If more than 1/2 have failed, then this likely is not a column that contains URLs.
                // Only check this condition after we have processed some of them - don't want to error just because 
                // one of the first two failed.
                if ((numprocessed >= 50) && (numFailed >= numprocessed/2))
                {
                    rc = ReturnCode.NotURLColumn;
                    break;
                }
            }

            return rc;
        }

        private void AddURLSplitColumns(string fileName, string sheetName)
        {
            dataModel.AddColumn(fileName, sheetName, newURLSplitColumns);
        }

        private void GetFacebookParts(Uri uri)
        {
            string userID;
            string storyID;

            string[] segments = uri.Segments;
            if (segments[1].Equals("permalink.php"))
            {
                NameValueCollection queryStrings = HttpUtility.ParseQueryString(uri.Query);
                userID = queryStrings["id"];
                storyID = queryStrings["story_fbid"];
            }
            else
            {
                userID = segments[1].Trim('/');
                storyID = segments[3].Trim('/');
            }
        }

        private void GetTwitterParts(Uri uri)
        {
            string[] segments = uri.Segments;
            string user = segments[1].Trim('/');
            string tweetID = segments[3].Trim('/');
        }

        private void GetBloggerParts(Uri uri)
        {
            // Call blogs/byurl with the full URL
            // Then call posts/bypath using the user ID from the first one and the part of the URL after the domain.
            string[] segments = uri.Segments;
        }


    }
}
