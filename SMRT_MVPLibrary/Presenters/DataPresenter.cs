using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwinArch.SMRT_MVPLibrary.Interfaces;

namespace TwinArch.SMRT_MVPLibrary.Presenters
{
    public class DataPresenter : Presenter<ISMRTMainView>, IDisposable
    {
        public DataPresenter(ISMRTMainView view, int useAutomation)
            : base(view, useAutomation)
        {
        }

        public void Dispose()
        {
            Model.Dispose();
        }

        public void DisplaySheetNames(string filename)
        {
            try
            {
                if (View != null) View.SheetNames = Model.GetSheetNames(filename);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }
        }

        public void DisplayColumnNames(string filename, string sheetname)
        {
            try
            {
                if (View != null) View.ColumnNames = Model.GetColumnNames(filename, sheetname);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }
        }

        public void EmptySheetNames()
        {
            if (View != null) View.SheetNames = null;
        }

        public void EmptyColumnNames()
        {
            if (View != null) View.ColumnNames = null;
        }

        public ReturnCode ParseURLs(string fileName, string sheetName, string columnName, bool overwriteExistingData, bool ignoreFirstRow)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.SplitURLs(fileName, sheetName, columnName, overwriteExistingData, ignoreFirstRow);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode AddTwitterInfo(string fileName, string sheetName, bool overwriteExistingData, bool ignoreFirstRow, int numUsersToRetrieve)
        {
            ReturnCode rc = ReturnCode.Failed;
            try
            {
                rc = Model.AddTwitterInfo(fileName, sheetName, overwriteExistingData, ignoreFirstRow, numUsersToRetrieve);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode Autocode(string fileNameMentionFile, string sheetName, string firstColumnName, string secondColumnName, string fileNameAutocodeFile, bool overwriteExistingData, bool ignoreFirstRow)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.Autocode(fileNameMentionFile, sheetName, firstColumnName, secondColumnName, fileNameAutocodeFile, overwriteExistingData, ignoreFirstRow);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode RandomSelectOneGroup(string fileNameMentionFile, string sheetName, string columnNameAutocodeCounts, string columnNameRandomSelect, int percent, int floor, int ceiling, bool overwriteExistingData, bool ignoreFirstRow)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.RandomSelectOneGroup(fileNameMentionFile, sheetName, columnNameAutocodeCounts, columnNameRandomSelect, percent, floor, ceiling, overwriteExistingData, ignoreFirstRow);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode RandomSelectMultipleGroups(string fileNameMentionFile, string sheetName, string columnNameAutocodeCounts, string columnNameRandomSelect, int numberOfGroups, bool overwriteExistingData, bool ignoreFirstRow)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.RandomSelectMultipleGroups(fileNameMentionFile, sheetName, columnNameAutocodeCounts, columnNameRandomSelect, numberOfGroups, overwriteExistingData, ignoreFirstRow);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode CalculateWordFrequency(string fileNameMentionFile, string sheetName, string columnNameWordText, string fileNameStopList, string fileNameOutput, int minPhraseLength, int maxPhraseLength, int minFrequency, bool ignoreNumericOnlyWords, bool ignoreFirstRow)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.CalculateWordFrequency(fileNameMentionFile, sheetName, columnNameWordText, fileNameStopList, fileNameOutput, minPhraseLength, maxPhraseLength, minFrequency, ignoreNumericOnlyWords, ignoreFirstRow);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }

        public ReturnCode GetTwitterConnections(string fileNameSourceIDsFile, string sheetName, string columnNameTwitterIDs, int maxConnections, string fileNameOutput)
        {
            ReturnCode rc = ReturnCode.Failed;

            try
            {
                rc = Model.GetTwitterConnections(fileNameSourceIDsFile, sheetName, columnNameTwitterIDs, maxConnections, fileNameOutput);
            }
            catch (System.IO.FileNotFoundException)
            {
                View.IsFileValid = false;
            }
            catch (Exception e)
            {
                View.UnhandledException = e.Message + "\n\nThis additional info should be passed on to your tech support:\n\n\n" + e.ToString();
            }

            return rc;
        }
    }
}
