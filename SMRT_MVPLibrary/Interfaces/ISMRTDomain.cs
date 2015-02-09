using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTDomain : ISMRTBase, IDisposable
    {
        /// <summary>
        /// Gets the names of worksheets in an Excel file.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook file</param>
        /// <returns>A list of the names of the columns in the workbook</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        List<string> GetSheetNames(string fileName);

        /// <summary>
        /// Gets the column identifiers and names from a worksheet in an Excel file
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <returns>A list of the identifiers and names of the columns in the worksheet.</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        Dictionary<string, string> GetColumnNames(string fileName, string sheetName);

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
        /// <returns>A return code indicating success or failure</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        ReturnCode SplitURLs(string fileName, string sheetName, string urlColumnID, bool overwriteExistingData, bool ignoreFirstRow);

        /// <summary>
        /// After verifying that the sheet has the expected fields (those created by the SplitURLs function), for each row that
        /// has a MentionType of "Twitter" it will add info about the tweet and tweeter to additional columns.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the worksheet</param>
        /// <param name="overwriteExistingData"></param>
        /// <returns>A return code indicating success or failure</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the file is not a valid Excel file. It must
        /// exist, not be open by another process, and be an XLSX file (not an older XLS file).</exception>
        ReturnCode AddTwitterInfo(string fileName, string sheetName, bool overwriteExistingData, bool ignoreFirstRow, int numUsersToRetrieve);

        ReturnCode Autocode(string fileNameMentionFile, string sheetName, string columnName, string fileNameAutocodeFile, bool overwriteExistingData, bool ignoreFirstRow);
        ReturnCode RandomSelect(string fileNameMentionFile, string sheetName, string columnNameAutocodeCounts, string columnNameRandomSelect, int percent, int floor, int ceiling, bool overwriteExistingData, bool ignoreFirstRow);
    }
}
