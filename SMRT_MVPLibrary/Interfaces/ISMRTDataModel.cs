using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace TwinArch.SMRT_MVPLibrary.Interfaces
{
    public interface ISMRTDataModel : ISMRTBase, IDisposable
    {
        /// <summary>
        /// Get the list of sheets in an Excel workbook.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <returns>A list of sheet names</returns>
        List<string> GetSheetNames(string fileName);

        /// <summary>
        /// Get the column identifiers and names found within a worksheet of an Excel workbook.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the sheet</param>
        /// <returns>A dictionary containing the column identifiers, as the key, and the column names, as the value, of the 
        /// columns in the sheet</returns>
        Dictionary<string, string> GetColumnNames(string fileName, string sheetName);

        /// <summary>
        /// Get all values appearing within a column in a sheet in an Excel workbook.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the sheet</param>
        /// <param name="columnNames">The name of the column</param>
        /// <returns>A list of cell identifiers and values for all cells in the column that contain a value</returns>
        DataTable GetColumnValuesForColumnNames(string fileName, string sheetName, string[] columnNames, bool firstRowHasHeaders);

        /// <summary>
        /// Adds columns to a sheet in a workbook. If the column already exists, it is not added.
        /// </summary>
        /// <param name="fileName">The full path to the Excel workbook</param>
        /// <param name="sheetName">The name of the sheet</param>
        /// <param name="columnNames">The names of the columns to be added</param>
        /// <returns>A ReturnCode value.</returns>
        ReturnCode AddColumns(string fileName, string sheetName, string[] columnNames);

        /// <summary>
        /// Determines whether the file is a valid Excel file to use with this DataModel. The file must exist and be
        /// in XLSX format.
        /// </summary>
        /// <param name="FileName">The full path to the Excel workbook.</param>
        /// <returns></returns>
        bool FileIsValid(string fileName);

        /// <summary>
        /// Writes values to existing columns in an existing spreadsheet.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNames"></param>
        /// <param name="newValues"></param>
        /// <param name="firstRow"></param>
        /// <returns></returns>
        ReturnCode WriteColumnValues(string fileName, string sheetName, DataTable newValuesTable, bool firstRowHasHeaders);

        /// <summary>
        /// Writes values to a new sheet.
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="sheetName"></param>
        /// <param name="columnNames"></param>
        /// <param name="newValues"></param>
        /// <param name="firstRow"></param>
        /// <returns></returns>
        ReturnCode WriteColumnValuesToNewSheet(string fileName, string sheetName, DataTable newValuesTable, bool firstRowHasHeaders);

        ReturnCode GetTwitterUserInfo(string userID, ref TwitterUserInfo twitterUserInfo);
    }
}
