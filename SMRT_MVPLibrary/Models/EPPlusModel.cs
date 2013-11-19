using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;
using System.IO;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using OfficeOpenXml;


namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class EPPlusModel : ISMRTDataModel
    {

        public void Dispose()
        {
        }

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName))
            {
                using (ExcelPackage pkg = new ExcelPackage(new FileInfo(fileName)))
                {
                    foreach (ExcelWorksheet sheet in pkg.Workbook.Worksheets)
                    {
                        string sheetName = sheet.Name;
                        sheetNames.Add(sheetName);
                    }
                }
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {
                using (ExcelPackage pkg = new ExcelPackage(new FileInfo(fileName)))
                {
                    ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];
                    ExcelRange firstRow = sheet.Cells["A1:XDR1"];
                    foreach (ExcelRangeBase cell in firstRow)
                        columnNames.Add(cell.Address, cell.Text);
                }
            }

            return columnNames;
        }

        public List<KeyValuePair<string, string>> GetColumnValues(string fileName, string sheetName, string columnID)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && !string.IsNullOrEmpty(columnID))
            {
                using (ExcelPackage pkg = new ExcelPackage(new FileInfo(fileName)))
                {
                    ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                    ExcelRange cells = sheet.Cells[columnID + ":" + columnID[0] + "1000000"];
                    foreach (ExcelRangeBase cell in cells)
                    {
                        columnValues.Add(new KeyValuePair<string, string>(cell.Address, cell.Text));
                    }

                }
            }

            return columnValues;
        }

        public ReturnCode AddColumn(string fileName, string sheetName, string[] columnNames)
        {
            ReturnCode rc = ReturnCode.Failed;

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName) && (columnNames.Length >= 0))
            {
                using (ExcelPackage pkg = new ExcelPackage(new FileInfo(fileName)))
                {
                    ExcelWorksheet sheet = pkg.Workbook.Worksheets[sheetName];

                    ExcelRange firstRow = sheet.Cells["A1:XDR1"];
                    foreach (string columnName in columnNames)
                    {
                        if (firstRow.FirstOrDefault<ExcelRangeBase>(cell => cell.Text.Equals(columnName)) == null)
                        {
                            int columnNumber = firstRow.Count<ExcelRangeBase>() + 1;
                            sheet.Cells[1, columnNumber].Value = columnName;
                        }
                    }

                    pkg.Save();
                    rc = ReturnCode.Success;
                }
            }
        

            return rc;
        }
    }
}
