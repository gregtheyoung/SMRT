using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TwinArch.SMRT_MVPLibrary.Interfaces;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;

namespace TwinArch.SMRT_MVPLibrary.Models
{
    public class ExcelAutomationModel : ISMRTDataModel
    {

        private Application app;

        private Application ExcelApp
        {
            get
            {
                if (app == null)
                    app = new Application();
                return app;
            }
        }

        public void Dispose()
        {
            if (app != null)
                app.Quit();
        }

        private Workbook OpenWorkbook(string fileName)
        {
            Workbook book = null;

            if (!String.IsNullOrEmpty(fileName))
                book = ExcelApp.Workbooks.Open(fileName);

            return book;
        }

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            Workbook book = OpenWorkbook(fileName);
            if (book != null)
            {
                foreach (Worksheet sheet in book.Sheets)
                    sheetNames.Add(sheet.Name);

                book.Close(false);
            }
                            
            return sheetNames;
        }

        public Dictionary<string, string> GetColumnNames(string fileName, string sheetName)
        {
            Dictionary<string, string> columnNames = new Dictionary<string, string>();

            Workbook book = OpenWorkbook(fileName);
            if ((book != null) && !String.IsNullOrEmpty(sheetName))
            {
                Worksheet sheet = book.Sheets[sheetName];
                Range firstRow = sheet.get_Range("A1", sheet.get_Range("A1").get_End(XlDirection.xlToRight));

                foreach (Range cell in firstRow.Cells)
                    columnNames.Add(cell.get_Address(Type.Missing, false).Substring(0,1), cell.get_Value().ToString());

                book.Close(false);
            }

            return columnNames;
        }


        public List<KeyValuePair<string, string>> GetColumnValues(string fileName, string sheetName, string columnName)
        {
            List<KeyValuePair<string, string>> columnValues = new List<KeyValuePair<string, string>>();

            Workbook book = OpenWorkbook(fileName);

            if ((book != null) && !String.IsNullOrEmpty(sheetName) && !String.IsNullOrEmpty(columnName))
            {
                Worksheet sheet = book.Sheets[sheetName];
                Range range = sheet.get_Range(columnName + ":" + columnName);

                foreach (Range cell in range.Cells)
                {
                    if (cell.Value2 != null)
                        columnValues.Add(new KeyValuePair<string, string>(cell.get_Address(), cell.Value2));
                }

                book.Close(false);
            }

            return columnValues;
        }


    }
}
