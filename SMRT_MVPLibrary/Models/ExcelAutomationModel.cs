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
    public class ExcelAutomationModel : ISMRTModel
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

        public List<string> GetSheetNames(string fileName)
        {
            List<string> sheetNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName))
            {
                Workbook book = ExcelApp.Workbooks.Open(fileName);

                foreach (Worksheet sheet in book.Sheets)
                    sheetNames.Add(sheet.Name);

                ExcelApp.Quit();
            }
                            
            return sheetNames;
        }

        public List<string> GetColumnNames(string fileName, string sheetName)
        {
            List<string> columnNames = new List<string>();

            if (!String.IsNullOrEmpty(fileName) && !String.IsNullOrEmpty(sheetName))
            {

                Workbook book = ExcelApp.Workbooks.Open(fileName);
                Worksheet sheet = book.Sheets[sheetName];
                Range range = sheet.get_Range("A1");

                Range firstRow = sheet.get_Range("A1", sheet.get_Range("A1").get_End(XlDirection.xlToRight));

                foreach (Range cell in firstRow.Cells)
                    columnNames.Add(cell.get_Value().ToString());

                ExcelApp.Quit();
            }

            return columnNames;
        }
    }
}
