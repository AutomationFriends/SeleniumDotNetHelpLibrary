using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumDotNetHelpLibrary.Selenium.Support.DataImport
{
    public class ReadFromExcel
    {
        public List<string> GetWorkSheet(string fileName, int worksheetNumber, string rangeStart, string rangeEnd)
        {
            // ->>>>> int userRow, int userCol
            // Reference to Excel Application.

            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook book = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range range = null;

            string startupPath = System.IO.Directory.GetCurrentDirectory();
     
            try
            {
                app = new Excel.Application();
                app.DisplayAlerts = false;
                books = app.Workbooks;
                //book = books.Open(Path.GetFullPath(Path.GetFullPath(@"TestData\" + fileName + ".xlsx")));

                book = books.Open(fileName);

                sheets = book.Sheets;
                sheet = sheets.get_Item(1);

                range = sheet.get_Range(rangeStart, rangeEnd);
                //range.NumberFormat = "General";
                //range.Value2 = data;
                
                object[,] cellValues = (object[,])range.Value2;
                //List<double> lst = cellValues.Cast<object>().ToList().ConvertAll(x=> Convert.ToDouble(x));

                List<string> lst = cellValues.Cast<object>().ToList().ConvertAll(x => Convert.ToString(x));

                book.Close();
                app.Quit();

                //Console.WriteLine(lst[0]);
                //Console.WriteLine(lst[1]);
                //Console.WriteLine(lst[2]);

                return lst;

            }

            finally
            {
                if (range != null) Marshal.ReleaseComObject(range);
                if (sheet != null) Marshal.ReleaseComObject(sheet);
                if (sheets != null) Marshal.ReleaseComObject(sheets);
                if (book != null) Marshal.ReleaseComObject(book);
                if (books != null) Marshal.ReleaseComObject(books);
                if (app != null) Marshal.ReleaseComObject(app);
            }


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
             
            //Excel.Application xlApp = new Excel.Application();
            


            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.GetFullPath(@"TestData\" + fileName + ".xlsx"));
            //Excel.Worksheet xlWorksheet = (Excel.Worksheet)xlWorkbook.Sheets.get_Item(worksheetNumber);

            //Excel.Range xlRange = xlWorksheet.UsedRange;


            
            //

            //double[] valueArray = (double[])xlRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault);

            /*

            // Cleanup
            xlWorkbook.Close(false);
            xlApp.Quit();

            // Manual disposal because of COM
            while (Marshal.ReleaseComObject(xlApp) != 0) { }
            while (Marshal.ReleaseComObject(xlWorkbook) != 0) { }
            while (Marshal.ReleaseComObject(xlRange) != 0) { }

            xlApp = null;
            xlWorkbook = null;
            xlRange = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();


            */
            
        }

    }
}
