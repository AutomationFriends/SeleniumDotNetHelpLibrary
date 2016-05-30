using Microsoft.Office.Interop.Excel;
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
    public class WriteToExcel
    {
        public void WriteToRange(string fileName, double[] data, string startCell, string endCell)
        {

            Excel.Application app = null;
            Excel.Workbooks books = null;
            Excel.Workbook book = null;
            Excel.Sheets sheets = null;
            Excel.Worksheet sheet = null;
            Excel.Range range = null;

            try
            {
               app = new Excel.Application();
               books = app.Workbooks;
               book = books.Open(Path.GetFullPath(@"TestData\" + fileName + ".xlsx"), Type.Missing, false,
                                              Type.Missing, Type.Missing, Type.Missing, false, Type.Missing,
                                              Type.Missing, true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
               sheets = book.Sheets;
               sheet = sheets.get_Item(1);
               range = sheet.get_Range(startCell, endCell);
               range.NumberFormat = "General";
               range.Value2 = data;
               book.Save();
               book.Close();
               app.Quit();

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

            /*
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Open(Path.GetFullPath(@"TestData\" + fileName + ".xlsx"), Type.Missing, false, 
                                              Type.Missing, Type.Missing, Type.Missing, false, Type.Missing, 
                                              Type.Missing, true, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            Excel.Range writeRange = xlWorkSheet.get_Range(startCell, endCell);
            writeRange.NumberFormat = "General";
            writeRange.Value2 = data;

            xlWorkBook.Save();
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            

            while (Marshal.ReleaseComObject(xlApp) != 0) { }
            while (Marshal.ReleaseComObject(xlWorkBook) != 0) { }
            while (Marshal.ReleaseComObject(xlWorkSheet) != 0) { }

            xlApp = null;
            xlWorkBook = null;
            xlWorkSheet = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            */

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
