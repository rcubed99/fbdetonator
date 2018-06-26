using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace fbdetonator
{
    public class ExcelReader
    {
        public static void getExcelFile()
        {
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            string path = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\fbdonate.xlsx");
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"c:\Users\rrojas\source\repos\fbdetonator\fbdetonator\data\fbdonate.xlsx");
            //Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"Data\fbdonate.xlsx"));
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@path);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            //int pageDonateCount = 0;
            int postDonateCount = 0;
            //int fundraiserDonateCount = 0;

            string tmpStr = "";
            string rowStr = "";
            string colStr = "";

            FBDonateRules FBRules = new FBDonateRules();

            MainWindow win = (MainWindow)System.Windows.Application.Current.MainWindow;
            win.SetStatusText("Getting started!");

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= rowCount; i++)
            {

                for (int j = 1; j <= colCount; j++)
                {
                    //new line
                    if (j == 1)
                    {
                        Debug.WriteLine("\r\n");
                    }
                    //write the value to the console
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        colStr = xlRange.Cells[i, j].Value2.ToString();
                        //tmpStr = xlRange.Cells[i, j].Value2.ToString() + " ";
                        tmpStr = colStr + " ";
                        rowStr += tmpStr;
                        Debug.Write(tmpStr);
                        if (j == (int)FBDonateRules.Col.Source && rowCount > 1)
                        {
                            if (FBRules.IsPostDonation(colStr))
                            {
                                postDonateCount++;
                                win.SetPostDonateCount(postDonateCount);
                            }
                        }
                        colStr = "";
                    }
                }
                win.SetStatusText(rowStr);
                rowStr = "";
                win.SetRowCount(i);
            }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
    }
}
