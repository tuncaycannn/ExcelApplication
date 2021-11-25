using System;
using Microsoft.Office.Interop.Excel;

namespace ExcelApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            Application application = new Application();

            if (application == null)
            {
                Console.WriteLine("Excel is not installed!!");
                return;
            }

            application.Workbooks.Open(@"C:\Users\user\Desktop\FindeksRapor_30.06.2021-dönüştürüldü.xlsx");
            var excelSheet1 = application.Sheets.Count;

            for (int i = 1; i <= 8; i++)
            {
                _Worksheet excelSheets = application.Sheets[i];
                Range excelRange = excelSheets.UsedRange;
                int rowCount = excelRange.Rows.Count;
                int colCount = excelRange.Columns.Count;


                Console.Write("\r\n");

                for (int n = 3; n <= 8; n++)
                {
                    Console.Write("\r\n");
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (excelRange.Cells[n, j] != null && excelRange.Cells[n, j].Value2 != null)
                            Console.Write(excelRange.Cells[n, j].Value2.ToString() + "\t");
                    }
                }
                Console.Write("\r\n");

                for (int n = 9; n <= 14; n++)
                {
                    Console.Write("\r\n");
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (excelRange.Cells[n, j] != null && excelRange.Cells[n, j].Value2 != null)
                            Console.Write(excelRange.Cells[n, j].Value2.ToString() + "\t");
                    }
                }
                Console.Write("\r\n");

                for (int n = 15; n <= 20; n++)
                {
                    Console.Write("\r\n");
                    for (int j = 1; j <= colCount; j++)
                    {
                        if (excelRange.Cells[n, j] != null && excelRange.Cells[n, j].Value2 != null)
                            Console.Write(excelRange.Cells[n, j].Value2.ToString() + "\t");
                    }
                }
            }

            application.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(application);
            Console.ReadLine();
        }

    }
}
