using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace xlsx_to_sqlite
{
    class Program
    {
        static void Main(string[] args)
        {
            Excel.Application exApp = new Excel.Application();

            Console.WriteLine("Please enter the full path of excel workbook");
            string bookDir = Console.ReadLine();

            Excel.Workbook exBook = exApp.Workbooks.Open(@bookDir);
            Excel.Worksheet exSheet = exBook.Sheets[1];
            Excel.Range exRange = exSheet.UsedRange;

            Console.ReadLine();
        }
    }
}
