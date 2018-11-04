using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelClass;

namespace excelCodeTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string fPath = @"C:\Users\Wantai\Desktop\ExcelUpdate\ExcelUpdate\bin\Debug\test.xlsx";
            List<string> sheets = ExcelSaveAndRead.GetSheets(fPath);
            foreach (var ss in sheets)
            {
                Console.WriteLine(ss);
            }

            string[,] tdata = ExcelSaveAndRead.ReadBySheetName(fPath, 1, 1, "萬萬好帥");


        }
    }
}
