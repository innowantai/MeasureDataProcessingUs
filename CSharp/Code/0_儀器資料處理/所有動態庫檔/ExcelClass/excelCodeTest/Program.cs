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
            string fPath = @"D:\Users\95074\Desktop\ExcelSheetDataReadTest\SheetTest.xlsx";



            string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            List<int> res = new List<int>();
            int tat = 26789789;
            TransExcelPo(tat, ref res);
            res.Reverse();
            Console.WriteLine(tat);
            string po = "";
            foreach (int r in res)
            {
                po += Engpo[r];
            }

            Console.WriteLine(po);
            //EXCEL excel = new EXCEL(fPath);
            //List<string> sheets = excel.sheets; 
            //foreach (string ss in sheets)
            //{
            //    Console.WriteLine(ss);
            //}
            //string[,] data = excel.ReadBySheetName(1, 1, excel.sheets[0]);

            //Dictionary<string, string[,]>  resData = excel.GetSheetsData(); 
        }


        private static void TransExcelPo(int Num, ref List<int> res)
        {
            if (Num == 0) return;

            if (Num % 26 != 0 )
            {
                int rr = Num % 26;
                res.Add(rr);
                Num = (Num - rr) / 26 > 0 ? (Num - rr) / 26 : Num - rr;
                TransExcelPo(Num, ref res);
            }
            else
            { 
                res.Add(Num);
            } 
        }
    }
}
