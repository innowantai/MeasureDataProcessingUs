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
            string fPath = @"C:\Users\Wantai\Desktop\test\tt.xlsx";



            //string[] Engpo = new string[] { "", "A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z" };
            //List<int> res = new List<int>();
            //int tat = 5;
            //TransExcelPo(tat, ref res);
            //res.Reverse();
            //Console.WriteLine(tat);
            //string po = "";
            //foreach (int r in res)
            //{
            //    po += Engpo[r];
            //}

            //Console.WriteLine(po);

            EXCEL excel = new EXCEL(fPath);


            //List<string> sheets = excel.sheets;
            //foreach (string SS in sheets)
            //{
            //    Console.WriteLine(SS);
            //}
            //string[,] data = excel.GetDataBySheetName(1, 1, excel.sheets[0]);

            string[,] data = new string[10, 10];
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    data[i, j] = i.ToString();
                }
            }

            //excel.Save_To(@"C:\Users\Wantai\Desktop\test\tt.xlsx","test",6,1, data); 
            excel.Save( "1234", 5, 5, data);
            excel.close();
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
