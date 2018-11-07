using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelClass;

namespace excelCodeTest
{
    class Program
    {
        /// 桌面的路徑
        public static string DeskPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

        static void Main(string[] args)
        {

            /// 要讀取的Excel檔案名稱
            string ExcelName = "新增 Microsoft Excel 工作表.xlsx";

            /// 要讀取的Excel檔案路徑
            string fPath = Path.Combine(DeskPath, ExcelName);

            /// 建立Excel 物件
            EXCEL Excel = new EXCEL(fPath);

            /// 讀取所有sheet的資料
            Dictionary<string, string[,]> AllSheetsData = Excel.GetSheetsData();

            /// 顯示所有Sheet資料
            ShowAllExcelData(AllSheetsData);


        }


        private static void Example()
        {
            /// 要讀取的Excel檔案名稱
            string ExcelName = "sheetData.xlsx";

            /// 要另外儲存的Excel檔案名稱
            string SaveName = "test.xlsx";

            /// 要讀取的Excel檔案路徑
            string fPath = Path.Combine(DeskPath, ExcelName);

            /// 建立Excel 物件
            EXCEL Excel = new EXCEL(fPath);

            /// 取得Excel所有sheet名稱
            List<string> sheets = Excel.sheets;

            /// 讀取第ii個sheet資料,起始位置為(1,1)
            int ii = 0;
            string[,] data = Excel.GetDataBySheetName(1, 1, sheets[ii]);


            /// 讀取所有sheet的資料
            Dictionary<string, string[,]> AllSheetsData = Excel.GetSheetsData();


            /// 隨意創立要儲存的資料陣列
            string[,] saveData = new string[10, 10];
            for (int i = 0; i < 10; i++)
            {
                for (int j = 0; j < 10; j++)
                {
                    saveData[i, j] = i.ToString();
                }
            }


            /// 儲存至所開啟之Excel
            Excel.Save("儲存資料至本Excel中", 1, 1, saveData);


            /// 另外儲存的檔案路徑
            string savePath = Path.Combine(DeskPath, SaveName);
            Excel.Save_To(savePath, "儲存資料至其他Excel", 1, 1, saveData);

            Excel.close();



            //////// 顯示資料

            /// 顯示所有Sheets
            ShowSheetsName(sheets);

            /// 顯示讀取Data
            Console.WriteLine("========= 第{0}個sheet : {1} 的Data  ========= ", ii + 1, sheets[ii]);
            ShowData(data);

            /// 顯示所有Sheet資料
            ShowAllExcelData(AllSheetsData);

        }


        private static void ShowSheetsName(List<string> sheets)
        {

            /// 顯示所有sheet名稱
            Console.WriteLine("========= Sheets  ========= ");
            foreach (string sh in sheets)
            {
                Console.WriteLine(sh);
            }
            Console.WriteLine("\n");
        }


        private static void ShowAllExcelData(Dictionary<string, string[,]> AllSheetsData)
        {
            /// 顯示所有讀取的Data
            Console.WriteLine("========= 顯示所有sheet的Data  ========= ");
            foreach (KeyValuePair<string, string[,]> DATA in AllSheetsData)
            {
                Console.WriteLine("========= sheet {0} 的 Data  ========= ", DATA.Key);
                ShowData(DATA.Value);
            }
        }



        /// <summary>
        /// 顯示ExcelData
        /// </summary>
        /// <param name="data"></param>
        private static void ShowData(string[,] data)
        {
            string ss = "";
            for (int i = 0; i < data.GetLength(0); i++)
            {
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    ss += data[i, j] + " ";
                }
                ss += "\r\n";
            }
            Console.WriteLine(ss);

        }
    
         
    }
}
