using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
using ExcelClass;

/*
 2018-06-14 14:15
 程式概要 :
 1.讀取GPS csv檔將其CODE轉換至 C1 C2 C3 DESP 格式後儲存至excel表格中
 2.儲存之excel有一定格式，除了原csv檔及C1,C2,C3,DESP,資料外其餘補1
 */
namespace 儀器資料處理
{
    public class GPS_SORT
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";


        public static string GPSSORT_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";

            EXCEL Excel = new EXCEL(Path.Combine(dataPath, fileName));
            //string[,] data = ExcelClass.ExcelSaveAndRead.Read(Path.Combine(dataPath, fileName), 1, 1, 1);
            string[,] data = Excel.GetDataBySheetName(1, 1, Excel.sheets[0]);
            string[,] excelData = TransToExcelFormat(data);

            string SaveName = "三次元資料.xlsx";
            string sheetName = fileName.Replace(".csv", "");
            sheetName = sheetName.Replace(".CSV", "");
            //ExcelClass.ExcelSaveAndRead.SaveCreat(strPath: Path.Combine(savePath, SaveName), sheetName: sheetName, poRow: 1, poCol: 1, Data: excelData);
            Excel.Save_To(Path.Combine(savePath, SaveName), sheetName, 1, 1, excelData);

            outPutTxt = "處理完成\r\n";
            Excel.close();
            return outPutTxt;
        }




        private static string[,] TransToExcelFormat(string[,] data)
        {
            string[,] excelData = new string[data.GetLength(0) + 1, 13];
            string[] headName = new string[] { " 點號 (測站)", " 覘標高", " 編碼 (後視)", "   水平角 ", "天頂距", " 斜距", "  縱座標", " 橫座標 ", "高程", "  C1 ", " C2 ", " C3 ", "DESP" };
            int[] equalOnePo = new int[] { 1, 3, 4, 5 };
            int[] catToExcelPo = new int[] { 0, 2, 6, 7, 8 };
            int[] oriDataPo = new int[] { 0, 4, 1, 2, 3 };
            for (int i = 1; i < data.GetLength(0) + 1; i++)
            {
                string[] index = data[i - 1, 4].Split('.');
                string index1 = index[0];
                string index2 = index.Length == 1 ? "" : index[1];

                //// Part : C1 C2 C3 
                double index1Len = index1.Length;
                double num1 = Math.Ceiling(index1Len / 3);
                for (int jj = 0; jj < num1; jj++)
                {
                    int endPo = 3 + 3 * jj > index1.Length ? index1.Length - 3 * jj : 3;
                    string tmpIndex = index1.Substring(0 + 3 * jj, endPo);
                    excelData[i, 9 + jj] = tmpIndex;
                }

                //// Part : DESP 
                index2 = index2.Length == 3 ? index2 : (index2.Length == 2 ? index2 + "0" : (index2.Length == 1 ? index2 + "00" : ""));
                try
                {
                    index2 = index2 != "" ? Convert.ToInt32(index2).ToString() : "";
                }
                catch (Exception)
                {

                }
                excelData[i, 12] = index2;

                //// Catch Else Data : PointName,Code,X,Y,H; 
                for (int ii = 0; ii < 5; ii++)
                {
                    excelData[i, catToExcelPo[ii]] = data[i - 1, oriDataPo[ii]];
                }

                //// Full One
                foreach (int ii in equalOnePo)
                {
                    excelData[i, ii] = "1";
                }
            }

            ////Given HeadName
            for (int i = 0; i < 13; i++)
            {
                excelData[0, i] = headName[i];
            }
            return excelData;
        }
    }
}
