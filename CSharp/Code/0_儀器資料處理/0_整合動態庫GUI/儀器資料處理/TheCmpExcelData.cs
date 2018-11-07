using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;
/*
 程式概要 2018-06-12 16:30
 1.讀取並比較兩excel表 ("檔名a.xls" V.S "檔名.xls" -> 讀取檔名相同只差一個a
 2.比較結果儲存至表格 : (1) "檔名a.xls"之結果放左邊
                        (2) "檔名.xls" 之結果放右邊
                        (3) 中間以judge隔開用來
                            (a) 判斷兩表格不同之"測站"
                            (b) 表格左邊有重複之測站位置
                        (4) 表格左邊與右邊相同測站對應位置擺放，不同處於中間judge標註1
 
 */

namespace 儀器資料處理
{
    public class TheCmpExcelData
    {

        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";
        public static string TheCmp_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";
            string fName1 = fileName.Replace("原始過程", "a原始過程");
            string fName2 = fileName;
            string[,] Data1 = null;
            try
            {
                Data1 = ExcelClass.ExcelSaveAndRead.Read(Path.Combine(dataPath, fName1), 2, 1, 11);
            }
            catch (Exception)
            {
                return "缺少 " + fName1 + " 檔案 \r\n";
            }

            string[,] Data2 = ExcelClass.ExcelSaveAndRead.Read(Path.Combine(dataPath, fName2), 2, 1, 11);

            int dataLen1 = Data1.GetLength(0);
            int dataLen2 = Data2.GetLength(0);
            int[,] inform1 = new int[dataLen1, 2];
            int[,] inform2 = new int[dataLen2, 2];

            int caseName = dataLen1 >= dataLen2 ? 1 : 2;
            string[,] resData_;
            string[,] resData;

            setPosition(Data1, Data2, ref inform1, ref inform2);
            resData_ = combineResData(Data1, Data2, inform1, inform2);
            resData = resDataAfterProcess(resData_);

            string indexSaveName = fileName.Replace("原始過程", "原始過程Result");
            File.Delete(Path.Combine(savePath, indexSaveName));
            indexSaveName = indexSaveName.Replace(".xls", ".xlsx");
            File.Copy(Path.Combine(oriPath, "Ori_excel.xlsx"), Path.Combine(savePath, indexSaveName));
            ExcelClass.ExcelSaveAndRead.Save(strPath: Path.Combine(savePath, indexSaveName), sheetNumber: 13, poRow: 2, poCol: 1, Data: resData);

            outPutTxt += "處理完成\r\n";
            return outPutTxt;
        }

        /// <summary>
        /// 流程 : 
        /// 1 : 比較兩個excelData, [L,R] = [a-part,non a-part]
        /// 2 : 以L-Data為準, 依序搜尋R-Data之測站名稱有相同與無相同之資訊 
        /// 3 : 找出無相同測站，給與位置，給予前一測站之位置之最後數值+1
        /// 4 : 依序將大於(3)所給予之無相同測站位置數值 + 1
        /// </summary>
        /// <param name="Data1"></param> L-Data
        /// <param name="Data2"></param> R-Data
        /// <param name="inform1"></param> L-Information[row,col] : row 判斷是否有相同,無相同為-1, col :位置
        /// <param name="inform2"></param> R-Information[row,col] : row 判斷是否有相同,無相同為-1, col :位置
        public static void setPosition(string[,] Data1, string[,] Data2, ref int[,] inform1, ref int[,] inform2)
        {
            int dataLen1 = Data1.GetLength(0);
            int dataLen2 = Data2.GetLength(0);
            for (int i = 0; i < dataLen2; i++) { inform2[i, 0] = -1; }
            for (int i = 0; i < dataLen1; i++) { inform1[i, 0] = -1; }

            //// According to the L-Data to compare the R-Data that have same station-Name
            //// same ? record position : -1
            bool flag = true;                           //// use to judge that isn't have same sataion-Name
            int KK = -1;                                //// use to record position
            string index1 = null, index2 = null;
            for (int i = 0; i < dataLen1; i++)
            {
                index1 = Data1[i, 0] + Data1[i, 1];
                flag = true;
                for (int j = 0; j < dataLen2; j++)
                {
                    index2 = Data2[j, 0] + Data2[j, 1];
                    if (index1 == index2 && inform2[j, 0] == -1 && index2.Trim() != "")
                    {
                        KK++;
                        inform2[j, 0] = i;
                        inform2[j, 1] = KK;
                        inform1[i, 0] = 1;
                        inform1[i, 1] = KK;
                        flag = false;
                    }
                }

                if (flag)
                {
                    KK++;
                    inform1[i, 1] = KK;
                }
            }

            //// Given non-same sataionName position by that before 1-step position plus 1
            int target;
            for (int i = 1; i < dataLen2; i++)
            {
                if (inform2[i, 0] == -1)
                {
                    //// Find the position according before 1-step is have positon or not? "use final postion by L-Data"  : "direct use position"
                    inform2[i, 1] = inform2[i - 1, 0] == -1 ? inform2[i - 1, 1] + 1 : inform1[inform2[i - 1, 0], 1] + 1;
                    target = inform2[i, 1];

                    for (int j = 0; j < dataLen1; j++)
                    {
                        inform1[j, 1] = inform1[j, 1] >= target ? inform1[j, 1] + 1 : inform1[j, 1];
                    }

                    for (int j = 0; j < dataLen2; j++)
                    {
                        inform2[j, 1] = (inform2[j, 0] != -1 && inform2[j, 1] >= target) ? inform2[j, 1] + 1 : inform2[j, 1];
                    }
                }
                //Console.WriteLine(i.ToString() + " : " + Data2[i, 0] + " , " + Data2[i, 1] + " , " + inform2[i, 0] + "," + inform2[i, 1]);
            }
        }

        /// <summary>
        /// 跟據inform1 與 inform2 所紀錄之位置，將Data1與Data2組合至resData裡面
        /// </summary>
        /// <param name="Data1"></param>
        /// <param name="Data2"></param>
        /// <param name="inform1"></param>
        /// <param name="inform2"></param>
        /// <returns></returns>
        public static string[,] combineResData(string[,] Data1, string[,] Data2, int[,] inform1, int[,] inform2)
        {
            int dataLen1 = Data1.GetLength(0);
            int dataLen2 = Data2.GetLength(0);
            int dataLen = dataLen1 >= dataLen2 ? dataLen1 : dataLen2;

            string[,] resData = new string[dataLen * 2, 21];
            for (int i = 0; i < dataLen1; i++)
            {
                int po = inform1[i, 1];
                if (po != -1)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        resData[po, j] = Data1[i, j];
                    }
                    if (inform1[i, 0] == -1)
                    {
                        resData[po, 10] = "1";
                    }
                }
            }

            for (int i = 0; i < dataLen2; i++)
            {
                int po = inform2[i, 1];
                if (po != -1)
                {
                    for (int j = 0; j < 10; j++)
                    {
                        resData[po, j + 11] = Data2[i, j];
                    }
                    if (inform2[i, 0] == -1)
                    {
                        resData[po, 10] = "1";
                    }
                }
            }
            return resData;
        }

        /// <summary>
        /// 將空白之row數拿掉，並比較L-Data是否有重複與 L-Data與R-Data是否有沒出現過之測站Case? "給予重複位置" : "1" at Judge column
        /// </summary>
        /// <param name="resData"></param>
        /// <returns></returns>
        public static string[,] resDataAfterProcess(string[,] resData)
        {
            //// Assgin the old resData to new ResData thar take off block row
            string[,] RESData = new string[resData.GetLength(0), resData.GetLength(1)];

            int kk = 0;
            bool open = true;
            for (int i = 0; i < resData.GetLength(0); i++)
            {
                open = true;
                string index = "";
                for (int j = 0; j < 21; j++) { index += resData[i, j]; }

                if (index.Trim() == "" | index.Trim() == "1") { open = false; }

                if (open)
                {
                    for (int p = 0; p < 21; p++)
                    {
                        RESData[kk, p] = resData[i, p];
                    }
                    kk++;
                }
            }

            //// Fine repeat stationName and position
            bool[] rec = new bool[RESData.GetLength(0)];
            for (int i = 0; i < RESData.GetLength(0); i++) { rec[i] = true; }
            for (int i = 0; i < RESData.GetLength(0); i++)
            {
                string index1 = RESData[i, 0] + RESData[i, 1];
                if (rec[i])
                {
                    for (int j = 0; j < RESData.GetLength(0); j++)
                    {
                        string index2 = RESData[j, 0] + RESData[j, 1];
                        string resOut = "與第" + (i + 2).ToString() + "列測站重複";
                        if (index1 == index2 && index1.Trim() != "" && i != j)
                        {
                            RESData[j, 10] = resOut;
                            rec[j] = false;
                        }
                    }
                }
            }

            return RESData;
        }
    }
}
