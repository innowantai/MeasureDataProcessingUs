using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

/*
 程式流程:
 1.載入SUM檔
 2.載入Excel sheet-4 第二欄位(judge)是否為空白,判斷是否,新增#EDIT or 不動
 3.若非空白,則儲存所對應第三欄位之Path --> targetString
 4.搜尋一次SUM檔案找出符合targetString 之位置並於poCheck變數中標注
 5.targetString 可能為全路徑也可能為 \xxx.ooo 最後一個路徑
 */

namespace Assembled
{
    public static class GPS_SUM
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";
        public static string addString = "$EDIT";


        public static string GPSSUM_main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";

            //// Load SUM and Excel data
            fileName = fileName.Replace(".SUM", "");
            ArrayList sumData = LoadSUM(dataPath, fileName + ".SUM");
            string[,] excelData = ExcelClass.ExcelSaveAndRead.Read(Path.Combine(dataPath, fileName + ".xlsx"), 2, 1, 4);

            //// Find TargetString
            ArrayList addTarget = getTarget(excelData);

            //// If judge all empty , return
            if (addTarget.Count == 0) { return "不需增加#EDIT\r\n"; }

            //// Mark need add row position
            int[] poCheck = poMark(sumData, addTarget);

            //// Save SUM file with adding #EDIT
            int addNum = saveSum(savePath, fileName, sumData, poCheck);

            outPutTxt += "資料處理完成,總共於" + addNum + "處新增#EDIT\r\n";
            return outPutTxt;
        }


        /// <summary>
        /// 儲存並複蓋SUM檔,根據poCheck = 1的位置之隔一行新增#EDIT
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        /// <param name="sumData"></param>
        /// <param name="poCheck"></param>
        /// <returns></returns>
        public static int saveSum(string savePath, string fileName, ArrayList sumData, int[] poCheck)
        {
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, fileName + ".SUM"));
            int kk = 0;
            for (int i = 0; i < sumData.Count; i++)
            {
                string index = Convert.ToString(sumData[i]);
                if (poCheck[i] == 0)
                {
                    sw.WriteLine(index);
                    sw.Flush();
                }
                else
                {
                    sw.WriteLine(index);
                    sw.Flush();
                    sw.WriteLine(addString);
                    sw.Flush();
                    kk++;
                }
            }
            sw.Close();
            return kk;
        }


        /// <summary>
        /// 搜尋一次sumData找到符合targetString之位置，於其iter後2位做標注，用來判斷是否新增#EDIT
        /// </summary>
        /// <param name="sumData"></param>
        /// <param name="addTarget"></param>
        /// <returns></returns>
        public static int[] poMark(ArrayList sumData, ArrayList addTarget)
        {
            ArrayList po = new ArrayList();
            int[] poCheck = new int[sumData.Count + 2];
            foreach (string targetString in addTarget)
            {
                string[] index2 = targetString.Split('\\');     //// Split path by symbol "\" and saving in the Array-variable of index2
                string index3 = index2[index2.Length - 1];      //// Take partial path that in the end position of Array-variable of index2
                for (int i = 0; i < sumData.Count; i++)
                {
                    string index = Convert.ToString(sumData[i]);
                    if (index.Contains(targetString) ? true : false | index.Contains(index3) ? true : false)   //// Two Case : Full and partial path
                    {
                        po.Add(i + 2);
                    }
                }
            }

            foreach (int ii in po)
            {
                poCheck[ii] = 1;
            }

            return poCheck;
        }


        /// <summary>
        /// 讀取excel judge 欄位不為空白之path
        /// </summary>
        /// <param name="excelData"></param>
        /// <returns></returns>
        public static ArrayList getTarget(string[,] excelData)
        {
            ArrayList addTarget = new ArrayList();
            for (int i = 0; i < excelData.GetLength(0); i++)
            {
                string index = excelData[i, 1];
                if (index.Trim() != "")
                {
                    addTarget.Add(excelData[i, 3]);
                }
            }
            return addTarget;
        }

        /// <summary>
        /// 讀取SUM檔並且拿掉$EDIT之列數
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static ArrayList LoadSUM(string savePath, string fileName)
        {
            StreamReader sr = new StreamReader(Path.Combine(savePath, fileName));
            ArrayList data = new ArrayList();
            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine();
                if (!index.Contains(addString))
                {
                    data.Add(index);
                }
            }
            sr.Close();
            return data;
        }
    }
}
