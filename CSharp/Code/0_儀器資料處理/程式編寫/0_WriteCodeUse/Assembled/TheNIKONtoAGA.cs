using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

/*
 程式概要 2018-06-12 16:30
1.將NIKON之txt檔轉換至AGA格式,輸出為DAT檔案
2.因檔案副檔名與ZTSNtoAGA一樣為txt檔，需使用判斷格式準則
3.只讀取ST與SS之列數
4.以ST-SS-SS 三個為一組,判斷組數是否有誤
5.ST擷取 1 5 3 7 位置之資料
6.SS擷取 1 7 2 4 5 3 位置之資料
7.可能還需增加判斷組數有誤之程式
 */

namespace Assembled
{
    public static class TheNIKONtoAGA
    {


        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";

        static void Main_(string[] args)
        {
            string fileName = "RAW-0000.txt";
            string dataPath = oriPath;
            string savePath = oriPath;
            string res;

            res = TheNIKONtoAGA_Main(dataPath, savePath, fileName);

            Console.WriteLine(res);
        }

        public static string TheNIKONtoAGA_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";
            ArrayList Data = new ArrayList();
            string dateName = "";
            string stationName = "";
            int[] checkNumber = new int[2];
            bool errCheck1 = Load(dataPath, fileName, ref Data, ref dateName, ref stationName, ref checkNumber);
            if (errCheck1)
            {
                outPutTxt += "不是NIKON格式\r\n";
                return outPutTxt;
            }

            bool errCheck2 = checkNumber[0] == Data.Count / 3 && checkNumber[0] + checkNumber[1] == Data.Count ? false : true;
            if (errCheck2)
            {
                outPutTxt += "組數有誤\r\n";
                return outPutTxt;
            }

            ArrayList resData = catchData(Data, stationName, dateName);
            saveToDAT(resData, savePath, fileName);

            outPutTxt += "處理完成\r\n";
            return outPutTxt;
        }

        /// <summary>
        /// 讀檔且判斷此txt格式是否為NIKONtoAGA，針對"ST","SS"之列數讀檔，並讀取Created列數之 DateTime and SatationName
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="fileName"></param>
        /// <param name="Data"></param>
        /// <param name="dateName"></param>
        /// <param name="stationName"></param>
        /// <param name="checkNumber"></param>
        public static bool Load(string dataPath, string fileName, ref ArrayList Data, ref string dateName, ref string stationName, ref int[] checkNumber)
        {
            bool err = false;
            StreamReader sr = new StreamReader(Path.Combine(dataPath, fileName));
            int st = 0;
            int ss = 0;
            int kk = 0;
            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine();
                if (kk == 0 && !index.Contains("CO"))
                {
                    err = true;
                    return err;
                }

                if (index.Contains("Created"))
                {
                    string[] tmp1 = index.Split(',');
                    string[] tmp2 = index.Split(' ');
                    stationName = tmp1[1].Substring(0, tmp1[1].IndexOf(" "));
                    dateName = tmp2[3];
                }
                else if (index.Contains("ST,") | index.Contains("SS,"))
                {
                    st = index.Contains("ST,") ? ++st : st;
                    ss = index.Contains("SS,") ? ++ss : ss;
                    Data.Add(index);
                }
                kk++;
            }
            sr.Close();
            checkNumber[0] = st;
            checkNumber[1] = ss;

            return err;
        }

        /// <summary>
        /// 將ST與SS列之資料使用","分隔
        /// ST 擷取第1 5 3 7位置並包含前兩個stationName and dataName
        /// SS 擷取第 1 7 2 4 5 3 位置
        /// </summary>
        /// <param name="Data"></param>
        /// <param name="stationName"></param>
        /// <param name="dateName"></param>
        /// <returns></returns>
        public static ArrayList catchData(ArrayList Data, string stationName, string dateName)
        {
            int kk = 0;
            string[] index = null;
            string resIndex = null;
            ArrayList resData = new ArrayList();
            int[] caseIter1 = new int[] { 1, 5, 3, 7 };
            int[] caseIter2 = new int[] { 1, 7, 2, 4, 5, 3 };
            string[] caseStr1 = new string[] { " 2=", " 3=", " 62=", " 21=" };
            string[] caseStr2 = new string[] { " 5=", " 4=", " 6=", " 7=", " 8=", " 9=" };
            foreach (string ff in Data)
            {
                resIndex = "";
                index = ff.Split(',');
                if (kk % 3 == 0)
                {
                    resIndex += "50=" + stationName + " 51=" + dateName;
                    int iter = 0;
                    foreach (int ii in caseIter1)
                    {
                        resIndex += caseStr1[iter] + index[ii].Trim();
                        iter++;
                    }
                }
                else
                {
                    int iter = 0;
                    foreach (int ii in caseIter2)
                    {
                        resIndex += caseStr2[iter] + index[ii].Trim();
                        iter++;
                    }
                }
                resData.Add(resIndex.Trim());
                // Console.WriteLine(resIndex.Trim());
                kk++;
            }

            return resData;
        }

        /// <summary>
        /// 儲存至DAT檔
        /// </summary>
        /// <param name="resData"></param>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        public static void saveToDAT(ArrayList resData, string savePath, string fileName)
        {
            fileName = fileName.Replace(".txt", ".DAT");
            fileName = fileName.Replace(".TXT", ".DAT");
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, fileName));
            foreach (string ff in resData)
            {
                sw.WriteLine(ff);
                sw.Flush();
            }

            sw.Close();
        }
    }
}
