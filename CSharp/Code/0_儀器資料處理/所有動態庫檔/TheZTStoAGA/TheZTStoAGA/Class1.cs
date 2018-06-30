using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Collections;

/*
 程式概要 2018-06-12 16:30
1.將ZTS之txt檔轉換至AGA格式,輸出為DAT檔案
2.因檔案副檔名與KINONtoAGA一樣為txt檔，需使用判斷格式準則
3.主要擷取文字檔中 STA {2=.3=}= {po[1],po[2]}, BS {62=} = {po[1]}, SD{5=,6=,4=} = {po[1],po[3],po[2]}, HVD {7=,8=,9=} = {po[1],po[2],po[3]}的資料
4.HVD中H與V單位為秒,要轉換至"度分秒格式"
5.STA BS SD HVD 擷取時須注意與小心測站名子是否有相同類似，可能會誤判
 */

namespace TheZTStoAGA
{
    public class TheZTStoAGA
    {
        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";

        static void Main(string[] args)
        {
            string fileName = "0813.txt";
            string dataPath = oriPath;
            string savePath = oriPath;
            string res;

            res = TheZTStoAGA_Main(dataPath, savePath, fileName);

            Console.WriteLine(res);
        }


        /// <summary>
        /// 流程 :
        /// 1.讀檔
        /// 2.擷取資料
        /// 3.儲存DAT
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="savePath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static string TheZTStoAGA_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";
            bool errCheck = false;
            ArrayList data = Read(dataPath, fileName);
            errCheck = Convert.ToString((System.String)data[0]).Contains("JOB") ? false : true;
            if (errCheck)
            {
                return outPutTxt += "格式有誤\r\n";
            }

            ArrayList resData = catchData(data, fileName);
            saveToDAT(savePath, fileName, resData);

            outPutTxt += "處理完成\r\n";
            return outPutTxt;
        }


        /// <summary>
        /// 很單純->讀全部檔案
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static ArrayList Read(string dataPath, string fileName)
        {
            StreamReader sr = new StreamReader(Path.Combine(dataPath, fileName));
            ArrayList data = new ArrayList();
            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine().Trim();
                data.Add(index);
            }
            sr.Close();
            return data;
        }

        /// <summary>
        /// 1.擷取所需資料 : STA, BS, SD, HVD,儲存至陣列中
        /// 2.將陣列資料轉換成字串並忽略空白列數
        /// </summary>
        /// <param name="data"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        public static ArrayList catchData(ArrayList data, string fileName)
        {
            string[,] dataSet = new string[data.Count, 6];
            int kk = 0;
            bool open = false;
            string nowTime = DateTime.Now.ToString("yyyy-MM-dd");   //使用資料處理當時的時間當作時間格式
            foreach (string ff in data)
            {
                //// If ff (another said the row of data) is block, Ignore it or take off head text 
                string index1 = ff.IndexOf(" ") == -1 ? ff : index1 = ff.Substring(ff.IndexOf(" "), ff.Length - ff.IndexOf(" ")).Trim();
                string[] index2 = index1.Split(',');

                if (ff.Contains("STA  "))
                {
                    dataSet[kk, 0] = "50=" + fileName.Replace(".txt", "");
                    dataSet[kk, 1] = "51=" + nowTime;
                    dataSet[kk, 2] = "2=" + index2[0];
                    dataSet[kk, 3] = "3=" + index2[1];
                }
                else if (ff.Contains("BS  "))
                {
                    dataSet[kk, 4] = "62=" + index2[0];
                    dataSet[kk, 5] = "21=0.0000";
                    kk++;
                }
                else if (ff.Contains("SD  "))
                {
                    dataSet[kk, 0] = "5=" + index2[0];
                    dataSet[kk, 1] = "4=" + index2[2];
                    dataSet[kk, 2] = "6=" + index2[1];
                    open = true;

                }
                else if (ff.Contains("HVD  ") && open)
                {
                    dataSet[kk, 3] = "7=" + transAngle(index2[0]);
                    dataSet[kk, 4] = "8=" + transAngle(index2[1]);
                    dataSet[kk, 5] = "9=" + index2[2];
                    kk++;
                    open = false;
                }
                else if (ff.Trim() == "")
                {
                    kk++;
                }
            }

            //// Translate Array to string output-formation
            ArrayList resData = new ArrayList();
            resData.Add("12345678901234567890123456789012345678901234567890123456789012345678901234567890");
            resData.Add("");
            resData.Add("");
            for (int i = 0; i < data.Count; i++)
            {
                string index = "";
                for (int j = 0; j < 6; j++)
                {
                    index += dataSet[i, j] + " ";
                }
                if (index.Trim() != "")
                {
                    resData.Add(index);
                }
            }

            return resData;
        }

        /// <summary>
        /// 很單純的將資料輸出DAT檔
        /// </summary>
        /// <param name="savePath"></param>
        /// <param name="saveName"></param>
        /// <param name="data"></param>
        public static void saveToDAT(string savePath, string saveName, ArrayList data)
        {
            saveName = saveName.Replace(".txt", ".DAT");
            saveName = saveName.Replace(".TXT", ".DAT");
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, saveName));
            foreach (string ff in data)
            {
                sw.WriteLine(ff);
                sw.Flush();
            }
            sw.Close();
        }

        /// <summary>
        /// 將"秒數"轉換成"度分秒"
        /// </summary>
        /// <param name="Data"></param>
        /// <returns></returns>
        public static string transAngle(string Data)
        {
            double data = Convert.ToInt32(Data);
            int dd, mm, ss;
            double index;
            string result;
            double test = data / 3600;

            dd = Convert.ToInt32(Math.Floor(data / 3600));
            index = data % 3600;
            mm = Convert.ToInt32(Math.Floor(index / 60));
            ss = Convert.ToInt32(index % 60);
            result = dd + "." + mm + ss;
            return result;
        }

    }
}
