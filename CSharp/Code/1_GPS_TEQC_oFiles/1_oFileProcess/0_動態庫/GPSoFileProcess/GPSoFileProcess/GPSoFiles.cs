using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;



/*
 2018-06-17 18:30
 程式概述 :
 1.讀取選取資料夾之excel檔其備註欄位為o檔檔名,其餘欄位為儀器資訊
 2.將資料夾所有o檔檔名與excel檔備註欄之o檔檔名批配,若有找到則將所需資訊填入o檔檔名中
 3.所需資訊為 :
    (a) 點號       : MARKER NAME && MARKER NUMBER
    (b) 接收器序號 : "REC # / TYPE / VERS"           ----> 用於判斷o檔所填寫之接收器訊號是否有與excel表中相同
    (c) 天線盤序號 : "ANT # / TYPE"
    (d) 天線高     :"ANTENNA: DELTA H/E/N"
 4.天線高之 E and N 資訊各於excel中建立sheet儲存資訊

 */
namespace GPSoFileProcess
{

    public class GPSoFiles
    {

        public static string outPutTxt = "";


        public static string GPSoFile_Main(string excelFullPath, string dataPath, string oriPath, string savePath)
        {
            //// 讀取excel檔並將無用資訊去除
            string[,] data = GetUseData(ExcelClass.ExcelSaveAndRead.Read(excelFullPath, 1, 1, 1));

            //// 將excel檔轉成陣列格式儲存至,將所對應檔名資料 儲存至 字典DicExcelData中
            string[] excelOFileName = new string[data.GetLength(0)];
            Dictionary<string, string> DicExcelData = new Dictionary<string, string>();
            for (int i = 0; i < data.GetLength(0); i++)
            {
                string resIndex = "";
                for (int j = 0; j < data.GetLength(1); j++)
                {
                    resIndex += data[i, j] + "|";
                } 

                excelOFileName[i] = data[i, 10];
                DicExcelData[data[i, 10]] = resIndex;
            }


            //// 找出資料夾所有o檔檔名
            Dictionary<string, string> DicOFiles = new Dictionary<string, string>();
            string[] OFiles = FindAlloFilesAndFileName(dataPath, ref DicOFiles);

            //// 儲存至excel天線高資訊
            string[,] ARPexcel = new string[data.GetLength(0), 3];
            string[,] APCexcel = new string[data.GetLength(0), 3];
            int kk = 0;
            int num = 0;
            outPutTxt += "  總共 " + OFiles.Count().ToString() + " 筆o檔\r\n";
            foreach (var ff in OFiles)
            {
                string index = ff.ToString();
                bool check = excelOFileName.Any(cc => ff.Contains(cc));
                if (check)
                {
                    index = index.Substring(0, index.IndexOf("."));
                    string[] ArrayData = DicExcelData[index].Split('|');
                    string ofileName = ArrayData[10];
                    outPutTxt += "  第" + (num + 1).ToString() + "筆(" + DicOFiles[ofileName] + ") : ";
                    outPutTxt += oFileProcessing(dataPath, DicOFiles[ofileName], ArrayData, ARPexcel, APCexcel, kk);
                    ARPexcel[kk, 0] = ofileName;
                    APCexcel[kk, 0] = ofileName;
                    kk++;
                }
                else
                {
                    outPutTxt += "  第" + (num + 1).ToString() + "筆(" + index + ") : Excel檔無資料\r\n";
                }
                num++;
            }

            ExcelClass.ExcelSaveAndRead.SaveCreat(excelFullPath, "ARP", 1, 1, ARPexcel);
            ExcelClass.ExcelSaveAndRead.SaveCreat(excelFullPath, "APC", 1, 1, APCexcel);

            outPutTxt += "  總共" + num.ToString() + "資料處理完畢";


            return outPutTxt;
        }






        /// <summary>
        /// 找出資料夾所有o檔與o檔檔名對應儲存至res與DicOFiles中,
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="DicOFiles"></param>
        /// <returns>res.split(',') , DicOFiles</returns>
        private static string[] FindAlloFilesAndFileName(string dataPath, ref Dictionary<string, string> DicOFiles)
        {
            DirectoryInfo Dir = new DirectoryInfo(dataPath);
            var Files = Dir.GetFiles();
            string res = "";
            int kk = 0;
            foreach (var ff in Files)
            {
                if (Path.GetExtension(ff.ToString()).Contains("o") | Path.GetExtension(ff.ToString()).Contains("O"))
                {
                    string subName = Path.GetExtension(ff.ToString());
                    string index = ff.ToString().Substring(0, ff.ToString().IndexOf(subName));
                    DicOFiles[index] = ff.ToString();
                    res += ff.ToString() + ",";
                }
                kk++;
            }
            res = res.Substring(0, res.Length - 1);
            return res.Split(',');
        }


        /// <summary>
        /// 替換txt檔中資訊
        /// </summary>
        /// <param name="dataPath"></param>
        /// <param name="fileName"></param>
        /// <param name="paraData"></param>
        /// <param name="ARPexcel"></param>
        /// <param name="APCexcel"></param>
        /// <param name="lastPO"></param>
        /// <returns></returns>
        private static string oFileProcessing(string dataPath, string fileName, string[] paraData, string[,] ARPexcel, string[,] APCexcel, int lastPO)
        {
            string res = "";
            string[] tarString = { "MARKER NAME", "MARKER NUMBER", "REC # / TYPE / VERS", "ANT # / TYPE", "ANTENNA: DELTA H/E/N" };
            string[] block = creatTypeArray(" ");
            string[] zeros = creatTypeArray("0");
             

            string fullPath = Path.Combine(dataPath, fileName);

            //// 儲存targetString 對應至data-array中之列數
            Dictionary<string, int> tarDic = new Dictionary<string, int>();
            string[] oData = Read(fullPath, tarString, ref tarDic);

            //// tarString 有五項,若字典裡面沒有五個的話表示有缺資料
            if (tarDic.Count() != 5) { return "  o檔有缺資料\r\n"; }

            int kk = 0;
            string tmp1 = "";
            string tmp2 = "";
            bool checkImpNo = true;
            foreach (string ff in tarString)
            {
                int po = tarDic[ff];
                string index = oData[po];

                if (kk == 0 | kk == 1)
                {
                    tmp1 = paraData[1];
                    tmp2 = index.Substring(0, tmp1.Length);
                    index = index.Replace(tmp2, tmp1);
                    ARPexcel[lastPO, 1] = tmp1;
                    APCexcel[lastPO, 1] = tmp1;
                }
                else if (kk == 2)
                {
                    tmp1 = index.Split(' ')[0];
                    checkImpNo = tmp1 == paraData[4];
                }
                else if (kk == 3)
                {
                    tmp1 = paraData[6];
                    index = block[20] + tmp1 + block[40 - tmp1.Length] + ff;
                }
                else if (kk == 4)
                {
                    index = "";
                    for (int i = 0; i < 3; i++)
                    {
                        tmp1 = Math.Round(Convert.ToDouble(paraData[7 + i]), 4).ToString();
                        tmp2 = tmp1.Split('.')[1];
                        tmp1 = tmp2.Length == 4 ? tmp1 : tmp1 + zeros[4 - tmp2.Length];
                        index += block[8] + tmp1;
                        if (i == 1)
                        {
                            ARPexcel[lastPO, 2] = tmp1;
                        }
                        else if (i == 2)
                        {
                            APCexcel[lastPO, 2] = tmp1;
                        }
                    }
                    index += "                  " + ff;
                }
                oData[po] = index;
                kk++;
            }
            Save(Path.Combine(dataPath, fileName), oData);

            res = checkImpNo ? "  處理完成\r\n" : "  處理完成 -> 儀器序號不匹配\r\n";
            return res;
        }
         

        /// <summary>
        /// 讀取o檔資料儲存於"陣列中"輸出並找出targetString於陣列之位置儲存至po中
        /// </summary>
        /// <param name="fullPath"></param>
        /// <param name="targetString"></param>
        /// <param name="po"></param>
        /// <returns>res , po </returns>
        private static string[] Read(string fullPath, string[] targetString, ref Dictionary<string, int> targetDic)
        {
            StreamReader sr = new StreamReader(fullPath);
            ArrayList Data = new ArrayList();
            int kk = 0; 
            while (sr.Peek() != -1)
            {
                string index = sr.ReadLine();
                Data.Add(index);

                foreach (string ff in targetString)
                {
                    if (index.Contains(ff))
                    {
                        targetDic[ff] = kk;
                        break;
                    }
                }
                kk++;
            }
            sr.Close();

            string[] res = new string[Data.Count];
            for (int i = 0; i < Data.Count; i++)
            {
                res[i] = (System.String)Data[i];
            }


            return res;
        }
        
        /// <summary>
        /// 很單純儲存新的o檔
        /// </summary>
        /// <param name="fullPath"></param>
        /// <param name="data"></param>
        private static void Save(string fullPath, string[] data)
        {
            StreamWriter sw = new StreamWriter(fullPath); 
            foreach (string ff in data)
            {
                sw.WriteLine(ff);
                sw.Flush();
            } 
            sw.Close(); 
        }



        /// <summary>
        /// 以excel第8行與第12行之資料判斷此列是否為可使用之資料,可使用則儲存
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static string[,] GetUseData(string[,] data)
        {

            int[] po = new int[data.GetLength(0)];
            for (int i = 0; i < data.GetLength(0); i++)
            {
                po[i] = Double.TryParse(data[i, 7], out double check) && !String.IsNullOrEmpty(data[i, 11]) ? 1 : 0;
            }

            int num = po.Count(n => n == 1);
            int kk = 0;
            string[,] Data = new string[num, 12];
            for (int i = 0; i < data.GetLength(0); i++)
            {
                if (po[i] == 1)
                {
                    for (int j = 0; j < 12; j++)
                    {
                        Data[kk, j] = data[i, j];
                    }
                    kk++;
                }
            }
            return Data;
        }


        private static string[] creatTypeArray(string creatType)
        {
            string[] block = new string[50];
            for (int i = 0; i < 50; i++)
            {
                for (int j = 0; j < i; j++)
                {
                    block[i] += creatType;
                }
            }
            return block;
        }
    }
}
