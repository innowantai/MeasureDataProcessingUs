using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

/// <summary>
/// Parameters : 
/// 50 : 檔名
///  2 : 測站
///  3 : 測站儀器高
/// 62 : 後視點號
///  6 : 後視儀器高
/// 21 : 歸0角度
///  5 : 觀測方向的點(前視)
/// 程式流程 :
/// 1.讀檔
/// 2.將檔案轉成原始輸出格式transData
/// 3.將transData每列之資料分割，依照Parameters位置填入陣列中，儲存至ArrayData中
/// 4.將ArrayData轉換output格式，其中每行資料對齊，小數點也要對齊
/// </summary>
/// ;

namespace 儀器資料處理
{
    public class TheJOBtoT01
    {

        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";

        public static string TheJOBtoT01_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";

            string[] data = LoadingData_sub(Path.Combine(dataPath, fileName));
            string[] transData = transData_sub(data);
            string[,] ArrayData = TransToArrayForm_sub(transData);
            List<string> resData = FormateToOutput_sub(ArrayData);

            SavingData_sub(fileName, savePath, resData);

            outPutTxt += "處理完成\r\n";
            return outPutTxt;


        }


        /// <summary>
        /// 讀檔
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private static string[] LoadingData_sub(string path)
        {
            List<string> data = new List<string>();
            StreamReader sr = new StreamReader(path);
            while (sr.Peek() != -1)
            {
                data.Add(sr.ReadLine());
            }
            return data.ToArray();
        }

        /// <summary>
        /// 將讀檔之資料轉換成output格式(2=, 5=開頭)
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private static string[] transData_sub(string[] data)
        {
            List<string> resData = new List<string>();
            string connectData = "";
            for (int i = 0; i < data.GetLength(0); i++)
            {
                string[] cmp = data[i].Split('=');
                if (cmp[0].IndexOf("54") != -1) { continue; }

                if (cmp[0].Length == 1 && (data[i].IndexOf("2=") != -1 || data[i].IndexOf("5=") != -1))
                {
                    resData.Add(connectData);
                    connectData = "";
                }
                connectData += data[i] + " ";
            }
            return resData.ToArray();
        }

        /// <summary>
        /// 將output格式資料，每列資料之塞入陣列所對應代號位置中中，63=xxxxx -> 將 xxxxx 塞入陣列63的位置
        /// </summary>
        /// <param name="transData"></param>
        /// <returns></returns>
        private static string[,] TransToArrayForm_sub(string[] transData)
        {
            int kk = 0;
            string[,] ArrayData = new string[transData.GetLength(0), 63];
            foreach (var subData in transData)
            {
                string[] subString = subData.Split(' ');
                foreach (string ss in subString)
                {
                    if (ss == "") { continue; }
                    string[] ssSplit = ss.Split('=');
                    int po = Convert.ToInt32(ssSplit[0]);
                    ArrayData[kk, po] = ssSplit[1];
                    // Console.Write(ssSplit[1] + " ");
                }

                // Console.WriteLine("");
                kk++;
            }
            return ArrayData;
        }

        /// <summary>
        /// 將資料轉成輸出格式，其中每行資料對齊，小數點也要對齊
        /// </summary>
        /// <param name="ArrayData"></param>
        /// <returns></returns>
        private static List<string> FormateToOutput_sub(string[,] ArrayData)
        {
            string[] block = CreateBlockArray_sub();
            int[] target50 = new int[] { 50 };
            int[] target2 = new int[] { 2, 3, 62, 6, 21 };
            int[] target5 = new int[] { 5, 6, 21, 7, 8, 9, 17, 18, 16, 19 };

            List<string> resData = new List<string>();
            string res = "";
            for (int i = 0; i < ArrayData.GetLength(0); i++)
            {
                /// 判斷此列資料為何種格式，將其格式所需截取之資料陣列儲存至target中
                int[] target = ArrayData[i, 50] != null ? target50 : (ArrayData[i, 2] != null ? target2 : (ArrayData[i, 5] != null ? target5 : null));

                /// 若所需截取之資列皆不存在，跳過此回合
                if (target == null) continue;


                res = "";
                int kk = 0;
                foreach (int po in target)
                {
                    /// 5=xxxxx之列數第一次會歸零所以會有21=xxxxxx，第二次則不會有21=xxxxx，所以會有空値狀況，給予""字串以免出錯
                    string tmpString = ArrayData[i, po] == null ? "" : ArrayData[i, po];
                    string headString = ArrayData[i, po] == null ? "" : po.ToString() + "=";
                    /// 字串資料長度，用來判斷要補多少空白値
                    int Len = tmpString.Length;

                    /// 要將資料對齊，設定要補滿之空格位數，kk代表每行資料的參數，po代表指定位置之參數
                    /// 若資料為上述所說之21=xxxxx空値，則補予10+3("21="三個字元拿掉)格空白
                    int targetLen = ArrayData[i, po] == null ? 13
                        : (kk == 1) ? 7
                        : (po == 16) ? 8
                        : (kk == 5 | kk == 3) ? 9 : 10;

                    /// 第3、4、5、6行之資料因要對齊小數點位置以小數點前面3位(xxx.oooooo)為基準來判斷要補多少空白
                    int tmp = (kk == 3 || kk == 4 | kk == 5 | kk == 6) ? tmp = 3 - tmpString.Split('.')[0].Length : 0;

                    /// 於第4行時，位置為8僅有1位數，其於位置為21，位了將21與8對齊，則在8前面補一個空白
                    headString = (kk == 4 && po == 8) ? "  " + headString : headString;

                    //// 於第3行時，位置為6其小數點後僅有3位數字，其餘為4位，為對齊所以於最後補1個0
                    tmpString = (kk == 3 && po == 6) ? tmpString + "0" : tmpString;

                    /// 將資料組合
                    res += headString + block[tmp] + tmpString + block[targetLen - Len - tmp];
                    kk++;
                }

                resData.Add(res);
            }

            return resData;
        }

        /// <summary>
        /// 儲存.T01檔案
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="savePath"></param>
        /// <param name="resData"></param>
        private static void SavingData_sub(string fileName, string savePath, List<string> resData)
        {
            string saveName = fileName.Substring(0, fileName.IndexOf(".")) + ".T01";
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, saveName));

            foreach (var ff in resData)
            {
                sw.WriteLine(ff);
                sw.Flush();
            }

            sw.Close();
        }

        /// <summary>
        /// 用來創造空白字串
        /// </summary>
        /// <returns></returns>
        private static string[] CreateBlockArray_sub()
        {
            string[] res = new string[100];
            string tmp = "";
            for (int i = 0; i < 100; i++)
            {
                res[i] = tmp;
                tmp += " ";
            }
            return res;
        }
    }
}
