using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace TheJOBtoT01
{
    public class TheJOBtoT01
    {

        public static string oriPath = System.Environment.CurrentDirectory;
        public static string outPutTxt = "";

        public static string TheJOBtoT01_Main(string dataPath, string savePath, string fileName)
        {
            outPutTxt = "";
            int[] target50 = new int[] { 50 };
            int[] target2 =new int[] { 2, 3, 62, 6, 21 };
            int[] target5 = new int[] { 5, 6, 21, 7, 8, 9, 17, 18, 16, 19 };
            string[] block = CreateBlockArray_sub();

            string[] data = LoadingData_sub(Path.Combine(dataPath,fileName));
            string[] transData = transData_sub(data);
            string[,] ArrayData = TransToArrayForm_sub(transData);

            List<string> resData = new List<string>();
            string res = "";
            for (int i = 0; i < ArrayData.GetLength(0); i++)
            { 
                int[] target = ArrayData[i, 50] != null ? target50 : (ArrayData[i, 2] != null ? target2 : (ArrayData[i, 5] != null ? target5 : null));
                if (target == null) continue;


                res = "";
                int kk = 0;
                foreach (int po in target)
                {
                    string tmpString  = ArrayData[i, po] == null ? "" : ArrayData[i, po]; 
                    string headString = ArrayData[i, po] == null ? "" : po.ToString() + "=";
                    int Len = tmpString.Length;
                    int targetLen = ArrayData[i, po] == null ? 13 
                        : (kk == 1)  ? 7 
                        : (po == 16) ? 8 
                        : (kk == 5 | kk == 3)  ? 9 : 10;

                    int tmp = (kk == 3 || kk == 4 | kk == 5 | kk == 6) ? tmp = 3 - tmpString.Split('.')[0].Length : 0;
                    headString = (kk == 4 && po == 8) ? "  " + headString : headString;
                    tmpString = (kk == 3 && po == 6) ? tmpString + "0" : tmpString;

                    res += headString  + block[tmp] + tmpString + block[targetLen - Len - tmp];
                    kk++;
                }

                resData.Add(res); 
            }
             


            string saveName = fileName.Substring(0,fileName.IndexOf(".")) + ".T01";
            StreamWriter sw = new StreamWriter(Path.Combine(savePath, saveName));

            foreach (var ff in resData)
            {
                sw.WriteLine(ff);
                sw.Flush();
            }

            sw.Close();
            outPutTxt += "處理完成\r\n";

            return outPutTxt;


        }

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

        private static string[] transData_sub(string[] data)
        {
            List<string> resData = new List<string>();
            string connectData = "";
            for (int i = 0; i < data.GetLength(0); i++)
            {
                string[] cmp = data[i].Split('=');
                if (cmp[0].IndexOf("54") != -1)
                {
                    continue;
                }


                if (cmp[0].Length == 1 && (data[i].IndexOf("2=") != -1 || data[i].IndexOf("5=") != -1))
                {
                    resData.Add(connectData);
                    connectData = "";
                }
                connectData += data[i] + " ";
            }

            return resData.ToArray();
        }


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
