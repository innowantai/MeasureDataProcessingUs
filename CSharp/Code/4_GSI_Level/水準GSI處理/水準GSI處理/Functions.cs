using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 水準GSI處理
{
    public static class Functions
    {

        public static class DataProcess
        {


            public static Dictionary<string, string> CrackGSIData(string data)
            {
                string[] tmpData = data.Split(' ');
                string Iter = tmpData[0].Split('+')[0].Replace("*", "");
                string Name = tmpData[0].Split('+')[1];
                string D = tmpData[1].Split('+')[1];
                string H = tmpData[2].Split('+')[1];
                Dictionary<string, string> res = new Dictionary<string, string>();


                TransToMeter(D);
                res["D"] = TransToMeter(D);
                res["H"] = TransToMeter(H);
                res["Iter"] = Iter.Substring(2, Iter.Length - 2) ;
                res["Name"] = TakeOffZeros(Name);


                return res;
            }


            private static string TransToMeter(string Data)
            {
                int PointAtLocation = 5;
                string subD1 = Data.Substring(0, Data.Length - PointAtLocation);
                string subD2 = Data.Substring(Data.Length - PointAtLocation, PointAtLocation);
                string res = subD1 + "." + subD2;
                return res;

            }

            public static List<List<string>> LoadingAndClassGSIData(string path)
            {
                /// 讀取全部檔案
                List<string> tmpData = new List<string>();
                using (StreamReader sr = new StreamReader(path))
                {
                    while (sr.Peek() != -1)
                    {
                        tmpData.Add(sr.ReadLine());
                    }
                }

                /// 檔案分類
                List<string> tmp = new List<string>();
                List<List<string>> GroupDatas = new List<List<string>>();
                for (int i = 0; i < tmpData.Count; i++)
                {
                    if (tmpData[i].Length == 25)
                    {
                        if (tmp.Count != 0) GroupDatas.Add(tmp);
                        tmp = new List<string>();
                    }
                    else if (tmpData[i].Length == 145)
                    {
                        tmp.Add(tmpData[i - 4]);
                        tmp.Add(tmpData[i - 3]);
                        tmp.Add(tmpData[i - 2]);
                        tmp.Add(tmpData[i - 1]);
                    }
                }
                GroupDatas.Add(tmp);

                return GroupDatas;
            }


            private static string TakeOffZeros(string data)
            {

                string NewName = "";
                for (int i = 0; i < data.Length; i++)
                {
                    if (data.Substring(i, 1) != "0")
                    {
                        NewName = data.Substring(i, data.Length - i);
                        break;
                    }
                }

                return NewName;

            }
        }
    }
}
