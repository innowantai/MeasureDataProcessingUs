using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace _3_0_Consloe_RINEX_TOOLS_1
{
    public class ListData
    {
        private int ROUND_NUMBER = 3;
        public string FileName;
        public string MarkerName;
        public string NUMBER;
        public string REC_TYPE_VERs;
        public string ANT_TYPE;
        public string ANTENNA_DETAIL_HEN;
        public double R;
        public double C;
        public double A;
        public double RAW_ANATANNA_HEIGHT;
        public double APC;
        public double ARP;
        public string MARK;
        public string filePath;
        private List<string> O_Datas;
        public ListData(string data, string diro)
        {
            string[] tmp = data.Split(',');
            this.FileName = tmp[0];
            this.MarkerName = tmp[1];
            this.NUMBER = tmp[2];
            this.REC_TYPE_VERs = tmp[3];
            this.ANT_TYPE = tmp[4];
            this.ANTENNA_DETAIL_HEN = tmp[5];
            this.R = Convert.ToDouble(tmp[6]);
            this.C = Convert.ToDouble(tmp[7]);
            this.A = Convert.ToDouble(tmp[8]);
            this.RAW_ANATANNA_HEIGHT = Convert.ToDouble(tmp[9]);
            this.RAW_ANATANNA_HEIGHT = 1.485;
            this.MARK = tmp[12];
            this.APC = Math.Sqrt(RAW_ANATANNA_HEIGHT * RAW_ANATANNA_HEIGHT - this.R * this.R) + this.C;
            this.ARP = this.APC - this.A;
            this.filePath = Path.Combine(diro, this.FileName + ".20o");

            this.O_Datas = Function.LoadingData.Loading_O_Data(this.filePath);
        }



        public void DataProcessing()
        {
            /// Replace MARKER NAME
            string r1 = this.ReplaceMarkerName(out int num1);
            /// Replace MARKER NUMBER
            string r2 = this.ReplaceMarkerNumber(out int num2);

            /// Replace ANTENNA: DELTA H/E/N
            string r3 = this.ReplaceRECTYPE(out int num3);
            string r4 = this.ReplaceANTNNA_HEN(out int num4);

            this.O_Datas[num1] = r1;
            this.O_Datas[num2] = r2;
            this.O_Datas[num3] = r3;
            this.O_Datas[num4] = r4;
            string[] tmp = r4.Split(' ');
            this.ANTENNA_DETAIL_HEN = tmp[8] + "/" + tmp[16] + "/" + tmp[24];
        } 

        private string ReplaceANTNNA_HEN(out int num)
        {
            num = -1;
            double tmpData =this.MARK.Trim().ToLower() == "apc" ? this.APC : this.ARP;
            bool IsMinus = tmpData < 0;
            string tmp = tmpData.ToString();
            tmp = tmp.Replace("-", "");

            char[] replaceData = tmp.ToCharArray();
            int loopNum = 6;

            for (int i = 0; i < this.O_Datas.Count; i++)
            {
                string t = this.O_Datas[i];
                if (!t.Contains("COMMENT") && t.Contains("ANTENNA: DELTA H/E/N"))
                {
                    num = i;
                }
            }

            string data = this.O_Datas.Find(t => !t.Contains("COMMENT") && t.Contains("ANTENNA: DELTA H/E/N"));
            char[] dd = data.ToCharArray();
            for (int i = 8; i < loopNum + 8; i++)
            {
                dd[i] = replaceData[i - 8];
            }
            dd[7] = IsMinus ? '-' : ' ';

            return this.charArrayToString(dd);
        }

        private string ReplaceMarkerName(out int num)
        {
            num = -1;
            char[] replaceData = this.MarkerName.ToCharArray();
            int loopNum = replaceData.Count() < 4 ? replaceData.Count() : 4;

            for (int i = 0; i < this.O_Datas.Count; i++)
            {
                string t = this.O_Datas[i];
                if (!t.Contains("COMMENT") && t.Contains("MARKER NAME"))
                {
                    num = i;
                }
            }

            string data = this.O_Datas.Find(t => !t.Contains("COMMENT") && t.Contains("MARKER NAME"));
            char[] dd = data.ToCharArray();
            for (int i = 0; i < loopNum; i++)
            {
                dd[i] = replaceData[i];
            }
            string res = this.charArrayToString(dd);
            return res;
        }

        private string ReplaceMarkerNumber(out int num)
        {
            num = -1;
            char[] replaceData = this.NUMBER.ToCharArray();
            int loopNum = replaceData.Count() < 4 ? replaceData.Count() : 4;

            for (int i = 0; i < this.O_Datas.Count; i++)
            {
                string t = this.O_Datas[i];
                if (!t.Contains("COMMENT") && t.Contains("MARKER NUMBER"))
                {
                    num = i;
                }
            }

            string data = this.O_Datas.Find(t => !t.Contains("COMMENT") && t.Contains("MARKER NUMBER"));
            char[] dd = data.ToCharArray();
            for (int i = 0; i < loopNum; i++)
            {
                dd[i] = replaceData[i];
            }
            string res = this.charArrayToString(dd);
            return res;
        }


        private string ReplaceRECTYPE(out int num)
        {
            num = -1;
            string tmp = this.REC_TYPE_VERs.Split(':')[1].Split('/')[0];
            char[] replaceData = tmp.ToCharArray();
            int loopNum = replaceData.Count() < 10 ? replaceData.Count() : 10;
            for (int i = 0; i < this.O_Datas.Count; i++)
            {
                string t = this.O_Datas[i];
                if (!t.Contains("COMMENT") && t.Contains("REC # / TYPE / VERS"))
                {
                    num = i;
                }
            }
            string data = this.O_Datas.Find(t => !t.Contains("COMMENT") && t.Contains("REC # / TYPE / VERS"));
            char[] dd = data.ToCharArray();

            for (int i = 0; i < loopNum; i++)
            {
                dd[i] = replaceData[i];
            }

            string res = this.charArrayToString(dd);
            return res;
        }


        public void SavOData(string newPath)
        {
            using (StreamWriter sw = new StreamWriter(newPath))
            {
                foreach (string str in this.O_Datas)
                {
                    sw.WriteLine(str);
                    sw.Flush();
                }
            }
        }
        private string charArrayToString(char[] data)
        {
            string res = "";
            foreach (char item in data)
            {
                res += item;
            }
            return res;

        }
    }
}
