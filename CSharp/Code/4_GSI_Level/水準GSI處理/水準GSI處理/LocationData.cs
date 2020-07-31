using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mime;
using System.Text;
using System.Windows.Forms.VisualStyles;

namespace 水準GSI處理
{
    public class LocationData
    {

        public string Name;
        public string Iter;
        public double D;
        public double H;
        public string Data;
        public string Case;
        public string GroupNumber;
        public LocationData(string Data,string CA,int　GroupNumber)
        {
            this.GroupNumber = GroupNumber.ToString();
            this.Case = CA;
            this.Data = Data;
            Dictionary<string, string> TreansData = Functions.
                                                    DataProcess.
                                                    CrackGSIData(Data);
            this.Name = TreansData["Name"];
            this.Iter = Convert.ToInt32(TreansData["Iter"]).ToString();
            this.D = Convert.ToDouble(TreansData["D"]);
            this.H = Convert.ToDouble(TreansData["H"]);


            string R_Iter = this.AddBlock(Iter, 6) + "|KD1";
            string R_Name = this.AddBlock(Name, 9);
            string R_time = "      00:00:000   " + GroupNumber + "|" + Case;
            string R_H = " |HD" + this.AddBlock(this.H.ToString(), 15) + " m   ";
            string R_D = this.AddBlock(this.D.ToString(), 15) + " m   |";
            string FormularData = "For M5|Adr" + R_Iter + R_Name + R_time + R_D + R_H;

        }


        private string AddBlock(string str,int targetNum)
        {

            return String.Format("{0," + targetNum.ToString() + "}", str);
        }
    }
}
