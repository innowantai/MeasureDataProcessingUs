using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace 水準GSI處理
{
    public class GroupData
    {
        List<LocationData> FrontData;
        List<LocationData> BackData;
        public int GroupNumber = 0;
        public List<string> Data;
        string StartLine = "For M5|Adr     2|TO  Start-Line         BF     X|                      |                      |                      | ";
        string EndLine = "For M5|Adr    33|TO  End-Line                  X|                      |                      |                      | ";
        public GroupData(List<string> Data)
        {
            string GN = String.Format("{0,5}", this.GroupNumber.ToString()); ;
            this.Data = Data;
            this.StartLine = this.StartLine.Replace("    X", GN);
            this.EndLine = this.StartLine.Replace("    X", GN);
        }


        public void Process()
        {
            this.FrontData = new List<LocationData>();
            for (int i = 0; i < this.Data.Count; i = i + 4)
            {
                this.FrontData.Add(new LocationData(this.Data[i],"Rb",GroupNumber));
                this.FrontData.Add(new LocationData(this.Data[i + 1], "Rf", GroupNumber));
            }

            this.BackData = new List<LocationData>();
            Dictionary<int, LocationData> BackData_Dic = new Dictionary<int, LocationData>();
            int kk = 0;
            for (int i = 2; i < this.Data.Count; i = i + 4)
            {
                BackData_Dic[kk] = new LocationData(this.Data[i + 1], "Rb", GroupNumber);
                BackData_Dic[kk + 1] = new LocationData(this.Data[i], "Rf", GroupNumber);
                kk = kk + 2;
            }

            foreach (int item in BackData_Dic.Keys.OrderByDescending(t => t))
            {
                this.BackData.Add(BackData_Dic[item]);
            } 
        }

         

    }
}
