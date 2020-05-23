using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace _3_0_Consloe_RINEX_TOOLS_1
{
    public class RINEX
    {
        List<ListData> Datas;
        public RINEX(string ListFilePath, string diro)
        {
            this.Datas = Function.LoadingData.LoadingData_List_Data(ListFilePath, diro);
            this.Datas[0].DataProcessing();
        }







    }
}
