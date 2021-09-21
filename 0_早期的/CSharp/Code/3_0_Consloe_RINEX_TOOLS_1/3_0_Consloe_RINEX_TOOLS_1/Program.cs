using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace _3_0_Consloe_RINEX_TOOLS_1
{
    class Program
    {
        static void Main(string[] args)
        {
            string BasePath = @"C:\Users\innow\桌面\test\RINEX_TOOLS_1";
            string ListData = Path.Combine(BasePath, "5700-01-LIST.csv");
            string diro = @"C:\Users\innow\桌面\test\RINEX_TOOLS_1\5700-1";
            RINEX rr = new RINEX(ListData, diro);


        }
    }
}
