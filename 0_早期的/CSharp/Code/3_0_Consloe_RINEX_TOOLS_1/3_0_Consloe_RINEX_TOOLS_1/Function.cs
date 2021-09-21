using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace _3_0_Consloe_RINEX_TOOLS_1
{
    public static class Function
    {
        public static class LoadingData
        {
            public static List<ListData> LoadingData_List_Data(string path, string diro)
            {
                List<ListData> data = new List<ListData>();
                using (StreamReader sr = new StreamReader(path, Encoding.Default, true))
                {
                    sr.ReadLine();
                    while (sr.Peek() != -1)
                    {
                        data.Add(new ListData(sr.ReadLine(), diro));
                    }
                }

                return data;
            }

            public static List<string> Loading_O_Data(string path)
            {
                List<string> data = new List<string>();
                using (StreamReader sr = new StreamReader(path))
                {
                    while (sr.Peek() != -1)
                    {
                        data.Add(sr.ReadLine());
                    }
                }
                return data;
            }
        }
    }
}
