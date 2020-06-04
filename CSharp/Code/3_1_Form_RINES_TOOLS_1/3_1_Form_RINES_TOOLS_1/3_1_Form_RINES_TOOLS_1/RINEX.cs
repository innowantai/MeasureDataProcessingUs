using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;

namespace _3_1_Form_RINES_TOOLS_1
{
    public class RINEX
    {
        List<ListData> Datas;
        public RINEX(string ListFilePath, string diro)
        {
            List<string> csvRes = new List<string>();
            csvRes.Add("FILE NAME,MARKER NAME,MARKER MUMBER,REC # /TYPE/VERS,ANT #/TYPE,ANTENNA:DELTA H/E/N,R,C,A,RAW ANTANNA HEIGHT,垂高(至相位中心APC),垂高(至天線盤底ARP),MARK");
            this.Datas = Function.LoadingData.LoadingData_List_Data(ListFilePath, diro);
            foreach (ListData item in this.Datas)
            {
                item.DataProcessing();
                string FullfilaName = Path.GetFileName(item.filePath);
                string DirPath = Path.GetDirectoryName(item.filePath);
                string FileName = FullfilaName.Split('.')[0];
                string subName = FullfilaName.Split('.')[1];
                string NewName = FileName + "_New." + subName;
                item.SavOData(Path.Combine(DirPath,NewName));

                string csvStr = item.FileName + "," +
                                item.MarkerName + "," +
                                item.NUMBER + "," +
                                item.REC_TYPE_VERs + "," +
                                item.ANT_TYPE + "," +
                                item.ANTENNA_DETAIL_HEN + "," +
                                item.R + "," +
                                item.C + "," +
                                item.A + "," +
                                item.RAW_ANATANNA_HEIGHT + "," +
                                item.APC + "," +
                                item.ARP + "," +
                                item.MARK;
                csvRes.Add(csvStr);
            }

            newCSV(ListFilePath, csvRes);

        }

        private void newCSV(string path,List<string> data)
        {
            string FullfilaName = Path.GetFileName(path);
            string DirPath = Path.GetDirectoryName(path);
            string FileName = FullfilaName.Split('.')[0];
            string subName = FullfilaName.Split('.')[1];
            string NewName = FileName + "_New." + subName;
            string nPath = Path.Combine(DirPath, NewName);
            using (StreamWriter sw = new StreamWriter(nPath, false,Encoding.UTF8))
            {
                foreach (string item in data)
                {
                    sw.WriteLine(item);
                    sw.Flush();
                }
            }

        }
        




    }
}
