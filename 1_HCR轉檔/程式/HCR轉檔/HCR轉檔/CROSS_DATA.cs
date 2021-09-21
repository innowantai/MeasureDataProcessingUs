using System;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HCR轉檔
{
    [DebuggerDisplay("Name = {NAME}, Count = {data.Count}")]
    public class CROSS_DATA
    {
        public static string year;
        public static string month;
        public static string number;

        public string NAME { get; private set; }
        public List<RowData> data { get; set; }
        public int dataCount { get; set; }
        public int dividCount { get; set; }
        public List<RowData> Results { get; set; }

        public CROSS_DATA(List<RowData> data)
        {
            this.NAME = data.First().data[1];
            data.RemoveAt(0);
            this.data = data.Select(t => t.Distint()).ToList();
        }



        public void ToClass()
        {

            Dictionary<int, List<RowData>> res = new Dictionary<int, List<RowData>>();


            int num = this.data.Count;
            int gNum = Convert.ToInt32(Math.Ceiling(Convert.ToDouble(num ) / 4));

            this.dataCount = num;
            this.dividCount = gNum;


            int iter = 0;
            bool IsSave = false;
            List<RowData> rd = new List<RowData>();
            for (int kk = 0; kk < num; kk++)
            {
                IsSave = false;
                rd.Add(this.data[kk]);
                if ((kk+1) % gNum == 0)
                {
                    res[iter] = new List<RowData>(rd);
                    iter++;
                    rd = new List<RowData>();
                    IsSave = true;
                }
            }

            if (!IsSave) res[iter] = rd;

            List<ColData> CLs = new List<ColData>();
            foreach (List<RowData> item in res.Values)
            {
                List<ColData> ii = item.ToStringArray().ToColData();
                CLs.AddRange(ii);
            }

            RowData header = new RowData(new string[] {
            CROSS_DATA.year,CROSS_DATA.number,this.NAME,this.dataCount.ToString(),
            this.dividCount.ToString(),CROSS_DATA.month});

            RowData blockRow = new RowData(new string[] { "" });
            List<RowData> newRd = CLs.ToStringArray().ToRowData();
            newRd.Insert(0, blockRow);
            newRd.Insert(0, header);
            Results = newRd;
        }

        public void Save(string saveFolder)
        {

            string ff = Path.Combine(saveFolder, CROSS_DATA.year + "-" + this.NAME + ".txt");
            using (StreamWriter sw = new StreamWriter(ff))
            {
                foreach (RowData item in Results)
                {
                    sw.WriteLine(item.ToOutFormat());
                    sw.Flush();
                }
            }
        }

        private void TmpSav(List<RowData> data)
        {
            string ff = @"C:\Users\innow\Desktop\HCR轉檔\參考資歷\22650-24350\tmp.txt";
            using (StreamWriter sw = new StreamWriter(ff))
            {
                foreach (RowData item in data)
                {
                    sw.WriteLine(item.ToOutFormat());
                    sw.Flush();
                }
            }
        }


    }
}
