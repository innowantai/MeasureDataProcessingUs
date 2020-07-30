using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace 水準GSI處理
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void btn_執行_Click(object sender, EventArgs e)
        {
            string filePath = @"C:\Users\innow\Desktop\MeasureDataProcessingUs\CSharp\Code\4_GSI_Level\TryData\M_0616JD\M_0616JD.GSI";

            this.LoadingGSIData(filePath);
        }



        private void LoadingGSIData(string filePath)
        {
            List<string> tmpData = new List<string>();
            using ( StreamReader sr = new StreamReader(filePath))
            {
                while (sr.Peek() != -1)
                {
                    tmpData.Add(sr.ReadLine());
                }
            }

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
            /// 25    145
            /// 

        }
    }
}
