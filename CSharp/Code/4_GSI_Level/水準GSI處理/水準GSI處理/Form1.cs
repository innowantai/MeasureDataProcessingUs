using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;

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

            List<List<string>> GroupDatgas = Functions.
                                             DataProcess.
                                             LoadingAndClassGSIData(filePath);

            foreach (List<string> item in GroupDatgas)
            {
                GroupData gd = new GroupData(item); ;
                gd.Process();
            }
        }

        private void toolTip2_Popup(object sender, PopupEventArgs e)
        {

        }
    }
}
