using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace HCR轉檔
{
    public partial class Form1 : Form
    {

        string filName = "";
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }



        private void button1_Click(object sender, EventArgs e)
        {
            if (this.filName.Trim() == "")
            {
                MessageBox.Show("請填寫年份");
                return;
            }

            if (this.txt_Year.Text.Trim() == "" )
            {
                MessageBox.Show("請填寫月份");
                return;
            }
            if (this.txt_Month.Text.Trim() == "")
            {
                MessageBox.Show("請填寫河川代碼");
                return;
            }
            if (this.txt_RiverNumber.Text.Trim() == "")
            {
                MessageBox.Show("請先選擇檔案");
                return;
            }

            CROSS_DATA.year = this.txt_Year.Text;
            CROSS_DATA.month = this.txt_Month.Text;
            CROSS_DATA.number = this.txt_RiverNumber.Text;


            List<RowData> data = LoadHCR(this.filName);
            List<CROSS_DATA> CDs = data.ToCROSS_DATA();


            string SavePath = Path.GetDirectoryName(this.filName);
            string SaveFolder = Path.Combine(SavePath, "HCR轉檔結果");
            if (!Directory.Exists(SaveFolder)) Directory.CreateDirectory(SaveFolder);
            foreach (CROSS_DATA item in CDs)
            {
                item.ToClass();
                item.Save(SaveFolder);
            }



            MessageBox.Show($"執行完成, 總共{CDs.Count}筆資料");
            //this.filName = "";
        }




        List<RowData> LoadHCR(string filepath)
        {
            List<RowData> res = new List<RowData>();
            using (StreamReader sr = new StreamReader(filepath))
            {
                while (sr.Peek() != -1)
                {
                    string d = sr.ReadLine();
                    res.Add(new RowData(d.SplieToArray()));
                }
            }

            return res;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog off = new OpenFileDialog();
            off.Filter = "(HCR)|*.hcr";
            if (off.ShowDialog() != DialogResult.OK) return;
            this.filName = off.FileName;
        }
    }
}
