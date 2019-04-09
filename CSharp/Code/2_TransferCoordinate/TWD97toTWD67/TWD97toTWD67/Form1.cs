using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace TWD97toTWD67
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.ShowIcon = false;
            this.Loading();

        }

        private string TmpDataSavePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),"LastPath9767.txt");
        string lastPath = "";


        private void btnSingle_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "選擇excel檔案";
            dialog.Filter = "xlsx files (*.*)|*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string SavePath = Path.GetDirectoryName(dialog.FileName);
                string SaveFileName = this.txtSingleFileName.Text.Trim() == "" ?
                                       (Path.GetFileNameWithoutExtension(dialog.FileName) + "_圖號座標轉換.xlsx") : 
                                       this.txtSingleFileName.Text + "_圖號座標轉換.xlsx";
                TransCoordinateSystem TCS = new TransCoordinateSystem();
                TCS.Main_start(dialog.FileName, SavePath, SaveFileName);
                MessageBox.Show("檔案處理完成!"); 
            }
        }



        private void btnGroup_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog Folder = new FolderBrowserDialog();
            Folder.SelectedPath = this.lastPath;

            if (Folder.ShowDialog() == DialogResult.OK)
            {
                this.lastPath = Folder.SelectedPath;
                List<string> FILEs = ShiftFileNames(Directory.GetFiles(lastPath).ToList());
                if (FILEs.Count > 0)
                {
                    TransCoordinateSystem TCS = new TransCoordinateSystem();
                    foreach (string file in FILEs)
                    { 
                        string SaveFileName = Path.GetFileNameWithoutExtension(file) + "_圖號座標轉換.xlsx";
                        TCS.Main_start(file, lastPath, SaveFileName);
                    }
                    MessageBox.Show("檔案處理完成!");
                    this.Saveing();
                    return;
                }

                MessageBox.Show("無excel檔案!");
            };
        }


        private List<string> ShiftFileNames(List<string> Files)
        {
            List<string> result = new List<string>();
            foreach (string item in Files) if (item.Contains(".xlsx")) result.Add(item); 

            return result;
        }


        private void Saveing()
        {
            using (StreamWriter sw = new StreamWriter(TmpDataSavePath))
            {
                sw.WriteLine(this.lastPath);
                sw.Flush();
            } 
        }

        private void Loading()
        {
            try
            { 
                using (StreamReader sr = new StreamReader(TmpDataSavePath))
                {
                    this.lastPath = sr.ReadLine();
                }
            }
            catch (Exception)
            { 
            }
        }

    }
}
