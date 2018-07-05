using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Collections;

namespace 儀器資料處理
{
    public partial class 儀器資料處理 : Form
    {
        public string oriPath = System.Environment.CurrentDirectory;
        public string lastPath = "";
        public int selectIndex1 = 0;
        public int selectIndex2 = 0;
        public string obError = "";
        public string[,] subName = new string[,] { { ".DAT", "", "", "" }, 
                                                   { ".GSI", ".xls", ".txt", ".txt" }, 
                                                   { ".RES", ".SUM", ".csv", "" } };

        public 儀器資料處理()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoadPastData();
            listImplement.SelectedIndex = 0;

            listImplement.SelectedIndex = selectIndex1;
            listMethod.SelectedIndex = selectIndex2;

        }




        private void 讀檔ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.SelectedPath = lastPath;

            if (path.ShowDialog() == DialogResult.OK)
            {
                lastPath = path.SelectedPath;
                int selIndex1 = listImplement.SelectedIndex;
                int selIndex2 = listMethod.SelectedIndex;
                string fileSubName = subName[selIndex1, selIndex2];
                ArrayList files = GetFileName_sub(lastPath, fileSubName);
                 

                string index = txtOutput.Text;
                obError = txtOutput.Text.Contains("DAT誤差設置(m)") ? index.Substring(index.IndexOf(":") + 1, index.Length - index.IndexOf(":") - 1).Trim() : obError;

                bool keepGo = files.Count == 0 ? true : false;
                if (keepGo) {
                    txtOutput.Text = "無" + fileSubName + "資料";
                }
                else
                {
                    files = TheCmpDataLoadingProcess(files);
                    txtOutput.Text = "總共有" + files.Count + "筆資料\r\n";
                    int kk = 1;
                    foreach (string ff in files)
                    {
                        txtOutput.Text += "第" + kk.ToString() + "筆資料(" + ff + ") :";
                        txtOutput.Text += ProcessMain(oriPath, lastPath, ff, obError);
                        kk++;
                    }
                }


                SavePastData();
            };
        }
         


        private string ProcessMain(string oriPath, string savePath, string fileName, string Standard)
        {
            string res = "";
            if (listImplement.GetSelected(0))
            {
                if (listMethod.GetSelected(0))
                {
                    res = OBDAT.OBDAT.OBMain_sub(oriPath, savePath, fileName, Standard);
                }
            }
            else if (listImplement.GetSelected(1))
            {
                if (listMethod.GetSelected(0))
                {
                    res = TheGSI.TheGSI.TheGSI_Main(savePath, savePath, fileName); 
                }
                else if (listMethod.GetSelected(1))
                {
                    res = TheCmpExcelData.TheCmpExcelData.TheCmp_Main(savePath, savePath, fileName);
                }
                else if (listMethod.GetSelected(2))
                {
                    res = TheNIKONtoAGA.TheNIKONtoAGA.TheNIKONtoAGA_Main(savePath, savePath, fileName);
                }
                else if (listMethod.GetSelected(3))
                {
                    res = TheZTStoAGA.TheZTStoAGA.TheZTStoAGA_Main(savePath, savePath, fileName);
                }
            }
            else if (listImplement.GetSelected(2))
            {

                if (listMethod.GetSelected(0))
                {
                    res = GPS_RES.GPSRES.GPSRES_Main(fileName, savePath, oriPath, savePath); 
                }
                else if (listMethod.GetSelected(1))
                {
                    res = GPS_SUM.GPSSUM.GPSSUM_main(savePath, savePath, fileName);
                }
                else if (listMethod.GetSelected(2))
                {
                    res = GPS_SORT.GPSSORT.GPSSORT_Main(savePath, savePath, fileName);
                }
            }

            return res;
        }

        /// <summary>
        /// TheCmpExcelData 讀檔檔名需要特別處理
        /// </summary>
        /// <param name="files"></param>
        /// <returns></returns>
        public ArrayList TheCmpDataLoadingProcess(ArrayList files)
        {
            if (listImplement.SelectedIndex == 1 && listMethod.SelectedIndex == 1)
            {
                ArrayList files2 = new ArrayList();
                foreach (string ff in files)
                {
                    if (!ff.Contains("a") && !ff.Contains("Result"))
                    {
                        files2.Add(ff);
                    }

                }
                return files2;
            }
            return files;
        }


        /// <summary>
        /// Get files from indicated Path
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public ArrayList GetFileName_sub(string Path, string fileSubName)
        {
            DirectoryInfo Dir = new DirectoryInfo(Path);
            ArrayList FIleName = new ArrayList();

            foreach (FileInfo f in Dir.GetFiles()) //查詢附檔名為""的文件  
            {
                string index = f.ToString();
                if (index.Contains(fileSubName.ToLower()) | index.Contains(fileSubName.ToUpper()))
                {
                    FIleName.Add(index);
                }
            }
            return FIleName;
        }





        private void listImplement_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listImplement.GetSelected(0))
            {
                listMethod.Items.Clear();
                listMethod.Items.Add("DAT");
            }
            else if (listImplement.GetSelected(1))
            {
                listMethod.Items.Clear();
                listMethod.Items.Add("GSI");
                listMethod.Items.Add("CmpExcel");
                listMethod.Items.Add("NIKONtoAGA");
                listMethod.Items.Add("ZTStoAGA");
            }
            else if (listImplement.GetSelected(2))
            {
                listMethod.Items.Clear();
                listMethod.Items.Add("RES");
                listMethod.Items.Add("SUM");
                listMethod.Items.Add("SORT");
            }
            listMethod.SelectedIndex = 0;
        }


        private void listMethod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listImplement.GetSelected(0) && listMethod.GetSelected(0))
            {
                txtOutput.Text = "DAT誤差設置(m) : " + obError;
            }
            else if(listImplement.GetSelected(2) && listMethod.GetSelected(1))
            {
                txtOutput.Text = "先執行完'RES程式'產生excel檔後再執行此程式";
            }
            else
            {
                txtOutput.Text = "";
            }
        }

        public void LoadPastData()
        {
            StreamReader sr = new StreamReader(Path.Combine(oriPath, "PastData.txt"));
            lastPath = sr.ReadLine();
            selectIndex1 = Convert.ToInt32(sr.ReadLine());
            selectIndex2 = Convert.ToInt32(sr.ReadLine());
            obError = sr.ReadLine();
            sr.Close();
        }

        public void SavePastData()
        {
            StreamWriter sw = new StreamWriter(Path.Combine(oriPath, "PastData.txt"));
            sw.WriteLine(lastPath);
            sw.Flush();
            sw.WriteLine(listImplement.SelectedIndex);
            sw.Flush();
            sw.WriteLine(listMethod.SelectedIndex);
            sw.Flush();
            sw.WriteLine(obError);
            sw.Flush();
            sw.Close();
        }

        private void listImplement_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.Graphics.FillRectangle(new SolidBrush(e.BackColor), e.Bounds);
            if (e.Index >= 0)
            {
                StringFormat sStringFormat = new StringFormat();
                sStringFormat.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(((ListBox)sender).Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds, sStringFormat);
            }
            e.DrawFocusRectangle();
        }


        private void listImplement_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = e.ItemHeight + 4;
        }

        private void listMethod_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.Graphics.FillRectangle(new SolidBrush(e.BackColor), e.Bounds);
            if (e.Index >= 0)
            {
                StringFormat sStringFormat = new StringFormat();
                sStringFormat.LineAlignment = StringAlignment.Center;
                e.Graphics.DrawString(((ListBox)sender).Items[e.Index].ToString(), e.Font, new SolidBrush(e.ForeColor), e.Bounds, sStringFormat);
            }
            e.DrawFocusRectangle();
        }

        private void listMethod_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            e.ItemHeight = e.ItemHeight + 10;
        }

        private void txtOutput_TextChanged(object sender, EventArgs e)
        {

        }

    }
}
