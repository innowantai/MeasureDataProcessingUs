using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.IO;
using System.Diagnostics;

namespace TeqcTransfer
{
    public partial class Form1 : Form
    {
        public string OriPath;
        public ArrayList G_Targets = new ArrayList();
        public ArrayList G_Method = new ArrayList();

        public Form1()
        {
            InitializeComponent();
            OriPath = System.Windows.Forms.Application.StartupPath;
             
            //// Loading Commands from the file of CommandList.txt and append to ListBox
            ArrayList lisBoxString = ReadText_sub(Path.Combine(OriPath, "CommandList.txt"));
            int i = 1;
            foreach (string ff in lisBoxString)
            {
                string[] index = ff.Split(';');
                G_Targets.Add(index[0].Trim());
                G_Method.Add(index[1].Trim());
                lisMethod.Items.Add(i.ToString() + " : " + index[2].Trim() + "(" + index[1].Trim() + ")");
                i++;
            }

            
            //// Loading past Path and Listbox selected Index
            StreamReader sr = new StreamReader(Path.Combine(OriPath, "lastPath.txt"));
            txtPath.Text = sr.ReadLine();
            lisMethod.SelectedIndex = Convert.ToInt32(sr.ReadLine());
            sr.Close();

        }

        private void btnPath_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog path = new FolderBrowserDialog();
            path.SelectedPath = txtPath.Text;

            if (path.ShowDialog() == DialogResult.OK)
            {
                //// Get Selected Path and sign to txtPath,saving selected Path and selected Index to LastPath.txt
                txtPath.Text = path.SelectedPath;
                string proPath = path.SelectedPath;
                SaveText_sub(Path.Combine(OriPath, "LastPath.txt"), path.SelectedPath, lisMethod.SelectedIndex.ToString());

                //// Loading All files from selected path
                ArrayList FIleName = GetFileName_sub(proPath);
                if (FIleName.Count > 0)
                { 
                    txtOutput.Text += "---- 總共有 " + FIleName.Count.ToString() + "筆檔案\r\n";

                    BATFileProcess_sub(FIleName, proPath);
                    MoveCompletedFiles_sub(proPath, FIleName);

                    txtOutput.Text += "---- 處理完成";
                }
                else
                {
                    txtOutput.Text += "---- 無檔案";

                }

            }; 
        }


        /// <summary>
        /// Moving Processed File to Indicated Folder
        /// </summary>
        /// <param name="proPath"></param>
        /// <param name="FIleName"></param>
        public void MoveCompletedFiles_sub(string proPath, ArrayList FIleName)
        {
            //// Get Command Name and Creat SaveFolder
            string index = lisMethod.SelectedItem.ToString();
            index = index.Replace(" : ", "_ ");
            Directory.CreateDirectory(Path.Combine(proPath, index));

            //// Moving All processed files
            foreach (string ff in FIleName)
            {
                string ff2 = ff.Replace(".", "a.");
                try
                {
                    File.Move(Path.Combine(proPath, ff2), Path.Combine(proPath, index, ff2));
                }
                catch (Exception)
                {
                    txtOutput.Text += "---- " + ff2 + " 檔案已存在\r\n";
                }
            }

        }


        /// <summary>
        /// Combinate Teqc command from each parameters
        /// </summary>
        /// <param name="FIleName"></param>
        /// <param name="proPath"></param>
        public void BATFileProcess_sub(ArrayList FIleName,string proPath)
        {
            //// Copy teqc.exe to target Folder if the file is not exists
            if (!File.Exists(Path.Combine(proPath, "teqc.exe")))
            {
                File.Copy(Path.Combine(OriPath, "teqc.exe"), Path.Combine(proPath, "teqc.exe"));
            }

            //// Create BAT commands
            ArrayList BATText = new ArrayList();
            int selected = lisMethod.SelectedIndex;
            foreach (string ff in FIleName)
            {
                string Combination = "teqc " + G_Targets[selected] + " " + G_Method[selected] + " " + ff + " > " + ff.Replace(".", "a.");
                BATText.Add(Combination);
            }
            SaveText_sub(Path.Combine(proPath, "CC.BAT"), BATText);
            

            //// Call BAT and Processing
            Directory.SetCurrentDirectory(proPath);
            Process p = new Process();
            p.StartInfo.FileName = Path.Combine(proPath, "CC.BAT");
            p.Start();
            p.WaitForExit(); //' 指示 Process 元件無限期等候相關處理序的結束。
        }


        /// <summary>
        /// Get files from indicated Path
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public ArrayList GetFileName_sub(string Path)
        {
            DirectoryInfo Dir = new DirectoryInfo(Path);
            ArrayList FIleName = new ArrayList();

            foreach (FileInfo f in Dir.GetFiles()) //查詢附檔名為""的文件  
            {
                string index = f.ToString();
                bool check = index.Contains("BAT") == true ? false : (index.Contains("exe") == true ? false : (index.Contains("a.") == true ? false : true));
                if (check)
                {
                    FIleName.Add(index); 
                } 
            }

            return FIleName;
        }


        /// <summary>
        /// Save text to indicate files and extension parameters General Case
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="list"></param>
        public void SaveText_sub(string Path, params string[] list)
        {
            ArrayList SaveData = new ArrayList();
            StreamWriter sw = new StreamWriter(Path);
            foreach (string ff in list)
            { 
                sw.WriteLine(ff);
                sw.Flush();
            }
            sw.Close();
            //SaveText_sub(Path, SaveData);
        }
         
        /// <summary>
        /// Save text to indicate files and Only one ArrayList Case
        /// </summary>
        /// <param name="Path"></param>
        /// <param name="SaveData"></param>
        public void SaveText_sub(string Path, ArrayList SaveData)
        {
            StreamWriter sw = new StreamWriter(Path);
            foreach (string ff in SaveData)
            {
                sw.WriteLine(ff);
                sw.Flush();
            }
            sw.Close();
        }
         
        /// <summary>
        /// Reading indicated file and saving the Variable of Data
        /// </summary>
        /// <param name="Path"></param>
        /// <returns></returns>
        public ArrayList ReadText_sub(string Path)
        {
            StreamReader sr = new StreamReader(Path, System.Text.Encoding.Default);
            ArrayList Data = new ArrayList();
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                Data.Add(line); 
            };
            sr.Close();
            return Data;
        }
         

        private void lisMethod_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtOutput_TextChanged(object sender, EventArgs e)
        {

        }



    }
}
