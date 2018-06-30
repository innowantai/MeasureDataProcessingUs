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
using System.Collections;

namespace oFIleProGUI
{
    public partial class Form1 : Form
    {
        public string saveTxtName = "lastPath.txt";
        public string oriPath = System.Environment.CurrentDirectory;

        public Form1()
        {
            InitializeComponent();
            StreamReader sr = new StreamReader(Path.Combine(oriPath, saveTxtName));
            txtPath.Text += sr.ReadLine();
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
                saveText();

                DirectoryInfo dir = new DirectoryInfo(path.SelectedPath);
                var AllFiles = dir.GetFiles();
                ArrayList Files = new ArrayList();
                foreach (var ff in AllFiles)
                {
                    if (ff.ToString().Contains(".xls"))
                    {
                        Files.Add(ff.ToString());
                    }
                }

                if (Files.Count != 0)
                {
                    string dataPath = txtPath.Text;
                    txtOutput.Text += "總共 " + Files.Count.ToString() + " 筆excel檔案 \r\n";
                    foreach (string ff in Files)
                    {
                        txtOutput.Text += GPSoFileProcess.GPSoFiles.GPSoFile_Main(Path.Combine(dataPath, ff), dataPath, dataPath, dataPath);

                    }
                }
                else
                {
                    txtOutput.Text += "無 xls 檔案";
                }

            };
        }


        private void saveText()
        {

            StreamWriter sw = new StreamWriter(Path.Combine(oriPath,saveTxtName));
            sw.WriteLine(txtPath.Text);
            sw.Flush();
            sw.Close(); 
        }


        private void txtPath_TextChanged(object sender, EventArgs e)
        {
        }



        private void txtOutput_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
