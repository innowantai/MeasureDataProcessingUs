using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace _3_1_Form_RINES_TOOLS_1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        FormGeneralSavingAndLoading fs;
        private void Form1_Load(object sender, EventArgs e)
        {
            this.fs = new FormGeneralSavingAndLoading(this.Controls,"O_DATA_PROCESSING");
            this.fs.Loading();
            if (this.txt_csvFilePath.Text == "0") this.txt_csvFilePath.Text = "";
            if (this.txt_FilesFolder.Text == "0") this.txt_FilesFolder.Text = ""; 
        }

        private void button1_Click(object sender, EventArgs e)
        {

            string csvPath = this.txt_csvFilePath.Text;
            string diro = this.txt_FilesFolder.Text;
            RINEX rr = new RINEX(csvPath, diro);
            MessageBox.Show("處理結束");

        }

        private void btn_csvFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog of = new OpenFileDialog();
            if (of.ShowDialog() != DialogResult.OK) return;
            this.txt_csvFilePath.Text = of.FileName;
            this.fs.Saving();
        }

        private void btn_FileFolder_Click(object sender, EventArgs e)
        {

            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.SelectedPath = this.txt_FilesFolder.Text;
            if (fbd.ShowDialog() != DialogResult.OK) return;
            this.txt_FilesFolder.Text = fbd.SelectedPath;
            this.fs.Saving();
        }

    }
}
