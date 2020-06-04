namespace _3_1_Form_RINES_TOOLS_1
{
    partial class Form1
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.btn_執行 = new System.Windows.Forms.Button();
            this.txt_csvFilePath = new System.Windows.Forms.TextBox();
            this.txt_FilesFolder = new System.Windows.Forms.TextBox();
            this.btn_csvFile = new System.Windows.Forms.Button();
            this.btn_FileFolder = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_執行
            // 
            this.btn_執行.Font = new System.Drawing.Font("新細明體", 14F);
            this.btn_執行.Location = new System.Drawing.Point(659, 191);
            this.btn_執行.Name = "btn_執行";
            this.btn_執行.Size = new System.Drawing.Size(92, 39);
            this.btn_執行.TabIndex = 0;
            this.btn_執行.Text = "執行";
            this.btn_執行.UseVisualStyleBackColor = true;
            this.btn_執行.Click += new System.EventHandler(this.button1_Click);
            // 
            // txt_csvFilePath
            // 
            this.txt_csvFilePath.Font = new System.Drawing.Font("新細明體", 14F);
            this.txt_csvFilePath.Location = new System.Drawing.Point(152, 50);
            this.txt_csvFilePath.Name = "txt_csvFilePath";
            this.txt_csvFilePath.Size = new System.Drawing.Size(550, 30);
            this.txt_csvFilePath.TabIndex = 1;
            // 
            // txt_FilesFolder
            // 
            this.txt_FilesFolder.Font = new System.Drawing.Font("新細明體", 14F);
            this.txt_FilesFolder.Location = new System.Drawing.Point(152, 127);
            this.txt_FilesFolder.Name = "txt_FilesFolder";
            this.txt_FilesFolder.Size = new System.Drawing.Size(550, 30);
            this.txt_FilesFolder.TabIndex = 2;
            // 
            // btn_csvFile
            // 
            this.btn_csvFile.Font = new System.Drawing.Font("新細明體", 9F);
            this.btn_csvFile.Location = new System.Drawing.Point(719, 49);
            this.btn_csvFile.Name = "btn_csvFile";
            this.btn_csvFile.Size = new System.Drawing.Size(32, 32);
            this.btn_csvFile.TabIndex = 3;
            this.btn_csvFile.Text = "......";
            this.btn_csvFile.UseVisualStyleBackColor = true;
            this.btn_csvFile.Click += new System.EventHandler(this.btn_csvFile_Click);
            // 
            // btn_FileFolder
            // 
            this.btn_FileFolder.Font = new System.Drawing.Font("新細明體", 9F);
            this.btn_FileFolder.Location = new System.Drawing.Point(719, 126);
            this.btn_FileFolder.Name = "btn_FileFolder";
            this.btn_FileFolder.Size = new System.Drawing.Size(32, 30);
            this.btn_FileFolder.TabIndex = 4;
            this.btn_FileFolder.Text = "......";
            this.btn_FileFolder.UseVisualStyleBackColor = true;
            this.btn_FileFolder.Click += new System.EventHandler(this.btn_FileFolder_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("新細明體", 14F);
            this.label1.Location = new System.Drawing.Point(40, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(106, 19);
            this.label1.TabIndex = 5;
            this.label1.Text = "CSV檔路徑:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("新細明體", 14F);
            this.label2.Location = new System.Drawing.Point(9, 134);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(137, 19);
            this.label2.TabIndex = 6;
            this.label2.Text = "o檔資料夾路徑:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(824, 253);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btn_FileFolder);
            this.Controls.Add(this.btn_csvFile);
            this.Controls.Add(this.txt_FilesFolder);
            this.Controls.Add(this.txt_csvFilePath);
            this.Controls.Add(this.btn_執行);
            this.Name = "Form1";
            this.Text = "O檔資料處理";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btn_執行;
        private System.Windows.Forms.TextBox txt_csvFilePath;
        private System.Windows.Forms.TextBox txt_FilesFolder;
        private System.Windows.Forms.Button btn_csvFile;
        private System.Windows.Forms.Button btn_FileFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
    }
}

