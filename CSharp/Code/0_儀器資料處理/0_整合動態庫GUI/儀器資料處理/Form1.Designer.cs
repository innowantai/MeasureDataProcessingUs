namespace 儀器資料處理
{
    partial class 儀器資料處理
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
            this.listImplement = new System.Windows.Forms.ListBox();
            this.listMethod = new System.Windows.Forms.ListBox();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.檔案ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.讀檔ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.水準儀DAT誤差ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.menuStrip1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // listImplement
            // 
            this.listImplement.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listImplement.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listImplement.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.listImplement.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.listImplement.FormattingEnabled = true;
            this.listImplement.ItemHeight = 18;
            this.listImplement.Items.AddRange(new object[] {
            "水準儀",
            "經緯儀",
            "GPS"});
            this.listImplement.Location = new System.Drawing.Point(18, 28);
            this.listImplement.Name = "listImplement";
            this.listImplement.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.listImplement.Size = new System.Drawing.Size(141, 126);
            this.listImplement.TabIndex = 0;
            this.listImplement.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.listImplement_DrawItem);
            this.listImplement.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.listImplement_MeasureItem);
            this.listImplement.SelectedIndexChanged += new System.EventHandler(this.listImplement_SelectedIndexChanged);
            // 
            // listMethod
            // 
            this.listMethod.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.listMethod.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.listMethod.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawVariable;
            this.listMethod.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.listMethod.FormattingEnabled = true;
            this.listMethod.Location = new System.Drawing.Point(18, 28);
            this.listMethod.Name = "listMethod";
            this.listMethod.Size = new System.Drawing.Size(141, 269);
            this.listMethod.TabIndex = 1;
            this.listMethod.DrawItem += new System.Windows.Forms.DrawItemEventHandler(this.listMethod_DrawItem);
            this.listMethod.MeasureItem += new System.Windows.Forms.MeasureItemEventHandler(this.listMethod_MeasureItem);
            this.listMethod.SelectedIndexChanged += new System.EventHandler(this.listMethod_SelectedIndexChanged);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.Window;
            this.menuStrip1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.檔案ToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(827, 28);
            this.menuStrip1.TabIndex = 3;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // 檔案ToolStripMenuItem
            // 
            this.檔案ToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.讀檔ToolStripMenuItem});
            this.檔案ToolStripMenuItem.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.檔案ToolStripMenuItem.Name = "檔案ToolStripMenuItem";
            this.檔案ToolStripMenuItem.Size = new System.Drawing.Size(53, 24);
            this.檔案ToolStripMenuItem.Text = "檔案";
            // 
            // 讀檔ToolStripMenuItem
            // 
            this.讀檔ToolStripMenuItem.Name = "讀檔ToolStripMenuItem";
            this.讀檔ToolStripMenuItem.Size = new System.Drawing.Size(158, 24);
            this.讀檔ToolStripMenuItem.Text = "讀檔與處理";
            this.讀檔ToolStripMenuItem.Click += new System.EventHandler(this.讀檔ToolStripMenuItem_Click);
            // 
            // 水準儀DAT誤差ToolStripMenuItem
            // 
            this.水準儀DAT誤差ToolStripMenuItem.Name = "水準儀DAT誤差ToolStripMenuItem";
            this.水準儀DAT誤差ToolStripMenuItem.Size = new System.Drawing.Size(32, 19);
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.listImplement);
            this.groupBox1.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.groupBox1.Location = new System.Drawing.Point(12, 34);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(177, 168);
            this.groupBox1.TabIndex = 5;
            this.groupBox1.TabStop = false;
            // 
            // label2
            // 
            this.label2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label2.AutoSize = true;
            this.label2.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label2.Location = new System.Drawing.Point(49, 5);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(73, 20);
            this.label2.TabIndex = 1;
            this.label2.Text = "儀器類型";
            // 
            // groupBox2
            // 
            this.groupBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.listMethod);
            this.groupBox2.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.groupBox2.Location = new System.Drawing.Point(12, 208);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(177, 316);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            // 
            // label3
            // 
            this.label3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left)));
            this.label3.AutoSize = true;
            this.label3.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.label3.Location = new System.Drawing.Point(49, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(73, 20);
            this.label3.TabIndex = 2;
            this.label3.Text = "檔案類型";
            // 
            // txtOutput
            // 
            this.txtOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.txtOutput.Font = new System.Drawing.Font("微軟正黑體", 12F);
            this.txtOutput.Location = new System.Drawing.Point(195, 46);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtOutput.Size = new System.Drawing.Size(610, 477);
            this.txtOutput.TabIndex = 7;
            this.txtOutput.TextChanged += new System.EventHandler(this.txtOutput_TextChanged);
            // 
            // 儀器資料處理
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(827, 547);
            this.Controls.Add(this.txtOutput);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.menuStrip1);
            this.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.Name = "儀器資料處理";
            this.ShowIcon = false;
            this.Text = "儀器資料處理";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listImplement;
        private System.Windows.Forms.ListBox listMethod;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 檔案ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 讀檔ToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem 水準儀DAT誤差ToolStripMenuItem;
        private System.Windows.Forms.TextBox txtOutput;
    }
}

