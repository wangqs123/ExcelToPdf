namespace ExcelToPdf
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnConvert = new System.Windows.Forms.Button();
            this.btnOpenSFile = new System.Windows.Forms.Button();
            this.btnOpenTPath = new System.Windows.Forms.Button();
            this.tbSFile = new System.Windows.Forms.TextBox();
            this.tbTFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.lbProcess = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnBatConvert = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.btnBatSPath = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBatTPath = new System.Windows.Forms.Button();
            this.tbBatTPath = new System.Windows.Forms.TextBox();
            this.tbBatSPath = new System.Windows.Forms.TextBox();
            this.folderBatS = new System.Windows.Forms.FolderBrowserDialog();
            this.folderBatT = new System.Windows.Forms.FolderBrowserDialog();
            this.label5 = new System.Windows.Forms.Label();
            this.rbNo = new System.Windows.Forms.RadioButton();
            this.rbMonth = new System.Windows.Forms.RadioButton();
            this.rbYear = new System.Windows.Forms.RadioButton();
            this.rbDay = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.lbSubPath = new System.Windows.Forms.Label();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnConvert
            // 
            this.btnConvert.Location = new System.Drawing.Point(205, 206);
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Size = new System.Drawing.Size(127, 61);
            this.btnConvert.TabIndex = 0;
            this.btnConvert.Text = "开始转换";
            this.btnConvert.UseVisualStyleBackColor = true;
            this.btnConvert.Click += new System.EventHandler(this.btnConvert_Click);
            // 
            // btnOpenSFile
            // 
            this.btnOpenSFile.Location = new System.Drawing.Point(440, 39);
            this.btnOpenSFile.Name = "btnOpenSFile";
            this.btnOpenSFile.Size = new System.Drawing.Size(39, 23);
            this.btnOpenSFile.TabIndex = 1;
            this.btnOpenSFile.Text = "...";
            this.btnOpenSFile.UseVisualStyleBackColor = true;
            this.btnOpenSFile.Click += new System.EventHandler(this.btnOpenSFile_Click);
            // 
            // btnOpenTPath
            // 
            this.btnOpenTPath.Location = new System.Drawing.Point(440, 86);
            this.btnOpenTPath.Name = "btnOpenTPath";
            this.btnOpenTPath.Size = new System.Drawing.Size(39, 23);
            this.btnOpenTPath.TabIndex = 2;
            this.btnOpenTPath.Text = "...";
            this.btnOpenTPath.UseVisualStyleBackColor = true;
            this.btnOpenTPath.Click += new System.EventHandler(this.btnOpenTPath_Click);
            // 
            // tbSFile
            // 
            this.tbSFile.Location = new System.Drawing.Point(112, 41);
            this.tbSFile.Name = "tbSFile";
            this.tbSFile.Size = new System.Drawing.Size(322, 21);
            this.tbSFile.TabIndex = 3;
            // 
            // tbTFile
            // 
            this.tbTFile.Location = new System.Drawing.Point(112, 88);
            this.tbTFile.Name = "tbTFile";
            this.tbTFile.Size = new System.Drawing.Size(322, 21);
            this.tbTFile.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(32, 44);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(71, 12);
            this.label1.TabIndex = 5;
            this.label1.Text = "源Excel文件";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(50, 93);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 12);
            this.label2.TabIndex = 6;
            this.label2.Text = "保存路径";
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl1.Location = new System.Drawing.Point(0, 0);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(528, 344);
            this.tabControl1.TabIndex = 7;
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.btnConvert);
            this.tabPage1.Controls.Add(this.label2);
            this.tabPage1.Controls.Add(this.btnOpenSFile);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.btnOpenTPath);
            this.tabPage1.Controls.Add(this.tbTFile);
            this.tabPage1.Controls.Add(this.tbSFile);
            this.tabPage1.Location = new System.Drawing.Point(4, 22);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(520, 318);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "单个";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.lbSubPath);
            this.tabPage2.Controls.Add(this.lbProcess);
            this.tabPage2.Controls.Add(this.progressBar1);
            this.tabPage2.Controls.Add(this.groupBox1);
            this.tabPage2.Controls.Add(this.label5);
            this.tabPage2.Controls.Add(this.btnBatConvert);
            this.tabPage2.Controls.Add(this.label3);
            this.tabPage2.Controls.Add(this.btnBatSPath);
            this.tabPage2.Controls.Add(this.label4);
            this.tabPage2.Controls.Add(this.btnBatTPath);
            this.tabPage2.Controls.Add(this.tbBatTPath);
            this.tabPage2.Controls.Add(this.tbBatSPath);
            this.tabPage2.Location = new System.Drawing.Point(4, 22);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(520, 318);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "批量";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // lbProcess
            // 
            this.lbProcess.AutoSize = true;
            this.lbProcess.Font = new System.Drawing.Font("SimSun", 17F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(134)));
            this.lbProcess.Location = new System.Drawing.Point(6, 264);
            this.lbProcess.Name = "lbProcess";
            this.lbProcess.Size = new System.Drawing.Size(70, 23);
            this.lbProcess.TabIndex = 15;
            this.lbProcess.Text = "--/--";
            // 
            // progressBar1
            // 
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Location = new System.Drawing.Point(3, 292);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(514, 23);
            this.progressBar1.TabIndex = 14;
            // 
            // btnBatConvert
            // 
            this.btnBatConvert.Location = new System.Drawing.Point(191, 241);
            this.btnBatConvert.Name = "btnBatConvert";
            this.btnBatConvert.Size = new System.Drawing.Size(138, 45);
            this.btnBatConvert.TabIndex = 7;
            this.btnBatConvert.Text = "开始转换";
            this.btnBatConvert.UseVisualStyleBackColor = true;
            this.btnBatConvert.Click += new System.EventHandler(this.btnBatConvert_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(39, 95);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 12);
            this.label3.TabIndex = 13;
            this.label3.Text = "保存至目录";
            // 
            // btnBatSPath
            // 
            this.btnBatSPath.Location = new System.Drawing.Point(449, 49);
            this.btnBatSPath.Name = "btnBatSPath";
            this.btnBatSPath.Size = new System.Drawing.Size(39, 23);
            this.btnBatSPath.TabIndex = 8;
            this.btnBatSPath.Text = "...";
            this.btnBatSPath.UseVisualStyleBackColor = true;
            this.btnBatSPath.Click += new System.EventHandler(this.btnBatSPath_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(33, 56);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(71, 12);
            this.label4.TabIndex = 12;
            this.label4.Text = "源Excel目录";
            // 
            // btnBatTPath
            // 
            this.btnBatTPath.Location = new System.Drawing.Point(449, 93);
            this.btnBatTPath.Name = "btnBatTPath";
            this.btnBatTPath.Size = new System.Drawing.Size(39, 23);
            this.btnBatTPath.TabIndex = 9;
            this.btnBatTPath.Text = "...";
            this.btnBatTPath.UseVisualStyleBackColor = true;
            this.btnBatTPath.Click += new System.EventHandler(this.btnBatTPath_Click);
            // 
            // tbBatTPath
            // 
            this.tbBatTPath.Location = new System.Drawing.Point(110, 95);
            this.tbBatTPath.Name = "tbBatTPath";
            this.tbBatTPath.Size = new System.Drawing.Size(333, 21);
            this.tbBatTPath.TabIndex = 11;
            // 
            // tbBatSPath
            // 
            this.tbBatSPath.Location = new System.Drawing.Point(110, 51);
            this.tbBatSPath.Name = "tbBatSPath";
            this.tbBatSPath.Size = new System.Drawing.Size(333, 21);
            this.tbBatSPath.TabIndex = 10;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(39, 169);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(89, 12);
            this.label5.TabIndex = 7;
            this.label5.Text = "保存子目录模式";
            // 
            // rbNo
            // 
            this.rbNo.AutoSize = true;
            this.rbNo.Location = new System.Drawing.Point(6, 15);
            this.rbNo.Name = "rbNo";
            this.rbNo.Size = new System.Drawing.Size(35, 16);
            this.rbNo.TabIndex = 0;
            this.rbNo.TabStop = true;
            this.rbNo.Text = "无";
            this.rbNo.UseVisualStyleBackColor = true;
            this.rbNo.Click += new System.EventHandler(this.rbSubPath_Click);
            // 
            // rbMonth
            // 
            this.rbMonth.AutoSize = true;
            this.rbMonth.Location = new System.Drawing.Point(88, 15);
            this.rbMonth.Name = "rbMonth";
            this.rbMonth.Size = new System.Drawing.Size(35, 16);
            this.rbMonth.TabIndex = 3;
            this.rbMonth.TabStop = true;
            this.rbMonth.Text = "月";
            this.rbMonth.UseVisualStyleBackColor = true;
            this.rbMonth.Click += new System.EventHandler(this.rbSubPath_Click);
            // 
            // rbYear
            // 
            this.rbYear.AutoSize = true;
            this.rbYear.Location = new System.Drawing.Point(47, 15);
            this.rbYear.Name = "rbYear";
            this.rbYear.Size = new System.Drawing.Size(35, 16);
            this.rbYear.TabIndex = 1;
            this.rbYear.TabStop = true;
            this.rbYear.Text = "年";
            this.rbYear.UseVisualStyleBackColor = true;
            this.rbYear.Click += new System.EventHandler(this.rbSubPath_Click);
            // 
            // rbDay
            // 
            this.rbDay.AutoSize = true;
            this.rbDay.Location = new System.Drawing.Point(129, 15);
            this.rbDay.Name = "rbDay";
            this.rbDay.Size = new System.Drawing.Size(35, 16);
            this.rbDay.TabIndex = 2;
            this.rbDay.TabStop = true;
            this.rbDay.Text = "日";
            this.rbDay.UseVisualStyleBackColor = true;
            this.rbDay.Click += new System.EventHandler(this.rbSubPath_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.rbDay);
            this.groupBox1.Controls.Add(this.rbYear);
            this.groupBox1.Controls.Add(this.rbMonth);
            this.groupBox1.Controls.Add(this.rbNo);
            this.groupBox1.Location = new System.Drawing.Point(134, 151);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(168, 41);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            // 
            // lbSubPath
            // 
            this.lbSubPath.AutoSize = true;
            this.lbSubPath.Location = new System.Drawing.Point(308, 170);
            this.lbSubPath.Name = "lbSubPath";
            this.lbSubPath.Size = new System.Drawing.Size(23, 12);
            this.lbSubPath.TabIndex = 9;
            this.lbSubPath.Text = "...";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 344);
            this.Controls.Add(this.tabControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Excel转Pdf";
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Button btnOpenSFile;
        private System.Windows.Forms.Button btnOpenTPath;
        private System.Windows.Forms.TextBox tbSFile;
        private System.Windows.Forms.TextBox tbTFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Button btnBatConvert;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button btnBatSPath;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBatTPath;
        private System.Windows.Forms.TextBox tbBatTPath;
        private System.Windows.Forms.TextBox tbBatSPath;
        private System.Windows.Forms.FolderBrowserDialog folderBatS;
        private System.Windows.Forms.FolderBrowserDialog folderBatT;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lbProcess;
        private System.Windows.Forms.Label lbSubPath;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RadioButton rbDay;
        private System.Windows.Forms.RadioButton rbYear;
        private System.Windows.Forms.RadioButton rbMonth;
        private System.Windows.Forms.RadioButton rbNo;
        private System.Windows.Forms.Label label5;
    }
}

