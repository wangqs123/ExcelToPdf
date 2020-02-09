using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ExcelToPdf
{
    public partial class Form1 : Form
    {
        string fileNameModel = "";
        public Form1()
        {
            InitializeComponent();
            fileNameModel = ConfigurationManager.AppSettings.Get("FileNameModel");
            this.tbBatSPath.Text = ConfigurationManager.AppSettings.Get("SPath");
            this.tbBatTPath.Text = ConfigurationManager.AppSettings.Get("TPath");
            var subPathModel = ConfigurationManager.AppSettings.Get("SubPathModel");
            if (subPathModel == "1"){
                rbYear.Checked = true;
                lbSubPath.Text = DateTime.Now.ToString("yyyy");
            }
            else if (subPathModel == "2"){
                this.rbMonth.Checked = true;
                lbSubPath.Text = DateTime.Now.ToString("yyyyMM");
            }
            else if (subPathModel == "3"){
                this.rbDay.Checked = true;
                lbSubPath.Text = DateTime.Now.ToString("yyyyMMdd");
            }
            else
            {
                this.rbNo.Checked = true;
                lbSubPath.Text = "";
            }
        }

        #region 单个
        private void btnConvert_Click(object sender, EventArgs e)
        {
            ConverterToPdf(tbSFile.Text, tbTFile.Text);
        }

        /// <summary>  
        /// 转换excel 成PDF文档  
        /// </summary>  
        /// <param name="_lstrInputFile">原文件路径</param>  
        /// <param name="_lstrOutFile">pdf文件输出路径</param>  
        /// <returns>true 成功</returns>  
        public bool ConverterToPdf(string _lstrInputFile, string _lstrOutFile)
        {
            if (!File.Exists(_lstrInputFile))
            {
                MessageBox.Show("源文件不存在", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (string.IsNullOrEmpty(_lstrOutFile) || _lstrOutFile.IndexOf("\\") < 0)
            {
                MessageBox.Show("请指定目标文件", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (!Directory.Exists(_lstrOutFile.Substring(0, _lstrOutFile.LastIndexOf("\\") + 1)))
            {
                MessageBox.Show("保存目录不存在", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }
            if (File.Exists(_lstrOutFile) && MessageBox.Show("文件【" + _lstrOutFile.Substring(_lstrOutFile.LastIndexOf("\\")+1) + "】已存在，是否覆盖？", "", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != System.Windows.Forms.DialogResult.Yes)
            {
                return false;   
            }
            _lstrOutFile = _lstrOutFile.Substring(0, _lstrOutFile.LastIndexOf("."));
            Microsoft.Office.Interop.Excel.Application lobjExcelApp = null;
            Microsoft.Office.Interop.Excel.Workbooks lobjExcelWorkBooks = null;
            Microsoft.Office.Interop.Excel.Workbook lobjExcelWorkBook = null;

            string lstrTemp = string.Empty;
            object lobjMissing = System.Reflection.Missing.Value;

            try
            {
                lobjExcelApp = new Microsoft.Office.Interop.Excel.Application();
                lobjExcelApp.Visible = true;
                lobjExcelWorkBooks = lobjExcelApp.Workbooks;
                lobjExcelWorkBook = lobjExcelWorkBooks.Open(_lstrInputFile, true, true, lobjMissing, lobjMissing, lobjMissing, true,
                    lobjMissing, lobjMissing, lobjMissing, lobjMissing, lobjMissing, false, lobjMissing, lobjMissing);

                //Microsoft.Office.Interop.Excel 12.0.0.0之后才有这函数              
                //注释掉好像没问题  lstrTemp = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xls" + (lobjExcelWorkBook.HasVBProject ? 'm' : 'x');
                //lstrTemp = System.IO.Path.GetTempPath() + Guid.NewGuid().ToString() + ".xls";  
                //注释掉好像没问题lobjExcelWorkBook.SaveAs(lstrTemp, Microsoft.Office.Interop.Excel.XlFileFormat.xlExcel4Workbook, Type.Missing, Type.Missing, Type.Missing, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing,
                //注释掉好像没问题   false, Type.Missing, Type.Missing, Type.Missing);
                //输出为PDF 第一个选项指定转出为PDF,还可以指定为XPS格式  
                lobjExcelWorkBook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, _lstrOutFile, Microsoft.Office.Interop.Excel.XlFixedFormatQuality.xlQualityStandard, Type.Missing, false, Type.Missing, Type.Missing, false, Type.Missing);
                lobjExcelWorkBooks.Close();
                lobjExcelApp.Quit();
            }
            catch (Exception ex)
            {
                //其他日志操作；  
                return false;
            }
            finally
            {
                try
                {
                    if (lobjExcelWorkBook != null)
                    {
                        lobjExcelWorkBook.Close(Type.Missing, Type.Missing, Type.Missing);
                        Marshal.ReleaseComObject(lobjExcelWorkBook);
                        lobjExcelWorkBook = null;

                    }
                    if (lobjExcelWorkBooks != null)
                    {
                        lobjExcelWorkBooks.Close();
                        Marshal.ReleaseComObject(lobjExcelWorkBooks);
                        lobjExcelWorkBooks = null;
                    }
                    if (lobjExcelApp != null)
                    {
                        lobjExcelApp.Quit();
                        Marshal.ReleaseComObject(lobjExcelApp);
                        lobjExcelApp = null;

                    }
                }
                catch (Exception e)
                {

                }
                //主动激活垃圾回收器，主要是避免超大批量转文档时，内存占用过多，而垃圾回收器并不是时刻都在运行！  
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return true;
        }

        /// <summary>
        /// 根据文件名模板获取文件名
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="fileNameModel"></param>
        /// <returns></returns>
        public string GetFileName(string filePath, string fileNameModel)
        {
            if (string.IsNullOrEmpty(filePath) || !File.Exists(filePath))
                return "";
            var dicParams = AnalyzeExp(fileNameModel);
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(filePath);
            try
            {
                FileStream fs = File.OpenRead(filePath);
                if (extension.Equals(".xls"))
                {
                    //把xls文件中的数据写入wk中
                    wk = new HSSFWorkbook(fs);
                }
                else
                {
                    //把xlsx文件中的数据写入wk中
                    wk = new XSSFWorkbook(fs);
                }

                fs.Close();
                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);
                foreach (var item in dicParams)
                {
                    var arr = item.Key.TrimStart('[').TrimEnd(']').Split(new string[]{",","，"}, StringSplitOptions.RemoveEmptyEntries);
                    if (arr.Length == 2 && Sysfun.IsNum(arr[0]) && Sysfun.IsNum(arr[1]))
                    {
                        IRow row = sheet.GetRow(Sysfun.ConverToInt(arr[0], 0)-1);  //读取当前行数据
                        if(row == null)
                            continue;
                        var cell = row.GetCell(Sysfun.ConverToInt(arr[1], 0) - 1);
                        if (cell == null)
                            continue;
                        fileNameModel = fileNameModel.Replace(item.Key, cell.ToString());
                    }
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("出现异常：" + e.Message, "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return "";
            }
            return fileNameModel;
        }

        /// <summary>
        /// 解析表达式中的变量
        /// </summary>
        /// <param name="node"></param>
        /// <param name="expression"></param>
        /// <returns></returns>
        private Dictionary<string, string> AnalyzeExp(string expression)
        {
            var dic = new Dictionary<string, string>();
            if (string.IsNullOrEmpty(expression))
                return dic;
            Regex reg = new Regex(@"\[(.+?)]");
            foreach (Match m in reg.Matches(expression))
            {
                if (!dic.ContainsKey(m.Value))
                    dic.Add(m.Value, "");
            }
            return dic;
        }

        private void btnOpenSFile_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Excel2007(.xlsx)|*.xlsx|Excel(.xls)|*.xls";
            openFileDialog1.DefaultExt = "";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                tbSFile.Text = openFileDialog1.FileName;//获取文件路径 
            }
        }

        private void btnOpenTPath_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                tbTFile.Text = folderBrowserDialog1.SelectedPath;
                var fileName = GetFileName(tbSFile.Text, fileNameModel);
                if (!string.IsNullOrEmpty(fileName))
                    tbTFile.Text = tbTFile.Text + @"\" + fileName + ".pdf";
            }
        }
        #endregion

        #region 批量
        private void btnBatSPath_Click(object sender, EventArgs e)
        {
            if (this.folderBatS.ShowDialog() == DialogResult.OK)
            {
                this.tbBatSPath.Text = folderBatS.SelectedPath;
            }
        }

        private void btnBatTPath_Click(object sender, EventArgs e)
        {
            if (this.folderBatT.ShowDialog() == DialogResult.OK)
            {
                this.tbBatTPath.Text = folderBatT.SelectedPath;
            }
        }
        private void btnBatConvert_Click(object sender, EventArgs e)
        {
            if (!Directory.Exists(tbBatSPath.Text))
            {
                MessageBox.Show("源目录不存在", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!Directory.Exists(tbBatTPath.Text))
            {
                MessageBox.Show("目标目录不存在", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["TPath"].Value = this.tbBatTPath.Text;
            config.AppSettings.Settings["SPath"].Value = this.tbBatSPath.Text;
            config.Save(ConfigurationSaveMode.Modified);
            System.Configuration.ConfigurationManager.RefreshSection("appSettings");
            DirectoryInfo TheFolder = new DirectoryInfo(tbBatSPath.Text);
            List<FileInfo> fileList = new List<FileInfo>();
            //遍历文件
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Extension == ".xlsx" || NextFile.Extension == ".xls")
                    fileList.Add(NextFile);
            }
            if (fileList.Count == 0)
            {
                MessageBox.Show("源目录[" + tbBatSPath.Text + "]下无Excel文件");
                return;
            }
            progressBar1.Maximum = fileList.Count;
            progressBar1.Step = 1;
            progressBar1.Value = 0;
            lbProcess.Text = "0/" + fileList.Count;
            int succCount = 0;
            foreach (var file in fileList)
            {
                var fileName = GetFileName(file.FullName, fileNameModel);
                if (!Directory.Exists(tbBatTPath.Text + @"\" + this.lbSubPath.Text))
                {
                    Directory.CreateDirectory(tbBatTPath.Text + @"\" + this.lbSubPath.Text);
                }
                if (ConverterToPdf(file.FullName, tbBatTPath.Text + @"\" + this.lbSubPath.Text + @"\" + fileName + ".pdf"))
                {
                    succCount++;
                }

                if (this.progressBar1.Value < this.progressBar1.Maximum)
                {
                    this.progressBar1.PerformStep();
                }
                else
                {
                    this.progressBar1.Value = this.progressBar1.Maximum;
                }
                lbProcess.Text = progressBar1.Value+ "/" + fileList.Count;
            }
            MessageBox.Show("共处理文件"+fileList.Count+"个，成功"+succCount+"个");
        }

        private void rbSubPath_Click(object sender, EventArgs e)
        {
            string model = "0";
            if (this.rbYear.Checked)
            {
                model = "1";
                this.lbSubPath.Text = DateTime.Now.ToString("yyyy");
            }
            else if (this.rbMonth.Checked)
            {
                model = "2";
                this.lbSubPath.Text = DateTime.Now.ToString("yyyyMM");
            }
            else if (this.rbDay.Checked)
            {
                model = "3";
                this.lbSubPath.Text = DateTime.Now.ToString("yyyyMMdd");
            }
            else
            {
                model = "0";
                this.lbSubPath.Text = "";
            }
            Configuration config = System.Configuration.ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings["SubPathModel"].Value = model;
            config.Save(ConfigurationSaveMode.Modified);
            System.Configuration.ConfigurationManager.RefreshSection("appSettings");
        }
        #endregion
    }
}
