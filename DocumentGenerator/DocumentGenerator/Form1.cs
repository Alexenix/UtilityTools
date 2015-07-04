using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace DocumentGenerator
{
    public partial class Form1 : Form
    {
        private string csvInputFilePath = string.Empty;
        private string docTemplateFilePath = string.Empty;
        private string targetOutputFolder = string.Empty;
        private object missing = System.Reflection.Missing.Value;
        private const string EMPTY_PLACE_HOLDER = "          ";

        public Form1()
        {
            InitializeComponent();
        }

        private void buttonCSV_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择投资人列表文件";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "CSV Files(*.csv)|*.csv|Text Files(*.txt)|*.txt";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBoxCSVInput.Text = openFileDialog1.FileName;
                csvInputFilePath = openFileDialog1.FileName;
                toolStripStatusLabel.Text = string.Format("投资人列表文件选择成功: {0}", csvInputFilePath);
            }

            openFileDialog1.Dispose();
        }

        private void buttonDocTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "请选择Word文档模板";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "Word Document(*.docx)|*.docx|Word 97-2003 Document(*.doc)|*.doc";
            openFileDialog1.FilterIndex = 1;
            openFileDialog1.CheckFileExists = true;
            openFileDialog1.CheckPathExists = true;
            openFileDialog1.Multiselect = false;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //t .Text = openFileDialog1.FileName;
                textBoxDocTemplate.Text = openFileDialog1.FileName;
                docTemplateFilePath = openFileDialog1.FileName;
                toolStripStatusLabel.Text = string.Format("Word文档模板选择成功: {0}", docTemplateFilePath);
            }

            openFileDialog1.Dispose();
        }

        private void buttonTargetFolder_Click(object sender, EventArgs e)
        {
            // Show the FolderBrowserDialog.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBoxTargetFolder.Text = folderBrowserDialog1.SelectedPath;
                targetOutputFolder = folderBrowserDialog1.SelectedPath;
                toolStripStatusLabel.Text = string.Format("目标文件夹选择成功: {0}", targetOutputFolder);
            }

            folderBrowserDialog1.Dispose();
        }
       
        private void button1_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(csvInputFilePath) || string.IsNullOrEmpty(docTemplateFilePath) || string.IsNullOrEmpty(targetOutputFolder)) 
            {
                MessageBox.Show("请确保投资人列表文件、文档模板和目标文件夹都已经正确选择。");
                return;
            }

            try
            {
                if (!File.Exists(csvInputFilePath))
                {
                    MessageBox.Show(string.Format("投资人列表文件不存在: {0}", csvInputFilePath));
                    return;
                }

                if (!File.Exists(docTemplateFilePath))
                {
                    MessageBox.Show(string.Format("文档模板文件不存在: {0}", docTemplateFilePath));
                    return;
                }

                if (!Directory.Exists(targetOutputFolder))
                {
                    Directory.CreateDirectory(targetOutputFolder);
                }

                string[] allLines = null;
                allLines = File.ReadAllLines(csvInputFilePath);
                if (allLines == null || allLines.Length == 0) 
                {
                    MessageBox.Show(string.Format("投资人列表文件为空: {0}", csvInputFilePath));
                    return;
                }

                toolStripStatusLabel.Text = "文档批量生成中，请等待...";
                DisableButtons();
                toolStripProgressBar1.Maximum = allLines.Length - 1;
                toolStripProgressBar1.Minimum = 0;
                toolStripProgressBar1.Step = 1;
                toolStripProgressBar1.Value = 0;

                //construct column keys
                IDictionary<string, string> columnKeyValuePairs = new Dictionary<string, string>();
                string[] keyArray = allLines[0].Split(new char[] { '|' },StringSplitOptions.None);
                for (int i = 0; i < keyArray.Length; i++) 
                {
                    columnKeyValuePairs.Add(new KeyValuePair<string, string>(keyArray[i], string.Empty));
                }
                
                string currentLine = string.Empty;
                string tempOutputFilePath = string.Empty;
                string tempOutputFileName = string.Empty;
                string outputFileExtension = docTemplateFilePath.EndsWith("docx") ? ".docx" : ".doc";                
                Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application { Visible = false };

                for (int i = 1; i < allLines.Length; i++) 
                {
                    currentLine = allLines[i];
                    string[] valueArray = currentLine.Split(new char[] { '|' }, StringSplitOptions.None);
                    if (valueArray.Length != keyArray.Length) 
                    {
                        //Todo: Add error prompt
                        toolStripProgressBar1.PerformStep();
                        continue;
                    }

                    bool containsEmptyField = false;
                    for (int k = 0; k < keyArray.Length; k++)
                    {
                        if (string.IsNullOrEmpty(valueArray[k]))
                        {
                            columnKeyValuePairs[keyArray[k]] = EMPTY_PLACE_HOLDER;
                            containsEmptyField = true;
                        }
                        else
                        {
                            columnKeyValuePairs[keyArray[k]] = valueArray[k];
                        }
                    }

                    if (containsEmptyField == true)
                    {
                        tempOutputFileName = string.Format("{0}-{1}-{2}-Incomplete{3}", i, valueArray[0], valueArray[1], outputFileExtension);
                    }
                    else 
                    {
                        tempOutputFileName = string.Format("{0}-{1}-{2}{3}", i, valueArray[0], valueArray[1], outputFileExtension);                        
                    }

                    tempOutputFilePath = Path.Combine(targetOutputFolder, tempOutputFileName);
                    File.Copy(docTemplateFilePath, tempOutputFilePath, true);

                    ReplaceDocumentContent(wordApp, tempOutputFilePath, columnKeyValuePairs);
                    toolStripProgressBar1.PerformStep();               
                }

                wordApp.Quit(ref missing, ref missing, ref missing);
                toolStripStatusLabel.Text = string.Format("批量生成文档成功, 目标文件夹: {0}", targetOutputFolder);

            }
            catch (Exception exp)
            {
                MessageBox.Show(string.Format("Exception Happened: {0}", exp.ToString()));
            }
            finally 
            {
                EnableButtons();
            }
        }

        private void ReplaceDocumentContent(Microsoft.Office.Interop.Word.Application wordApp, object outputFilePath, IDictionary<string, string> columnKeyValuePairs) 
        {
            Microsoft.Office.Interop.Word.Document docObj = wordApp.Documents.Open(ref outputFilePath, ReadOnly: false, Visible: true);
            docObj.Activate();

            Microsoft.Office.Interop.Word.Find fnd = wordApp.ActiveWindow.Selection.Find;

            foreach (KeyValuePair<string, string> pair in columnKeyValuePairs)
            {
                fnd.ClearFormatting();
                fnd.Replacement.ClearFormatting();
                fnd.Forward = true;
                fnd.Wrap = Microsoft.Office.Interop.Word.WdFindWrap.wdFindContinue;

                fnd.Text = pair.Key;
                fnd.Replacement.Text = pair.Value;
                
                fnd.Execute(Replace: Microsoft.Office.Interop.Word.WdReplace.wdReplaceAll);
            }

            docObj.Save();
            docObj.Close(ref missing, ref missing, ref missing);
        }

        private void DisableButtons() 
        {
            button1.Enabled = false;
            buttonCSV.Enabled = false;
            buttonDocTemplate.Enabled = false;
            buttonTargetFolder.Enabled = false;
            toolStripProgressBar1.Visible = true;
        }

        private void EnableButtons() 
        {
            button1.Enabled = true;
            buttonCSV.Enabled = true;
            buttonDocTemplate.Enabled = true;
            buttonTargetFolder.Enabled = true;
            toolStripProgressBar1.Visible = false;
        }
    }
}
