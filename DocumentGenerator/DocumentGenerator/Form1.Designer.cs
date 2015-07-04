namespace DocumentGenerator
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
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxCSVInput = new System.Windows.Forms.TextBox();
            this.textBoxDocTemplate = new System.Windows.Forms.TextBox();
            this.textBoxTargetFolder = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.buttonCSV = new System.Windows.Forms.Button();
            this.buttonDocTemplate = new System.Windows.Forms.Button();
            this.buttonTargetFolder = new System.Windows.Forms.Button();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.toolStripProgressBar1 = new System.Windows.Forms.ToolStripProgressBar();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 62);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(129, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "文档模板 (Doc/Docx)： ";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(13, 24);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(162, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "投资人列表文件 (CSV/TXT)： ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(13, 100);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(82, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "目标文件夹： ";
            // 
            // textBoxCSVInput
            // 
            this.textBoxCSVInput.Location = new System.Drawing.Point(171, 22);
            this.textBoxCSVInput.Name = "textBoxCSVInput";
            this.textBoxCSVInput.ReadOnly = true;
            this.textBoxCSVInput.Size = new System.Drawing.Size(389, 20);
            this.textBoxCSVInput.TabIndex = 3;
            // 
            // textBoxDocTemplate
            // 
            this.textBoxDocTemplate.Location = new System.Drawing.Point(171, 59);
            this.textBoxDocTemplate.Name = "textBoxDocTemplate";
            this.textBoxDocTemplate.ReadOnly = true;
            this.textBoxDocTemplate.Size = new System.Drawing.Size(389, 20);
            this.textBoxDocTemplate.TabIndex = 4;
            // 
            // textBoxTargetFolder
            // 
            this.textBoxTargetFolder.Location = new System.Drawing.Point(171, 93);
            this.textBoxTargetFolder.Name = "textBoxTargetFolder";
            this.textBoxTargetFolder.ReadOnly = true;
            this.textBoxTargetFolder.Size = new System.Drawing.Size(389, 20);
            this.textBoxTargetFolder.TabIndex = 5;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(160, 139);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(336, 23);
            this.button1.TabIndex = 6;
            this.button1.Text = "确认生成";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonCSV
            // 
            this.buttonCSV.Location = new System.Drawing.Point(568, 19);
            this.buttonCSV.Name = "buttonCSV";
            this.buttonCSV.Size = new System.Drawing.Size(87, 23);
            this.buttonCSV.TabIndex = 8;
            this.buttonCSV.Text = "选择文件...";
            this.buttonCSV.UseVisualStyleBackColor = true;
            this.buttonCSV.Click += new System.EventHandler(this.buttonCSV_Click);
            // 
            // buttonDocTemplate
            // 
            this.buttonDocTemplate.Location = new System.Drawing.Point(568, 58);
            this.buttonDocTemplate.Name = "buttonDocTemplate";
            this.buttonDocTemplate.Size = new System.Drawing.Size(87, 23);
            this.buttonDocTemplate.TabIndex = 9;
            this.buttonDocTemplate.Text = "选择文件...";
            this.buttonDocTemplate.UseVisualStyleBackColor = true;
            this.buttonDocTemplate.Click += new System.EventHandler(this.buttonDocTemplate_Click);
            // 
            // buttonTargetFolder
            // 
            this.buttonTargetFolder.Location = new System.Drawing.Point(568, 90);
            this.buttonTargetFolder.Name = "buttonTargetFolder";
            this.buttonTargetFolder.Size = new System.Drawing.Size(87, 23);
            this.buttonTargetFolder.TabIndex = 10;
            this.buttonTargetFolder.Text = "选择文件夹...";
            this.buttonTargetFolder.UseVisualStyleBackColor = true;
            this.buttonTargetFolder.Click += new System.EventHandler(this.buttonTargetFolder_Click);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel,
            this.toolStripProgressBar1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 204);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(681, 22);
            this.statusStrip1.TabIndex = 11;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(180, 17);
            this.toolStripStatusLabel.Text = "欢迎使用Word文档批量生成工具";
            // 
            // folderBrowserDialog1
            // 
            this.folderBrowserDialog1.Description = "请选择生成文档的目标文件夹";
            // 
            // toolStripProgressBar1
            // 
            this.toolStripProgressBar1.Name = "toolStripProgressBar1";
            this.toolStripProgressBar1.Size = new System.Drawing.Size(100, 16);
            this.toolStripProgressBar1.Visible = false;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(681, 226);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.buttonTargetFolder);
            this.Controls.Add(this.buttonDocTemplate);
            this.Controls.Add(this.buttonCSV);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.textBoxTargetFolder);
            this.Controls.Add(this.textBoxDocTemplate);
            this.Controls.Add(this.textBoxCSVInput);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "Form1";
            this.Text = "Word文档批量生成工具";
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxCSVInput;
        private System.Windows.Forms.TextBox textBoxDocTemplate;
        private System.Windows.Forms.TextBox textBoxTargetFolder;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button buttonCSV;
        private System.Windows.Forms.Button buttonDocTemplate;
        private System.Windows.Forms.Button buttonTargetFolder;
        private System.Windows.Forms.StatusStrip statusStrip1;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.ToolStripProgressBar toolStripProgressBar1;
    }
}

