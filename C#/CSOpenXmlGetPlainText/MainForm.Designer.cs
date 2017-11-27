namespace CSUsingOpenXmlPlainText
{
    partial class MainForm
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnGetPlainText = new System.Windows.Forms.Button();
            this.btnOpen = new System.Windows.Forms.Button();
            this.tbxFile = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.rtbText = new System.Windows.Forms.RichTextBox();
            this.btnSaveas = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.txtExcelPath = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.button2 = new System.Windows.Forms.Button();
            this.txtResult = new System.Windows.Forms.TextBox();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.button2);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.txtExcelPath);
            this.groupBox1.Controls.Add(this.button1);
            this.groupBox1.Controls.Add(this.btnGetPlainText);
            this.groupBox1.Controls.Add(this.btnOpen);
            this.groupBox1.Controls.Add(this.tbxFile);
            this.groupBox1.Controls.Add(this.label1);
            this.groupBox1.Location = new System.Drawing.Point(17, 15);
            this.groupBox1.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Padding = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.groupBox1.Size = new System.Drawing.Size(623, 202);
            this.groupBox1.TabIndex = 1;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Microsoft Word Docx File";
            // 
            // btnGetPlainText
            // 
            this.btnGetPlainText.Location = new System.Drawing.Point(18, 169);
            this.btnGetPlainText.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnGetPlainText.Name = "btnGetPlainText";
            this.btnGetPlainText.Size = new System.Drawing.Size(156, 27);
            this.btnGetPlainText.TabIndex = 3;
            this.btnGetPlainText.Text = "Create Excel";
            this.btnGetPlainText.UseVisualStyleBackColor = true;
            this.btnGetPlainText.Click += new System.EventHandler(this.btnGetPlainText_Click);
            // 
            // btnOpen
            // 
            this.btnOpen.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnOpen.Location = new System.Drawing.Point(505, 40);
            this.btnOpen.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(100, 27);
            this.btnOpen.TabIndex = 2;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // tbxFile
            // 
            this.tbxFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.tbxFile.Location = new System.Drawing.Point(19, 43);
            this.tbxFile.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.tbxFile.Name = "tbxFile";
            this.tbxFile.Size = new System.Drawing.Size(477, 25);
            this.tbxFile.TabIndex = 1;
            this.tbxFile.Text = "C:\\Users\\User\\Desktop\\規格代號編碼原則\\52kd開關保險絲類及保安器材.docx";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 24);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Docx file:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(19, 352);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(127, 15);
            this.label2.TabIndex = 2;
            this.label2.Text = "Plain text in Word：";
            // 
            // rtbText
            // 
            this.rtbText.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.rtbText.BackColor = System.Drawing.SystemColors.Window;
            this.rtbText.Location = new System.Drawing.Point(22, 370);
            this.rtbText.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.rtbText.Name = "rtbText";
            this.rtbText.ReadOnly = true;
            this.rtbText.Size = new System.Drawing.Size(621, 40);
            this.rtbText.TabIndex = 3;
            this.rtbText.Text = "";
            // 
            // btnSaveas
            // 
            this.btnSaveas.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnSaveas.Location = new System.Drawing.Point(13, 322);
            this.btnSaveas.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.btnSaveas.Name = "btnSaveas";
            this.btnSaveas.Size = new System.Drawing.Size(125, 27);
            this.btnSaveas.TabIndex = 4;
            this.btnSaveas.Text = "Save as Text";
            this.btnSaveas.UseVisualStyleBackColor = true;
            this.btnSaveas.Click += new System.EventHandler(this.btnSaveas_Click);
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(181, 169);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(135, 27);
            this.button1.TabIndex = 4;
            this.button1.Text = "Reset Excel";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txtExcelPath
            // 
            this.txtExcelPath.Location = new System.Drawing.Point(18, 138);
            this.txtExcelPath.Name = "txtExcelPath";
            this.txtExcelPath.Size = new System.Drawing.Size(478, 25);
            this.txtExcelPath.TabIndex = 5;
            this.txtExcelPath.Text = "C:\\Users\\User\\Desktop\\規格代號編碼原則\\test.xlsx";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(16, 120);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(72, 15);
            this.label3.TabIndex = 6;
            this.label3.Text = "Excel Path:";
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(323, 169);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(122, 27);
            this.button2.TabIndex = 7;
            this.button2.Text = "Final Excel";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // txtResult
            // 
            this.txtResult.Location = new System.Drawing.Point(36, 224);
            this.txtResult.Name = "txtResult";
            this.txtResult.Size = new System.Drawing.Size(477, 25);
            this.txtResult.TabIndex = 5;
            this.txtResult.Text = "C:\\Users\\User\\Desktop\\規格代號編碼原則\\result.xlsx";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(656, 422);
            this.Controls.Add(this.txtResult);
            this.Controls.Add(this.btnSaveas);
            this.Controls.Add(this.rtbText);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.groupBox1);
            this.Margin = new System.Windows.Forms.Padding(4, 3, 4, 3);
            this.MinimumSize = new System.Drawing.Size(461, 339);
            this.Name = "MainForm";
            this.Text = "GetPlainTextFromWord";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button btnGetPlainText;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.TextBox tbxFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.RichTextBox rtbText;
        private System.Windows.Forms.Button btnSaveas;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtExcelPath;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.TextBox txtResult;
    }
}

