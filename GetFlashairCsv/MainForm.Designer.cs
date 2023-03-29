using System.Text;

namespace GetFlashairCsv
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.FlashairUrlTextBox = new System.Windows.Forms.TextBox();
            this.WriteInifileButton = new System.Windows.Forms.Button();
            this.CsvFileListBox = new System.Windows.Forms.ListBox();
            this.UpdateCsvFileListButton = new System.Windows.Forms.Button();
            this.WriteExcelButton = new System.Windows.Forms.Button();
            this.ExcelFileNameLabel = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.OpenExcelButton = new System.Windows.Forms.Button();
            this.CloseButton = new System.Windows.Forms.Button();
            this.ExcelLastDataLabel = new System.Windows.Forms.Label();
            this.ClockLabel = new System.Windows.Forms.Label();
            this.CsvFileNameLabel = new System.Windows.Forms.Label();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.EdgeRadioButton = new System.Windows.Forms.RadioButton();
            this.ChromeRadioButton = new System.Windows.Forms.RadioButton();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.SuspendLayout();
            // 
            // FlashairUrlTextBox
            // 
            this.FlashairUrlTextBox.Font = new System.Drawing.Font("Yu Gothic UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.FlashairUrlTextBox.Location = new System.Drawing.Point(9, 22);
            this.FlashairUrlTextBox.Name = "FlashairUrlTextBox";
            this.FlashairUrlTextBox.Size = new System.Drawing.Size(230, 23);
            this.FlashairUrlTextBox.TabIndex = 0;
            // 
            // WriteInifileButton
            // 
            this.WriteInifileButton.Location = new System.Drawing.Point(257, 11);
            this.WriteInifileButton.Name = "WriteInifileButton";
            this.WriteInifileButton.Size = new System.Drawing.Size(47, 43);
            this.WriteInifileButton.TabIndex = 1;
            this.WriteInifileButton.Text = "設定\r\n保存";
            this.WriteInifileButton.UseVisualStyleBackColor = true;
            this.WriteInifileButton.Click += new System.EventHandler(this.WriteInifileButton_Click);
            // 
            // CsvFileListBox
            // 
            this.CsvFileListBox.Font = new System.Drawing.Font("ＭＳ ゴシック", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.CsvFileListBox.FormattingEnabled = true;
            this.CsvFileListBox.Location = new System.Drawing.Point(82, 27);
            this.CsvFileListBox.Name = "CsvFileListBox";
            this.CsvFileListBox.ScrollAlwaysVisible = true;
            this.CsvFileListBox.Size = new System.Drawing.Size(216, 108);
            this.CsvFileListBox.TabIndex = 3;
            this.CsvFileListBox.SelectedIndexChanged += new System.EventHandler(this.CsvFileListBox_SelectedIndexChanged);
            // 
            // UpdateCsvFileListButton
            // 
            this.UpdateCsvFileListButton.Location = new System.Drawing.Point(14, 75);
            this.UpdateCsvFileListButton.Name = "UpdateCsvFileListButton";
            this.UpdateCsvFileListButton.Size = new System.Drawing.Size(49, 43);
            this.UpdateCsvFileListButton.TabIndex = 2;
            this.UpdateCsvFileListButton.Text = "リスト\r\n更新";
            this.UpdateCsvFileListButton.UseVisualStyleBackColor = true;
            this.UpdateCsvFileListButton.Click += new System.EventHandler(this.UpdateCsvFileListButton_Click);
            // 
            // WriteExcelButton
            // 
            this.WriteExcelButton.Location = new System.Drawing.Point(130, 14);
            this.WriteExcelButton.Name = "WriteExcelButton";
            this.WriteExcelButton.Size = new System.Drawing.Size(172, 43);
            this.WriteExcelButton.TabIndex = 0;
            this.WriteExcelButton.Text = "CSVファイルをダウンロードして\r\nExcelファイルに書き込む";
            this.WriteExcelButton.UseVisualStyleBackColor = true;
            this.WriteExcelButton.Click += new System.EventHandler(this.WriteExcelButton_Click);
            // 
            // ExcelFileNameLabel
            // 
            this.ExcelFileNameLabel.Location = new System.Drawing.Point(11, 17);
            this.ExcelFileNameLabel.Name = "ExcelFileNameLabel";
            this.ExcelFileNameLabel.Size = new System.Drawing.Size(285, 49);
            this.ExcelFileNameLabel.TabIndex = 0;
            // 
            // label1
            // 
            this.label1.Font = new System.Drawing.Font("Yu Gothic UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(12, 66);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(46, 15);
            this.label1.TabIndex = 1;
            this.label1.Text = "最終行:";
            // 
            // OpenExcelButton
            // 
            this.OpenExcelButton.Location = new System.Drawing.Point(19, 365);
            this.OpenExcelButton.Name = "OpenExcelButton";
            this.OpenExcelButton.Size = new System.Drawing.Size(152, 34);
            this.OpenExcelButton.TabIndex = 1;
            this.OpenExcelButton.Text = "Excelファイルを開く";
            this.OpenExcelButton.UseVisualStyleBackColor = true;
            this.OpenExcelButton.Click += new System.EventHandler(this.OpenExcelButton_Click);
            // 
            // CloseButton
            // 
            this.CloseButton.Location = new System.Drawing.Point(247, 365);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(66, 34);
            this.CloseButton.TabIndex = 2;
            this.CloseButton.Text = "終了";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // ExcelLastDataLabel
            // 
            this.ExcelLastDataLabel.Font = new System.Drawing.Font("Yu Gothic UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.ExcelLastDataLabel.Location = new System.Drawing.Point(59, 66);
            this.ExcelLastDataLabel.Name = "ExcelLastDataLabel";
            this.ExcelLastDataLabel.Size = new System.Drawing.Size(237, 15);
            this.ExcelLastDataLabel.TabIndex = 2;
            // 
            // ClockLabel
            // 
            this.ClockLabel.Font = new System.Drawing.Font("Yu Gothic UI", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.ClockLabel.Location = new System.Drawing.Point(19, 402);
            this.ClockLabel.Name = "ClockLabel";
            this.ClockLabel.Size = new System.Drawing.Size(208, 23);
            this.ClockLabel.TabIndex = 3;
            // 
            // CsvFileNameLabel
            // 
            this.CsvFileNameLabel.Location = new System.Drawing.Point(12, 29);
            this.CsvFileNameLabel.Name = "CsvFileNameLabel";
            this.CsvFileNameLabel.Size = new System.Drawing.Size(113, 15);
            this.CsvFileNameLabel.TabIndex = 1;
            // 
            // timer1
            // 
            this.timer1.Enabled = true;
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CsvFileNameLabel);
            this.groupBox1.Controls.Add(this.WriteExcelButton);
            this.groupBox1.Location = new System.Drawing.Point(11, 212);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(309, 60);
            this.groupBox1.TabIndex = 6;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "CSVファイル";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.ExcelFileNameLabel);
            this.groupBox2.Controls.Add(this.ExcelLastDataLabel);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(11, 275);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(309, 84);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Excelファイル";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.EdgeRadioButton);
            this.groupBox3.Controls.Add(this.ChromeRadioButton);
            this.groupBox3.Controls.Add(this.UpdateCsvFileListButton);
            this.groupBox3.Controls.Add(this.CsvFileListBox);
            this.groupBox3.Location = new System.Drawing.Point(9, 63);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(310, 143);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "FlashAirに格納されているCSVファイル";
            // 
            // EdgeRadioButton
            // 
            this.EdgeRadioButton.AutoSize = true;
            this.EdgeRadioButton.Location = new System.Drawing.Point(10, 43);
            this.EdgeRadioButton.Name = "EdgeRadioButton";
            this.EdgeRadioButton.Size = new System.Drawing.Size(51, 19);
            this.EdgeRadioButton.TabIndex = 1;
            this.EdgeRadioButton.Text = "Edge";
            this.EdgeRadioButton.UseVisualStyleBackColor = true;
            // 
            // ChromeRadioButton
            // 
            this.ChromeRadioButton.AutoSize = true;
            this.ChromeRadioButton.Checked = true;
            this.ChromeRadioButton.Location = new System.Drawing.Point(10, 25);
            this.ChromeRadioButton.Name = "ChromeRadioButton";
            this.ChromeRadioButton.Size = new System.Drawing.Size(66, 19);
            this.ChromeRadioButton.TabIndex = 0;
            this.ChromeRadioButton.TabStop = true;
            this.ChromeRadioButton.Text = "Chrome";
            this.ChromeRadioButton.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.WriteInifileButton);
            this.groupBox4.Controls.Add(this.FlashairUrlTextBox);
            this.groupBox4.Location = new System.Drawing.Point(9, 3);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(311, 58);
            this.groupBox4.TabIndex = 4;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "FlashAirのURL";
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(330, 425);
            this.Controls.Add(this.groupBox4);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.ClockLabel);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.OpenExcelButton);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "MainForm";
            this.groupBox1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox4.ResumeLayout(false);
            this.groupBox4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private TextBox FlashairUrlTextBox;
        private Button WriteInifileButton;
        private ListBox CsvFileListBox;
        private Button UpdateCsvFileListButton;
        private Button WriteExcelButton;
        private Label ExcelFileNameLabel;
        private Label label1;
        private Button OpenExcelButton;
        private Button CloseButton;
        private Label ExcelLastDataLabel;
        private Label ClockLabel;
        private Label CsvFileNameLabel;
        private System.Windows.Forms.Timer timer1;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private GroupBox groupBox4;
        private RadioButton EdgeRadioButton;
        private RadioButton ChromeRadioButton;
    }
}