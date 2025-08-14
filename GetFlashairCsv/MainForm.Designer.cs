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
        private void InitializeComponent() {
            components = new System.ComponentModel.Container();
            FlashairUrlTextBox = new TextBox();
            WriteInifileButton = new Button();
            CsvFileListBox = new ListBox();
            UpdateCsvFileListButton = new Button();
            WriteExcelButton = new Button();
            ExcelFileNameLabel = new Label();
            OpenExcelButton = new Button();
            CloseButton = new Button();
            ExcelLastDataLabel = new Label();
            ClockLabel = new Label();
            CsvFileNameLabel = new Label();
            ChromeRadioButton = new RadioButton();
            EdgeRadioButton = new RadioButton();
            timer1 = new System.Windows.Forms.Timer(components);
            groupBox1 = new GroupBox();
            groupBox2 = new GroupBox();
            label1 = new Label();
            groupBox3 = new GroupBox();
            groupBox4 = new GroupBox();
            FindFlashairButton = new Button();
            groupBox1.SuspendLayout();
            groupBox2.SuspendLayout();
            groupBox3.SuspendLayout();
            groupBox4.SuspendLayout();
            SuspendLayout();
            // 
            // FlashairUrlTextBox
            // 
            FlashairUrlTextBox.Font = new Font("Yu Gothic UI", 9F, FontStyle.Bold);
            FlashairUrlTextBox.Location = new Point(9, 22);
            FlashairUrlTextBox.Name = "FlashairUrlTextBox";
            FlashairUrlTextBox.Size = new Size(192, 23);
            FlashairUrlTextBox.TabIndex = 0;
            // 
            // WriteInifileButton
            // 
            WriteInifileButton.Location = new Point(257, 11);
            WriteInifileButton.Name = "WriteInifileButton";
            WriteInifileButton.Size = new Size(47, 43);
            WriteInifileButton.TabIndex = 1;
            WriteInifileButton.Text = "設定\r\n保存";
            WriteInifileButton.UseVisualStyleBackColor = true;
            WriteInifileButton.Click += WriteInifileButton_Click;
            // 
            // CsvFileListBox
            // 
            CsvFileListBox.Font = new Font("ＭＳ ゴシック", 9.75F, FontStyle.Bold);
            CsvFileListBox.FormattingEnabled = true;
            CsvFileListBox.Location = new Point(82, 27);
            CsvFileListBox.Name = "CsvFileListBox";
            CsvFileListBox.ScrollAlwaysVisible = true;
            CsvFileListBox.Size = new Size(216, 108);
            CsvFileListBox.TabIndex = 3;
            CsvFileListBox.SelectedIndexChanged += CsvFileListBox_SelectedIndexChanged;
            // 
            // UpdateCsvFileListButton
            // 
            UpdateCsvFileListButton.Location = new Point(14, 75);
            UpdateCsvFileListButton.Name = "UpdateCsvFileListButton";
            UpdateCsvFileListButton.Size = new Size(49, 43);
            UpdateCsvFileListButton.TabIndex = 2;
            UpdateCsvFileListButton.Text = "リスト\r\n更新";
            UpdateCsvFileListButton.UseVisualStyleBackColor = true;
            UpdateCsvFileListButton.Click += UpdateCsvFileListButton_Click;
            // 
            // WriteExcelButton
            // 
            WriteExcelButton.Location = new Point(130, 14);
            WriteExcelButton.Name = "WriteExcelButton";
            WriteExcelButton.Size = new Size(172, 43);
            WriteExcelButton.TabIndex = 0;
            WriteExcelButton.Text = "CSVファイルをダウンロードして\r\nExcelファイルに書き込む";
            WriteExcelButton.UseVisualStyleBackColor = true;
            WriteExcelButton.Click += WriteExcelButton_Click;
            // 
            // ExcelFileNameLabel
            // 
            ExcelFileNameLabel.Location = new Point(11, 17);
            ExcelFileNameLabel.Name = "ExcelFileNameLabel";
            ExcelFileNameLabel.Size = new Size(285, 70);
            ExcelFileNameLabel.TabIndex = 0;
            // 
            // OpenExcelButton
            // 
            OpenExcelButton.Location = new Point(19, 393);
            OpenExcelButton.Name = "OpenExcelButton";
            OpenExcelButton.Size = new Size(152, 34);
            OpenExcelButton.TabIndex = 1;
            OpenExcelButton.Text = "Excelファイルを開く";
            OpenExcelButton.UseVisualStyleBackColor = true;
            OpenExcelButton.Click += OpenExcelButton_Click;
            // 
            // CloseButton
            // 
            CloseButton.Location = new Point(247, 393);
            CloseButton.Name = "CloseButton";
            CloseButton.Size = new Size(66, 34);
            CloseButton.TabIndex = 2;
            CloseButton.Text = "終了";
            CloseButton.UseVisualStyleBackColor = true;
            CloseButton.Click += CloseButton_Click;
            // 
            // ExcelLastDataLabel
            // 
            ExcelLastDataLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Bold);
            ExcelLastDataLabel.Location = new Point(59, 89);
            ExcelLastDataLabel.Name = "ExcelLastDataLabel";
            ExcelLastDataLabel.Size = new Size(237, 15);
            ExcelLastDataLabel.TabIndex = 2;
            // 
            // ClockLabel
            // 
            ClockLabel.Font = new Font("Yu Gothic UI", 12F, FontStyle.Bold);
            ClockLabel.Location = new Point(11, 430);
            ClockLabel.Name = "ClockLabel";
            ClockLabel.Size = new Size(208, 23);
            ClockLabel.TabIndex = 3;
            // 
            // CsvFileNameLabel
            // 
            CsvFileNameLabel.Location = new Point(12, 29);
            CsvFileNameLabel.Name = "CsvFileNameLabel";
            CsvFileNameLabel.Size = new Size(113, 15);
            CsvFileNameLabel.TabIndex = 1;
            // 
            // ChromeRadioButton
            // 
            ChromeRadioButton.AutoSize = true;
            ChromeRadioButton.Checked = true;
            ChromeRadioButton.Location = new Point(10, 25);
            ChromeRadioButton.Name = "ChromeRadioButton";
            ChromeRadioButton.Size = new Size(66, 19);
            ChromeRadioButton.TabIndex = 0;
            ChromeRadioButton.TabStop = true;
            ChromeRadioButton.Text = "Chrome";
            ChromeRadioButton.UseVisualStyleBackColor = true;
            // 
            // EdgeRadioButton
            // 
            EdgeRadioButton.AutoSize = true;
            EdgeRadioButton.Location = new Point(10, 43);
            EdgeRadioButton.Name = "EdgeRadioButton";
            EdgeRadioButton.Size = new Size(51, 19);
            EdgeRadioButton.TabIndex = 1;
            EdgeRadioButton.Text = "Edge";
            EdgeRadioButton.UseVisualStyleBackColor = true;
            // 
            // timer1
            // 
            timer1.Enabled = true;
            timer1.Interval = 1000;
            timer1.Tick += timer1_Tick;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(CsvFileNameLabel);
            groupBox1.Controls.Add(WriteExcelButton);
            groupBox1.Location = new Point(11, 212);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(309, 60);
            groupBox1.TabIndex = 6;
            groupBox1.TabStop = false;
            groupBox1.Text = "CSVファイル";
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(label1);
            groupBox2.Controls.Add(ExcelFileNameLabel);
            groupBox2.Controls.Add(ExcelLastDataLabel);
            groupBox2.Location = new Point(11, 275);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(309, 112);
            groupBox2.TabIndex = 0;
            groupBox2.TabStop = false;
            groupBox2.Text = "Excelファイル";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(7, 89);
            label1.Name = "label1";
            label1.Size = new Size(46, 15);
            label1.TabIndex = 3;
            label1.Text = "最終行:";
            // 
            // groupBox3
            // 
            groupBox3.Controls.Add(EdgeRadioButton);
            groupBox3.Controls.Add(ChromeRadioButton);
            groupBox3.Controls.Add(UpdateCsvFileListButton);
            groupBox3.Controls.Add(CsvFileListBox);
            groupBox3.Location = new Point(9, 63);
            groupBox3.Name = "groupBox3";
            groupBox3.Size = new Size(310, 143);
            groupBox3.TabIndex = 5;
            groupBox3.TabStop = false;
            groupBox3.Text = "FlashAirに格納されているCSVファイル";
            // 
            // groupBox4
            // 
            groupBox4.Controls.Add(FindFlashairButton);
            groupBox4.Controls.Add(WriteInifileButton);
            groupBox4.Controls.Add(FlashairUrlTextBox);
            groupBox4.Location = new Point(9, 3);
            groupBox4.Name = "groupBox4";
            groupBox4.Size = new Size(311, 58);
            groupBox4.TabIndex = 4;
            groupBox4.TabStop = false;
            groupBox4.Text = "FlashAirのURL";
            // 
            // FindFlashairButton
            // 
            FindFlashairButton.Location = new Point(207, 11);
            FindFlashairButton.Name = "FindFlashairButton";
            FindFlashairButton.Size = new Size(44, 43);
            FindFlashairButton.TabIndex = 2;
            FindFlashairButton.Text = "検索";
            FindFlashairButton.UseVisualStyleBackColor = true;
            FindFlashairButton.Click += FindFlashairButton_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(330, 456);
            Controls.Add(groupBox4);
            Controls.Add(groupBox3);
            Controls.Add(groupBox2);
            Controls.Add(groupBox1);
            Controls.Add(ClockLabel);
            Controls.Add(CloseButton);
            Controls.Add(OpenExcelButton);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            Name = "MainForm";
            Shown += MainForm_Shown;
            groupBox1.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            groupBox3.ResumeLayout(false);
            groupBox3.PerformLayout();
            groupBox4.ResumeLayout(false);
            groupBox4.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        public TextBox FlashairUrlTextBox;
        private Button WriteInifileButton;
        private Button UpdateCsvFileListButton;
        private Button WriteExcelButton;
        private Button OpenExcelButton;
        private Button CloseButton;
        private Label ExcelFileNameLabel;
        private Label ExcelLastDataLabel;
        private Label ClockLabel;
        private Label CsvFileNameLabel;
        private ListBox CsvFileListBox;
        private System.Windows.Forms.Timer timer1;
        private GroupBox groupBox1;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private GroupBox groupBox4;
        private RadioButton EdgeRadioButton;
        private RadioButton ChromeRadioButton;
        private Label label1;
        private Button FindFlashairButton;
    }
}