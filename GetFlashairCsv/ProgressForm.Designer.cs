namespace GetFlashairCsv
{
    partial class ProgressForm
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
            progressBar = new ProgressBar();
            CsvLabel = new Label();
            ExcelLabel = new Label();
            textLabel = new Label();
            SuspendLayout();
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 82);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(215, 16);
            progressBar.TabIndex = 2;
            // 
            // CsvLabel
            // 
            CsvLabel.Location = new Point(12, 49);
            CsvLabel.Name = "CsvLabel";
            CsvLabel.Size = new Size(215, 16);
            CsvLabel.TabIndex = 3;
            // 
            // ExcelLabel
            // 
            ExcelLabel.Location = new Point(12, 65);
            ExcelLabel.Name = "ExcelLabel";
            ExcelLabel.Size = new Size(215, 14);
            ExcelLabel.TabIndex = 4;
            // 
            // textLabel
            // 
            textLabel.AutoSize = true;
            textLabel.Font = new Font("Yu Gothic UI", 11F, FontStyle.Bold, GraphicsUnit.Point);
            textLabel.Location = new Point(12, 25);
            textLabel.Name = "textLabel";
            textLabel.Size = new Size(92, 20);
            textLabel.TabIndex = 5;
            textLabel.Text = "お待ちください";
            // 
            // ProgressForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(246, 112);
            Controls.Add(textLabel);
            Controls.Add(ExcelLabel);
            Controls.Add(CsvLabel);
            Controls.Add(progressBar);
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ProgressForm";
            StartPosition = FormStartPosition.Manual;
            Text = "処理中...";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        //public Label label1;
        public ProgressBar progressBar;
        public Label CsvLabel;
        public Label ExcelLabel;
        private Label textLabel;
    }
}