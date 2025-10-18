namespace GetFlashairCsv {
    partial class ProgressForm {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing) {
            if (disposing && (components != null)) {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent() {
            ProgressBar = new ProgressBar();
            CsvLabel = new Label();
            ExcelLabel = new Label();
            StatusLabel = new Label();
            AbortButton = new Button();
            SuspendLayout();
            // 
            // ProgressBar
            // 
            ProgressBar.Location = new Point(12, 44);
            ProgressBar.Name = "ProgressBar";
            ProgressBar.Size = new Size(282, 18);
            ProgressBar.TabIndex = 2;
            // 
            // CsvLabel
            // 
            CsvLabel.Location = new Point(12, 68);
            CsvLabel.Name = "CsvLabel";
            CsvLabel.Size = new Size(215, 16);
            CsvLabel.TabIndex = 3;
            // 
            // ExcelLabel
            // 
            ExcelLabel.Location = new Point(12, 88);
            ExcelLabel.Name = "ExcelLabel";
            ExcelLabel.Size = new Size(215, 14);
            ExcelLabel.TabIndex = 4;
            // 
            // StatusLabel
            // 
            StatusLabel.AutoSize = true;
            StatusLabel.Font = new Font("Yu Gothic UI", 11F, FontStyle.Bold);
            StatusLabel.Location = new Point(12, 13);
            StatusLabel.Name = "StatusLabel";
            StatusLabel.Size = new Size(66, 20);
            StatusLabel.TabIndex = 5;
            StatusLabel.Text = "処理中...";
            // 
            // AbortButton
            // 
            AbortButton.Enabled = false;
            AbortButton.Location = new Point(232, 76);
            AbortButton.Name = "AbortButton";
            AbortButton.Size = new Size(63, 28);
            AbortButton.TabIndex = 6;
            AbortButton.Text = "中断";
            AbortButton.UseVisualStyleBackColor = true;
            // 
            // ProgressForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(306, 112);
            Controls.Add(AbortButton);
            Controls.Add(StatusLabel);
            Controls.Add(ExcelLabel);
            Controls.Add(CsvLabel);
            Controls.Add(ProgressBar);
            Cursor = Cursors.WaitCursor;
            FormBorderStyle = FormBorderStyle.FixedSingle;
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "ProgressForm";
            StartPosition = FormStartPosition.Manual;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public ProgressBar ProgressBar;
        public Label CsvLabel;
        public Label ExcelLabel;
        public Label StatusLabel;
        public Button AbortButton;
    }
}