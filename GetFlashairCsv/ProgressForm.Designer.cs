﻿namespace GetFlashairCsv {
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
            progressBar = new ProgressBar();
            CsvLabel = new Label();
            ExcelLabel = new Label();
            textLabel = new Label();
            abortButton = new Button();
            SuspendLayout();
            // 
            // progressBar
            // 
            progressBar.Location = new Point(12, 82);
            progressBar.Name = "progressBar";
            progressBar.Size = new Size(272, 18);
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
            textLabel.Size = new Size(66, 20);
            textLabel.TabIndex = 5;
            textLabel.Text = "処理中...";
            // 
            // abortButton
            // 
            abortButton.Cursor = Cursors.Arrow;
            abortButton.Enabled = false;
            abortButton.Location = new Point(221, 23);
            abortButton.Name = "abortButton";
            abortButton.Size = new Size(63, 28);
            abortButton.TabIndex = 6;
            abortButton.Text = "中断";
            abortButton.UseVisualStyleBackColor = true;
            // 
            // ProgressForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(296, 112);
            Controls.Add(abortButton);
            Controls.Add(textLabel);
            Controls.Add(ExcelLabel);
            Controls.Add(CsvLabel);
            Controls.Add(progressBar);
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

        //public Label label1;
        public ProgressBar progressBar;
        public Label CsvLabel;
        public Label ExcelLabel;
        public Label textLabel;
        public Button abortButton;
    }
}