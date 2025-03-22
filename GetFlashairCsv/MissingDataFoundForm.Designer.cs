namespace GetFlashairCsv
{
    partial class MissingDataFoundForm
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
        private void InitializeComponent() {
            InformationLabel = new Label();
            DontShowAgainCheckBox = new CheckBox();
            YesButton = new Button();
            NoButton = new Button();
            SuspendLayout();
            // 
            // InformationLabel
            // 
            InformationLabel.AutoSize = true;
            InformationLabel.Location = new Point(32, 9);
            InformationLabel.Name = "InformationLabel";
            InformationLabel.Size = new Size(213, 15);
            InformationLabel.TabIndex = 0;
            InformationLabel.Text = "CSVファイルにデータ欠落の可能性があります";
            // 
            // DontShowAgainCheckBox
            // 
            DontShowAgainCheckBox.AutoSize = true;
            DontShowAgainCheckBox.Location = new Point(32, 88);
            DontShowAgainCheckBox.Name = "DontShowAgainCheckBox";
            DontShowAgainCheckBox.Size = new Size(102, 19);
            DontShowAgainCheckBox.TabIndex = 1;
            DontShowAgainCheckBox.Text = "今後確認しない";
            DontShowAgainCheckBox.UseVisualStyleBackColor = true;
            // 
            // YesButton
            // 
            YesButton.Location = new Point(182, 84);
            YesButton.Name = "YesButton";
            YesButton.Size = new Size(76, 25);
            YesButton.TabIndex = 2;
            YesButton.Text = "はい";
            YesButton.UseVisualStyleBackColor = true;
            // 
            // NoButton
            // 
            NoButton.Location = new Point(283, 84);
            NoButton.Name = "NoButton";
            NoButton.Size = new Size(73, 26);
            NoButton.TabIndex = 3;
            NoButton.Text = "いいえ";
            NoButton.UseVisualStyleBackColor = true;
            // 
            // MissingDataFoundForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(368, 114);
            Controls.Add(NoButton);
            Controls.Add(YesButton);
            Controls.Add(DontShowAgainCheckBox);
            Controls.Add(InformationLabel);
            Name = "MissingDataFoundForm";
            Text = "確認";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label InformationLabel;
        public CheckBox DontShowAgainCheckBox;
        public Button YesButton;
        public Button NoButton;
    }
}