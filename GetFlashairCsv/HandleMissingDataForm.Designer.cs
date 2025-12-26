namespace GetFlashairCsv
{
    partial class HandleMissingDataForm
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
            OKButton = new Button();
            CancelButton = new Button();
            SuspendLayout();
            // 
            // InformationLabel
            // 
            InformationLabel.AutoSize = true;
            InformationLabel.Location = new Point(32, 9);
            InformationLabel.Name = "InformationLabel";
            InformationLabel.Size = new Size(97, 15);
            InformationLabel.TabIndex = 0;
            InformationLabel.Text = "InformationLabel";
            // 
            // DontShowAgainCheckBox
            // 
            DontShowAgainCheckBox.AutoSize = true;
            DontShowAgainCheckBox.Location = new Point(32, 88);
            DontShowAgainCheckBox.Name = "DontShowAgainCheckBox";
            DontShowAgainCheckBox.Size = new Size(173, 19);
            DontShowAgainCheckBox.TabIndex = 3;
            DontShowAgainCheckBox.Text = "今後このメッセージを表示しない";
            DontShowAgainCheckBox.UseVisualStyleBackColor = true;
            // 
            // OKButton
            // 
            OKButton.Location = new Point(213, 84);
            OKButton.Name = "OKButton";
            OKButton.Size = new Size(76, 25);
            OKButton.TabIndex = 1;
            OKButton.Text = "OK";
            OKButton.UseVisualStyleBackColor = true;
            // 
            // CancelButton
            // 
            CancelButton.Location = new Point(300, 84);
            CancelButton.Name = "CancelButton";
            CancelButton.Size = new Size(73, 26);
            CancelButton.TabIndex = 2;
            CancelButton.Text = "キャンセル";
            CancelButton.UseVisualStyleBackColor = true;
            // 
            // HandleMissingDataForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(391, 114);
            Controls.Add(CancelButton);
            Controls.Add(OKButton);
            Controls.Add(DontShowAgainCheckBox);
            Controls.Add(InformationLabel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "HandleMissingDataForm";
            Text = "確認";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label InformationLabel;
        public CheckBox DontShowAgainCheckBox;
        public Button OKButton;
        public Button CancelButton;
    }
}