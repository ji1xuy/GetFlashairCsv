namespace GetFlashairCsv
{
    partial class FindIpAddrForm
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
            flashairMacAddrLabel = new Label();
            applyButton = new Button();
            closeButton = new Button();
            ipAddrLabel = new Label();
            text2Label = new Label();
            macAddrLabel = new Label();
            text3Label = new Label();
            statusLabel = new Label();
            text1Label = new Label();
            SuspendLayout();
            // 
            // flashairMacAddrLabel
            // 
            flashairMacAddrLabel.AutoSize = true;
            flashairMacAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            flashairMacAddrLabel.Location = new Point(150, 9);
            flashairMacAddrLabel.Name = "flashairMacAddrLabel";
            flashairMacAddrLabel.Size = new Size(12, 15);
            flashairMacAddrLabel.TabIndex = 0;
            flashairMacAddrLabel.Text = "-";
            // 
            // applyButton
            // 
            applyButton.Location = new Point(140, 76);
            applyButton.Name = "applyButton";
            applyButton.Size = new Size(133, 23);
            applyButton.TabIndex = 1;
            applyButton.Text = "FlashAirのURLへ反映";
            applyButton.UseVisualStyleBackColor = true;
            // 
            // closeButton
            // 
            closeButton.Location = new Point(281, 76);
            closeButton.Name = "closeButton";
            closeButton.Size = new Size(75, 23);
            closeButton.TabIndex = 2;
            closeButton.Text = "閉じる";
            closeButton.UseVisualStyleBackColor = true;
            // 
            // ipAddrLabel
            // 
            ipAddrLabel.AutoSize = true;
            ipAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            ipAddrLabel.Location = new Point(150, 24);
            ipAddrLabel.Name = "ipAddrLabel";
            ipAddrLabel.Size = new Size(12, 15);
            ipAddrLabel.TabIndex = 3;
            ipAddrLabel.Text = "-";
            // 
            // text2Label
            // 
            text2Label.AutoSize = true;
            text2Label.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            text2Label.Location = new Point(24, 24);
            text2Label.Name = "text2Label";
            text2Label.Size = new Size(52, 15);
            text2Label.TabIndex = 4;
            text2Label.Text = "IPアドレス";
            // 
            // macAddrLabel
            // 
            macAddrLabel.AutoSize = true;
            macAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            macAddrLabel.Location = new Point(150, 39);
            macAddrLabel.Name = "macAddrLabel";
            macAddrLabel.Size = new Size(12, 15);
            macAddrLabel.TabIndex = 4;
            macAddrLabel.Text = "-";
            // 
            // text3Label
            // 
            text3Label.AutoSize = true;
            text3Label.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            text3Label.Location = new Point(24, 39);
            text3Label.Name = "text3Label";
            text3Label.Size = new Size(68, 15);
            text3Label.TabIndex = 4;
            text3Label.Text = "MACアドレス";
            // 
            // statusLabel
            // 
            statusLabel.AutoSize = true;
            statusLabel.Font = new Font("Yu Gothic UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            statusLabel.Location = new Point(24, 54);
            statusLabel.Name = "statusLabel";
            statusLabel.Size = new Size(65, 19);
            statusLabel.TabIndex = 4;
            statusLabel.Text = "検索開始";
            // 
            // text1Label
            // 
            text1Label.AutoSize = true;
            text1Label.Location = new Point(24, 9);
            text1Label.Name = "text1Label";
            text1Label.Size = new Size(120, 15);
            text1Label.TabIndex = 5;
            text1Label.Text = "FlashAirのMACアドレス";
            // 
            // FindIpAddrForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(368, 109);
            Controls.Add(text1Label);
            Controls.Add(statusLabel);
            Controls.Add(macAddrLabel);
            Controls.Add(text3Label);
            Controls.Add(text2Label);
            Controls.Add(ipAddrLabel);
            Controls.Add(closeButton);
            Controls.Add(applyButton);
            Controls.Add(flashairMacAddrLabel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FindIpAddrForm";
            Text = "IPアドレス検索";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label flashairMacAddrLabel;
        public Button applyButton;
        public Button closeButton;
        public Label ipAddrLabel;
        private Label text2Label;
        public Label macAddrLabel;
        private Label text3Label;
        public Label statusLabel;
        private Label text1Label;
    }
}