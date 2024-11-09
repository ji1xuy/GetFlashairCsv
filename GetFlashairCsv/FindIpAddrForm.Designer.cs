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
            FlashairMacAddrLabel = new Label();
            applyButton = new Button();
            closeButton = new Button();
            IpAddrLabel = new Label();
            text2Label = new Label();
            MacAddrLabel = new Label();
            text3Label = new Label();
            statusLabel = new Label();
            text1Label = new Label();
            SuspendLayout();
            // 
            // FlashairMacAddrLabel
            // 
            FlashairMacAddrLabel.AutoSize = true;
            FlashairMacAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            FlashairMacAddrLabel.Location = new Point(150, 9);
            FlashairMacAddrLabel.Name = "FlashairMacAddrLabel";
            FlashairMacAddrLabel.Size = new Size(12, 15);
            FlashairMacAddrLabel.TabIndex = 0;
            FlashairMacAddrLabel.Text = "-";
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
            // IpAddrLabel
            // 
            IpAddrLabel.AutoSize = true;
            IpAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            IpAddrLabel.Location = new Point(150, 24);
            IpAddrLabel.Name = "IpAddrLabel";
            IpAddrLabel.Size = new Size(12, 15);
            IpAddrLabel.TabIndex = 3;
            IpAddrLabel.Text = "-";
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
            // MacAddrLabel
            // 
            MacAddrLabel.AutoSize = true;
            MacAddrLabel.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            MacAddrLabel.Location = new Point(150, 39);
            MacAddrLabel.Name = "MacAddrLabel";
            MacAddrLabel.Size = new Size(12, 15);
            MacAddrLabel.TabIndex = 4;
            MacAddrLabel.Text = "-";
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
            Controls.Add(MacAddrLabel);
            Controls.Add(text3Label);
            Controls.Add(text2Label);
            Controls.Add(IpAddrLabel);
            Controls.Add(closeButton);
            Controls.Add(applyButton);
            Controls.Add(FlashairMacAddrLabel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FindIpAddrForm";
            Text = "IPアドレス検索";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label FlashairMacAddrLabel;
        public Button applyButton;
        public Button closeButton;
        public Label IpAddrLabel;
        private Label text2Label;
        public Label MacAddrLabel;
        private Label text3Label;
        public Label statusLabel;
        private Label text1Label;
    }
}