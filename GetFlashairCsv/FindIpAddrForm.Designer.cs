namespace GetFlashairCsv
{
    partial class FindFlashairForm
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
            ApplyButton = new Button();
            CloseButton = new Button();
            IpAddrLabel = new Label();
            TextLabel2 = new Label();
            MacAddrLabel = new Label();
            TextLabel3 = new Label();
            StatusLabel = new Label();
            TextLabel1 = new Label();
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
            // ApplyButton
            // 
            ApplyButton.Location = new Point(140, 76);
            ApplyButton.Name = "ApplyButton";
            ApplyButton.Size = new Size(133, 23);
            ApplyButton.TabIndex = 1;
            ApplyButton.Text = "FlashAirのURLへ反映";
            ApplyButton.UseVisualStyleBackColor = true;
            // 
            // CloseButton
            // 
            CloseButton.Location = new Point(281, 76);
            CloseButton.Name = "CloseButton";
            CloseButton.Size = new Size(75, 23);
            CloseButton.TabIndex = 2;
            CloseButton.Text = "閉じる";
            CloseButton.UseVisualStyleBackColor = true;
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
            // TextLabel2
            // 
            TextLabel2.AutoSize = true;
            TextLabel2.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            TextLabel2.Location = new Point(24, 24);
            TextLabel2.Name = "TextLabel2";
            TextLabel2.Size = new Size(52, 15);
            TextLabel2.TabIndex = 4;
            TextLabel2.Text = "IPアドレス";
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
            // TextLabel3
            // 
            TextLabel3.AutoSize = true;
            TextLabel3.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            TextLabel3.Location = new Point(24, 39);
            TextLabel3.Name = "TextLabel3";
            TextLabel3.Size = new Size(68, 15);
            TextLabel3.TabIndex = 4;
            TextLabel3.Text = "MACアドレス";
            // 
            // StatusLabel
            // 
            StatusLabel.AutoSize = true;
            StatusLabel.Font = new Font("Yu Gothic UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            StatusLabel.Location = new Point(24, 54);
            StatusLabel.Name = "StatusLabel";
            StatusLabel.Size = new Size(65, 19);
            StatusLabel.TabIndex = 4;
            StatusLabel.Text = "検索開始";
            // 
            // TextLabel1
            // 
            TextLabel1.AutoSize = true;
            TextLabel1.Location = new Point(24, 9);
            TextLabel1.Name = "TextLabel1";
            TextLabel1.Size = new Size(120, 15);
            TextLabel1.TabIndex = 5;
            TextLabel1.Text = "FlashAirのMACアドレス";
            // 
            // FindFlashairForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(368, 109);
            Controls.Add(TextLabel1);
            Controls.Add(StatusLabel);
            Controls.Add(MacAddrLabel);
            Controls.Add(TextLabel3);
            Controls.Add(TextLabel2);
            Controls.Add(IpAddrLabel);
            Controls.Add(CloseButton);
            Controls.Add(ApplyButton);
            Controls.Add(FlashairMacAddrLabel);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "FindFlashairForm";
            Text = "FlashAir検索";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        public Label FlashairMacAddrLabel;
        public Button ApplyButton;
        public Button CloseButton;
        public Label IpAddrLabel;
        private Label TextLabel2;
        public Label MacAddrLabel;
        private Label TextLabel3;
        public Label StatusLabel;
        private Label TextLabel1;
    }
}