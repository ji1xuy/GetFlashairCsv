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
            textLabel2 = new Label();
            MacAddrLabel = new Label();
            textLabel3 = new Label();
            StatusLabel = new Label();
            textLabel1 = new Label();
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
            // textLabel2
            // 
            textLabel2.AutoSize = true;
            textLabel2.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            textLabel2.Location = new Point(24, 24);
            textLabel2.Name = "textLabel2";
            textLabel2.Size = new Size(52, 15);
            textLabel2.TabIndex = 4;
            textLabel2.Text = "IPアドレス";
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
            // textLabel3
            // 
            textLabel3.AutoSize = true;
            textLabel3.Font = new Font("Yu Gothic UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
            textLabel3.Location = new Point(24, 39);
            textLabel3.Name = "textLabel3";
            textLabel3.Size = new Size(68, 15);
            textLabel3.TabIndex = 4;
            textLabel3.Text = "MACアドレス";
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
            // textLabel1
            // 
            textLabel1.AutoSize = true;
            textLabel1.Location = new Point(24, 9);
            textLabel1.Name = "textLabel1";
            textLabel1.Size = new Size(120, 15);
            textLabel1.TabIndex = 5;
            textLabel1.Text = "FlashAirのMACアドレス";
            // 
            // FindFlashairForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(368, 109);
            Controls.Add(textLabel1);
            Controls.Add(StatusLabel);
            Controls.Add(MacAddrLabel);
            Controls.Add(textLabel3);
            Controls.Add(textLabel2);
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
        private Label textLabel2;
        public Label MacAddrLabel;
        private Label textLabel3;
        public Label StatusLabel;
        private Label textLabel1;
    }
}