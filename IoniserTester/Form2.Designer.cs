namespace IoniserTester
{
    partial class Form2
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
            components = new System.ComponentModel.Container();
            textBoxReceivedData = new TextBox();
            buttonClose = new Button();
            buttonOpen = new Button();
            Timer = new System.Windows.Forms.Timer(components);
            SuspendLayout();
            // 
            // textBoxReceivedData
            // 
            textBoxReceivedData.Location = new Point(82, 159);
            textBoxReceivedData.MaximumSize = new Size(178, 50);
            textBoxReceivedData.MinimumSize = new Size(400, 150);
            textBoxReceivedData.Multiline = true;
            textBoxReceivedData.Name = "textBoxReceivedData";
            textBoxReceivedData.Size = new Size(400, 150);
            textBoxReceivedData.TabIndex = 8;
            // 
            // buttonClose
            // 
            buttonClose.Location = new Point(227, 104);
            buttonClose.Name = "buttonClose";
            buttonClose.Size = new Size(113, 41);
            buttonClose.TabIndex = 7;
            buttonClose.Text = "CLOSE";
            buttonClose.UseVisualStyleBackColor = true;
            buttonClose.Click += buttonClose_Click;
            // 
            // buttonOpen
            // 
            buttonOpen.Location = new Point(227, 49);
            buttonOpen.Name = "buttonOpen";
            buttonOpen.Size = new Size(113, 40);
            buttonOpen.TabIndex = 6;
            buttonOpen.Text = "OPEN";
            buttonOpen.UseVisualStyleBackColor = true;
            buttonOpen.Click += buttonOpen_Click;
            // 
            // Timer
            // 
            Timer.Enabled = true;
            Timer.Interval = 1000;
            Timer.Tick += Timer_Tick;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(8F, 20F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(588, 364);
            Controls.Add(textBoxReceivedData);
            Controls.Add(buttonClose);
            Controls.Add(buttonOpen);
            Name = "Form2";
            Text = "Form2";
            Load += Form2_Load;
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private TextBox textBoxReceivedData;
        private Button buttonClose;
        private Button buttonOpen;
        private System.Windows.Forms.Timer Timer;
    }
}