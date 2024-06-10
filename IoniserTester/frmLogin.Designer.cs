namespace IoniserTester
{
    partial class frmLogin
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
            panel1 = new Panel();
            panel2 = new Panel();
            label3 = new Label();
            button1 = new Button();
            label2 = new Label();
            label1 = new Label();
            textPassword = new TextBox();
            textUsername = new TextBox();
            panel1.SuspendLayout();
            panel2.SuspendLayout();
            SuspendLayout();
            // 
            // panel1
            // 
            panel1.BackColor = SystemColors.GradientActiveCaption;
            panel1.Controls.Add(panel2);
            panel1.Controls.Add(button1);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(textPassword);
            panel1.Controls.Add(textUsername);
            panel1.Location = new Point(20, 18);
            panel1.Margin = new Padding(3, 2, 3, 2);
            panel1.Name = "panel1";
            panel1.Size = new Size(657, 313);
            panel1.TabIndex = 70;
            // 
            // panel2
            // 
            panel2.BackColor = SystemColors.GradientInactiveCaption;
            panel2.Controls.Add(label3);
            panel2.Location = new Point(96, 26);
            panel2.Margin = new Padding(3, 2, 3, 2);
            panel2.Name = "panel2";
            panel2.Size = new Size(481, 58);
            panel2.TabIndex = 75;
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Font = new Font("Segoe UI Historic", 19.8000011F, FontStyle.Bold, GraphicsUnit.Point);
            label3.Location = new Point(15, 10);
            label3.Name = "label3";
            label3.Size = new Size(418, 37);
            label3.TabIndex = 73;
            label3.Text = "Welcome to X950 Ionizer tester";
            // 
            // button1
            // 
            button1.BackColor = Color.CornflowerBlue;
            button1.Font = new Font("Segoe UI", 10.2F, FontStyle.Bold, GraphicsUnit.Point);
            button1.Location = new Point(268, 252);
            button1.Margin = new Padding(3, 2, 3, 2);
            button1.Name = "button1";
            button1.Size = new Size(120, 40);
            button1.TabIndex = 74;
            button1.Text = "LogIn";
            button1.UseVisualStyleBackColor = false;
            button1.Click += button1_Click;
            // 
            // label2
            // 
            label2.AutoSize = true;
            label2.Font = new Font("Segoe UI Historic", 12F, FontStyle.Bold, GraphicsUnit.Point);
            label2.Location = new Point(284, 221);
            label2.Name = "label2";
            label2.Size = new Size(85, 21);
            label2.TabIndex = 73;
            label2.Text = "Password";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Font = new Font("Segoe UI Historic", 12F, FontStyle.Bold, GraphicsUnit.Point);
            label1.Location = new Point(279, 148);
            label1.Name = "label1";
            label1.Size = new Size(97, 21);
            label1.TabIndex = 72;
            label1.Text = "User Name";
            // 
            // textPassword
            // 
            textPassword.Font = new Font("Segoe UI", 13.8F, FontStyle.Regular, GraphicsUnit.Point);
            textPassword.Location = new Point(219, 177);
            textPassword.Margin = new Padding(3, 2, 3, 2);
            textPassword.MaximumSize = new Size(219, 250);
            textPassword.MinimumSize = new Size(79, 40);
            textPassword.Name = "textPassword";
            textPassword.PasswordChar = '•';
            textPassword.Size = new Size(219, 40);
            textPassword.TabIndex = 71;
            textPassword.TextAlign = HorizontalAlignment.Center;
            // 
            // textUsername
            // 
            textUsername.Font = new Font("Segoe UI", 13.8F, FontStyle.Regular, GraphicsUnit.Point);
            textUsername.Location = new Point(219, 106);
            textUsername.Margin = new Padding(3, 2, 3, 2);
            textUsername.MaximumSize = new Size(219, 250);
            textUsername.MinimumSize = new Size(79, 40);
            textUsername.Name = "textUsername";
            textUsername.Size = new Size(219, 40);
            textUsername.TabIndex = 70;
            textUsername.TextAlign = HorizontalAlignment.Center;
            // 
            // frmLogin
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            AutoScroll = true;
            AutoValidate = AutoValidate.EnableAllowFocusChange;
            BackColor = SystemColors.ActiveCaption;
            ClientSize = new Size(700, 351);
            Controls.Add(panel1);
            Margin = new Padding(3, 2, 3, 2);
            MaximizeBox = false;
            MinimizeBox = false;
            Name = "frmLogin";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "USER LOG IN";
            panel1.ResumeLayout(false);
            panel1.PerformLayout();
            panel2.ResumeLayout(false);
            panel2.PerformLayout();
            ResumeLayout(false);
        }

        #endregion

        private Panel panel1;
        private Label label2;
        private Label label1;
        private TextBox textPassword;
        private TextBox textUsername;
        private Panel panel2;
        private Label label3;
        private Button button1;
    }
}