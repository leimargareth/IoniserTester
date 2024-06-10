using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;

namespace IoniserTester
{
    public partial class datatable : Form
    {
        private System.Windows.Forms.Timer timer;
        public datatable()
        {
            InitializeComponent();

        }

        private void button1_Click(object sender, EventArgs e)
        {
            int ff1 = int.Parse(f1.Text);
            int ff2 = int.Parse(f2.Text);
            int ff3 = int.Parse(f3.Text);
            int ff4 = int.Parse(f4.Text);
            int ff5 = int.Parse(f5.Text);
            int ss1 = int.Parse(s1.Text);
            int ss2 = int.Parse(s2.Text);
            int ss3 = int.Parse(s3.Text);
            int ss4 = int.Parse(s4.Text);
            int ss5 = int.Parse(s5.Text);
            int tt1 = int.Parse(t1.Text);
            int tt2 = int.Parse(t2.Text);
            int tt3 = int.Parse(t3.Text);
            int tt4 = int.Parse(t4.Text);
            int tt5 = int.Parse(t5.Text);

            if (ss1 >= ff1 && ss1 <= tt1 && ss2 >= ff2 && ss2 <= tt2 && ss3 >= ff3 && ss3 <= tt3 &&
                ss4 >= ff4 && ss4 <= tt4 && ss5 >= ff5 && ss5 <= tt5)
            {
                output.Text = "PASSED";
                output.BackColor = Color.Green;
            }
            else
            {
                output.Text = "FAILED";
                output.BackColor = Color.Red;
            }

            if (ss1 >= ff1 && ss1 <= tt1)
            {
                s1.BackColor = Color.Green;
            }
            else
            {
                s1.BackColor = Color.Red;
            }

            if (ss2 >= ff2 && ss2 <= tt2)
            {
                s2.BackColor = Color.Green;
            }
            else
            {
                s2.BackColor = Color.Red;
            }

            if (ss3 >= ff3 && ss3 <= tt3)
            {
                s3.BackColor = Color.Green;
            }
            else
            {
                s3.BackColor = Color.Red;
            }

            if (ss4 >= ff4 && ss4 <= tt4)
            {
                s4.BackColor = Color.Green;
            }
            else
            {
                s4.BackColor = Color.Red;
            }

            if (ss5 >= ff5 && ss5 <= tt5)
            {
                s5.BackColor = Color.Green;
            }
            else
            {
                s5.BackColor = Color.Red;
            }
        }
    }
}
