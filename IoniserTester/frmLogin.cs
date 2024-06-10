using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using System.Runtime.InteropServices;
using Microsoft.Identity.Client.NativeInterop;
using System.Xml.Linq;

namespace IoniserTester
{
    public partial class frmLogin : Form
    {
        public static frmLogin instance;

        //private const string connectionString = "Data Source=DESKTOP-59DG91J\\SQLEXPRESS;Initial Catalog=IonizerTester; User ID=sa; Password=1234; Integrated Security=True";
        private const string connectionString = "Data Source=localhost;Initial Catalog=IonizerTester; User ID=sa; Password=1234; Integrated Security=True";
        
        public string AccountLevelFromLogin { get; set; }

        public frmLogin()
        {
            
            InitializeComponent();
            instance = this;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (textUsername.Text == "" || textPassword.Text == "")
            {
                MessageBox.Show("Please provide UserName and Password");
                return;
            }
            try
            {
                //Create SqlConnection
                SqlConnection con = new SqlConnection(connectionString);
                SqlCommand cmd = new SqlCommand("Select * from userAccount where UserName=@Username and PassWord=@Password", con);
                cmd.Parameters.AddWithValue("@Username", textUsername.Text);
                cmd.Parameters.AddWithValue("@Password", textPassword.Text);



                con.Open();
                SqlDataAdapter adapt = new SqlDataAdapter(cmd);
                DataSet ds = new DataSet();
                adapt.Fill(ds);
                con.Close();
                int count = ds.Tables[0].Rows.Count;
                //If count is equal to 1, than show frmMain form
                if (count == 1)
                {
                    con.Open();
                    SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        string fullname = reader["FirstName"].ToString();
                        string name = reader["AccLevel"].ToString();
                        if (name == "Enginneer")
                        {

                            //MessageBox.Show("Login Successful!");
                            this.Hide();
                            frmMain.instance.accInformation.Text = "Welcome back " + fullname + "!";
                            frmMain.instance.AccLvlAccess.Text = "Engineer";

                        }
                        else
                        {
                            //if (_frmMain != null)
                            //{
                            //    // Access the TabPage from Form1
                            //    TextBox tabPage4 = _frmMain.tabPage;
                            //    tabPage4.Visible = false;
                            //    MessageBox.Show("Login Successful!");
                            //    this.Hide();
                            //    frmMain fm = new frmMain();
                            //    fm.Show();
                            //    // Now you can work with tabPageFromForm1 as needed
                            //}
                            
                            //MessageBox.Show("Login Successful!");
                            this.Hide();
                            frmMain.instance.accInformation.Text = "Welcome back "+fullname+"!";
                            frmMain.instance.AccLvlAccess.Text = "User";
                            
                            

                        }
                        // Close the reader
                        reader.Close();

                        // Assign the retrieved data to the TextBox control
                        
                    }


                }
                else
                {
                    MessageBox.Show("Login Failed!");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }

}
