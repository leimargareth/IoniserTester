namespace IoniserTester;
using System.Security.Cryptography;
using System;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Text;
using System.Data;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using System.Runtime.Intrinsics.X86;
using Microsoft.VisualBasic.Logging;
using Excel = Microsoft.Office.Interop.Excel;
using GemBox.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using System.Configuration;
using System.Drawing;
using System.ComponentModel;
using System.Collections.Generic;
using System.Net.Sockets;
using Advantech.Adam;
using System.IO.Ports;
using System.Diagnostics;



//using DocumentFormat.OpenXml.Drawing;

public partial class frmMain : Form
{
    public static frmMain instance;
    public TextBox AccLvlAccess;
    public TextBox accInformation;

    private bool m_bStart;
    private AdamSocket adamModbus;
    private Adam6000Type m_Adam6000Type;
    private string m_szIP;
    private int m_iPort;
    private int m_iDoTotal, m_iDiTotal, m_iCount;

    public static SerialPort spDLY = new SerialPort();
    public const int RxDataSize = 140;
    public static byte[] rx_data = new byte[RxDataSize];
    public static int rx_len = 0;
    public static bool reading_OL = false;
    public static bool reading_valid = false;
    public static DateTime dtLastRxValidReading = DateTime.Now;
    public static string reading_string = "";
    public static double reading = 0.0;
    public static double ion_multiply_factor = 0.01;

    private System.Windows.Forms.Timer timer;
    private bool isFirstResultDisplayed = false;
    private int fetchCount = 0;
    private bool isTesting = false;

    public string AccountLevelTEST { get; set; }

    private Stopwatch stopwatch;

    //private const string connectionString = "Data Source=DESKTOP-59DG91J\\SQLEXPRESS;Initial Catalog=IonizerTester; User ID= INNOMFG-PBTS; Password= InnoMFG-PBTS; Integrated Security=True"; //Integrated Security=True
    //SqlConnection con = new SqlConnection(@"Data Source=DESKTOP-59DG91J\\SQLEXPRESS;Initial Catalog=IonizerTester; User ID= INNOMFG-PBTS; Password= InnoMFG-PBTS; Integrated Security=True"); // Integrated Security=True

    private const string connectionString = "Data Source=localhost;Initial Catalog=IonizerTester; User ID=sa; Password=1234; Integrated Security=True";
    SqlConnection con = new SqlConnection(@"Data Source=localhost;Initial Catalog=IonizerTester; User ID=sa; Password=1234; Integrated Security=True");
    SqlCommand cmd;

    public frmMain()
    {
        InitializeComponent();
        //OpenPort("COM4");
        ConnecttoAdam();
        instance = this;
        AccLvlAccess = accessAcc;
        accInformation = accInfobox;

        SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
        timer1.Interval = 1000; // Update every 1 second (1000 milliseconds)
        timer1.Tick += Timer1_Tick; // Attach event handler for the Timer tick
        timer1.Start();

        timer2.Interval = 1; // Update every 1 second (1000 milliseconds)
        timer2.Tick += timer2_Tick; // Attach event handler for the Timer tick
        timer2.Start();

        timer = new System.Windows.Forms.Timer();
        timer.Interval = 15000; // 15 seconds in milliseconds
        timer.Tick += new EventHandler(Timer_Tick);

        btnCh_Click(14, txtCh14);
        newTestButton.Enabled = false;
        newTestButton.BackColor = Color.LightGray;

        stopwatch = new Stopwatch();

    }

    public static void OpenPort(string portName)
    {
        if (!spDLY.IsOpen)
        {
            spDLY.PortName = portName;
            spDLY.BaudRate = 2400;
            spDLY.DataBits = 8;
            spDLY.Parity = Parity.None;
            spDLY.Open();
        }
        var taskDLYMainLoop = new Thread(MainLoop);
        taskDLYMainLoop.IsBackground = true;
        taskDLYMainLoop.Start();
    }

    //==== Process for DISPLAY of all DATA FROM SQL DATABASE ====//
    private void TabControlMain_SelectedIndexChanged(object sender, EventArgs e)
    {
        if (TabControlMain.SelectedTab == Passwords)
        {
            displayData();
        }
        else if (TabControlMain.SelectedTab == Reports)
        {
            displayResult();
        }
        else if (TabControlMain.SelectedTab == configMode)
        {
            configurationData();
        }

    }
    //==== End of process for DISPLAY of all DATA FROM SQL DATABASE ====//

    private void ExitButton_Click(object sender, EventArgs e)
    {
        DialogResult result = MessageBox.Show("Are you sure you want to exit?", "Confirmation", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        btnCh_Click(14, txtCh14);
        // Check the user's response
        if (result == DialogResult.Yes)
        {
            Close();
            frmLogin fm = new frmLogin();
            fm.Show();
        }
    }

    private void loginButton_Click(object sender, EventArgs e)
    {
        frmLogin frmLogin = new frmLogin();
        frmLogin.Show();
    }

    private void Timer1_Tick(object sender, EventArgs e)
    {

        lblDateTime.Text = DateTime.Now.ToString("dd MMM yyyy HH:mm:ss tt");

    }

    private void timer2_Tick(object sender, EventArgs e)
    {
        // Update the label with the elapsed time
        TimeSpan elapsed = stopwatch.Elapsed;
        testTimeTxtbox.Text = string.Format("{1:00}:{2:00}.{3:0}",
                                       elapsed.Hours, elapsed.Minutes, elapsed.Seconds, elapsed.Milliseconds / 100);

        if (accessAcc.Text == " ")
        {
            TabControlMain.Refresh();
            accessAcc.Text = "No Account";
            accInfobox.Text = "Welcome!";
            List<TabPage> tabsToRemove = new List<TabPage>();

            foreach (TabPage tab in TabControlMain.TabPages)
            {
                string tabPageText = tab.Text;

                if (tabPageText == "Configuration Mode")
                {
                    tabsToRemove.Add(tab);
                }
                else if (tabPageText == "Debug Mode")
                {
                    tabsToRemove.Add(tab);
                }
                else if (tabPageText == "Password")
                {
                    tabsToRemove.Add(tab);
                }
            }

            foreach (TabPage tab in tabsToRemove)
            {
                TabControlMain.TabPages.Remove(tab);
            }

            return;
        }

        else if (accessAcc.Text == "User")
        {
            logoutButton.Location = new System.Drawing.Point(1795, 70);
            logoutButton.Visible = true;
            logoutLabel.Location = new System.Drawing.Point(1800, 135);
            logoutLabel.Visible = true;
            loginButton.Visible = false;
            loginLabel.Visible = false;
            exitButton.Visible = false;
            exitLabel.Visible = false;

            TabControlMain.TabPages.Add(debugMode);
            TabControlMain.TabPages.Add(configMode);
            TabControlMain.TabPages.Add(Passwords);



            accessAcc.Text = "Login";
        }
    }


    private void passFailTxtBox_TextChanged(object sender, EventArgs e)
    {
        TextBox passFailTxtBox = sender as TextBox;
        if (passFailTxtBox != null)
        {

            if (passFailTxtBox.Text == "PASSED")
            {
                passfailPanel.BackColor = Color.Green;
                passFailTxtBox.BackColor = Color.Green;

            }
            else if (passFailTxtBox.Text == "FAILED")
            {
                passfailPanel.BackColor = Color.Red;
                passFailTxtBox.BackColor = Color.Red;

            }
            else
            {
                passfailPanel.BackColor = Color.White;
                passFailTxtBox.BackColor = Color.White;
            }
        }
    }


    // ========= START CODE FOR PASSWORD TAB ========== //

    //==== Function for INSERTION of NEW USER ACCOUNT ====//
    private void conButton_Click(object sender, EventArgs e)
    {

        string firstname = fname.Text;
        string lastname = lname.Text;
        string Position = position.Text;
        string Department = department.Text;
        string username = uname.Text;
        string password = accpass.Text;
        string accountlevel = accLev.Text;


        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            if (fname.Text == null || fname.Text == "")
            {
                MessageBox.Show("Please input your First Name!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (lname.Text == null || lname.Text == "")
            {
                MessageBox.Show("Please input your Last Name!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (position.Text == null || position.Text == "")
            {
                MessageBox.Show("Please input your Position!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (department.Text == null || department.Text == "")
            {
                MessageBox.Show("Please input your Department!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (uname.Text == null || uname.Text == "")
            {
                MessageBox.Show("Please input Username!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (accpass.Text == null || accpass.Text == "")
            {
                MessageBox.Show("Please input Password!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (rpassword.Text == null || rpassword.Text == "")
            {
                MessageBox.Show("Please re-type your Password!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (accLev.Text == "Account Level")
            {
                MessageBox.Show("Please input Account Level!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (accpass.Text == rpassword.Text)
            {
                string query = "INSERT INTO UserAccount (FirstName, LastName, Position, Department, UserName, PassWord, AccLevel) VALUES (@Firstname, @Lastname, @Position, @Department, @Username, @Password, @AccLvl)";

                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Firstname", firstname);
                command.Parameters.AddWithValue("@Lastname", lastname);
                command.Parameters.AddWithValue("@Position", Position);
                command.Parameters.AddWithValue("@Department", Department);
                command.Parameters.AddWithValue("@Username", username);
                command.Parameters.AddWithValue("@Password", password);
                command.Parameters.AddWithValue("@AccLvl", accountlevel);

                try
                {
                    connection.Open();
                    int rowsAffected = command.ExecuteNonQuery();
                    MessageBox.Show("Account successfully added", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    createLabel.Text = "Create New User Account";
                    fname.Clear();
                    lname.Clear();
                    position.Clear();
                    department.Clear();
                    uname.Clear();
                    accpass.Clear();
                    rpassword.Clear();
                    accLev.Text = "Account Level";
                    displayData();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Password is do not match!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }

    }

    //==== End of Function for INSERTION of NEW USER ACCOUNT ====//


    //==== Function for CALLING all DATA of User Account FROM SQL DATABASE ====//
    private void displayData()
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            SqlDataAdapter accData = new SqlDataAdapter("SELECT * FROM UserAccount", connection);
            DataTable dt = new DataTable();
            accData.Fill(dt);

            AccView.DataSource = dt;

            DataGridViewColumn centerFirstName = AccView.Columns["FirstName"];
            DataGridViewColumn centerLastName = AccView.Columns["LastName"];
            DataGridViewColumn centerPosition = AccView.Columns["Position"];
            DataGridViewColumn centerDepartment = AccView.Columns["Department"];
            DataGridViewColumn centerUsername = AccView.Columns["UserName"];
            DataGridViewColumn centerPassword = AccView.Columns["Password"];
            DataGridViewColumn centerAccLevel = AccView.Columns["AccLevel"];

            AccView.Columns["FirstName"].HeaderText = "First Name";
            AccView.Columns["LastName"].HeaderText = "Last Name";
            AccView.Columns["UserName"].HeaderText = "User";
            AccView.Columns["AccLevel"].HeaderText = "Account Level";

            //foreach (DataGridViewColumn column in configView.Columns)
            //{
            //    column.Width = 339;
            //}

            DataGridViewColumn columnToHide = AccView.Columns["ID"];
            if (columnToHide != null)
            {
                columnToHide.Visible = false;
            }
            else
            {
                MessageBox.Show("Column not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


            centerFirstName.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerFirstName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerFirstName.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerLastName.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerLastName.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerLastName.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerPosition.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPosition.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPosition.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerDepartment.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerDepartment.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerDepartment.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerUsername.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerUsername.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerUsername.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerPassword.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPassword.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPassword.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);

            centerAccLevel.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerAccLevel.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerAccLevel.HeaderCell.Style.Font = new Font(AccView.Font, FontStyle.Bold);
        }
    }
    //==== End of function for CALLING all DATA of User Account FROM SQL DATABASE ====//

    //==== Function for CALLING all DATA of Test Result Information FROM SQL DATABASE ====//
    private void displayResult()
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            //SqlDataAdapter resData = new SqlDataAdapter("SELECT * FROM ResultInfo", connection);
            SqlDataAdapter resData = new SqlDataAdapter("WITH NumberedRows AS (SELECT *, ROW_NUMBER() OVER (PARTITION BY [SerialNumber] ORDER BY " +
            "[SerialNumber]) AS RowNum FROM [IonizerTester].[dbo].[ResultInfo]) SELECT * FROM NumberedRows WHERE (RowNum - 1) % 100 = 0; ", connection);
            DataTable dt = new DataTable();
            resData.Fill(dt);

            resultView.DataSource = dt;

            resultView.Columns["PalletSN"].HeaderText = "Pallet SN";
            resultView.Columns["SerialNumber"].HeaderText = "Serial Number";
            resultView.Columns["PassedorFailed"].HeaderText = "   Result";
            resultView.Columns["DateTimeInfo"].HeaderText = "Date";
            resultView.Columns["TestTime"].HeaderText = "Test Time";

            DataGridViewColumn centerPalletSN = resultView.Columns["PalletSN"];
            DataGridViewColumn centerSerialNumber = resultView.Columns["SerialNumber"];
            DataGridViewColumn centerModel = resultView.Columns["Model"];
            DataGridViewColumn centerPassedorFailed = resultView.Columns["PassedorFailed"];
            DataGridViewColumn centerPass = resultView.Columns["Pass"];
            DataGridViewColumn centerReject = resultView.Columns["Reject"];
            DataGridViewColumn centerTotal = resultView.Columns["Total"];
            DataGridViewColumn centerYield = resultView.Columns["Yield"];
            DataGridViewColumn centerTestTime = resultView.Columns["TestTime"];
            DataGridViewColumn centerDateTimeInfo = resultView.Columns["DateTimeInfo"];

            DataGridViewColumn hideLow1 = resultView.Columns["Low1"];
            DataGridViewColumn hideLow2 = resultView.Columns["Low2"];
            DataGridViewColumn hideLow3 = resultView.Columns["Low3"];
            DataGridViewColumn hideLow4 = resultView.Columns["Low4"];
            DataGridViewColumn hideResult1 = resultView.Columns["Result1"];
            DataGridViewColumn hideResult2 = resultView.Columns["Result2"];
            DataGridViewColumn hideResult3 = resultView.Columns["Result3"];
            DataGridViewColumn hideResult4 = resultView.Columns["Result4"];
            DataGridViewColumn hideResult5 = resultView.Columns["Result5"];
            DataGridViewColumn hideHigh1 = resultView.Columns["High1"];
            DataGridViewColumn hideHigh2 = resultView.Columns["High2"];
            DataGridViewColumn hideHigh3 = resultView.Columns["High3"];
            DataGridViewColumn hideHigh4 = resultView.Columns["High4"];
            DataGridViewColumn hideHigh5 = resultView.Columns["High5"];
            //DataGridViewColumn hideTestTime = resultView.Columns["TestTime"];
            DataGridViewColumn hideUserName = resultView.Columns["UserName"];
            //DataGridViewColumn hidePass = resultView.Columns["Pass"];
            //DataGridViewColumn hideReject = resultView.Columns["Reject"];
            //DataGridViewColumn hideTotal = resultView.Columns["Total"];
            //DataGridViewColumn hideYield = resultView.Columns["Yield"];
            DataGridViewColumn hideRowNum = resultView.Columns["RowNum"];
            DataGridViewColumn hidePassedorFailed = resultView.Columns["PassedorFailed"];
            DataGridViewColumn hideSerialNumber = resultView.Columns["SerialNumber"];
            DataGridViewColumn hidePalletSN = resultView.Columns["PalletSN"];


            DataGridViewColumn columnToHide = resultView.Columns["ID"];
            if (columnToHide != null || hideLow1 != null || hideLow2 != null || hideLow3 != null || hideLow4 != null || hideResult1 != null || hideResult2 != null || hideResult3 != null
                || hideResult4 != null || hideResult5 != null || hideHigh1 != null || hideHigh2 != null || hideHigh3 != null || hideHigh4 != null || hideHigh5 != null
                || hideUserName != null || hideRowNum != null || hidePassedorFailed != null || hideSerialNumber != null || hidePalletSN != null)
            {
                hideLow1.Visible = false;
                hideLow2.Visible = false;
                hideLow3.Visible = false;
                hideLow4.Visible = false;
                hideResult1.Visible = false;
                hideResult2.Visible = false;
                hideResult3.Visible = false;
                hideResult4.Visible = false;
                hideResult5.Visible = false;
                hideHigh1.Visible = false;
                hideHigh2.Visible = false;
                hideHigh3.Visible = false;
                hideHigh4.Visible = false;
                hideHigh5.Visible = false;
                hideUserName.Visible = false;
                columnToHide.Visible = false;
                //hideTestTime.Visible = false;
                //hideYield.Visible = false;
                //hidePass.Visible = false;
                //hideReject.Visible = false;
                //hideTotal.Visible = false;
                hideRowNum.Visible = false;
                hidePassedorFailed.Visible = false;
                hidePalletSN.Visible = false;
                hideSerialNumber.Visible = false;
            }
            else
            {
                MessageBox.Show("Column not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            //resultView.Columns["SerialNumber"].Width = 100;
            //resultView.Columns["Model"].Width = 110;
            //resultView.Columns["PassedorFailed"].Width = 55;
            //resultView.Columns["DateTimeInfo"].Width = 150;

            centerTestTime.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerTestTime.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerTestTime.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerPalletSN.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPalletSN.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPalletSN.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerSerialNumber.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerSerialNumber.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerSerialNumber.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerModel.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerModel.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerModel.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerPassedorFailed.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPassedorFailed.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPassedorFailed.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerPass.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPass.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerPass.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerReject.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerReject.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerReject.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerTotal.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerTotal.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerYield.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerYield.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerYield.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

            centerDateTimeInfo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerDateTimeInfo.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerDateTimeInfo.HeaderCell.Style.Font = new Font(resultView.Font, FontStyle.Bold);

        }
    }

    private void setColumnWithResultView()
    {
        resultView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

    }
    //==== End of function for CALLING all DATA of Test Result Information FROM SQL DATABASE ====//


    //==== Function for LIVE SEARCHING a INFORMATION ====//
    private void textSearch_TextChanged(object sender, EventArgs e)
    {
        if (string.IsNullOrEmpty(textSearch.Text))
        {
            displayData();
        }

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            using (DataTable dt = new DataTable("IonizerTester"))
            {
                using (SqlCommand accData = new SqlCommand("SELECT * FROM UserAccount WHERE FirstName LIKE @Firstname or LastName LIKE @Lastname  ", connection))
                {
                    accData.Parameters.AddWithValue("FirstName", "%" + textSearch.Text + "%");
                    accData.Parameters.AddWithValue("LastName", "%" + textSearch.Text + "%");
                    SqlDataAdapter adapter = new SqlDataAdapter(accData);
                    adapter.Fill(dt);
                    AccView.DataSource = dt;
                }
            }
        }
    }
    //==== End of function for LIVE SEARCHING a INFORMATION ====//

    //==== Function for VIEWING the INFORMATION of DATA that you want to UPDATE ====//
    private void AccView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.RowIndex >= 0)
        {
            conButton.Visible = false;
            deleteButton.Visible = true;
            updateButton.Visible = true;
            cancelButton.Visible = true;
            deleteButton.Location = new System.Drawing.Point(293, 255);
            updateButton.Location = new System.Drawing.Point(108, 255);
            cancelButton.Location = new System.Drawing.Point(813, 295);

            DataGridViewRow row = this.AccView.Rows[e.RowIndex];
            createLabel.Text = row.Cells["FirstName"].Value.ToString() + " information details";
            IDsText.Text = row.Cells["ID"].Value.ToString();
            fname.Text = row.Cells["FirstName"].Value.ToString();
            lname.Text = row.Cells["LastName"].Value.ToString();
            position.Text = row.Cells["Position"].Value.ToString();
            department.Text = row.Cells["Department"].Value.ToString();
            uname.Text = row.Cells["UserName"].Value.ToString();
            accpass.Text = row.Cells["Password"].Value.ToString();
            rpassword.Text = row.Cells["Password"].Value.ToString();
            accLev.Text = row.Cells["AccLevel"].Value.ToString();

        }
        else
        {
            conButton.Visible = true;
            deleteButton.Visible = false;
            updateButton.Visible = false;
            cancelButton.Visible = false;
            IDsText.Clear();
            fname.Clear();
            lname.Clear();
            position.Clear();
            department.Clear();
            uname.Clear();
            accpass.Clear();
            rpassword.Clear();
            accLev.Text = "Account Level";

        }
    }
    //==== End of function for VIEWING the INFORMATION of DATA that you want to UPDATE ====//

    //==== Function for DELETE DATA in User Account ====//
    private void deleteButton_Click(object sender, EventArgs e)
    {
        string firstname = fname.Text;
        string lastname = lname.Text;
        string Position = position.Text;
        string Department = department.Text;
        string username = uname.Text;
        string password = accpass.Text;
        string accountlevel = accLev.Text;


        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            if (fname.Text != "" || lname.Text != "" || position.Text != "" || department.Text != "" || uname.Text != "" || accpass.Text != "" || rpassword.Text != "")
            {
                string query = "DELETE FROM UserAccount WHERE FirstName=@Firstname or LastName=@Lastname";

                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Firstname", firstname);
                command.Parameters.AddWithValue("@Lastname", lastname);
                command.Parameters.AddWithValue("@Position", Position);
                command.Parameters.AddWithValue("@Department", Department);
                command.Parameters.AddWithValue("@Username", username);
                command.Parameters.AddWithValue("@Password", password);
                command.Parameters.AddWithValue("@AccLvl", accountlevel);

                try
                {
                    connection.Open();
                    int rowsAffected = command.ExecuteNonQuery();
                    MessageBox.Show("Account successfully deleted!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    createLabel.Text = "Data Information";
                    fname.Clear();
                    lname.Clear();
                    position.Clear();
                    department.Clear();
                    uname.Clear();
                    accpass.Clear();
                    rpassword.Clear();
                    accLev.Text = "Account Level";
                    displayData();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please select a data you want to delete!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
    //==== End of function for DELETE DATA in User Account ====//

    //==== Function for UPDATING the SELECTED DATA INFORMATION ====//
    private void updateButton_Click(object sender, EventArgs e)
    {
        string firstname = fname.Text;
        string lastname = lname.Text;
        string Position = position.Text;
        string Department = department.Text;
        string username = uname.Text;
        string password = accpass.Text;
        string accountlevel = accLev.Text;
        string IDs = IDsText.Text;


        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            if (fname.Text != "" || lname.Text != "" || position.Text != "" || department.Text != "" || uname.Text != "" || accpass.Text != "" || rpassword.Text != "")
            {
                if (accpass.Text == rpassword.Text)
                {
                    string query = "UPDATE UserAccount SET FirstName = @Firstname, LastName = @Lastname, Position = @Position, Department = @Department, UserName = @Username, PassWord = @Password, AccLevel =  @AccLvl WHERE ID = @ID";

                    SqlCommand command = new SqlCommand(query, connection);
                    command.Parameters.AddWithValue("@ID", IDs);
                    command.Parameters.AddWithValue("@Firstname", firstname);
                    command.Parameters.AddWithValue("@Lastname", lastname);
                    command.Parameters.AddWithValue("@Position", Position);
                    command.Parameters.AddWithValue("@Department", Department);
                    command.Parameters.AddWithValue("@Username", username);
                    command.Parameters.AddWithValue("@Password", password);
                    command.Parameters.AddWithValue("@AccLvl", accountlevel);

                    try
                    {
                        connection.Open();
                        int rowsAffected = command.ExecuteNonQuery();
                        MessageBox.Show("Account successfully updated!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        AccView.Refresh();
                        displayData();

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
                else
                {
                    MessageBox.Show("Password do not match!", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Please select a data you want to update!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
    //==== End of function for UPDATING the SELECTED DATA INFORMATION ====//

    //==== Function for CANCELLATION of UPDATING INFORMATION ====//
    private void cancelButton_Click(object sender, EventArgs e)
    {
        conButton.Visible = true;
        updateButton.Visible = false;
        deleteButton.Visible = false;
        cancelButton.Visible = false;
        createLabel.Text = "Create New User Account";
        fname.Text = null;
        lname.Text = null;
        position.Text = null;
        department.Text = null;
        uname.Text = null;
        accpass.Text = null;
        rpassword.Text = null;
        accLev.Text = "Account Level";
    }
    //==== End of function for CANCELLATION of UPDATING INFORMATION ====//

    // ========= END CODE FOR PASSWORD TAB ========== //

    // ========= START CODE FOR REPORTS TAB ========== //
    private void searchResult_Click(object sender, EventArgs e)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            using (DataTable dt = new DataTable("IonizerTester"))
            {
                using (SqlCommand accData = new SqlCommand("SELECT * FROM ResultInfo WHERE (PalletSN LIKE @PalletSN OR SerialNumber LIKE @SerialNumber OR Model LIKE @Model) OR CAST(DateTimeInfo AS DATE) >= @StartDate AND CAST(DateTimeInfo AS DATE) <= @EndDate", connection))
                {
                    accData.Parameters.AddWithValue("@PalletSN", "%" + searchPallet.Text + "%");
                    accData.Parameters.AddWithValue("@SerialNumber", "%" + searchPallet.Text + "%");
                    accData.Parameters.AddWithValue("@Model", "%" + searchPallet.Text + "%");
                    accData.Parameters.AddWithValue("@StartDate", searchFrom.Value.Date);
                    accData.Parameters.AddWithValue("@EndDate", searchTo.Value.Date);


                    SqlDataAdapter adapter = new SqlDataAdapter(accData);
                    adapter.Fill(dt);
                    if (dt.Rows.Count == 0)
                    {
                        noResult.Visible = true;
                    }
                    else
                    {
                        noResult.Visible = false;
                        resultView.DataSource = dt;
                    }

                }
            }
        }
    }


    private void exportButton_Click(object sender, EventArgs e)
    {

        if (disModel.Text == "")//disPallet.Text == "" && disSerial.Text == ""
        {
            MessageBox.Show("Please select a data to export!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        else
        {
            // Load the Excel template file
            ExcelFile excelTemplate = ExcelFile.Load(@"C:\Users\BSS-PROGRAMMER\Documents\PROJECTS - JAMESON\DYSON\IoniserTester_ADVANCE\dyson.xlsx");
            ExcelWorksheet excelWorksheet = excelTemplate.Worksheets[0];

            // Insert specific data from text boxes into specific cells
            excelWorksheet.Cells["C4"].Value = disModel.Text;
            excelWorksheet.Cells["C5"].Value = disPallet.Text;
            excelWorksheet.Cells["C6"].Value = disSerial.Text;

            //excelWorksheet.Cells["C8"].Value = disTestTime.Text;

            //excelWorksheet.Cells[""].Value = disPassed.Text;
            //excelWorksheet.Cells[""].Value = disReject.Text;
            //excelWorksheet.Cells[""].Value = disTotal.Text;
            //excelWorksheet.Cells[""].Value = disYield.Text;

            excelWorksheet.Cells["C9"].Value = displayLow1.Text;
            excelWorksheet.Cells["C10"].Value = displayLow2.Text;
            excelWorksheet.Cells["C11"].Value = displayLow3.Text;
            excelWorksheet.Cells["C13"].Value = displayLow4.Text;

            excelWorksheet.Cells["D9"].Value = displayResult1.Text;
            excelWorksheet.Cells["D10"].Value = displayResult2.Text;
            excelWorksheet.Cells["D11"].Value = displayResult3.Text;
            excelWorksheet.Cells["D12"].Value = displayResult4.Text;
            excelWorksheet.Cells["D13"].Value = displayResult5.Text;

            excelWorksheet.Cells["E9"].Value = displayHigh1.Text;
            excelWorksheet.Cells["E10"].Value = displayHigh2.Text;
            excelWorksheet.Cells["E11"].Value = displayHigh3.Text;
            excelWorksheet.Cells["E12"].Value = displayHigh4.Text;
            excelWorksheet.Cells["E13"].Value = displayHigh5.Text;

            excelWorksheet.Cells["F9"].Value = displayPassFail.Text;
            excelWorksheet.Cells["F10"].Value = displayPassFail.Text;
            excelWorksheet.Cells["F11"].Value = displayPassFail.Text;
            excelWorksheet.Cells["F12"].Value = displayPassFail.Text;
            excelWorksheet.Cells["F13"].Value = displayPassFail.Text;

            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            // Define the file name
            string fileName = disModel.Text + "-" + disPallet.Text + timestamp + ".xlsx"; // Set your desired file name

            // Specify the full file path
            string outputDirectory = @"C:\Users\BSS-PROGRAMMER\Documents\X590-Output\Per Test Result";
            string filePath = Path.Combine(outputDirectory, fileName); // Set your specific folder path

            // Save the modified Excel file to the specified file path
            excelTemplate.Save(filePath);

            // Show a message box indicating successful saving
            MessageBox.Show($"Successfully exported {fileName}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

        }
    }



    private void resultView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {

        if (e.RowIndex >= 0) // Ensure the click is not on the header row
        {
            // Check if the row is blank
            bool isBlankRow = true;
            for (int i = 0; i < resultView.Columns.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(resultView.Rows[e.RowIndex].Cells[i].Value?.ToString()))
                {
                    isBlankRow = false;
                    break;
                }
            }

            if (isBlankRow)
            {
                MessageBox.Show("Invalid", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                // Existing code for handling non-blank rows
                exportLotButton.Visible = true;
                backToLot.Visible = true;
                resultView.Visible = false;
                perTestView.Visible = true;
                noResult.Visible = true;
                DataResultPanel.Visible = true;

                DataGridViewRow row = this.resultView.Rows[e.RowIndex];
                lotBox.Text = row.Cells["SerialNumber"].Value.ToString();

                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "SELECT * FROM ResultInfo WHERE SerialNumber LIKE @SerialNumber";
                    using (SqlCommand cmd = new SqlCommand(query, connection))
                    {
                        // Use the parameter @SerialNumber
                        cmd.Parameters.AddWithValue("@SerialNumber", lotBox.Text);

                        using (SqlDataAdapter resData = new SqlDataAdapter(cmd))
                        {
                            DataTable dt = new DataTable();
                            resData.Fill(dt);
                            perTestView.DataSource = dt;

                            perTestView.Columns["PalletSN"].HeaderText = "Pallet SN";
                            perTestView.Columns["SerialNumber"].HeaderText = "Serial Number";
                            perTestView.Columns["PassedorFailed"].HeaderText = "   Result";
                            perTestView.Columns["DateTimeInfo"].HeaderText = "Date";

                            DataGridViewColumn centerPalletSN = perTestView.Columns["PalletSN"];
                            DataGridViewColumn centerSerialNumber = perTestView.Columns["SerialNumber"];
                            DataGridViewColumn centerModel = perTestView.Columns["Model"];
                            DataGridViewColumn centerPassedorFailed = perTestView.Columns["PassedorFailed"];
                            DataGridViewColumn centerPass = perTestView.Columns["Pass"];
                            DataGridViewColumn centerReject = perTestView.Columns["Reject"];
                            DataGridViewColumn centerTotal = perTestView.Columns["Total"];
                            DataGridViewColumn centerYield = perTestView.Columns["Yield"];
                            DataGridViewColumn centerDateTimeInfo = perTestView.Columns["DateTimeInfo"];

                            DataGridViewColumn hideLow1 = perTestView.Columns["Low1"];
                            DataGridViewColumn hideLow2 = perTestView.Columns["Low2"];
                            DataGridViewColumn hideLow3 = perTestView.Columns["Low3"];
                            DataGridViewColumn hideLow4 = perTestView.Columns["Low4"];
                            DataGridViewColumn hideResult1 = perTestView.Columns["Result1"];
                            DataGridViewColumn hideResult2 = perTestView.Columns["Result2"];
                            DataGridViewColumn hideResult3 = perTestView.Columns["Result3"];
                            DataGridViewColumn hideResult4 = perTestView.Columns["Result4"];
                            DataGridViewColumn hideResult5 = perTestView.Columns["Result5"];
                            DataGridViewColumn hideHigh1 = perTestView.Columns["High1"];
                            DataGridViewColumn hideHigh2 = perTestView.Columns["High2"];
                            DataGridViewColumn hideHigh3 = perTestView.Columns["High3"];
                            DataGridViewColumn hideHigh4 = perTestView.Columns["High4"];
                            DataGridViewColumn hideHigh5 = perTestView.Columns["High5"];
                            DataGridViewColumn hideTestTime = perTestView.Columns["TestTime"];
                            DataGridViewColumn hideUserName = perTestView.Columns["UserName"];
                            DataGridViewColumn hidePass = perTestView.Columns["Pass"];
                            DataGridViewColumn hideReject = perTestView.Columns["Reject"];
                            DataGridViewColumn hideTotal = perTestView.Columns["Total"];
                            DataGridViewColumn hideYield = perTestView.Columns["Yield"];
                            DataGridViewColumn hideRowNum = perTestView.Columns["RowNum"];

                            DataGridViewColumn columnToHide = perTestView.Columns["ID"];
                            if (columnToHide != null || hideLow1 != null || hideLow2 != null || hideLow3 != null || hideLow4 != null || hideResult1 != null || hideResult2 != null || hideResult3 != null
                                || hideResult4 != null || hideResult5 != null || hideHigh1 != null || hideHigh2 != null || hideHigh3 != null || hideHigh4 != null || hideHigh5 != null || hideTestTime != null
                                || hideUserName != null || hidePass != null || hideReject != null || hideTotal != null || hideYield != null || hideTotal != null)
                            {
                                hideLow1.Visible = false;
                                hideLow2.Visible = false;
                                hideLow3.Visible = false;
                                hideLow4.Visible = false;
                                hideResult1.Visible = false;
                                hideResult2.Visible = false;
                                hideResult3.Visible = false;
                                hideResult4.Visible = false;
                                hideResult5.Visible = false;
                                hideHigh1.Visible = false;
                                hideHigh2.Visible = false;
                                hideHigh3.Visible = false;
                                hideHigh4.Visible = false;
                                hideHigh5.Visible = false;
                                hideTestTime.Visible = false;
                                hideUserName.Visible = false;
                                columnToHide.Visible = false;
                                hideYield.Visible = false;
                                hidePass.Visible = false;
                                hideReject.Visible = false;
                                hideTotal.Visible = false;
                            }
                            else
                            {
                                MessageBox.Show("Column not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }

                            centerPalletSN.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPalletSN.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPalletSN.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerSerialNumber.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerSerialNumber.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerSerialNumber.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerModel.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerModel.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerModel.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerPassedorFailed.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPassedorFailed.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPassedorFailed.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerPass.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPass.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerPass.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerReject.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerReject.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerReject.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerTotal.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerTotal.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerTotal.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerYield.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerYield.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerYield.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            centerDateTimeInfo.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerDateTimeInfo.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
                            centerDateTimeInfo.HeaderCell.Style.Font = new Font(perTestView.Font, FontStyle.Bold);

                            resultView.Visible = false;
                        }
                    }
                }
            }


        }
        //    DataGridViewRow row = this.resultView.Rows[e.RowIndex];
        //    dataResultLabel.Text = "Data Result for " + row.Cells["PalletSN"].Value.ToString();
        //    IDssText.Text = row.Cells["ID"].Value.ToString();
        //    disPassed.Text = row.Cells["Pass"].Value.ToString();
        //    disReject.Text = row.Cells["Reject"].Value.ToString();
        //    disTotal.Text = row.Cells["Total"].Value.ToString();
        //    disYield.Text = row.Cells["Yield"].Value.ToString();
        //    disModel.Text = row.Cells["Model"].Value.ToString();
        //    disPallet.Text = row.Cells["PalletSN"].Value.ToString();
        //    disSerial.Text = row.Cells["SerialNumber"].Value.ToString();
        //    disTestTime.Text = row.Cells["TestTime"].Value.ToString();
        //    displayLow1.Text = row.Cells["Low1"].Value.ToString();
        //    displayLow2.Text = row.Cells["Low2"].Value.ToString();
        //    displayLow3.Text = row.Cells["Low3"].Value.ToString();
        //    displayLow4.Text = row.Cells["Low4"].Value.ToString();
        //    displayResult1.Text = row.Cells["Result1"].Value.ToString();
        //    displayResult2.Text = row.Cells["Result2"].Value.ToString();
        //    displayResult3.Text = row.Cells["Result3"].Value.ToString();
        //    displayResult4.Text = row.Cells["Result4"].Value.ToString();
        //    displayResult5.Text = row.Cells["Result5"].Value.ToString();
        //    displayHigh1.Text = row.Cells["High1"].Value.ToString();
        //    displayHigh2.Text = row.Cells["High2"].Value.ToString();
        //    displayHigh3.Text = row.Cells["High3"].Value.ToString();
        //    displayHigh4.Text = row.Cells["High4"].Value.ToString();
        //    displayHigh5.Text = row.Cells["High5"].Value.ToString();
        //    displayPassFail.Text = row.Cells["PassedorFailed"].Value.ToString();

        //    if (disPallet.Text == "" && disSerial.Text == "")
        //    {
        //        dataResultLabel.Text = "Data Result";
        //    }

        //}
        //else
        //{
        //    IDssText.Clear();
        //    disPassed.Clear();
        //    disReject.Clear();
        //    disTotal.Clear();
        //    disYield.Clear();
        //    disModel.Clear();
        //    disPallet.Clear();
        //    disSerial.Clear();
        //    disTestTime.Clear();
        //    displayLow1.Clear();
        //    displayLow2.Clear();
        //    displayLow3.Clear();
        //    displayLow4.Clear();
        //    displayResult1.Clear();
        //    displayResult2.Clear();
        //    displayResult3.Clear();
        //    displayResult4.Clear();
        //    displayResult5.Clear();
        //    displayHigh1.Clear();
        //    displayHigh2.Clear();
        //    displayHigh3.Clear();
        //    displayHigh4.Clear();
        //    displayHigh5.Clear();
        //    displayPassFail.Clear();
        //}

    }

    // ========= END CODE FOR REPORTS TAB ========== //

    // ========= START CODE FOR TEST MODE TAB ========== //
    private void modelComboBox_TextChanged(object sender, EventArgs e)
    {
        SqlConnection con = new SqlConnection(connectionString);
        SqlCommand cmd = new SqlCommand("SELECT * FROM ConfigurationData WHERE ModelData = @Modeldata", con);
        cmd.Parameters.AddWithValue("@Modeldata", modelComboBox.SelectedItem.ToString());

        con.Open();
        SqlDataAdapter adapt = new SqlDataAdapter(cmd);
        DataSet ds = new DataSet();
        adapt.Fill(ds);
        con.Close();
        int count = ds.Tables[0].Rows.Count;
        if (count == 1)
        {
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            if (reader.Read())
            {
                lowTxtBox1.Text = reader["lowOne"].ToString();
                lowTxtBox2.Text = reader["lowTwo"].ToString();
                lowTxtBox3.Text = reader["lowThree"].ToString();
                lowTxtBox4.Text = reader["lowFour"].ToString();
                highTxtBox1.Text = reader["highOne"].ToString();
                highTxtBox2.Text = reader["highTwo"].ToString();
                highTxtBox3.Text = reader["highThree"].ToString();
                highTxtBox4.Text = reader["highFour"].ToString();
                highTxtBox5.Text = reader["highFive"].ToString();
                totalTxtBox.Text = reader["Total"].ToString();
                onGoingBox.Text = "READY";
                onGoingBox.BackColor = Color.Yellow;
                reader.Close();
            }
        }
        else
        {
            lowTxtBox1.Text = "0";
            lowTxtBox2.Text = "0";
            lowTxtBox3.Text = "0";
            lowTxtBox4.Text = "0";
            highTxtBox1.Text = "0";
            highTxtBox2.Text = "0";
            highTxtBox3.Text = "0";
            highTxtBox4.Text = "0";
            highTxtBox5.Text = "0";
            resultTxtBox1.Text = "0";
            resultTxtBox2.Text = "0";
            resultTxtBox3.Text = "0";
            resultTxtBox4.Text = "0";
            resultTxtBox5.Text = "0";
            resultTxtBox1.BackColor = System.Drawing.Color.White;
            resultTxtBox2.BackColor = System.Drawing.Color.White;
            resultTxtBox3.BackColor = System.Drawing.Color.White;
            resultTxtBox4.BackColor = System.Drawing.Color.White;
            resultTxtBox5.BackColor = System.Drawing.Color.White;
            onGoingBox.Text = "STAND BY";
            onGoingBox.BackColor = Color.White;
        }
    }


    private void frmMain_Load(object sender, EventArgs e)
    {
        setColumnWithResultView();
        SqlConnection con = new SqlConnection(connectionString);
        SqlCommand cmd = new SqlCommand("SELECT DISTINCT ModelData FROM ConfigurationData", con);

        try
        {
            con.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            modelComboBox.Items.Add("Select...");
            while (reader.Read())
            {
                modelComboBox.Items.Add(reader["ModelData"].ToString());
            }
            reader.Close();

            if (modelComboBox.Items.Count > 0)
            {
                modelComboBox.SelectedIndex = 0;
            }

        }
        catch (Exception ex)
        {
            MessageBox.Show("An error occurred: " + ex.Message);
        }
    }



    private void logoutButton_Click(object sender, EventArgs e)
    {
        loginButton.Visible = true;
        loginLabel.Visible = true;
        exitButton.Visible = true;
        exitLabel.Visible = true;
        logoutButton.Visible = false;
        logoutLabel.Visible = false;
        accessAcc.Text = " ";

    }


    // ========= START CODE FOR ADAM6050 ========== //
    private void communication()
    {
        m_bStart = false;           // the action stops at the beginning
                                    //m_szIP = "10.0.0.2";	// modbus slave IP address
                                    //m_iPort = 502; 				// modbus TCP port is 502

        m_szIP = txtip.Text.Trim();//*/	// modbus slave IP address
        m_iPort = 502;
        adamModbus = new AdamSocket();
        adamModbus.SetTimeout(1000, 1000, 1000); // set timeout for TCP

        m_Adam6000Type = Adam6000Type.Adam6050;

        if (m_Adam6000Type == Adam6000Type.Adam6050 ||
            m_Adam6000Type == Adam6000Type.Adam6050W)
            InitAdam6050();

    }

    protected void InitChannelItems(bool i_bVisable, bool i_bIsDI, ref int i_iDI, ref int i_iDO, ref Panel panelCh, ref Button btnCh)
    {
        int iCh;
        if (i_bVisable)
        {
            panelCh.Visible = true;
            iCh = i_iDI + i_iDO;
            if (i_bIsDI) // DI
            {
                btnCh.Text = "DI " + i_iDI.ToString();
                btnCh.Enabled = false;
                i_iDI++;
            }
            else // DO
            {
                btnCh.Text = "DO " + i_iDO.ToString();
                i_iDO++;
            }
        }
        else
            panelCh.Visible = false;
    }


    private void btncnnct_Click(object sender, EventArgs e)
    {
        if (m_bStart) // was started
        {
            communication();
            //panelDIO.Enabled = false;
            m_bStart = false;       // starting flag
            timer3.Enabled = false; // disable timer
            adamModbus.Disconnect(); // disconnect slave
            btncnnct.Text = "Start";
            //btnApplyWDT.Enabled = false;

            txtCh0.Text = "";
            txtCh1.Text = "";
            txtCh2.Text = "";
            txtCh3.Text = "";
            txtCh4.Text = "";
            txtCh5.Text = "";
            txtCh6.Text = "";
            txtCh7.Text = "";
            txtCh8.Text = "";
            txtCh9.Text = "";
            txtCh10.Text = "";
            txtCh11.Text = "";
            txtCh12.Text = "";
            txtCh13.Text = "";
            txtCh14.Text = "";
            txtCh15.Text = "";
            txtCh16.Text = "";
            txtCh17.Text = "";
            txtReadCount.Text = "";

            txtCh0.BackColor = Color.White;
            txtCh1.BackColor = Color.White;
            txtCh2.BackColor = Color.White;
            txtCh3.BackColor = Color.White;
            txtCh4.BackColor = Color.White;
            txtCh5.BackColor = Color.White;
            txtCh6.BackColor = Color.White;
            txtCh7.BackColor = Color.White;
            txtCh8.BackColor = Color.White;
            txtCh9.BackColor = Color.White;
            txtCh10.BackColor = Color.White;
            txtCh11.BackColor = Color.White;
            txtCh12.BackColor = Color.White;
            txtCh13.BackColor = Color.White;
            txtCh14.BackColor = Color.White;
            txtCh15.BackColor = Color.White;
            txtCh16.BackColor = Color.White;
            txtCh17.BackColor = Color.White;

        }
        else    // was stoped
        {
            communication();
            AdamDevice adamDevice = new AdamDevice();
            if (adamModbus.Connect(m_szIP, ProtocolType.Tcp, m_iPort))
            {
                //RefreshWDT();
                //panelDIO.Enabled = true;
                m_iCount = 0; // reset the reading counter
                timer3.Enabled = true; // enable timer
                btncnnct.Text = "Stop";
                m_bStart = true; // starting flag


            }
            else
                MessageBox.Show("Connect to " + m_szIP + " failed", "Error");
        }
    }

    private void ConnecttoAdam()
    {
        if (m_bStart) // was started
        {
            communication();
            //panelDIO.Enabled = false;
            m_bStart = false;       // starting flag
            timer3.Enabled = false; // disable timer
            adamModbus.Disconnect(); // disconnect slave
            btncnnct.Text = "Start";
            //btnApplyWDT.Enabled = false;

            txtCh0.Text = "";
            txtCh1.Text = "";
            txtCh2.Text = "";
            txtCh3.Text = "";
            txtCh4.Text = "";
            txtCh5.Text = "";
            txtCh6.Text = "";
            txtCh7.Text = "";
            txtCh8.Text = "";
            txtCh9.Text = "";
            txtCh10.Text = "";
            txtCh11.Text = "";
            txtCh12.Text = "";
            txtCh13.Text = "";
            txtCh14.Text = "";
            txtCh15.Text = "";
            txtCh16.Text = "";
            txtCh17.Text = "";
            txtReadCount.Text = "";

            txtCh0.BackColor = Color.White;
            txtCh1.BackColor = Color.White;
            txtCh2.BackColor = Color.White;
            txtCh3.BackColor = Color.White;
            txtCh4.BackColor = Color.White;
            txtCh5.BackColor = Color.White;
            txtCh6.BackColor = Color.White;
            txtCh7.BackColor = Color.White;
            txtCh8.BackColor = Color.White;
            txtCh9.BackColor = Color.White;
            txtCh10.BackColor = Color.White;
            txtCh11.BackColor = Color.White;
            txtCh12.BackColor = Color.White;
            txtCh13.BackColor = Color.White;
            txtCh14.BackColor = Color.White;
            txtCh15.BackColor = Color.White;
            txtCh16.BackColor = Color.White;
            txtCh17.BackColor = Color.White;

        }
        else    // was stoped
        {
            communication();
            AdamDevice adamDevice = new AdamDevice();
            if (adamModbus.Connect(m_szIP, ProtocolType.Tcp, m_iPort))
            {
                //RefreshWDT();
                //panelDIO.Enabled = true;
                m_iCount = 0; // reset the reading counter
                timer3.Enabled = true; // enable timer
                btncnnct.Text = "Stop";
                m_bStart = true; // starting flag


            }
            else
                MessageBox.Show("Connect to " + m_szIP + " failed", "Error");
        }
    }

    private void frmMain_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (m_bStart)
        {
            timer3.Enabled = false;
            adamModbus.Disconnect(); // disconnect slave
        }
    }

    private void RefreshDIO()
    {
        int iDiStart = 1, iDoStart = 17;
        int iChTotal;
        bool[] bDiData, bDoData, bData;

        if (m_Adam6000Type == Adam6000Type.Adam6055)
        {
            if (adamModbus.Modbus().ReadCoilStatus(iDiStart, m_iDiTotal, out bDiData))
            {
                iChTotal = m_iDiTotal;
                bData = new bool[iChTotal];
                Array.Copy(bDiData, 0, bData, 0, m_iDiTotal);
                if (iChTotal > 0)
                    txtCh0.Text = bData[0].ToString();
                txtCh0.BackColor = Color.Green;
                if (iChTotal > 1)
                    txtCh1.Text = bData[1].ToString();
                txtCh0.BackColor = Color.Red;
                if (iChTotal > 2)
                    txtCh2.Text = bData[2].ToString();
                if (iChTotal > 3)
                    txtCh3.Text = bData[3].ToString();
                if (iChTotal > 4)
                    txtCh4.Text = bData[4].ToString();
                if (iChTotal > 5)
                    txtCh5.Text = bData[5].ToString();
                if (iChTotal > 6)
                    txtCh6.Text = bData[6].ToString();
                if (iChTotal > 7)
                    txtCh7.Text = bData[7].ToString();
                if (iChTotal > 8)
                    txtCh8.Text = bData[8].ToString();
                if (iChTotal > 9)
                    txtCh9.Text = bData[9].ToString();
                if (iChTotal > 10)
                    txtCh10.Text = bData[10].ToString();
                if (iChTotal > 11)
                    txtCh11.Text = bData[11].ToString();
                if (iChTotal > 12)
                    txtCh12.Text = bData[12].ToString();
                if (bData[12] == true)
                {
                    txtCh12.Text = "Turn On";
                }
                else
                {
                    txtCh12.Text = "Turn Off";
                }
                if (iChTotal > 13)
                    txtCh13.Text = bData[13].ToString();
                if (iChTotal > 14)
                    txtCh14.Text = bData[14].ToString();
                if (iChTotal > 15)
                    txtCh15.Text = bData[15].ToString();
                if (iChTotal > 16)
                    txtCh16.Text = bData[16].ToString();
                if (iChTotal > 17)
                    txtCh17.Text = bData[17].ToString();
            }
            else
            {
                txtCh0.Text = "Fail";
                txtCh1.Text = "Fail";
                txtCh2.Text = "Fail";
                txtCh3.Text = "Fail";
                txtCh4.Text = "Fail";
                txtCh5.Text = "Fail";
                txtCh6.Text = "Fail";
                txtCh7.Text = "Fail";
                txtCh8.Text = "Fail";
                txtCh9.Text = "Fail";
                txtCh10.Text = "Fail";
                txtCh11.Text = "Fail";
                txtCh12.Text = "Fail";
                txtCh13.Text = "Fail";
                txtCh14.Text = "Fail";
                txtCh15.Text = "Fail";
                txtCh16.Text = "Fail";
                txtCh17.Text = "Fail";
            }
        }
        else
        {
            if (adamModbus.Modbus().ReadCoilStatus(iDiStart, m_iDiTotal, out bDiData) &&
                adamModbus.Modbus().ReadCoilStatus(iDoStart, m_iDoTotal, out bDoData))
            {
                iChTotal = m_iDiTotal + m_iDoTotal;
                bData = new bool[iChTotal];
                Array.Copy(bDiData, 0, bData, 0, m_iDiTotal);
                Array.Copy(bDoData, 0, bData, m_iDiTotal, m_iDoTotal);
                if (iChTotal > 0)
                    txtCh0.Text = bData[0].ToString();
                txtCh0.ForeColor = Color.Black;
                txtCh0.BackColor = Color.Green;


                if (bData[0] == false)
                {
                    txtCh0.BackColor = Color.Red;

                }

                if (txtCh0.Text == "True")
                {
                    btnCh_Click(12, txtCh12);
                }

                if (iChTotal > 1)
                    txtCh1.Text = bData[1].ToString();
                txtCh1.BackColor = Color.Green;

                if (txtCh1.Text == "True")
                {
                    MessageBox.Show("The system succesfully reset!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    passTxtBox.Text = "0";
                    rejectTxtBox.Text = "0";
                    total.Text = "0";
                    yieldTxtBox.Text = "0";
                    palletTxtBox.Clear();
                    serialTxtBox.Clear();
                    lowTxtBox1.Clear();
                    lowTxtBox2.Clear();
                    lowTxtBox3.Clear();
                    lowTxtBox4.Clear();
                    resultTxtBox1.Clear();
                    resultTxtBox2.Clear();
                    resultTxtBox3.Clear();
                    resultTxtBox4.Clear();
                    resultTxtBox5.Clear();
                    highTxtBox1.Clear();
                    highTxtBox2.Clear();
                    highTxtBox3.Clear();
                    highTxtBox4.Clear();
                    highTxtBox5.Clear();
                    passFailTxtBox.Clear();
                    passedResults = 0;
                    failedResults = 0;
                    totalResults = 0;

                }

                txtCh14.Text = bData[1].ToString();

                if (bData[1] == true)
                {
                    txtCh14.BackColor = Color.Gold;
                    txtCh14.Text = "Turn On";
                }
                else
                {
                    txtCh14.BackColor = Color.WhiteSmoke;
                    txtCh14.Text = "Turn Off";
                }
                if (bData[1] == false)
                {
                    txtCh1.BackColor = Color.Red;
                }

                if (iChTotal > 2)
                    txtCh2.Text = bData[2].ToString();
                txtCh2.BackColor = Color.Green;
                if (bData[2] == false)
                {
                    txtCh2.BackColor = Color.Red;
                }
                if (iChTotal > 3)
                    txtCh3.Text = bData[3].ToString();
                txtCh3.BackColor = Color.Green;
                if (bData[3] == false)
                {
                    txtCh3.BackColor = Color.Red;
                }
                if (iChTotal > 4)
                    txtCh4.Text = bData[4].ToString();
                txtCh4.BackColor = Color.Green;
                if (bData[4] == false)
                {
                    txtCh4.BackColor = Color.Red;
                }
                if (iChTotal > 5)
                    txtCh5.Text = bData[5].ToString();
                txtCh5.BackColor = Color.Green;
                if (bData[5] == false)
                {
                    txtCh5.BackColor = Color.Red;
                }
                if (iChTotal > 6)
                    txtCh6.Text = bData[6].ToString();
                txtCh6.BackColor = Color.Green;
                if (bData[6] == false)
                {
                    txtCh6.BackColor = Color.Red;
                }
                if (iChTotal > 7)
                    txtCh7.Text = bData[7].ToString();
                txtCh7.BackColor = Color.Green;
                if (bData[7] == false)
                {
                    txtCh7.BackColor = Color.Red;
                }
                if (iChTotal > 8)
                    txtCh8.Text = bData[8].ToString();
                txtCh8.BackColor = Color.Green;
                if (bData[8] == false)
                {
                    txtCh8.BackColor = Color.Red;
                }
                if (iChTotal > 9)
                    txtCh9.Text = bData[9].ToString();
                txtCh9.BackColor = Color.Green;
                if (bData[9] == false)
                {
                    txtCh9.BackColor = Color.Red;
                }
                if (iChTotal > 10)
                    txtCh10.Text = bData[10].ToString();
                txtCh10.BackColor = Color.Green;
                if (bData[10] == false)
                {
                    txtCh10.BackColor = Color.Red;
                }
                if (iChTotal > 11)
                    txtCh11.Text = bData[11].ToString();
                txtCh11.BackColor = Color.Green;
                if (bData[11] == false)
                {
                    txtCh11.BackColor = Color.Red;
                }
                if (iChTotal > 12)
                    //txtCh12.Text = bData[12].ToString();
                    //MessageBox.Show("Hello, world!", "Greetings");


                    if (bData[12] == true)
                    {
                        txtCh12.BackColor = Color.Gold;
                        txtCh12.Text = "True";


                    }
                    else
                    {
                        txtCh12.BackColor = Color.WhiteSmoke;
                        txtCh12.Text = "False";
                    }
                if (iChTotal > 13)
                    txtCh13.Text = bData[13].ToString();

                if (bData[13] == true)
                {
                    txtCh13.BackColor = Color.Gold;
                    txtCh13.Text = "True";
                }
                else
                {
                    txtCh13.BackColor = Color.WhiteSmoke;
                    txtCh13.Text = "False";
                }
                if (iChTotal > 14)
                    txtCh14.Text = bData[14].ToString();

                if (bData[14] == true)
                {
                    txtCh14.BackColor = Color.Gold;
                    txtCh14.Text = "True";
                }
                else
                {
                    txtCh14.BackColor = Color.WhiteSmoke;
                    txtCh14.Text = "False";
                }
                if (iChTotal > 15)
                    txtCh15.Text = bData[15].ToString();

                if (bData[15] == true)
                {
                    txtCh15.BackColor = Color.Gold;
                    txtCh15.Text = "True";
                }
                else
                {
                    txtCh15.BackColor = Color.WhiteSmoke;
                    txtCh15.Text = "False";
                }
                if (iChTotal > 16)
                    txtCh16.Text = bData[16].ToString();

                if (bData[16] == true)
                {
                    txtCh16.BackColor = Color.Gold;
                    txtCh16.Text = "True";
                }
                else
                {
                    txtCh16.BackColor = Color.WhiteSmoke;
                    txtCh16.Text = "False";
                }
                if (iChTotal > 17)
                    txtCh17.Text = bData[17].ToString();

                if (bData[17] == true)
                {
                    txtCh17.BackColor = Color.Gold;
                    txtCh17.Text = "True";
                }
                else
                {
                    txtCh17.BackColor = Color.WhiteSmoke;
                    txtCh17.Text = "False";
                }


            }
            else
            {
                txtCh0.Text = "Fail";
                txtCh1.Text = "Fail";
                txtCh2.Text = "Fail";
                txtCh3.Text = "Fail";
                txtCh4.Text = "Fail";
                txtCh5.Text = "Fail";
                txtCh6.Text = "Fail";
                txtCh7.Text = "Fail";
                txtCh8.Text = "Fail";
                txtCh9.Text = "Fail";
                txtCh10.Text = "Fail";
                txtCh11.Text = "Fail";
                txtCh12.Text = "Fail";
                txtCh13.Text = "Fail";
                txtCh14.Text = "Fail";
                txtCh15.Text = "Fail";
                txtCh16.Text = "Fail";
                txtCh17.Text = "Fail";
            }
        }

        System.GC.Collect();
    }

    protected void InitAdam6050()
    {
        int iDI = 0, iDO = 0;

        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh0, ref btnCh0);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh1, ref btnCh1);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh2, ref btnCh2);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh3, ref btnCh3);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh4, ref btnCh4);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh5, ref btnCh5);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh6, ref btnCh6);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh7, ref btnCh7);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh8, ref btnCh8);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh9, ref btnCh9);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh10, ref btnCh10);
        InitChannelItems(true, true, ref iDI, ref iDO, ref panelCh11, ref btnCh11);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh12, ref btnCh12);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh13, ref btnCh13);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh14, ref btnCh14);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh15, ref btnCh15);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh16, ref btnCh16);
        InitChannelItems(true, false, ref iDI, ref iDO, ref panelCh17, ref btnCh17);

        m_iDoTotal = iDO;
        m_iDiTotal = iDI;
    }

    private void btnCh12_Click(object sender, EventArgs e)
    {
        btnCh_Click(12, txtCh12);

    }

    private void btnCh13_Click(object sender, EventArgs e)
    {
        btnCh_Click(13, txtCh13);
    }

    private void btnCh14_Click(object sender, EventArgs e)
    {
        btnCh_Click(14, txtCh14);
    }

    private void btnCh15_Click(object sender, EventArgs e)
    {
        btnCh_Click(15, txtCh15);
    }

    private void btnCh16_Click(object sender, EventArgs e)
    {
        btnCh_Click(16, txtCh16);
    }

    private void btnCh17_Click(object sender, EventArgs e)
    {
        btnCh_Click(17, txtCh17);
    }

    private void btnCh_Click(int i_iCh, TextBox txtBox)
    {
        int iOnOff, iStart = 17 + i_iCh - m_iDiTotal;

        timer3.Enabled = false;
        if (txtBox.Text == "True") // was ON, now set to OFF
        {
            iOnOff = 0;


            RefreshDIO();
        }
        else
        {
            iOnOff = 1;
        }
        if (adamModbus.Modbus().ForceSingleCoil(iStart, iOnOff))
            RefreshDIO();
        else
            MessageBox.Show("Set digital output failed!", "Error");
        timer3.Enabled = true;
    }
    private void Form2_Load(object sender, EventArgs e)
    {

    }

    private void timer3_Tick(object sender, EventArgs e)
    {
        timer3.Enabled = false;

        m_iCount++; // increment the reading counter
        txtReadCount.Text = "Read coil " + m_iCount.ToString() + " times...";
        RefreshDIO();

        timer3.Enabled = true;
    }



    // ========= START CODE FOR CONFIGURATION TAB ========== //
    private void configurationData()
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();
            SqlDataAdapter resData = new SqlDataAdapter("SELECT * FROM ConfigurationData", connection);
            DataTable dt = new DataTable();
            resData.Fill(dt);

            configView.DataSource = dt;

            DataGridViewColumn centerModelData = configView.Columns["ModelData"];


            foreach (DataGridViewColumn column in configView.Columns)
            {
                column.Width = 339;
            }

            DataGridViewColumn hideLow1 = configView.Columns["lowOne"];
            DataGridViewColumn hideLow2 = configView.Columns["lowTwo"];
            DataGridViewColumn hideLow3 = configView.Columns["lowThree"];
            DataGridViewColumn hideLow4 = configView.Columns["lowFour"];
            DataGridViewColumn hideHigh1 = configView.Columns["highOne"];
            DataGridViewColumn hideHigh2 = configView.Columns["highTwo"];
            DataGridViewColumn hideHigh3 = configView.Columns["highThree"];
            DataGridViewColumn hideHigh4 = configView.Columns["highFour"];
            DataGridViewColumn hideHigh5 = configView.Columns["highFive"];
            DataGridViewColumn hideTotal = configView.Columns["Total"];



            DataGridViewColumn columnToHide = configView.Columns["ID"];
            if (columnToHide != null || hideLow1 != null || hideLow2 != null || hideLow3 != null || hideLow4 != null || hideHigh1 != null
                || hideHigh2 != null || hideHigh3 != null || hideHigh4 != null || hideHigh5 != null)
            {
                hideLow1.Visible = false;
                hideLow2.Visible = false;
                hideLow3.Visible = false;
                hideLow4.Visible = false;
                hideHigh1.Visible = false;
                hideHigh2.Visible = false;
                hideHigh3.Visible = false;
                hideHigh4.Visible = false;
                hideHigh5.Visible = false;
                columnToHide.Visible = false;
                hideTotal.Visible = false;
            }
            else
            {
                MessageBox.Show("Column not found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            configView.Columns["ModelData"].HeaderText = "Model Data";
            centerModelData.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerModelData.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            centerModelData.HeaderCell.Style.Font = new Font(configView.Font, FontStyle.Bold);
        }
    }

    private void configView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {

        if (e.RowIndex >= 0)
        {
            configConButton.Visible = false;
            configUpdate.Visible = true;
            configCancel.Visible = true;
            configUpdate.Location = new System.Drawing.Point(380, 438);
            configCancel.Location = new System.Drawing.Point(585, 438);

            DataGridViewRow row = this.configView.Rows[e.RowIndex];
            ConfigureSettings.Text = "Information details";
            ConfigureSettings.Location = new System.Drawing.Point(402, 23);
            IDText.Text = row.Cells["ID"].Value.ToString();
            setModel.Text = row.Cells["ModelData"].Value.ToString();
            setLow1.Text = row.Cells["lowOne"].Value.ToString();
            setLow2.Text = row.Cells["lowTwo"].Value.ToString();
            setLow3.Text = row.Cells["lowThree"].Value.ToString();
            setLow4.Text = row.Cells["lowFour"].Value.ToString();
            setHigh1.Text = row.Cells["highOne"].Value.ToString();
            setHigh2.Text = row.Cells["highTwo"].Value.ToString();
            setHigh3.Text = row.Cells["highThree"].Value.ToString();
            setHigh4.Text = row.Cells["highFour"].Value.ToString();
            setHigh5.Text = row.Cells["highFive"].Value.ToString();
            setTotal.Text = row.Cells["Total"].Value.ToString();
        }
        else
        {
            configConButton.Visible = true;
            configUpdate.Visible = false;
            configCancel.Visible = false;
            IDText.Clear();
            setModel.Clear();
            setLow1.Clear();
            setLow2.Clear();
            setLow3.Clear();
            setLow4.Clear();
            setHigh1.Clear();
            setHigh2.Clear();
            setHigh3.Clear();
            setHigh4.Clear();
            setHigh5.Clear();
            setTotal.Clear();
        }
    }

    private void configCancel_Click(object sender, EventArgs e)
    {
        ConfigureSettings.Text = "Configure Settings";
        setModel.Text = null;
        setLow1.Text = null;
        setLow2.Text = null;
        setLow3.Text = null;
        setLow4.Text = null;
        setHigh1.Text = null;
        setHigh2.Text = null;
        setHigh3.Text = null;
        setHigh4.Text = null;
        setHigh5.Text = null;
        configUpdate.Visible = false;
        configCancel.Visible = false;
        configConButton.Visible = true;
    }

    private void configConButton_Click(object sender, EventArgs e)
    {
        string Modeldata = setModel.Text;
        string Lowone = setLow1.Text;
        string Lowtwo = setLow2.Text;
        string Lowthree = setLow3.Text;
        string Lowfour = setLow4.Text;
        string Highone = setHigh1.Text;
        string Hightwo = setHigh2.Text;
        string Highthree = setHigh3.Text;
        string Highfour = setHigh4.Text;
        string Highfive = setHigh5.Text;
        string TotalItems = setTotal.Text;

        if (!int.TryParse(setLow1.Text, out _) || !int.TryParse(setLow2.Text, out _) || !int.TryParse(setLow3.Text, out _) || !int.TryParse(setLow4.Text, out _) || !int.TryParse(setHigh1.Text, out _)
            || !int.TryParse(setHigh2.Text, out _) || !int.TryParse(setHigh3.Text, out _) || !int.TryParse(setHigh4.Text, out _) || !int.TryParse(setHigh5.Text, out _) || !int.TryParse(setTotal.Text, out _))
        {
            MessageBox.Show("Please insert valid numbers", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            setLow1.Focus();
            setLow2.Focus();
            setLow3.Focus();
            setLow4.Focus();
            setHigh1.Focus();
            setHigh2.Focus();
            setHigh3.Focus();
            setHigh4.Focus();
            setHigh5.Focus();
            setTotal.Focus();
        }
        else
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query = "INSERT INTO ConfigurationData (ModelData, lowOne, lowTwo, lowThree, lowFour, highOne, highTwo, highThree, highFour, highFive, Total) " +
                               "VALUES (@Modeldata, @Lowone, @Lowtwo, @Lowthree, @Lowfour, @Highone, @Hightwo, @Highthree, @Highfour, @Highfive, @Total)";

                SqlCommand command = new SqlCommand(query, connection);
                command.Parameters.AddWithValue("@Modeldata", Modeldata);
                command.Parameters.AddWithValue("@Lowone", Lowone);
                command.Parameters.AddWithValue("@Lowtwo", Lowtwo);
                command.Parameters.AddWithValue("@Lowthree", Lowthree);
                command.Parameters.AddWithValue("@Lowfour", Lowfour);
                command.Parameters.AddWithValue("@Highone", Highone);
                command.Parameters.AddWithValue("@Hightwo", Hightwo);
                command.Parameters.AddWithValue("@Highthree", Highthree);
                command.Parameters.AddWithValue("@Highfour", Highfour);
                command.Parameters.AddWithValue("@Highfive", Highfive);
                command.Parameters.AddWithValue("@Total", TotalItems);

                try
                {
                    connection.Open();
                    int rowsAffected = command.ExecuteNonQuery();
                    MessageBox.Show("New settings successfully added", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //createLabel.Text = "Create New User Account";

                    configurationData();

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }

            }
        }
    }

    private void configUpdate_Click(object sender, EventArgs e)
    {

        string model = setModel.Text;
        string low1 = setLow1.Text;
        string low2 = setLow2.Text;
        string low3 = setLow3.Text;
        string low4 = setLow4.Text;
        string high1 = setHigh1.Text;
        string high2 = setHigh2.Text;
        string high3 = setHigh3.Text;
        string high4 = setHigh4.Text;
        string high5 = setHigh5.Text;
        string ID = IDText.Text;
        string TotalItems = setTotal.Text;

        if (!int.TryParse(setLow1.Text, out _) || !int.TryParse(setLow2.Text, out _) || !int.TryParse(setLow3.Text, out _) || !int.TryParse(setLow4.Text, out _) || !int.TryParse(setHigh1.Text, out _)
            || !int.TryParse(setHigh2.Text, out _) || !int.TryParse(setHigh3.Text, out _) || !int.TryParse(setHigh4.Text, out _) || !int.TryParse(setHigh5.Text, out _) || !int.TryParse(setTotal.Text, out _))
        {
            MessageBox.Show("Please insert valid numbers", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            setLow1.Focus();
            setLow2.Focus();
            setLow3.Focus();
            setLow4.Focus();
            setHigh1.Focus();
            setHigh2.Focus();
            setHigh3.Focus();
            setHigh4.Focus();
            setHigh5.Focus();
            setTotal.Focus();
        }
        else
        {
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                string query1 = "UPDATE ConfigurationData SET ModelData = @ModelData, lowOne = @lowOne, lowTwo = @lowTwo, lowThree = @lowThree, lowFour =@lowFour, highOne =@highOne, " +
                                "highTwo = @highTwo, highThree = @highThree, highFour = @highFour, highFive = @highFive, Total = @Total WHERE ID = @ID";

                SqlCommand command1 = new SqlCommand(query1, connection);
                command1.Parameters.AddWithValue("@ID", ID);
                command1.Parameters.AddWithValue("@ModelData", model);
                command1.Parameters.AddWithValue("@lowOne", low1);
                command1.Parameters.AddWithValue("@lowTwo", low2);
                command1.Parameters.AddWithValue("@lowThree", low3);
                command1.Parameters.AddWithValue("@lowFour", low4);
                command1.Parameters.AddWithValue("@highOne", high1);
                command1.Parameters.AddWithValue("@highTwo", high2);
                command1.Parameters.AddWithValue("@highThree", high3);
                command1.Parameters.AddWithValue("@highFour", high4);
                command1.Parameters.AddWithValue("@highFive", high5);
                command1.Parameters.AddWithValue("@Total", TotalItems);



                connection.Open();
                int rowsAffected = command1.ExecuteNonQuery();
                MessageBox.Show($"{setModel.Text} update successfully!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                configView.Refresh();
                configurationData();
            }
        }


    }
    // ========= END CODE FOR CONFIGURATION TAB ========== //

    private void ValidationInput()
    {

        if (!int.TryParse(setLow1.Text, out _) || !int.TryParse(setLow2.Text, out _) || !int.TryParse(setLow3.Text, out _) || !int.TryParse(setLow4.Text, out _) || !int.TryParse(setHigh1.Text, out _)
            || !int.TryParse(setHigh2.Text, out _) || !int.TryParse(setHigh3.Text, out _) || !int.TryParse(setHigh4.Text, out _) || !int.TryParse(setHigh5.Text, out _))
        {
            MessageBox.Show("Please insert valid numbers", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Error);
            setLow1.Clear();
            setLow2.Clear();
            setLow3.Clear();
            setLow4.Clear();
            setHigh1.Clear();
            setHigh2.Clear();
            setHigh3.Clear();
            setHigh4.Clear();
            setHigh5.Clear();
            setLow1.Focus();
            setLow2.Focus();
            setLow3.Focus();
            setLow4.Focus();
            setHigh1.Focus();
            setHigh2.Focus();
            setHigh3.Focus();
            setHigh4.Focus();
            setHigh5.Focus();
        }
    }

    int passedResults = 0;
    int failedResults = 0;
    int totalResults = 0;


    private void startTestButton_Click(object sender, EventArgs e)
    {

        isTesting = true;
        isFirstResultDisplayed = false;



        // Fetch and display data
        FetchAndDisplayData();
        // Check if testing is already ongoing
        //if (!isTesting)
        //{
        //    isTesting = true; 

        //fetchCount = 0;     

        timer.Stop();

        //int result1 = int.Parse(resultTxtBox1.Text);
        //int lowCom1 = int.Parse(lowTxtBox1.Text);
        //int highCom1 = int.Parse(highTxtBox1.Text);

        //int result2 = int.Parse(resultTxtBox2.Text);
        //int lowCom2 = int.Parse(lowTxtBox2.Text);
        //int highCom2 = int.Parse(highTxtBox2.Text);

        double result3 = ((int)reading);
        int lowCom3 = int.Parse(lowTxtBox3.Text);
        int highCom3 = int.Parse(highTxtBox3.Text);

        double result4 = reading;
        int highCom4 = int.Parse(highTxtBox4.Text);

        //int result5 = int.Parse(resultTxtBox5.Text);
        //int lowCom4 = int.Parse(lowTxtBox4.Text);
        //int highCom5 = int.Parse(highTxtBox5.Text);

        string Pallet = palletTxtBox.Text;
        string Serial = serialTxtBox.Text;
        string Low1 = lowTxtBox1.Text;
        string Low2 = lowTxtBox2.Text;
        string Low3 = lowTxtBox3.Text;
        string Low4 = lowTxtBox4.Text;
        string Result1 = resultTxtBox1.Text;
        string Result2 = resultTxtBox2.Text;
        string Result3 = resultTxtBox3.Text;
        string Result4 = resultTxtBox4.Text;
        string Result5 = resultTxtBox5.Text;
        string High1 = highTxtBox1.Text;
        string High2 = highTxtBox2.Text;
        string High3 = highTxtBox3.Text;
        string High4 = highTxtBox4.Text;
        string High5 = highTxtBox5.Text;
        string PassFail = passFailTxtBox.Text;
        string TestTime = testTimeTxtbox.Text;
        string Pass = passTxtBox.Text;
        string Reject = rejectTxtBox.Text;
        string Total = total.Text;
        string Yield = yieldTxtBox.Text;
        string TimeDate = lblDateTime.Text;

        int totalItem = int.Parse(total.Text); // Assuming totalTxtBox contains the total number of items

        int stop = int.Parse(totalTxtBox.Text);


        if (Low1 != "0" || Low2 != "0" || Low3 != "0" || Low4 != "0" || High1 != "0" || High2 != "0" || High3 != "0"
            || High4 != "0" || High5 != "0")
        {
            if (modelComboBox.Text == "No Selected")
            {
                MessageBox.Show("PLEASE INPUT MODEL!", "", MessageBoxButtons.OK);
                timer.Stop();
            }
            else
            {
                onGoingBox.Text = "ON GOING TESTING";
                onGoingBox.ForeColor = Color.Black;
                onGoingBox.BackColor = Color.Yellow;

                // Start the timer
                timer.Start();

                btnCh_Click(14, txtCh14);
                btnCh_Click(13, txtCh14);
                stopwatch.Start();
                timer2.Start();
                if (totalItem < stop)
                {

                    if (result3 >= lowCom3 && result3 <= highCom3 && result4 <= highCom4)
                    {
                        passedResults += 1;
                        totalResults += 1;
                        passTxtBox.Text = passedResults.ToString();
                        total.Text = totalResults.ToString();

                        passFailTxtBox.Text = "PASSED";
                        passFailTxtBox.BackColor = Color.Green;

                        int passedItems = int.Parse(passTxtBox.Text); // Assuming passTxtBox contains the number of passed items
                        int totalItems = int.Parse(total.Text); // Assuming totalTxtBox contains the total number of items


                        if (totalItems != 0)
                        {
                            // Calculate the percentage
                            double passPercentage = (double)passedItems / totalItems * 100;

                            // Display the percentage in a TextBox or Label
                            yieldTxtBox.Text = passPercentage.ToString("0.##") + "%";
                        }
                        else
                        {
                            // If the total number of items is zero, display "0%" or handle it as per your requirement
                            yieldTxtBox.Text = "0%";
                        }



                    }
                    else
                    {
                        int passedItemsFail = int.Parse(passTxtBox.Text); // Assuming passTxtBox contains the number of passed items
                        int totalItemsFail = int.Parse(total.Text); // Assuming totalTxtBox contains the total number of items

                        if (totalItemsFail != 0)
                        {
                            // Calculate the percentage
                            double passPercentageFail = (double)passedItemsFail / totalItemsFail * 100;

                            // Display the percentage in a TextBox or Label
                            yieldTxtBox.Text = passPercentageFail.ToString("0.##") + "%";
                        }
                        else
                        {
                            // If the total number of items is zero, display "0%" or handle it as per your requirement
                            yieldTxtBox.Text = "0%";
                        }

                        failedResults += 1;
                        totalResults += 1;
                        rejectTxtBox.Text = failedResults.ToString();
                        total.Text = totalResults.ToString();
                        passFailTxtBox.Text = "FAILED";
                        passFailTxtBox.BackColor = Color.Red;
                    }



                    //if (result1 >= lowCom1 && result1 <= highCom1)
                    //{
                    //    resultTxtBox1.BackColor = Color.Green;
                    //}
                    //else
                    //{
                    //    resultTxtBox1.BackColor = Color.Red;
                    //}

                    //if (result2 >= lowCom2 && result2 <= highCom2)
                    //{
                    //    resultTxtBox2.BackColor = Color.Green;
                    //}
                    //else
                    //{
                    //    resultTxtBox2.BackColor = Color.Red;
                    //}

                    if (result3 >= lowCom3 && result3 <= highCom3)
                    {
                        resultTxtBox3.BackColor = Color.Green;
                    }
                    else
                    {
                        resultTxtBox3.BackColor = Color.Red;
                    }

                    if (result4 <= highCom4)
                    {
                        resultTxtBox4.BackColor = Color.Green;
                    }
                    else
                    {
                        resultTxtBox4.BackColor = Color.Red;
                    }

                    //if (result5 >= lowCom4 && result5 <= highCom5)
                    //{
                    //    resultTxtBox5.BackColor = Color.Green;
                    //}
                    //else
                    //{
                    //    resultTxtBox5.BackColor = Color.Red;
                    //} 
                    //}
                }
                else
                {
                    startTestButton.Enabled = false;

                }
            }

        }
        else
        {


            resultTxtBox1.Text = "";
            resultTxtBox2.Text = "";
            resultTxtBox3.Text = "";
            resultTxtBox4.Text = "";
            resultTxtBox5.Text = "";

            btnCh_Click(14, txtCh14);
            btnCh_Click(12, txtCh12);
            btnCh_Click(15, txtCh15);

            DialogResult result = MessageBox.Show("Please scan your items!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            timer.Stop();
            if (result == DialogResult.OK)
            {
                btnCh_Click(12, txtCh12);
                btnCh_Click(14, txtCh14);
                btnCh_Click(15, txtCh15);
                timer.Stop();

            }


        }

    }

    private void Timer_Tick(object sender, EventArgs e)
    {
        string Model = modelComboBox.Text;
        string Pallet = palletTxtBox.Text;
        string Serial = serialTxtBox.Text;
        string Low1 = lowTxtBox1.Text;
        string Low2 = lowTxtBox2.Text;
        string Low3 = lowTxtBox3.Text;
        string Low4 = lowTxtBox4.Text;
        string Result1 = resultTxtBox1.Text;
        string Result2 = resultTxtBox2.Text;
        string Result3 = resultTxtBox3.Text;
        string Result4 = resultTxtBox4.Text;
        string Result5 = resultTxtBox5.Text;
        string High1 = highTxtBox1.Text;
        string High2 = highTxtBox2.Text;
        string High3 = highTxtBox3.Text;
        string High4 = highTxtBox4.Text;
        string High5 = highTxtBox5.Text;
        string PassFail = passFailTxtBox.Text;
        string TestTime = testTimeTxtbox.Text;
        string Pass = passTxtBox.Text;
        string Reject = rejectTxtBox.Text;
        string Total = total.Text;
        string Yield = yieldTxtBox.Text;
        string TimeDate = lblDateTime.Text;

        if (Low1 != "0" || Low2 != "0" || Low3 != "0" || Low4 != "0" || High1 != "0" || High2 != "0" || High3 != "0"
            || High4 != "0" || High5 != "0")
        {
            if (isTesting)
            {
                //Thread.Sleep(15000);
                // Fetch and display data
                FetchAndDisplayData();

                // Increment the fetch count
                fetchCount++;
                //resultTxtBox1.Text = fetchCount.ToString();
                //int test = int.Parse(resultTxtBox1.Text);
                // Stop the timer after the second fetch
                if (fetchCount >= 2)
                {
                    timer.Stop();
                    onGoingBox.Text = "COMPLETED";
                    onGoingBox.ForeColor = Color.Black;
                    onGoingBox.BackColor = Color.LimeGreen;

                    isTesting = false;

                    DialogResult result = MessageBox.Show("TESTING IS COMPLETED!", "Confirmation", MessageBoxButtons.OK);
                    modelComboBox.SelectedIndex = 0;

                    // Check if the user clicked OK
                    if (result == DialogResult.OK)
                    {

                        fetchCount = 0;
                        onGoingBox.Text = "";
                        passFailTxtBox.Text = "";
                        palletTxtBox.Text = "";
                        serialTxtBox.Text = "";
                        //passTxtBox.Text = "0";
                        //rejectTxtBox.Text = "0";
                        //totalTxtBox.Text = "0";
                        //yieldTxtBox.Text = "0";
                        lowTxtBox1.Text = "0";
                        lowTxtBox2.Text = "0";
                        lowTxtBox3.Text = "0";
                        lowTxtBox4.Text = "0";
                        highTxtBox1.Text = "0";
                        highTxtBox2.Text = "0";
                        highTxtBox3.Text = "0";
                        highTxtBox4.Text = "0";
                        highTxtBox5.Text = "0";
                        resultTxtBox1.Text = "0";
                        resultTxtBox2.Text = "0";
                        resultTxtBox3.Text = "0";
                        resultTxtBox4.Text = "0";
                        resultTxtBox5.Text = "0";
                        resultTxtBox1.BackColor = System.Drawing.Color.White;
                        resultTxtBox2.BackColor = System.Drawing.Color.White;
                        resultTxtBox3.BackColor = System.Drawing.Color.White;
                        resultTxtBox4.BackColor = System.Drawing.Color.White;
                        resultTxtBox5.BackColor = System.Drawing.Color.White;
                        onGoingBox.BackColor = System.Drawing.Color.White;
                        onGoingBox.Text = "STAND BY";

                        btnCh_Click(13, txtCh13);
                        btnCh_Click(14, txtCh14);

                        int totalItem = int.Parse(total.Text); // Assuming totalTxtBox contains the total number of items

                        int stop = int.Parse(totalTxtBox.Text);


                        if (totalItem == stop)
                        {
                            newTestButton.Enabled = true;
                            startTestButton.Enabled = false;
                            newTestButton.BackColor = System.Drawing.Color.DeepSkyBlue;
                            startTestButton.BackColor = System.Drawing.Color.LightGray;

                            totalResults = 0;
                            passTxtBox.Text = "0";
                            rejectTxtBox.Text = "0";
                            totalTxtBox.Text = "0";
                            yieldTxtBox.Text = "0";
                            total.Text = "0";
                            testTimeTxtbox.Text = "00:00:0";

                            stopwatch.Stop();
                            timer2.Stop();
                            stopwatch.Reset();

                        }


                    }

                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        string query = "INSERT INTO ResultInfo (PalletSN, SerialNumber, Model, Low1, Low2, Low3, Low4, Result1, Result2, Result3, Result4, Result5, " +
                                       "High1, High2, High3, High4, High5, PassedorFailed, TestTime, Pass, Reject, Total, Yield, DateTimeInfo) " +
                                       "VALUES (@Pallet, @Serial, @Model, @Low1, @Low2, @Low3, @Low4, @Result1 ,@Result2 ,@Result3 ,@Result4, @Result5," +
                                       "@High1 ,@High2 ,@High3 ,@High4 ,@High5 ,@PassFail ,@TestTime ,@Pass ,@Reject ,@Total ,@Yield ,@TimeDate)";

                        SqlCommand command = new SqlCommand(query, connection);
                        command.Parameters.AddWithValue("@Pallet", Pallet);
                        command.Parameters.AddWithValue("@Serial", Serial);
                        command.Parameters.AddWithValue("@Model", Model);
                        command.Parameters.AddWithValue("@Low1", Low1);
                        command.Parameters.AddWithValue("@Low2", Low2);
                        command.Parameters.AddWithValue("@Low3", Low3);
                        command.Parameters.AddWithValue("@Low4", Low4);
                        command.Parameters.AddWithValue("@Result1", Result1);
                        command.Parameters.AddWithValue("@Result2", Result2);
                        command.Parameters.AddWithValue("@Result3", Result3);
                        command.Parameters.AddWithValue("@Result4", Result4);
                        command.Parameters.AddWithValue("@Result5", Result5);
                        command.Parameters.AddWithValue("@High1", High1);
                        command.Parameters.AddWithValue("@High2", High2);
                        command.Parameters.AddWithValue("@High3", High3);
                        command.Parameters.AddWithValue("@High4", High4);
                        command.Parameters.AddWithValue("@High5", High5);
                        command.Parameters.AddWithValue("@PassFail", PassFail);
                        command.Parameters.AddWithValue("@TestTime", TestTime);
                        command.Parameters.AddWithValue("@Pass", Pass);
                        command.Parameters.AddWithValue("@Reject", Reject);
                        command.Parameters.AddWithValue("@Total", Total);
                        command.Parameters.AddWithValue("@Yield", Yield);
                        command.Parameters.AddWithValue("@TimeDate", TimeDate);


                        try
                        {
                            connection.Open();
                            int rowsAffected = command.ExecuteNonQuery();
                            //MessageBox.Show($"{fname} account successfully created");
                            //MessageBox.Show("Data Saved!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            isTesting = false;
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Error: " + ex.Message);
                        }

                    }


                }
            }
        }
        else
        {
        //    timer.Stop();
        }


    }

    private void FetchAndDisplayData()
    {
        // Fetch data from the source
        double ion_reading;
        string ion_reading_string;
        bool over_flow;
        int result = Read_Ions_Reading(out ion_reading, out ion_reading_string, out over_flow);

        string data = result == 0 ? $" {ion_reading_string}" : "0";

        // Update the TextBox on the UI thread
        //if (palletTxtBox.InvokeRequired)
        //{
        //    palletTxtBox.BeginInvoke((Action)(() =>
        //    {
        //        palletTxtBox.Text = data;
        //    }));
        //}
        //else
        //{
        //    serialTxtBox.Text = data;
        //}

        // Determine which TextBox to update
        if (fetchCount == 0)
        {
            if (!isFirstResultDisplayed)
            {
                if (resultTxtBox3.InvokeRequired)
                {
                    resultTxtBox3.BeginInvoke((Action)(() =>
                    {
                        resultTxtBox3.Text = Math.Round(reading, 2).ToString();
                    }));
                }
                else
                {
                    resultTxtBox3.Text = Math.Round(reading, 2).ToString();
                }
                isFirstResultDisplayed = true;
            }
        }
        else if (fetchCount == 1)
        {
            if (resultTxtBox4.InvokeRequired)
            {
                resultTxtBox4.BeginInvoke((Action)(() =>
                {
                    resultTxtBox4.Text = Math.Round(reading, 2).ToString();
                }));
            }
            else
            {
                resultTxtBox4.Text = Math.Round(reading, 2).ToString();
            }
        }
    }

    private static int Read_Ions_Reading(out double ion_reading, out string ion_reading_string, out bool over_flow)
    {
        DateTime dtStart = DateTime.Now;
        reading_valid = false;
        over_flow = false;
        ion_reading = 0;
        ion_reading_string = "";
        do
        {
            Thread.Sleep(10);
            if (reading_valid)
            {
                ion_reading = reading;
                ion_reading_string = reading_string;
                over_flow = reading_OL;
                return 0;
            }
        } while ((DateTime.Now - dtStart).TotalSeconds < 1);
        return -1;
    }

    public static void MainLoop()
    {
        for (; ; )
        {
            Thread.Sleep(100);
            if (!spDLY.IsOpen) continue;
            TimeSpan ts = DateTime.Now - dtLastRxValidReading;
            if (ts.TotalSeconds > 2)
            {
                reading_valid = false;
                reading_OL = false;
            }

            int len = spDLY.BytesToRead;
            if (len > 0)
            {
                Thread.Sleep(30);
                byte[] rx_buf = new byte[141];
                if (len > RxDataSize) len = RxDataSize;
                spDLY.Read(rx_buf, 0, len);
                for (int i = 0; i < len; i++) rx_data[i] = rx_buf[i];
                rx_len = len;
                if (rx_len >= 14)
                {
                    if ((rx_data[0] == 0x15) && (rx_data[13] == 0xE4))
                    {
                        string sSign = "";
                        if ((rx_buf[1] & 0x07) == 0x00)
                        {
                            reading_OL = true;
                        }
                        else
                        {
                            reading_OL = false;
                        }
                        if ((rx_data[1] & 0x08) == 0x08) sSign = "-";
                        byte[] data = new byte[4];
                        string[] sData = new string[4];
                        bool invalid_format = false;
                        for (int i = 0; i < 4; i++)
                        {
                            byte msb = (byte)(rx_data[i * 2 + 1] & 0x07);
                            byte lsb = (byte)(rx_data[i * 2 + 2] & 0x0F);
                            data[i] = (byte)(msb * 16 + lsb);
                            string s = "0";
                            switch (data[i])
                            {
                                case 0x7D: s = "0"; break;
                                case 0x05: s = "1"; break;
                                case 0x5B: s = "2"; break;
                                case 0x1F: s = "3"; break;
                                case 0x27: s = "4"; break;
                                case 0x3E: s = "5"; break;
                                case 0x7E: s = "6"; break;
                                case 0x15: s = "7"; break;
                                case 0x7F: s = "8"; break;
                                case 0x3F: s = "9"; break;
                                default: invalid_format = true; break;
                            }
                            sData[i] = s;
                        }
                        if (!invalid_format)
                        {
                            reading_string = sSign + sData[0] + sData[1] + sData[2] + sData[3];

                            try
                            {
                                reading = ion_multiply_factor * double.Parse(reading_string);
                                dtLastRxValidReading = DateTime.Now;
                                reading_valid = true;
                            }
                            catch (Exception ex)
                            {
                                // Handle exception
                            }
                        }
                    }
                }
            }
        }
    }

    private void newTestButton_Click(object sender, EventArgs e)
    {
        totalResults = 0;
        passedResults = 0;
        failedResults = 0;
        passTxtBox.Text = "0";
        totalTxtBox.Text = "0";
        yieldTxtBox.Text = "0";
        startTestButton.Enabled = true;
        newTestButton.Enabled = false;
        startTestButton.BackColor = System.Drawing.Color.Green;
        newTestButton.BackColor = System.Drawing.Color.LightGray;
    }

    private void backToLot_Click(object sender, EventArgs e)
    {
        displayResult();
        exportLotButton.Visible = false;
        backToLot.Visible = false;
        DataResultPanel.Visible = false;
        noResult.Visible = false;
        resultView.Visible = true;

        IDssText.Clear();
        disPassed.Clear();
        disReject.Clear();
        disTotal.Clear();
        disYield.Clear();
        disModel.Clear();
        disPallet.Clear();
        disSerial.Clear();
        //disTestTime.Clear();
        displayLow1.Clear();
        displayLow2.Clear();
        displayLow3.Clear();
        displayLow4.Clear();
        displayResult1.Clear();
        displayResult2.Clear();
        displayResult3.Clear();
        displayResult4.Clear();
        displayResult5.Clear();
        displayHigh1.Clear();
        displayHigh2.Clear();
        displayHigh3.Clear();
        displayHigh4.Clear();
        displayHigh5.Clear();
        displayPassFail.Clear();

    }

    private void perTestView_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
    {
        if (e.RowIndex >= 0) //&& e.RowIndex < resultView.Rows.Count - 1
        {
            backToLot.Visible = true;
            resultView.Visible = false;
            perTestView.Visible = true;
            noResult.Visible = true;
            DataResultPanel.Visible = true;

            DataGridViewRow row = this.perTestView.Rows[e.RowIndex];
            lotBox.Text = row.Cells["SerialNumber"].Value.ToString();


            dataResultLabel.Text = "Data Result for " + row.Cells["PalletSN"].Value.ToString();
            IDssText.Text = row.Cells["ID"].Value.ToString();
            disPassed.Text = row.Cells["Pass"].Value.ToString();
            disReject.Text = row.Cells["Reject"].Value.ToString();
            disTotal.Text = row.Cells["Total"].Value.ToString();
            disYield.Text = row.Cells["Yield"].Value.ToString();
            disModel.Text = row.Cells["Model"].Value.ToString();
            disPallet.Text = row.Cells["PalletSN"].Value.ToString();
            disSerial.Text = row.Cells["SerialNumber"].Value.ToString();
            //disTestTime.Text = row.Cells["TestTime"].Value.ToString();
            displayLow1.Text = row.Cells["Low1"].Value.ToString();
            displayLow2.Text = row.Cells["Low2"].Value.ToString();
            displayLow3.Text = row.Cells["Low3"].Value.ToString();
            displayLow4.Text = row.Cells["Low4"].Value.ToString();
            displayResult1.Text = row.Cells["Result1"].Value.ToString();
            displayResult2.Text = row.Cells["Result2"].Value.ToString();
            displayResult3.Text = row.Cells["Result3"].Value.ToString();
            displayResult4.Text = row.Cells["Result4"].Value.ToString();
            displayResult5.Text = row.Cells["Result5"].Value.ToString();
            displayHigh1.Text = row.Cells["High1"].Value.ToString();
            displayHigh2.Text = row.Cells["High2"].Value.ToString();
            displayHigh3.Text = row.Cells["High3"].Value.ToString();
            displayHigh4.Text = row.Cells["High4"].Value.ToString();
            displayHigh5.Text = row.Cells["High5"].Value.ToString();
            displayPassFail.Text = row.Cells["PassedorFailed"].Value.ToString();

            if (disPallet.Text == "" && disSerial.Text == "")
            {
                dataResultLabel.Text = "Data Result";
            }

        }
        else
        {
            IDssText.Clear();
            disPassed.Clear();
            disReject.Clear();
            disTotal.Clear();
            disYield.Clear();
            disModel.Clear();
            disPallet.Clear();
            disSerial.Clear();
            //disTestTime.Clear();
            displayLow1.Clear();
            displayLow2.Clear();
            displayLow3.Clear();
            displayLow4.Clear();
            displayResult1.Clear();
            displayResult2.Clear();
            displayResult3.Clear();
            displayResult4.Clear();
            displayResult5.Clear();
            displayHigh1.Clear();
            displayHigh2.Clear();
            displayHigh3.Clear();
            displayHigh4.Clear();
            displayHigh5.Clear();
            displayPassFail.Clear();
        }

    }

    private void exportLotButton_Click(object sender, EventArgs e)
    {
        // Load the Excel template file
        string templatePath = @"C:\Users\BSS-PROGRAMMER\Documents\PROJECTS - JAMESON\DYSON\IoniserTester_ADVANCE\dyson-per-lot.xlsx";
        var workbook = ExcelFile.Load(templatePath);
        var worksheet = workbook.Worksheets[0];

        // Add DataGridView data to the Excel file starting at cell A20
        int startRow = 14;
        int startColumn = 1; // Column B in Excel notation

        //for (int i = 0; i < perTestView.Rows.Count; i++)
        //{
        //    for (int j = 0; j < perTestView.Columns.Count; j++)
        //    {
        //        worksheet.Cells[startRow + i, startColumn + j].Value = perTestView.Rows[i].Cells[j].Value?.ToString();
        //    }
        //}
        string[] columnsToExport = { "DateTimeInfo", "SerialNumber", "PalletSN", "Model", "PassedorFailed", "Low1", "Result1", "High1", "Low2", "Result2", "High2",
                                     "Low3", "Result3", "High3", "Result4", "High4", "Low4", "Result5", "High5", "TestTime"};
        int currentRow = startRow;

        foreach (DataGridViewRow row in perTestView.Rows)
        {
            if (!row.IsNewRow)
            {
                for (int colIndex = 0; colIndex < columnsToExport.Length; colIndex++)
                {
                    var cellValue = row.Cells[columnsToExport[colIndex]].Value?.ToString();
                    worksheet.Cells[currentRow, startColumn + colIndex].Value = cellValue ?? string.Empty;
                }
                currentRow++;
            }
        }

        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
        // Define the file name
        string fileName = $"X590Ionizer - {timestamp}.xlsx"; // Set your desired file name

        // Specify the full file path
        string outputDirectory = @"C:\Users\BSS-PROGRAMMER\Documents\X590-Output\Per Lot Result";
        string filePath = Path.Combine(outputDirectory, fileName); // Set your specific folder path

        // Save the modified Excel file to the specified file path
        workbook.Save(filePath);

        // Show a message box indicating successful saving
        MessageBox.Show($"Successfully exported {fileName}!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

    }


}