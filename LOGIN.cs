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
using System.Configuration;
using System.Net.Mail;
using System.Reflection;
using Microsoft.Win32;

namespace In_House_Calibration
{
    public partial class LOGIN : Form
    {
        public static string loginuser = "";
        public static string loginName = "";
        public static string loginAutho = "";
        public static string LoginSection = "";
        public static string loginNumber = "";
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        public LOGIN()
        {
            RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Control Panel\International", true); key.SetValue("sShortDate", "MM/dd/yyyy");
            InitializeComponent();
        }

        public void UserVerification()
        {

            if (txtUsername.Text == "" && txtPassword.Text == "")
            {
                Application.ExitThread();
            }
            else if (txtUsername.Text == "")
            {
                MessageBox.Show("ID Number Required!", "Login", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtUsername.Focus();
            }
            else if (txtPassword.Text == "")
            {
                MessageBox.Show("Password Required!", "Login", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtPassword.Focus();
            }
            else
            {
                con.Open();
                SqlDataAdapter sda = new SqlDataAdapter("SELECT COUNT(*) FROM tbl_UserInfo WHERE Username='" + txtUsername.Text + "' AND Password = '" + txtPassword.Text + "'", con);
                DataTable dt = new DataTable();
                sda.Fill(dt);


                if (dt.Rows[0][0].ToString() == "1")
                {
                    SqlCommand cmd = new SqlCommand("select * from tbl_UserInfo where Username ='" + txtUsername.Text + "' AND Password = '" + txtPassword.Text + "'", con);
                    SqlDataReader reader = cmd.ExecuteReader();
                    if (reader.Read())
                    {
                        string login_user;
                        string login_name;
                        string autho;
                        string Section;
                        string login_number;

                        login_user = reader[1].ToString();
                        login_name = reader[3].ToString();
                        autho = reader[6].ToString();
                        Section = reader[5].ToString();
                        login_number = reader[9].ToString();

                        loginuser = login_user;
                        loginName = login_name;
                        loginAutho = autho;
                        LoginSection = Section;
                        loginNumber = login_number;
                    }
                    this.Hide();
                    new Home().Show();
                }
                else
                {

                    MessageBox.Show("Invalid Password.", "Login", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    txtPassword.Text = "";
                    txtPassword.Focus();
                    con.Close();
                }
            }
        }

        private void txtUsername_Enter(object sender, EventArgs e)
        {
            txtUsername.Focus();
            
            if (txtUsername.Text == "Username")
            {
                txtUsername.Text = "";
                txtUsername.ForeColor = Color.FromArgb(26,44,47);
            }
        }

        private void txtUsername_Leave(object sender, EventArgs e)
        {
            if (txtUsername.Text == "")
            {
                txtUsername.Text = "Username";
                txtUsername.ForeColor = Color.DarkGray;
            }
        }

        private void txtPassword_Enter(object sender, EventArgs e)
        {
            //txtPassword.Focus();
            if (txtPassword.Text == "Password")
            {
                txtPassword.Text = "";
                txtPassword.ForeColor = Color.FromArgb(26, 44, 47);
               
            }
            
        }

        private void txtPassword_Leave(object sender, EventArgs e)
        {
            if (txtPassword.Text == "")
            {
                txtPassword.Text = "Password";
                txtPassword.ForeColor = Color.DarkGray;
            }
        }

        private void txtPassword_KeyPress(object sender, KeyPressEventArgs e)
        {
            txtPassword.PasswordChar = '•';
            //UserVerification();
        }

        private void lblClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        static string IHC = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        private void LOGIN_Load(object sender, EventArgs e)
        {
            database.Text = IHC.Replace("Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd", "");
            database.Text = database.Text.Replace("Data Source=", "  Server Name: ");
            database.Text = database.Text.Replace(";Initial Catalog=", "    Database: ");
            version.Text = "AA." + Assembly.GetExecutingAssembly().GetName().Version.ToString();

        }

        private void LOGIN_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ReleaseCapture();
                Common.SendMessage(Handle, Common.WM_NCLBUTTONDOWN, Common.HT_CAPTION, 0);
            }
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            UserVerification();
        }

        private void HidePassword_Click(object sender, EventArgs e)
        {
            if (txtPassword.PasswordChar == '•')
            {
                txtPassword.PasswordChar = '\0';
            }
            else
            {
                txtPassword.PasswordChar = '•';
            }

            ShowPassword.BringToFront();
            HidePassword.SendToBack();
        }

        private void ShowPassword_Click(object sender, EventArgs e)
        {
            if (txtPassword.PasswordChar == '•')
            {
                txtPassword.PasswordChar = '\0';
            }
            else
            {
                txtPassword.PasswordChar = '•';
            }

            ShowPassword.SendToBack();
            HidePassword.BringToFront();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ReleaseCapture();
                Common.SendMessage(Handle, Common.WM_NCLBUTTONDOWN, Common.HT_CAPTION, 0);
            }
        }

        private void txtPassword_KeyDown(object sender, KeyEventArgs e)
        {
           //UserVerification();
        }

        private void lblFGP_Click(object sender, EventArgs e)
        {
            ForgetPass aaa = new ForgetPass();
            aaa.ShowDialog();
        }

        private void r(object sender, EventArgs e)
        {
            Application.Exit();

        }

        private void btnClose_Click_1(object sender, EventArgs e)
        {
 
        }

        private void txtUsername_TabIndexChanged(object sender, EventArgs e)
        {
            //if(txtUsername.TabIndex > 0)
            //{
            //    txtPassword.Select();
            //    txtPassword.Focus();
            //}
        }

        private void btnClose_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
