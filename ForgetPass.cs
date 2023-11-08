using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class ForgetPass : Form
    {
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");

        SqlCommand cmd;
        SqlDataAdapter adapt;
        DataTable dt;
        int ID = 0;
        public ForgetPass()
        {
            InitializeComponent();
        }

        private void btnGO_Click(object sender, EventArgs e)
        {
            con.Open();
            string sqlsss = "Select * from tbl_UserInfo where Email = '" + txtADID.Text + "'";
            SqlCommand cmdsss = new SqlCommand(sqlsss, con);
            SqlDataReader sdrsss = cmdsss.ExecuteReader();
            sdrsss.Read();
            if (sdrsss.HasRows)
            {
                lblName.Text = sdrsss["Name"].ToString();
            }

            con.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (txtADID.Text != "" && txtPass.Text != "")
            {
                cmd = new SqlCommand("update tbl_UserInfo set Password=@pass where Email = '" + txtADID.Text + "'", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@pass", txtPass.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Updated Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                Application.Exit();
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
