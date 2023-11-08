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
    public partial class AddAccount : Form
    {
        public static AddAccount _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        string cs = "Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd";
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        DataTable dt;
        int ID = 0;
        public AddAccount()
        {
            _instance = this;
            InitializeComponent();
        }


        public void refreshdata()
        {
            DisplayData1();
        }

        private void DisplayData1()
        {

            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select * from tbl_UserInfo", con);
          
            adapt.Fill(dt);
            dataGridView.DataSource = dt;
            dataGridView.Columns["Password"].Visible = false;
            dataGridView.Columns["EmpNo"].Visible = false;

        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void AddAccount_Load(object sender, EventArgs e)
        {
            DisplayData1();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtName.Text != "" && txtPass.Text != "" && txtDept.Text != "" && txtSection.Text != "" && cmbAuthority.Text != "")
            {
                string SElectDouble_SQL = string.Empty;
                SElectDouble_SQL = "SELECT ADID from tbl_UserInfo where ADID = '" + txtADID.Text + "'";
                bool DoubleData = false;
                DoubleData = CRUD.RETRIEVESINGLE(SElectDouble_SQL);
                if (DoubleData == true)
                {
                    MessageBox.Show("DUPLICATE ENTRY DETECTED!", "In-House Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    cmd = new SqlCommand("insert into tbl_UserInfo (Username, Password, Name, Department, Section, Authority, ADID, Email,EmpNo) values (@username, @pass, @name, @department, @section, @autho, @adid, @email, @empno)", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.Parameters.AddWithValue("@username", txtADID.Text);
                    cmd.Parameters.AddWithValue("@pass", txtPass.Text);
                    cmd.Parameters.AddWithValue("@name", txtName.Text);
                    cmd.Parameters.AddWithValue("@department", txtDept.Text);
                    cmd.Parameters.AddWithValue("@autho", cmbAuthority.Text);
                    cmd.Parameters.AddWithValue("@section", txtSection.Text);
                    cmd.Parameters.AddWithValue("@adid", txtADID.Text);
                    cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                    cmd.Parameters.AddWithValue("@empno", lblEmpNo.Text);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Account Added Successfully!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    DisplayData1();
                }
               
            }
        }

        private void dataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dataGridView.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtADID.Text = dataGridView.Rows[e.RowIndex].Cells[1].Value.ToString();
            txtPass.Text = dataGridView.Rows[e.RowIndex].Cells[2].Value.ToString();
            txtName.Text = dataGridView.Rows[e.RowIndex].Cells[3].Value.ToString();
            txtDept.Text = dataGridView.Rows[e.RowIndex].Cells[4].Value.ToString();
            txtSection.Text = dataGridView.Rows[e.RowIndex].Cells[5].Value.ToString();
            cmbAuthority.Text = dataGridView.Rows[e.RowIndex].Cells[6].Value.ToString();
            txtEmail.Text = dataGridView.Rows[e.RowIndex].Cells[7].Value.ToString();
            lblEmpNo.Text = dataGridView.Rows[e.RowIndex].Cells[9].Value.ToString();
        }

        private void txtADID_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                con.Open();

                string sqlsss = "Select * from V_EMS where ADID = '" + txtADID.Text + "'";
                SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                SqlDataReader sdrsss = cmdsss.ExecuteReader();
                sdrsss.Read();
                if (sdrsss.HasRows)
                {
                    txtName.Text = sdrsss["Full_Name"].ToString();
                    txtDept.Text = sdrsss["Department"].ToString();
                    txtSection.Text = sdrsss["Section"].ToString();
                    txtEmail.Text = sdrsss["Email"].ToString();
                    lblEmpNo.Text = sdrsss["EmpNo"].ToString();
                }

                con.Close();
            }
        }

        private void btnDel_Click(object sender, EventArgs e)
        {
            if (ID != 0)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to PROCEED?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    cmd = new SqlCommand("delete tbl_UserInfo where ID=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Deleted Successfully!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    DisplayData1();

                }
                else if (dialogResult == DialogResult.No)
                {
                    //do Nothing
                }
                else
                {
                    MessageBox.Show("Please Select Record to Delete", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (txtADID.Text != "" && txtName.Text != "" && txtPass.Text != "" && txtName.Text != "" && txtDept.Text != "" && txtSection.Text != "" && cmbAuthority.Text != "" && txtEmail.Text != "")
            {
                cmd = new SqlCommand("update tbl_UserInfo set Username=@idnum,Password=@pass,Name=@name,Department=@dept,Section=@section,Authority=@autho,Email=@email,ADID=@adid where ID=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@idnum", txtADID.Text);
                cmd.Parameters.AddWithValue("@pass", txtPass.Text);
                cmd.Parameters.AddWithValue("@name", txtName.Text);
                cmd.Parameters.AddWithValue("@dept", txtDept.Text);
                cmd.Parameters.AddWithValue("@section", txtSection.Text);
                cmd.Parameters.AddWithValue("@autho", cmbAuthority.Text);
                cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                cmd.Parameters.AddWithValue("@adid", txtADID.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Updated Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                DisplayData1();

            }
            else
            {
                MessageBox.Show("Please Select Record to Update!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

        }
    }
}
