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
    public partial class InstrumentReceiving : Form
    {
        public static InstrumentReceiving _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        public InstrumentReceiving()
        {
            _instance = this;
            InitializeComponent();
        }
        public void RequestSec()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select SectionList from tbl_SectionList";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "SectionList");
                cmbRequestSection.DisplayMember = "SectionList";
                cmbRequestSection.ValueMember = "SectionList";
                cmbRequestSection.DataSource = dt.Tables["SectionList"];
            }
        }

        public void InstrumentCat()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select InstrumentCategory from tbl_InstrumentCat";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "InstrumentCategory");
                cmbInstrumentCat.DisplayMember = "InstrumentCategory";
                cmbInstrumentCat.ValueMember = "InstrumentCategory";
                cmbInstrumentCat.DataSource = dt.Tables["InstrumentCategory"];
            }
        }

        public void CalibrationFreq()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select CalibrationFreq from tbl_CalibrationFreq";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "CalibrationFreq");
                cmbCalFreq.DisplayMember = "CalibrationFreq";
                cmbCalFreq.ValueMember = "CalibrationFreq";
                cmbCalFreq.DataSource = dt.Tables["CalibrationFreq"];
            }
        }

        public void CalibrationPlan()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select CalibrationPlan from tbl_CalibrationPlan";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "CalibrationPlan");
                cmbCalPlan.DisplayMember = "CalibrationPlan";
                cmbCalPlan.ValueMember = "CalibrationPlan";
                cmbCalPlan.DataSource = dt.Tables["CalibrationPlan"];
            }
        }
        private void InstrumentReceiving_Load(object sender, EventArgs e)
        { 
            dateTimePicker1.Text = DateTime.Now.ToString("MM/dd/yyyy");

            if(lblStatstat.Text != "NOT YET RECEIVE")
            {
                btnUpdatestat.Enabled = false;
                tbEmpNo.Enabled = false;
                tbName.Enabled = false;
                txtEmpNo.Enabled = false;
                txtName1.Enabled = false;
                dateTimePicker1.Enabled = false;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            tbName.Text = "";
            txtName1.Text = "";
            tbEmpNo.Text = "";
            txtEmpNo.Text = "";
            this.Close();
        }

        private void tbEmpNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                con.Open();

                string sqlsss = "Select * from V_EMS where EmpNo = '" + tbEmpNo.Text + "'";
                SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                SqlDataReader sdrsss = cmdsss.ExecuteReader();
                sdrsss.Read();
                if (sdrsss.HasRows)
                {
                    tbName.Text = sdrsss["Full_Name"].ToString();
                }

                con.Close();
            }
        }

        private void txtEmpNo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                con.Open();

                string sqlsss = "Select * from V_EMS where EmpNo = '" + txtEmpNo.Text + "'";
                SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                SqlDataReader sdrsss = cmdsss.ExecuteReader();
                sdrsss.Read();
                if (sdrsss.HasRows)
                {
                    txtName1.Text = sdrsss["Full_Name"].ToString();
                }

                con.Close();
            }
        }

        private void btnUpdatestat_Click(object sender, EventArgs e)
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);

            cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set DateReceived = @daterec, EndorsedBy = @endorseby, ReceivedBy = @receivedby, OverallStatus = 'FOR CALIBRATION' where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
            con.Open();
            cmd.Parameters.AddWithValue("@daterec", dateTimePicker1.Text);
            cmd.Parameters.AddWithValue("@endorseby", txtName.Text);
            cmd.Parameters.AddWithValue("@receivedby", txtName1.Text);
            cmd.ExecuteNonQuery();
            con.Close();
            this.Close();
         
            MessageBox.Show("Successfully Received!");
            Home._instance.refreshdata4();
            tbName.Text = "";
            txtName1.Text = "";
            tbEmpNo.Text = "";
            txtEmpNo.Text = "";

        }
    }
}
