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
    public partial class EditCalRequest : Form
    {
        public static EditCalRequest _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        public EditCalRequest()
        {
            _instance = this;
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
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

        private void EditCalRequest_Load(object sender, EventArgs e)
        {
            
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ReleaseCapture();
                Common.SendMessage(Handle, Common.WM_NCLBUTTONDOWN, Common.HT_CAPTION, 0);
            }
        }

        private void btnUpdateReq_Click(object sender, EventArgs e)
        {
            if (txtEmpNumber.Text != "" && txtInstrumentLoc.Text != "" && txtInstrumentRefNo.Text != "" && txtManufucturer.Text != "" && txtMinstrumentName.Text != "" && txtModel.Text != "" && txtName.Text != "" && txtSerialNo.Text != "" && txtEmail.Text != "" && cmbCalFreq.Text != "" && cmbCalPlan.Text != "" && txtCostCenter.Text != "" && cmbInstrumentCat.Text != "" && cmbRequestSection.Text != "" && txtInstrumentLoc.Text != "" && dtDueDate.Text != "" && dtReqDate.Text != "")
            {
                cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set InstrumentCat=@instrumentcat, Picture_Link=@piclink, MeasureInstName=@measureinstname, InstRefNo=@instrefno, Manufacturer=@manufacturer, Model=@model, SerialNo=@serialno, InstrumentLoc=@instrumentloc, CalibrationFreq=@calibfreq, CalibrationDueDate=@calibduedate, CalibrationPlan=@calibplan, Remarks=@remarks where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
                con.Open();
                cmd.Parameters.AddWithValue("@instrumentcat", EditCalRequest._instance.cmbInstrumentCat.Text);
                cmd.Parameters.AddWithValue("@measureinstname", txtMinstrumentName.Text);
                cmd.Parameters.AddWithValue("@instrefno", txtInstrumentRefNo.Text);
                cmd.Parameters.AddWithValue("@manufacturer", txtManufucturer.Text);
                cmd.Parameters.AddWithValue("@model", txtModel.Text);
                cmd.Parameters.AddWithValue("@serialno", txtSerialNo.Text);
                cmd.Parameters.AddWithValue("@instrumentloc", txtInstrumentLoc.Text);
                cmd.Parameters.AddWithValue("@calibfreq", EditCalRequest._instance.cmbCalFreq.Text);
                cmd.Parameters.AddWithValue("@calibduedate", dtDueDate.Text);
                cmd.Parameters.AddWithValue("@calibplan", EditCalRequest._instance.cmbCalPlan.Text);
                cmd.Parameters.AddWithValue("@remarks", txtReqRemarks.Text);
                cmd.Parameters.AddWithValue("@piclink", textBox1.Text);
                cmd.ExecuteNonQuery();
                Home._instance.refreshdata2();
                con.Close();
                MessageBox.Show("Calibration Updated Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Close();
            }
            else
            {
                MessageBox.Show("Please input details to update!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            }

        private void btnSubmitReq_Click(object sender, EventArgs e)
        {

        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            // open file dialog   
            OpenFileDialog open = new OpenFileDialog();
            // image filters  
            open.Filter = "Image Files(*.jpg; *.jpeg; *.gif; *.bmp)|*.jpg; *.jpeg; *.gif; *.bmp";
            if (open.ShowDialog() == DialogResult.OK)
            {
                // display image in picture box  
                picInstrument.Image = new Bitmap(open.FileName);
                // image file path  
                textBox1.Text = open.FileName;
            }
        }
    }
}
