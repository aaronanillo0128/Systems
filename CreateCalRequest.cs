using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class CreateCalRequest : Form
    {
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        int ID = 0;
        public CreateCalRequest()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Hide();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ReleaseCapture();
                Common.SendMessage(Handle, Common.WM_NCLBUTTONDOWN, Common.HT_CAPTION, 0);
            }
        }
        private void CreateCalRequest_Load(object sender, EventArgs e)
        {
            InstrumentCat();
            Section();
            CalibrationFreq();
            CalibrationPlan();
            DateTime nm = DateTime.Now;
            string date = nm.ToString("ddHHmmss");
            string cnt = Convert.ToString(date);
            dtReqDate.Text = DateTime.Now.ToString();
            dtDueDate.Text = DateTime.Now.ToString();
        }

        //EMAIL REQUEST CALIB
        public string emailto;
        private string innerString;
        public void audit_email_CreateCalReq()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                string sql = "select distinct Email from tbl_InhouseCalibrationMAIN where Email = '" + txtEmail.Text + "'";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(dt);
                foreach (DataRow dr in dt.Rows)
                {
                    emailto = dr["Email"].ToString();
                    EmailBody_Audit_CreateCalReq();
                    EmailNotif_Audit_CreateCalReq();
                }
            }
        }
        private void EmailNotif_Audit_CreateCalReq()
        {
            MailMessage mail = new MailMessage("inhousecalibrationadmin@brother-biph.com.ph", emailto);
            mail.CC.Add(new MailAddress("christine.sotaso@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("alpie.hatulan@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("mary.mendoza @brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("roise.bordeos@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("kessie.blanca@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("markaljon.opena@brother-biph.com.ph"));
            mail.Bcc.Add(new MailAddress("zzpde178@brother.co.jp"));
            mail.Bcc.Add(new MailAddress("aaronpaul.anillo@brother-biph.com.ph"));
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.UseDefaultCredentials = false;
            client.Host = "10.113.10.1";
            mail.Subject = "【IHC】REQUEST FOR ACCEPTANCE";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);  
        }
        public void EmailBody_Audit_CreateCalReq()
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();
            builder.Append("<h1><font color=Black>Good day!</h1>");
            builder.Append("<b><h2><font color=blue>Request for In-House Calibration has been submitted for <span style='Background-color: Yellow'>ACCEPTANCE.</span></h2></b></font>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";
            //{
            mailBody = "<h3><font color=black>Request Details: </font> " + "" + "</h3><table width='50%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Request Date/Time</th> <th>Section</th> <th>Instrument Reference Number</th><th>Calibration Plan</th>";
            mailBody += "</tr>";

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + DateTime.Now + "</td>";
                mailBody += "<td style='color:black;'>" + cmbSection.Text + "</td>";
                mailBody += "<td style='color:black;'>" + txtInstrumentRefNo.Text + "</td>";
                mailBody += "<td style='color:black;'>" + cmbCalPlan.Text + "</td>";
                mailBody += "</tr>";
            

            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><font color=black>Thank you!</font></h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }

        private void btnSubmitReq_Click(object sender, EventArgs e)
        {
            string getID = "Select TOP 1 ID from tbl_InhouseCalibrationMAIN order by ID desc";

            bool isResult = CRUD.RETRIEVESINGLE(getID);
            if (isResult == true)
            {
                int ID1 = Convert.ToInt16(CRUD.dt.Rows[0]["ID"].ToString()) + 1;
                txtWorkOrderNo.Text = ("BIPH" + "-" + "CAL" + "-" + "23" + "-" + "000" + ID1).Replace(" ", "");
            }

            if (txtEmpNumber.Text != "" && txtInstrumentLoc.Text != "" && txtInstrumentRefNo.Text != "" && txtManufucturer.Text != "" && txtMinstrumentName.Text != "" && txtModel.Text != "" && txtName.Text != "" && txtSerialNo.Text != "" && txtEmail.Text != "" && cmbCalFreq.Text != "" && cmbCalPlan.Text != "" && txtCostCenter.Text != "" && cmbInstrumentCat.Text != "" && cmbSection.Text != "" && txtInstrumentLoc.Text != "" && dtDueDate.Text != "" && dtReqDate.Text != "")
            {
                cmd = new SqlCommand("insert into tbl_InhouseCalibrationMAIN (WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, CalibrationNo, Status, OverallStatus, Remarks, Picture_Link) values (@workorderno, @datereq, @reqsection, @Empno, @reqby, @email, @costcenter, @instrumentcat, @measureinstname, @instrefno, @manufacturer, @model, @serialno, @instrumentloc, @calibreq, @calibduedate, @calibplan, @calibrationno, @status, @overallstat, @remarks, @piclink)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@workorderno", txtWorkOrderNo.Text);
                cmd.Parameters.AddWithValue("@datereq", dtReqDate.Text);
                cmd.Parameters.AddWithValue("@reqsection", cmbSection.Text);
                cmd.Parameters.AddWithValue("@EmpNo", txtEmpNumber.Text);
                cmd.Parameters.AddWithValue("@reqby", txtName.Text);                                                                                                                                                                                                                                                     
                cmd.Parameters.AddWithValue("@email", txtEmail.Text);
                cmd.Parameters.AddWithValue("@costcenter", txtCostCenter.Text);
                cmd.Parameters.AddWithValue("@instrumentcat", cmbInstrumentCat.Text);
                cmd.Parameters.AddWithValue("@measureinstname", txtMinstrumentName.Text);
                cmd.Parameters.AddWithValue("@instrefno", txtInstrumentRefNo.Text);
                cmd.Parameters.AddWithValue("@manufacturer", txtManufucturer.Text);
                cmd.Parameters.AddWithValue("@model", txtModel.Text);
                cmd.Parameters.AddWithValue("@serialno", txtSerialNo.Text);
                cmd.Parameters.AddWithValue("@instrumentloc", txtInstrumentLoc.Text);
                cmd.Parameters.AddWithValue("@calibreq", cmbCalFreq.Text);
                cmd.Parameters.AddWithValue("@calibduedate", dtDueDate.Text);
                cmd.Parameters.AddWithValue("@calibplan", cmbCalPlan.Text);
                cmd.Parameters.AddWithValue("@calibrationno", txtWorkOrderNo.Text);
                cmd.Parameters.AddWithValue("@status", "ONGOING");
                cmd.Parameters.AddWithValue("@overallstat", "FOR ACCEPTANCE");
                cmd.Parameters.AddWithValue("@remarks", txtReqRemarks.Text);
                cmd.Parameters.AddWithValue("@piclink", textBox1.Text);
                cmd.ExecuteNonQuery();
                Home._instance.refreshdata2();
                audit_email_CreateCalReq();
                con.Close();
                MessageBox.Show("Calibration Request Submitted Successfully!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                this.Close();
            }
         
        }

        private void txtEmpNumber_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                con.Open();

                string sqlsss = "Select * from V_EMS where EmpNo = '" + txtEmpNumber.Text + "'";
                SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                SqlDataReader sdrsss = cmdsss.ExecuteReader();
                sdrsss.Read();
                if (sdrsss.HasRows)
                {
                    txtName.Text = sdrsss["Full_Name"].ToString();
                    txtEmail.Text = sdrsss["Email"].ToString();
                    txtCostCenter.Text = sdrsss["CostCode"].ToString();
                }

                con.Close();
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

        public void Section()
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
                cmbSection.DisplayMember = "SectionList";
                cmbSection.ValueMember = "SectionList";
                cmbSection.DataSource = dt.Tables["SectionList"];
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

        private void btnUpdateReq_Click(object sender, EventArgs e)
        {

        }
    }
}

