using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class UpdateCal_Manager : Form
    {
        public static UpdateCal_Manager _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        public UpdateCal_Manager()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void UpdateCal_Manager_Load(object sender, EventArgs e)
        {

        }

        public void InstrumentCat_Manager()
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

        public void CalibrationFreq_Manager()
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

        public void CalibrationPlan_Manager()
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

        public void CalibIncharge_Manager()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select CalibrationInCharge from tbl_CalibrationInCharge";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "CalibrationInCharge");
                cmbCalInCharge.DisplayMember = "CalibrationInCharge";
                cmbCalInCharge.ValueMember = "CalibrationInCharge";
                cmbCalInCharge.DataSource = dt.Tables["CalibrationInCharge"];
            }
        }


        public string emailto;
        private string innerString;

        //FINISHED CALIBRATION - STATUS - APPROVED
        public void audit_email_FINISHEDCALIBRATION()
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
                    EmailBody_Audit_FINISHEDCALIBRATION();
                    EmailNotif_Audit_FINISHEDCALIBRATION();
                }
            }
        }
        private void EmailNotif_Audit_FINISHEDCALIBRATION()
        {
            MailMessage mail = new MailMessage("inhousecalibrationadmin@brother-biph.com.ph", emailto);
            mail.CC.Add(new MailAddress("christine.sotaso@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("alpie.hatulan@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("mary.mendoza @brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("bernard.marquez@brother-biph.com.ph"));
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.UseDefaultCredentials = false;
            client.Host = "10.113.10.1";
            mail.Subject = "【IHC】INSTRUMENT FOR RETRIEVAL";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
        public void EmailBody_Audit_FINISHEDCALIBRATION()
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();
            builder.Append("<h1><font color=black>Good day! </font></h1>");

            builder.Append("<b><h2><span style='Background-color: Green'><font color=blue>INSTRUMENT </span> for In-House Calibration is now <span style='Background-color: Yellow'>FINISHED Calibration.</span></h2></b></font>");
            builder.Append("<b><h2><font color=blue>You may retrieve your instruments at PQC 3D Room.</h2></b></font>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font> " + "" + "</h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <b><th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th><th>Calibrator In-Charge</th></b>";
            mailBody += "</tr>";

            mailBody += "<tr align='Center'>";
            mailBody += "<td style='color:black;'>" + txtInstrumentRefNo.Text + "</td>";
            mailBody += "<td style='color:black;'>" + txtWorkOrderNo.Text + "</td>";
            mailBody += "<td style='color:black;'>" + txtReqSection.Text + "</td>";
            mailBody += "<td style='color:black;'>" + cmbCalInCharge.Text + "</td>";
            mailBody += "</tr>";


            mailBody += "</table>";
            builder.Append("" + mailBody);

            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            innerString = builder.ToString();
        }

        //CALIBRATION DISAPPROVED
        public void audit_email_FINISHEDCALIBRATION_Disapproved()
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
                    EmailBody_Audit_FINISHEDCALIBRATION_Disapproved();
                    EmailNotif_Audit_FINISHEDCALIBRATION_Disapproved();
                }
            }
        }
        private void EmailNotif_Audit_FINISHEDCALIBRATION_Disapproved()
        {
            MailMessage mail = new MailMessage("inhousecalibrationadmin@brother-biph.com.ph", emailto);
            mail.CC.Add(new MailAddress("christine.sotaso@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("alpie.hatulan@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("mary.mendoza @brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("bernard.marquez@brother-biph.com.ph"));
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.UseDefaultCredentials = false;
            client.Host = "10.113.10.1";
            mail.Subject = "[IHC] INSTRUMENT FOR RETRIEVAL";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
        public void EmailBody_Audit_FINISHEDCALIBRATION_Disapproved()
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();
            builder.Append("<h1><font color=black>Good day! </font></h1>");

            builder.Append("<b><h2><font color=blue>INSTRUMENT for In-House Calibration has been </font><span style='Background-color: Yellow'><font color=red>DISAPPROVED.</span></h2></b></font>");
            string mailBody = "<table width='60%' style='border:Solid 1px gray;'>";
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font> " + "" + "</h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <b><th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th><th>Calibrator In-Charge</th></b>";
            mailBody += "</tr>";

            mailBody += "<tr align='Center'>";
            mailBody += "<td style='color:black;'>" + txtInstrumentRefNo.Text + "</td>";
            mailBody += "<td style='color:black;'>" + txtWorkOrderNo.Text + "</td>";
            mailBody += "<td style='color:black;'>" + txtReqSection.Text + "</td>";
            mailBody += "<td style='color:black;'>" + cmbCalInCharge.Text + "</td>";
            mailBody += "</tr>";


            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            innerString = builder.ToString();
        }

        private void btnUpdateStat_Click(object sender, EventArgs e)
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);

            cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set Status = 'ONGOING', CalibrationResult = 'PASS', CertificateStats = 'UPLOADED', CalibrationIncharge = @calcharge, CalibrationStartDate = @calstartdate, CalibrationEndDate = @calenddate, StartTime = @starttime, EndTime = @endtime, OverallStatus = 'FOR 1ST APPROVAL', CalibrationTimehrs = @calibtime, InstrumentCat = @instcat, MeasureInstName = @measureinsname, InstRefNo = @insrefno, Manufacturer = @manufa, Model = @model, SerialNo = @sno, CalibrationFreq = @calfreq, CalibrationDueDate = @calduedate, CalibrationPlan = @calplan, InstrumentLoc = @insloc where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
            con.Open();
            cmd.Parameters.AddWithValue("@calcharge", cmbCalInCharge.Text);
            cmd.Parameters.AddWithValue("@calstartdate", dtCalStartDate.Text);
            cmd.Parameters.AddWithValue("@calenddate", dtCalEndDate.Text);
            cmd.Parameters.AddWithValue("@starttime", cmbStartTime.Text);
            cmd.Parameters.AddWithValue("@endtime", cmbEndTime.Text);
            cmd.Parameters.AddWithValue("@calibtime", txtAutoCompute.Text);
            cmd.Parameters.AddWithValue("@instcat", cmbInstrumentCat.Text);
            cmd.Parameters.AddWithValue("@measureinsname", txtMinstrumentName.Text);
            cmd.Parameters.AddWithValue("@insrefno", txtInstrumentRefNo.Text);
            cmd.Parameters.AddWithValue("@manufa", txtManufucturer.Text);
            cmd.Parameters.AddWithValue("@model", txtModel.Text);
            cmd.Parameters.AddWithValue("@sno", txtSerialNo.Text);
            cmd.Parameters.AddWithValue("@calfreq", cmbCalFreq.Text);
            cmd.Parameters.AddWithValue("@calduedate", dtDueDate.Text);
            cmd.Parameters.AddWithValue("@calplan", cmbCalPlan.Text);
            cmd.Parameters.AddWithValue("@insloc", txtInstrumentLoc.Text);
            cmd.ExecuteNonQuery();
            con.Close();
            this.Close();
            audit_email_FINISHEDCALIBRATION();
            Home._instance.refreshdata4();
            MessageBox.Show("Approved Successfully!");
        }

        private void btnUploadCerti_Click(object sender, EventArgs e)
        {

        }

        private void btnDisapproved_Click(object sender, EventArgs e)
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);

            cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set Status = 'VOID', CalibrationResult = 'FAILED', CertificateStats = 'UPLOADED', CalibrationIncharge = @calcharge, CalibrationStartDate = @calstartdate, CalibrationEndDate = @calenddate, StartTime = @starttime, EndTime = @endtime, OverallStatus = @overallstat, CalibrationTimehrs = @calibtime, InstrumentCat = @instcat, MeasureInstName = @measureinsname, InstRefNo = @insrefno, Manufacturer = @manufa, Model = @model, SerialNo = @sno, CalibrationFreq = @calfreq, CalibrationDueDate = @calduedate, CalibrationPlan = @calplan, InstrumentLoc = @insloc where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
            con.Open();
            cmd.Parameters.AddWithValue("@calcharge", cmbCalInCharge.Text);
            cmd.Parameters.AddWithValue("@calstartdate", dtCalStartDate.Text);
            cmd.Parameters.AddWithValue("@calenddate", dtCalEndDate.Text);
            cmd.Parameters.AddWithValue("@starttime", cmbStartTime.Text);
            cmd.Parameters.AddWithValue("@endtime", cmbEndTime.Text);
            cmd.Parameters.AddWithValue("@calibtime", txtAutoCompute.Text);
            cmd.Parameters.AddWithValue("@instcat", cmbInstrumentCat.Text);
            cmd.Parameters.AddWithValue("@measureinsname", txtMinstrumentName.Text);
            cmd.Parameters.AddWithValue("@insrefno", txtInstrumentRefNo.Text);
            cmd.Parameters.AddWithValue("@manufa", txtManufucturer.Text);
            cmd.Parameters.AddWithValue("@model", txtModel.Text);
            cmd.Parameters.AddWithValue("@sno", txtSerialNo.Text);
            cmd.Parameters.AddWithValue("@calfreq", cmbCalFreq.Text);
            cmd.Parameters.AddWithValue("@ calduedate", dtDueDate.Text);
            cmd.Parameters.AddWithValue("@calplan", cmbCalPlan.Text);
            cmd.Parameters.AddWithValue("@insloc", txtInstrumentLoc.Text);
            cmd.Parameters.AddWithValue("@overallstat", "DISAPPROVED");
            cmd.ExecuteNonQuery();
            con.Close();
            this.Close();
            audit_email_FINISHEDCALIBRATION_Disapproved();
            Home._instance.refreshdata4();
            MessageBox.Show("Instrument Calibration Disapproved!");
        }
    }
}
