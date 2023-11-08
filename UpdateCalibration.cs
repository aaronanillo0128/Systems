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
    public partial class UpdateCalibration : Form
    {
        public static UpdateCalibration _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        public UpdateCalibration()
        {
            _instance = this;
            InitializeComponent();
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

        public void CalibIncharge()
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

        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        string input;
        private void btnDownloadForm_Click(object sender, EventArgs e)
        {
            string dir = @"C:\EXPORT\" + cmbInstrumentCat.Text;
            if (!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            if (SelectForm.Text != "")
            {
                oXL = new Microsoft.Office.Interop.Excel.Application();

                if (cmbInstrumentCat.Text.ToUpper() == "Caliper".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                else if (cmbInstrumentCat.Text.ToUpper() == "Micrometer".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                else if (cmbInstrumentCat.Text.ToUpper() == "Dial Test Indicator".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                else if (cmbInstrumentCat.Text.ToUpper() == "Pin Gauge".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                else if (cmbInstrumentCat.Text.ToUpper() == "Height Gauge".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                else if (cmbInstrumentCat.Text.ToUpper() == "Dial Gauge".ToUpper())
                {
                    oWB = oXL.Workbooks.Open(lolo + @"\" + SelectForm.Text + ".xlsx");
                }
                oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                int intRow = 0;
                //Row Column

                oSheet.Cells[4, 3] = txtWorkOrderNo.Text;
                oSheet.Cells[9, 5] = txtMinstrumentName.Text;
                oSheet.Cells[10, 5] = txtManufucturer.Text;
                oSheet.Cells[11, 5] = txtModel.Text;
                oSheet.Cells[12, 5] = txtSerialNo.Text;
                oSheet.Cells[13, 5] = txtInstrumentRefNo.Text;
                oSheet.Cells[9, 11] = txtReqSection.Text;
                oSheet.Cells[10, 11] = dtCalStartDate.Text;
                oSheet.Cells[11, 11] = dtCalEndDate.Text;

                oSheet.Cells[42, 2] = LOGIN.loginName;

                oXL.Visible = false;
                oXL.UserControl = true;
                try
                {
                    Microsoft.Office.Interop.Excel.Range oRanges = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[37, 2];
                    float Lefts = (float)((double)oRanges.Left) + 40;
                    float Tops = (float)((double)oRanges.Top);
                    const float ImageSizes = 80;

                    oSheet.Shapes.AddPicture(@"\\Apbiphbpswb01\ihcs$\" + Home.lblIDNUM + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Lefts, Tops, ImageSizes, ImageSizes);
                }
                catch
                {
                    oSheet.Cells[41, 3] = "NOE - " + Home.lblIDNUM;
                }
               

                string location = @"C:\EXPORT\" + cmbInstrumentCat.Text + "\\" + txtWorkOrderNo.Text + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + ".xlsx";
                oSheet.SaveAs(location);
                oXL.Quit();

                System.Diagnostics.Process.Start(location);
            }

            else
            {
                MessageBox.Show("Please select Certicate Name");
            }

        }
    
        

        private void btnUploadCerti_Click(object sender, EventArgs e)
        {
            UploadCertificate eron = new UploadCertificate();

            eron.txtWorkOrderNo.Text = txtWorkOrderNo.Text;

            eron.ShowDialog();
        }

        public string emailto;
        private string innerString;

        //FINISHED CALIBRATION - STATUS
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
            mail.CC.Add(new MailAddress("roise.bordeos@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("kessie.blanca@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("markaljon.opena@brother-biph.com.ph"));
            mail.Bcc.Add(new MailAddress("zzpde178@brother.co.jp"));
            mail.Bcc.Add(new MailAddress("aaronpaul.anillo@brother-biph.com.ph"));
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

            builder.Append("<b><h2><span style='Background-color: LightGreen'><font color=blue>INSTRUMENT </span> for In-House Calibration is now <span style='Background-color: Yellow'>FINISHED Calibration.</span></h2></b></font>");
            builder.Append("<b><h2><font color=blue>You may retrieve your instruments at PQC 3D Room.</h2></b></font>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font> " + "" + "</h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <b><th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th><th>Calibrator In-Charge</th></b>";
            mailBody += "</tr>";

            mailBody += "<tr align='Center'>";
            mailBody += "<td style='color:blue;'>" + txtInstrumentRefNo.Text + "</td>";
            mailBody += "<td style='color:blue;'>" + txtWorkOrderNo.Text + "</td>";
            mailBody += "<td style='color:blue;'>" + txtReqSection.Text + "</td>";
            mailBody += "<td style='color:blue;'>" + cmbCalInCharge.Text + "</td>";
            mailBody += "</tr>";


            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }
        private void btnUpdateStat_Click(object sender, EventArgs e)
        {
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);

            cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set CertificateName = @certname, CalibrationIncharge = @calcharge, CalibrationStartDate = @calstartdate, CalibrationEndDate = @calenddate, StartTime = @starttime, EndTime = @endtime, OverallStatus = @overallstat, CalibrationTimehrs = @calibtime, InstrumentCat = @instcat, CalibrationResult = @calresult, MeasureInstName = @measureinsname, InstRefNo = @insrefno,  Manufacturer = @manufa, Model = @model, SerialNo = @sno, CalibrationFreq = @calfreq, CalibrationDueDate = @calduedate, CalibrationPlan = @calplan, InstrumentLoc = @insloc, Remarks = @remarks where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
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
            cmd.Parameters.AddWithValue("@certname", SelectForm.Text);
            cmd.Parameters.AddWithValue("@overallstat", comboBox1.Text);
            cmd.Parameters.AddWithValue("@calresult", cmbCalResult.Text);
            cmd.Parameters.AddWithValue("@remarks", txtReqRemarks.Text);
            cmd.ExecuteNonQuery();
            con.Close();
            this.Close();
            if(comboBox1.Text == "FOR 1ST APPROVAL")
            {
                audit_email_FINISHEDCALIBRATION();
            }
            
            Home._instance.refreshdata4();
            Home._instance.refreshdata5();
            MessageBox.Show("Update Successfully!");
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UpdateCalibration_Load(object sender, EventArgs e)
        {
            //if (lblStat.Text == "DISAPPROVED")
            //{
            //    btnUpdateStat.Enabled = false;
            //} 
        }

       

        private void txtAutoCompute_TextChanged(object sender, EventArgs e)
        {

        
        }

        private void btnCalculate_Click(object sender, EventArgs e)
        {
            TimeSpan ts1 = TimeSpan.Parse(cmbStartTime.Text); //"1:35"
            TimeSpan ts2 = TimeSpan.Parse(cmbEndTime.Text); //"3:30"
            TimeSpan Calcu = TimeSpan.Parse(cmbEndTime.Text)- TimeSpan.Parse(cmbStartTime.Text) ;
            txtAutoCompute.Text = DateTime.Parse(Calcu.ToString()).ToString("HH:mm").Replace(":",".");
        }
        string lolo;
        private void cmbInstrumentCat_TextChanged(object sender, EventArgs e)
        {


            if (cmbInstrumentCat.Text.ToUpper() == "Caliper".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\SC (4) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);
                }
            }
        

            else if (cmbInstrumentCat.Text.ToUpper() == "Micrometer".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\M1000 (4) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);

                }
            }

            else if (cmbInstrumentCat.Text.ToUpper() == "Height Gauge".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\HG (3) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);

                }
            }

            else if (cmbInstrumentCat.Text.ToUpper() == "Dial Test Indicator".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\DTI (7) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);

                }
            }

            else if (cmbInstrumentCat.Text.ToUpper() == "Dial Gauge".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\DG (2) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);
  
                }
            }

            else if (cmbInstrumentCat.Text.ToUpper() == "Pin Gauge".ToUpper())
            {
                lolo = @"\\apbiphsh07\D0_ShareBrotherGroup\15_PQC\01 IQC\02. Inspection\01. Inspection Measurement\FY2022 In-House Calibration System Continuation Modification Request\01 Report\System Overview\Certificate Formats\PIN (2) ok";
                string filePaths1 = lolo;
                string[] files = System.IO.Directory.GetFiles(filePaths1);
                for (int i = 0; i < files.Length; i++)
                {
                    input = Path.GetFileName(files[i].Replace(".xlsx", ""));
                    SelectForm.Items.Add(input);
      
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            SelectForm.Text = "";
            SelectForm.Items.Remove(input);
        }
    }
    }

