using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class Home : Form
    {
        public static Home _instance;
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");

        SqlConnection con2 = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlDataAdapter adapt;
        int ID;
        public Home()
        {
            InitializeComponent();
            _instance = this;
        }

        public void refreshdata1()
        {
            DisplayData1();
        }
        public void refreshdata2()
        {
            DisplayData2();
        }
        public void refreshdata3()
        {
            DisplayData3();
        }
        public void refreshdata4()
        {
            DisplayData4();
        }

        public void refreshdata5()
        {
            DisplayData5();
        }
        public void refreshdata6()
        {
            DisplayData6();
        }

        private void AddCheckedBoxColumn()
        {
            // Add checkbox column in datagrid            
            DataGridViewCheckBoxColumn checkColumn = new DataGridViewCheckBoxColumn();
            checkColumn.Name = "Select";
            checkColumn.HeaderText = "Select";
            checkColumn.ReadOnly = false;
            checkColumn.FillWeight = 10; //if the datagridview is resized (on form resize) the checkbox won't take up too much; value is relative to the other columns' fill values            
            dataGridView3.Columns.Add(checkColumn);
            checkColumn.DisplayIndex = 0;
            checkColumn.Frozen = true;
        }

        private void AddCheckedBoxColumn1()
        {
            // Add checkbox column in datagrid            
            DataGridViewCheckBoxColumn checkColumn1 = new DataGridViewCheckBoxColumn();
            checkColumn1.Name = "Select";
            checkColumn1.HeaderText = "Select";
            checkColumn1.ReadOnly = false;
            checkColumn1.FillWeight = 10; //if the datagridview is resized (on form resize) the checkbox won't take up too much; value is relative to the other columns' fill values            
            dataGridView5.Columns.Add(checkColumn1);
            checkColumn1.DisplayIndex = 0;
            checkColumn1.Frozen = true;
        }




        private void Home_Load(object sender, EventArgs e)
        {
            timer1.Start();
            SelectSection1();
            SelectCalMonth1();
            SelectStats();
            DisplayData1();
            DisplayData2();
            DisplayData3();
            AddCheckedBoxColumn();
            AddCheckedBoxColumn1();
            DisplayData4();
            DisplayData5();
            DisplayData6();
            tabPage1.Text = "HOME " + "(" + dataGridView1.RowCount.ToString() + ")";
            tabPage2.Text = "CALIBRATION REQUEST " + "(" + dataGridView2.RowCount.ToString() + ")";
            tabPage3.Text = "ACCEPTANCE APPROVAL " + "(" + dataGridView3.RowCount.ToString() + ")";
            tabPage4.Text = "ACTUAL CALIBRATION " + "(" + dataGridView4.RowCount.ToString() + ")";
            tabPage5.Text = "CERTIFICATE " + "(" + dataGridView5.RowCount.ToString() + ")";
            tabPage6.Text = "MH SUMMARY " + "(" + dataGridView6.RowCount.ToString() + ")";

            if (LOGIN.loginuser != null)
            {
                lblUser.Text = LOGIN.loginuser;
                lblSection.Text = LOGIN.LoginSection;
                lblAutho.Text = LOGIN.loginAutho;
                lblEmpNo.Text = LOGIN.loginNumber;
            }

            if (LOGIN.loginAutho == "Requestor")
            {
                tabPage4.Enabled = false;
                tabPage3.Enabled = false;
                lblNoAccess.Visible = true;
                lblNoAccess2.Visible = true;
                LinkAccount.Enabled = false;
                dataGridView5.Columns["Select"].ReadOnly = true;
                btnChecker.Enabled = false;
                btnApprover.Enabled = false;
                btnRecieving.Enabled = false;
                btnUpdateDetails.Enabled = false;
                btnTempCert.Visible = false;
                btnUpdateDetail2.Enabled = false;
            }
            if (LOGIN.loginAutho == "Manager")
            {
                btnCreateCalibrationReq.Enabled = false;
                btnEditReq.Enabled = false;
                btnUpdateDetails.Enabled = false;
                btnRecieving.Enabled = false;
            }


            //if(LOGIN.loginAutho == "Manager")
            //{
            //    btnApproveManager.Visible = true;
            //    btnDisapproveManager.Visible = true;
            //    btnApprove.Visible = false;
            //    btnDisapprove.Visible = false;

            //}
            lblUsers = lblUser.Text;
            lblIDNUM = lblEmpNo.Text;

        }
        public static string lblUsers;
        public static string lblIDNUM;
        //HOME
        private void DisplayData1()
        {
            if (LOGIN.loginAutho == "Admin")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns["EmpNo"].Visible = false;
                dataGridView1.Columns["Email"].Visible = false;
                dataGridView1.Columns["InstrumentLoc"].Visible = false;
                dataGridView1.Columns["Approved_Link"].Visible = false;
                dataGridView1.Columns["PDF_Link"].Visible = false;
                dataGridView1.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView1.Columns["Picture_Link"].Visible = false;
                dataGridView1.Columns["CertificateName"].Visible = false;
                this.dataGridView1.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView1.Columns[0].HeaderText = "No.";
                dataGridView1.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView1.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView1.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Supervisor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns["EmpNo"].Visible = false;
                dataGridView1.Columns["Email"].Visible = false;
                dataGridView1.Columns["InstrumentLoc"].Visible = false;
                dataGridView1.Columns["Approved_Link"].Visible = false;
                dataGridView1.Columns["PDF_Link"].Visible = false;
                dataGridView1.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView1.Columns["Picture_Link"].Visible = false;
                dataGridView1.Columns["CertificateName"].Visible = false;
                this.dataGridView1.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView1.Columns[0].HeaderText = "No.";
                dataGridView1.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView1.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView1.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns["EmpNo"].Visible = false;
                dataGridView1.Columns["Email"].Visible = false;
                dataGridView1.Columns["InstrumentLoc"].Visible = false;
                dataGridView1.Columns["Approved_Link"].Visible = false;
                dataGridView1.Columns["PDF_Link"].Visible = false;
                dataGridView1.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView1.Columns["Picture_Link"].Visible = false;
                dataGridView1.Columns["CertificateName"].Visible = false;
                this.dataGridView1.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView1.Columns[0].HeaderText = "No.";
                dataGridView1.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView1.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView1.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Requestor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + LOGIN.LoginSection + "'", con);
                adapt.Fill(dt);
                dataGridView1.DataSource = dt;
                dataGridView1.Columns["EmpNo"].Visible = false;
                dataGridView1.Columns["Email"].Visible = false;
                dataGridView1.Columns["InstrumentLoc"].Visible = false;
                dataGridView1.Columns["Approved_Link"].Visible = false;
                dataGridView1.Columns["PDF_Link"].Visible = false;
                dataGridView1.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView1.Columns["Picture_Link"].Visible = false;
                dataGridView1.Columns["CertificateName"].Visible = false;
                this.dataGridView1.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView1.Columns[0].HeaderText = "No.";
                dataGridView1.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView1.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView1.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
        }
        //CALIBRATION REQUEST
        private void DisplayData2()
        {
            if (LOGIN.loginAutho == "Admin")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, Picture_Link, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE' OR OverallStatus = 'FOR FINAL ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView2.DataSource = dt;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["EmpNo"].Visible = false;
                dataGridView2.Columns["Email"].Visible = false;
                dataGridView2.Columns["InstrumentLoc"].Visible = false;
                dataGridView2.Columns["DateReceived"].Visible = false;
                dataGridView2.Columns["EndorsedBy"].Visible = false;
                dataGridView2.Columns["ReceivedBy"].Visible = false;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["Picture_Link"].Visible = false;
                this.dataGridView2.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView2.Columns[0].HeaderText = "No.";
                dataGridView2.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView2.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView2.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            else if (LOGIN.loginAutho == "Supervisor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, Picture_Link, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE' OR OverallStatus = 'FOR FINAL ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView2.DataSource = dt;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["EmpNo"].Visible = false;
                dataGridView2.Columns["Email"].Visible = false;
                dataGridView2.Columns["InstrumentLoc"].Visible = false;
                dataGridView2.Columns["DateReceived"].Visible = false;
                dataGridView2.Columns["EndorsedBy"].Visible = false;
                dataGridView2.Columns["ReceivedBy"].Visible = false;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["Picture_Link"].Visible = false;
                this.dataGridView2.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView2.Columns[0].HeaderText = "No.";
                dataGridView2.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView2.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView2.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID,Picture_Link, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE' OR OverallStatus = 'FOR FINAL ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView2.DataSource = dt;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["EmpNo"].Visible = false;
                dataGridView2.Columns["Email"].Visible = false;
                dataGridView2.Columns["InstrumentLoc"].Visible = false;
                dataGridView2.Columns["DateReceived"].Visible = false;
                dataGridView2.Columns["EndorsedBy"].Visible = false;
                dataGridView2.Columns["ReceivedBy"].Visible = false;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["Picture_Link"].Visible = false;
                this.dataGridView2.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView2.Columns[0].HeaderText = "No.";
                dataGridView2.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView2.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView2.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Requestor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, Picture_Link, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = '" + LOGIN.LoginSection + "' AND (OverallStatus = 'FOR ACCEPTANCE' OR OverallStatus = 'FOR FINAL ACCEPTANCE')", con);
                adapt.Fill(dt);
                dataGridView2.DataSource = dt;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["EmpNo"].Visible = false;
                dataGridView2.Columns["Email"].Visible = false;
                dataGridView2.Columns["InstrumentLoc"].Visible = false;
                dataGridView2.Columns["DateReceived"].Visible = false;
                dataGridView2.Columns["EndorsedBy"].Visible = false;
                dataGridView2.Columns["ReceivedBy"].Visible = false;
                dataGridView2.Columns["Remarks"].Visible = false;
                dataGridView2.Columns["Picture_Link"].Visible = false;
                this.dataGridView2.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView2.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView2.Columns[0].HeaderText = "No.";
                dataGridView2.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView2.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView2.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

        }

        //ACCEPTANCE APPROVAL

        private void DisplayData3()
        {
            if (LOGIN.loginAutho == "Admin")
            {

                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView3.DataSource = dt;
                dataGridView3.Columns["Remarks"].Visible = false;
                dataGridView3.Columns["EmpNo"].Visible = false;
                dataGridView3.Columns["ID"].Visible = false;
                dataGridView3.Columns["Email"].Visible = false;
                dataGridView3.Columns["InstrumentLoc"].Visible = false;
                dataGridView3.Columns["DateReceived"].Visible = false;
                dataGridView3.Columns["EndorsedBy"].Visible = false;
                dataGridView3.Columns["ReceivedBy"].Visible = false;
                this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                //dataGridView3.Columns[1].HeaderText = "Select";
                dataGridView3.Columns[0].HeaderText = "Select";
                dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
                //dataGridView3.Columns["Select"].ReadOnly = true;

            }
            else if (LOGIN.loginAutho == "Supervisor")
            {

                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView3.DataSource = dt;
                dataGridView3.Columns["Remarks"].Visible = false;
                dataGridView3.Columns["EmpNo"].Visible = false;
                dataGridView3.Columns["ID"].Visible = false;
                dataGridView3.Columns["Email"].Visible = false;
                dataGridView3.Columns["InstrumentLoc"].Visible = false;
                dataGridView3.Columns["DateReceived"].Visible = false;
                dataGridView3.Columns["EndorsedBy"].Visible = false;
                dataGridView3.Columns["ReceivedBy"].Visible = false;
                this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                //dataGridView3.Columns[1].HeaderText = "Select";
                dataGridView3.Columns[0].HeaderText = "Select";
                dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
                //dataGridView3.Columns["Select"].ReadOnly = true;

            }

            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL ACCEPTANCE'", con);
                adapt.Fill(dt);
                dataGridView3.DataSource = dt;
                dataGridView3.Columns["Remarks"].Visible = false;
                dataGridView3.Columns["EmpNo"].Visible = false;
                dataGridView3.Columns["ID"].Visible = false;
                dataGridView3.Columns["Email"].Visible = false;
                dataGridView3.Columns["InstrumentLoc"].Visible = false;
                dataGridView3.Columns["DateReceived"].Visible = false;
                dataGridView3.Columns["EndorsedBy"].Visible = false;
                dataGridView3.Columns["ReceivedBy"].Visible = false;
                this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                //dataGridView3.Columns[1].HeaderText = "Select";
                dataGridView3.Columns[0].HeaderText = "Select";
                dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Requestor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = '" + LOGIN.LoginSection + "' AND (OverallStatus = 'FOR FINAL ACCEPTANCE' OR OverallStatus = 'FOR ACCEPTANCE')", con);
                adapt.Fill(dt);
                dataGridView3.DataSource = dt;
                dataGridView3.Columns["Remarks"].Visible = false;
                dataGridView3.Columns["EmpNo"].Visible = false;
                dataGridView3.Columns["ID"].Visible = false;
                dataGridView3.Columns["Email"].Visible = false;
                dataGridView3.Columns["InstrumentLoc"].Visible = false;
                dataGridView3.Columns["DateReceived"].Visible = false;
                dataGridView3.Columns["EndorsedBy"].Visible = false;
                dataGridView3.Columns["ReceivedBy"].Visible = false;
                this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView3.Columns[0].HeaderText = "Select";
                dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

        }
        //ACTUAL CALIBRATION
        private void DisplayData4()
        {
            if (LOGIN.loginAutho == "Admin")
            {

                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where (OverallStatus = 'FOR CALIBRATION' OR OverallStatus = 'NOT YET RECEIVE' OR OverallStatus = 'ON-GOING CALIBRATION' OR OverallStatus = 'DISAPPROVED')", con);
                adapt.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.Columns["EmpNo"].Visible = false;
                dataGridView4.Columns["Email"].Visible = false;
                dataGridView4.Columns["InstrumentLoc"].Visible = false;
                dataGridView4.Columns["Approved_Link"].Visible = false;
                dataGridView4.Columns["PDF_Link"].Visible = false;
                dataGridView4.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView4.Columns["Picture_Link"].Visible = false;
                dataGridView4.Columns["CertificateName"].Visible = false;
                this.dataGridView4.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView4.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView4.Columns[0].HeaderText = "No.";
                dataGridView4.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView4.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView4.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Supervisor")
            {

                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where (OverallStatus = 'FOR CALIBRATION' OR OverallStatus = 'NOT YET RECEIVE' OR OverallStatus = 'ON-GOING CALIBRATION' OR OverallStatus = 'DISAPPROVED')", con);
                adapt.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.Columns["EmpNo"].Visible = false;
                dataGridView4.Columns["Email"].Visible = false;
                dataGridView4.Columns["InstrumentLoc"].Visible = false;
                dataGridView4.Columns["Approved_Link"].Visible = false;
                dataGridView4.Columns["PDF_Link"].Visible = false;
                dataGridView4.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView4.Columns["Picture_Link"].Visible = false;
                dataGridView4.Columns["CertificateName"].Visible = false;
                this.dataGridView4.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView4.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView4.Columns[0].HeaderText = "No.";
                dataGridView4.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView4.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView4.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where (OverallStatus = 'FOR CALIBRATION' OR OverallStatus = 'NOT YET RECEIVE' OR OverallStatus = 'ON-GOING CALIBRATION' OR OverallStatus = 'DISAPPROVED')", con);
                adapt.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.Columns["EmpNo"].Visible = false;
                dataGridView4.Columns["Email"].Visible = false;
                dataGridView4.Columns["InstrumentLoc"].Visible = false;
                dataGridView4.Columns["Approved_Link"].Visible = false;
                dataGridView4.Columns["PDF_Link"].Visible = false;
                dataGridView4.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView4.Columns["Picture_Link"].Visible = false;
                dataGridView4.Columns["CertificateName"].Visible = false;
                this.dataGridView4.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView4.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView4.Columns[0].HeaderText = "No.";
                dataGridView4.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView4.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView4.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Requestor")
            {

                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + LOGIN.LoginSection + "' AND (OverallStatus = 'FOR CALIBRATION' OR OverallStatus = 'ON-GOING CALIBRATION' OR OverallStatus = 'NOT YET RECEIVE'  OR OverallStatus = 'DISAPPROVED')", con);
                adapt.Fill(dt);
                dataGridView4.DataSource = dt;
                dataGridView4.Columns["EmpNo"].Visible = false;
                dataGridView4.Columns["Email"].Visible = false;
                dataGridView4.Columns["InstrumentLoc"].Visible = false;
                dataGridView4.Columns["Approved_Link"].Visible = false;
                dataGridView4.Columns["PDF_Link"].Visible = false;
                dataGridView4.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView4.Columns["Picture_Link"].Visible = false;
                dataGridView4.Columns["CertificateName"].Visible = false;
                this.dataGridView4.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView4.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView4.Columns[0].HeaderText = "No.";
                dataGridView4.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView4.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView4.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

        }

        //CERTIFICATE TAB
        private void DisplayData5()
        {
            if (LOGIN.loginAutho == "Admin")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR 1ST APPROVAL'", con);
                adapt.Fill(dt);
                dataGridView5.DataSource = dt;
                dataGridView5.Columns["EmpNo"].Visible = false;
                dataGridView5.Columns["Email"].Visible = false;
                dataGridView5.Columns["InstrumentLoc"].Visible = false;
                dataGridView5.Columns["Approved_Link"].Visible = false;
                dataGridView5.Columns["PDF_Link"].Visible = false;
                dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView5.Columns["Picture_Link"].Visible = false;
                dataGridView5.Columns["CertificateName"].Visible = false;
                this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView5.Columns[0].HeaderText = "No.";
                dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
            else if (LOGIN.loginAutho == "Supervisor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR 1ST APPROVAL'", con);
                adapt.Fill(dt);
                dataGridView5.DataSource = dt;
                dataGridView5.Columns["EmpNo"].Visible = false;
                dataGridView5.Columns["Email"].Visible = false;
                dataGridView5.Columns["InstrumentLoc"].Visible = false;
                dataGridView5.Columns["Approved_Link"].Visible = false;
                dataGridView5.Columns["PDF_Link"].Visible = false;
                dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView5.Columns["Picture_Link"].Visible = false;
                dataGridView5.Columns["CertificateName"].Visible = false;
                this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView5.Columns[0].HeaderText = "No.";
                dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL APPROVAL'", con);
                adapt.Fill(dt);
                dataGridView5.DataSource = dt;
                dataGridView5.Columns["EmpNo"].Visible = false;
                dataGridView5.Columns["Email"].Visible = false;
                dataGridView5.Columns["InstrumentLoc"].Visible = false;
                dataGridView5.Columns["Approved_Link"].Visible = false;
                dataGridView5.Columns["PDF_Link"].Visible = false;
                dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView5.Columns["Picture_Link"].Visible = false;
                dataGridView5.Columns["CertificateName"].Visible = false;
                this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView5.Columns[0].HeaderText = "No.";
                dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            if (LOGIN.loginAutho == "Requestor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN ReqSection = '" + LOGIN.LoginSection + "' AND OverallStatus = 'COMPLETED : CERTIFICATE READY'", con);
                adapt.Fill(dt);
                dataGridView5.DataSource = dt;
                dataGridView5.Columns["EmpNo"].Visible = false;
                dataGridView5.Columns["Email"].Visible = false;
                dataGridView5.Columns["InstrumentLoc"].Visible = false;
                dataGridView5.Columns["Approved_Link"].Visible = false;
                dataGridView5.Columns["PDF_Link"].Visible = false;
                dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView5.Columns["Picture_Link"].Visible = false;
                dataGridView5.Columns["CertificateName"].Visible = false;
                this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView5.Columns[0].HeaderText = "No.";
                dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

        }

        //MH TAB
        private void DisplayData6()
        {
            if (LOGIN.loginAutho == "Admin")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView6.DataSource = dt;
                dataGridView6.Columns["EmpNo"].Visible = false;
                dataGridView6.Columns["Email"].Visible = false;
                dataGridView6.Columns["InstrumentLoc"].Visible = false;
                dataGridView6.Columns["Approved_Link"].Visible = false;
                dataGridView6.Columns["PDF_Link"].Visible = false;
                dataGridView6.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView6.Columns["Picture_Link"].Visible = false;
                dataGridView6.Columns["CertificateName"].Visible = false;
                this.dataGridView6.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView6.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView6.Columns[0].HeaderText = "No.";
                dataGridView6.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView6.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView6.Columns["CalibrationFreq"].HeaderText = "Frequency";

            }
            else if (LOGIN.loginAutho == "Supervisor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView6.DataSource = dt;
                dataGridView6.Columns["EmpNo"].Visible = false;
                dataGridView6.Columns["Email"].Visible = false;
                dataGridView6.Columns["InstrumentLoc"].Visible = false;
                dataGridView6.Columns["Approved_Link"].Visible = false;
                dataGridView6.Columns["PDF_Link"].Visible = false;
                dataGridView6.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView6.Columns["Picture_Link"].Visible = false;
                dataGridView6.Columns["CertificateName"].Visible = false;
                this.dataGridView6.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView6.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView6.Columns[0].HeaderText = "No.";
                dataGridView6.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView6.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView6.Columns["CalibrationFreq"].HeaderText = "Frequency";

            }
            else if (LOGIN.loginAutho == "Manager")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN", con);
                adapt.Fill(dt);
                dataGridView6.DataSource = dt;
                dataGridView6.Columns["EmpNo"].Visible = false;
                dataGridView6.Columns["Email"].Visible = false;
                dataGridView6.Columns["InstrumentLoc"].Visible = false;
                dataGridView6.Columns["Approved_Link"].Visible = false;
                dataGridView6.Columns["PDF_Link"].Visible = false;
                dataGridView6.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView6.Columns["Picture_Link"].Visible = false;
                dataGridView6.Columns["CertificateName"].Visible = false;
                this.dataGridView6.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView6.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView6.Columns[0].HeaderText = "No.";
                dataGridView6.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView6.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView6.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }

            else if (LOGIN.loginAutho == "Requestor")
            {
                DataTable dt = new DataTable();
                adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + LOGIN.LoginSection + "'", con);
                adapt.Fill(dt);
                dataGridView6.DataSource = dt;
                dataGridView6.Columns["EmpNo"].Visible = false;
                dataGridView6.Columns["Email"].Visible = false;
                dataGridView6.Columns["InstrumentLoc"].Visible = false;
                dataGridView6.Columns["Approved_Link"].Visible = false;
                dataGridView6.Columns["PDF_Link"].Visible = false;
                dataGridView6.Columns["CertificateReleaseDate"].Visible = false;
                dataGridView6.Columns["Picture_Link"].Visible = false;
                dataGridView6.Columns["CertificateName"].Visible = false;
                this.dataGridView6.DefaultCellStyle.Font = new Font("Calibri", 8);
                dataGridView6.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
                dataGridView6.Columns[0].HeaderText = "No.";
                dataGridView6.Columns["CostCenter"].HeaderText = "CostCode";
                dataGridView6.Columns["CalibrationDueDate"].HeaderText = "Due Date";
                dataGridView6.Columns["CalibrationFreq"].HeaderText = "Frequency";
            }
        }

        //DATAGRIDVIEW 1
        //SEARCH ....
        public void searchACC()
        {
           // dataGridView1.DataSource = null;
           
          
        }


        public void searchAC3()
        {
            dataGridView2.DataSource = null;
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where WorkOrderNo like @workorderno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@workorderno", "%" + tbWorkOrder2.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                sda.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
                dataGridView2.Columns[0].HeaderText = "No.";

            }
        }

        public void searchACC4()
        {

            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where InstRefNo like @instrefno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@instrefno", "%" + tbInstRefNo2.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView2.DataSource = ds.Tables[0];
                dataGridView2.Columns[0].HeaderText = "No.";

            }
        }

        public void searchACC5()
        {
            dataGridView4.DataSource = null;
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where WorkOrderNo like @workorderno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@workorderno", "%" + txtWorkOrder2.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView4.DataSource = ds.Tables[0];
                dataGridView4.Columns[0].HeaderText = "No.";

            }
        }

        public void searchACC6()
        {

            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where InstRefNo like @instrefno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@instrefno", "%" + txtInsRefNo2.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView4.DataSource = ds.Tables[0];
                dataGridView4.Columns[0].HeaderText = "No.";

            }
        }

        public void searchACC7()
        {
            dataGridView5.DataSource = null;
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where WorkOrderNo like @workorderno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@workorderno", "%" + txtWorkOrder5.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView5.DataSource = ds.Tables[0];
                dataGridView5.Columns[0].HeaderText = "No.";

            }
        }


        private void btnLogOut_Click(object sender, EventArgs e)
        {
            this.Hide();
            LOGIN aaa = new LOGIN();
            aaa.ShowDialog();
            this.Close();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void txtWorkOrder_Enter(object sender, EventArgs e)
        {
            //if (txtWorkOrder.Text == "Input")
            //{
            //    txtWorkOrder.Text = "";
            //    txtWorkOrder.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void txtWorkOrder_Leave(object sender, EventArgs e)
        {
            //if (txtWorkOrder.Text == "")
            //{
            //    txtWorkOrder.Text = "Input";
            //    txtWorkOrder.ForeColor = Color.DarkGray;
            //}
        }

        private void txtInReNumber_Enter(object sender, EventArgs e)
        {
            //if (txtInReNumber.Text == "Input")
            //{
            //    txtInReNumber.Text = "";
            //    txtInReNumber.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }


        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ReleaseCapture();
                Common.SendMessage(Handle, Common.WM_NCLBUTTONDOWN, Common.HT_CAPTION, 0);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            DateTime now = DateTime.Now;
            Console.WriteLine(now.ToString("F"));

            DateTime dt = new DateTime(2019, 2, 22, 14, 0, 0);

            DateTime dt1 = dt.AddSeconds(55);
            DateTime dt2 = dt.AddMinutes(30);
            DateTime dt3 = dt.AddHours(72);
            DateTime dt4 = dt.AddDays(65);
            DateTime dt5 = dt.AddDays(-65);
            DateTime dt6 = dt.AddMonths(3);
            DateTime dt7 = dt.AddYears(4);

            Console.WriteLine(dt1.ToString("F"));
            Console.WriteLine(dt2.ToString("F"));
            Console.WriteLine(dt3.ToString("F"));
            Console.WriteLine(dt4.ToString("F"));
            Console.WriteLine(dt5.ToString("F"));
            Console.WriteLine(dt6.ToString("F"));
            Console.WriteLine(dt7.ToString("F"));

            lblDate.Text = DateTime.Now.ToString("F");
        }

        private void txtInReNumber_Leave_1(object sender, EventArgs e)
        {
            //if (txtInReNumber.Text == "")
            //{
            //    txtInReNumber.Text = "Input";
            //    txtInReNumber.ForeColor = Color.DarkGray;
            //}
        }

        private void btnCreateCalibrationReq_Click(object sender, EventArgs e)
        {
            CreateCalRequest aaa = new CreateCalRequest();

  

            aaa.ShowDialog();
        }

        private void btnEditReq_Click(object sender, EventArgs e)
        {

            EditCalRequest lookup = new EditCalRequest();

            EditCalRequest._instance.RequestSec();
            EditCalRequest._instance.InstrumentCat();
            EditCalRequest._instance.CalibrationFreq();
            EditCalRequest._instance.CalibrationPlan();
            lookup.txtWorkOrderNo.Text = this.dataGridView2.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
            lookup.dtReqDate.Text = this.dataGridView2.CurrentRow.Cells["DateReq"].Value.ToString();
            lookup.cmbRequestSection.Text = this.dataGridView2.CurrentRow.Cells["ReqSection"].Value.ToString();
            lookup.txtEmpNumber.Text = this.dataGridView2.CurrentRow.Cells["EmpNo"].Value.ToString();
            lookup.txtName.Text = this.dataGridView2.CurrentRow.Cells["ReqBy"].Value.ToString();
            lookup.txtEmail.Text = this.dataGridView2.CurrentRow.Cells["Email"].Value.ToString();
            lookup.txtCostCenter.Text = this.dataGridView2.CurrentRow.Cells["CostCenter"].Value.ToString();
            lookup.cmbInstrumentCat.Text = this.dataGridView2.CurrentRow.Cells["InstrumentCat"].Value.ToString();
            lookup.txtMinstrumentName.Text = this.dataGridView2.CurrentRow.Cells["MeasureInstName"].Value.ToString();
            lookup.txtInstrumentRefNo.Text = this.dataGridView2.CurrentRow.Cells["InstRefNo"].Value.ToString();
            lookup.txtManufucturer.Text = this.dataGridView2.CurrentRow.Cells["Manufacturer"].Value.ToString();
            lookup.txtModel.Text = this.dataGridView2.CurrentRow.Cells["Model"].Value.ToString();
            lookup.txtSerialNo.Text = this.dataGridView2.CurrentRow.Cells["SerialNo"].Value.ToString();
            lookup.txtInstrumentLoc.Text = this.dataGridView2.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
            lookup.cmbCalFreq.Text = this.dataGridView2.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
            lookup.dtDueDate.Text = this.dataGridView2.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
            lookup.cmbCalPlan.Text = this.dataGridView2.CurrentRow.Cells["calibrationPlan"].Value.ToString();
            lookup.txtReqRemarks.Text = this.dataGridView2.CurrentRow.Cells["Remarks"].Value.ToString();
            lookup.picInstrument.ImageLocation = this.dataGridView2.CurrentRow.Cells["Picture_Link"].Value.ToString();
            lookup.textBox1.Text = this.dataGridView2.CurrentRow.Cells["Picture_Link"].Value.ToString();

            lookup.ShowDialog();
        }
        private void btnSearch_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            if (txtInReNumber.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + txtWorkOrder.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }

            if (txtWorkOrder.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where InstRefNo = '" + txtInReNumber.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }

            else
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {

                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + txtWorkOrder.Text + "' AND InstRefNo = '" + txtInReNumber.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }

        }

        private void btnRefresh2_Click(object sender, EventArgs e)
        {
            DisplayData2();
            tabPage2.Text = "CALIBRATION REQUEST " + "(" + dataGridView2.RowCount.ToString() + ")";
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        private void tbWorkOrder2_Enter(object sender, EventArgs e)
        {
            //if (tbWorkOrder2.Text == "Input")
            //{
            //    tbWorkOrder2.Text = "";
            //    tbWorkOrder2.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void tbWorkOrder2_Leave(object sender, EventArgs e)
        {
            //if (tbWorkOrder2.Text == "")
            //{
            //    tbWorkOrder2.Text = "Input";
            //    tbWorkOrder2.ForeColor = Color.DarkGray;
            //}
        }

        private void tbInstRefNo2_Enter(object sender, EventArgs e)
        {
            //if (tbInstRefNo2.Text == "Input")
            //{
            //    tbInstRefNo2.Text = "";
            //    tbInstRefNo2.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void tbInstRefNo2_Leave(object sender, EventArgs e)
        {
            //if (tbInstRefNo2.Text == "")
            //{
            //    tbInstRefNo2.Text = "Input";
            //    tbInstRefNo2.ForeColor = Color.DarkGray;
            //}
        }

        public void SelectSection1()
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
                cmbSelectSection2.DisplayMember = "SectionList";
                cmbSelectSection2.DataSource = dt.Tables["SectionList"];
                btnSelSection.DisplayMember = "SectionList";
                btnSelSection.DataSource = dt.Tables["SectionList"];
                cmbSelSection4.DisplayMember = "SectionList";
                cmbSelSection4.DataSource = dt.Tables["SectionList"];
                cmbSelSection5.DisplayMember = "SectionList";
                cmbSelSection5.DataSource = dt.Tables["SectionList"];
                cmbSelSection6.DisplayMember = "SectionList";
                cmbSelSection6.DataSource = dt.Tables["SectionList"];
                conn.Close();

            }
        }

        public void SelectCalMonth1()
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
                cmbMonth.DisplayMember = "CalibrationPlan";
                cmbMonth.ValueMember = "CalibrationPlan";
                cmbMonth.DataSource = dt.Tables["CalibrationPlan"];
                cmbSelectMonth2.DisplayMember = "CalibrationPlan";
                cmbSelectMonth2.DataSource = dt.Tables["CalibrationPlan"];
                cmbSelMonth.DisplayMember = "CalibrationPlan";
                cmbSelMonth.DataSource = dt.Tables["CalibrationPlan"];
                cmbSelMonth4.DisplayMember = "CalibrationPlan";
                cmbSelMonth4.DataSource = dt.Tables["CalibrationPlan"];
                cmbSelMonth5.DisplayMember = "CalibrationPlan";
                cmbSelMonth5.DataSource = dt.Tables["CalibrationPlan"];
                cmbCalPlan6.DisplayMember = "CalibrationPlan";
                cmbCalPlan6.DataSource = dt.Tables["CalibrationPlan"];
                conn.Close();

            }
        }

        public void SelectStats()
        {
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                conn.Open();
                string sql = "select CalibrationStatus from tbl_CalibrationStatus";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataSet dt = new DataSet();
                sda.Fill(dt, "CalibrationStatus");
                cmbStatus.DisplayMember = "CalibrationStatus";
                cmbStatus.ValueMember = "CalibrationStatus";
                cmbStatus.DataSource = dt.Tables["CalibrationStatus"];
                cmbStatus2.DisplayMember = "CalibrationStatus";
                cmbStatus2.DataSource = dt.Tables["CalibrationStatus"];
                btnSelStat.DisplayMember = "CalibrationStatus";
                btnSelStat.DataSource = dt.Tables["CalibrationStatus"];
                cmbSelStat4.DisplayMember = "CalibrationStatus";
                cmbSelStat4.DataSource = dt.Tables["CalibrationStatus"];
                cmbSelStat5.DisplayMember = "CalibrationStatus";
                cmbSelStat5.DataSource = dt.Tables["CalibrationStatus"];
                conn.Close();
            }
        }

        //DATAGRIDVIEW 1
        private void btnFilter_Click(object sender, EventArgs e)
        {
            //HOME DATAGRIDVIEW 1
            //YEAR FILTERING

            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CertificateName,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSection.Text + "' AND OverallStatus = '" + cmbStatus.Text + "' AND CalibrationPlan = '" + cmbMonth.Text + "'  ";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];

            }



            if (cmbMonth.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSection.Text + "' AND OverallStatus = '" + cmbStatus.Text + "' ";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];

                }
            }

            if (cmbSection.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbMonth.Text + "' AND OverallStatus = '" + cmbStatus.Text + "' ";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }

            if (cmbStatus.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbMonth.Text + "' AND ReqSection = '" + cmbSection.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                }
            }

           if (cmbMonth.Text == "" && cmbSection.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = '" + cmbStatus.Text + "' ";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];

                }
            }
          if (cmbSection.Text == "" && cmbStatus.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbMonth.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];

                }
            }

            if (cmbMonth.Text == "" && cmbStatus.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, PDF_Link, EmpNo, ReqBy, Picture_Link, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateName,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSection.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];

                }
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int norun = 0;
            int count = dataGridView1.Rows.Count;
            for (int i = 0; i <= count - 1;)
            {
                string dgvValue = dataGridView1.Rows[i].Cells["OverallStatus"].Value.ToString();
                if (dgvValue == "FOR ACCEPTANCE" && dataGridView1.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                }
                else if (dgvValue == "DISAPPROVED" && dataGridView1.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                {
                    norun++;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                }

                else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView1.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                {
                    norun++;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }
                else if (dgvValue == "NOT YET RECEIVE" && dataGridView1.Rows[i].Cells["OverallStatus"].Value.ToString() == "NOT YET RECEIVE")
                {
                    norun++;
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                }

                i++;

            }
            dataGridView1.Columns["ID"].Width = 30;
            dataGridView1.Columns["WorkOrderNo"].Width = 120;
            dataGridView1.Columns["CostCenter"].Width = 55;
            dataGridView1.Columns["OverallStatus"].Width = 200;
        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            int norun = 0;
            int count = dataGridView2.Rows.Count;
            for (int i = 0; i <= count - 1;)
            {
                string dgvValue = dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString();
                if (dgvValue == "FOR ACCEPTANCE" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                {
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                }

                else if (dgvValue == "DISAPPROVED" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                {
                    norun++;
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                }

                else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                {
                    norun++;
                    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                }

                //else if (dgvValue == "FOR 1ST APPROVAL" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR 1ST APPROVAL")
                //{
                //    norun++;
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                //}
                //else if (dgvValue == "ON-GOING CALIBRATION" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "ON-GOING CALIBRATION")
                //{
                //    norun++;
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                //}
                //else if (dgvValue == "CERTIFICATE FOR APPROVAL" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "CERTIFICATE FOR APPROVAL")
                //{
                //    norun++;
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                //}
                //else if (dgvValue == "FOR FINAL APPROVAL" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR FINAL APPROVAL")
                //{
                //    norun++;
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                //}
                //else if (dgvValue == "FINISHED CALIBRATION" && dataGridView2.Rows[i].Cells["OverallStatus"].Value.ToString() == "FINISHED CALIBRATION")
                //{
                //    norun++;
                //    dataGridView2.Rows[i].DefaultCellStyle.BackColor = Color.LightSteelBlue;
                //}
                i++;
                //dataGridView1.Rows[1].DefaultCellStyle.BackColor = Color.Red;
            }
            dataGridView2.Columns["ID"].Width = 30;
            dataGridView2.Columns["WorkOrderNo"].Width = 120;
            dataGridView2.Columns["CostCenter"].Width = 55;
            dataGridView2.Columns["OverallStatus"].Width = 200;


        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView1.Columns[0].Frozen = true;
            dataGridView1.Columns[1].Frozen = true;
            dataGridView1.Columns[2].Frozen = true;
            dataGridView1.Columns[3].Frozen = true;
            dataGridView1.Columns[4].Frozen = true;
            dataGridView1.Columns[5].Frozen = true;
            dataGridView1.Columns[6].Frozen = true;
        }

        private void dataGridView2_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView2.Columns[0].Frozen = true;
            dataGridView2.Columns[1].Frozen = true;
            dataGridView2.Columns[2].Frozen = true;
            dataGridView2.Columns[3].Frozen = true;
            dataGridView2.Columns[4].Frozen = true;
            dataGridView2.Columns[5].Frozen = true;
            dataGridView2.Columns[6].Frozen = true;
        }

        private void dataGridView3_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {


                int norun = 0;
                int count = dataGridView3.Rows.Count;
                for (int i = 0; i <= count - 1;)
                {
                    string dgvValue = dataGridView3.Rows[i].Cells["OverallStatus"].Value.ToString();
                    if (dgvValue == "FOR ACCEPTANCE" && dataGridView3.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                    {
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                    }

                    else if (dgvValue == "DISAPPROVED" && dataGridView3.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                    {
                        norun++;
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                    }

                    else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView3.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                    {
                        norun++;
                        dataGridView3.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    }

                    i++;

                }

                dataGridView3.Columns["ID"].Width = 30;
                dataGridView3.Columns["WorkOrderNo"].Width = 120;
                dataGridView3.Columns["CostCenter"].Width = 55;
                dataGridView3.Columns["OverallStatus"].Width = 200;
                dataGridView3.Columns["Select"].Width = 55;
            }
            catch
            {

            }
        }


        private void btnSearch2_Click(object sender, EventArgs e)
        {
            searchAC3();

        }

        private void btnRefresh_Click(object sender, EventArgs e)
        {
            DisplayData1();
            tabPage1.Text = "HOME " + "(" + dataGridView1.RowCount.ToString() + ")";
        }

        private void btnDownloadFilter_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }

        private void btnDownloadFilter2_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView2.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView2.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView2.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView2.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }


        //DATAGRIDVIEW 3 ACCEPTANCE
        private void btnFilter3_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
           

        }

        private void btnFilter2_Click(object sender, EventArgs e)
        {
            //HOME DATAGRIDVIEW 2
            //YEAR FILTERING
            DataSet ds = new DataSet();
            if (cmbMonth.Text == "APR 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "MAY 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "JUN 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];

                }
            }
            else if (cmbMonth.Text == "JUL 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "AUG 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "SEP 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "OCT 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "NOV 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "DEC 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "JAN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "FEB 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbMonth.Text == "MAR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "APR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "MAY 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "JUN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "JUL 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "AUG 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }

            else if (cmbMonth.Text == "SEP 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "OCT 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "NOV 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "DEC 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "JAN 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }

            else if (cmbMonth.Text == "FEB 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "MAR 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }

            else if (cmbMonth.Text == "APR 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            //SECTION FILTERING DATAGRIDVIEW2
            else if (cmbSelectSection2.Text == "IQC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'IQC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelectSection2.Text == "EE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'EE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelectSection2.Text == "INK CARTRIDGE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'INK CARTRIDGE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelectSection2.Text == "TAPE CASSETTE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TAPE CASSETTE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelectSection2.Text == "MOLDING")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'MOLDING'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelectSection2.Text == "PRINTER(A3)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(A3)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelectSection2.Text == "PRINTER(MINI)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(MINI)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelectSection2.Text == "PT")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PT'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelectSection2.Text == "TC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }
            }

            else if (cmbStatus2.Text == "FOR ACCEPTANCE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }
            else if (cmbStatus2.Text == "FOR FINAL ACCEPTANCE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL ACCEPTANCE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }
            else if (cmbStatus2.Text == "FOR CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }
            else if (cmbStatus2.Text == "ON-GOING CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'ON-GOING CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }

            else if (cmbStatus2.Text == "FINISHED CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FINISHED CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }
            else if (cmbStatus2.Text == "FOR 1ST APPROVAL")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR 1ST APPROVAL'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }

            else if (cmbStatus2.Text == "FOR FINAL APPROVAL")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL APPROVAL'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }

            else if (cmbStatus2.Text == "COMPLETED : CERTIFICATE READY")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'COMPLETED : CERTIFICATE READY'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }

            else if (cmbStatus2.Text == "DISAPPROVED")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'DISAPPROVED'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView2.DataSource = ds.Tables[0];

                }

            }
        }


        private void btnSearch4_Click(object sender, EventArgs e)
        {
            searchACC5();

        }

        private void btnGraphs_Click(object sender, EventArgs e)
        {
            Graphs yyy = new Graphs();
            yyy.ShowDialog();
        }

        private void dataGridView4_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView4.Columns[0].Frozen = true;
            dataGridView4.Columns[1].Frozen = true;
            dataGridView4.Columns[2].Frozen = true;
            dataGridView4.Columns[3].Frozen = true;
            dataGridView4.Columns[4].Frozen = true;
            dataGridView4.Columns[5].Frozen = true;
            dataGridView4.Columns[6].Frozen = true;
        }



        private void dataGridView5_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView5.Columns[0].Frozen = true;
            dataGridView5.Columns[1].Frozen = true;
            dataGridView5.Columns[2].Frozen = true;
            dataGridView5.Columns[3].Frozen = true;
            dataGridView5.Columns[4].Frozen = true;
            dataGridView5.Columns[5].Frozen = true;
            dataGridView5.Columns[6].Frozen = true;
        }

        private void dataGridView6_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            dataGridView6.Columns[0].Frozen = true;
            dataGridView6.Columns[1].Frozen = true;
            dataGridView6.Columns[2].Frozen = true;
            dataGridView6.Columns[3].Frozen = true;
            dataGridView6.Columns[4].Frozen = true;
            dataGridView6.Columns[5].Frozen = true;
            dataGridView6.Columns[6].Frozen = true;
        }

        private void dataGridView4_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {


                int norun = 0;
                int count = dataGridView4.Rows.Count;
                for (int i = 0; i <= count - 1;)
                {
                    string dgvValue = dataGridView4.Rows[i].Cells["OverallStatus"].Value.ToString();
                    if (dgvValue == "FOR ACCEPTANCE" && dataGridView4.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                    {
                        dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                    }

                    else if (dgvValue == "DISAPPROVED" && dataGridView4.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                    {
                        norun++;
                        dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                    }

                    else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView4.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                    {
                        norun++;
                        dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else if (dgvValue == "NOT YET RECEIVE" && dataGridView4.Rows[i].Cells["OverallStatus"].Value.ToString() == "NOT YET RECEIVE")
                    {
                        norun++;
                        dataGridView4.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    }

                    //else if (dgvValue == "FOR 1ST APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR 1ST APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                    //}
                    //else if (dgvValue == "ON-GOING CALIBRATION" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "ON-GOING CALIBRATION")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                    //}
                    //else if (dgvValue == "CERTIFICATE FOR APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "CERTIFICATE FOR APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                    //}
                    //else if (dgvValue == "FOR FINAL APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR FINAL APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                    //}
                    //else if (dgvValue == "FINISHED CALIBRATION" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FINISHED CALIBRATION")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightSteelBlue;
                    //}
                    i++;
                    //dataGridView1.Rows[1].DefaultCellStyle.BackColor = Color.Red;
                }

                dataGridView4.Columns["ID"].Width = 30;
                dataGridView4.Columns["WorkOrderNo"].Width = 120;
                dataGridView4.Columns["CostCenter"].Width = 55;
                dataGridView4.Columns["OverallStatus"].Width = 200;
     
            }
            catch
            {
                //MessageBox.Show("Error");
            }

        }

        private void dataGridView5_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {


                int norun = 0;
                int count = dataGridView5.Rows.Count;
                for (int i = 0; i <= count - 1;)
                {
                    string dgvValue = dataGridView5.Rows[i].Cells["OverallStatus"].Value.ToString();
                    if (dgvValue == "FOR ACCEPTANCE" && dataGridView5.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                    {
                        dataGridView5.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                    }

                    else if (dgvValue == "DISAPPROVED" && dataGridView5.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                    {
                        norun++;
                        dataGridView5.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                    }

                    else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView5.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                    {
                        norun++;
                        dataGridView5.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    }
                    else if (dgvValue == "NOT YET RECEIVE" && dataGridView5.Rows[i].Cells["OverallStatus"].Value.ToString() == "NOT YET RECEIVE")
                    {
                        norun++;
                        dataGridView5.Rows[i].DefaultCellStyle.BackColor = Color.Red;
                    }

                    //else if (dgvValue == "FOR 1ST APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR 1ST APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.Orange;
                    //}
                    //else if (dgvValue == "ON-GOING CALIBRATION" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "ON-GOING CALIBRATION")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCyan;
                    //}
                    //else if (dgvValue == "CERTIFICATE FOR APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "CERTIFICATE FOR APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                    //}
                    //else if (dgvValue == "FOR FINAL APPROVAL" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR FINAL APPROVAL")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightCoral;
                    //}
                    //else if (dgvValue == "FINISHED CALIBRATION" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FINISHED CALIBRATION")
                    //{
                    //    norun++;
                    //    dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightSteelBlue;
                    //}
                    i++;
                    //dataGridView1.Rows[1].DefaultCellStyle.BackColor = Color.Red;
                }

                dataGridView5.Columns["ID"].Width = 30;
                dataGridView5.Columns["WorkOrderNo"].Width = 120;
                dataGridView5.Columns["CostCenter"].Width = 55;
                dataGridView5.Columns["OverallStatus"].Width = 200;
                dataGridView5.Columns["Select"].Width = 55;
            }
            catch
            {
                //MessageBox.Show("Error");
            }

        }



        private void dataGridView6_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {


                int norun = 0;
                int count = dataGridView6.Rows.Count;
                for (int i = 0; i <= count - 1;)
                {
                    string dgvValue = dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString();
                    if (dgvValue == "FOR ACCEPTANCE" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                    {
                        dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;
                    }

                    else if (dgvValue == "DISAPPROVED" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
                    {
                        norun++;
                        dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.DarkGray;
                    }

                    else if (dgvValue == "COMPLETED : CERTIFICATE READY" && dataGridView6.Rows[i].Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
                    {
                        norun++;
                        dataGridView6.Rows[i].DefaultCellStyle.BackColor = Color.LightGreen;
                    }

              
                    i++;

                }
                dataGridView6.Columns["ID"].Width = 30;
                dataGridView6.Columns["WorkOrderNo"].Width = 120;
                dataGridView6.Columns["CostCenter"].Width = 55;
                dataGridView6.Columns["OverallStatus"].Width = 200;
            }
            catch
            {
                //MessageBox.Show("Error");
            }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void btnRefresh3_Click(object sender, EventArgs e)
        {
            dataGridView3.Columns["Select"].ReadOnly = false;
            tabPage3.Text = "ACCEPTANCE APPROVAL " + "(" + dataGridView3.RowCount.ToString() + ")";
            DisplayData3();
        }

        private void btnApprove_Click(object sender, EventArgs e)
        {

            dataGridView3.Columns["Select"].ReadOnly = false;

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)
            {
                DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);
                if (dialog == DialogResult.OK)
                {
                    dataGridView3.Enabled = true;
                }
            }
            else
            {
                int i = 0;
                panel10.Visible = true;
                foreach (DataGridViewRow rows in selectedRows)
                {


                    i++;
                    if (i == 1)
                    {
                        lbl1.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl2.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email10.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 2)
                    {
                        lbl3.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl4.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email9.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 3)
                    {
                        lbl5.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl6.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email8.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 4)
                    {
                        lbl7.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl8.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email7.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 5)
                    {
                        lbl9.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl10.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email6.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 6)
                    {
                        label125.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label124.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email5.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 7)
                    {
                        label123.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label122.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email4.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 8)
                    {
                        label121.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label120.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email3.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 9)
                    {
                        label119.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label118.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email2.Text = rows.Cells["Email"].Value.ToString();
                    }
                    else if (i == 10)
                    {
                        label117.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label116.Text = rows.Cells["InstRefNo"].Value.ToString();
                        email1.Text = rows.Cells["Email"].Value.ToString();
                    }
                }

            }
        }


        private void btn1stAccept_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE'", con);
            adapt.Fill(dt);
            dataGridView3.DataSource = dt;
            dataGridView3.Columns["Remarks"].Visible = false;
            dataGridView3.Columns["EmpNo"].Visible = false;
            dataGridView3.Columns["Email"].Visible = false;
            dataGridView3.Columns["InstrumentLoc"].Visible = false;
            this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
            dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
            dataGridView3.Columns[0].HeaderText = "Select";
            dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
            dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
            dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
        }

        private void btnFinalAccept_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL ACCEPTANCE'", con);
            adapt.Fill(dt);
            dataGridView3.DataSource = dt;
            dataGridView3.Columns["Remarks"].Visible = false;
            dataGridView3.Columns["EmpNo"].Visible = false;
            dataGridView3.Columns["Email"].Visible = false;
            dataGridView3.Columns["InstrumentLoc"].Visible = false;
            this.dataGridView3.DefaultCellStyle.Font = new Font("Calibri", 8);
            dataGridView3.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
            dataGridView3.Columns[0].HeaderText = "Select";
            dataGridView3.Columns["CostCenter"].HeaderText = "CostCode";
            dataGridView3.Columns["CalibrationDueDate"].HeaderText = "Due Date";
            dataGridView3.Columns["CalibrationFreq"].HeaderText = "Frequency";
        }


        //ACCEPTANCE APPROVAL - ADMIN
        public string emailto;
        private string innerString;
        public void audit_email_AcceptanceApproval_ADMIN()
        {

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "select Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        emailto = rows.Cells["Email"].Value.ToString();
                        //emailto = dr["Email"].ToString();
                        EmailBody_Audit_AcceptanceApproval_ADMIN();
                    }
                }
            }
            EmailNotif_Audit_AcceptanceApproval_ADMIN();

        }
        private void EmailNotif_Audit_AcceptanceApproval_ADMIN()
        {

            MailMessage mail = new MailMessage("inhousecalibrationadmin@brother-biph.com.ph", emailto);
            mail.CC.Add(new MailAddress("christine.sotaso@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("alpie.hatulan@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("mary.mendoza @brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("roise.bordeos@brother-biph.com.ph"));
            mail.CC.Add(new MailAddress("kessie.blanca@brother-biph.com.ph"));
            mail.Bcc.Add(new MailAddress("zzpde178@brother.co.jp"));
            mail.Bcc.Add(new MailAddress("aaronpaul.anillo@brother-biph.com.ph"));
            SmtpClient client = new SmtpClient();
            client.Port = 25;
            client.UseDefaultCredentials = false;
            client.Host = "10.113.10.1";
            mail.Subject = "【IHC】REQUEST FOR FINAL APPROVAL";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
        public void EmailBody_Audit_AcceptanceApproval_ADMIN()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h1><font color=black>Good Day!</h1></b></font>");
            builder.Append("<br>");
            builder.Append("<h2><font color=Blue>Request for In-House Calibration has been ACCEPTED.</h2>");
            builder.Append("<h2><font color=Blue>Request ready for <span style='Background-color: Yellow'>APPROVAL.</span></h2>");
            builder.Append("<br>");

            string mailBody = "<table width='60%' style='border:Solid 1px gray;'>";

            int count = dataGridView3.RowCount;
            //{
            mailBody = "<h3><font color=black>Request Details: </font></h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Request Date</th><th>Instrument Reference Number</th> <th>Calibration Plan</th> <th>Section</th>";
            mailBody += "</tr>";

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["DateReq"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationPlan"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "</tr>";
            }

            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }
        string tempModel_to;
        //ACCEPTANCE APPROVAL - MANAGER
        public void audit_email_AcceptanceApproval_MANAGER()
        {

            string truncatetemp = "TRUNCATE TABLE [dbo].[tbl_TempEmail_Admin]";
            CRUD.CUD(truncatetemp);

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                string sql = "select DISTINCT Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                CRUD.RETRIEVESINGLE(sql);
                var emailto = CRUD.dt.Rows[0]["Email"].ToString();
                        EmailBody_Audit_AcceptanceApproval_MANAGER();
                string insertEmail = "INSERT INTO [dbo].[tbl_TempEmail_Admin] (EmailAdd) values ('" + emailto + "' )";
                CRUD.CUD(insertEmail);
            }
            string SelectRecepients = "SELECT DISTINCT EmailAdd FROM [dbo].[tbl_TempEmail_Admin]";
            CRUD.RETRIEVESINGLE(SelectRecepients);
            var emailtoTrue = CRUD.dt.AsEnumerable().Select(row => row["EmailAdd"].ToString());
            tempModel_to = String.Join(", ", emailtoTrue.ToArray());

            SENDEMAIL_URGENT(tempModel_to.ToString(), "【IHC】REQUEST APPROVED!");
        }
        string emailtos;
        void SENDEMAIL_URGENT(string WhoToEmail, string subject)
        {
            emailtos = WhoToEmail;
            if (emailtos == "")
            {
               
            }
            else
            {
                try
                {
                    //Email structure.
                    MailMessage mail = new MailMessage();
                    mail.Headers.Add("Importance", "High");
                    mail.From = new MailAddress("inhousecalibration@brother-biph.com.ph");
                    mail.To.Add(FormatMultipleEmailAddresses(emailtos));
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
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = "10.113.10.1";
                    mail.Subject = subject;



                    mail.Body = innerString;
                    mail.IsBodyHtml = true;
                    client.Send(mail);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                }
            }
        }

        private string FormatMultipleEmailAddresses(string emailAddresses)
        {
            var delimiters = new[] { ',', ';' };
            var addresses = emailAddresses.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(",", addresses);
        }

        public void EmailBody_Audit_AcceptanceApproval_MANAGER()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h1><font color=black>Good Day!</h1></b><br></font>");
            builder.Append("<br>");
            builder.Append("<h2><font color=Blue>Request for In-House Calibration has been <span style='Background-color: Yellow'>APPROVED.</span></h2>");
            builder.Append("<h2><span style='Background-color: Yellow'><font color=Blue>You may now bring below instruments at PQC 3D Room.</span></h2>");

            string mailBody = "<table width='60%' style='border:Solid 1px gray;'>";

            int count = dataGridView3.RowCount;

            mailBody = "<h3><font color=black>Request Details:</font></h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Instrument Reference Number</th> <th>Calibration Plan</th> <th>Section</th>";
            mailBody += "</tr>";

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationPlan"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "</tr>";
            }

            mailBody += "</table>";
            builder.Append("" + mailBody);

            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }

        //ACCEPTANCE TAB - DISAPPROVED
        public void audit_email_AcceptanceDisapproved()
        {
            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "select Distinct Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "' AND ReqSection = '" + rows.Cells["ReqSection"].Value.ToString() + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        emailto = dr["Email"].ToString();
                        EmailBody_Audit_AcceptanceDisapproved();

                    }
                }

            }
            EmailNotif_Audit_AcceptanceDisapproved();

        }
        private void EmailNotif_Audit_AcceptanceDisapproved()
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
            mail.Subject = "【IHC】REQUEST DECLINED";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
        public void EmailBody_Audit_AcceptanceDisapproved()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h2><font color=black>Good Day!</h2></b></font>");
            builder.Append("<h2><font color=Blue>Request for In-House calibration has been <span style='Background-color: Yellow'><font color=red> DECLINED.</span> </font> </h2>");
            string mailBody = "";
            int count = dataGridView3.RowCount;
            //{
            mailBody = "<h3><font color=black>Request Details: </font> </h3><table width='50%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Instrument Reference Number</th> <th>Calibration Plan</th> <th>Section</th>";
            mailBody += "</tr>";

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationPlan"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "</tr>";
                builder.Append("<br>");
            }
            builder.Append("<br>");
            mailBody += "<tr >";
            mailBody += "<th style = 'background-color: #D6EEEE'  colspan ='3'> REMARKS </th>";
            mailBody += "</tr>";
            mailBody += "<tr style>";
            builder.Append("<br>");
            mailBody += "<td style='color:blue;'><font color=red> DECLINED REMARKS: <br> " + txtRemarks.Text + "</font></td>";
            mailBody += "</tr>";




            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br><font color=black>Thank you!</font></h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }


        //CERTIFICATE APPROVED / ADMIN

        public void audit_email_CertificateApproved_ADMIN()
        {
            string truncatetemp_Admin = "TRUNCATE TABLE [dbo].[tbl_TempEmail_Admin]";
            CRUD.CUD(truncatetemp_Admin);

            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                string sql = "select DISTINCT Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                CRUD.RETRIEVESINGLE(sql);
                var emailto_Admin = CRUD.dt.Rows[0]["Email"].ToString();
                EmailBody_Audit_CertificateApproved();
                string insertEmail_Admin = "INSERT INTO [dbo].[tbl_TempEmail_Admin](EmailAdd) values ('" + emailto_Admin + "' )";
                CRUD.CUD(insertEmail_Admin);
            }

            string SelectRecepients = "SELECT DISTINCT EmailAdd FROM [dbo].[tbl_TempEmail_Admin]";
            CRUD.RETRIEVESINGLE(SelectRecepients);
            var emailtoTrue_Admin = CRUD.dt.AsEnumerable().Select(row => row["EmailAdd"].ToString());
            tempModel_to = String.Join(", ", emailtoTrue_Admin.ToArray());


            SENDEMAIL_URGENT_CERTIFICATEADMIN(tempModel_to.ToString(), "【IHC_CERTIFICATE】CERTIFICATE FOR APPROVAL");
        }
        string emailtos_Admin;
        void SENDEMAIL_URGENT_CERTIFICATEADMIN(string WhoToEmail_Certificate_Admin, string subject_Certificate_Admin)
        {
            emailtos_Admin = WhoToEmail_Certificate_Admin;
            if (emailtos_Admin == "")
            {
             
            }
            else
            {
                try
                {
                    //Email structure.
                    MailMessage mail = new MailMessage();
                    mail.Headers.Add("Importance", "High");
                    mail.From = new MailAddress("inhousecalibration@brother-biph.com.ph");
                    mail.To.Add(FormatMultipleEmailAddresses_CertificateADMIN(emailtos_Admin));
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
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = "10.113.10.1";
                    mail.Subject = subject_Certificate_Admin;



                    mail.Body = innerString;
                    mail.IsBodyHtml = true;
                    client.Send(mail);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                }



            }
        }

        private string FormatMultipleEmailAddresses_CertificateADMIN(string emailAddresses_Certificate_Admin)
        {
            var delimiters_Admin = new[] { ',', ';' };



            var addresses_Admin = emailAddresses_Certificate_Admin.Split(delimiters_Admin, StringSplitOptions.RemoveEmptyEntries);



            return string.Join(",", addresses_Admin);
        }
        public void EmailBody_Audit_CertificateApproved()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h1><font color=black>Good Day!</h1></b></font>");
            builder.Append("<br>");
            builder.Append("<h2><font color=Blue>CERTIFICATE for In-House Calibration has been signed by checker.</h2>");
            builder.Append("<h2><font color=Blue>CERTIFICATE ready for <span style='Background-color: Yellow'>FINAL APPROVAL.</span></h2>");
            builder.Append("<br>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";

            int count = dataGridView3.RowCount;
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font></h3><table width='50%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th> <th>Calibrator In-Charge</th>";
            mailBody += "</tr>";



            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationIncharge"].Value.ToString() + "</td>";
                mailBody += "</tr>";
            }

            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }
        //CERTIFICATE APPROVAL - MANAGER
        public void audit_email_CertificateApproved_MANAGER()
        {
            string truncatetemp_Manager = "TRUNCATE TABLE [dbo].[tbl_TempEmail_Manager]";
            CRUD.CUD(truncatetemp_Manager);

            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                string sql = "select DISTINCT Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                CRUD.RETRIEVESINGLE(sql);
                var emailto_Manager = CRUD.dt.Rows[0]["Email"].ToString();
                EmailBody_Audit_CertificateApproved_MANAGER();
                string insertEmail_Manager = "INSERT INTO [dbo].[tbl_TempEmail_Manager](EmailAdd) values ('" + emailto_Manager + "' )";
                CRUD.CUD(insertEmail_Manager);
            }

            string SelectRecepients_Manager = "SELECT DISTINCT EmailAdd FROM [dbo].[tbl_TempEmail_Manager]";
            CRUD.RETRIEVESINGLE(SelectRecepients_Manager);
            var emailtoTrue_Manager = CRUD.dt.AsEnumerable().Select(row => row["EmailAdd"].ToString());
            tempModel_to = String.Join(", ", emailtoTrue_Manager.ToArray());

            SENDEMAIL_URGENT_CERTIFICATEMANAGER(tempModel_to.ToString(), "【IHC_CERTIFICATE】CERTIFICATE READY!");
        }
       string emailtos_Manager;
        void SENDEMAIL_URGENT_CERTIFICATEMANAGER(string WhoToEmail_Certificate_Manager, string subject_Certificate_manager)
        {

            emailtos_Manager = WhoToEmail_Certificate_Manager;



            if (emailtos_Manager == "")
            {
                
            }
            else
            {
                try
                {
                    //Email structure.
                    MailMessage mail = new MailMessage();
                    mail.Headers.Add("Importance", "High");
                    mail.From = new MailAddress("inhousecalibration@brother-biph.com.ph");
                    mail.To.Add(FormatMultipleEmailAddresses_CertificateMANAGER(emailtos_Manager));
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
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.UseDefaultCredentials = false;
                    client.Host = "10.113.10.1";
                    mail.Subject = subject_Certificate_manager;

                    mail.Body = innerString;
                    mail.IsBodyHtml = true;
                    client.Send(mail);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error " + ex.Message);
                }



            }
        }

        private string FormatMultipleEmailAddresses_CertificateMANAGER(string emailAddresses_Certificate_Manager)
        {
            var delimiters = new[] { ',', ';' };



            var addresses_Manager = emailAddresses_Certificate_Manager.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);



            return string.Join(",", addresses_Manager);
        }
        public void EmailBody_Audit_CertificateApproved_MANAGER()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h1><font color=black>Good Day!</h1></b></font>");
            builder.Append("<br>");
            builder.Append("<h2><font color=Blue>CERTIFICATE for In-House Calibration is now <span style='Background-color: Yellow'> READY. </span></h2>");
            builder.Append("<h2><font color=Blue>You may download certificate through <a href='\\\\apbiphsh07\\D0_ShareBrotherGroup\\19_BPS\\Installer\\InHouseCalibration\\setup.exe'> In-House Calibration System </a></h2>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";

            int count = dataGridView3.RowCount;
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font> </h3><table width='60%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th> <th>Calibrator In-Charge</th> <th>Calibration Plan</th>";
            mailBody += "</tr>";



            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationIncharge"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationPlan"].Value.ToString() + "</td>";
                mailBody += "</tr>";
            }

            mailBody += "</table>";
            builder.Append("" + mailBody);
            // }


            builder.Append("<br>");
            builder.Append("<h2><br>Thank you!</h2>").AppendLine();
            builder.Append("<br>");
            builder.Append("<h2><font color=Red>Note: This is a System Generated Mail. Do not reply!</font></h2>").AppendLine();
            innerString = builder.ToString();
        }

        //CERTIFICATE DISAPPROVED
        public void audit_email_CertificateDisapproved()
        {
            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "select Distinct Email from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(dt);
                    foreach (DataRow dr in dt.Rows)
                    {
                        emailto = dr["Email"].ToString();
                        EmailBody_Audit_CertificateDisapproved();

                    }
                }

            }
            EmailNotif_Audit_CertificateDisapproved();

        }
        private void EmailNotif_Audit_CertificateDisapproved()
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
            mail.Subject = "【IHC】CERTIFICATE DECLINED!";
            mail.Body = innerString;
            mail.IsBodyHtml = true;
            client.Send(mail);
        }
        public void EmailBody_Audit_CertificateDisapproved()
        {
            DateTime currentTime = DateTime.Now;
            string timeString = currentTime.ToString("hh:mm tt", CultureInfo.InvariantCulture);
            DateTime aa = DateTime.Now;
            string date1 = aa.ToString("MM/dd/yyyy");
            string cnt = Convert.ToString(date1);
            StringBuilder builder = new StringBuilder();
            builder.AppendLine();

            builder.Append("<b><h1><font color=black>Good Day!</h1></b></font>");
            builder.Append("<br>");
            builder.Append("<h2><font color=Blue>CERTIFICATE for In-House calibration has been </font> <span style='Background-color: Yellow'><font color=red> DISAPPROVED </span></font>by checker</h2>");

            string mailBody = "<table width='50%' style='border:Solid 1px gray;'>";

            int count = dataGridView3.RowCount;
            //{
            mailBody = "<h3><font color=black>Certificate Details: </font></h3><table width='50%' style='border:Solid 1px Black;'>";
            mailBody += "<tr style = 'background-color: #D6EEEE'> <th>Instrument Reference Number</th> <th>BIPH Certificate Number</th> <th>Section</th> <th>Calibrator In-Charge</th>";
            mailBody += "</tr>";



            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            foreach (DataGridViewRow rows in selectedRows)
            {

                mailBody += "<tr align='Center'>";
                mailBody += "<td style='color:black;'>" + rows.Cells["InstRefNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationNo"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["ReqSection"].Value.ToString() + "</td>";
                mailBody += "<td style='color:black;'>" + rows.Cells["CalibrationIncharge"].Value.ToString() + "</td>";
                mailBody += "</tr>";
                builder.Append("<br>");
            }

            builder.Append("<br>");
            mailBody += "<tr >";
            mailBody += "<th style = 'background-color: #D6EEEE'  colspan ='3'> REMARKS </th>";
            mailBody += "</tr>";
            mailBody += "<tr style>";
            builder.Append("<br>");
            mailBody += "<td style='color:blue;'><font color=red> DECLINED REMARKS: <br> " + txtRemarks.Text + "</font></td>";
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
        private void btnDisapprove_Click(object sender, EventArgs e)
        {


            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)
            {
                DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);
            }
            else
            {
                panel11.Visible = true;
                int i = 0;
                foreach (DataGridViewRow rows in selectedRows)
                {
                    i++;
                    if (i == 1)
                    {
                        lbl11.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl12.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 2)
                    {
                        lbl13.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl14.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 3)
                    {
                        lbl15.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl16.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 4)
                    {
                        lbl17.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl18.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 5)
                    {
                        lbl19.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        lbl20.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 6)
                    {
                        label93.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label92.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 7)
                    {
                        label91.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label90.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 8)
                    {
                        label89.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label88.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 9)
                    {
                        label87.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label86.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 10)
                    {
                        label85.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label84.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                }
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    row.Cells["Select"].Value = true;
                }
            }
            else if (checkBox1.Checked == false)
            {
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    row.Cells["Select"].Value = false;
                }
            }
        }

        private void btnYes_Click(object sender, EventArgs e)
        {

            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)

            {

                //dataGridView3.Enabled = false;
                //DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);

                //if (dialog == DialogResult.OK)
                //{
                //    dataGridView3.Enabled = true;
                //}             
            }

            SqlCommand cmd;

            con.Open();

            foreach (DataGridViewRow rows in selectedRows)
            {

                cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set OverallStatus = 'NOT YET RECEIVE' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                cmd.ExecuteNonQuery();




                panel10.Visible = false;
                lbl1.Text = "";
                lbl2.Text = "";
                lbl3.Text = "";
                lbl4.Text = "";
                lbl5.Text = "";
                lbl6.Text = "";
                lbl7.Text = "";
                lbl8.Text = "";
                lbl9.Text = "";
                lbl10.Text = "";
                label116.Text = "";
                label117.Text = "";
                label118.Text = "";
                label119.Text = "";
                label120.Text = "";
                label121.Text = "";
                label122.Text = "";
                label123.Text = "";
                label124.Text = "";
                label125.Text = "";

            }

            audit_email_AcceptanceApproval_MANAGER();

            MessageBox.Show("Calibration Confirm Successfully");
            Home._instance.refreshdata3();
            Home._instance.refreshdata4();
            con.Close();

        }
        //else if (LOGIN.loginAutho == "Manager")
        //{
        //    SqlCommand cmd;

        //    con.Open();

        //    foreach (DataGridViewRow rows in selectedRows)
        //    {
        //        cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set OverallStatus = 'NOT YET RECEIVE' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

        //        cmd.ExecuteNonQuery();


        //        panel10.Visible = false;


        //        lbl1.Text = "";
        //        lbl2.Text = "";
        //        lbl3.Text = "";
        //        lbl4.Text = "";
        //        lbl5.Text = "";
        //        lbl6.Text = "";
        //        lbl7.Text = "";
        //        lbl8.Text = "";
        //        lbl9.Text = "";
        //        lbl10.Text = "";
        //        label116.Text = "";
        //        label117.Text = "";
        //        label118.Text = "";
        //        label119.Text = "";
        //        label120.Text = "";
        //        label121.Text = "";
        //        label122.Text = "";
        //        label123.Text = "";
        //        label124.Text = "";
        //        label125.Text = "";

        //    }

        //    audit_email_AcceptanceApproval_MANAGER();
        //    MessageBox.Show("Calibration Confirm Successfully");
        //    con.Close();
        //    Home._instance.refreshdata3();
        //    foreach (DataGridViewRow row in dataGridView3.Rows)

        //    {

        //        row.Cells["Select"].Value = false;

        //    }



        private void button7_Click(object sender, EventArgs e)
        {
            panel10.Visible = false;
            lbl1.Text = "";
            lbl2.Text = "";
            lbl3.Text = "";
            lbl4.Text = "";
            lbl5.Text = "";
            lbl6.Text = "";
            lbl7.Text = "";
            lbl8.Text = "";
            lbl9.Text = "";
            lbl10.Text = "";
            label116.Text = "";
            label117.Text = "";
            label118.Text = "";
            label119.Text = "";
            label120.Text = "";
            label121.Text = "";
            label122.Text = "";
            label123.Text = "";
            label124.Text = "";
            label125.Text = "";

        }

        private void btnYes2_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)

            {

                dataGridView3.Enabled = false;

            }
            else

            {

                SqlCommand cmd;

                con.Open();

                foreach (DataGridViewRow rows in selectedRows)
                {
                    cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set OverallStatus = 'DISAPPROVED', Status = 'VOID', Remarks2 = '" + txtRemarks.Text + "' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                    cmd.ExecuteNonQuery();



                    panel11.Visible = false;

                    lbl11.Text = "";
                    lbl12.Text = "";
                    lbl13.Text = "";
                    lbl14.Text = "";
                    lbl15.Text = "";
                    lbl16.Text = "";
                    lbl17.Text = "";
                    lbl18.Text = "";
                    lbl19.Text = "";
                    lbl20.Text = "";
                    label84.Text = "";
                    label85.Text = "";
                    label86.Text = "";
                    label87.Text = "";
                    label88.Text = "";
                    label89.Text = "";
                    label90.Text = "";
                    label91.Text = "";
                    label92.Text = "";
                    label93.Text = "";
                }

                audit_email_AcceptanceDisapproved();
                MessageBox.Show("Calibration Declined!");
                Home._instance.refreshdata3();
                con.Close();
            }


            foreach (DataGridViewRow row in dataGridView3.Rows)

            {
                row.Cells["Select"].Value = false;

            }
        }


        private void btnCancel2_Click(object sender, EventArgs e)
        {
            panel11.Visible = false;
            dataGridView5.Columns["Select"].ReadOnly = false;

        }

        private void button10_Click(object sender, EventArgs e)
        {
            panel11.Visible = false;
            lbl11.Text = "";
            lbl12.Text = "";
            lbl13.Text = "";
            lbl14.Text = "";
            lbl15.Text = "";
            lbl16.Text = "";
            lbl17.Text = "";
            lbl18.Text = "";
            lbl19.Text = "";
            lbl20.Text = "";
            label84.Text = "";
            label85.Text = "";
            label86.Text = "";
            label87.Text = "";
            label88.Text = "";
            label89.Text = "";
            label90.Text = "";
            label91.Text = "";
            label92.Text = "";
            label93.Text = "";
        }

        private void LinkAccount_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {

            AdminPage eron = new AdminPage();
            eron.ShowDialog();

        }

        private void btnRecieving_Click(object sender, EventArgs e)
        {
            
            InstrumentReceiving lookup = new InstrumentReceiving();

            InstrumentReceiving._instance.RequestSec();
            InstrumentReceiving._instance.InstrumentCat();
            InstrumentReceiving._instance.CalibrationFreq();
            InstrumentReceiving._instance.CalibrationPlan();
            lookup.txtWorkOrderNo.Text = this.dataGridView4.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
            lookup.dtReqDate.Text = this.dataGridView4.CurrentRow.Cells["DateReq"].Value.ToString();
            lookup.cmbRequestSection.Text = this.dataGridView4.CurrentRow.Cells["ReqSection"].Value.ToString();
            lookup.txtEmpNumber.Text = this.dataGridView4.CurrentRow.Cells["EmpNo"].Value.ToString();
            lookup.txtName.Text = this.dataGridView4.CurrentRow.Cells["ReqBy"].Value.ToString();
            lookup.txtEmail.Text = this.dataGridView4.CurrentRow.Cells["Email"].Value.ToString();
            lookup.txtCostCenter.Text = this.dataGridView4.CurrentRow.Cells["CostCenter"].Value.ToString();
            lookup.cmbInstrumentCat.Text = this.dataGridView4.CurrentRow.Cells["InstrumentCat"].Value.ToString();
            lookup.txtMinstrumentName.Text = this.dataGridView4.CurrentRow.Cells["MeasureInstName"].Value.ToString();
            lookup.txtInstrumentRefNo.Text = this.dataGridView4.CurrentRow.Cells["InstRefNo"].Value.ToString();
            lookup.txtManufucturer.Text = this.dataGridView4.CurrentRow.Cells["Manufacturer"].Value.ToString();
            lookup.txtModel.Text = this.dataGridView4.CurrentRow.Cells["Model"].Value.ToString();
            lookup.txtSerialNo.Text = this.dataGridView4.CurrentRow.Cells["SerialNo"].Value.ToString();
            lookup.txtInstrumentLoc.Text = this.dataGridView4.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
            lookup.cmbCalFreq.Text = this.dataGridView4.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
            lookup.dtDueDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
            lookup.cmbCalPlan.Text = this.dataGridView4.CurrentRow.Cells["calibrationPlan"].Value.ToString();
            lookup.txtReqRemarks.Text = this.dataGridView4.CurrentRow.Cells["Remarks"].Value.ToString();
            lookup.txtEmpNo.Text = this.dataGridView4.CurrentRow.Cells["EmpNo"].Value.ToString();
            lookup.tbEmpNo.Text = this.dataGridView4.CurrentRow.Cells["EmpNo"].Value.ToString();
            lookup.lblFileName.Text = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
            lookup.picInstrument.ImageLocation = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
            lookup.dateTimePicker1.Text = this.dataGridView4.CurrentRow.Cells["DateReceived"].Value.ToString();
            lookup.tbName.Text = this.dataGridView4.CurrentRow.Cells["EndorsedBy"].Value.ToString();
            lookup.txtName1.Text = this.dataGridView4.CurrentRow.Cells["ReceivedBy"].Value.ToString();
            lookup.lblStatstat.Text = this.dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString();

            lookup.ShowDialog();
        }

        private void btnUpdateDetails_Click(object sender, EventArgs e)
        {
            if (LOGIN.loginAutho == "Admin")
            {
                UpdateCalibration lookdown = new UpdateCalibration();
                UpdateCalibration._instance.InstrumentCat();
                UpdateCalibration._instance.CalibrationFreq();
                UpdateCalibration._instance.CalibrationPlan();
                UpdateCalibration._instance.CalibIncharge();
                // UpdateCalibration._instance.SelectForm();
                lookdown.txtWorkOrderNo.Text = this.dataGridView4.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                lookdown.cmbInstrumentCat.Text = this.dataGridView4.CurrentRow.Cells["InstrumentCat"].Value.ToString();
                lookdown.txtMinstrumentName.Text = this.dataGridView4.CurrentRow.Cells["MeasureInstName"].Value.ToString();
                lookdown.txtInstrumentRefNo.Text = this.dataGridView4.CurrentRow.Cells["InstRefNo"].Value.ToString();
                lookdown.txtManufucturer.Text = this.dataGridView4.CurrentRow.Cells["Manufacturer"].Value.ToString();
                lookdown.txtModel.Text = this.dataGridView4.CurrentRow.Cells["Model"].Value.ToString();
                lookdown.txtSerialNo.Text = this.dataGridView4.CurrentRow.Cells["SerialNo"].Value.ToString();
                lookdown.txtInstrumentLoc.Text = this.dataGridView4.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
                lookdown.cmbCalFreq.Text = this.dataGridView4.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
                lookdown.dtDueDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
                lookdown.cmbCalPlan.Text = this.dataGridView4.CurrentRow.Cells["calibrationPlan"].Value.ToString();
                lookdown.txtReqRemarks.Text = this.dataGridView4.CurrentRow.Cells["Remarks"].Value.ToString();
                lookdown.cmbCalInCharge.Text = this.dataGridView4.CurrentRow.Cells["CalibrationIncharge"].Value.ToString();
                lookdown.dtCalStartDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationStartDate"].Value.ToString();
                lookdown.dtCalEndDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationEndDate"].Value.ToString();
                lookdown.cmbStartTime.Text = this.dataGridView4.CurrentRow.Cells["StartTime"].Value.ToString();
                lookdown.cmbEndTime.Text = this.dataGridView4.CurrentRow.Cells["EndTime"].Value.ToString();
                lookdown.txtAutoCompute.Text = this.dataGridView4.CurrentRow.Cells["CalibrationTimehrs"].Value.ToString();
                lookdown.txtReqSection.Text = this.dataGridView4.CurrentRow.Cells["ReqSection"].Value.ToString();
                lookdown.txtEmail.Text = this.dataGridView4.CurrentRow.Cells["Email"].Value.ToString();
                lookdown.picInstrument.ImageLocation = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown.txtLinkPic.Text = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown.lblStat.Text = this.dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown.SelectForm.Text = this.dataGridView4.CurrentRow.Cells["CertificateName"].Value.ToString();
                lookdown.comboBox1.Text = this.dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString();
                //  lookdown.SelectForm.Enabled = false;
                lookdown.ShowDialog();
            }

            else if (LOGIN.loginAutho == "Supervisor")
            {
                UpdateCalibration lookdown = new UpdateCalibration();
                UpdateCalibration._instance.InstrumentCat();
                UpdateCalibration._instance.CalibrationFreq();
                UpdateCalibration._instance.CalibrationPlan();
                UpdateCalibration._instance.CalibIncharge();
                // UpdateCalibration._instance.SelectForm();
                lookdown.txtWorkOrderNo.Text = this.dataGridView4.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                lookdown.cmbInstrumentCat.Text = this.dataGridView4.CurrentRow.Cells["InstrumentCat"].Value.ToString();
                lookdown.txtMinstrumentName.Text = this.dataGridView4.CurrentRow.Cells["MeasureInstName"].Value.ToString();
                lookdown.txtInstrumentRefNo.Text = this.dataGridView4.CurrentRow.Cells["InstRefNo"].Value.ToString();
                lookdown.txtManufucturer.Text = this.dataGridView4.CurrentRow.Cells["Manufacturer"].Value.ToString();
                lookdown.txtModel.Text = this.dataGridView4.CurrentRow.Cells["Model"].Value.ToString();
                lookdown.txtSerialNo.Text = this.dataGridView4.CurrentRow.Cells["SerialNo"].Value.ToString();
                lookdown.txtInstrumentLoc.Text = this.dataGridView4.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
                lookdown.cmbCalFreq.Text = this.dataGridView4.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
                lookdown.dtDueDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
                lookdown.cmbCalPlan.Text = this.dataGridView4.CurrentRow.Cells["calibrationPlan"].Value.ToString();
                lookdown.txtReqRemarks.Text = this.dataGridView4.CurrentRow.Cells["Remarks"].Value.ToString();
                lookdown.cmbCalInCharge.Text = this.dataGridView4.CurrentRow.Cells["CalibrationIncharge"].Value.ToString();
                lookdown.dtCalStartDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationStartDate"].Value.ToString();
                lookdown.dtCalEndDate.Text = this.dataGridView4.CurrentRow.Cells["CalibrationEndDate"].Value.ToString();
                lookdown.cmbStartTime.Text = this.dataGridView4.CurrentRow.Cells["StartTime"].Value.ToString();
                lookdown.cmbEndTime.Text = this.dataGridView4.CurrentRow.Cells["EndTime"].Value.ToString();
                lookdown.txtAutoCompute.Text = this.dataGridView4.CurrentRow.Cells["CalibrationTimehrs"].Value.ToString();
                lookdown.txtReqSection.Text = this.dataGridView4.CurrentRow.Cells["ReqSection"].Value.ToString();
                lookdown.txtEmail.Text = this.dataGridView4.CurrentRow.Cells["Email"].Value.ToString();
                lookdown.picInstrument.ImageLocation = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown.txtLinkPic.Text = this.dataGridView4.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown.lblStat.Text = this.dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown.SelectForm.Text = this.dataGridView4.CurrentRow.Cells["CertificateName"].Value.ToString();
                lookdown.comboBox1.Text = this.dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString();
                //  lookdown.SelectForm.Enabled = false;
                lookdown.ShowDialog();
            }
           
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            panel10.Visible = false;
            lbl1.Text = "";
            lbl2.Text = "";
            lbl3.Text = "";
            lbl4.Text = "";
            lbl5.Text = "";
            lbl6.Text = "";
            lbl7.Text = "";
            lbl8.Text = "";
            lbl9.Text = "";
            lbl10.Text = "";
            label116.Text = "";
            label117.Text = "";
            label118.Text = "";
            label119.Text = "";
            label120.Text = "";
            label121.Text = "";
            label122.Text = "";
            label123.Text = "";
            label124.Text = "";
            label125.Text = "";

        }

        private void btnApprove1_Click(object sender, EventArgs e)
        {

            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)
            {
                DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);
            }
            else
            {
                panel13.Visible = true;
                int i = 0;
                foreach (DataGridViewRow rows in selectedRows)
                {
                    i++;
                    if (i == 1)
                    {
                        label111.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label110.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 2)
                    {
                        label109.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label108.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 3)
                    {
                        label107.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label106.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 4)
                    {
                        label105.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label104.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 5)
                    {
                        label103.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label102.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 6)
                    {
                        label155.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label154.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 7)
                    {
                        label153.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label152.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 8)
                    {
                        label151.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label150.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 9)
                    {
                        label149.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label148.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 10)
                    {
                        label147.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label146.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }

                }
            }
        }

        private void btnDisapprove1_Click(object sender, EventArgs e)
        {

            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)
            {
                DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);
            }
            else
            {
                panel12.Visible = true;
                int i = 0;
                foreach (DataGridViewRow rows in selectedRows)
                {
                    i++;
                    if (i == 1)
                    {
                        label76.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label75.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 2)
                    {
                        label74.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label73.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 3)
                    {
                        label72.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label71.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 4)
                    {
                        label67.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label66.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 5)
                    {
                        label65.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label64.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 6)
                    {
                        label140.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label139.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 7)
                    {
                        label138.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label137.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 8)
                    {
                        label136.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label135.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 9)
                    {
                        label134.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label133.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }
                    else if (i == 10)
                    {
                        label132.Text = rows.Cells["WorkOrderNo"].Value.ToString();
                        label131.Text = rows.Cells["InstRefNo"].Value.ToString();
                    }

                }
            }
        }

        private void btnYes6_Click(object sender, EventArgs e)
        {
            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)

            {

                dataGridView5.Enabled = false;
            }
            else

            {

                SqlCommand cmd;

                con.Open();

                foreach (DataGridViewRow rows in selectedRows)
                {
                    cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set OverallStatus = 'DISAPPROVED', Status = 'Void', Remarks3 = '" + tbRemarks5.Text + "' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                    cmd.ExecuteNonQuery();


                    panel12.Visible = false;

                    label76.Text = "";
                    label75.Text = "";
                    label74.Text = "";
                    label73.Text = "";
                    label72.Text = "";
                    label71.Text = "";
                    label67.Text = "";
                    label66.Text = "";
                    label65.Text = "";
                    label64.Text = "";
                    label140.Text = "";
                    label139.Text = "";
                    label138.Text = "";
                    label137.Text = "";
                    label136.Text = "";
                    label135.Text = "";
                    label134.Text = "";
                    label133.Text = "";
                    label132.Text = "";
                    label131.Text = "";
                }
                con.Close();
                audit_email_CertificateDisapproved();
                MessageBox.Show("Certificate has been DISAPPROVED");
                Home._instance.refreshdata5();
            }


            foreach (DataGridViewRow row in dataGridView5.Rows)

            {
                row.Cells["Select"].Value = false;

            }
        }

        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;

        private void dataGridView5_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

            if (dataGridView5.CurrentRow.Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
            {
                CalibrationCertificate aaron = new CalibrationCertificate();

                aaron.txtWorkOrderNo.Text = this.dataGridView5.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                openPDF = this.dataGridView5.CurrentRow.Cells["PDF_Link"].Value.ToString();

                aaron.ShowDialog();
            }


            //else if (dataGridView5.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR 1ST APPROVAL")
            //{
            //    openPDF = this.dataGridView5.CurrentRow.Cells["Temporary_Link"].Value.ToString();
            //    System.Diagnostics.Process.Start(location2);
            //}
        }


        private void btnYes5_Click(object sender, EventArgs e)
        {


            List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                  where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                  select row).ToList();
            int selectedRowsss = selectedRows.Count;
            if (selectedRowsss == 0)

            {

                // dataGridView5.Enabled = false;
                //DialogResult dialog = MessageBox.Show("No Checkbox Selected!", "Error!", MessageBoxButtons.OK);

                //if (dialog == DialogResult.OK)
                //{
                //    dataGridView5.Enabled = true;
                //}
            }
            else if (LOGIN.loginAutho == "Admin")
            {
                SqlCommand cmd;

                con.Open();

                foreach (DataGridViewRow rows in selectedRows)
                {
                    cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set CertificateStats = 'APPROVED - INITIAL', OverallStatus = 'FOR FINAL APPROVAL' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                    cmd.ExecuteNonQuery();

                    con.Close();
                    con.Open();
                    string sqlsss = "Select * from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                    SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                    SqlDataReader sdrsss = cmdsss.ExecuteReader();
                    sdrsss.Read();
                    if (sdrsss.HasRows)
                    {


                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = false;

                        oWB = oXL.Workbooks.Open(sdrsss["Temporary_Link"].ToString());
                       // oWB = oXL.Workbooks.Open(sdrsss["Temporary_Link_View"].ToString());
                        con.Close();
                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
                        int intRow = 0;
                        //Row Column


                        oSheet.Cells[48, 2] = LOGIN.loginName;

                        oXL.Visible = false;
                        oXL.UserControl = false;
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range oRanges = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[44, 2];
                            float Lefts = (float)((double)oRanges.Left) + 40;
                            float Tops = (float)((double)oRanges.Top);
                            const float ImageSizes = 80;
                            oSheet.Shapes.AddPicture(@"\\Apbiphbpswb01\ihcs$\" + Home.lblIDNUM + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Lefts, Tops, ImageSizes, ImageSizes);
                        }
                        catch
                        {
                            oSheet.Cells[47, 3] = "NOE - " + Home.lblIDNUM;
                        }

                        //string location1 = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + "View" + ".xlsx";
                        string location = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + ".xlsx";
                        System.IO.Directory.CreateDirectory(@"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\");
                        oSheet.SaveAs(location);

                        con.Open();
                        cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set Temporary_Link = '" + location + "' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                        cmd.ExecuteNonQuery();

                        con.Close();


                        oXL.Quit();

                        System.Diagnostics.Process.Start(location);



                        panel13.Visible = false;

                        label111.Text = "";
                        label110.Text = "";
                        label109.Text = "";
                        label108.Text = "";
                        label107.Text = "";
                        label106.Text = "";
                        label105.Text = "";
                        label104.Text = "";
                        label103.Text = "";
                        label102.Text = "";
                        label155.Text = "";
                        label154.Text = "";
                        label153.Text = "";
                        label152.Text = "";
                        label151.Text = "";
                        label150.Text = "";
                        label149.Text = "";
                        label148.Text = "";
                        label147.Text = "";
                        label146.Text = "";
                        label125.Text = "";

                    }
                    audit_email_CertificateApproved_ADMIN();
                }

                Home._instance.refreshdata5();
                MessageBox.Show("Calibration Confirm Successfully");
                con.Close();

            }


            else if (LOGIN.loginAutho == "Supervisor")
            {
                SqlCommand cmd;


                foreach (DataGridViewRow rows in selectedRows)
                {

                    con.Open();
                    cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set CertificateStats = 'APPROVED - INITIAL', OverallStatus = 'FOR FINAL APPROVAL' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                    cmd.ExecuteNonQuery();

                    string sqlsss = "Select * from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                    SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                    SqlDataReader sdrsss = cmdsss.ExecuteReader();
                    sdrsss.Read();


                    if (sdrsss.HasRows)
                    {

                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = false;

                        oWB = oXL.Workbooks.Open(sdrsss["Temporary_Link"].ToString());
                     //   oWB = oXL.Workbooks.Open(sdrsss["Temporary_Link_View"].ToString());

                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets["CERT1"];

                        int intRow = 0;
                        //Row Column


                        oSheet.Cells[48, 2] = LOGIN.loginName;

                        oXL.Visible = false;
                        oXL.UserControl = false;
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range oRanges = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[44, 2];
                            float Lefts = (float)((double)oRanges.Left) + 40;
                            float Tops = (float)((double)oRanges.Top);
                            const float ImageSizes = 80;
                            oSheet.Shapes.AddPicture(@"\\Apbiphbpswb01\ihcs$\" + Home.lblIDNUM + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Lefts, Tops, ImageSizes, ImageSizes);
                        }
                        catch
                        {
                            oSheet.Cells[47, 3] = "NOE - " + Home.lblIDNUM;
                        }

                        //string location1 = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + "View" + ".xlsx";
                        string location = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + ".xlsx";
                        System.IO.Directory.CreateDirectory(@"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\");
                        oSheet.SaveAs(location);
                        con2.Open();
                        cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set Temporary_Link = '" + location + "' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con2);

                        cmd.ExecuteNonQuery();

                        con2.Close();


                        oXL.Quit();

                        System.Diagnostics.Process.Start(location);


                        panel13.Visible = false;

                        label111.Text = "";
                        label110.Text = "";
                        label109.Text = "";
                        label108.Text = "";
                        label107.Text = "";
                        label106.Text = "";
                        label105.Text = "";
                        label104.Text = "";
                        label103.Text = "";
                        label102.Text = "";
                        label155.Text = "";
                        label154.Text = "";
                        label153.Text = "";
                        label152.Text = "";
                        label151.Text = "";
                        label150.Text = "";
                        label149.Text = "";
                        label148.Text = "";
                        label147.Text = "";
                        label146.Text = "";
                        label125.Text = "";

                    }
                    con.Close();

                    audit_email_CertificateApproved_ADMIN();
                }
                Home._instance.refreshdata5();
                MessageBox.Show("Calibration Confirm Successfully");


            }




            else if (LOGIN.loginAutho == "Manager")
            {
                SqlCommand cmd;



                foreach (DataGridViewRow rows in selectedRows)
                {
                    con.Open();
                    cmd = new SqlCommand("UPDATE tbl_InhouseCalibrationMAIN set CertificateStats = 'APPROVED - FINAL', OverallStatus = 'COMPLETED : CERTIFICATE READY', Status = 'DONE' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con);

                    cmd.ExecuteNonQuery();

                    //Home._instance.refreshdata5();
                    panel13.Visible = false;

                    string sqlsss = "Select * from tbl_InhouseCalibrationMAIN where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'";
                    SqlCommand cmdsss = new SqlCommand(sqlsss, con);
                    SqlDataReader sdrsss = cmdsss.ExecuteReader();
                    sdrsss.Read();
                    if (sdrsss.HasRows)
                    {

                        oXL = new Microsoft.Office.Interop.Excel.Application();
                        oXL.Visible = false;

                        oWB = oXL.Workbooks.Open(sdrsss["Temporary_Link"].ToString());

                        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.Sheets["CERT1"];

                        int intRow = 0;
                        //Row Column


                        oSheet.Cells[48, 9] = LOGIN.loginName;


                        oXL.Visible = false;
                        oXL.UserControl = false;
                        try
                        {
                            Microsoft.Office.Interop.Excel.Range oRanges = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[44, 9];
                            float Lefts = (float)((double)oRanges.Left) + 40;
                            float Tops = (float)((double)oRanges.Top);
                            const float ImageSizes = 80;

                            oSheet.Shapes.AddPicture(@"\\Apbiphbpswb01\ihcs$\" + Home.lblIDNUM + ".png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Lefts, Tops, ImageSizes, ImageSizes);

                        }
                        catch
                        {
                            oSheet.Cells[47, 10] = "NOE - " + Home.lblIDNUM;
                        }
                        con.Close();
                        string workbookPath = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + ".xlsx";
                        string outputPath = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\Approved\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + ".pdf";

                        System.IO.Directory.CreateDirectory(@"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\Approved\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\");
                        oSheet.SaveAs(workbookPath);

                        //try
                        //{
                        //    Microsoft.Office.Interop.Excel.Range oRanges = (Microsoft.Office.Interop.Excel.Range)oSheet.Cells[1, 1];
                        //    float Lefts = (float)((double)oRanges.Left) - 100;
                        //    float Tops = (float)((double)oRanges.Top);
                        //    const float ImageSizes = 700;

                        //    oSheet.Shapes.AddPicture(@"\\Apbiphbpswb01\ihcs$\watermark.png", Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, Lefts, Tops, ImageSizes, ImageSizes);

                        //}
                        //catch
                        //{
                        //    oSheet.Cells[47, 10] = "NOE - " + Home.lblIDNUM;
                        //}
                        //    string workbookPathView = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\For1stApproval\View\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + "View" + ".xlsx";
                        //    string outputPathView = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\Approved\View\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() + "_" + DateTime.Now.Month.ToString() + "-" + System.DateTime.Now.ToString("yyyy MM dd HH MM ss") + "View" + ".pdf";

                        //    System.IO.Directory.CreateDirectory(@"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\Approved\View\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\");
                        //    oSheet.SaveAs(workbookPath);
                            //((Microsoft.Office.Interop.Excel._Worksheet)
                            //oWB.ActiveSheet).PageSetup.Orientation =
                            //Microsoft.Office.Interop.Excel.XlPageOrientation.xlLandscape;

                            //oWB.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                            //Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location) +
                            //    "\\" + @"\\apbiphsh04\B1_BIPHCommon\15_IQC\Certificates InHouse\Approved\" + rows.Cells["WorkOrderNo"].Value.ToString() + "\\" + rows.Cells["WorkOrderNo"].Value.ToString() +".pdf");
                         
                            oXL.Quit();

                            con2.Open();
                            cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set Approved_Link = '" + workbookPath + "', PDF_Link = '" + outputPath + "' where WorkOrderNo = '" + rows.Cells["WorkOrderNo"].Value.ToString() + "'", con2);


                            cmd.ExecuteNonQuery();
                            con2.Close();

                            Microsoft.Office.Interop.Excel.Application excelApplication;
                            Microsoft.Office.Interop.Excel.Workbook excelWorkbook;



                            // Create new instance of Excel
                            excelApplication = new Microsoft.Office.Interop.Excel.Application();

                            // Make the process invisible to the user
                            excelApplication.ScreenUpdating = false;



                            // Make the process silent
                            excelApplication.DisplayAlerts = false;



                            // Open the workbook that you wish to export to PDF
                            excelWorkbook = excelApplication.Workbooks.Open(workbookPath);



                            // If the workbook failed to open, stop, clean up, and bail out
                            if (excelWorkbook == null)
                            {
                                excelApplication.Quit();



                                excelApplication = null;
                                excelWorkbook = null;




                            }



                            var exportSuccessful = true;
                            try
                            {
                                // Call Excel's native export function (valid in Office 2007 and Office 2010, AFAIK)
                                excelWorkbook.ExportAsFixedFormat(Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF, outputPath);
                            }
                            catch (System.Exception ex)
                            {
                                // Mark the export as failed for the return value...
                                exportSuccessful = false;



                                // Do something with any exceptions here, if you wish...
                                // MessageBox.Show...        
                            }
                            finally
                            {
                                // Close the workbook, quit the Excel, and clean up regardless of the results...
                                excelWorkbook.Close();
                                excelApplication.Quit();



                                excelApplication = null;
                                excelWorkbook = null;
                            }



                            // You can use the following method to automatically open the PDF after export if you wish
                            // Make sure that the file actually exists first...
                            if (System.IO.File.Exists(outputPath))
                            {
                                System.Diagnostics.Process.Start(outputPath);
                            }
                            //System.Diagnostics.Process.Start(location);
                            ////Instantiate the Workbook object with the Excel file
                            //Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(location);
                            //workbook.Save(location.Replace(".xlsx", ".pdf"), Aspose.Cells.SaveFormat.Pdf);
                            //workbook.Save(location.Replace(".xlsx", ".png"), Aspose.Cells.SaveFormat.Png);
                            //var uri = location.Replace(".xlsx", ".pdf");
                            //var psi = new System.Diagnostics.ProcessStartInfo();
                            //psi.UseShellExecute = true;
                            //psi.FileName = uri;
                        }

                        label111.Text = "";
                        label110.Text = "";
                        label109.Text = "";
                        label108.Text = "";
                        label107.Text = "";
                        label106.Text = "";
                        label105.Text = "";
                        label104.Text = "";
                        label103.Text = "";
                        label102.Text = "";
                        label155.Text = "";
                        label154.Text = "";
                        label153.Text = "";
                        label152.Text = "";
                        label151.Text = "";
                        label150.Text = "";
                        label149.Text = "";
                        label148.Text = "";
                        label147.Text = "";
                        label146.Text = "";
                        con.Close();

                    }

                    audit_email_CertificateApproved_MANAGER();
                }
                MessageBox.Show("Calibration Confirm Successfully");
                Home._instance.refreshdata5();
            

           
        }
        

        private void btnRefresh5_Click(object sender, EventArgs e)
        {
            DisplayData6();
            tabPage6.Text = "MH SUMMARY " + "(" + dataGridView6.RowCount.ToString() + ")";
        }

        private void btnRefresh4_Click(object sender, EventArgs e)
        {
            DisplayData4();
            tabPage4.Text = "ACTUAL CALIBRATION " + "(" + dataGridView4.RowCount.ToString() + ")";
        }

        private void button15_Click(object sender, EventArgs e)
        {
            panel13.Visible = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            panel12.Visible = false;
        }

        //FILTER DATAGRIDVIEW 4
        private void btnFilter4_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
           



            if (cmbSelMonth4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSelSection4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];

  



                }
            }

            if (cmbSelSection4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbSelMonth4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];

                }
            }

            if (cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbSelMonth4.Text + "' AND ReqSection = '" + cmbSelSection4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];
                }
            }

            if (cmbSelMonth4.Text == "" && cmbSelSection4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where OverallStatus = '" + cmbSelStat4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];

                }
            }
            if (cmbSelSection4.Text == "" && cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {

                    string sql = "Select * from tbl_InhouseCalibrationMAIN where CalibrationPlan = '" + cmbSelMonth4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];


                }
            }

            if (cmbSelMonth4.Text == "" && cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSelSection4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];

                }
            }
            else
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select * from tbl_InhouseCalibrationMAIN where ReqSection = '" + cmbSelSection4.Text + "' AND CalibrationPlan = '" + cmbSelMonth4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView4.DataSource = ds.Tables[0];


                }
            }
        }
        //FILTER DATAGRIDVIEW5
        private void btnFilter5_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            if (cmbSelMonth5.Text == "APR 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "MAY 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "JUN 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "JUL 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "AUG 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "SEP 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "OCT 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "NOV 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];
                }
            }
            else if (cmbSelMonth5.Text == "DEC 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "JAN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "FEB 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "MAR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "APR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "MAY 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }

            else if (cmbSelMonth5.Text == "JUN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "JUL 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "AUG 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "SEP 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "OCT 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "NOV 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelMonth5.Text == "DEC 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "JAN 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "FEB 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelMonth5.Text == "MAR 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection5.Text == "IQC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'IQC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection5.Text == "EE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'EE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection5.Text == "INK CARTRIDGE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'INK CARTRIDGE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection5.Text == "TAPE CASSETTE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TAPE CASSETTE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection5.Text == "MOLDING")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'MOLDING'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection5.Text == "PRINTER(A3)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(A3)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection5.Text == "PRINTER(MINI)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(MINI)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection5.Text == "PT")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PT'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection5.Text == "TC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,  Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }
            }

            else if (cmbSelStat5.Text == "FOR ACCEPTANCE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR ACCEPTANCE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }
            else if (cmbSelStat5.Text == "FOR FINAL ACCEPTANCE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL ACCEPTANCE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }
            else if (cmbSelStat5.Text == "FOR CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }
            else if (cmbSelStat5.Text == "ON-GOING CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'ON-GOING CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }

            else if (cmbSelStat5.Text == "FINISHED CALIBRATION")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FINISHED CALIBRATION'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }
            else if (cmbSelStat5.Text == "FOR 1ST APPROVAL")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR 1ST APPROVAL'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }

            else if (cmbSelStat5.Text == "FOR FINAL APPROVAL")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL APPROVAL'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }

            else if (cmbSelStat5.Text == "COMPLETED : CERTIFICATE READY")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'COMPLETED : CERTIFICATE READY'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }

            else if (cmbSelStat5.Text == "DISAPPROVED")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where OverallStatus = 'DISAPPROVED'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView5.DataSource = ds.Tables[0];

                }

            }
        }

        private void btnDownloadFilter3_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView3.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView3.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView3.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView3.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }

        private void btnDownloadFilter4_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView4.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView4.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView4.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView4.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }

        private void btnDownloadFilter5_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView5.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView5.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView5.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView5.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView5.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult dialogResult = MessageBox.Show("Are you sure you want to EXPORT to EXCEL?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dialogResult == DialogResult.Yes)
            {
                Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();

                Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);

                Microsoft.Office.Interop.Excel._Worksheet worksheet = null;

                app.Visible = true;

                worksheet = workbook.Sheets["Sheet1"];
                worksheet = workbook.ActiveSheet;

                worksheet.Name = "Export to EXCEL";
                for (int i = 1; i < dataGridView6.Columns.Count + 1; i++)
                {
                    worksheet.Cells[1, i] = dataGridView6.Columns[i - 1].HeaderText;
                }

                for (int i = 0; i < dataGridView6.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView6.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataGridView6.Rows[i].Cells[j].Value.ToString();
                    }
                }
            }
            else if (dialogResult == DialogResult.No)
            {
                //do Nothing
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {

        }

        private void btnSearch5_Click(object sender, EventArgs e)
        {
            searchACC7();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ViewMH view = new ViewMH();
            view.ShowDialog();
        }

        private void cmbSection_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {
            cmbMonth.Text = "";
            cmbSection.Text = "";
            cmbSelectMonth2.Text = "";
            cmbSelectSection2.Text = "";
            cmbSelMonth.Text = "";
            cmbStatus.Text = "";
            cmbStatus2.Text = "";
            btnSelSection.Text = "";
            btnSelStat.Text = "";
            cmbSelStat4.Text = "";
            cmbSelSection4.Text = "";
            cmbSelMonth4.Text = "";
            cmbSelStat5.Text = "";
            cmbSelSection5.Text = "";
            cmbSelMonth5.Text = "";
            cmbSelSection6.Text = "";
            cmbCalPlan6.Text = "";
        }

        private void txtWorkOrder2_Enter(object sender, EventArgs e)
        {
            //if (txtWorkOrder2.Text == "Input")
            //{
            //    txtWorkOrder2.Text = "";
            //    txtWorkOrder2.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void txtWorkOrder2_Leave(object sender, EventArgs e)
        {
            //if (txtWorkOrder2.Text == "")
            //{
            //    txtWorkOrder2.Text = "Input";
            //    txtWorkOrder2.ForeColor = Color.DarkGray;
            //}
        }

        private void txtWorkOrder5_Enter(object sender, EventArgs e)
        {
            //if (txtWorkOrder5.Text == "Input")
            //{
            //    txtWorkOrder5.Text = "";
            //    txtWorkOrder5.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void txtWorkOrder5_Leave(object sender, EventArgs e)
        {
            //if (txtWorkOrder5.Text == "")
            //{
            //    txtWorkOrder5.Text = "Input";
            //    txtWorkOrder5.ForeColor = Color.DarkGray;
            //}
        }

        private void textBox2_Enter(object sender, EventArgs e)
        {
            //if (txtWorkOrderNo6.Text == "Input")
            //{
            //    txtWorkOrderNo6.Text = "";
            //    txtWorkOrderNo6.ForeColor = Color.FromArgb(26, 44, 47);
            //}
        }

        private void textBox2_Leave(object sender, EventArgs e)
        {
            //if (txtWorkOrderNo6.Text == "")
            //{
            //    txtWorkOrderNo6.Text = "Input";
            //    txtWorkOrderNo6.ForeColor = Color.DarkGray;
            //}
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            ViewLegend view = new ViewLegend();
            view.ShowDialog();
        }

        private void dataGridView5_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView5.CurrentRow.Index;

            if (LOGIN.loginAutho == "Admin")
            {
                if (dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                {
                    List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                          where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                          select row).ToList();
                    int selectedRowsss = selectedRows.Count;
                    if (dataGridView5.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR FINAL APPROVAL")
                    {
                        dataGridView5.Columns["Select"].ReadOnly = true;
                        dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                        dataGridView5.Columns["Select"].ReadOnly = false;
                    }
                    else if (selectedRowsss == 9)
                    {
                        dataGridView5.Columns["Select"].ReadOnly = true;
                    }

                }

                else if (LOGIN.loginAutho == "Supervisor")
                {
                    if (dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                    {
                        List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                              where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                              select row).ToList();
                        int selectedRowsss = selectedRows.Count;
                        if (dataGridView5.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR FINAL APPROVAL")
                        {
                            dataGridView5.Columns["Select"].ReadOnly = true;
                            dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                            dataGridView5.Columns["Select"].ReadOnly = false;
                        }
                        else if (selectedRowsss == 9)
                        {
                            dataGridView5.Columns["Select"].ReadOnly = true;
                        }

                    }
                }
                else if (LOGIN.loginAutho == "Manager")
                {
                    if (dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                    {
                        List<DataGridViewRow> selectedRows = (from row in dataGridView5.Rows.Cast<DataGridViewRow>()
                                                              where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                              select row).ToList();
                        int selectedRowsss = selectedRows.Count;
                        if (dataGridView5.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR 1ST APPROVAL")
                        {
                            dataGridView5.Columns["Select"].ReadOnly = true;
                            dataGridView5.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                            dataGridView5.Columns["Select"].ReadOnly = false;
                        }
                        else if (selectedRowsss == 9)
                        {
                            dataGridView5.Columns["Select"].ReadOnly = true;
                        }

                    }
                }
                }
            }
        
        

        private void panel10_Paint(object sender, PaintEventArgs e)
        {

        }

        private void panel12_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnCancel6_Click(object sender, EventArgs e)
        {
            panel12.Visible = false;
        }


        private void btnApproveManager_Click(object sender, EventArgs e)
        {
            //panel10.Visible = true;
            //dataGridView3.Columns["Select"].ReadOnly = false;
        
            //List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
            //                                      where Convert.ToBoolean(row.Cells["Select"].Value) == true
            //                                      select row).ToList();
            //int selectedRowsss = selectedRows.Count;
            //if (selectedRowsss == 0)
            //{

            //}
            //else
            //{
            //    int i = 0;
            //    foreach (DataGridViewRow rows in selectedRows)
            //    {
            //        i++;
            //        if (i == 1)
            //        {
            //            label218.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label217.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 2)
            //        {
            //            label216.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label215.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 3)
            //        {
            //            label214.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label213.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 4)
            //        {
            //            label212.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label211.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 5)
            //        {
            //            label210.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label209.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 6)
            //        {
            //            label203.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label202.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 7)
            //        {
            //            label201.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label200.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 8)
            //        {
            //            label199.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label198.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 9)
            //        {
            //            label197.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label196.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }
            //        else if (i == 10)
            //        {
            //            label195.Text = rows.Cells["WorkOrderNo"].Value.ToString();
            //            label194.Text = rows.Cells["InstRefNo"].Value.ToString();
            //        }

            //    }
            //}
        }

        private void btnDisapproveManager_Click(object sender, EventArgs e)
        {

        }


        private void btnFilter6_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            if (cmbCalPlan6.Text == "APR 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "MAY 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "JUN 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "JUL 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "AUG 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "SEP 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "OCT 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "NOV 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];
                }
            }
            else if (cmbCalPlan6.Text == "DEC 2022")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2022'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "JAN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "FEB 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "MAR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "APR 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'APR 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "MAY 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAY 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }

            else if (cmbCalPlan6.Text == "JUN 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUN 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "JUL 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JUL 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "AUG 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'AUG 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "SEP 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'SEP 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "OCT 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'OCT 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "NOV 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'NOV 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbCalPlan6.Text == "DEC 2023")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'DEC 2023'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "JAN 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'JAN 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "FEB 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'FEB 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbCalPlan6.Text == "MAR 2024")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where CalibrationPlan = 'MAR 2024'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection6.Text == "IQC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'IQC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection6.Text == "EE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'EE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection6.Text == "INK CARTRIDGE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'INK CARTRIDGE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection6.Text == "TAPE CASSETTE")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TAPE CASSETTE'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection6.Text == "MOLDING")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'MOLDING'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection6.Text == "PRINTER(A3)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(A3)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection6.Text == "PRINTER(MINI)")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PRINTER(MINI)'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
            else if (cmbSelSection6.Text == "PT")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'PT'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];


                }
            }
            else if (cmbSelSection6.Text == "TC")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where ReqSection = 'TC'";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView6.DataSource = ds.Tables[0];

                }
            }
        }

        private void SearchACC8()
        {
            dataGridView6.DataSource = null;
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                
                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks,Picture_Link, OverallStatus from tbl_InhouseCalibrationMAIN where WorkOrderNo like @workorderno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@workorderno", "%" + txtWorkOrderNo6.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
                dataGridView6.Columns[0].HeaderText = "No.";

            }
        }

        private void SearchACC9()
        {
            dataGridView6.DataSource = null;
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {

                string sql = "Select ID, WorkOrderNo, DateReq, ReqSection, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan, Remarks, Picture_Link,OverallStatus from tbl_InhouseCalibrationMAIN where InstRefNo like @workorderno";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                cmd.Parameters.Add(new SqlParameter("@workorderno", "%" + txtInsRefNo6.Text + "%"));
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView6.DataSource = ds.Tables[0];
                dataGridView6.Columns[0].HeaderText = "No.";

            }
        }

        private void btnSearch6_Click(object sender, EventArgs e)
        {
            SearchACC8();
        }

        private void btnChecker_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR 1ST APPROVAL'", con);
            adapt.Fill(dt);
            dataGridView5.DataSource = dt;
            dataGridView5.Columns["EmpNo"].Visible = false;
            dataGridView5.Columns["Email"].Visible = false;
            dataGridView5.Columns["InstrumentLoc"].Visible = false;
            dataGridView5.Columns["Approved_Link"].Visible = false;
            dataGridView5.Columns["PDF_Link"].Visible = false;
            dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
            dataGridView5.Columns["Picture_Link"].Visible = false;
            dataGridView5.Columns["CertificateName"].Visible = false;
            this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
            dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
            dataGridView5.Columns[0].HeaderText = "No.";
            dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
            dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
            dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
        }

        private void btnApprover_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select ID, WorkOrderNo, DateReq, ReqSection, Approved_Link, Picture_Link, CertificateName, PDF_Link, EmpNo, ReqBy, Email, CostCenter, InstrumentCat, MeasureInstName, InstRefNo, Manufacturer, Model, SerialNo, InstrumentLoc, CalibrationFreq, CalibrationDueDate, CalibrationPlan,  DateReceived, EndorsedBy, ReceivedBy, CalibrationIncharge, CalibrationStartDate, StartTime,CalibrationEndDate,EndTime,CalibrationTimehrs,Status,CalibrationResult,CalibrationNo,CertificateReleaseDate,OverallStatus,Remarks from tbl_InhouseCalibrationMAIN where OverallStatus = 'FOR FINAL APPROVAL'", con);
            adapt.Fill(dt);
            dataGridView5.DataSource = dt;
            dataGridView5.Columns["EmpNo"].Visible = false;
            dataGridView5.Columns["Email"].Visible = false;
            dataGridView5.Columns["InstrumentLoc"].Visible = false;
            dataGridView5.Columns["Approved_Link"].Visible = false;
            dataGridView5.Columns["PDF_Link"].Visible = false;
            dataGridView5.Columns["CertificateReleaseDate"].Visible = false;
            dataGridView5.Columns["Picture_Link"].Visible = false;
            dataGridView5.Columns["CertificateName"].Visible = false;
            this.dataGridView5.DefaultCellStyle.Font = new Font("Calibri", 8);
            dataGridView5.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 9, FontStyle.Bold);
            dataGridView5.Columns[0].HeaderText = "No.";
            dataGridView5.Columns["CostCenter"].HeaderText = "CostCode";
            dataGridView5.Columns["CalibrationDueDate"].HeaderText = "Due Date";
            dataGridView5.Columns["CalibrationFreq"].HeaderText = "Frequency";
        }

        private void btnDownloadForm_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            //dataGridView3.Columns[0].Frozen = true;
            //dataGridView3.Columns[1].Frozen = true;
            //dataGridView3.Columns[2].Frozen = true;
            //dataGridView3.Columns[3].Frozen = true;
            //dataGridView3.Columns[4].Frozen = true;
            //dataGridView3.Columns[5].Frozen = true;
            //dataGridView3.Columns[6].Frozen = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView5.Columns["Select"].ReadOnly = false;
            DisplayData5();
            tabPage5.Text = "CERTIFICATE " + "(" + dataGridView5.RowCount.ToString() + ")";
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            int index = dataGridView3.CurrentRow.Index;

            if (LOGIN.loginAutho == "Admin")
            {
                if (dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                {
                    List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                          where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                          select row).ToList();
                    int selectedRowsss = selectedRows.Count;
                    if (dataGridView3.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR FINAL ACCEPTANCE")
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                        dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                        dataGridView3.Columns["Select"].ReadOnly = false;
                    }
                    else if (selectedRowsss == 9)
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                    }

                }
            }
           else if (LOGIN.loginAutho == "Supervisor")
            {
                if (dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                {
                    List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                          where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                          select row).ToList();
                    int selectedRowsss = selectedRows.Count;
                    if (dataGridView3.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR FINAL ACCEPTANCE")
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                        dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                        dataGridView3.Columns["Select"].ReadOnly = false;
                    }
                    else if (selectedRowsss == 9)
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                    }

                }
            }
            else if (LOGIN.loginAutho == "Manager")
            {
                if (dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex] is DataGridViewCheckBoxCell)
                {
                    List<DataGridViewRow> selectedRows = (from row in dataGridView3.Rows.Cast<DataGridViewRow>()
                                                          where Convert.ToBoolean(row.Cells["Select"].Value) == true
                                                          select row).ToList();
                    int selectedRowsss = selectedRows.Count;
                    if (dataGridView3.CurrentRow.Cells["OverallStatus"].Value.ToString() == "FOR ACCEPTANCE")
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                        dataGridView3.Rows[e.RowIndex].Cells[e.ColumnIndex].Value = false;
                        dataGridView3.Columns["Select"].ReadOnly = false;
                    }
                    else if (selectedRowsss == 9)
                    {
                        dataGridView3.Columns["Select"].ReadOnly = true;
                    }

                }
            }
            
            
        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click_1(object sender, EventArgs e)
        {

        }

        private void button5_Click(object sender, EventArgs e)
        {
            searchACC4();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            SearchACC9();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            searchACC6();
        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString() == "NOT YET RECEIVE")
            {
                btnUpdateDetails.Enabled = false;
                btnRecieving.Enabled = true;
            }
            else
            {
                btnUpdateDetails.Enabled = true;
                btnRecieving.Enabled = false;
            }

            if (dataGridView4.CurrentRow.Cells["OverallStatus"].Value.ToString() == "DISAPPROVED")
            {
                btnUpdateDetails.Enabled = true;
            }
        }

        public static string openPDF;
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
            {
                CalibrationCertificate aaron = new CalibrationCertificate();

                aaron.txtWorkOrderNo.Text = this.dataGridView1.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                openPDF = this.dataGridView1.CurrentRow.Cells["PDF_Link"].Value.ToString();

                aaron.ShowDialog();
            }
            else
            {
               // openPDF = this.dataGridView1.CurrentRow.Cells[""].Value.ToString();
            }
        }

        private void dataGridView5_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            try
            {
                btnTempCert.Enabled = true;
                ID = Convert.ToInt32(dataGridView5.Rows[e.RowIndex].Cells["ID"].Value.ToString());
                lblWorkOrderNo.Text = dataGridView5.Rows[e.RowIndex].Cells["WorkOrderNo"].Value.ToString();
                lblTempLink.Text = dataGridView5.Rows[e.RowIndex].Cells["Temporary_Link"].Value.ToString();
            }
            catch
            {

            }
           
        }

        private void btnTempCert_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(lblTempLink.Text);
            MessageBox.Show("OPENING. . . .");
        }

        private void btnUpdateDetail2_Click(object sender, EventArgs e)
        {
            if (LOGIN.loginAutho == "Admin" && dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
            {
                UpdateCalibration lookdown2 = new UpdateCalibration();
                UpdateCalibration._instance.InstrumentCat();
                UpdateCalibration._instance.CalibrationFreq();
                UpdateCalibration._instance.CalibrationPlan();
                UpdateCalibration._instance.CalibIncharge();
                // UpdateCalibration._instance.SelectForm();
                lookdown2.txtWorkOrderNo.Text = this.dataGridView1.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                lookdown2.cmbInstrumentCat.Text = this.dataGridView1.CurrentRow.Cells["InstrumentCat"].Value.ToString();
                lookdown2.txtMinstrumentName.Text = this.dataGridView1.CurrentRow.Cells["MeasureInstName"].Value.ToString();
                lookdown2.txtInstrumentRefNo.Text = this.dataGridView1.CurrentRow.Cells["InstRefNo"].Value.ToString();
                lookdown2.txtManufucturer.Text = this.dataGridView1.CurrentRow.Cells["Manufacturer"].Value.ToString();
                lookdown2.txtModel.Text = this.dataGridView1.CurrentRow.Cells["Model"].Value.ToString();
                lookdown2.txtSerialNo.Text = this.dataGridView1.CurrentRow.Cells["SerialNo"].Value.ToString();
                lookdown2.txtInstrumentLoc.Text = this.dataGridView1.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
                lookdown2.cmbCalFreq.Text = this.dataGridView1.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
                lookdown2.dtDueDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
                lookdown2.cmbCalPlan.Text = this.dataGridView1.CurrentRow.Cells["calibrationPlan"].Value.ToString();
                lookdown2.txtReqRemarks.Text = this.dataGridView1.CurrentRow.Cells["Remarks"].Value.ToString();
                lookdown2.cmbCalInCharge.Text = this.dataGridView1.CurrentRow.Cells["CalibrationIncharge"].Value.ToString();
                lookdown2.dtCalStartDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationStartDate"].Value.ToString();
                lookdown2.dtCalEndDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationEndDate"].Value.ToString();
                lookdown2.cmbStartTime.Text = this.dataGridView1.CurrentRow.Cells["StartTime"].Value.ToString();
                lookdown2.cmbEndTime.Text = this.dataGridView1.CurrentRow.Cells["EndTime"].Value.ToString();
                lookdown2.txtAutoCompute.Text = this.dataGridView1.CurrentRow.Cells["CalibrationTimehrs"].Value.ToString();
                lookdown2.txtReqSection.Text = this.dataGridView1.CurrentRow.Cells["ReqSection"].Value.ToString();
                lookdown2.txtEmail.Text = this.dataGridView1.CurrentRow.Cells["Email"].Value.ToString();
                lookdown2.picInstrument.ImageLocation = this.dataGridView1.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown2.txtLinkPic.Text = this.dataGridView1.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown2.lblStat.Text = this.dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown2.SelectForm.Text = this.dataGridView1.CurrentRow.Cells["CertificateName"].Value.ToString();
                lookdown2.comboBox1.Text = this.dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown2.cmbCalResult.Text = this.dataGridView1.CurrentRow.Cells["CalibrationResult"].Value.ToString();
                //  lookdown.SelectForm.Enabled = false;
                lookdown2.ShowDialog();
            }

            else if (LOGIN.loginAutho == "Supervisor" && dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString() == "COMPLETED : CERTIFICATE READY")
            {
                UpdateCalibration lookdown2 = new UpdateCalibration();
                UpdateCalibration._instance.InstrumentCat();
                UpdateCalibration._instance.CalibrationFreq();
                UpdateCalibration._instance.CalibrationPlan();
                UpdateCalibration._instance.CalibIncharge();
                // UpdateCalibration._instance.SelectForm();
                lookdown2.txtWorkOrderNo.Text = this.dataGridView1.CurrentRow.Cells["WorkOrderNo"].Value.ToString();
                lookdown2.cmbInstrumentCat.Text = this.dataGridView1.CurrentRow.Cells["InstrumentCat"].Value.ToString();
                lookdown2.txtMinstrumentName.Text = this.dataGridView1.CurrentRow.Cells["MeasureInstName"].Value.ToString();
                lookdown2.txtInstrumentRefNo.Text = this.dataGridView1.CurrentRow.Cells["InstRefNo"].Value.ToString();
                lookdown2.txtManufucturer.Text = this.dataGridView1.CurrentRow.Cells["Manufacturer"].Value.ToString();
                lookdown2.txtModel.Text = this.dataGridView1.CurrentRow.Cells["Model"].Value.ToString();
                lookdown2.txtSerialNo.Text = this.dataGridView1.CurrentRow.Cells["SerialNo"].Value.ToString();
                lookdown2.txtInstrumentLoc.Text = this.dataGridView1.CurrentRow.Cells["InstrumentLoc"].Value.ToString();
                lookdown2.cmbCalFreq.Text = this.dataGridView1.CurrentRow.Cells["CalibrationFreq"].Value.ToString();
                lookdown2.dtDueDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationDueDate"].Value.ToString();
                lookdown2.cmbCalPlan.Text = this.dataGridView1.CurrentRow.Cells["calibrationPlan"].Value.ToString();
                lookdown2.txtReqRemarks.Text = this.dataGridView1.CurrentRow.Cells["Remarks"].Value.ToString();
                lookdown2.cmbCalInCharge.Text = this.dataGridView1.CurrentRow.Cells["CalibrationIncharge"].Value.ToString();
                lookdown2.dtCalStartDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationStartDate"].Value.ToString();
                lookdown2.dtCalEndDate.Text = this.dataGridView1.CurrentRow.Cells["CalibrationEndDate"].Value.ToString();
                lookdown2.cmbStartTime.Text = this.dataGridView1.CurrentRow.Cells["StartTime"].Value.ToString();
                lookdown2.cmbEndTime.Text = this.dataGridView1.CurrentRow.Cells["EndTime"].Value.ToString();
                lookdown2.txtAutoCompute.Text = this.dataGridView1.CurrentRow.Cells["CalibrationTimehrs"].Value.ToString();
                lookdown2.txtReqSection.Text = this.dataGridView1.CurrentRow.Cells["ReqSection"].Value.ToString();
                lookdown2.txtEmail.Text = this.dataGridView1.CurrentRow.Cells["Email"].Value.ToString();
                lookdown2.picInstrument.ImageLocation = this.dataGridView1.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown2.txtLinkPic.Text = this.dataGridView1.CurrentRow.Cells["Picture_Link"].Value.ToString();
                lookdown2.lblStat.Text = this.dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown2.SelectForm.Text = this.dataGridView1.CurrentRow.Cells["CertificateName"].Value.ToString();
                lookdown2.comboBox1.Text = this.dataGridView1.CurrentRow.Cells["OverallStatus"].Value.ToString();
                lookdown2.cmbCalResult.Text = this.dataGridView1.CurrentRow.Cells["CalibrationResult"].Value.ToString();
                //                                   lookdown.SelectForm.Enabled = false;
                lookdown2.ShowDialog();
            }
        }

        private void btnAboutUs_Click(object sender, EventArgs e)
        {
            ABOUT about = new ABOUT();
            about.ShowDialog();
        }
    }
    }

    
 
    

