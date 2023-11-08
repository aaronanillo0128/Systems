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
    public partial class ViewMH : Form
    {
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        DataTable dt;
        int ID;
        public ViewMH()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void SelectCalMonth()
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
                cmbSelMonth4.DisplayMember = "CalibrationPlan";
                cmbSelMonth4.ValueMember = "CalibrationPlan";
                cmbSelMonth4.DataSource = dt.Tables["CalibrationPlan"];
                conn.Close();

            }
        }

        public void SelectSection()
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
                cmbSelSection4.DisplayMember = "SectionList";
                cmbSelSection4.ValueMember = "SectionList";
                cmbSelSection4.DataSource = dt.Tables["SectionList"];
                conn.Close();

            }
        }

        public void Status()
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
                cmbSelStat4.DisplayMember = "CalibrationStatus";
                cmbSelStat4.ValueMember = "CalibrationStatus";
                cmbSelStat4.DataSource = dt.Tables["CalibrationStatus"];
                conn.Close();

            }
        }
        private void DisplayData()
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN GROUP BY ReqSection", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 13, FontStyle.Bold);
            dataGridView1.Columns["ReqSection"].HeaderText = "Section";
        }
        string TOTALMH;
        private void ViewMH_Load(object sender, EventArgs e)
        {
            
            DisplayData();
            con.Open();
          
          
            SqlCommand cmd = new SqlCommand("Select SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN", con);
            SqlDataReader reader = cmd.ExecuteReader();

            if (reader.Read())
            {
                
                TOTALMH= reader[0].ToString();
                txtTotalMH.Text = TOTALMH;
            }
            con.Close();
            SelectCalMonth();
            SelectSection();
            Status();
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dataGridView1.Columns["ReqSection"].Width = 280;
           dataGridView1.Columns["TOTAL MH"].Width = 247;
        }

        private void filteringCalPlan()
        {
            DataSet ds = new DataSet();
            using (SqlConnection conn = new SqlConnection(AAA))
            {
                string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where CalibrationPlan = '" + cmbSelMonth4.Text + "' AND ReqSection = '" + cmbSelSection4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "' Group by ReqSection";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.CommandType = CommandType.Text;
                SqlDataAdapter sda = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                sda.Fill(ds);
                dataGridView1.DataSource = ds.Tables[0];

            }
           
          if (cmbSelMonth4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where ReqSection = '" + cmbSelSection4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                  
                }
  
            }
          if(cmbSelSection4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where CalibrationPlan = '" + cmbSelMonth4.Text + "' AND OverallStatus = '" + cmbSelStat4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    txtTotalMH.Text = TOTALMH;
                }
            }
          if(cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where CalibrationPlan = '" + cmbSelMonth4.Text + "' AND ReqSection = '" + cmbSelSection4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    txtTotalMH.Text = TOTALMH;
                }
            }

          if (cmbSelMonth4.Text == "" && cmbSelSection4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where OverallStatus = '" + cmbSelStat4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    txtTotalMH.Text = TOTALMH;
                }
            }

          if(cmbSelSection4.Text == "" && cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where CalibrationPlan = '" + cmbSelMonth4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    txtTotalMH.Text = TOTALMH;
                }
            }

          if(cmbSelMonth4.Text == "" && cmbSelStat4.Text == "")
            {
                using (SqlConnection conn = new SqlConnection(AAA))
                {
                    string sql = "Select ReqSection, SUM(ISNULL(CASE WHEN ISNUMERIC([CalibrationTimehrs]) = 1 THEN CONVERT(DECIMAL(10,2),[CalibrationTimehrs]) ELSE 0 END, 0)) AS [TOTAL MH] from tbl_InhouseCalibrationMAIN Where ReqSection = '" + cmbSelSection4.Text + "' Group by ReqSection";
                    SqlCommand cmd = new SqlCommand(sql, conn);
                    cmd.CommandType = CommandType.Text;
                    SqlDataAdapter sda = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    sda.Fill(ds);
                    dataGridView1.DataSource = ds.Tables[0];
                    txtTotalMH.Text = TOTALMH;
                }
            }

          if (cmbSelStat4.Text == "" && cmbSelSection4.Text == "" && cmbSelMonth4.Text == "")
            {
                DisplayData();
            }
        }
        private void btnFilter4_Click(object sender, EventArgs e)
        {
            filteringCalPlan();
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

        private void button1_Click(object sender, EventArgs e)
        {
          //  filteringSection();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void btnTableau_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
        
        }

        private void cmbSelMonth4_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
