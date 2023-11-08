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
    public partial class AddCalPlan : Form
    {
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        SqlDataAdapter adapt;
        DataTable dt;
        int ID;
        public AddCalPlan()
        {
            InitializeComponent();
        }

        private void AddCalPlan_Load(object sender, EventArgs e)
        {
            DisplayData();
        }
        private void DisplayData()
        {
            DataTable dt = new DataTable();
            adapt = new SqlDataAdapter("Select * from tbl_CalibrationPlan", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 13, FontStyle.Bold);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtCalPlan.Text != "")
            {
                string SElectDouble_SQL = string.Empty;
                SElectDouble_SQL = "SELECT CalibrationPlan from tbl_CalibrationPlan where CalibrationPlan = '" + txtCalPlan.Text + "'";
                bool DoubleData = false;
                DoubleData = CRUD.RETRIEVESINGLE(SElectDouble_SQL);
                if (DoubleData == true)
                {
                    MessageBox.Show("DUPLICATE ENTRY DETECTED!", "In-House Calibration", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                else
                {
                    cmd = new SqlCommand("insert into tbl_CalibrationPlan(CalibrationPlan) Values (@calibrationPlan)", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@calibrationPlan", txtCalPlan.Text);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Inserted Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    DisplayData();

                }

            }
            else
            {
                MessageBox.Show("Please Provide Details!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            dataGridView1.Columns["CalibrationPlan"].Width = 271;
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (txtCalPlan.Text != "")
            {
                cmd = new SqlCommand("update tbl_CalibrationFreq set CalibrationFreq=@calibrationfreq where ID=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@calibrationfreq", txtCalPlan.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Updated Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                DisplayData();

            }
            else
            {
                MessageBox.Show("Please Select Record to Update!", "Warning!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (ID != 0)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to PROCEED?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    cmd = new SqlCommand("delete tbl_CalibrationPlan where ID=@id", con);
                    con.Open();
                    cmd.Parameters.AddWithValue("@id", ID);
                    cmd.ExecuteNonQuery();
                    con.Close();
                    MessageBox.Show("Record Deleted Successfully!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    DisplayData();

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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtCalPlan.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
        }
    }
}
