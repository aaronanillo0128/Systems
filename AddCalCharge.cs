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
    public partial class AddCalCharge : Form
    {
        static string AAA = System.Configuration.ConfigurationManager.ConnectionStrings["dbconnection"].ConnectionString;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        SqlCommand cmd;
        int ID;
        public AddCalCharge()
        {
            InitializeComponent();
        }

        private void DisplayData()
        {
            DataTable dt = new DataTable();
            SqlDataAdapter adapt = new SqlDataAdapter("Select * from [dbo].[tbl_CalibrationInCharge]", con);
            adapt.Fill(dt);
            dataGridView1.DataSource = dt;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new System.Drawing.Font("Calibri", 13, FontStyle.Bold);
        }

        private void btnDelete_Click(object sender, EventArgs e)
        {
            if (ID != 0)
            {
                DialogResult dialogResult = MessageBox.Show("Are you sure you want to PROCEED?", "Verification!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (dialogResult == DialogResult.Yes)
                {
                    cmd = new SqlCommand("delete [dbo].[tbl_CalibrationInCharge] where ID=@id", con);
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

        private void btnAdd_Click(object sender, EventArgs e)
        {
            if(txtCalIncharge.Text != "")
            {
                cmd = new SqlCommand("insert into [dbo].[tbl_CalibrationInCharge](CalibrationInCharge) Values (@calibrationincharge)", con);
                con.Open();
                cmd.Parameters.AddWithValue("@calibrationincharge", txtCalIncharge.Text);
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Record Inserted Successfully", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                DisplayData();
            }
            else
            {
                MessageBox.Show("NO DATA TO BE ADD!");
            }
           
        }

        private void btnUpdate_Click(object sender, EventArgs e)
        {
            if (txtCalIncharge.Text != "")
            {
                cmd = new SqlCommand("update [dbo].[tbl_CalibrationInCharge] set CalibrationInCharge=@calibrationincharge where ID=@id", con);
                con.Open();
                cmd.Parameters.AddWithValue("@id", ID);
                cmd.Parameters.AddWithValue("@calibrationincharge", txtCalIncharge.Text);
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

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            ID = Convert.ToInt32(dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString());
            txtCalIncharge.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void AddCalCharge_Load(object sender, EventArgs e)
        {
            DisplayData();
        }
    }
}
