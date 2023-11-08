using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class UploadCertificate : Form
    {
        SqlCommand cmd;
        SqlConnection con = new SqlConnection("Data Source=apbiphdb18;Initial Catalog=InHouse_Calibration;Persist Security Info=True;User ID=IHCS1;Password=P@ssw0rd");
        public UploadCertificate()
        {
            InitializeComponent();
        }

        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;

            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xlsx") == 0)//compare the extension of the file
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtLink.Text + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';";//for below excel 2007
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + txtLink.Text + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1';";//for above excel 2007
            using (OleDbConnection con = new OleDbConnection(conn))
            {

                if (con.State == ConnectionState.Closed)
                {
                    con.Open();
                }
                try
                {
                    string Sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();

                    Sheet1 = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null).Rows[0]["TABLE_NAME"].ToString();
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [" + Sheet1 + "]", con);//here we read data from sheet1
                    oleAdpt.Fill(dtexcel);//fill excel data into dataTable
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }// end ReadExcel
        DataTable dtExcel;
        private void LoadExcel()
        {
            dtExcel = ReadExcel(filePath, fileExt);//read excel file
           dataGridView1.Invoke(new Action(() => dataGridView1.Visible = true));
          dataGridView1.Invoke(new Action(() => dataGridView1.DataSource = dtExcel));

        }
        string fileExt;
        string filePath;
        private void btnSelectFile_Click(object sender, EventArgs e)
        {

            OpenFileDialog file = new OpenFileDialog();//open dialog to choose file
            if (file.ShowDialog() == System.Windows.Forms.DialogResult.OK)//if there is a file choosen by the user
            {
                filePath = file.FileName;//get the path of the file
                fileExt = Path.GetExtension(filePath);//get the file extension
                txtLink.Text = filePath;
                txtLink.Text = txtLink.Text.Replace(".xlsx.xlsx", ".xlsx").Replace(".xls.xls", ".xls");

                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        LoadExcel();

                    }
                    catch (Exception ex)
                    {
                       // MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);//custom messageBox to show error
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnUploadFile_Click(object sender, EventArgs e)
        {
            if (txtLink.Text == "")
            {
                MessageBox.Show("Please select File");
            }
            else
            {
                string input = txtLink.Text;
                int index = input.LastIndexOf("\\");
                if (index > 0)
                {
                    input = input.Substring(0, index);
                }



                string fileName1 = "" + txtLink.Text + "";
                string filename = Path.GetFileName(fileName1);



                string source1Path = @"" + input + "";
                string targetPath = @"\\apbiphsh04\B1_BIPHCommon\15_IQC\00_Certificates IHCS (Do Not Delete)\" + txtWorkOrderNo.Text;
                //Console.WriteLine("The directory was created successfully at {0}.", Directory.GetCreationTime(path));
                // Use Path class to manipulate file and directory paths.
                string source1File = System.IO.Path.Combine(source1Path, filename);
                string destFile = System.IO.Path.Combine(targetPath, filename.Replace(".xlsx","").Replace("xls","") + " INHOUSE" + Path.GetExtension(fileName1));



                // To copy a folder's contents to a new location:
                // Create a new target folder.
                // If the directory already exists, this method does not create a new directory.
                System.IO.Directory.CreateDirectory(targetPath);

                cmd = new SqlCommand("update tbl_InhouseCalibrationMAIN set CertificateStats = 'UPLOADED', Temporary_Link = '" + targetPath +"\\"+filename.Replace(".xlsx", "").Replace("xls", "") + " INHOUSE" + Path.GetExtension(fileName1) + "', Status = 'Ongoing', Approver1 = '"+ cmb1stApprover.Text+ "', Approver2 = '" + cmbFinalApprover.Text + "' where WorkOrderNo = '" + txtWorkOrderNo.Text + "'", con);
                con.Open();
                cmd.ExecuteNonQuery();
                
                MessageBox.Show("Certificate Uploaded!", "Notification", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                con.Close();
                this.Close();

                // To copy a file to another location and
                // overwrite the destination file if it already exists.
                System.IO.File.Copy(source1File, destFile, true);
               // comboBox1.Text = "";
                txtLink.Text = "";



                dataGridView1.DataSource = null;
                dataGridView1.Rows.Clear();
              //  ReloadAlls();
            }
        }
    }
}
