using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace In_House_Calibration
{
    public partial class AdminPage : Form
    {
        public AdminPage()
        {
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AddAccount add = new AddAccount();
            add.ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AddCalStatus add2 = new AddCalStatus();
            add2.ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            AddInsCat add3 = new AddInsCat();
            add3.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            AddSection add4 = new AddSection();
            add4.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            AddCalPlan add5 = new AddCalPlan();
            add5.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            AddCalFreq add6 = new AddCalFreq();
            add6.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            AddCalCharge add7 = new AddCalCharge();
            add7.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            AddCertStats add8 = new AddCertStats();
            add8.ShowDialog();
        }
    }
}
