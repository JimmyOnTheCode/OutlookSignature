using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OutlookSignature
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoggedInUser loggedInUser = new LoggedInUser();
            this.LabelFullnameData.Text = loggedInUser.getFullname();
            this.LabelUsernameData.Text = loggedInUser.getUsername();
            this.LabelOrgData.Text = loggedInUser.getOrganization();
        }

        private bool ValidateInputs()
        {
            bool condition = false;
            int countFilled = 0;
            
            if(this.TextboxUnitData.Text != "")
            {
                countFilled++;
            }

            if (this.TextboxDepartmentData.Text != "")
            {
                countFilled++;
            }

            if (this.TextboxDivisionData.Text != "")
            {
                countFilled++;
            }

            if (countFilled >= 2)
            {
                condition = true;
            }

            return condition;
        }
        private void LabelFullnameData_Click(object sender, EventArgs e)
        {

        }

        private void Label1_Click(object sender, EventArgs e)
        {

        }

        private void Label2_Click(object sender, EventArgs e)
        {

        }

        private void LabelUsernameData_Click(object sender, EventArgs e)
        {

        }

        private void Label6_Click(object sender, EventArgs e)
        {

        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {

            if (!ValidateInputs())
            {
                MessageBox.Show("Please fill at least 2/3 fields above!");
            }
            else
            {
                ButtonGenerate.Enabled = false;
                MessageBox.Show("success");
            }

        }

        private void LabelOrgData_Click(object sender, EventArgs e)
        {

        }
    }
}
