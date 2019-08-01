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
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            StartPosition = FormStartPosition.CenterScreen;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoggedInUser loggedInUser = new LoggedInUser();
            this.LabelFullnameData.Text = loggedInUser.GetFullname(); ;
            this.LabelUsernameData.Text = loggedInUser.GetUsername();
            this.LabelOrgData.Text = loggedInUser.GetOrganization();

            ActiveDirectoryFunctions activeDirectoryFunctions = new ActiveDirectoryFunctions();
            Dictionary<string, string> userData = activeDirectoryFunctions.LoadUserProperties(this.LabelUsernameData.Text);
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
                try
                {
                    SignatureFunctions signatureFunctions = new SignatureFunctions();
                    
                    signatureFunctions.CopyFiles(this.LabelUsernameData.Text);
                    signatureFunctions.GetCopiedFiles();
                    signatureFunctions.UpdateCopiedFiles(this.LabelFullnameData.Text, this.TextboxUnitData.Text, this.TextboxDepartmentData.Text, this.TextboxDivisionData.Text, this.TextboxAddressData.Text, this.TextboxTelData.Text);
                    ButtonGenerate.Enabled = false;
                    MessageBox.Show("Success!");
                    Application.Exit();
                }
                catch(Exception exc)
                {
                    MessageBox.Show("Error! " + exc.ToString());
                }
            }
        }
        private void LabelOrgData_Click(object sender, EventArgs e)
        {

        }
    }
}