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
            ButtonGenerate.Select();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            LoggedInUser loggedInUser = new LoggedInUser();
            this.LabelFullnameData.Text = loggedInUser.GetFullname(); ;
            this.LabelUsernameData.Text = loggedInUser.GetUsername();
            

            ActiveDirectoryFunctions activeDirectoryFunctions = new ActiveDirectoryFunctions();
            Dictionary<string, string> userData = activeDirectoryFunctions.LoadUserProperties(this.LabelUsernameData.Text);
            this.TextboxJobPosition.Text = userData["title"];
            this.TextboxDepartment.Text = userData["department"];
            this.TextboxAddress.Text = userData["streetAddress"];
            this.TextboxTelephone.Text = userData["telephoneNumber"];
            this.TextboxMobile.Text = userData["mobile"];
            this.LabelOrgData.Text = userData["company"];
        }

        private bool ValidateInputs()
        {
            string[] mandatoryTextboxes = { "TextboxJobPosition", "TextboxDepartment", "TextboxAddress", "TextboxTelephone" };
            bool condition = true;

            foreach(Control controlElement in this.Controls)
            {
                if (controlElement is TextBox)
                {
                    if (controlElement.Text == "")
                    {
                        if (mandatoryTextboxes.Contains(controlElement.Name))
                        {
                            condition = false;
                        }
                    }
                }
            }

            return condition;
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {
            if (!ValidateInputs())
            {
                MessageBox.Show("Please fill all mandatory fields: \n\u2022 Job Position\n\u2022 Branch/Dept.\n\u2022 Address\n\u2022 Telephone ");
            }
            else
            {
                try
                {
                    SignatureFunctions signatureFunctions = new SignatureFunctions();
                    
                    signatureFunctions.CopyFiles(this.LabelUsernameData.Text);
                    signatureFunctions.GetCopiedFiles();
                    signatureFunctions.UpdateCopiedFiles(
                        this.LabelFullnameData.Text,
                        this.TextboxJobPosition.Text,
                        this.TextboxDepartment.Text,
                        this.TextboxAddress.Text,
                        this.TextboxTelephone.Text,
                        this.TextboxMobile.Text
                    );
                    ButtonGenerate.Enabled = false;
                    MessageBox.Show("Success!\nYou can now select the new signature from Outlook.");
                    Application.Exit();
                }
                catch(Exception exc)
                {
                    MessageBox.Show("Error! " + exc.ToString());
                }
            }
        }

    }
}