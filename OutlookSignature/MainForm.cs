using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;

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

        private void MainForm_Load(object sender, EventArgs e)
        {
            LoggedInUser loggedInUser = new LoggedInUser();
            this.LabelFullnameData.Text = loggedInUser.GetFullname();
            this.LabelUsernameData.Text = loggedInUser.GetUsername();

            ActiveDirectoryFunctions activeDirectoryFunctions = new ActiveDirectoryFunctions();
            Dictionary<string, string> userData = activeDirectoryFunctions.LoadUserProperties(this.LabelUsernameData.Text);
            this.TextboxJobPosition.Text = userData["title"];
            this.TextboxDepartment.Text = userData["department"]; 
            this.TextboxAddress.Text = userData["streetAddress"];
            this.TextboxTelephone.Text = userData["telephoneNumber"];
            this.TextboxMobile.Text = userData["mobile"];
            this.LabelOrgData.Text = userData["company"];
            if (userData["departmentNumber"] != "")
            {
                this.TextboxDepartment.Text = userData["department"] + ", " + userData["departmentNumber"];
            }
 
        }

        private bool ValidateInputs()
        {
            string[] mandatoryTextboxes = { "TextboxJobPosition","TextboxDepartment", "TextboxAddress" };
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
        
        private void UpdateRegistry(string value)
        {
            try
            {
                using (RegistryKey key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\16.0\Outlook\Profiles\Outlook\9375CFF0413111d3B88A00104B2A6676\00000002", true))
                {
                    if(key != null)
                    {
                        key.SetValue("New Signature", value);
                        key.SetValue("Reply-Forward Signature", value);
                    }
                }
            }catch(Exception exc)
            {
                return;
            }
        }

        private void ButtonGenerate_Click(object sender, EventArgs e)
        {   
            //validate inputs, skip if certain user logged in...
            if (!ValidateInputs() && this.LabelUsernameData.Text != "uniqueusername")
            {
                MessageBox.Show("Please fill all mandatory fields: \n\u2022 Job Position\n\u2022 Branch/Dept.\n\u2022 Address");
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
                    UpdateRegistry("RBAL Signature");
                    MessageBox.Show("Success!\nSimply restart Outlook and enjoy your new signature.");
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
