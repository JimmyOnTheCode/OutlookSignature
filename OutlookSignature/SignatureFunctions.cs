using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;

namespace OutlookSignature
{
    class SignatureFunctions
    {
        private readonly string sourcePath = @"\\pathorip\e\RBAL Signature Template";
        public string destinationPath;
        private readonly string[] filesInScope = { "RBAL Signature.txt", "RBAL Signature.htm" };
        private string[] filePaths;    

        public void CopyFiles(string username)
        {
            destinationPath = @"C:\Users\" + username + @"\AppData\Roaming\Microsoft\Signatures";

            try
            {
                //Create all of the directories
                foreach (string dirPath in Directory.GetDirectories(sourcePath, "*", SearchOption.AllDirectories))
                {
                    Directory.CreateDirectory(dirPath.Replace(sourcePath, destinationPath));
                }

                //Copy all the files & Replaces any files with the same name
                foreach (string newPath in Directory.GetFiles(sourcePath, "*.*", SearchOption.AllDirectories))
                {
                    File.Copy(newPath, newPath.Replace(sourcePath, destinationPath), true);
                }
            }catch(Exception exc)
            {
                MessageBox.Show("Could not find Template!\n" + exc.Message);
                Application.Exit();
            }
        }

        public void GetCopiedFiles()
        {
            filePaths = Directory.GetFiles(destinationPath, "*", SearchOption.AllDirectories);
        }

        public void UpdateCopiedFiles(string fullname, string jobPosition, string department, string address, string tel, string mobile)
        {
            Dictionary<string, string> varsToReplace = new Dictionary<string, string>();

            //split AD Display Name - to handle name and surname separately also for email
            string name = fullname.Split(' ')[0]; ;
            string surname = fullname.Split(' ')[1];

            varsToReplace.Add("varName", name);
            varsToReplace.Add("varSurname", surname);
            varsToReplace.Add("varJobPosition", jobPosition);
            varsToReplace.Add("varDepartment", department);
            varsToReplace.Add("varAddress", address);

            //Only affect template if not empty
            if (tel != "")
            {
                varsToReplace.Add("varTel", tel);
            }

            if (mobile != "")
            {
                varsToReplace.Add("varMobile", mobile);
            }
            
            foreach (string filePath in filePaths)
            {
                if (filesInScope.Contains(Path.GetFileName(filePath)))
                {
                    string fileText = File.ReadAllText(filePath);

                    foreach(var item in varsToReplace)
                    {
                        if (fileText.Contains(item.Key))
                        {
                            fileText = fileText.Replace(item.Key, item.Value);
                        }
                    }

                    //if tel and mobile are empty, remove the entire signature line
                    if (tel == "")
                    {
                        fileText = fileText.Replace("Tel: varTel | ", "");
                    }

                    if (mobile == "")
                    {
                        fileText = fileText.Replace("Mob: varMobile", "");
                    }
                    
                    File.WriteAllText(filePath, fileText);
                }
            }
        }

    }
}
