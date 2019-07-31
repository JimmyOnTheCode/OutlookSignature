using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.AccountManagement;

namespace OutlookSignature
{
    class LoggedInUser
    {

        public string getUsername()
        {
            return (string) Environment.UserName;
        }

        public string getFullname()
        {
            return (string) System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
        }

        public string getOrganization()
        {
            return (string)"Raiffeisen Bank Albania";
        }

    }
}
