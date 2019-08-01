using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices.AccountManagement;

namespace OutlookSignature
{
    class LoggedInUser
    {

        public string GetUsername()
        {
            return (string) Environment.UserName;
        }

        public string GetFullname()
        {
            return (string) System.DirectoryServices.AccountManagement.UserPrincipal.Current.DisplayName;
        }

        public string GetOrganization()
        {
            return (string)"Raiffeisen Bank Albania";
        }

    }
}
