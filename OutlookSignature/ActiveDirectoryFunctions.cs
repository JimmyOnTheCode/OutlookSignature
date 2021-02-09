using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.DirectoryServices;

namespace OutlookSignature
{
    class ActiveDirectoryFunctions
    {
        public string GetPropertyValue(SearchResult searchResult, string PropertyName)
        {
            if (searchResult.Properties.Contains(PropertyName))
            {
                return searchResult.Properties[PropertyName][0].ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        public Dictionary<string, string> LoadUserProperties(string username)
        {
            DirectoryEntry entry = new DirectoryEntry("LDAP://addomainnamegoeshere");
            DirectorySearcher adSearch = new DirectorySearcher(entry);
            adSearch.Filter = $"(sAMAccountName={username})";
            string[] propertiesToLoad = { "telephoneNumber", "department", "departmentNumber", "mobile", "title", "streetAddress", "company" };

            foreach (string property in propertiesToLoad)
            {
                adSearch.PropertiesToLoad.Add(property);
            }

            SearchResult adSearchResult = adSearch.FindOne();
            Dictionary<string, string> loadedProperties = new Dictionary<string, string>();

            foreach (string property in propertiesToLoad)
            {
                loadedProperties.Add(property, GetPropertyValue(adSearchResult, property));
            }
            return loadedProperties;
        }

    }
}
