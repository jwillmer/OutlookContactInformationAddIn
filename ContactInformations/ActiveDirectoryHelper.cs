using System;
using System.DirectoryServices;
using System.Threading;

namespace ContactInformations
{
    public class ActiveDirectoryHelper
    {
        private SearchResult searchResult = null;

        public bool CanConnectToActiveDirectory { get; private set; }

        public ActiveDirectoryHelper(string[] propertiesToLoad)
        {
            var thread = new Thread(delegate()
                                    {
                                        if (InitSearchResult(propertiesToLoad)) 
                                            CanConnectToActiveDirectory = true;
                                    });
            thread.Start();
        }

        private bool InitSearchResult(string[] propertiesToLoad)
        {
            using (var entry = new DirectoryEntry())
            {
                using (var adSearcher = new DirectorySearcher(entry))
                {
                    adSearcher.Filter = string.Format("(sAMAccountName={0})", Environment.UserName);
                    adSearcher.PropertiesToLoad.Add("allowedAttributesEffective");  //all attributes with write access

                    foreach (var property in propertiesToLoad)
                    {
                        adSearcher.PropertiesToLoad.Add(property);
                    }                  

                    try
                    {
                        searchResult = adSearcher.FindOne();
                    }
                    catch (Exception)
                    {
                        return false;
                    }
                    return true;
                }
            }
        }

        public bool CheckPropertyWritePermission(string property)
        {
            if (searchResult != null && searchResult.Properties.Contains("allowedAttributesEffective"))
                return searchResult.Properties["allowedAttributesEffective"].Contains(property);
            return false;
        }

        public bool SetProperty(string propertyName, object propertyValue)
        {
            if (propertyName == null || propertyValue == null || searchResult == null) return false;

            var de = searchResult.GetDirectoryEntry();
            if (de.Properties.Contains(propertyName))
                de.Properties[propertyName].Value = propertyValue;
            else
                de.Properties[propertyName].Add(propertyValue);

            try
            {
                de.CommitChanges();
                return true;
            }
            catch
            {
                return false;
            }

        }

        public object GetValue(string propertyName)
        {
            if (searchResult == null || propertyName == null || !searchResult.Properties.Contains(propertyName)) return null;
            return searchResult.Properties[propertyName][0];
        }
    }
}
