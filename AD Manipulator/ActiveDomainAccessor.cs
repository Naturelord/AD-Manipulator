using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices;
using System.Windows;

namespace AD_Manipulator
{
    class ActiveDomainAccessor
    {
        string filter = "";
        string ADName = "";
        string adminUsername = "";
        string adminPassword = "";
        DirectoryEntry de = null;
        DirectorySearcher ds = null;


        /// <summary>
        /// Will Create a new Active Domain Accessor with all the objects in the domain included
        /// </summary>
        /// <param name="adName"></param>
        /// <param name="aduser"></param>
        /// <param name="adpass"></param>
        public ActiveDomainAccessor(string adName, string aduser, string adpass)
        {
            // Setting global variables for possible future use
            ADName = adName;
            adminUsername = aduser;
            adminPassword = adpass;
            // Open connection with domain
            try
            {
                de = new DirectoryEntry($"LDAP://{adName}", aduser, adpass);
                de.AuthenticationType = AuthenticationTypes.Secure;
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            // Search the domain based on no filter to retrieve all objects
            if (de != null)
            {
                DirectorySearcher ds = new DirectorySearcher(de);
                // Find all
                SearchResultCollection sr = ds.FindAll();
            }
        }
        /// <summary>
        /// Will Create a new Active Domain Accessor with only the filtered objects in the domain included.
        /// </summary>
        /// <param name="adName"></param>
        /// <param name="aduser"></param>
        /// <param name="adpass"></param>
        /// <param name="filter"></param>
        public ActiveDomainAccessor(string adName, string aduser, string adpass, string filter)
        {

        }

        /// <summary>
        /// Given a unique email address and an AD element - Delete the AD Element from the user's profile
        /// </summary>
        /// <param name="adElement"> AD Element you wish to delete</param>
        /// <param name="userEmail"> Email Address of User (unique email) </param>
        public void deleteElementFromUser(string adElement, string userEmail)
        {

        }

        public void findAll()
        {

        }

        /// <summary>
        /// Returns the correct form of the number given (without +1), if the number is already correct it returns the given number
        /// </summary>
        /// <param name="number"></param>
        /// <returns></returns>
        public string correctPhoneNumberNoPlusOne( string number)
        {
            string ret = number;
            if(number != null)
            {
                ret = ret.Trim();
                string check = ret.Substring(0, 2);
                if( check.CompareTo("+1") == 0)
                {
                    ret = ret.Substring(2);
                }
            }
            return ret;
        }


    }// End Class
}// End Namespace
