using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.DirectoryServices;
using System.DirectoryServices.ActiveDirectory;
using System.DirectoryServices.AccountManagement;
using Excel = Microsoft.Office.Interop.Excel;
using Window = System.Windows.Window;
using AD_Manipulator.ActiveDirectoryHelper;
namespace AD_Manipulator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string domain = "murraycompany.local";
        private string password = "";
        private string domainName = Environment.UserDomainName;
        private DirectorySearcher dirSearch = null;
        private string excelpath = @"C:\Egnyte\Private\bpatton\Data\murray_users.xls";
        private ExcelAccessor accessor;



        public MainWindow()
        {
            InitializeComponent();
            //accessor = new ExcelAccessor(excelpath);
            //changeADBasedOnExcel(accessor.Sheet);
            //+testTwo();
            removeOnesFromPhoneNumbersInAD();
            MessageBox.Show("Done");
        }
        
        // New Script to remove All +1's from everyones phone numbers
        private void removeOnesFromPhoneNumbersInAD()
        {
            string murray = "murraycompany.local";
            // Create the Object of our directoy
            DirectoryEntry de = new DirectoryEntry($"LDAP://{murray}", "bpatton", "Meganega13");
            de.AuthenticationType = AuthenticationTypes.Secure;
            // find the Directory
            DirectorySearcher ds = new DirectorySearcher(de);
            ds.Filter= "(&(objectcategory=user))";
            
            // Officially assign that directory
            SearchResultCollection sr = null;
            try
            {
                sr = ds.FindAll();
            }
            catch(Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            

            int i = 0;
            string add = "";
            foreach ( SearchResult userResult in sr)
            {
                
                de = userResult.GetDirectoryEntry();



                string fax = Convert.ToString(de.Properties[ADProperties.FAX].Value);
                if (!String.IsNullOrEmpty(fax))
                {
                    fax = correctPhoneNumberNoPlusOne(fax);
                    de.Properties[ADProperties.FAX].Value = fax;
                    de.CommitChanges();
                }



                //string show = Convert.ToString(de.Properties[ADProperties.TELEPHONE].Value);
                //show = show + " " + Convert.ToString(de.Properties[ADProperties.FAX].Value);
                //show = show + " " + Convert.ToString(de.Properties[ADProperties.NAME].Value);
                //add = add + show + "\n";
                //if (i == 24)
                //{
                //        MessageBox.Show(add);
                //        i = 0;
                //        add = "";
                //}
            de.CommitChanges();
                i++; 
            }
            

        }

        private void changeADBasedOnExcel(Excel.Worksheet Sheet)
        {
            for (int i = 2; i < 360; i++)
            {
                string mail = (string) Sheet.Cells[1][i].Value2;
                string mobile = (string)Sheet.Cells[3][i].Value2;
                string rcNumber = (string)Sheet.Cells[2][i].Value2;

                try
                {
                    // Make Authentication Type Secure to grab data?? Posibly because of group policy
                    DirectoryEntry de = new DirectoryEntry("LDAP://murraycompany.local","bpatton","Meganega13");
                    de.AuthenticationType = AuthenticationTypes.Secure;
                    DirectorySearcher search = new DirectorySearcher(de);


                    // look for the object based on mail property.
                    search.Filter= "(mail=" + mail + ")";
                    
                    
                    // Find the 1 corresponding email
                    SearchResult result = search.FindOne();
                    // Create a new de based on result
                    if (result != null)
                    {
                        de = result.GetDirectoryEntry();

                        string telephone = Convert.ToString(de.Properties[ADProperties.TELEPHONE].Value);



                        // NOW THAT setup is complete we need to remove the Extension

                        if (telephone.Length > 14)
                        {
                            telephone = telephone.Substring(0, 14);
                            if (!String.IsNullOrEmpty(telephone))
                            {
                                de.Properties[ADProperties.TELEPHONE].Value = telephone;
                                de.CommitChanges();

                            }
                        }
                        // Set Mobile equal to mobile
                        if (!String.IsNullOrEmpty(mobile))
                        {
                            de.Properties[ADProperties.MOBILE].Value = mobile;
                            de.CommitChanges();
                        }

                        // Fax = RC Number
                        if (!String.IsNullOrEmpty(rcNumber))
                        {
                            de.Properties[ADProperties.FAX].Value = rcNumber;
                            de.CommitChanges();
                        }
                        // Check Result
                        //string one = Convert.ToString(de.Properties[ADProperties.MOBILE].Value);
                        //string two = Convert.ToString(de.Properties[ADProperties.FAX].Value);
                        //string three = Convert.ToString(de.Properties[ADProperties.TELEPHONE].Value);
                        //MessageBox.Show("mail: " + mail + "\nMobile: " + one + "\nFax Number: " + two + "\nTelephone Number: " + three);
                    }// If there is no result do nothing

                }
                catch (Exception e)
                {
                    MessageBox.Show(e.ToString());
                }
            }// End For loop
            
        }

        private DirectorySearcher GetDirectorySearcher()
        {
            dirSearch = new DirectorySearcher( new DirectoryEntry("LDAP://murraycompany.local") );
            if(dirSearch != null)
            {
                return dirSearch;
            }
            return null;
        }

        private SearchResult SearchByEmail(DirectorySearcher ds, string email)
        {
            ds.Filter = "(&((&(objectCategory=Person)(objectClass=User)))(mail="+email+"))";


            ds.SearchScope = SearchScope.Subtree;
            ds.ServerTimeLimit = TimeSpan.FromSeconds(90);

            SearchResult userObject = ds.FindOne();

            if(userObject != null)
            {
                return userObject;
            }
            return null;
        }

        public string correctPhoneNumberNoPlusOne(string number)
        {
            string ret = number;
            if (number != null)
            {
                ret = ret.Trim();
                string check = ret.Substring(0, 2);
                if (check.CompareTo("+1") == 0)
                {
                    ret = ret.Substring(2);
                }
            }
            return ret;
        }

        /// <summary>
        /// Dead Code Must have to not pull Errors... further research into this issue will be looked later.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Submit_Click(object sender, RoutedEventArgs e)
        {

        }

    }// End Class

    public class User
    {
        public string email { get; set; }
        public string name { get; set; }
        public string phoneNumber { get; set; }
        public string rcNumber { get; set; }

    }

    
}// End Namespace
