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
            accessor = new ExcelAccessor(excelpath);
            changeADBasedOnExcel(accessor.Sheet);
            //+testTwo();
            MessageBox.Show("Done");
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
