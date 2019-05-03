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
        private DirectorySearcher ds = null;
        private string excelpath = @"C:\Egnyte\Private\bpatton\Data\murray_users.xls";

       

        public MainWindow()
        {
            InitializeComponent();
            ExcelAccessor accessor = new ExcelAccessor(excelpath);
            GetADUsers();
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


        // Utilizes the System.
        private void GetADUsers()
        {
            using (var context = new PrincipalContext(ContextType.Domain, "murraycompany.local"))
            {
                using (var searcher = new PrincipalSearcher(new UserPrincipal(context)))
                {
                    foreach (var result in searcher.FindAll())
                    {
                        DirectoryEntry de = result.GetUnderlyingObject() as DirectoryEntry;
                        var clsuser = new User();
                        MessageBox.Show(Convert.ToString(de.Properties["samAccountName"].Value) + " , "+ Convert.ToString(de.Properties["givenName"].Value
                             + " , " + Convert.ToString(de.Properties["mail"].Value) ));
                    }
                    
                }
            }
        }
    }

    public class User
    {
        public string email { get; set; }
        public string name { get; set; }
        public string phoneNumber { get; set; }
    }
}
