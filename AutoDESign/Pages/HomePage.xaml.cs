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

namespace AutoDESign.Pages
{
    /// <summary>
    /// Interaction logic for HomePage.xaml
    /// </summary>
    public partial class HomePage : Page
    {
        public HomePage()
        {
            InitializeComponent();
        }

        private void btnAddSelfWeight_OnClick(object sender, RoutedEventArgs e)
        {
            JediMethods.AddSelfWeights();

        }

        private void btnSetInertiaChords_OnClick(object sender, RoutedEventArgs e)
        {
            var dictionary = JediMethods.GetAdditionalTakeoffInfo();
        }

        private void btnModifyJoists_OnClick(object sender, RoutedEventArgs e)
        {
            var selectModificationsPage = new AutoDESign.Pages.SelectModifications();
            this.NavigationService.Navigate(selectModificationsPage);
        }
    }
}
