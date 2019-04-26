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
using System.Windows.Shell;

namespace SalesBot.Pages
{
    /// <summary>
    /// Interaction logic for HomePage.xaml
    /// </summary>
    public partial class HomePage : Window
    {
        public HomePage()
        {
            InitializeComponent();
        }

        private void btnTakeoffConversion_OnClick(object sender, RoutedEventArgs e)
        {
            this.MainFrame.Source = new Uri("./ConvertTakeoffPage.xaml", UriKind.Relative);
        }

        private void btnJediModifications_OnClick(object sender, RoutedEventArgs e)
        {
            this.MainFrame.Source = new Uri("./SelectModifications.xaml", UriKind.Relative);
        }
    }
}
