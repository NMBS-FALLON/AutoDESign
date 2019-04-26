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

namespace SalesBot.Pages
{
    /// <summary>
    /// Interaction logic for ConvertTakeoffPage.xaml
    /// </summary>
    public partial class ConvertTakeoffPage : Page
    {
        public ConvertTakeoffPage()
        {
            InitializeComponent();
        }

        private void btnConvertSeiTakeoff_Click(object sender, RoutedEventArgs e)
        {
            Conversions.ConvertSeiTakeoff();
        }
    }
}
