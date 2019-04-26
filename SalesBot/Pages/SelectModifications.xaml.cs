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
    /// Interaction logic for SelectModifications.xaml
    /// </summary>
    public partial class SelectModifications : Page
    {
        public SelectModifications()
        {
            InitializeComponent();
            btnApplyModifications.Width = modificationList.Width;
        }

        public void btnApplyModifications_OnClick(object sender, RoutedEventArgs e)
        {
            var modifications =
                (
                    AddSelfWeight: (bool)(AddSelfWeightCb.IsChecked),
                    SetChordsForInertia: (bool)(SetChordsForInertiaCb.IsChecked),
                    ApplyAdditionalTakeoffInfo: (bool)(ApplyAdditionalTakeoffInfoCb.IsChecked)
                );

            ApplyModifications(modifications);

        }

        public static void ApplyModifications((bool AddSelfWeight, bool SetChordsForInertia, bool ApplyAdditionalTakeoffInfo) modifications)
        {
            
        }

    }
}
