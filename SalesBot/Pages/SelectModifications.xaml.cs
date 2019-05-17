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
using LanguageExt;
using static LanguageExt.Prelude;
using OfficeOpenXml;
using AutoIt;
using System.Collections.ObjectModel;

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

            try
            {
                JediMethods.ApplyModifications(modifications);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }

        }

        private void btnTest_OnClick(object sender, RoutedEventArgs e)
        {
            var fileName = General.GetFileName("Select Excel File", "Excel Documents | *.xlsx; *.xlsm");
            var package = match(fileName, Some: s => General.GetExcelPackage(s), None: () => null);
            JediMethods.PrintTables(package);
        }

        private void btnDrawAllModifications_OnClick(object sender, RoutedEventArgs e)
        {
            JediMethods.DrawAllProfiles();
        }
    }
}
