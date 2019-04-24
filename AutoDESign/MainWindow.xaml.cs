﻿using System;
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
using AutoIt;
using System.Drawing;
using OfficeOpenXml;

namespace AutoDESign
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnTest_OnClick(object sender, RoutedEventArgs e)
        {
            JediMethods.AddSelfWeights();

        }
    }


    public class JediMethods
    {
        static public void GoToJoistList()
        {
            AutoIt.AutoItX.AutoItSetOption("WinTitleMatchMode", 2);
            AutoIt.AutoItX.WinActivate("Joist Design");
            AutoIt.AutoItX.WinWaitActive("Joist Design");
            AutoIt.AutoItX.ControlClick("Joist Design", "", "TListBox1", "left", 1, 56, 84);
            AutoIt.AutoItX.Sleep(500);
            AutoIt.AutoItX.ControlClick("Joist Design", "", "TDBGridExt1", "left", 1, 20, 30);
            AutoIt.AutoItX.Sleep(500);
            AutoIt.AutoItX.Send("{HOME}");
            AutoIt.AutoItX.Sleep(1000);
        }

        static public void ApplyToJoists(Action func)
        {
            GoToJoistList();
            var previousMark = "";
            var isFinalMark = false;
            GetSelfWeights();
            do
            {
                AutoIt.AutoItX.Send("{ENTER}");
                AutoIt.AutoItX.WinWaitActive("Joist Properties");
                var currentMark = AutoIt.AutoItX.ControlGetText("Joist Properties", "", "TDBEdit20");
                if (currentMark != previousMark)
                {
                    func();
                }
                else
                {
                    isFinalMark = true;
                }
                previousMark = currentMark;
                AutoIt.AutoItX.WinClose("Joist Properties");
                AutoIt.AutoItX.WinActivate("Joist Design");
                AutoIt.AutoItX.WinWaitActive("Joist Design");
                AutoIt.AutoItX.Send("{DOWN}");
            } while (!isFinalMark);
        }

        static public void AddSelfWeights()
        {
            var selfWeightDictionary = GetSelfWeights();
            GoToJoistList();
            var previousMark = "";
            var isFinalMark = false;
            do
            {
                AutoIt.AutoItX.Send("{ENTER}");
                AutoIt.AutoItX.WinWaitActive("Joist Properties");
                var currentMark = AutoIt.AutoItX.ControlGetText("Joist Properties", "", "TDBEdit20");
                if (currentMark != previousMark)
                {
                    if (selfWeightDictionary.ContainsKey(currentMark))
                    {
                        var selfWeight = selfWeightDictionary[currentMark];
                        AutoIt.AutoItX.ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 154, 13);
                        AutoIt.AutoItX.ControlFocus("Joist Properties", "", "TStringGrid1");
                        AutoIt.AutoItX.Send("{INSERT}");
                        AutoIt.AutoItX.WinWaitActive("New Load");
                        AutoIt.AutoItX.Send("1");
                        AutoIt.AutoItX.Send("{TAB}");
                        AutoIt.AutoItX.Send("2");
                        AutoIt.AutoItX.Send("{TAB}");
                        AutoItX.Send(selfWeight.ToString());
                        AutoItX.Sleep(200);
                        AutoItX.Send("{TAB}");
                        AutoItX.Send("{TAB}");
                        AutoItX.ControlFocus("New Load", "", "TBitBtn1");
                        AutoItX.ControlClick("New Load", "", "TBitBtn1", "left", 1, 40, 13);
                        AutoItX.WinWaitActive("Joist Properties");
                        AutoItX.ControlClick("Joist Properties", "", "TBitBtn2", "left", 1, 46, 11);
                        AutoItX.WinWaitClose("Joist Properties");
                        AutoItX.Sleep(500);
                    }
                }
                else
                {
                    isFinalMark = true;
                    AutoItX.WinClose("Joist Properties");
                }
                previousMark = currentMark;
                AutoIt.AutoItX.WinWaitActive("Joist Design");
                AutoIt.AutoItX.Send("{DOWN}");
            } while (!isFinalMark);

        }

        static public Dictionary<string,double> GetSelfWeights()
        {
            var selfWeightDictionary = new Dictionary<string, double>();
            var openFile = new Microsoft.Win32.OpenFileDialog();
            openFile.Filter = "Excel Documents|*.xlsx;*.xlsm";

            var result = openFile.ShowDialog();
            if (result == true)
            {
                var fileName = openFile.FileName;
                using (var stream = new System.IO.FileStream(fileName, System.IO.FileMode.Open))
                {
                    using (var package = new ExcelPackage(stream))
                    {
                        var selfWeightWorksheet = package.Workbook.Worksheets["Self Weight"];
                        var lastRow = selfWeightWorksheet.Dimension.End.Row;
                        for (int i = 1; i <= lastRow; i++)
                        {
                            var mark = selfWeightWorksheet.GetValue<string>(i + 1, 5);
                            if (mark != null && mark != "")
                            {
                                var selfWeight = selfWeightWorksheet.GetValue<double>(i + 1, 6);
                                selfWeightDictionary.Add(mark, selfWeight);
                            }
                        }
                       
                    }
                }
            }
            return selfWeightDictionary;
        }

    }
}
