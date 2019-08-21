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
using System.Drawing;
using OfficeOpenXml;
using LanguageExt;
using static LanguageExt.Prelude;
using AutoIt;
using static AutoIt.AutoItX;
using OfficeOpenXml.Table;
using System.Linq.Expressions;

namespace SalesBot
{
    public class JediMethods
    {

        static public void PrintTables(ExcelPackage package)
        {
            Predicate<ExcelTable> tableContainsMarkColumn =
                tbl =>
                    tbl.Columns
                    .Where(col => col.Name.ToUpper() == "MARK")
                    .Any();

            var tables = TryGetTablesWhere(package, tableContainsMarkColumn);
            tables.Iter(tbl => MessageBox.Show(tbl.Name));
        }

        static public void PrintTableColumnNames(ExcelPackage package, Func<ExcelTable, bool> predicate)
        {
            var tables =
                package.Workbook.Worksheets.AsEnumerable()
                .SelectMany(sheet => sheet.Tables);
            var table =
                tables
                .Where(tbl => predicate(tbl))
                .First();
            var columnNames =
                table.Columns.Select(col => col.Name);
            columnNames.Iter(colName => MessageBox.Show(colName));

        }

        static public IEnumerable<ExcelTable> TryGetTablesWhere(ExcelPackage package, Predicate<ExcelTable> whereFunction)
        {
            var tables =
               package.Workbook.Worksheets
               .AsEnumerable()
               .SelectMany(sheet => sheet.Tables.Where(tbl => whereFunction(tbl)));
            return tables;
        }

        static public Option<ExcelTable> TryGetTableWhere(ExcelPackage package, Predicate<ExcelTable> whereFunction)
        {
            var possibleTable =
               package.Workbook.Worksheets
               .AsEnumerable()
               .SelectMany(sheet => sheet.Tables.Where(tbl => whereFunction(tbl)));
            var tableOption =
                Enumerable.Any(possibleTable) ?
                Some(possibleTable.First()) :
                Option<ExcelTable>.None;
            return tableOption;
        }


        static public void GoToJoistList()
        {
            AutoItSetOption("WinTitleMatchMode", 2);
            WinActivate("Joist Design");
            WinWaitActive("Joist Design");
            ControlClick("Joist Design", "", "TListBox1", "left", 1, 56, 84);
            Sleep(500);
            ControlClick("Joist Design", "", "TDBGridExt1", "left", 1, 20, 30);
            Sleep(500);
            Send("{HOME}");
            Sleep(1000);
        }

        static public void ApplyToJoists<ArgType>(Action<ArgType> func, ArgType arg)
        {
            GoToJoistList();
            var previousMark = "";
            var isFinalMark = false;
            do
            {
                Send("{ENTER}");
                WinWaitActive("Joist Properties");
                var currentMark = ControlGetText("Joist Properties", "", "TDBEdit20");
                if (currentMark != previousMark)
                {
                    func(arg);
                }
                else
                {
                    isFinalMark = true;
                }
                previousMark = currentMark;
                WinClose("Joist Properties");
                WinActivate("Joist Design");
                WinWaitActive("Joist Design");
                Send("{DOWN}");
            } while (!isFinalMark);
        }

        static public void AddSelfWeight(double selfWeight, out bool changeHappend)
        {
            changeHappend = false;
            if (WinExists("Joist Properties") == 1)
            {
                changeHappend = true;
                WinActivate("Joist Properties");
                ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 154, 13);
                ControlFocus("Joist Properties", "", "TStringGrid1");
                Send("{INSERT}");
                WinWaitActive("New Load");
                Send("1");
                Send("{TAB}");
                Send("2");
                Send("{TAB}");
                Send(selfWeight.ToString());
                Sleep(200);
                Send("{TAB}");
                Send("{TAB}");
                ControlFocus("New Load", "", "TBitBtn1");
                ControlClick("New Load", "", "TBitBtn1", "left", 1, 40, 13);
                WinWaitActive("Joist Properties");
            }
            else
            {
                failwith("'AddSelfWeight(double selfWeight)' must be called with the 'Joist Properties' window open");
            }
        }

        static public Dictionary<string, double> GetSelfWeights(ExcelPackage package)
        {
            var selfWeightDictionary = new Dictionary<string, double>();

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

            return selfWeightDictionary;
        }

        static public Dictionary<string, (bool TopChordChanged, bool BottomChordChanged, string TopChordSize, string BottomChordSize)> GetChordsFromInertiaCheck(ExcelPackage package)
        {
            var chordDictionary = new Dictionary<string, (bool, bool, string, string)>();
            var inertiaCheckWorksheet = package.Workbook.Worksheets["Inertia Check"];
            var lastRow = inertiaCheckWorksheet.Dimension.End.Row;
            for (int i = 1; i <= lastRow; i++)
            {
                var mark = inertiaCheckWorksheet.GetValue<string>(i + 1, 5);
                if (mark != null && mark != "")
                {
                    var topChordChanged = inertiaCheckWorksheet.GetValue<bool>(i + 1, 35);
                    var bottomChordChanged = inertiaCheckWorksheet.GetValue<bool>(i + 1, 36);
                    var topChordSize = inertiaCheckWorksheet.GetValue<string>(i + 1, 37);
                    var bottomChordSize = inertiaCheckWorksheet.GetValue<string>(i + 1, 38);
                    if (!chordDictionary.ContainsKey(mark))
                    {
                        chordDictionary.Add(mark, (topChordChanged, bottomChordChanged, topChordSize, bottomChordSize));
                    }
                }
            }
            return chordDictionary;
        }

        public struct AdditionalTakeoffInfo
        {
            public double Mf { get; }
            public double IMin { get; }
            public double TlDeflection { get; }
            public double LlDeflection { get; }
            public bool ErfoAtLe { get; }
            public bool ErfoAtRe { get; }
            public Option<double> WnSpacing { get; }

            public AdditionalTakeoffInfo(double mf, double iMin, double tlDeflection, double llDeflection, bool erfoAtLe, bool erfoAtRe, Option<double> wnSpacing)
            {
                Mf = mf;
                IMin = iMin;
                TlDeflection = tlDeflection;
                LlDeflection = llDeflection;
                ErfoAtLe = erfoAtLe;
                ErfoAtRe = erfoAtRe;
                WnSpacing = wnSpacing;
            }
        }

        static public Dictionary<string, AdditionalTakeoffInfo> GetAdditionalTakeoffInfo(ExcelPackage package)
        {
            var additionalTakeoffInfoDictionary = new Dictionary<string, AdditionalTakeoffInfo>();

            var additionalTakeoffInfoSheet = package.Workbook.Worksheets["Additional Takeoff Info"];
            var lastRow = additionalTakeoffInfoSheet.Dimension.End.Row;
            for (int i = 1; i <= lastRow; i++)
            {
                var mark = additionalTakeoffInfoSheet.GetValue<string>(i + 1, 1);
                if (mark != null && mark != "")
                {
                    if (!additionalTakeoffInfoDictionary.ContainsKey(mark))
                    {
                        var mf = additionalTakeoffInfoSheet.GetValue<double>(i + 1, 2);
                        var iMin = additionalTakeoffInfoSheet.GetValue<double>(i + 1, 3);
                        var tlDeflection = additionalTakeoffInfoSheet.GetValue<double>(i + 1, 4);
                        var llDeflection = additionalTakeoffInfoSheet.GetValue<double>(i + 1, 5);
                        var erfoAtLe = additionalTakeoffInfoSheet.GetValue<bool>(i + 1, 6);
                        var erfoAtRe = additionalTakeoffInfoSheet.GetValue<bool>(i + 1, 7);
                        var wnSpacing =
                            additionalTakeoffInfoSheet.GetValue<double>(i + 1, 8) == 0 ?
                            Option<double>.None :
                            Option<double>.Some(additionalTakeoffInfoSheet.GetValue<double>(i + 1, 8));
                        var additionalTakeoffInfo = new AdditionalTakeoffInfo(mf, iMin, tlDeflection, llDeflection, erfoAtLe, erfoAtRe, wnSpacing);
                        additionalTakeoffInfoDictionary.Add(mark, additionalTakeoffInfo);
                    }
                }
            }

            return additionalTakeoffInfoDictionary;
        }

        public static void SetDeflection(double tlDeflection, double llDeflection, out bool changeHappened)
        {
            changeHappened = false;
            if (WinExists("Joist Properties") == 1)
            {
                changeHappened = true;
                WinActivate("Joist Properties");
                ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 27, 13);
                Sleep(100);
                ControlSetText("Joist Properties", "", "TDBEdit1", tlDeflection.ToString());
                Sleep(100);
                ControlSetText("Joist Properties", "", "TDBEdit2", llDeflection.ToString());
                Sleep(100);
                WinWaitActive("Joist Properties");
            }
            else
            {
                failwith("'AddSelfWeight(double selfWeight)' must be called with the 'Joist Properties' window open");
            }
        }

        public static void SetWoodnailer(double screwSpacing, out bool changeHappened)
        {
            changeHappened = false;
            if (WinExists("Joist Properties") == 1)
            {
                WinActivate("Joist Properties");
                if (screwSpacing != 0.0)
                {
                    changeHappened = true;
                    ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 223, 13);
                    Sleep(100);
                    ControlCommand("Joist Properties", "", "TDBCheckBox3", "Check", "");
                    Sleep(100);
                    ControlFocus("Joist Properties", "", "Edit1");
                    Sleep(100);
                    Send(Math.Floor(screwSpacing).ToString());
                    Sleep(200);
                    Send("{TAB}");
                    WinWaitActive("Joist Properties");
                }
            }
            else
            {
                failwith("'AddSelfWeight(double selfWeight)' must be called with the 'Joist Properties' window open");
            }
        }

        public static void SetErfos(bool hasErfoAtLe, bool hasErfoAtRe, out bool changeHappened)
        {
            changeHappened = false;
            if (WinExists("Joist Properties") == 1)
            {
                changeHappened = true;
                WinActivate("Joist Properties");
                ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 275, 13);
                Sleep(200);
                if (hasErfoAtLe)
                {
                    ControlCommand("Joist Properties", "", "TDBCheckBox4", "Check", "");
                }
                else
                {
                    ControlCommand("Joist Properties", "", "TDBCheckBox4", "UnCheck", "");
                }

                if (hasErfoAtRe)
                {
                    ControlCommand("Joist Properties", "", "TDBCheckBox1", "Check", "");
                }
                else
                {
                    ControlCommand("Joist Properties", "", "TDBCheckBox1", "UnCheck", "");
                }

                Sleep(100);
                WinWaitActive("Joist Properties");
            }
            else
            {
                failwith("'AddSelfWeight(double selfWeight)' must be called with the 'Joist Properties' window open");
            }
        }

        public static void SetChords(Option<string> topChordSize, Option<string> bottomChordSize, out bool changeHappened)
        {
            changeHappened = false;
            if (WinExists("Joist Properties") == 1)
            {
                if (topChordSize.IsSome || bottomChordSize.IsSome)
                {
                    changeHappened = true;
                    WinActivate("Joist Properties");
                    ControlClick("Joist Properties", "", "TPageControl1", "left", 1, 241, 10);
                    if (topChordSize.IsSome)
                    {
                        Sleep(100);
                        ControlClick("Joist Properties", "", "TComboBox2", "left", 1, 81, 10);
                        Sleep(200);
                        match(topChordSize, Some: s => Send(s.ToString()), None: () => { });
                        Sleep(200);
                        Send("{ENTER}");
                        Send("{TAB}");
                        Sleep(100);
                    }
                    if (bottomChordSize.IsSome)
                    {
                        Sleep(100);
                        ControlClick("Joist Properties", "", "TComboBox3", "left", 1, 81, 10);
                        Sleep(200);
                        match(bottomChordSize, Some: s => Send(s.ToString()), None: () => { });
                        Sleep(200);
                        Send("{ENTER}");
                        Send("{TAB}");
                        Sleep(100);
                    }
                    WinWaitActive("Joist Properties");
                }
            }
            else
            {
                failwith("'AddSelfWeight(double selfWeight)' must be called with the 'Joist Properties' window open");
            }
        }

        public static void ApplyModifications((bool AddSelfWeight, bool SetChordsForInertia, bool ApplyAdditionalTakeoffInfo) modifications)
        {
            var excelFileFilter = "Excel Documents | *.xlsx; *.xlsm";

            var requiresCoordinatorTools = (modifications.AddSelfWeight || modifications.SetChordsForInertia);
            var requiresTakeoff = modifications.ApplyAdditionalTakeoffInfo;

            Func<ExcelPackage> getCoordinatorToolsPackage =
                () =>
                     match(
                         General.GetFileName("Select Coordinator Tools", excelFileFilter),
                         Some: s => General.GetExcelPackage(s),
                         None: () => throw new System.Exception("Coordinator Tools was not selected."));

            Func<ExcelPackage> getTakeoffPackage =
                () =>
                     match(
                         General.GetFileName("Select Takeoff", excelFileFilter),
                         Some: s => General.GetExcelPackage(s),
                         None: () => throw new System.Exception("Takeoff was not selected."));

            ExcelPackage coordinatorToolsPackage = requiresCoordinatorTools ? getCoordinatorToolsPackage() : null;
            ExcelPackage takeoffPackge = requiresTakeoff ? getTakeoffPackage() : null;



            if (requiresCoordinatorTools || requiresTakeoff)
            {
                GoToJoistList();
                var previousMark = "";
                var isFinalMark = false;
                do
                {
                    Send("{ENTER}");
                    WinWaitActive("Joist Properties");
                    var currentMark = ControlGetText("Joist Properties", "", "TDBEdit20");
                    if (currentMark != previousMark)
                    {
                        bool requiresSelfWeight, requiresChordsForInertia, requiresDeflectionSet, requiresErfos, requiresWoodNailerSet;
                        requiresSelfWeight = requiresChordsForInertia = requiresDeflectionSet = requiresErfos = requiresWoodNailerSet = false;

                        if (modifications.AddSelfWeight)
                        {
                            var selfWeightDict = GetSelfWeights(coordinatorToolsPackage);
                            if (selfWeightDict.ContainsKey(currentMark))
                            {
                                var selfWeight = selfWeightDict[currentMark];
                                if (selfWeight != 0.0)
                                {
                                    AddSelfWeight(selfWeight, out requiresSelfWeight);
                                }
                            }
                        }
                        if (modifications.SetChordsForInertia)
                        {
                            var inertiaDict = GetChordsFromInertiaCheck(coordinatorToolsPackage);
                            if (inertiaDict.ContainsKey(currentMark))
                            {
                                var inertiaInfo = inertiaDict[currentMark];
                                var tc =
                                    inertiaInfo.TopChordChanged ?
                                    Some(inertiaInfo.TopChordSize) :
                                    Option<string>.None;
                                var bc =
                                    inertiaInfo.BottomChordChanged ?
                                    Some(inertiaInfo.BottomChordSize) :
                                    Option<string>.None;
                                SetChords(tc, bc, out requiresChordsForInertia);
                            }
                        }
                        if (modifications.ApplyAdditionalTakeoffInfo)
                        {
                            var modificationInfo = GetAdditionalTakeoffInfo(takeoffPackge);
                            var currentMarkWithoutLetters =
                                new string(
                                    currentMark
                                    .Where(c => (char.IsDigit(c)))
                                    .ToArray());
                            if (modificationInfo.ContainsKey(currentMarkWithoutLetters))
                            {
                                var addTakeoffInfo = modificationInfo[currentMarkWithoutLetters];
                                SetDeflection(addTakeoffInfo.TlDeflection, addTakeoffInfo.LlDeflection, out requiresDeflectionSet);
                                match(addTakeoffInfo.WnSpacing, Some: space => { SetWoodnailer(space, out requiresDeflectionSet); }, None: () => { });
                                SetErfos(addTakeoffInfo.ErfoAtLe, addTakeoffInfo.ErfoAtRe, out requiresErfos);
                            }
                        }
                        WinWaitActive("Joist Properties");
                        var requiresModification = requiresSelfWeight || requiresChordsForInertia || requiresDeflectionSet || requiresErfos || requiresWoodNailerSet;
                        if (requiresModification)
                        {
                            ControlClick("Joist Properties", "", "TBitBtn2", "left", 1, 46, 11);
                        }
                        else
                        {
                            WinClose("Joist Properties", "");
                        }
                        WinWaitClose("Joist Properties");
                        Sleep(500);
                    }
                    else
                    {
                        isFinalMark = true;
                        WinClose("Joist Properties");
                    }
                    previousMark = currentMark;
                    WinWaitActive("Joist Design");
                    Send("{DOWN}");
                } while (!isFinalMark);
            }

            MessageBox.Show("Modifications Complete!");

        }

    }
}
