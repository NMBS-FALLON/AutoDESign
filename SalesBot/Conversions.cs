using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LanguageExt;
using static LanguageExt.Prelude;
using System.Windows;


namespace SalesBot
{
    class Conversions
    {
        public static void ConvertSeiTakeoff()
        {
            var seiTakeoffFileName = General.GetFileName("Select SEI Takeoff", "Excel Documents|*.xlsx;*.xlsm");
            match(
                seiTakeoffFileName,
                Some: fileName =>
                    {
                        Design.SalesTools.Sei.CreateTakeoff.CreateTakeoff(fileName);
                        MessageBox.Show("Conversion Complete!");
                    },
                None: () => { }
                );
            
        }

        public static void ConvertGemTakeoff()
        {
            var gemTakeoffFileName = General.GetFileName("Select GEM Takeoff", "Excel Documents|*.xlsx;*.xlsm;*.xls");
            match(
                gemTakeoffFileName,
                Some: fileName =>
                {
                    Design.SalesTools.Gem.CreateTakeoff.CreateTakeoff(fileName);
                    MessageBox.Show("Conversion Complete!");
                },
                None: () => { }
                );
        }
    }
}
