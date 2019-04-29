using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Design.SalesTools.Sei;
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
                Some: fileName => CreateTakeoff.CreateTakeoff(fileName),
                None: () => { }
                );
            MessageBox.Show("Conversion Complete!");
        }

        public static void ConvertGemTakeoff()
        {

        }
    }
}
