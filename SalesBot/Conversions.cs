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
            var seiTakeoffFileName = General.GetFileName("Select SEI Takeoff", "Excel Documents|*.xlsx;*.xlsm;*.xls");
            match(
                seiTakeoffFileName,
                Some: fileName =>
                    {
                        try
                        {
                            if (System.IO.Path.GetExtension(fileName) == ".xls")
                            {
                                Design.SalesTools.Gem.CreateTakeoff.CreateTakeoff(fileName);
                            }
                            else
                            {
                                Design.SalesTools.Sei.CreateTakeoff.CreateTakeoff(fileName);
                            }
                            MessageBox.Show("Conversion Complete!");
                        }
                        catch (Exception e)
                        {
                            MessageBox.Show("Conversion Failed: " + e.Message);
                        }

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
                    try
                    {
                        Design.SalesTools.Gem.CreateTakeoff.CreateTakeoff(fileName);
                        MessageBox.Show("Conversion Complete!");
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show("Conversion Failed: " + e.Message);
                    }

                },
                None: () => { }
                );
        }
    }
}
