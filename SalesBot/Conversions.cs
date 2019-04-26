using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Design.SalesTools.Sei;

namespace SalesBot
{
    class Conversions
    {
        public static void ConvertSeiTakeoff()
        {
            var openFile = new Microsoft.Win32.OpenFileDialog();
            openFile.Filter = "Excel Documents|*.xlsx;*.xlsm";

            var result = openFile.ShowDialog();
            if (result == true)
            {
                var fileName = openFile.FileName;
                CreateTakeoff.CreateTakeoff(fileName);
            }
        }
    }
}
