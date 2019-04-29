using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LanguageExt;
using OfficeOpenXml;

namespace SalesBot
{
    public class General
    {
        public static Option<string> GetFileName(string title, string filterString)
        {
            var openFile = new Microsoft.Win32.OpenFileDialog();
            openFile.Filter = filterString;
            openFile.Title = title;

            var result = openFile.ShowDialog();
            if (result == true)
            {
                return Option<string>.Some(openFile.FileName);
            }
            else
            {
                return Option<string>.None;
            }
        }

        public static ExcelPackage GetExcelPackage(string fileName)
        {
            var fileInfo = new System.IO.FileInfo(fileName);
            return new ExcelPackage(fileInfo, false);
        }
    }
}
