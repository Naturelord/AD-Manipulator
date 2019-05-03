using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows;
using System.Windows.Controls;

namespace AD_Manipulator
{
    class ExcelAccessor
    {
        public Excel.Workbook wb { get; private set; }
        public Excel.Worksheet Sheet { get; private set; }
        public Excel.Application excel;
        private string path { get; set; }

        public ExcelAccessor(string pathGiven)
        {
            this.path = pathGiven;
            Open(path);
            // [ Row ] [ Column ]
        }

        /// <summary>
        /// Must call after you have accessed and pulled data from Excel
        /// </summary>
        public void closeOut()
        {
            wb.Close();
            excel.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
        }
        public void Save()
        {
            wb.Save();
        }
        public void Open(string path)
        {
            excel = new Excel.Application();
            wb = excel.Workbooks.Open(path);
            Sheet = wb.Worksheets[1];
            MessageBox.Show(Sheet.Cells[1][3].Value2, "Hello");
        }



    }
}
