using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WindowsFormsApp1
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int Sheet) // constructor
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public string readCell(int i, int j) // i and j refer to row and column
        {
            // Excel iterates from 1 onward, not 0 onward
            i++;
            j++;
            // check that cell values are not null
            if (ws.Cells[i,j].Value2 != null)
            {
                return ws.Cells[i, j].Value;
            }
            else
            {
                return "";
            }
        }
    }
}
