using System;
using _Excel = Microsoft.Office.Interop.Excel;

namespace erice
{
    class ExcelReader
    {
        string path = "";
        _Excel.Application excel;
        _Excel.Workbook wb;
        _Excel.Worksheet ws;

        public ExcelReader(string path, int sheet)
        {
            this.path = path;
            excel = new _Excel.Application();
            wb = excel.Workbooks.Open(path);
            ws = (_Excel.Worksheet)wb.Worksheets[sheet];
        }

        public string ReadCell(int i, int j)
        {
            if (ws.Cells[i, j].Value2 != null)
            {
                return Convert.ToString(ws.Cells[i, j].Value2);
            }
            else
            {
                return "";
            }
        }

        public void Close()
        {
            wb.Close();
            excel.Quit();

            // Release COM objects to avoid memory leaks
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ws);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wb);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);

            ws = null;
            wb = null;
            excel = null;

            GC.Collect();
        }
    }
}
