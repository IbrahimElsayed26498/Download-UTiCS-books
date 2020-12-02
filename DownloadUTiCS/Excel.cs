using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace DownloadUTiCS
{
    class Excel
    {
        private readonly _Application _excel = new Application();
        private readonly Workbook _wb;
        private readonly Worksheet _ws;
        private readonly int _numberOfRows;
        private readonly int _numberOfColumns;
        private Dictionary<string, int> _dictionary;
        public Excel(string path, int sheet)
        {
            _wb = _excel.Workbooks.Open(path);
            _ws = _excel.Worksheets[sheet];
            _numberOfRows = _ws.UsedRange.Rows.Count;
            _numberOfColumns = _ws.UsedRange.Columns.Count;
            Console.WriteLine("Excel Sheet Opened!");
            _dictionary = new Dictionary<string, int>();
            StoreColumnsNames();
        }

        private void StoreColumnsNames()
        {
            for (var i = 0; i < _numberOfColumns; i++)
            {
                _dictionary.Add(ReadCell(0, i), i);
            }
        }

        public string ReadCell(int r, int c)
        {
            r++;
            c++;
            var cell = _ws.Cells[r, c].Value2;
            return cell != null ? cell.ToString() : "";
        }
        public int GetColumnNumber(string columnName)
        {
            int columnNumber;
            _dictionary.TryGetValue(columnName, out columnNumber);
            return columnNumber;
        }

        public int RowsNumbers()
        {
            return _numberOfRows;
        }
        public bool Close()
        {
            try
            {
                _wb.Close(0);
                _excel.Quit();
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }
    }

}
