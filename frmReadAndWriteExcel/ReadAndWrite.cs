using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using LinqToExcel;

namespace frmReadAndWriteExcel
{
    public class ReadAndWrite
    {
        private string address;

        public string Address
        {
            get { return address; }
            set { address = value; }
        }
        public ReadAndWrite()
        {

        }
        public ReadAndWrite(string s)
        {
            this.Address = s;
        }

        public Task SaveAsync(List<string> list)
        {
            return Task.Factory.StartNew(() => SaveFile(list));
        }

        private void SaveFile(List<string> Headerlist)
        {
            _Application _app = new Application();
            Workbook workbook = _app.Workbooks.Add(XlSheetType.xlWorksheet);
            Worksheet worksheet = (Worksheet)_app.ActiveSheet;
            _app.Visible = false;
            int i = 0;
            foreach (var item in Headerlist)
            {
                i++;
                worksheet.Cells[1, i] = item.ToString();
            }
            workbook.SaveAs(address, XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, true, false,
                XlSaveAsAccessMode.xlNoChange, XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing);
            _app.Quit();
        }

        public Task<Tuple<Dictionary<string, string>>> ReadAsync()
        {
            return Task.Factory.StartNew(() => ReadWithInteropExcel());
        }

        private Tuple<Dictionary<string, string>> ReadWithInteropExcel()
        {
            object row;
            object row2;
            Dictionary<string, string> dictionary = new Dictionary<string, string>();
            _Application application = new Application();
            Workbook workbook = application.Workbooks.Open(address);
            Worksheet worksheet = workbook.Worksheets[1];
            Excel.Range range = worksheet.UsedRange;
            int countRow = range.Rows.Count;
            int countColumn = range.Columns.Count;
            for (int i = 2; i <= countRow; i++)
            {
                row = worksheet.Cells[i, 1].Value2;
                row2 = worksheet.Cells[i, 2].Value2;
                dictionary.Add(row.ToString(), row2.ToString());
            }
            return new Tuple<Dictionary<string, string>>(dictionary);
        }
    }
}
