using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using Microsoft.Office.Interop.Excel;
using LinqToExcel;

namespace frmReadAndWriteExcel
{
    public partial class FrmReadAndSaveExcel : Form
    {
        public FrmReadAndSaveExcel()
        {
            InitializeComponent();
            LsvLoadData.Columns.Add(new ColumnHeader() { Text = "ID" });
            LsvLoadData.Columns.Add(new ColumnHeader() { Text = "Name" });

        }

        private async void BtnSave_Click(object sender, EventArgs e)
        {
            Cursor.Current = Cursors.WaitCursor;
            using (var save =new SaveFileDialog() {Filter= "Excel Workbook|*.xlsx",ValidateNames=true })
            {
                List<string> list = new List<string>() { "D" };
                if (save.ShowDialog()==DialogResult.OK)
                {
                    
                    ReadAndWrite readAndWrite = new ReadAndWrite(save.FileName);
                    await readAndWrite.SaveAsync(list);

                }
            }
            MessageBox.Show("Done!","Save file",MessageBoxButtons.OK,MessageBoxIcon.Information,MessageBoxDefaultButton.Button1);
            Cursor.Current = Cursors.Default;
        }

        private async void BtnLoad_Click(object sender, EventArgs e)
        {
            LsvLoadData.Items.Clear();
            Cursor.Current = Cursors.WaitCursor;
            using (var Opf = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if (Opf.ShowDialog() == DialogResult.OK)
                {
                    ReadAndWrite readAndWrite = new ReadAndWrite(Opf.FileName);
                    var tuple = await readAndWrite.ReadAsync();//ReadWithInteropExcel();
                    foreach (var item in tuple.Item1)
                    {
                        ListViewItem listViewItem = new ListViewItem() {Text=item.Key.ToString() };
                        listViewItem.SubItems.Add(item.Value.ToString());
                        LsvLoadData.Items.Add(listViewItem);
                    }
                    
                    /*int debug = 0;
                    //using (var stream=File.Open(Opf.FileName,FileMode.Open,FileAccess.Read))
                    var op = new ExcelQueryFactory(Opf.FileName);
                    var data = from d in op.Worksheet<Info>("Info")
                               select d;
                    _Application _app = new Excel.Application();
                    Workbook workbook = _app.Workbooks.Open(Opf.FileName);
                    Worksheet worksheet = _app.ActiveSheet as Worksheet;
                    //Worksheet worksheet = workbook.Sheets[1];
                    Excel.Range xlRange = worksheet.UsedRange;

                    int count = xlRange.Columns.Count;

                    var a = worksheet.Cells[2, 2].Value2;
                    workbook.Close();
                    _app.Quit();*/
                }
            }
            Cursor.Current = Cursors.Default;
        }
    }
}
