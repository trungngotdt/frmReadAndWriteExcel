using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel.Attributes;

namespace frmReadAndWriteExcel
{
    public class Info
    {
        [ExcelColumn("ID")]
        public string ID { get; set; }
        [ExcelColumn("Name")]
        public string Name { get; set; }
    }
}
