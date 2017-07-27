using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1A
{
    class ExcelData
    {
        public List<ExpandoObject> lstRecords { get; set; }
        public List<TitleHeader> headers { get; set; }
        public ExcelData() {
            lstRecords = new List<ExpandoObject>();
            headers = new List<TitleHeader>();
        }
    }
}
