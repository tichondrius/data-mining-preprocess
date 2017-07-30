using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Bai1A
{
    class BinningWidth
    {
        public double Min { get; set; }
        public double Max { get; set; }

        public List<ExpandoObject> lstRecord { get; set; }

        public BinningWidth(double Min, double Max) {
            this.Min = Min;
            this.Max = Max;
            lstRecord = new List<ExpandoObject>();
        }
    }
}
