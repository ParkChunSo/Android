using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.WtagClasses
{
    public class TopName
    {
        //public double time { get; set; }
        public string formatt { get; set; }
        public int progress { get; set; }
        public string version { get; set; }
        public string creator { get; set; }
        public IList<object> data { get; set; }
    }
}
