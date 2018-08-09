using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.CtagClasses
{
    class Cross_TopTag    {
        public string version { get; set; }
        public string creator { get; set; }
        public int progress { get; set; }
        public string formatt { get; set; }
        public double time { get; set; }
        public bool check { get; set; }
        public string firstfile { get; set; }
        public string secondfile { get; set; }
        public IList<object> data { get; set; }
    }
}
