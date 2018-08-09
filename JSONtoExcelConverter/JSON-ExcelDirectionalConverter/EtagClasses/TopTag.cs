using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.EtagClasses
{
    class ETRI_TopTag
    {
        public string version { get; set; }
        public string creator { get; set; }

        public IList<object> data { get; set; }
    }
}
