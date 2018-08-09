using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.EtagClasses
{
    class ETRI_Paragraphs
    {
        public string context { get; set; }
        public string context_en { get; set; }
        public string context_tagged { get; set; }

        public IList<object> qas { get; set; }
    }
}
