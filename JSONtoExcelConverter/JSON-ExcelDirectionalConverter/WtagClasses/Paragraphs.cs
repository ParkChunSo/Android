using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.WtagClasses
{
    public class Paragraphs
    {
        public string context { get; set; }
        //public string context_original { get; set; }
        public string context_en { get; set; }
        //public IList<string> context_tagged { get; set; }
        public string context_tagged { get; set; }
        public IList<object> qas { get; set; }
    }
}
