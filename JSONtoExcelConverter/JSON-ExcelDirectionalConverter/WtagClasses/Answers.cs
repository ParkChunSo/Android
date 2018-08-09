using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.WtagClasses
{
    public class Answers
    {
        public string text { get; set; }
        //public string text_original { get; set; }
        public string text_en { get; set; }
        //public IList<string> text_tagged { get; set; }
        //public IList<string> text_syn { get; set; }
        public string text_tagged { get; set; }
        public string text_syn { get; set; }
        public int answer_start { get; set; }
        public int answer_end { get; set; }
    }
}
