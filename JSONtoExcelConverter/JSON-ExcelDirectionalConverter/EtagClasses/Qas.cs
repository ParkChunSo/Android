using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.EtagClasses
{
    class ETRI_Qas
    {
        public string id { get; set; }
        public string question { get; set; }
        public string question_en { get; set; }
        public string question_tagged { get; set; }
        public string questionType { get; set; }
        public string questionFocus { get; set; }
        public string questionSAT { get; set; }
        public string questionLAT { get; set; }
        
        public IList<object> answers { get; set; }
    }
}
