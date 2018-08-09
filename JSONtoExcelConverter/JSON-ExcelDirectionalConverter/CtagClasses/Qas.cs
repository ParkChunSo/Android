using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.CtagClasses
{
    class Cross_Qas
    {
        public string id { get; set; }
        public bool confuseQt1 { get; set; }
        public bool confuseQf1 { get; set; }
        public bool confuseSat1 { get; set; }
        public bool confuseLat1 { get; set; }

        public string question { get; set; }
        public string question_en { get; set; }
        public string question_tagged1 { get; set; }
        public string questionType1 { get; set; }
        public string questionFocus1 { get; set; }
        public string questionSAT1 { get; set; }
        public string questionLAT1 { get; set; }

        public bool confuseQt2 { get; set; }
        public bool confuseQf2 { get; set; }
        public bool confuseSat2 { get; set; }
        public bool confuseLat2 { get; set; }

        public string question_tagged2 { get; set; }
        public string questionType2 { get; set; }
        public string questionFocus2 { get; set; }
        public string questionSAT2 { get; set; }
        public string questionLAT2 { get; set; }

        public bool confuseQt3 { get; set; }
        public bool confuseQf3 { get; set; }
        public bool confuseSat3 { get; set; }
        public bool confuseLat3 { get; set; }

        public string question_tagged3 { get; set; }
        public string questionType3 { get; set; }
        public string questionFocus3 { get; set; }
        public string questionSAT3 { get; set; }
        public string questionLAT3 { get; set; }


        public IList<object> answers { get; set; }
    }
}
