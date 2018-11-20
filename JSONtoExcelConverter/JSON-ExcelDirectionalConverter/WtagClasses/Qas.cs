using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JSON_ExcelDirectionalConverter.WtagClasses
{
    public class Qas
    {
        //public string confuse { get; set; }
        public bool confuseQt { get; set; }
        public bool confuseQf { get; set; }
        public bool confuseSat { get; set; }
        public bool confuseLat { get; set; }
        public double time { get; set; }//추가

        public string id { get; set; }
        public string question { get; set; }
        //public string question_original { get; set; }
        public string question_en { get; set; }
        //public IList<string> question_tagged { get; set; }
        public string question_tagged { get; set; }
        public string questionType { get; set; }
        public string questionFocus { get; set; }
        public string questionSAT { get; set; }
        public string questionLAT { get; set; }
        public bool etriQtCheck { get; set; }
        public bool etriQfCheck { get; set; }
        public bool etriLatCheck { get; set; }
        public bool etriSatCheck { get; set; }
        public string etriQt { get; set; }//추가
        public string etriQf { get; set; }//추가
        public string etriLat { get; set; }//추가
        public string etriSat { get; set; }//추가
        public bool checkIndividual { get; set; }//추가
       


        public IList<object> answers { get; set; }
        
    }
}
