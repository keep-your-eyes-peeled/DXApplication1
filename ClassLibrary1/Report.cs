using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassLibrary1
{
    public class Report
    {

        public Report(Dictionary<string, string> fieldValues)
        {
            FieldValues = fieldValues;
        }

        private Dictionary<string, string> fieldValues;

        

        public Dictionary<string, string> FieldValues { get => fieldValues; set => fieldValues = value; }
    }
}
