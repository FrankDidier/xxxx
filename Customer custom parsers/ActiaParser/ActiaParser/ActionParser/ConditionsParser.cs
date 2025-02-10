using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class ConditionsParser
    {
        public string Condition { get; set; }
        public int Line { get; set; }

        public ConditionsParser (string condition, int line)
        {
            Condition = condition;
            Line = line;
        }
    }
}
