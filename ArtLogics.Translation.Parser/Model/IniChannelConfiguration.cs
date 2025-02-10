using ArtLogics.TestSuite.TestResults;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Model
{
    public class IniChannelConfiguration
    {
        public Tuple<string, float> VoltemeterMeasure { get; set; }
        public Tuple<string, Severity> VoltmeterErrorType { get; set; }
        public Tuple<string, string> ErrorMessage { get; set; }
    }
}
