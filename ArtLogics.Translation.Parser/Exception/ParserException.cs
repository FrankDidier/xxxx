using NLog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Exception
{
    public class ParserException : System.Exception
    {
        public string Cell { get; set; }
        public ParserExceptionType Type { get; set; }
        public ParserExceptionMode Mode { get; set; }
        public string Subcondition { get; set; }
        public int SubconditionLine { get; set; }
        public string Sheet { get; set; }

        private static ILogger _log = LogManager.GetCurrentClassLogger();

        public ParserException (ParserExceptionType type) : base()
        {
            Type = type;
        }

        public void Log()
        {
            _log.Info(this.ToString());
        }

        public override string ToString()
        {
            if (Subcondition != null) {
                return "An Error occured (" + Type + ") with the test in the cell " + Cell + " of the sheet " + Sheet + " with the subcondition " + Subcondition + " it's located in the " + SubconditionLine + " in the cell, parser in Mode ( " + Mode + " )";
            } else if (Cell != null)
            {
                return "An Error occured (" + Type + ") with the test in the cell " + Cell + " of the sheet " + Sheet + " at Line " + SubconditionLine + ", parser in Mode ( " + Mode + " )";
            } else
            {
                return "An Error occured (" + Type + ") in the sheet " + Sheet + " at Line " + SubconditionLine + ", parser in Mode(" + Mode + ")";
            }
        }
    }
}
