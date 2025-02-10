using ActiaParser.Define;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public static class InputParser
    {
        public static decimal VBat;
        public static float PSUC;
        public static float PSUV;

        public static void ParseVBat(ExcelWorkbook Workbook)
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.OutputCurrentDeratingSheet];

            var vBat = Worksheet.Cells[ParserStaticVariable.VbatCell].Value.ToString();

            VBat = decimal.Parse(vBat);
        }
    }
}
