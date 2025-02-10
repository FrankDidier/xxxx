using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Exception;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class TestParameters
    {
        public FlowChartState State { get; set; }
        public ExcelWorksheet Worksheet { get; set; }
        public string Function { get; set; }
        public string Color { get; set; }
        public ExcelPicture ExcelPicture { get; set; }
        public List<ConditionsParser> Conditions { get; set; }

        public TestParameters(ExcelWorksheet worksheet, int i, int columnNumberFunction, int columnNumberDefinition, FlowChartState state, int colNumberColor = -1, bool testDefinitionCanBeNull = false)
        {
            State = state;
            Worksheet = worksheet;
            object colorValue = null;

            if (colNumberColor > 0)
            {
                colorValue = Worksheet.Cells[i, colNumberColor].Value;
            }
            Function = Worksheet.Cells[i, columnNumberFunction].Value.ToString().Replace("\n", " ").ToLower().Trim();
            if (colorValue != null) {
                Color = (colorValue != null) ? colorValue.ToString() : "";
            }

            if (!testDefinitionCanBeNull && Worksheet.Cells[i, columnNumberDefinition].Value == null)
            {
                var ex = new ParserException(ParserExceptionType.TEST_DEFINITION_NOT_FOUND);

                ex.Cell = Worksheet.Cells[i, columnNumberFunction].ToString();
                ex.SubconditionLine = i;
                ex.Mode = ParserExceptionMode.SETTER;
                ex.Sheet = Worksheet.Name;
                throw ex;
            }

            var conditionsTmp = Worksheet.Cells[i, columnNumberDefinition].Value?.ToString().Split('\n').ToList();

            ExcelPicture = Worksheet.Drawings["SYM_" + i] as ExcelPicture;

            Conditions = new List<ConditionsParser>();
            var e = 1;
            if (conditionsTmp != null)
            {
                foreach (var opWithComment in conditionsTmp)
                {
                    var op = opWithComment.Split(new string[] { "//" }, StringSplitOptions.None)[0];
                    if (op.ToLower().Contains("dm"))
                    {
                        Conditions.Add(new ConditionsParser(op + " singleframe", e));
                        Conditions.Add(new ConditionsParser(op + " multipleframe", e));
                    }
                    else
                    {
                        Conditions.Add(new ConditionsParser(op, e));
                    }
                    e++;
                }
            }
        }
    }
}
