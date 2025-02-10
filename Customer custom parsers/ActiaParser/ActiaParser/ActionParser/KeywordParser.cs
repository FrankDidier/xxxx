using ActiaParser.Define;
using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Shared;
using ArtLogics.Translation.Parser.Utils;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public static class KeywordParser
    {
        public static List<ChannelOutputInfo> FunctionChannelOutputMap { get; set; } = new List<ChannelOutputInfo>();
        public static List<ChannelInputInfo> FunctionChannelInputMap { get; set; } = new List<ChannelInputInfo>();
        public static Dictionary<string, ChannelInputInterpretations> ChannelInputInterpretations { get; set; } = new Dictionary<string, ChannelInputInterpretations>();
        public static Dictionary<string, string> KeywordFunctionMap { get; set; } = new Dictionary<string, string>();
        public static Dictionary<string, string> Keywords { get; set; } = new Dictionary<string, string>();
        public static Dictionary<string, string> StatusTestMap { get; set; } = new Dictionary<string, string>();

        private static ILogger _log = LogManager.GetCurrentClassLogger();

        public static void Reset()
        {
            FunctionChannelOutputMap.Clear();
            FunctionChannelInputMap.Clear();
            ChannelInputInterpretations.Clear();
            KeywordFunctionMap.Clear();
            Keywords.Clear();
            StatusTestMap.Clear();
        }

        public static void ParseFunctionMap(ExcelWorkbook Workbook)
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.KeywordsFunctionSheet];

            var columnNumberKeyword = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.KeywordsColName, Worksheet);
            var colNumberFunction = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.FunctionColName, Worksheet);

            for (var i = 3; Worksheet.Cells[i, columnNumberKeyword].Value != null; i++)
            {
                try
                {
                    KeywordFunctionMap[Worksheet.Cells[i, columnNumberKeyword].Value.ToString().ToLower().Trim()] = Worksheet.Cells[i, colNumberFunction].Value.ToString().ToLower().Trim();
                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtRow + " " + i);
                    continue;
                }
            }
        }

        public static void ParseKeywords(ExcelWorkbook Workbook)
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.KeywordsSheet];

            var columnNumberKeyword = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.KeywordsColName, Worksheet);
            var colNumberDescription = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.DescriptionColName, Worksheet);

            for (var i = 3; Worksheet.Cells[i, columnNumberKeyword].Value != null; i++)
            {
                try
                {
                    Keywords[Worksheet.Cells[i, columnNumberKeyword].Value.ToString().ToLower().Trim()] = Worksheet.Cells[i, colNumberDescription].Value.ToString().ToLower().Trim();

                } catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtRow + " " + i);
                    continue;
                }
            }
        }

        public static void ParseStatus(ExcelWorkbook Workbook)
        {
            /*var Worksheet = Workbook.Worksheets[ParserStaticVariable.StatusSheet];

            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.StatusFunctionColName, Worksheet);
            var colNumberDefinition = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.StatusDefinitionColName, Worksheet);

            for (var i = 4; Worksheet.Cells[i, columnNumberFunction].Value != null; i++)
            {
                try
                {
                    StatusTestMap[Worksheet.Cells[i, columnNumberFunction].Value.ToString()] = Worksheet.Cells[i, colNumberDefinition].Value.ToString();
                }
                catch
                {
                    _log.Info("Wrong value in sheet " + Worksheet.Name + " at row " + i);
                    continue;
                }
            }*/
        }
    }
}
