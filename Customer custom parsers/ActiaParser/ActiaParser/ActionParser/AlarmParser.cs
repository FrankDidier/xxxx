using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.Parser.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class AlarmParser : ActionParserBase
    {
        public AlarmParser(ExcelWorkbook Workbook, Project project) : base(Workbook, project)
        {
            System.IO.Directory.CreateDirectory(ParserStaticVariable.GlobalPath);
        }

        public void ParseActionsAlarm()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.AlarmsSheet];

            var columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.AlarmDefinitionColName, Worksheet);
            var columnNumberLangage = ExcelUtilsFunction.GetColumnNumber(1, ParserStaticVariable.AlarmLangageColName, Worksheet);
            var langage = Worksheet.Cells[2, columnNumberLangage].Value.ToString();

            var columnNumberName = ExcelUtilsFunction.GetColumnNumber(4, langage, Worksheet);
            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            for (var i = 5; Worksheet.Cells[i, columnNumberName].Value != null; i++)
            {
                try
                {
                    canLoopSignalUsed.Clear();
                    if (Worksheet.Cells[i, columnNumberDefinition].Value == null)
                        continue;

                    /*var name = Worksheet.Cells[i, columnNumberName].Value.ToString();
                    var TestDefinition = Worksheet.Cells[i, columnNumberDefinition].Value.ToString();*/

                    var testParameters = new TestParameters(Worksheet, i, columnNumberName, columnNumberDefinition, state);

                    //var conditions = TestDefinition.Split('\n').ToList();
                    var c = 1;
                    foreach (var conditionInfo in testParameters.Conditions)
                    {
                        var conditionWithComment = conditionInfo.Condition;

                        if (conditionWithComment == "" || conditionWithComment.Substring(0, 2) == "//")
                            continue;

                        var testInformation = " / " + c + " / " + conditionInfo.Condition;

                        var tableWithComment = conditionWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                        var condition = tableWithComment[0];

                        var tab = condition.Split(new string[] { "-", "—" }, 2, StringSplitOptions.None);
                        var result = tab[0].Trim().ToLower();
                        var operation = tab[1].Trim().ToLower();

                        operation = BuildOperation(operation, result, out result, false);

                        var tabOperation = operation.Split('且');

                        foreach (var op in tabOperation)
                        {
                            if (op != "")
                            {
                                try
                                {
                                    SetChannel(op, testParameters.Function, state, Resources.lang.ChannelActionSetter + " " + testParameters.Function + testInformation);
                                }
                                catch (ParserException ex)
                                {
                                    ex.Cell = Worksheet.Cells[i, columnNumberDefinition].ToString();
                                    ex.Subcondition = op;
                                    ex.SubconditionLine = c;
                                    ex.Mode = ParserExceptionMode.SETTER;
                                    ex.Sheet = Worksheet.Name;
                                    throw ex;
                                }
                            }
                        }
                        
                        ProjectUtilsFunction.BuildUserAction(state, Resources.lang.YouShouldSee + " : " + testParameters.Function + " " + result + testInformation, UserActionButtons.Pass | UserActionButtons.Fail);

                        CleanCanLoops(state, testInformation);
                        foreach (var op in tabOperation)
                        {
                            if (op != "")
                            {
                                try
                                {
                                    SetChannel(op, testParameters.Function, state, Resources.lang.ChannelActionReset + " " + testParameters.Function + testInformation, true);
                                }
                                catch (ParserException ex)
                                {
                                    ex.Cell = Worksheet.Cells[i, columnNumberDefinition].ToString();
                                    ex.Subcondition = op;
                                    ex.SubconditionLine = c;
                                    ex.Mode = ParserExceptionMode.SETTER;
                                    ex.Sheet = Worksheet.Name;
                                    throw ex;
                                }
                            }
                        }
                        c++;
                    }
                }
                catch (Exception err)
                {
                    HandleException(err, Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell  + " " + Worksheet.Cells[i, columnNumberDefinition]);
                    var error = err as ParserException;
                    if (error != null && error.Type == ParserExceptionType.RESISTIVE_VALUE_UNAVAILABLE)
                    {
                        var soundPath = "Sound/Sound.wav";
                        var soundPathDest = ParserStaticVariable.GlobalPath + soundPath.Substring(soundPath.LastIndexOf("/") + 1);

                        File.Copy(soundPath, soundPathDest, true);

                        ProjectUtilsFunction.BuildUserAction(state, Resources.lang.ConditionCannotBeTestedManually + error.Subcondition + Resources.lang.PleaseTestItManually, UserActionButtons.Fail, null, soundPathDest);
                    }
                    continue;
                }
            }
        }
    }
}
