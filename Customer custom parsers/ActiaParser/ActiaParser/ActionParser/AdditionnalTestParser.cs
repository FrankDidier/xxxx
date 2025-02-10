using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Utils;
using OfficeOpenXml;

namespace ActiaParser.ActionParser
{
    public class AdditionnalTestParser : ActionParserBase
    {
        public ExcelWorksheet Worksheet { get; set; }
        public FlowChartState currentState { get; set; }

        public AdditionnalTestParser(ExcelWorkbook workbook, Project project, List<FlowChartState> stateList, List<Transition> transitionList) : base(workbook, project)
        {
            this.stateList = stateList;
            this.transitionList = transitionList;
        }
        
        public void ParseAdditionnalTest()
        {
            Worksheet = Workbook.Worksheets[ParserStaticVariable.AdditionnalTestSheet];

            var columnResult = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.AdditionnalTestResultColName, Worksheet);
            var columnCondition = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.AdditionnalTestConditionColName, Worksheet);

            for (var i = 3; Worksheet.Cells[i, columnResult].Value != null; i++)
            {
                try
                {
                    /*var result = Worksheet.Cells[i, columnResult].Value.ToString().ToLower();
                    var conditionString = Worksheet.Cells[i, columnCondition].Value?.ToString().ToLower();*/

                    var testParameters = new TestParameters(Worksheet, i, columnResult, columnCondition, currentState, -1, true);

                    if (/*result*/testParameters.Function == "state")
                    {

                        CreateState(/*conditionString*/Worksheet.Cells[i, columnCondition].Value?.ToString().ToLower());
                        currentState = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
                    }
                    else
                    {
                        if (currentState == null)
                            throw new ParserException(ParserExceptionType.STATE_NOT_DEFINED);

                        if (/*conditionString*/testParameters.Conditions != null && testParameters.Conditions.Count > 0)
                        {
                            //var conditions = conditionString.Split('\n');

                            var c = 1;
                            foreach (var conditionWithComment in testParameters.Conditions/*conditions*/)
                            {
                                canLoopSignalUsed.Clear();
                                if (conditionWithComment.Condition == "" || conditionWithComment.Condition.Substring(0, 2) == "//")
                                    continue;

                                var tableWithComment = conditionWithComment.Condition.Split(new string[] { "//" }, StringSplitOptions.None);
                                var condition = tableWithComment[0];

                                var tabOperation = condition.Split('且');

                                foreach (var op in tabOperation)
                                {
                                    if (op != "")
                                    {
                                        try
                                        {
                                            SetChannel(op, testParameters.Function/*result*/, currentState, Resources.lang.ChannelActionSetter + " " + testParameters.Function/*result*/);
                                        }
                                        catch (ParserException ex)
                                        {
                                            ex.Cell = Worksheet.Cells[i, columnCondition].ToString();
                                            ex.Subcondition = op;
                                            ex.SubconditionLine = c;
                                            ex.Mode = ParserExceptionMode.SETTER;
                                            ex.Sheet = Worksheet.Name;
                                            throw ex;
                                        }
                                    }
                                }

                                string result = "";
                                BuildOperation(condition.ToLower(), testParameters.Function, out result);
                                ProjectUtilsFunction.BuildUserAction(currentState, result/*result*/, UserActionButtons.Pass | UserActionButtons.Fail);


                                CleanCanLoops(currentState);
                                foreach (var op in tabOperation)
                                {
                                    if (op != "")
                                    {
                                        try
                                        {
                                            SetChannel(op, testParameters.Function/*result*/, currentState, Resources.lang.ChannelActionReset + " " + testParameters.Function/*result*/, true);
                                        }
                                        catch (ParserException ex)
                                        {
                                            ex.Cell = Worksheet.Cells[i, columnCondition].ToString();
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
                        else
                        {
                            ProjectUtilsFunction.BuildUserAction(currentState, testParameters.Function/*result*/, UserActionButtons.Pass | UserActionButtons.Fail);
                        }
                    }
                }
                catch (Exception err)
                {
                    HandleException(err, Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell + " " + Worksheet.Cells[i, columnCondition]);
                    var error = err as ParserException;
                    if (error != null && error.Type == ParserExceptionType.RESISTIVE_VALUE_UNAVAILABLE)
                    {
                        var soundPath = "Sound/Sound.wav";
                        var soundPathDest = ParserStaticVariable.GlobalPath + soundPath.Substring(soundPath.LastIndexOf("/") + 1);

                        File.Copy(soundPath, soundPathDest, true);

                        ProjectUtilsFunction.BuildUserAction(currentState, Resources.lang.ConditionCannotBeTestedManually + error.Subcondition + Resources.lang.PleaseTestItManually, UserActionButtons.Fail, null, soundPathDest);
                    }
                    continue;
                }
            }
        }
    }
}
