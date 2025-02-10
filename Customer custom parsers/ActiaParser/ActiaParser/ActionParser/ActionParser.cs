using ActiaParser.Define;
using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Actions.Common;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.Limits.Comparisons.MultiRange;
using ArtLogics.TestSuite.Limits.Comparisons.SingleRange;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.CanLoop;
using ArtLogics.TestSuite.Testing.Actions.CaptureSensor;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.TestSuite.TestResults;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.Parser.Utils;
using NLog;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace ActiaParser.ActionParser
{
    public struct LimitInfo
    {
        public string ErrorMessage { get; set; }
        public Severity Severity { get; set; }
        public decimal Gap { get; set; }
    }

    public enum ActionType
    {
        NORMAL,
        BUZZERANDVIDEO,
        OVERLOADRATING
    }

    public class ActionParser : ActionParserBase
    {
        private ExcelWorksheet Worksheet { get; set; }

        private int columnNumberDefinition;
        private int columnNumberFunction;
        private int colNumberUsed;
        private int colNumberNoneValue;
        private List<string> prerequireListAll;

        private float limitValueHigh { get; set; }
        private float limitValueLow { get; set; }

        public ActionParser(ExcelWorkbook Workbook, Project project) : base (Workbook, project)
        {
        }

        public void ParseActionsOutput(ActionType actionType)
        {
            Worksheet = Workbook.Worksheets[ParserStaticVariable.OutputsSheet];

            columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsDefinitionColName, Worksheet);
            columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsFunctionColName, Worksheet);
            colNumberUsed = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsUsedColName, Worksheet);
            colNumberNoneValue = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsNoneValueColName, Worksheet);

            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            //SetAllPrerequire(state, "", "NONE", 4, false, true);
            prerequireListAll = GetAllActivePrerequire(4, true);
            
            for (var i = 4; Worksheet.Cells[i, colNumberUsed].Value != null; i++)
            {
                try
                {
                    if (Worksheet.Cells[i, colNumberUsed].Value.ToString() != "Y")
                        continue;

                    if (Worksheet.Cells[i, colNumberUsed].Value.ToString() != "Y")
                        continue;

                    var preRequireList = GetAllActivePrerequire(i);

                    specificTime = 0;
                    testType = KeywordsEnum.NOKEYWORD;

                    var testParameters = new TestParameters(Worksheet, i, columnNumberFunction, columnNumberDefinition, state);
                    
                        if (actionType == ActionType.BUZZERANDVIDEO && (testParameters.Function != "buzzer" &&
                            testParameters.Function != "video"))
                            continue;
                    
                        if ((actionType == ActionType.NORMAL || actionType == ActionType.OVERLOADRATING) && (testParameters.Function == "buzzer" ||
                            testParameters.Function == "video"))
                            continue;

                    try
                    {
                        SetGeneralUserAction(state,
                            columnNumberFunction, columnNumberDefinition, i);
                    }
                    catch (ParserException ex)
                    {
                        ex.SubconditionLine = i;
                        ex.Mode = ParserExceptionMode.SETTER;
                        ex.Sheet = Worksheet.Name;
                        throw ex;
                    }

                    //default test
                    /*SetAllPrerequire(state,
                        columnNumberFunction, columnNumberDefinition, testParameters.Function, "NONE", i, false);
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayAfterPrerquireNone + " " + i, delayAfterPrerequire);*/

                    BuildOutputTest(testParameters, i, delayTime, "NONE", preRequireList, Worksheet.Cells[i, columnNumberDefinition].ToString(), Worksheet.Cells[i, colNumberNoneValue], actionType, true);

                    /*SetAllPrerequire(state,
                        columnNumberFunction, columnNumberDefinition, testParameters.Function, "NONE", i, true);*/
                    /*ResetAllPRerequire(state, prerequireListAll);
                    SetAllPrerequire(state,
                        columnNumberFunction, columnNumberDefinition, "", "NONE", 4, false, true);*/
                    //dynamic test
                    for (var e = columnNumberFunction + 1; e < columnNumberDefinition; e++)
                    {

                        if (Worksheet.Cells[i, e].Value == null || Worksheet.Cells[i, e].Value.ToString() != "Y")
                            continue;

                        if (Worksheet.Cells[i, columnNumberDefinition].Value == null)
                            continue;
                        
                        var channelName = Worksheet.Cells[3, e].Value.ToString().Replace("\n", "");

                        //setPrerequire
                        /*SetAllPrerequire(state,
                                    columnNumberFunction, columnNumberDefinition, testParameters.Function, channelName, i, false, true);
                        ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayAfterPrerequire + " " + channelName + "" + Resources.lang.AtLine + " " + i, delayAfterPrerequire);*/

                        BuildOutputTest(testParameters, i, delayTime, channelName, preRequireList, Worksheet.Cells[i, columnNumberDefinition].ToString(), Worksheet.Cells[i, colNumberNoneValue], actionType);
                        //reset prerequire
                        /*SetAllPrerequire(state,
                                    columnNumberFunction, columnNumberDefinition, testParameters.Function, channelName, i, true);*/
                        /*ResetAllPRerequire(state, prerequireListAll);
                        SetAllPrerequire(state,
                            columnNumberFunction, columnNumberDefinition, "", "NONE", 4, false, true);*/
                    }

                    SetGeneralUserAction(state,
                        columnNumberFunction, columnNumberDefinition, i, true);
                }
                catch (Exception err)
                {
                    HandleException(err, Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell + " " + Worksheet.Cells[i, columnNumberDefinition]);
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
            //var prerequireListAll = GetAllActivePrerequire(columnNumberFunction, columnNumberDefinition, 4, true);
            //ResetAllPRerequire(state, prerequireListAll);
        }



        private void ResetAllPRerequire(FlowChartState state, List<string> preRequireList)
        {
            var delayAfter = true;
            foreach (var preRequireName in preRequireList)
            {
                var channels = KeywordParser.FunctionChannelInputMap.Where(o => o.Function != "0"
                    && preRequireName.Contains(o.Function));

                var channelInputInfos = GetChannelInfoFromKeyWord(preRequireName, channels, false);

                foreach (var channelInputInfo in channelInputInfos)
                {
                    if (channelInputInfo.BaseDefaultValue != null) {
                        channelInputInfo.DefaultValue = channelInputInfo.BaseDefaultValue;
                    }
                }

                foreach (var channel in channels)
                {
                    if (channel.BaseDefaultValue != null)
                    {
                        channel.DefaultValue = channel.BaseDefaultValue;
                    }
                }

                SetChannel(preRequireName + KeywordsEnum.OFF.ToString(), preRequireName, state, Resources.lang.ChannelActionResetPrerequire, true, delayAfter);
                delayAfter = false;
            }
        }

        private List<string> GetAllActivePrerequire(int line, bool getAll = false)
        {
            List<string> listPrerequire = new List<string>();

            for (var e = columnNumberFunction + 1; e < columnNumberDefinition; e++)
            {
                if ((Worksheet.Cells[line, e].Value == null || Worksheet.Cells[line, e].Value.ToString() != "Y") && !getAll)
                    continue;

                if ((Worksheet.Cells[line, e].Value == null || (Worksheet.Cells[line, e].Value.ToString() != "Y" &&
                    Worksheet.Cells[line, e].Value.ToString() != "N")) && getAll)
                    continue;

                if (Worksheet.Cells[4, columnNumberDefinition].Value == null)
                    continue;

                if (Worksheet.Cells[3, e].Value == null)
                {
                    var ex = new ParserException(ParserExceptionType.PREREQUIRE_NOT_FOUND);
                    ex.Cell = Worksheet.Cells[3, e].ToString();
                    ex.SubconditionLine = e;
                    ex.Mode = ParserExceptionMode.SETTER;
                    throw ex;
                }

                var channelName = Worksheet.Cells[3, e].Value.ToString().Replace("\n", "").ToLower();

                listPrerequire.Add(KeywordParser.KeywordFunctionMap[GetFunctionName(channelName)]);
            }

            return listPrerequire;
        }

        private void SetAllPrerequire(FlowChartState state,
                                    string function, string prerequireToSetName, int line, 
                                    bool reset = false, bool setDefaultValue = false)
        {
            var delayAfter = true;

            for (var e = columnNumberFunction + 1; e < columnNumberDefinition; e++)
            {

                if (Worksheet.Cells[line, e].Value == null || (Worksheet.Cells[line, e].Value.ToString() != "Y"
                    && Worksheet.Cells[line, e].Value.ToString() != "N"))
                    continue;

                if (Worksheet.Cells[line, columnNumberDefinition].Value == null)
                    continue;

                if (Worksheet.Cells[3, e].Value == null)
                {
                    var ex = new ParserException(ParserExceptionType.PREREQUIRE_NOT_FOUND);
                    ex.Cell = Worksheet.Cells[3, e].ToString();
                    ex.SubconditionLine = e;
                    ex.Mode = ParserExceptionMode.SETTER;
                    throw ex; 
                }

                var channelName = Worksheet.Cells[3, e].Value.ToString().Replace("\n", "");

                try
                {
                    if (Worksheet.Cells[3, e].Value.ToString().Replace("\n", "") == prerequireToSetName)
                    {
                        //set one Prerequire
                        SetChannel(channelName + KeywordsEnum.ON.ToString(), channelName, state, ((!reset) ? Resources.lang.ChannelActionSetPrerequire : Resources.lang.ChannelActionResetPrerequire) + " " + function, reset, delayAfter, setDefaultValue);
                    }
                    else
                    {
                        //set one Prerequire
                        SetChannel(channelName + KeywordsEnum.OFF.ToString(), channelName, state, ((!reset) ? Resources.lang.ChannelActionSetPrerequire : Resources.lang.ChannelActionResetPrerequire) + " " + function, reset, delayAfter, setDefaultValue);
                    }
                }
                catch (ParserException ex)
                {
                    ex.Cell = Worksheet.Cells[line, e].ToString();
                    ex.Subcondition = channelName;
                    ex.SubconditionLine = 1;
                    ex.Mode = reset? ParserExceptionMode.RESET:ParserExceptionMode.SETTER;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }
                delayAfter = false;
            }
        }

        private void SetGeneralUserAction(FlowChartState state, int columnNumberFunction, int columnNumberDefinition, int line, bool setDefaultValue = false)
        {
            for (var e = columnNumberFunction + 1; e < columnNumberDefinition; e++)
            {
                if (Worksheet.Cells[3, e].Value == null)
                    continue;

                var channelName = Worksheet.Cells[3, e].Value.ToString().ToLower();

                if (channelName != "useraction")
                    continue;

                var regexp = new Regex(@".*start[ ]{0,1}\-[ ]{0,1}(.*).*\n.*end[ ]{0,1}\-(.*).*");

                var value = Worksheet.Cells[line, e].Value;

                if (value == null)
                    continue;

                Match regexWipermatch = regexp.Match(value.ToString().ToLower());

                if (regexWipermatch.Success)
                {
                    if (setDefaultValue)
                    {
                        if (regexWipermatch.Groups[2].Value.Trim() != "")
                        {
                            ProjectUtilsFunction.BuildUserAction(state, regexWipermatch.Groups[2].Value.Trim(), UserActionButtons.Pass | UserActionButtons.Fail);
                        }
                    } else
                    {
                        ProjectUtilsFunction.BuildUserAction(state, regexWipermatch.Groups[1].Value.Trim(), UserActionButtons.Pass | UserActionButtons.Fail);
                    }
                } else
                {
                    throw new ParserException(ParserExceptionType.USERACTION_NOT_WELL_FORMATED);
                }
            }
        }

        public void BuildOutputTest(TestParameters testParameters, int line, int delayTime, string prerequireName,
                                List<string> prerequireList, string cell, ExcelRange noneCell, ActionType actionType, bool checkResultNone = false)
        {
            var delayAfterPrerequire = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterPrerequire"]);
            var delayAfterResetPrerequire = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterResetPrerequire"]);
            var prerequireKey = GetFunctionName(prerequireName.ToLower());
            var preRequireFunctionName = "";
            if (prerequireKey != null)
            {
                preRequireFunctionName = KeywordParser.KeywordFunctionMap[prerequireKey];
            }

            foreach (var conditionInfo in testParameters.Conditions)
            {
                canLoopSignalUsed.Clear();
                var conditionWithComment = conditionInfo.Condition;
                if (conditionWithComment == "" || conditionWithComment.Substring(0, 2) == "//")
                    continue;

                var tmpState = new FlowChartState();
                SetAllPrerequire(tmpState, testParameters.Function, prerequireName, line, false, true);
                if (prerequireName != "NONE")
                {
                    ProjectUtilsFunction.BuildDelay(tmpState, Resources.lang.DelayAfterPrerquireNone + " " + line, delayAfterPrerequire);
                }
                
                testType = KeywordsEnum.NOKEYWORD;
                var tableWithComment = conditionWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                var condition = tableWithComment[0];

                var testInformation = " / " + prerequireName + " / " + conditionInfo.Line + " / " + conditionInfo.Condition;

                var tab = condition.Split(new string[] { "-", "—" }, 2, StringSplitOptions.None);

                if (tab.Count() < 2)
                {
                    var ex = new ParserException(ParserExceptionType.RESULT_VALUE_EXPECTED_NOT_FOUND);
                    ex.Cell = cell;
                    ex.SubconditionLine = conditionInfo.Line;
                    ex.Mode = ParserExceptionMode.RESULT;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }

                var result = tab[0].Trim().ToLower();
                var operation = tab[1].Trim().ToLower();

                operation = BuildOperation(operation, result, testParameters, testInformation, checkResultNone, tmpState);

                var tabOperation = operation.Split('且');


                var validTest = true;

                foreach (var op in tabOperation)
                {
                    validTest = ExecuteOperation(op, result, preRequireFunctionName,
                                            cell, testInformation, testParameters,
                                            tmpState, prerequireList, conditionInfo.Line, validTest);
                }

                if (validTest)
                {
                    string noneValue = KeywordsEnum.OFF.ToString().ToLower();
                    if (noneCell.Value != null && noneCell.Value.ToString().Trim() != "")
                    {
                        var noneTab = noneCell.Value.ToString().Trim().ToLower().Split('\n');

                        if (noneTab.Count() < 2)
                        {
                            noneValue = noneTab[0];
                        } else
                        {
                            noneValue = noneTab[conditionInfo.Line - 1];
                        }
                    }

                    foreach (var action in tmpState.Actions)
                    {
                        testParameters.State.Actions.Add(action);
                    }
                    try
                    {
                        if (actionType == ActionType.OVERLOADRATING)
                        {
                            BuildOverloadRatingTest(testParameters, result, testInformation, checkResultNone);
                        }
                        else
                        {
                            BuildOutputResult(testParameters, delayTime, prerequireName,
                                    prerequireList, cell, noneCell, result, testInformation, noneValue, checkResultNone, operation);
                        }
                    }
                    catch (ParserException ex)
                    {
                        ex.Cell = cell;
                        ex.SubconditionLine = conditionInfo.Line;
                        ex.Mode = ParserExceptionMode.SETTER;
                        ex.Sheet = Worksheet.Name;
                        throw ex;
                    }
                    //reset test
                    CleanCanLoops(testParameters.State, testInformation);
                    foreach (var op in tabOperation)
                    {
                        if (op != "")
                        {
                            try
                            {
                                SetChannel(op, testParameters.Function, testParameters.State, Resources.lang.ChannelActionReset + " " + testParameters.Function + testInformation, true);
                            }
                            catch (ParserException ex)
                            {
                                ex.Cell = cell;
                                ex.Subcondition = op;
                                ex.SubconditionLine = conditionInfo.Line;
                                ex.Mode = ParserExceptionMode.RESET;
                                ex.Sheet = Worksheet.Name;
                                throw ex;
                            }
                        }
                    }
                    //ResetAllPRerequire(testParameters.State, prerequireListAll);
                    //ProjectUtilsFunction.BuildDelay(testParameters.State, Resources.lang.DelayResetPrerequire + " " + line, delayAfterResetPrerequire);
                }

                tmpState = new FlowChartState();
                ResetAllPRerequire(tmpState, prerequireListAll);
                ProjectUtilsFunction.BuildDelay(tmpState, Resources.lang.DelayResetPrerequire + " " + line, delayAfterResetPrerequire);
                if (validTest)
                {
                    foreach (var action in tmpState.Actions)
                    {
                        testParameters.State.Actions.Add(action);
                    }
                }
                //SetAllPrerequire(testParameters.State, "", "NONE", 4, false, true);
            }
        }

        private bool ExecuteOperation(string op, string result, string preRequireFunctionName,
                                        string cell, string testInformation, TestParameters testParameters,
                                        FlowChartState tmpState, List<string> prerequireList, int i, bool validTest)
        {
            if (op != "")
            {
                try
                {
                    /*if (result != KeywordsEnum.OFF.ToString().ToLower())
                    {*/
                        var opKey = GetFunctionName(op);
                        if (opKey != null)
                        {
                            var opName = KeywordParser.KeywordFunctionMap[opKey];

                            if (!IsCanChannel(op) && opName != null && opName == preRequireFunctionName)
                            {
                                var myChannelName = op.Replace(opKey, "");

                                if (myChannelName.Contains(KeywordsEnum.OFF.ToString().ToLower()))
                                {
                                    validTest = false;
                                }
                            }

                            if (validTest == false && !IsCanChannel(op) && opName != null && opName == preRequireFunctionName)
                            {
                                var myChannelName = op.Replace(opKey, "");

                                if (myChannelName.Contains(KeywordsEnum.ON.ToString().ToLower()))
                                {
                                    validTest = true;
                                }
                            }

                            if (prerequireList.Contains(opName) && opName != preRequireFunctionName)
                            {
                                var myChannelName = op.Replace(opKey, "");

                                if (myChannelName.Contains(KeywordsEnum.ON.ToString().ToLower()))
                                {
                                    validTest = false;
                                }
                            }
                        }
                    //}

                    SetChannel(op, testParameters.Function, tmpState, Resources.lang.ChannelActionSetter + " " + testParameters.Function + testInformation);
                }
                catch (ParserException ex)
                {
                    ex.Cell = cell;
                    ex.Subcondition = op;
                    ex.SubconditionLine = i;
                    ex.Mode = ParserExceptionMode.SETTER;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }
            }

            return validTest;
        }

        private string BuildOperation(string operation, string result, TestParameters testParameters,
            string testInformation, bool checkResultNone, FlowChartState tmpState)
        {
            var myOperation = ReplaceKeyword(operation);
            var myResult = ReplaceKeyword(result);
            if (myOperation.Contains(KeywordsEnum.AFTER.ToString().ToLower()))
            {
                operation = GetDelayTypeFromExcel(operation, KeywordsEnum.AFTER);
            }
            else if (myOperation.Contains(KeywordsEnum.DURING.ToString().ToLower()))
            {
                operation = GetDelayTypeFromExcel(operation, KeywordsEnum.DURING);
            }
            else if (myOperation.Contains(KeywordsEnum.HIGHLOW.ToString().ToLower()))
            {
                operation = GetDelayHighLowFromExcel(operation);
                BuildOutputOperationHighLow(testParameters.Function, result, tmpState, testInformation, false, true);
            }
            else if (myOperation.Contains(KeywordsEnum.HIGH.ToString().ToLower()))
            {
                operation = GetDelayTypeFromExcel(operation, KeywordsEnum.HIGH);
                BuildCaptureHighOrLowOutputTest(testParameters.Function, result, tmpState, testInformation, false, false);
            }
            else if (myOperation.Contains(KeywordsEnum.LOW.ToString().ToLower()))
            {
                operation = GetDelayTypeFromExcel(operation, KeywordsEnum.LOW);
                BuildCaptureHighOrLowOutputTest(testParameters.Function, result, tmpState, testInformation, false, true);
            } else if (myResult.Contains(KeywordsEnum.WIPER_WASH.ToString().ToLower()) && !checkResultNone)
            {
                testType = KeywordsEnum.WIPER_WASH;

                var channelOutputInfo = GetChannelOutputInfo(testParameters.Function);

                //Capture
                var channelActionCaptureStart = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + CaptureType.HIGH + " " + Resources.lang.START + " " + testParameters.Function + testInformation, channelOutputInfo.Channel, CaptureType.HIGH, true, channelOutputInfo?.Board);
                channelActionCaptureStart.Threshold1 = 0.8f * (float)InputParser.VBat;
                tmpState.Actions.Add(channelActionCaptureStart);
            } else if (myResult.Contains(KeywordsEnum.WIPER_INTERVAL.ToString().ToLower()) && !checkResultNone)
            {
                testType = KeywordsEnum.WIPER_INTERVAL;

                var channelOutputInfo = GetChannelOutputInfo(testParameters.Function);

                //Capture
                var channelActionCaptureStart = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + CaptureType.HIGH + " " + Resources.lang.START + " " + testParameters.Function + testInformation, channelOutputInfo.Channel, CaptureType.HIGH, true, channelOutputInfo?.Board);
                channelActionCaptureStart.Threshold1 = 0.8f * (float)InputParser.VBat;
                tmpState.Actions.Add(channelActionCaptureStart);
            }

            return operation;
        }

        private string GetDelayHighLowFromExcel(string operation)
        {
            testType = KeywordsEnum.HIGHLOW;
            var indexOfKeyword = operation.IndexOf(testType.ToString().ToLower());

            var regexp = new Regex(testType.ToString().ToLower() + @"\((\d+[.]?\d*),(.*)\)");
            var match = regexp.Match(operation);

            limitValueHigh = float.Parse(match.Groups[1].Value);
            try
            {
                limitValueLow = float.Parse(match.Groups[2].Value);
            }
            catch
            {
                limitValueLow = -1;
            }

            operation = operation.Remove(indexOfKeyword);

            return operation;
        }

        private void BuildOutputOperationHighLow(string function, string result, FlowChartState state, string testInformation, bool checkResultFalse, bool start)
        {
            var channelInfo = GetChannelOutputInfo(function);
            var gap = 700;

            if (channelInfo == null)
            {
                throw new ParserException(ParserExceptionType.RESULT_CHANNEL_NOT_FOUND);
            }

            var channelActionCapture = ProjectUtilsFunction.BuildCapture(((start)?Resources.lang.ChannelActionCaptureHighLowStart: Resources.lang.ChannelActionCaptureHighLowStop) + " " + testInformation, channelInfo.Channel, CaptureType.HIGHLOW, start, channelInfo.Board);
            channelActionCapture.Threshold1 = 0.8f * (float)InputParser.VBat;
            channelActionCapture.Threshold2 = 0.8f * (float)InputParser.VBat;

            var limitInfo = GetLimitInfo(channelInfo.Channel, function, result);
            Limit limitHigh;
            Limit limitLow;

            if (checkResultFalse && result == "off")
            {
                limitHigh = BuildLimit(limitInfo, "=" + 0, channelInfo, true, (decimal)limitValueHigh);
                if (limitValueLow < 0)
                {
                    limitLow = BuildLimit(limitInfo, "=0", channelInfo, true, (decimal)2000);
                }
                else
                {
                    limitLow = BuildLimit(limitInfo, "=" + 0, channelInfo, true, (decimal)limitValueLow);
                }
            }
            else
            {
                limitHigh = BuildLimit(limitInfo, "=" + (limitValueHigh + gap), channelInfo, true, (decimal)limitValueHigh);
                if (limitValueLow < 0)
                {
                    limitLow = BuildLimit(limitInfo, ">2000", channelInfo, true, (decimal)2000);
                }
                else
                {
                    limitLow = BuildLimit(limitInfo, "=" + (limitValueLow + gap), channelInfo, true, (decimal)limitValueLow);
                }
            }

            channelActionCapture.Limits.Add(limitHigh);
            channelActionCapture.ExtraLimits.Add(limitLow);

            state.Actions.Add(channelActionCapture);
        }

        private void BuildOutputResult(TestParameters testParameters, int delayTime, string prerequireName,
                                List<string> prerequireList, string cell, ExcelRange noneCell, string result, string testInformation, 
                                string noneValue, bool checkResultNone, string operation)
        {
            var delayAfterInput = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterInput"]);

            if (checkResultNone && !result.Contains("beep") && !result.Contains("video"))
            {
                result = noneValue;
            }
            var myResult = ReplaceKeyword(result);
            if ((testType == KeywordsEnum.AFTER || (testType == KeywordsEnum.DURING && !checkResultNone)) && !result.Contains("beep") && !result.Contains("video"))
            {
                BuildCaptureMinOrMaxOutputTest(testParameters.Function, result, testParameters.State, testInformation, checkResultNone, testType == KeywordsEnum.DURING);
            }
            else if (myResult.Contains(KeywordsEnum.FLASH.ToString().ToLower()))
            {
                BuildFlashOutputOperationHighLow(testParameters.Function, result, testParameters.State, testInformation, checkResultNone);

                var regExp = @"(\d+[.]?\d*)s: on, (\d+[.]?\d*)s: off";

                Match match = Regex.Match(result, regExp);

                var ONTime = decimal.Parse(match.Groups[1].Value);
                var OFFTime = decimal.Parse(match.Groups[2].Value);
                BuildFlashOutputOperationMax(testParameters.Function, result, testParameters.State, testInformation, checkResultNone);
            }
            else if (myResult.Contains(KeywordsEnum.AVG.ToString().ToLower()) && !result.Contains("beep") && !result.Contains("video"))
            {
                BuildAvgOutput(testParameters.Function, result, testParameters.State, testInformation, checkResultNone);
            }
            else if (myResult.Contains(KeywordsEnum.WIPER_WASH.ToString().ToLower()) && !result.Contains("beep") && !result.Contains("video"))
            {
                BuildWiperWashOutput(testParameters.Function, result, testParameters.State, testInformation);
            }
            else if (myResult.Contains(KeywordsEnum.WIPER_INTERVAL.ToString().ToLower()) && !result.Contains("beep") && !result.Contains("video"))
            {
                BuildWiperIntervalOutput(testParameters.Function, result, testParameters.State, testInformation);
            } else if (testType == KeywordsEnum.HIGHLOW)
            {
                ProjectUtilsFunction.BuildDelay(testParameters.State, Resources.lang.DelaySartCapture + " " + testType + " " + Resources.lang.AndStopCapture + " " + testType + " " + testParameters.Function + testInformation, 2000);
                BuildOutputOperationHighLow(testParameters.Function, result, testParameters.State, testInformation, checkResultNone, false);
            }
            else if ((testType == KeywordsEnum.HIGH || testType == KeywordsEnum.LOW) && !result.Contains("beep") && !result.Contains("video"))
            {
                long delayTmp = (long)(specificTime * 1.25);
                ProjectUtilsFunction.BuildDelay(testParameters.State, Resources.lang.DelaySartCapture + " " + testType + " " + Resources.lang.AndStopCapture + " " + testType + " " + testParameters.Function + testInformation, (delayTmp > 5000)? delayTmp :5000);
                BuildCaptureHighOrLowOutputTest(testParameters.Function, result, testParameters.State, testInformation, checkResultNone, testType == KeywordsEnum.LOW, false);
            }
            else
            {
                BuildStandardOutputTest(testParameters, result, delayTime, testInformation, checkResultNone, noneValue, operation);
            }
        }

        private void BuildWiperIntervalOutput(string function, string result, FlowChartState state, string testInformation)
        {
            int OccurrenceNumber = 3;
            var actionDelayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);
            var captureType = CaptureType.HIGH;

            var regexpWiper = new Regex(@".*wiper_interval\((.*)\).*");

            Match regexWipermatch = regexpWiper.Match(result);

            if (!regexWipermatch.Success)
            {
                throw new ParserException(ParserExceptionType.WIPER_WASH_ISSUE);
            }

            var channelInputName = regexWipermatch.Groups[1];
            var channelOutputInfo = GetChannelOutputInfo(function);

            for (var i = 0; i < OccurrenceNumber; i++)
            {
                //delay
                if (i == 0)
                {
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperInterval + " 100ms " + function + testInformation, 100 - actionDelayTime);
                } else
                {
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperInterval + " 1850ms " + function + testInformation, 1850 - actionDelayTime);
                }
                SetChannel(channelInputName + "ON", function, state, Resources.lang.SetWiper + " ON");
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperInterval + " 1300ms " + function + testInformation, 1300 - actionDelayTime);
                SetChannel(channelInputName + "OFF", function, state, Resources.lang.SetWiper + " OFF");
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperInterval + " 1850ms " + function + testInformation, 1850 - actionDelayTime);

                var limitInfo = GetLimitInfo(channelOutputInfo.Channel, function, result);
                var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.STOP + " " + function + testInformation, channelOutputInfo.Channel, captureType, false, channelOutputInfo?.Board);
                channelActionCaptureStop.Threshold1 = 0.8f * (float)InputParser.VBat;
                var limit = BuildLimit(limitInfo, "=" + 1500, channelOutputInfo, true, 1500);
                channelActionCaptureStop.Limits.Add(limit);
                state.Actions.Add(channelActionCaptureStop);

                if (i < 2)
                {
                    var channelActionCaptureStart = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.START + " " + function + testInformation, channelOutputInfo.Channel, captureType, true, channelOutputInfo?.Board);
                    channelActionCaptureStart.Threshold1 = 0.8f * (float)InputParser.VBat;
                    state.Actions.Add(channelActionCaptureStart);
                }
            }
        }

        private void BuildWiperWashOutput(string function, string result, FlowChartState state, string testInformation)
        {
            var captureType = CaptureType.HIGH;
            var channelOutputInfo = GetChannelOutputInfo(function);
            var actionDelayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);

            var regexpWiper = new Regex(@".*wiper_wash\((.*)\).*");

            Match regexWipermatch = regexpWiper.Match(result);

            if (!regexWipermatch.Success)
            {
                throw new ParserException(ParserExceptionType.WIPER_WASH_ISSUE);
            }

            var channelInputName = regexWipermatch.Groups[1];

            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 400ms " + function + testInformation, 400 - actionDelayTime);
            SetChannel(channelInputName + "ON", function, state, "SET wiper ON");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 1300ms " + function + testInformation, 1300 - actionDelayTime);
            SetChannel(channelInputName + "OFF", function, state, "SET wiper OFF");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 200ms " + function + testInformation, 200 - actionDelayTime);
            SetChannel(channelInputName + "ON", function, state, "SET wiper ON");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 1300ms " + function + testInformation, 1300 - actionDelayTime);
            SetChannel(channelInputName + "OFF", function, state, "SET wiper OFF");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 200ms " + function + testInformation, 200 - actionDelayTime);
            SetChannel(channelInputName + "ON", function, state, "SET wiper ON");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWash + " 1300ms " + function + testInformation, 1300 - actionDelayTime);
            SetChannel(channelInputName + "OFF", function, state, "SET wiper OFF");
            //delay
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayWiperWashBeforeGettingResult + " " + function + testInformation, 1000 - actionDelayTime);

            var limitInfo = GetLimitInfo(channelOutputInfo.Channel, function, result);
            var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.STOP +" " + function + testInformation, channelOutputInfo.Channel, captureType, false, channelOutputInfo?.Board);
            channelActionCaptureStop.Threshold1 = 0.8f * (float)InputParser.VBat;
            var limit = BuildLimit(limitInfo, "=" + 4500, channelOutputInfo, true, 4500);

            channelActionCaptureStop.Limits.Add(limit);
            state.Actions.Add(channelActionCaptureStop);
        }

        private void BuildAvgOutput(string function, string result, FlowChartState state, string testInformation, bool checkResultFalse)
        {
            var canChannel = GetChannelCanMessage(function);
            var channelInfo = GetChannelOutputInfo(function);
            var captureType = CaptureType.AVERAGE;

            result = result.Replace(KeywordsEnum.AVG.ToString().ToLower(), "");

            if (channelInfo == null && canChannel == null)
            {
                throw new ParserException(ParserExceptionType.CHANNEL_NOT_FOUND);
            }
            var channel = channelInfo?.Channel;
            if (channel == null)
            {
                channel = canChannel;
            }

            //Capture
            var channelActionCapture = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.START +" " + function + testInformation, channel, captureType, true, channelInfo?.Board);
            state.Actions.Add(channelActionCapture);
            //delai
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelaySartCapture + " " + captureType + " " + Resources.lang.AndStopCapture + " " + captureType + " " + function + testInformation, 2000);
            var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " STOP " + function + testInformation, channel, captureType, false, channelInfo?.Board);

            var limitInfo = GetLimitInfo(channel, function, result);
            Limit limit = null;
            if (checkResultFalse)
            {
                limit = BuildLimit(limitInfo, function + KeywordsEnum.OFF.ToString().ToLower(), channelInfo);
            }
            else
            {
                limit = BuildLimit(limitInfo, function + result, channelInfo);
            }

            channelActionCaptureStop.Limits.Add(limit);
            state.Actions.Add(channelActionCaptureStop);
        }

        public void BuildCaptureHighOrLowOutputTest(string function, string result, FlowChartState state, string testInformation, bool checkResultFalse, bool isLow, bool startCapture = true)
        {
            var channelInfo = GetChannelOutputInfo(function);

            if (channelInfo == null)
            {
                throw new ParserException(ParserExceptionType.RESULT_CHANNEL_NOT_FOUND);
            }

            var captureType = CaptureType.HIGH;
            var channel = channelInfo.Channel;

            if (isLow)
            {
                captureType = CaptureType.LOW;
            }

            //Capture

            var channelActionCapture = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + ((startCapture)?" " + Resources.lang.START + " ": " " + Resources.lang.STOP + " ") + function + testInformation, channel, captureType, startCapture, channelInfo?.Board);
            
            var limitInfo = GetLimitInfo(channel, function, result);

            if (captureType == CaptureType.HIGH) {
                channelActionCapture.Threshold1 = 0.8f * (float)InputParser.VBat;
            } else
            {
                if (channelInfo.OffValue == 0)
                {
                    channelActionCapture.Threshold1 = 0.2f * (float)InputParser.VBat;
                }
                else
                {
                    channelActionCapture.Threshold1 = (float)channelInfo.OffValue + 1.5f;
                }
            }

            Limit limit;

            if (checkResultFalse && result == "off")
            {
                limit = BuildLimit(limitInfo, "=" + 0, channelInfo);
            } else
            {
                limit = BuildLimit(limitInfo, "=" + specificTime, channelInfo, true, (decimal)specificTime);
            }

            channelActionCapture.Limits.Add(limit);

            state.Actions.Add(channelActionCapture);
        }

        public void BuildCaptureMinOrMaxOutputTest(string function, string result, FlowChartState state, string testInformation, bool checkResultFalse, bool isDuring)
        {
            var channelInfo = GetChannelOutputInfo(function);
            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterInput"]);

            if (channelInfo == null)
            {
                throw new ParserException(ParserExceptionType.RESULT_CHANNEL_NOT_FOUND);
            }
            var captureType = CaptureType.MAX;
            var channel = channelInfo.Channel;

            if ((!checkResultFalse && isDuring) || (!checkResultFalse && !isDuring && result == KeywordsEnum.OFF.ToString().ToLower()))
            {
                captureType = CaptureType.MIN;
            }

            if (isDuring)
            {
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBeforeStartCapture + " " + captureType + " " + function + testInformation, (int)((specificTime - ((isDuring) ? (250 + delayTime) : 0)) * 0.1));
            }

                //Capture
                var channelActionCapture = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.START + " " + function + testInformation, channel, captureType, true, channelInfo?.Board);
            state.Actions.Add(channelActionCapture);

            //delai
            if (isDuring)
            {
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelaySartCapture + " " + captureType + " " + Resources.lang.AndStopCapture + " " + captureType + " " + function + testInformation, (int)((specificTime - ((isDuring) ? (250 + delayTime) : 0)) * 0.7));
            } else
            {
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelaySartCapture + " " + captureType + " " + Resources.lang.AndStopCapture + " " + captureType + " " + function + testInformation, (int)(specificTime * 0.9 - 2000));
            }
            var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCapture + " " + captureType + " " + Resources.lang.STOP  + " " + function + testInformation, channel, captureType, false, channelInfo?.Board);

            var limitInfo = GetLimitInfo(channel, function, result);
            Limit limit = null;
            if (checkResultFalse)
            {
                limit = BuildLimit(limitInfo, KeywordsEnum.OFF.ToString().ToLower(), channelInfo);
            }
            else if (!isDuring && result == KeywordsEnum.OFF.ToString().ToLower())
            {
                limit = BuildLimit(limitInfo, "=" + InputParser.VBat, channelInfo);
            }
            else
            {
                if (isDuring)
                {
                    limit = BuildLimit(limitInfo, result, channelInfo);
                }
                else
                {
                    limit = BuildLimit(limitInfo, KeywordsEnum.OFF.ToString().ToLower(), channelInfo);
                }
            }

            channelActionCaptureStop.Limits.Add(limit);
            state.Actions.Add(channelActionCaptureStop);

            //delai
            if (isDuring)
            {
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBetweenStopCapture + " " + captureType + " " + Resources.lang.AndGetResult + function + testInformation, (int)((specificTime * 0.3) + ((isDuring) ? (250 - delayTime) : 0)));
            } else
            {
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBetweenStopCapture + " " + captureType + " " + Resources.lang.AndGetResult + function + testInformation, (int)(specificTime * 0.3 + 2000));
            }
            //get
            var channelActionResultAfterCapture = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionAfterCapture + " " + function + testInformation);

            if (checkResultFalse)
            {
                SetResultChannel(function + KeywordsEnum.OFF.ToString().ToLower(), function, channelActionResultAfterCapture);
            }
            else
            {
                if (isDuring)
                {
                    SetResultChannel(function + KeywordsEnum.OFF.ToString().ToLower(), function, channelActionResultAfterCapture);
                }
                else
                {
                    SetResultChannel(function + result, function, channelActionResultAfterCapture);
                }
            }
        }

        public void BuildStandardOutputTest(TestParameters testParameters, string result, int delayTime, string testInformation, 
            bool checkResultNone, string noneValue, string operation)
        {
            //add delay
            ProjectUtilsFunction.BuildDelay(testParameters.State, Resources.lang.DelayBetweenSetAndResultOf + " " + testParameters.Function + testInformation, delayTime);

            var regexpUserAction = new Regex(@".*useraction\((.*)\).*");

            Match matchUserAction = regexpUserAction.Match(result);

            if (matchUserAction.Success)
            {
                ProjectUtilsFunction.BuildUserAction(testParameters.State, matchUserAction.Groups[1] +
                    testInformation, UserActionButtons.Pass | UserActionButtons.Fail);
            }
            else if (!result.Contains("beep") && !result.Contains("video"))
            {
                var channelActionResult = ProjectUtilsFunction.BuildChannelAction(testParameters.State, Resources.lang.ChannelActionResult + " " + testParameters.Function + testInformation);

                if (checkResultNone)
                {
                    SetResultChannel(testParameters.Function + noneValue, testParameters.Function, channelActionResult);
                }
                else
                {
                    SetResultChannel(testParameters.Function + result, testParameters.Function, channelActionResult);
                }
            }
            else
            {
                operation = BuildOperation(operation, result, out result);

                var action = (result.Contains("beep")) ? Resources.lang.Hear : Resources.lang.See;
                if (!checkResultNone)
                {
                    ProjectUtilsFunction.BuildUserAction(testParameters.State, Resources.lang.YouShould + " " + action + " " + result + " ? " +
                    testParameters.Function + testInformation, UserActionButtons.Pass | UserActionButtons.Fail, null, null, testType == KeywordsEnum.EVERY);
                } else
                {
                    ProjectUtilsFunction.BuildUserAction(testParameters.State, Resources.lang.YouShouldNot + " " + action + " " + result + " ? " +
                    testParameters.Function + testInformation, UserActionButtons.Pass | UserActionButtons.Fail, null, null, testType == KeywordsEnum.EVERY);
                }
            }
        }

        public void ParseOverloadOutput()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.OutputsSheet];

            columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsDefinitionColName, Worksheet);
            columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsFunctionColName, Worksheet);
            colNumberUsed = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.OutputsUsedColName, Worksheet);

            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);
            var delayAfterPrerequire = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterPrerequire"]);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            SetAllPrerequire(state, "", "NONE", 4, true);
            for (var i = 4; Worksheet.Cells[i, colNumberUsed].Value != null; i++)
            {
                try
                {
                    if (Worksheet.Cells[i, colNumberUsed].Value.ToString() != "Y")
                    {
                        continue;
                    }

                    var prerequireList = GetAllActivePrerequire(i);

                    var function = Worksheet.Cells[i, columnNumberFunction].Value.ToString().ToLower();

                    if (Worksheet.Cells[i, columnNumberDefinition].Value == null)
                    {
                        var ex = new ParserException(ParserExceptionType.TEST_DEFINITION_NOT_FOUND);

                        ex.Cell = Worksheet.Cells[i, columnNumberFunction].ToString();
                        ex.SubconditionLine = i;
                        ex.Mode = ParserExceptionMode.SETTER;
                        ex.Sheet = Worksheet.Name;
                        throw ex;
                    }

                    var TestDefinition = Worksheet.Cells[i, columnNumberDefinition].Value.ToString();

                    var conditions = TestDefinition.Split('\n').ToList();

                    //dynamic test
                    for (var e = columnNumberFunction + 1; e < columnNumberDefinition; e++)
                    {
                        var overloadAvailabe = false;
                        foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
                        {
                            if (function.Contains(key) && OverloadParser.OverloadFunctionRelays.ContainsKey(key))
                            {
                                overloadAvailabe = true;
                            }
                        }

                        if (!overloadAvailabe)
                            continue;

                        if (Worksheet.Cells[i, e].Value == null || Worksheet.Cells[i, e].Value.ToString() != "Y")
                            continue;

                        if (Worksheet.Cells[i, columnNumberDefinition].Value == null)
                            continue;

                        var channelName = Worksheet.Cells[3, e].Value.ToString().Replace("\n", "");

                        SetAllPrerequire(state, function, channelName, i);
                        ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayAfterPrerequire + " " + channelName + " at line " + i, delayAfterPrerequire);
                        BuildOverloadTestList(conditions, state, function, delayTime, channelName, prerequireList, Worksheet.Cells[i, columnNumberDefinition].ToString());
                        //reset prerequire
                        /*SetAllPrerequire(Worksheet, state,
                            columnNumberFunction, columnNumberDefinition, function, channelName, i);*/
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
            var prerequireListAll = GetAllActivePrerequire(4, true);
            ResetAllPRerequire(state, prerequireListAll);
        }

        public void BuildOverloadTestList(List<string> conditions, FlowChartState state, string function, int delayTime, string prerequireName, List<string> prerequireList, string cell, bool checkResultFalse = false)
        {
            var i = 1;
            var prerequireKey = GetFunctionName(prerequireName.ToLower());
            var preRequireFunctionName = "";
            if (prerequireKey != null)
            {
                preRequireFunctionName = KeywordParser.KeywordFunctionMap[prerequireKey];
            }

            foreach (var conditionWithComment in conditions)
            {
                canLoopSignalUsed.Clear();
                testType = KeywordsEnum.NOKEYWORD;

                if (conditionWithComment == "" || conditionWithComment.Substring(0, 2) == "//")
                    continue;

                var tableWithComment = conditionWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                var condition = tableWithComment[0];

                var testInformation = " / " + prerequireName + " / " + i + " / " + conditionWithComment;

                //channelaction
                var tab = condition.Split(new string[] { "-", "—" }, 2, StringSplitOptions.None);

                if (tab.Count() < 2)
                {
                    var ex = new ParserException(ParserExceptionType.RESULT_VALUE_EXPECTED_NOT_FOUND);
                    ex.Cell = cell;
                    ex.SubconditionLine = i;
                    ex.Mode = ParserExceptionMode.RESULT;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }

                var result = tab[0].Trim().ToLower();
                var operation = tab[1].Trim().ToLower();

                var myOperation = ReplaceKeyword(operation);
                if (myOperation.Contains(KeywordsEnum.AFTER.ToString().ToLower()))
                {
                    operation = GetDelayTypeFromExcel(operation, KeywordsEnum.AFTER);
                }
                else if (myOperation.Contains(KeywordsEnum.DURING.ToString().ToLower()))
                {
                    var indexOfKeyword = operation.IndexOf(KeywordsEnum.DURING.ToString().ToLower());
                    operation = operation.Remove(indexOfKeyword);
                } else if (myOperation.Contains(KeywordsEnum.HIGH.ToString().ToLower()))
                {

                }

                var tabOperation = operation.Split('且');
                var validTest = true;
                var tmpState = new FlowChartState();

                foreach (var op in tabOperation)
                {
                    if (op != "")
                    {
                        try
                        {
                            if (result != KeywordsEnum.OFF.ToString().ToLower())
                            {
                                var opKey = GetFunctionName(op);
                                if (opKey != null)
                                {
                                    var opName = KeywordParser.KeywordFunctionMap[opKey];

                                    if (!IsCanChannel(op) && opName != null && opName == preRequireFunctionName)
                                    {
                                        var myChannelName = op.Replace(opKey, "");

                                        if (myChannelName.Contains(KeywordsEnum.OFF.ToString().ToLower()))
                                        {
                                            validTest = false;
                                            break;
                                        }
                                    }

                                    if (prerequireList.Contains(opName) && opName != preRequireFunctionName)
                                    {
                                        var myChannelName = op.Replace(opKey, "");

                                        if (myChannelName.Contains(KeywordsEnum.ON.ToString().ToLower()))
                                        {
                                            validTest = false;
                                            break;
                                        }
                                    }
                                }
                            }

                            SetChannel(op, function, tmpState, Resources.lang.ChannelActionSetter + " " + function + testInformation);
                        } catch (ParserException ex)
                        {
                            ex.Cell = cell;
                            ex.Subcondition = op;
                            ex.SubconditionLine = i;
                            ex.Mode = ParserExceptionMode.SETTER;
                            ex.Sheet = Worksheet.Name;
                            throw ex;
                        }
                    }
                }

                if (validTest)
                {
                    foreach (var action in tmpState.Actions)
                    {
                        state.Actions.Add(action);
                    }
                    if (testType == KeywordsEnum.DURING || testType == KeywordsEnum.AFTER)
                    {
                        BuildOverloadCapture(state, function + testInformation);
                    }

                    foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
                    {
                        if (function.Contains(key))
                        {
                            BuildOverloadMaxTest(result, function, key, function + testInformation, checkResultFalse);
                            break;
                        }
                    }

                    //reset test
                    foreach (var op in tabOperation)
                    {
                        if (op != "")
                        {
                            try
                            {
                                SetChannel(op, function, state, Resources.lang.ChannelActionReset + " " + function + testInformation, true);
                            }
                            catch (ParserException ex)
                            {
                                ex.Cell = cell;
                                ex.Subcondition = op;
                                ex.SubconditionLine = i + 1;
                                ex.Mode = ParserExceptionMode.RESET;
                                ex.Sheet = Worksheet.Name;
                                throw ex;
                            }
                        }
                    }
                }
                i++;
            }
        }

        public void StartOverloadTest(string key, FlowChartState state, string testInfo)
        {
            //put overloadmode
            var channelActionOverloadModeStart = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadModeStart + " " + testInfo);
            var relay = OverloadParser.OverloadFunctionRelays[key];
            channelActionOverloadModeStart.ChannelActions.Add(OverloadParser.BuildRelayAction(relay, RelayState.CLOSE, product));
        }

        public void StopOverloadTest(string key, FlowChartState state, string testInfo)
        {
            var relay = OverloadParser.OverloadFunctionRelays[key];
            //put overloadmode
            var channelActionOverloadModeStop = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadModeStop + " " + testInfo);
            channelActionOverloadModeStop.ChannelActions.Add(OverloadParser.BuildRelayAction(relay, RelayState.OPEN, product));
        }

        public void BuildOverloadRatingTest(TestParameters testParameters, string result, string testInfo, bool checkResultFalse)
        {
            string key = "";

            foreach (var keyString in KeywordParser.KeywordFunctionMap.Keys)
            {
                if (testParameters.Function.Contains(key))
                {
                    key = keyString;
                    break;
                }
            }

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
            StartOverloadTest(key, state, testInfo);
  
            //setup test
            var ChannelInfos = KeywordParser.FunctionChannelOutputMap.Where(ch => ch.Function == key);
            var ChannelInfo = ChannelInfos.FirstOrDefault();

            //do test
            ChannelOutputInfo RelayDefaultInfo = OverloadParser.GetLoadChannel(ChannelInfo);

            var defaultTestInfo = testInfo + " / " + Resources.lang.Load + " == " + RelayDefaultInfo.Load;

            var channelActionOverloadDefaultTestSetup = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadRatingTestSetup + " " + defaultTestInfo);
            channelActionOverloadDefaultTestSetup.ChannelActions.Add(OverloadParser.BuildRelayAction(RelayDefaultInfo, RelayState.CLOSE, product));

            var myResult = ReplaceKeyword(result);
            if (myResult.Contains("flash"))
            {
                ChannelOutputInfo voltmeterChannelInfo;
                if (ChannelInfo.Product == "multic s")
                {
                    voltmeterChannelInfo = OverloadParser.OverloadClusterVoltmeter;
                }
                else
                {
                    var productName = ChannelInfo.Product;
                    var productPos = int.Parse(productName.Replace("bcm ", ""));
                    voltmeterChannelInfo = OverloadParser.OverloadBCMVoltmeter[productPos];
                }

                BuildFlashOutputOperationHighLow(testParameters.Function, result, state, " " + Resources.lang.Rating + " " + defaultTestInfo, checkResultFalse, voltmeterChannelInfo);
                BuildFlashOutputOperationMax(testParameters.Function, result, state, " " + Resources.lang.Rating + " " + defaultTestInfo, checkResultFalse, voltmeterChannelInfo);
            }
            else
            {
                //delay 1s
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.ChannelActionOverloadDelay + " " + defaultTestInfo, 2000);

                //getvolt
                var channelActionOverloadDefaultTest = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadRatingTest + " " + defaultTestInfo);
                channelActionOverloadDefaultTest.ChannelActions.Add(OverloadParser.BuildGetVoltAction(ChannelInfo, product));

                /*BuildCaptureHighOrLowOutputTest(function, result, state, " " + Resources.lang.Rating + " " + defaultTestInfo, checkResultFalse, false, true);
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.ChannelActionOverloadDelay + " " + defaultTestInfo, captureTime);
                BuildCaptureHighOrLowOutputTest(function, result, state, " " + Resources.lang.Rating + " " + defaultTestInfo, checkResultFalse, false, false);*/

            }
            var channelActionOverloadDefaultTestReset = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadRatingTestReset + " " + defaultTestInfo);
            channelActionOverloadDefaultTestReset.ChannelActions.Add(OverloadParser.BuildRelayAction(RelayDefaultInfo, RelayState.OPEN, product));

            StopOverloadTest(key, state, testInfo);
        }

        public void BuildOverloadMaxTest(string result, string function, string key, string testInfo, bool checkResultFalse)
        {
            //OverloadBMSSecondPartUsed = false;
            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
            StartOverloadTest(key, state, testInfo);
            var captureTime = int.Parse(IniFile.IniDataRaw["Overload"]["CaptureTime"]);
            var ChannelInfos = KeywordParser.FunctionChannelOutputMap.Where(ch => ch.Function == key);
            var ChannelInfo = ChannelInfos.FirstOrDefault();
            /*var userAction = new UserInputAction();
            userAction.Description = "test normal overload";
            userAction.Buttons = UserActionButtons.Pass | UserActionButtons.Fail;
            state.Actions.Add(userAction);*/
            specificTime = captureTime;
            //set up max test
            float maxLoad = ChannelInfo.OverCurrent;

            //do test
            ChannelOutputInfo RelayMaxInfo = OverloadParser.GetLoadChannel(ChannelInfo, maxLoad);

            var maxTestInfo = testInfo + " / Load == " + RelayMaxInfo.Load;

            var channelActionOverloadMaxTestSetup = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadMaxTestSetup + " " + maxTestInfo);
            channelActionOverloadMaxTestSetup.ChannelActions.Add(OverloadParser.BuildRelayAction(RelayMaxInfo, RelayState.CLOSE, product));

            var myResult = ReplaceKeyword(result);
            if (myResult.Contains("flash"))
            {
                ChannelOutputInfo voltmeterChannelInfo;
                if (ChannelInfo.Product == "multic s")
                {
                    voltmeterChannelInfo = OverloadParser.OverloadClusterVoltmeter;
                }
                else
                {
                    var productName = ChannelInfo.Product;
                    var productPos = int.Parse(productName.Replace("bcm ", ""));
                    voltmeterChannelInfo = OverloadParser.OverloadBCMVoltmeter[productPos];
                }

                BuildFlashOutputOperationHighLow(function, result, state, " " + Resources.lang.Max + " " + maxTestInfo, checkResultFalse, voltmeterChannelInfo);
                BuildFlashOutputOperationMax(function, result, state, " " + Resources.lang.Max + " " + maxTestInfo, checkResultFalse, voltmeterChannelInfo);
            }
            else
            {

                //delay 1s
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.ChannelActionOverloadDelay + " " + maxTestInfo, 2000);

                //getvolt
                /*var channelActionOverloadMaxTest = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadMaxTest + " " + maxTestInfo);
                channelActionOverloadMaxTest.ChannelActions.Add(OverloadParser.BuildGetVoltAction(ChannelInfo, product));*/

                BuildCaptureHighOrLowOutputTest(function, result, state, " " + Resources.lang.Max + " " + maxTestInfo, checkResultFalse, false, true);
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.ChannelActionOverloadDelay + " " + maxTestInfo, captureTime);
                BuildCaptureHighOrLowOutputTest(function, result, state, " " + Resources.lang.Max + " " + maxTestInfo, checkResultFalse, false, false);
            }

            var channelActionOverloadMaxTestReset = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionOverloadMaxTestReset + " " + maxTestInfo);
            channelActionOverloadMaxTestReset.ChannelActions.Add(OverloadParser.BuildRelayAction(RelayMaxInfo, RelayState.OPEN, product));

            StopOverloadTest(key, state, testInfo);

            /*var userActionmax = new UserInputAction();
            userActionmax.Description = "test max overload";
            userActionmax.Buttons = UserActionButtons.Pass | UserActionButtons.Fail;
            state.Actions.Add(userActionmax);*/
        }

        private void BuildOverloadCapture(FlowChartState state, string testInfo)
        {
            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterInput"]);

            //delay for now maybe need capture alter
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayForOverloadCaptureAfter + " " + testInfo, (int)(specificTime * 1.1));
        }
    }
}
