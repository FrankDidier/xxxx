using ActiaParser.ActionParser;
using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Actions.Conditions;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Boards.Resources.ChannelSettings;
using ArtLogics.TestSuite.DevXlate.Resources.Bus;
using ArtLogics.TestSuite.Environment.Dbc;
using ArtLogics.TestSuite.Environment.GlobalVariables;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.Limits.Comparisons.Text;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.GlobalVariables;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.TestSuite.TestResults;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Utils;
using DevExpress.XtraEditors;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using static ArtLogics.TestSuite.Boards.Resources.ChannelSettings.CanSettings;
using static ArtLogics.TestSuite.Shared.Services.Data.BoardService;

namespace ActiaParser.MessageParser
{
    public class LimitSignalInfo
    {
        public decimal A { get; set; }
        public decimal B { get; set; }
        public int StartBit { get; set; }
        public int BitLength { get; set; }
        public string CanId { get; set; }
        public MessageFlowDirection Direction { get; set; }
        public string MessageName { get; set; }
        public Variable Variable { get; set; }
        public Board Board { get; internal set; }
        public Channel Channel { get; internal set; }
        public int Frequency { get; internal set; }
        public List<LimitSignalInfo> ForwardList { get; set; }
        public string DefaultValue { get; set; } = "0X0000000000000000";
        public string Test { get; set; } = null;
        public string SignalName { get; set; }
        public bool ResetLoop { get; set; } = false;
        public int Priority { get; set; }
        public string Endian { get; set; }
        public int FormulaLogic { get; set; }
        public int Line { get; internal set; }
        public string ChannelName { get; set; }
        public string Algoritm { get; set; }
    }

    public class MessageParser : ActionParserBase
    {
        private BindingList<ArtLogics.Translation.Parser.Model.Resources> resources;
        private BindingList<string> MessageName;

        private Dictionary<int, string> CanColumnMap { get; set; }
        public static List<LimitSignalInfo> LimitSignalInfoList { get; set; }

        public static Dictionary<string, string> NumberToLetter = new Dictionary<string, string>()
        {
            { "1", "A" },
            { "2", "B" },
            { "3", "C" },
            { "4", "D" },
            { "5", "E" },
            { "6", "F" },
            { "7", "G" },
            { "8", "H" },
            { "9", "I" },
            { "10", "J" },
            { "11", "K" },
            { "12", "L" },
            { "13", "M" },
            { "14", "N" },
            { "15", "O" },
            { "16", "P" },
            { "17", "Q" },
            { "18", "R" },
            { "19", "S" },
            { "20", "T" },
            { "21", "U" },
            { "22", "V" },
            { "23", "W" },
            { "24", "X" },
            { "25", "Y" },
            { "26", "Z" },
        };

        public static Dictionary<string, string> signToLetter = new Dictionary<string, string>()
        {
            { "+", "Plus" },
            { "&", "And" },
            { ">>", "ShiftR" },
            { "<<", "ShiftL" },
        };

        public BindingList<string> ManualTestMessage { get; set; }
        public List<NamedGlobalVariable> CheckSumGlobalVariables { get; private set; }

        private ExcelWorksheet worksheet;
        private int columnCanUsed;
        private int columnNumberIdentifier;
        private int beforeCanStartColumn;
        private int afterCanStartColumn;
        private int objectColumn;
        private int limitStartBitColumn;
        private int limitBitLengthColumn;
        private int limitAColumn;
        private int limitBColumn;
        private int limitFrequencyColumn;
        private int endianColumn;
        public int formulaLogicColumn;
        private int testColumn = -1;
        private int checkSumAlgoritmColumn = -1;
        private int priorityColumn = -1;

        public MessageParser(ExcelWorkbook Workbook, Project project, BindingList<ArtLogics.Translation.Parser.Model.Resources> resources) : base(Workbook, project)
        {
            this.resources = resources;
            CanColumnMap = new Dictionary<int, string>();
            MessageName = new BindingList<string>();
            ManualTestMessage = new BindingList<string>();
            LimitSignalInfoList = new List<LimitSignalInfo>();
        }

        public void ParseMessages()
        {
            ParseSender();
            ParseReceiver();
            ConfigureCanChannels();
        }

        private void ParseReceiver()
        {
            worksheet = Workbook.Worksheets[ParserStaticVariable.ReceiverSheet];

            columnCanUsed = ExcelUtilsFunction.GetColumnNumber(5, ParserStaticVariable.CanReceiverUsedColName, worksheet);

            priorityColumn = -1;
            testColumn = -1;

            columnNumberIdentifier = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverIdColName, worksheet);
            beforeCanStartColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverBeforeStartColName, worksheet);
            afterCanStartColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverAfterFinishColName, worksheet);
            objectColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverObjectColName, worksheet) + 1;

            limitStartBitColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanLimitStartBitColName, worksheet);
            limitBitLengthColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanLimitBitLengthColName, worksheet);
            limitAColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanLimitAColName, worksheet);
            limitBColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanLimitBColName, worksheet);
            limitFrequencyColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverLimitFrequencyColName, worksheet);
            priorityColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverPriorityColName, worksheet);
            endianColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanEndianColName, worksheet);
            formulaLogicColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanFormulaLogicColName, worksheet);

            BuildCans(8);
        }

        private void ParseSender()
        {
            worksheet = Workbook.Worksheets[ParserStaticVariable.SenderSheet];

            priorityColumn = -1;
            testColumn = -1;
            checkSumAlgoritmColumn = -1;

            columnCanUsed = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderUsedColName, worksheet);

            columnNumberIdentifier = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderIdColName, worksheet);
            beforeCanStartColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderBeforeStartColName, worksheet);
            afterCanStartColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderAfterFinishColName, worksheet);
            objectColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderObjectColName, worksheet);

            limitStartBitColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanLimitStartBitColName, worksheet);
            limitBitLengthColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanLimitBitLengthColName, worksheet);
            limitAColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanLimitAColName, worksheet);
            limitBColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanLimitBColName, worksheet);
            limitFrequencyColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderLimitFrequencyColName, worksheet);
            testColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderTestColName, worksheet);
            checkSumAlgoritmColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CheckSumAlgoritmColName, worksheet);
            endianColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanEndianColName, worksheet);
            formulaLogicColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanFormulaLogicColName, worksheet);

            BuildCans(6);
        }

        public void BuildMessageSenderTest()
        {
            worksheet = Workbook.Worksheets[ParserStaticVariable.SenderSheet];
            var signalsToTest = LimitSignalInfoList.Where(limit => limit.Test != null);
            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
            testColumn = ExcelUtilsFunction.GetColumnNumber(4, ParserStaticVariable.CanSenderTestColName, worksheet);

            BuildCheckSumGlobalVariable();

            foreach (var signalToTest in signalsToTest)
            {
                try
                {
                    if (signalToTest.Test != "")
                    {
                        BuildOutputTest(signalToTest, state, worksheet.Cells[signalToTest.Line, testColumn].ToString());
                    }
                }
                catch (Exception err)
                {
                    HandleException(err, Resources.lang.WrongValueInSheet + " " + worksheet.Name + " " + Resources.lang.AtCell + " " + worksheet.Cells[signalToTest.Line, testColumn]);
                    continue;
                }
            }
        }

        private void BuildCheckSumGlobalVariable()
        {
            CheckSumGlobalVariables = new List<NamedGlobalVariable>();

            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte1"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte2"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte3"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte4"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte5"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte6"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumByte7"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumCount"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumComp"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumIdByte1"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumIdByte2"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumIdByte3"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumIdByte4"/*, GlobalVariableType.Decimal*/));
            //CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumTemp"/*, GlobalVariableType.Decimal*/));
            //CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumTempShift3"/*, GlobalVariableType.Decimal*/));
            //CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumTempShift6"/*, GlobalVariableType.Decimal*/));
            //CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSum"/*, GlobalVariableType.Decimal*/));
            //CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumCalculated"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumCheck"/*, GlobalVariableType.Decimal*/));
            CheckSumGlobalVariables.Add(ProjectUtilsFunction.BuildGlobalVariable(product, "CheckSumCheckComp"/*, GlobalVariableType.Decimal*/));
        }

        public void BuildOutputTest(LimitSignalInfo signalToTest, FlowChartState state, string cell)
        {
            List<string> conditions = signalToTest.Test.Split('\n').ToList();
            var i = 1;

            foreach (var conditionWithComment in conditions)
            {
                canLoopSignalUsed.Clear();
                testType = KeywordsEnum.NOKEYWORD;
                if (conditionWithComment == "" || (conditionWithComment.Length >= 2 && conditionWithComment.Substring(0, 2) == "//"))
                    continue;

                var tableWithComment = conditionWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                var condition = tableWithComment[0];

                var testInformation = " / " + signalToTest.ChannelName + " / " + i + " / " + conditionWithComment;


                var regexp = new Regex(@"(=[ ]*-[ ]*\d+[ ]*[^-—]*[-—]|[=]{0,1}[^-—]*[^-—0-9]+[^-—]*[-—])(.*)");
                var match = regexp.Match(condition.Trim());

                if (match.Groups.Count < 2)
                {
                    var ex = new ParserException(ParserExceptionType.RESULT_VALUE_EXPECTED_NOT_FOUND);
                    ex.Cell = cell;
                    ex.SubconditionLine = i;
                    ex.Mode = ParserExceptionMode.RESULT;
                    ex.Sheet = worksheet.Name;
                    throw ex;
                }

                /*var tab = condition.Split(new string[] { "-", "—" }, 2, StringSplitOptions.None);

                if (tab.Count() < 2)
                {
                    var ex = new ParserException(ParserExceptionType.RESULT_VALUE_EXPECTED_NOT_FOUND);
                    ex.Cell = cell;
                    ex.SubconditionLine = i;
                    ex.Mode = ParserExceptionMode.RESULT;
                    ex.Sheet = worksheet.Name;
                    throw ex;
                }*/

                /*var result = tab[0].Trim().ToLower();
                var operation = tab[1].Trim().ToLower();*/

                var result = match.Groups[1].Value.Trim().ToLower();
                result = result.Remove(result.Length - 1);
                var operation = match.Groups[2].Value?.Trim().ToLower();

                if (operation.Contains(KeywordsEnum.AFTER.ToString().ToLower()))
                {
                    operation = GetDelayTypeFromExcel(operation, KeywordsEnum.AFTER);
                }
                else if (operation.Contains(KeywordsEnum.DURING.ToString().ToLower()))
                {
                    var indexOfKeyword = operation.IndexOf(KeywordsEnum.DURING.ToString().ToLower());
                    operation = operation.Remove(indexOfKeyword);
                }
                else if (operation.Contains(KeywordsEnum.HIGH.ToString().ToLower()))
                {
                    operation = GetDelayTypeFromExcel(operation, KeywordsEnum.HIGH);
                }

                var tabOperation = operation.Split('且');

                try
                {
                    foreach (var op in tabOperation)
                    {
                        if (op != "")
                        {
                            try
                            {
                                var newResult = SetChannel(op, signalToTest.SignalName, state, Resources.lang.ChannelActionSetter + " " + signalToTest.SignalName + testInformation);

                                if (result == "=setpoint")
                                {
                                    result = "=" + newResult;
                                }
                            }
                            catch (ParserException ex)
                            {
                                ex.Cell = cell;
                                ex.Subcondition = op;
                                ex.SubconditionLine = 1;
                                ex.Mode = ParserExceptionMode.SETTER;
                                ex.Sheet = worksheet.Name;

                                throw ex;
                            }
                        }
                    }
                }
                catch (ParserException err)
                {
                    var error = err as ParserException;
                    if (error != null && error.Type == ParserExceptionType.RESISTIVE_VALUE_UNAVAILABLE)
                    {
                        HandleException(err, Resources.lang.WrongValueInSheet + " " + worksheet.Name + " " + Resources.lang.AtCell + " " + worksheet.Cells[signalToTest.Line, testColumn]);
                        var soundPath = "Sound/Sound.wav";
                        var soundPathDest = ParserStaticVariable.GlobalPath + soundPath.Substring(soundPath.LastIndexOf("/") + 1);

                        File.Copy(soundPath, soundPathDest, true);

                        ProjectUtilsFunction.BuildUserAction(state, Resources.lang.ConditionCannotBeTestedManually + error.Subcondition + Resources.lang.PleaseTestItManually, UserActionButtons.Fail, null, soundPathDest);
                        continue;
                    }
                    else
                    {
                        throw err;
                    }
                }

                if (testType == KeywordsEnum.AFTER /*|| testDuringSpecificTime*/)
                {
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBeforeTestingMessage + " " + signalToTest.SignalName + testInformation, (int)specificTime);
                }
                else
                {
                    var delayTime = signalToTest.Frequency * 3;
                    if (delayTime < 2000)
                    {
                        delayTime = 2000;
                    }
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBeforeTestingMessage + " " + signalToTest.SignalName + testInformation, delayTime);
                }

                try
                {
                    BuildRxTest(signalToTest, signalToTest.SignalName, result, state, testInformation);

                    if (result.Contains("checksum"))
                    {
                        ComputeChecksum(signalToTest, signalToTest.SignalName, result, state, testInformation);
                    }
                }
                catch (ParserException ex)
                {
                    ex.Cell = cell;
                    ex.Subcondition = signalToTest.SignalName + result;
                    ex.SubconditionLine = 1;
                    ex.Mode = ParserExceptionMode.RESULT;
                    ex.Sheet = worksheet.Name;
                    throw ex;
                }
                //reset test
                CleanCanLoops(state, testInformation);
                foreach (var op in tabOperation)
                {
                    if (op != "")
                    {
                        try
                        {
                            SetChannel(op, signalToTest.SignalName, state, Resources.lang.ChannelActionReset + " " + signalToTest.SignalName + testInformation, true);
                        }
                        catch (ParserException ex)
                        {
                            ex.Cell = cell;
                            ex.Subcondition = op;
                            ex.SubconditionLine = 1;
                            ex.Mode = ParserExceptionMode.RESET;
                            ex.Sheet = worksheet.Name;
                            throw ex;
                        }
                    }
                }
                i++;
            }
        }

        private void ComputeChecksum(LimitSignalInfo signalToTest, string signalName, string result, FlowChartState state, string testInformation)
        {
            var globalVariableAction = new GlobalVariableAction();
            globalVariableAction.Description = "checksum Calculation" + testInformation;
            globalVariableAction.ImagePath = null;
            state.Actions.Add(globalVariableAction);

            var globalVariableComparisonCheckSum = new GVCompareAction();
            globalVariableComparisonCheckSum.Description = "checksum Comparison" + testInformation;
            globalVariableComparisonCheckSum.ErrorMessage = new ErrorMessage();
            globalVariableComparisonCheckSum.ErrorMessage.Severity = Severity.Error;
            globalVariableComparisonCheckSum.ErrorMessage.Name = "CheckSum error";
            state.Actions.Add(globalVariableComparisonCheckSum);

            var globalVariableComparisonCheckSumGlobal = new GVCompareAction();
            globalVariableComparisonCheckSumGlobal.Description = "checksum message Comparison" + testInformation;
            globalVariableComparisonCheckSumGlobal.ErrorMessage = new ErrorMessage();
            globalVariableComparisonCheckSumGlobal.ErrorMessage.Severity = Severity.Error;
            globalVariableComparisonCheckSumGlobal.ErrorMessage.Name = "CheckSum error";
            state.Actions.Add(globalVariableComparisonCheckSumGlobal);

            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumIdByte1"),
                signalToTest.CanId.Substring(0, 2), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumIdByte2"),
                signalToTest.CanId.Substring(2, 2), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumIdByte3"),
                signalToTest.CanId.Substring(4, 2), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumIdByte4"),
                signalToTest.CanId.Substring(6, 2), OperationKind.Set, globalVariableAction);

            ParseCheckSumAlgoritm(signalToTest, signalName, result, state, testInformation, globalVariableAction);

            /*ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte1"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte2"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte3"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte4"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte5"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte6"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte7"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCount"),
                15, OperationKind.And, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCount"), OperationKind.Add, globalVariableAction);



            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                signalToTest.CanId.Substring(0, 2), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                signalToTest.CanId.Substring(2, 2), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                signalToTest.CanId.Substring(4, 2), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"),
                signalToTest.CanId.Substring(6, 2), OperationKind.Add, globalVariableAction);


            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift3"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift6"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift3"),
                3, OperationKind.ShiftR, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift6"),
                6, OperationKind.ShiftR, globalVariableAction);

            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift6"),
                3, OperationKind.And, globalVariableAction);

            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSum"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift6"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSum"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTempShift3"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSum"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumTemp"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSum"),
                7, OperationKind.And, globalVariableAction);*/

            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte1"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte2"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte3"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte4"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte5"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte6"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                8, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumByte7"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                4, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCount"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                4, OperationKind.ShiftL, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheckComp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"), OperationKind.Set, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheckComp"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumComp"), OperationKind.Add, globalVariableAction);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "Checksum"), OperationKind.Add, globalVariableAction);


            ProjectUtilGlobalVariableFunction.AddGlobalVariableComparisonItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "Checksum"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumComp"), ComparisonOperator.Equal, globalVariableComparisonCheckSum);
            ProjectUtilGlobalVariableFunction.AddGlobalVariableComparisonItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheck"),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == "CheckSumCheckComp"), ComparisonOperator.Equal, globalVariableComparisonCheckSumGlobal);
        }

        private void ParseCheckSumAlgoritm(LimitSignalInfo signalToTest, string signalName, string result, 
                FlowChartState state, string testInformation, GlobalVariableAction globalVariableAction)
        {
            var algoritm = signalToTest.Algoritm.Split('\n');

            var regexp = new Regex(@"(\([^\(\);]*?\))");

            foreach (var calcul in algoritm)
            {
                var tab = calcul.Split('=');
                var resultName = tab[0].Trim();
                var operation = tab[1].Trim();

                var match = regexp.Match(operation);
                while (match.Groups.Count > 1)
                {
                    var variableName = CalculateOperation(match.Groups[1].Value, state, testInformation, globalVariableAction);
                    operation = operation.Replace(match.Groups[0].Value, variableName);
                    match = regexp.Match(operation);
                }
                operation = CalculateOperation(operation, state, testInformation, globalVariableAction);

                if (CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == resultName) == null)
                {
                    var globalVariableResult = ProjectUtilsFunction.BuildGlobalVariable(product, resultName);
                    CheckSumGlobalVariables.Add(globalVariableResult);
                }

                ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == resultName),
                CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == operation.Trim()), OperationKind.Set, globalVariableAction);
            }
        }

        private string CalculateOperation(string operation, FlowChartState state, string testInformation, GlobalVariableAction globalVariableAction)
        {
            var regexp = new Regex(@"(([^\(\);]*?)[ ]([+&]|<<|>>)[ ]([^\(\)+&<>;]*))");
            var match = regexp.Match(operation);

            while (match.Groups.Count > 1)
            {
                var variableName = match.Groups[2].Value.Trim() + signToLetter[match.Groups[3].Value.Trim()] + match.Groups[4].Value.Trim();

                if (CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName) == null)
                {
                    var globalVariable = ProjectUtilsFunction.BuildGlobalVariable(product, variableName);
                    CheckSumGlobalVariables.Add(globalVariable);
                }

                ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName),
                    CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == match.Groups[2].Value), OperationKind.Set, globalVariableAction);

                switch (match.Groups[3].Value)
                {
                    case "+":
                        ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName),
                            CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == match.Groups[4].Value.Trim()), OperationKind.Add, globalVariableAction);
                        break;
                    case "<<":
                        ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName),
                            int.Parse(match.Groups[4].Value.Trim()), OperationKind.ShiftL, globalVariableAction);
                        break;
                    case ">>":
                        ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName),
                            int.Parse(match.Groups[4].Value.Trim()), OperationKind.ShiftR, globalVariableAction);
                        break;
                    case "&":
                        ProjectUtilGlobalVariableFunction.AddGlobalVariableActionItem(CheckSumGlobalVariables.FirstOrDefault(gv => gv.Name == variableName),
                            match.Groups[4].Value.Trim(), OperationKind.And, globalVariableAction);
                        break;
                }
                operation = operation.Replace(match.Groups[0].Value.Trim(), variableName);
                match = regexp.Match(operation);
            }
            return operation.Replace("(", "").Replace(")", "");
        }

        private void BuildRxTest(LimitSignalInfo signalInfo, string signalName, string result, FlowChartState state, string testInformation)
        {
            var channelAction = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.TestCanSenderSignal + " " + signalName + testInformation);

            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.IsUsed = true;
            channelActionContainer.MainBoard = signalInfo.Board;
            channelActionContainer.Channel = signalInfo.Channel;

            channelAction.ChannelActions.Add(channelActionContainer);

            var canOperation = new CanOperation();
            channelActionContainer.Operation = canOperation;

            canOperation.MessageName = signalInfo.MessageName;
            canOperation.Action = CanAction.RECVMSG;
            canOperation.ByteNumber = 8;
            canOperation.FeedbackCtrl = CanFeedbackCtrl.NEWESTMSG;
            canOperation.Variable = signalInfo.Variable;

            if (result.Contains("checksum"))
            {
                var i = 0;
                foreach (var checkSumGlobalVariable in CheckSumGlobalVariables)
                {
                    SignalBitInfo signalBitInfo;

                    if (i < 14)
                    {
                        signalBitInfo = ProjectUtilsCanFunction.BuildSignalBitInfo(8, i * 4, (signalInfo.Endian == "motorola") ? SignalEndian.Motorola : SignalEndian.Intel);
                        i++;
                    } else
                    {
                        signalBitInfo = ProjectUtilsCanFunction.BuildSignalBitInfo(4, i * 4, (signalInfo.Endian == "motorola") ? SignalEndian.Motorola : SignalEndian.Intel);
                    }

                    var signalData = ProjectUtilsCanFunction.BuildSignalData(0, 0, (double)signalInfo.B, (double)signalInfo.A);
                    signalData.BitInfo = signalBitInfo;

                    var signal = new Signal();
                    signal.Alias = checkSumGlobalVariable.Name;
                    signal.Data = signalData;

                    var signalGvPair = new SignalGvPair();
                    signalGvPair.Signal = signal;
                    signalGvPair.Variable = checkSumGlobalVariable;
                    canOperation.SignalGVs.Add(signalGvPair);
                    i++;
                    if (i > 15)
                        break;
                }
            }
            else
            {
                SignalData data = new SignalData();

                var regexpBinary = new Regex(@".*=b([-]?\d+)");
                Match matchBinary = regexpBinary.Match(result);

                if (!matchBinary.Success)
                {
                    data = ProjectUtilsCanFunction.BuildSignalData(0, 0, (double)signalInfo.B, (double)signalInfo.A);
                }
                else
                {
                    data = ProjectUtilsCanFunction.BuildSignalData(0, 0, 0, 1);
                }
                var bitInfo = ProjectUtilsCanFunction.BuildSignalBitInfo(signalInfo.BitLength, signalInfo.StartBit, (signalInfo.Endian == "motorola") ? SignalEndian.Motorola : SignalEndian.Intel);
                var errorMessage = ProjectUtilsLimitFunction.BuildErrorMessage(Resources.lang.ErrorWithCanSender + " " + signalInfo.CanId + " " + Resources.lang.Signal + " " + signalName, Severity.Error);

                /*var EqualComparison = new EqualTextComparison();
                EqualComparison.Text = result;*/
                var limitInfo = GetLimitInfo(signalInfo.Channel, signalName, signalName + result, true);

                var max = (signalInfo.A * ((decimal)Math.Pow(2, signalInfo.BitLength) - 1)) + signalInfo.B;

                var limitContainer = CreateComparisonFromString(limitInfo, result, true, max);
                // var limitContainer = new ComparisonContainer(EqualComparison, ComparisonKind.EQUALTEXT);

                var comparison = new SignalComparison();
                if (limitContainer.ComparisonKind == ComparisonKind.EQUALTEXT)
                {
                    comparison.SignalDataType = DataType.Hex;
                }
                else
                {
                    comparison.SignalDataType = DataType.UnsignedInt64;
                }
                comparison.BaseComparison = limitContainer;

                var limit = ProjectUtilsLimitFunction.BuildLimit(errorMessage, Resources.lang.ErrorWithCanSender + " " + signalInfo.CanId + " " + Resources.lang.Signal + " " + signalName, comparison, ComparisonKind.SIGNAL, 1);

                canOperation.Limits.Add(limit);

                var signal = new Signal();
                signal.Alias = signalInfo.MessageName;
                data.BitInfo = bitInfo;
                signal.Data = data;

                comparison.Signal = signal;
            }
        }

        private void BuildCans(int startColumn)
        {
            for (var i = beforeCanStartColumn + 1; i < afterCanStartColumn; i++)
            {
                CanColumnMap[i] = worksheet.Cells[startColumn - 2, i].Value.ToString();
            }

            for (var i = startColumn; (string)worksheet.Cells[i, columnCanUsed].Value != null || worksheet.Cells[i, columnCanUsed].Merge == true; i++)
            {
                if (worksheet.Cells[i, columnNumberIdentifier].Value == null)
                    continue;

                if (worksheet.Cells[i, columnCanUsed].Value == null || worksheet.Cells[i, columnCanUsed].Value.ToString() != "Y")
                {
                    if (worksheet.Cells[i, columnCanUsed].Value != null && worksheet.Cells[i, columnCanUsed].Value.ToString() == "N") {
                        var limitInfo = new LimitSignalInfo();
                        limitInfo.Test = "";
                        limitInfo.Line = i;
                        LimitSignalInfoList.Add(limitInfo);
                    }
                    continue;
                }

                /*if (worksheet.Cells[i, objectColumn].Value != null && worksheet.Cells[i, objectColumn].Value.ToString().Contains("DM"))
                {
                    ManualTestMessage.Add(worksheet.Cells[i, objectColumn].Value.ToString() + "ID :" +
                        System.Environment.NewLine +
                        System.Environment.NewLine +
                        System.Environment.NewLine +
                        worksheet.Cells[i, columnNumberIdentifier].Value.ToString());
                    continue;
                }*/
                try
                {
                    if (worksheet.Cells[i, objectColumn].Value != null && worksheet.Cells[i, objectColumn].Value.ToString().Contains("DM"))
                    {
                        ParseDmMessage(i);
                        continue;
                    }
                    else
                    {
                        try
                        {
                            ParseMessage(worksheet.Cells[i, columnNumberIdentifier].Value.ToString().Trim().Replace("0X", "").Replace("0x", ""), i);
                        }
                        catch
                        {
                            _log.Info(Resources.lang.WrongValueInSheet + " " + worksheet.Name + " " + Resources.lang.AtCell  + " " + worksheet.Cells[i, columnNumberIdentifier]);
                            continue;
                        }
                    }
                }
                catch
                {
                    _log.Info(Resources.lang.TheMessageHasNotBeenAssociatedToOneCanChannelInSheet + " " + worksheet.Name + " " + Resources.lang.AtLine + " " + i);
                    continue;
                }
            }
        }

        private void ParseDmMessage(int i)
        {
            var ids = worksheet.Cells[i, columnNumberIdentifier].Value.ToString().Split('\n');

            foreach(var id in ids)
            {
                ParseMessage(id.Trim(), i);
            }
        }

        private void ParseMessage(string variableId, int i)
        {
            var forwardList = new List<LimitSignalInfo>();
            bool associatedToOneCan = false;
            for (var e = beforeCanStartColumn + 1; e < afterCanStartColumn; e++)
            {
                try
                {
                    if (worksheet.Cells[i, e].Value != null)
                    {
                        try
                        {
                            BuildCan(i, e, variableId, forwardList);
                            associatedToOneCan = true;
                        }
                        catch (ParserException ex)
                        {
                            ex.SubconditionLine = i;
                            ex.Mode = ParserExceptionMode.PARSEMESSAGE;
                            ex.Sheet = worksheet.Name;
                            ex.Log();
                        }
                    }
                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + worksheet.Name + " " + Resources.lang.AtLine + " " + i);
                    continue;
                }
            }

            if (!associatedToOneCan)
            {
                throw new Exception(Resources.lang.TheMessageHasNotBeenAssociatedToOneCanChannel);
            }
        }

        private void BuildCan(int i, int e, string variableId, List<LimitSignalInfo> forwardList)
        {
            MessageFlowDirection direction;

            if (worksheet.Cells[i, e].Value.ToString().ToLower() == "rx")
            {
                direction = MessageFlowDirection.TX;
            }
            else
            {
                direction = MessageFlowDirection.RX;
            }

            var variableName = NumberToLetter[CanColumnMap[e].Replace("CAN", "")] +
            direction.ToString() +
            variableId;

            var LimitSignalInfo = new LimitSignalInfo();
            if (worksheet.Cells[i, objectColumn].Value != null)
            {
                LimitSignalInfo = BuildLimitSignalInfo(direction, variableName, variableId, i);
            }

            Variable variable;
            bool setVariable = false;

            variable = product.Variables.FirstOrDefault(v => v.Name == variableName);

            if (variable == null)
            {
                variable = CreateVariable(variableName, out setVariable);
            }

            if (!MessageName.Contains(variableName))
            {
                MessageName.Add(variableName);
            }
            var setting = new CanMessageVariableSettings();

            var boardNumber = 0;

            var id = int.Parse(variableId.Replace("0x", "").Replace("0X", ""), NumberStyles.HexNumber);
            foreach (var resource in resources)
            {
                if (resource.CanMap.ContainsValue(CanColumnMap[e]))
                {
                    var XcuChannelName = resource.CanMap.FirstOrDefault(a => a.Value == CanColumnMap[e]).Key;
                    setting.Channel = product.Boards[boardNumber].Channels.FirstOrDefault(a => a.CurrentName.Replace(" - ", "") == XcuChannelName);
                    setting.Direction = direction;
                    setting.MessageId = id;
                    setting.Type = CanMessageFrameType.EXT;
                    if (setVariable)
                    {
                        variable.Settings = setting;
                        product.Variables.Add(variable);
                    }
                    LimitSignalInfo.Board = product.Boards[boardNumber];
                    LimitSignalInfo.Channel = setting.Channel;
                    LimitSignalInfo.ChannelName = CanColumnMap[e];
                    LimitSignalInfo.Variable = variable;
                    break;
                }
                boardNumber++;
            }

            if (worksheet.Cells[i, e].Value.ToString().Trim() != "" && worksheet.Cells[i, objectColumn].Value != null)
            {
                LimitSignalInfo.ForwardList = forwardList;
                LimitSignalInfoList.Add(LimitSignalInfo);
            }

            if (worksheet.Cells[i, e].Value.ToString().Trim() != "" && worksheet.Cells[i, e].Value.ToString().ToLower() != "rx")
            {
                forwardList.Add(LimitSignalInfo);
            }
        }

        private LimitSignalInfo BuildLimitSignalInfo(MessageFlowDirection direction, string variableName, string variableId, int i)
        {
            var limitSignalInfo = new LimitSignalInfo()
            {
                A = (worksheet.Cells[i, limitAColumn].Value == null) ? 0 : decimal.Parse(worksheet.Cells[i, limitAColumn].Value.ToString()),
                B = (worksheet.Cells[i, limitBColumn].Value == null) ? 0 : decimal.Parse(worksheet.Cells[i, limitBColumn].Value.ToString()),
                StartBit = (worksheet.Cells[i, limitStartBitColumn].Value == null) ? 0 : int.Parse(worksheet.Cells[i, limitStartBitColumn].Value.ToString()),
                BitLength = (worksheet.Cells[i, limitBitLengthColumn].Value == null) ? 0 : int.Parse(worksheet.Cells[i, limitBitLengthColumn].Value.ToString()),
                CanId = variableId,
                Frequency = (worksheet.Cells[i, limitFrequencyColumn].Value == null) ? 0 : int.Parse(worksheet.Cells[i, limitFrequencyColumn].Value.ToString().Replace("ms", "")),
                Direction = direction,
                MessageName = variableName,
                SignalName = worksheet.Cells[i, objectColumn].Value.ToString().ToLower().Replace("\n", " "),
                Endian = (worksheet.Cells[i, endianColumn].Value == null) ? "" : worksheet.Cells[i, endianColumn].Value.ToString().ToLower(),
                Line = i,
            };

            if (limitSignalInfo.Endian == "motorola" && (limitSignalInfo.BitLength > 8 * (limitSignalInfo.StartBit + 1)))
            {
                throw new ParserException(ParserExceptionType.CAN_LIMIT_DEFINITION_ERROR);
            }

            if (formulaLogicColumn > 0)
            {
                limitSignalInfo.FormulaLogic = (worksheet.Cells[i, formulaLogicColumn].Value == null) ? 0 : int.Parse(worksheet.Cells[i, formulaLogicColumn].Value.ToString());
            }

            if (priorityColumn > 0)
            {
                limitSignalInfo.Priority = (worksheet.Cells[i, priorityColumn].Value == null) ? 0 : int.Parse(worksheet.Cells[i, priorityColumn].Value.ToString());
            }

            if (testColumn > 0)
            {
                limitSignalInfo.Test = (worksheet.Cells[i, testColumn].Value == null) ? "" : worksheet.Cells[i, testColumn].Value.ToString();
            }

            if (checkSumAlgoritmColumn > 0)
            {
                limitSignalInfo.Algoritm = (worksheet.Cells[i, checkSumAlgoritmColumn].Value == null) ? "" : worksheet.Cells[i, checkSumAlgoritmColumn].Value.ToString();
            }

            return limitSignalInfo;
        }

        private Variable CreateVariable(string variableName, out bool setVariable)
        {
            var variable = new Variable();
            variable.Kind = VariableKind.CanMessage;
            variable.Name = variableName;
            setVariable = true;

            return variable;
        }

        public void ConfigureCanChannels()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.ReceiverSheet];
            var i = 7;
            var beforeCanStartColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverBeforeStartColName, Worksheet);
            var afterCanStartColumn = ExcelUtilsFunction.GetColumnNumber(6, ParserStaticVariable.CanReceiverAfterFinishColName, Worksheet);

            for (var e = beforeCanStartColumn + 1; e < afterCanStartColumn; e++)
            {
                if (Worksheet.Cells[i, e].Value != null)
                {
                    var CanBaudrate = (PresetBaudrate)Enum.Parse(typeof(PresetBaudrate), "Br" + Worksheet.Cells[i, e].Value.ToString().Replace("k", "K"));
                    var setting = new CanSettings();

                    setting.Baudrate = CanBaudrate;

                    var boardNumber = 0;
                    foreach (var resource in resources)
                    {
                        if (resource.CanMap.ContainsValue(CanColumnMap[e]))
                        {
                            var XcuChannelName = resource.CanMap.FirstOrDefault(a => a.Value == CanColumnMap[e]).Key;
                            var channel = product.Boards[boardNumber].Channels.FirstOrDefault(a => a.CurrentName.Replace(" - ", "") == XcuChannelName);

                            product.ChannelSettings.Add(
                                new ChannelSettings()
                                {
                                    Channel = channel,
                                    Settings = setting
                                });

                            break;
                        }
                        boardNumber++;
                    }
                }
            }
        }

        public void BuildManualTestMessage()
        {
            foreach (var message in ManualTestMessage)
            {
                var text = message +
                    System.Environment.NewLine +
                    System.Environment.NewLine +
                    "Need to be manually tested";

                var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
                ProjectUtilsFunction.BuildUserAction(state, text, UserActionButtons.Pass | UserActionButtons.Fail);
            }
        }
    }
}
