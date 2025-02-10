using ActiaParser.Define;
using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Shared.FlowChart;
using ArtLogics.TestSuite.Testing.Actions.CanLoop;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.PowerSupply;
using ArtLogics.TestSuite.Testing.Actions.Report;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.Configuration;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Utils;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class StateParser : ActionParserBase
    {
        public List<string> specialKeywords = new List<string>()
        {
            "DELAY",
            "PSUC",
            "PSUV",
            "PSU",
            "USERACTION",
        };

        private float Psuc = 0;
        private float Psuv = 0;
        private BindingList<WorkspaceUnit> Units;

        public StateParser(ExcelWorkbook Workbook, Project project) : base(Workbook, project)
        {
            Units = project.Workspaces[0].Units;
        }

        public void ParseStartState()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.StartStepsSheet];

            ParseState(Worksheet);
        }

        public void ParseEndState(string projectName)
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.EndStepsSheet];

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];
            
            ParseState(Worksheet);

            CleanCanLoops(state, "", false);

            CreateReport(state, projectName);
        }

        private void ParseState(ExcelWorksheet Worksheet)
        {
            var columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.StepsDefinitionColName, Worksheet);
            var columnNumberValue = ExcelUtilsFunction.GetColumnNumber(2, ParserStaticVariable.StepsValueColName, Worksheet);
            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            bool SpecialFeatureUsed = true;
            ChannelAction CurrentChannelAction = null;

            for (var i = 3; Worksheet.Cells[i, columnNumberDefinition].Value != null; i++)
            {
                try
                {
                    if (Worksheet.Cells[i, columnNumberDefinition].Value == null || Worksheet.Cells[i, columnNumberDefinition].Value.ToString() == "")
                        continue;

                    var Definition = Worksheet.Cells[i, columnNumberDefinition].Value.ToString();
                    var Value = Worksheet.Cells[i, columnNumberValue].Value.ToString();

                    var limitInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && l.SignalName == Definition.ToLower());

                    if (specialKeywords.Contains(Definition))
                    {
                        ApplySpecialFunction(state, Definition, Value);
                        SpecialFeatureUsed = true;
                    }
                    else if (limitInfo != null)
                    {
                        ParseStartCan(state, Definition, float.Parse(Value));
                    }
                    else
                    {
                        if (SpecialFeatureUsed)
                        {
                            CurrentChannelAction = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelAction);
                        }

                        try
                        {
                            BuildBasicAction(CurrentChannelAction, Definition, Value);
                        }
                        catch
                        {
                            try
                            {
                                SetChannelAndDefaultValue((Definition + Value).ToLower(), Definition.ToLower(), state);
                            }
                            catch
                            {
                                _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtLine + " " + i + ", " + Resources.lang.ThisRessourcesIsNotAvailableInYourProject);
                                continue;
                            }
                        }
                        SpecialFeatureUsed = false;
                    }
                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell + " " + Worksheet.Cells[i, columnNumberDefinition]);
                    continue;
                }
            }
        }

        private void SetChannelAndDefaultValue(string operation, string definition, FlowChartState state)
        {
            List<ChannelInputInfo> channelInputInfos = new List<ChannelInputInfo>();
            IEnumerable<ChannelInputInfo> channels = null;
            float value = -1;
            
            foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
            {
                if (key != "0" && operation.Contains(key.ToLower()))
                {
                    channels = KeywordParser.FunctionChannelInputMap.Where(o => o.Function != "0"
                    && KeywordParser.KeywordFunctionMap[key].Contains(o.Function));

                    operation = operation.Replace(key.ToLower(), "");

                    channelInputInfos = GetChannelInfoFromKeyWord(operation, channels, false);
                    break;
                }
            }


            if (channels != null && channels.Count() > 0 && channelInputInfos.Count <= 0)
            {
                channelInputInfos = ParseRelayResistive(operation, channels, true, out value);
            }

            ParseStandardAction(operation, definition, false, state, channelInputInfos, Resources.lang.SetDefaultValueFor + " " + definition, false, false);

            foreach(var channelInputInfo in channelInputInfos)
            {
                channelInputInfo.DefaultValue = channelInputInfo.CurrentValue;
            }
        }

        private void BuildBasicAction(ChannelAction CurrentChannelAction, string Definition, string Value)
        {
            Channel channel;

            var table = Definition.Split(',');
            var CurrentBoardNumber = int.Parse(table[0]) - 1;

            var RessourceKind = (ChannelKind)Enum.Parse(typeof(ChannelKind), table[1].Remove(table[1].IndexOf("-")));
            var RessourceNumber = int.Parse(table[1].Remove(0, table[1].IndexOf("-") + 1)) - 1;

            var channels = product.Boards[CurrentBoardNumber].Channels.Concat(product.Boards[CurrentBoardNumber].Extensions.SelectMany(f => f.Channels));
            channel = channels.Where(e => e.Kind.Equals(RessourceKind)).ElementAt(RessourceNumber);

            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.IsUsed = true;
            channelActionContainer.Channel = channel;
            channelActionContainer.MainBoard = product.Boards[CurrentBoardNumber];
            CurrentChannelAction.ChannelActions.Add(channelActionContainer);

            CreateInitialOperation(channelActionContainer, Value);
        }

        private void ParseStartCan(FlowChartState state, string Definition, float Value)
        {
            var limitCanInfo = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && l.SignalName == Definition.ToLower()).FirstOrDefault();
            var lastCanLoop = state.Actions.LastOrDefault(a => a is CanLoopAction && ((CanLoopAction)a).Variable.Name == limitCanInfo.Variable.Name) as CanLoopAction;

            if (lastCanLoop != null)
            {
                SetCanLoop(lastCanLoop, Definition.ToLower(), Value, lastCanLoop.Data.Remove(0, 2));

                var limitCanInfos = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && l.Variable.Name == limitCanInfo.Variable.Name && l.Direction == MessageFlowDirection.TX);
                foreach (var limit in limitCanInfos)
                {
                    limit.DefaultValue = lastCanLoop.Data;
                    limit.ResetLoop = true;
                }
            }
            else
            {
                var canLoopAction = new CanLoopAction();
                canLoopAction.CanLoopCmd = CanLoopCmd.LOOP;
                canLoopAction.Description = Resources.lang.StartCanLoop + " " + Definition;
                SetCanLoop(canLoopAction, Definition.ToLower(), Value);
                state.Actions.Add(canLoopAction);

                var limitCanInfos = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && l.Variable.Name == limitCanInfo.Variable.Name && l.Direction == MessageFlowDirection.TX);
                foreach (var limit in limitCanInfos)
                {
                    limit.DefaultValue = canLoopAction.Data;
                    limit.ResetLoop = true;
                }

                currentCanLoopAction.Add(new Tuple<string, object>(Definition.ToLower(), Value));
            }
        }

        private void CreateInitialOperation(ChannelActionContainer channelActionContainer, string value)
        {
            var channel = channelActionContainer.Channel;


            var channelOutputInfo = KeywordParser.FunctionChannelOutputMap.Where(ci => ci.Channel == channel).FirstOrDefault();
            var channelInputInfo = KeywordParser.FunctionChannelInputMap.Where(ci => ci.Channel == channel).FirstOrDefault();

            switch (channel.Kind)
            {
                case ChannelKind.AWGOUT:
                    ProjectUtilsFunction.BuildAwgOut(channelActionContainer, AwgAction.START, AwgFunction.SQUARE,
                                1, 1, float.Parse(value), float.Parse(value) / 2);
                    break;
                case ChannelKind.DCVOUT:
                    var floatValue = float.Parse(value);
                    ProjectUtilsFunction.BuildDCVout(channelActionContainer, floatValue);

                    if ((object)channelOutputInfo != null)
                    {
                        channelOutputInfo.DefaultValue = floatValue;
                    }

                    if ((object)channelInputInfo != null)
                    {
                        channelInputInfo.DefaultValue = floatValue;
                    }

                    break;
                case ChannelKind.RELAY:
                    var relayState = RelayState.CLOSE;
                    if (value == "OPEN")
                    {
                        relayState = RelayState.OPEN;
                    }

                    ProjectUtilsFunction.BuildRelayOperation(channelActionContainer, relayState);

                    if ((object)channelOutputInfo != null)
                    {
                        channelOutputInfo.DefaultValue = relayState;
                    }

                    if ((object)channelInputInfo != null)
                    {
                        channelInputInfo.DefaultValue = relayState;
                    }
                    break;
                case ChannelKind.FRQOUT:
                    var frqValue = float.Parse(value);
                    ProjectUtilsFunction.BuildFrqOut(channelActionContainer, frqValue);
                    break;
                case ChannelKind.None:
                default:
                    break;
            }
        }

        private void ApplySpecialFunction(FlowChartState state, string definition, string value)
        {
            switch (definition)
            {
                case "USERACTION":
                    BuildUserAction(state, value);
                    break;
                case "DELAY":
                    ProjectUtilsFunction.BuildDelay(state, Resources.lang.Delay, int.Parse(value));
                    break;
                case "PSUC":
                    Psuc = float.Parse(value);
                    InputParser.PSUC = Psuc;
                    break;
                case "PSUV":
                    Psuv = float.Parse(value);
                    InputParser.PSUV = Psuv;
                    break;
                case "PSU":
                    BuildPsuAction(state, value);
                    break;
            }
        }

        private void BuildUserAction(FlowChartState state, string value)
        {
            var userAction = new UserInputAction();

            var tab = value.Split('\n');

            string message = "";
            List<UserActionButtons> buttons = new List<UserActionButtons>();

            foreach(var text in tab)
            {
                var values = text.Split(':');

                if (values[0] == "Message")
                {
                    message = values[1];
                } else
                {
                    buttons.Add((UserActionButtons)Enum.Parse(typeof(UserActionButtons), values[1], true));
                }
            }

            userAction.Description = message;

            if (buttons.Count > 0)
            {
                userAction.Buttons = buttons[0];

                for (var i = 1; i < buttons.Count; i++)
                {
                    userAction.Buttons |= buttons[0];
                }
            }
            state.Actions.Add(userAction);
        }

        private void BuildPsuAction(FlowChartState state, string value)
        {
            foreach (var Unit in Units) {
                var groupResultAppearence = new GroupResultAppearance();
                groupResultAppearence.Name = "PSU " + Unit.Unit.Name;
                product.GroupAppearances.Add(groupResultAppearence);

                Unit.GroupId = groupResultAppearence.Guid;

                var psuAction = new PowerSupplyAction();
                psuAction.Current = Psuc;
                psuAction.Description = Resources.lang.PsuAction +  " " + Unit.Unit.Name;
                psuAction.Voltage = Psuv;
                psuAction.GroupId = groupResultAppearence.Guid;

                if (value == "ON")
                {
                    psuAction.Activate = true;
                } else
                {
                    psuAction.Activate = false;
                }

                state.Actions.Add(psuAction);
            }
        }
    }
}
