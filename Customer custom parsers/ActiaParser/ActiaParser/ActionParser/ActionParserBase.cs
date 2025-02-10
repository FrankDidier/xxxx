using ActiaParser.MessageParser;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.Limits.Comparisons.MultiRange;
using ArtLogics.TestSuite.Limits.Comparisons.SingleRange;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.TestResults;
using ArtLogics.Translation.Parser.Model;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ArtLogics.TestSuite.Environment.Dbc;
using System.Collections;
using ArtLogics.TestSuite.Testing.Actions.CanLoop;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Testing.StateMachines;
using System.Text.RegularExpressions;
using ArtLogics.TestSuite.Actions.Common;
using ArtLogics.TestSuite.Limits.Comparisons.Text;
using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Testing.Actions.CaptureSensor;
using ArtLogics.Translation.Parser;
using ArtLogics.TestSuite.Shared.FlowChart;
using ArtLogics.TestSuite.Testing.Actions.PowerSupply;
using ArtLogics.TestSuite.Actions;
using ArtLogics.Translation.Parser.Utils;
using ArtLogics.Translation.Parser.Exception;
using System.IO;
using ArtLogics.TestSuite.Testing.Actions.Report;
using ArtLogics.TestSuite.Testing.Configuration;
using ArtLogics.TestSuite.DevXlate.Units;
using System.Drawing;

namespace ActiaParser.ActionParser
{
    public abstract class ActionParserBase : BaseParser
    {
        protected ExcelWorkbook Workbook;
        protected Product product;
        protected List<Tuple<string, object>> currentCanLoopAction = new List<Tuple<string, object>>();
        protected float specificTime;
        protected KeywordsEnum testType = KeywordsEnum.NOKEYWORD;
        protected List<LimitSignalInfo> canLoopSignalUsed { get; set; } = new List<LimitSignalInfo>();

        public ActionParserBase(ExcelWorkbook Workbook, Project project)
        {
            this.Workbook = Workbook;
            this.product = project.Products[0];
            this.project = project;
        }

        protected float SetChannel(string ChannelName, string function, FlowChartState state,
                                        string channelActionDesc, bool resetChannel = false,
                                        bool buildAfterDelay = true, bool setDefaultValue = false)
        {
            ChannelName = ChannelName.ToLower();
            List<ChannelInputInfo> channelInputInfos = new List<ChannelInputInfo>();
            IEnumerable<ChannelInputInfo> channels = null;
            bool isCustomAction = false;
            string customAction = "";
            float value=-1;

            var functionName = GetFunctionName(ChannelName);

            if (functionName != null)
            {
                if (KeywordParser.KeywordFunctionMap[functionName] == "psu" || 
                    KeywordParser.KeywordFunctionMap[functionName] == "delay" ||
                    KeywordParser.KeywordFunctionMap[functionName] == "useraction")
                {
                    isCustomAction = true;
                    customAction = KeywordParser.KeywordFunctionMap[functionName];
                }
                else
                {
                    channels = KeywordParser.FunctionChannelInputMap.Where(o => o.Function != "0"
                    && KeywordParser.KeywordFunctionMap[functionName] == o.Function);

                    ChannelName = ChannelName.Replace(functionName.ToLower(), "");

                    channelInputInfos = GetChannelInfoFromKeyWord(ChannelName, channels, resetChannel);
                }
            }

            if (isCustomAction)
            {
                if (customAction == "psu")
                {
                    value = ParsePSU(ChannelName, state, resetChannel);
                } else if (customAction == "delay")
                {
                    if (!resetChannel)
                    {
                        ParseDelay(ChannelName, state);
                    }
                    buildAfterDelay = false;
                }
                else if (customAction == "useraction")
                {
                    ParseUserAction(ChannelName, state, resetChannel);
                }
            }
            else
            {
                if (channels != null && channels.Count() > 0 && channelInputInfos.Count <= 0)
                {
                    channelInputInfos = ParseRelayResistive(ChannelName, channels, setDefaultValue, out value);
                }

                var newChannelAction = buildAfterDelay;

                var sentValue = ParseStandardAction(ChannelName, function, resetChannel, state, channelInputInfos, channelActionDesc, newChannelAction, setDefaultValue);

                if (value == -1)
                {
                    value = sentValue;
                }
            }

            if (!(state.Actions.LastOrDefault() is DelayAction) && buildAfterDelay)
            {
                var channelAction = state.Actions.LastOrDefault() as ActionBase;
                var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["DelayAfterInput"]);
                ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayAfter + " " + channelAction.Description, delayTime);
            }

            return value;
        }

        protected string GetFunctionName(string channelName)
        {
            foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
            {
                if (key != "0" && channelName.Contains(key.ToLower()))
                {
                    return key;
                }
            }
            return null;
        }

        protected float ParseStandardAction(string ChannelName, string function, bool resetChannel,
            FlowChartState state, List<ChannelInputInfo> channelInputInfos,
            string channelActionDesc, bool newChannelAction, bool setDefaultValue)
        {
            float sentValue = 0;
            if (channelInputInfos.Count <= 0)
            {
                sentValue = CreateCanTxFromSignalList(ChannelName, state, resetChannel);
            }
            else if (channelInputInfos.Count > 0)
            {
                sentValue = (float)GetValueFromString(ChannelName);
                var channelActionContainers = new List<ChannelActionContainer>();
                foreach (var channelInputInfo in channelInputInfos)
                {
                    var channelActionContainer = new ChannelActionContainer();
                    channelActionContainer.IsUsed = true;
                    channelActionContainer.MainBoard = channelInputInfo.Board;
                    channelActionContainer.Channel = channelInputInfo.Channel;

                    CreateInputOperation(channelActionContainer, function, ChannelName, channelInputInfo, setDefaultValue, resetChannel);
                    channelActionContainers.Add(channelActionContainer);
                }

                ChannelActionContainersCompare comp = new ChannelActionContainersCompare();
                channelActionContainers.Sort(comp);

                ChannelAction channelAction = null;
                if (newChannelAction) {
                    channelAction = ProjectUtilsFunction.BuildChannelAction(state, channelActionDesc);
                } else
                {
                    channelAction = (ChannelAction)state.Actions.LastOrDefault(ch => ch is ChannelAction);
                }

                if (channelAction == null)
                {
                    channelAction = ProjectUtilsFunction.BuildChannelAction(state, channelActionDesc);
                }

                foreach (var channelActionContainer in channelActionContainers)
                {
                    channelAction.ChannelActions.Add(channelActionContainer);
                }
            }
            return sentValue;
        }

        protected List<ChannelInputInfo> ParseRelayResistive(string channelName, IEnumerable<ChannelInputInfo> channels, bool setDefaultValue, out float RealValue)
        {
            List<ChannelInputInfo> channelInputInfos = new List<ChannelInputInfo>();

            var functionName = channels.FirstOrDefault().Function;
            var value = GetValue(channelName);
            var defaultChannelInfo = channels.Where(ch => ch.Channel.Kind == ChannelKind.RELAY
                && (RelayState)ch.DefaultValue == RelayState.CLOSE).FirstOrDefault();
            if (defaultChannelInfo != null)
            {
                if (setDefaultValue) {
                    defaultChannelInfo.BaseDefaultValue = defaultChannelInfo.DefaultValue;
                    defaultChannelInfo.DefaultValue = RelayState.OPEN;
                }
                defaultChannelInfo.CurrentValue = RelayState.OPEN;
            }

            var channelInputInterpretation = KeywordParser.ChannelInputInterpretations[functionName];
            var closestValue = GetClosestValue(channelName, channelInputInterpretation, value);
            RealValue = (float)value;

            if (closestValue >= 500 || closestValue < 0)
            {
                throw new ParserException(ParserExceptionType.RESISTIVE_VALUE_UNAVAILABLE);
            }

            var firstChannelInterpretation = channelInputInterpretation.Where(ci => ci.InterpretedValue >= 0).ElementAt(0);
            var secondChannelInterpretation = channelInputInterpretation.Where(ci => ci.InterpretedValue >= 0).ElementAt(1);

            bool inverted = true;

            if (((firstChannelInterpretation.InterpretedValue - secondChannelInterpretation.InterpretedValue) > 0 && (firstChannelInterpretation.RealValue - secondChannelInterpretation.RealValue) > 0)
                || ((firstChannelInterpretation.InterpretedValue - secondChannelInterpretation.InterpretedValue) < 0 && (firstChannelInterpretation.RealValue - secondChannelInterpretation.RealValue) < 0))
            {
                inverted = false;
            }

            var closestInputInfo = channels.OrderBy(ch =>
            {
                decimal DecimalValue;

                if (channelName.Contains(KeywordsEnumList.KeyWordToString[KeywordsEnum.INF]))
                {
                    if (decimal.TryParse(ch.HwValue, out DecimalValue))
                    {
                        if (inverted && DecimalValue <= closestValue) {
                            return decimal.MaxValue;
                        }
                        else if (!inverted && DecimalValue >= closestValue)
                        {
                            return decimal.MaxValue;
                        } else
                        {
                            return Math.Abs(DecimalValue - closestValue);
                        }
                    }
                    else
                    {
                        return decimal.MaxValue;
                    }
                }
                else if (channelName.Contains(KeywordsEnumList.KeyWordToString[KeywordsEnum.SUP]))
                {
                    if (decimal.TryParse(ch.HwValue, out DecimalValue))
                    {
                        if (inverted && DecimalValue <= closestValue)
                        {
                            return Math.Abs(DecimalValue - closestValue);
                        }
                        else if (!inverted && DecimalValue >= closestValue)
                        {
                            return Math.Abs(DecimalValue - closestValue);
                        }
                        else
                        {
                            return decimal.MaxValue;
                        }
                    }
                    else
                    {
                        return decimal.MaxValue;
                    }
                }
                else
                {
                    if (decimal.TryParse(ch.HwValue, out DecimalValue))
                    {
                        return Math.Abs(DecimalValue - closestValue);
                    }
                    else
                    {
                        return decimal.MaxValue;
                    }
                }
            }).FirstOrDefault();
            if (setDefaultValue)
            {
                closestInputInfo.BaseDefaultValue = defaultChannelInfo.DefaultValue;
                closestInputInfo.DefaultValue = RelayState.CLOSE;
            }
            closestInputInfo.CurrentValue = RelayState.CLOSE;

            if (closestInputInfo.HwValue != null && closestInputInfo.HwValue != "")
            {
                RealValue = (float)channelInputInterpretation.Interpollation(decimal.Parse(closestInputInfo.HwValue), true);
            }

            if (defaultChannelInfo != null && defaultChannelInfo.Channel != closestInputInfo.Channel)
            {
                channelInputInfos.Add(defaultChannelInfo);
            }
            channelInputInfos.Add(closestInputInfo);

            return channelInputInfos;
        }

        private decimal GetClosestValue(string channelName, ChannelInputInterpretations channelInputInterpretation, decimal value)
        {
            /*if (channelName.Contains(KeywordsEnumList.KeyWordToString[KeywordsEnum.INF]))
            {
                return channelInputInterpretation.GetClosestInferiorValue(value);
            }
            else if (channelName.Contains(KeywordsEnumList.KeyWordToString[KeywordsEnum.SUP]))
            {
                return channelInputInterpretation.GetClosestSuperiorValue(value);
            }
            else
            {*/
                return channelInputInterpretation.Interpollation(value);
            //}

            throw new ParserException(ParserExceptionType.CALCULATION_ISSUE);
        }

        protected List<ChannelInputInfo> GetChannelInfoFromKeyWord(string ChannelName, IEnumerable<ChannelInputInfo> channels,
            bool resetChannel)
        {
            List<ChannelInputInfo> channelInputInfos = new List<ChannelInputInfo>();

            ChannelName = ReplaceKeyword(ChannelName);

            var keyword = KeywordsEnumList.GetKeyword(ChannelName);

            if (keyword != KeywordsEnum.NOKEYWORD)
            {
                if (channels.Count() == 1)
                {
                    channelInputInfos.Add(channels.FirstOrDefault());
                }
                else
                {
                    if ((keyword & KeywordsEnum.ON) == KeywordsEnum.ON ||
                        (keyword & KeywordsEnum.OFF) == KeywordsEnum.OFF)
                    {
                        var channelInputInfoVbat = channels.Where(c => c.HwValue == "vbat").FirstOrDefault();
                        if (channelInputInfoVbat != null)
                        {
                            channelInputInfos.Add(channelInputInfoVbat);
                        }

                        var channelInputInfoGnd = channels.Where(c => c.HwValue == "gnd").FirstOrDefault();
                        if (channelInputInfoGnd != null)
                        {
                            channelInputInfos.Add(channelInputInfoGnd);
                        }
                    }
                    else if ((keyword & KeywordsEnum.SUP) == KeywordsEnum.SUP ||
                      (keyword & KeywordsEnum.INF) == KeywordsEnum.INF ||
                      (keyword & KeywordsEnum.EQUAL) == KeywordsEnum.EQUAL)
                    {
                        if (channels.Count() > 0)
                        {
                            var logicalContact = channels.FirstOrDefault().LogicalContact;

                            if (logicalContact == "ana_v")
                            {
                                CalculateVoltageChannelValue(ChannelName, channels, channelInputInfos, keyword, resetChannel);

                            }
                            else if (logicalContact.Contains("freq"))
                            {
                                CalculateFrequencyChannelValue(ChannelName, channels, channelInputInfos, keyword, resetChannel);
                            } else if ((keyword & KeywordsEnum.OC) == KeywordsEnum.OC)
                            {
                                foreach(var channel in channels)
                                {
                                    if (channel.Channel.Kind != ChannelKind.RELAY)
                                        continue;
                                    channel.CurrentValue = RelayState.OPEN;
                                    channelInputInfos.Add(channel);
                                }
                            }
                            else if ((keyword & KeywordsEnum.GND) == KeywordsEnum.GND)
                            {
                                var defaultChannelInfo = channels.Where(ch => ch.Channel.Kind == ChannelKind.RELAY
                                    && (RelayState)ch.DefaultValue == RelayState.CLOSE).FirstOrDefault();
                                if (defaultChannelInfo != null)
                                {
                                    defaultChannelInfo.CurrentValue = RelayState.OPEN;
                                    channelInputInfos.Add(defaultChannelInfo);
                                }

                                var channelInputInfoGnd = channels.Where(c => c.HwValue == "gnd").FirstOrDefault();
                                channelInputInfoGnd.CurrentValue = RelayState.CLOSE;
                                channelInputInfos.Add(channelInputInfoGnd);
                            }
                        }
                    }
                }
            }

            return channelInputInfos;
        }

        private void CalculateVoltageChannelValue(string ChannelName, IEnumerable<ChannelInputInfo> channels,
                                                    List<ChannelInputInfo> channelInputInfos, KeywordsEnum keyword, bool resetChannel)
        {
            var channelInputInfoVout = channels.Where(c => c.HwValue == "vout").FirstOrDefault();
            if (channelInputInfoVout != null)
            {
                channelInputInfoVout.CurrentValue = RelayState.CLOSE;
                channelInputInfos.Add(channelInputInfoVout);
            }

            float value = 0;
            var channelVout = channels.Where(ch => ch.Channel.Kind == ChannelKind.DCVOUT).FirstOrDefault();
            if (!resetChannel)
            {
                var functionName = channels.FirstOrDefault().Function;
                var valueTable = ChannelName.Split(new string[] { KeywordsEnumList.KeyWordToString[keyword] }, StringSplitOptions.None);
                var RawValue = GetValue(KeywordsEnumList.KeyWordToString[keyword] + valueTable[1]);
                var channelInputInterpretation = KeywordParser.ChannelInputInterpretations[functionName];
                value = (float)channelInputInterpretation.Interpollation(RawValue);
                value = GetValueWithKeyWord(channelVout.Channel, value, keyword);
            } else
            {
                value = (float)channelVout.DefaultValue;
            }
            
            channelVout.CurrentValue = value;
            channelInputInfos.Add(channelVout);
        }


        private void CalculateFrequencyChannelValue(string ChannelName, IEnumerable<ChannelInputInfo> channels,
                                                    List<ChannelInputInfo> channelInputInfos, KeywordsEnum keyword, bool resetChannel)
        {
            var channelInputInfoFout = channels.Where(c => c.HwValue == "fout").FirstOrDefault();
            if (channelInputInfoFout != null)
            {
                channelInputInfoFout.CurrentValue = RelayState.CLOSE;
                channelInputInfos.Add(channelInputInfoFout);
            }

            var channelFreqOut = channels.Where(ch => ch.Channel.Kind == ChannelKind.FRQOUT).FirstOrDefault();

            float value = 0;
            if (!resetChannel)
            {
                var vehiculespeed = GetValue(ChannelName);

                value = (float)(vehiculespeed * channelFreqOut.Ratio * channelFreqOut.Pulse / 3600);
                value = GetValueWithKeyWord(channelFreqOut.Channel, value, keyword);
            } else
            {
                if (channelFreqOut.DefaultValue != null)
                {
                    value = (float)channelFreqOut.DefaultValue;
                }
            }

            channelFreqOut.CurrentValue = value;
            channelInputInfos.Add(channelFreqOut);
        }

        protected float GetValueWithKeyWord(Channel channel, float value, KeywordsEnum keyword, decimal CanScale = 0, bool isPSU = false)
        {
            float gap = 0;
            bool percent = false;
            string gapString = "";
            float gapValue = 0;
            bool useGapValue = false;

            if (channel != null)
            {
                switch (channel.Kind)
                {
                    case ChannelKind.DCVOUT:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERVALUE"];
                        gapValue = (float)InputParser.VBat;
                        break;
                    case ChannelKind.FRQOUT:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITFREQUENCYVALUE"];
                        gapValue = 5;
                        break;
                    case ChannelKind.DOUT:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITDIGITVALUE"];
                        break;
                    case ChannelKind.CAN:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITCANDECIMALVALUEVALUE"];
                        gapValue = (float)CanScale;
                        break;
                    default:
                        throw new Exception("No match found");
                        break;
                }
            }
            else if (isPSU)
            {
                gapString = IniFile.IniDataRaw["Channels"]["LIMITPSUVALUE"];
            }

            if (gapString.Contains("%"))
            {
                gapString = gapString.Replace("%", "");
                percent = true;
            }

            gap = float.Parse(gapString);

            if (percent)
            {
                gap = gap / 100;
            }

            if (channel != null && channel.Kind == ChannelKind.DCVOUT)
            {
                gapValue *= gap;
            }

            if (value * gap < gapValue)
                useGapValue = true;



            if ((keyword & KeywordsEnum.INF) == KeywordsEnum.INF)
            {
                value -= (useGapValue) ? gapValue : (value * gap);
            }
            else if ((keyword & KeywordsEnum.SUP) == KeywordsEnum.SUP) {
                value += (useGapValue) ? gapValue : (value * gap);
            }

            return value;
        }

        private float ParsePSU(string ChannelName, FlowChartState state, bool resetChannel)
        {
            float value;

            if (!resetChannel)
            {
                value = float.Parse(Regex.Match(ChannelName, @"\d+").Value);
                ChannelName = ReplaceKeyword(ChannelName);

                var keyword = KeywordsEnumList.GetKeyword(ChannelName);

                if (keyword != KeywordsEnum.NOKEYWORD)
                {
                    value = GetValueWithKeyWord(null, value, keyword, 0, true);
                }
            }
            else
            {
                value = InputParser.PSUV;
            }

            BuildPsuAction(state, value);
            return value;
        }

        private void ParseDelay(string ChannelName, FlowChartState state)
        {
            float value;

            ChannelName = ReplaceKeyword(ChannelName);
            value = float.Parse(Regex.Match(ChannelName, @"\d+[.]?\d*").Value);
            value = KeywordsEnumList.CalculateTIme(ChannelName, value);

            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayOf + " " + value + "ms", (int)value);
        }

        private void ParseUserAction(string ChannelName, FlowChartState state, bool reset)
        {
            var regExp = @".*useraction\((.*)\).*";

            Match match = Regex.Match(ChannelName, regExp);

            if (match.Success)
            {
                ProjectUtilsFunction.BuildUserAction(state, ((reset)?"reset ":"") + match.Groups[1].Value, UserActionButtons.Pass | UserActionButtons.Fail);
            } else
            {
                throw new ParserException(ParserExceptionType.USERACTION_NOT_WELL_FORMATED);
            }
        }

            private void BuildPsuAction(FlowChartState state, float value)
        {
            if (value > float.Parse(IniFile.IniDataRaw["Channels"]["LIMITPSUVALUEMAX"]))
            {
                throw new Exception(Resources.lang.ErrorTheVoltageSetIsTooHigh);
            }

            InputParser.VBat = (decimal)value;

            var Units = project.Workspaces[0].Units;

            foreach (var Unit in Units)
            {
                var psuAction = new PowerSupplyAction();
                psuAction.Current = InputParser.PSUC;
                psuAction.Description = Resources.lang.PsuAction + " " + Unit.Unit.Name + " " + value;
                psuAction.Voltage = value;
                psuAction.GroupId = Unit.GroupId;
                psuAction.Activate = true;

                state.Actions.Add(psuAction);
            }
        }

        protected void BuildFlashOutputOperationHighLow(string function, string result, FlowChartState state, 
                                                        string testInformation, bool checkResultFalse, ChannelOutputInfo channelInfo = null)
        {
            if (channelInfo == null)
            {
                channelInfo = GetChannelOutputInfo(function);
            }
            var channel = channelInfo.Channel;

            var channelActionCapture = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCaptureHighLowStart + " " + function + testInformation, channel, CaptureType.HIGHLOW, true, channelInfo.Board);
            channelActionCapture.Threshold1 = (float)InputParser.VBat / 2;
            channelActionCapture.Threshold2 = (float)InputParser.VBat / 2;

            state.Actions.Add(channelActionCapture);

            var regExp = @"(\d+[.]?\d*)s: on, (\d+[.]?\d*)s: off";

            Match match = Regex.Match(result, regExp);

            var ONTime = decimal.Parse(match.Groups[1].Value);
            var OFFTime = decimal.Parse(match.Groups[2].Value);

            var fullTime = (ONTime + OFFTime) * 2;

            //delai
            ProjectUtilsFunction.BuildDelay(state, "Delay for capture HIGHLOW of " + function + testInformation, (int)fullTime * 1000);

            var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCaptureHighLowStop + " " + function + testInformation, channel, CaptureType.HIGHLOW, false, channelInfo.Board);
            channelActionCaptureStop.Threshold1 = (float)InputParser.VBat / 2;
            channelActionCaptureStop.Threshold2 = (float)InputParser.VBat / 2;

            var limitInfo = GetLimitInfo(channel, function, result, true);
            Limit limitHigh = null;
            Limit limitLow = null;

            if (checkResultFalse && result == KeywordsEnum.OFF.ToString().ToLower())
            {
                limitHigh = BuildLimit(limitInfo, function + "=" + 0, channelInfo, true, ONTime * 1000);
                limitLow = BuildLimit(limitInfo, function + "=" + 0, channelInfo, true, OFFTime * 1000);
            }
            else
            {
                if (ONTime != 0)
                {
                    limitHigh = BuildLimit(limitInfo, function + "=" + ONTime * 1000, channelInfo, true, ONTime * 1000);
                } else
                {
                    limitHigh = BuildLimit(limitInfo, function + "=" + ONTime * 1000, channelInfo, true, (decimal)1000);
                }
                if (OFFTime != 0)
                {
                    limitLow = BuildLimit(limitInfo, function + "=" + OFFTime * 1000, channelInfo, true, OFFTime * 1000);
                } else
                {
                    limitLow = BuildLimit(limitInfo, function + "=" + OFFTime * 1000, channelInfo, true, (decimal)1000);
                }
            }

            channelActionCaptureStop.Limits.Add(limitHigh);
            channelActionCaptureStop.ExtraLimits.Add(limitHigh);
            state.Actions.Add(channelActionCaptureStop);
        }

        protected void BuildFlashOutputOperationMax(string function, string result, FlowChartState state, string testInformation,
                                                    bool checkResultFalse, ChannelOutputInfo channelInfo = null)
        {
            if (channelInfo == null)
            {
                channelInfo = GetChannelOutputInfo(function);
            }
           
            var channel = channelInfo.Channel;

            var channelActionCapture = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCaptureMaxStart + " " + function + testInformation, channel, CaptureType.MAX, true, channelInfo.Board);

            state.Actions.Add(channelActionCapture);

            var regExp = @"(\d+[.]?\d*)s: on, (\d+[.]?\d*)s: off";

            Match match = Regex.Match(result, regExp);

            var ONTime = decimal.Parse(match.Groups[1].Value);
            var OFFTime = decimal.Parse(match.Groups[2].Value);

            var fullTime = (ONTime + OFFTime) * 2;

            //delai
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayForCaptureMaxOf + " " + function + testInformation, (int)fullTime * 1000);

            var channelActionCaptureStop = ProjectUtilsFunction.BuildCapture(Resources.lang.ChannelActionCaptureMaxStop  + " " + function + testInformation, channel, CaptureType.MAX, false, channelInfo.Board);

            var limitInfo = GetLimitInfo(channel, function, result, true);

            Limit limit = null;
            if (ONTime == 0 && OFFTime == 0)
            {
                limit = BuildLimit(limitInfo, function + "off", channelInfo);
            }
            else
            {
                limit = BuildLimit(limitInfo, function + "=" + InputParser.VBat, channelInfo);
            }

            channelActionCaptureStop.Limits.Add(limit);

            state.Actions.Add(channelActionCaptureStop);
        }

        protected bool IsCanChannel(string ChannelName)
        {
            ChannelName = ChannelName.ToLower();

            var limit = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && ChannelName.Contains(l.SignalName));

            if (limit != null)
            {
                return true;
            }

            return false;
        }

        protected string GetCanChannelName(string ChannelName)
        {
            ChannelName = ChannelName.ToLower();

            var limit = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && ChannelName.Contains(l.SignalName));

            if (limit != null)
            {
                return limit.MessageName;
            }

            return "";
        }

        private void CreateCanTxFromSignalInfo(LimitSignalInfo limitCanInfo, FlowChartState state)
        {
            var channelAction = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.SendDefaultCanLoop + " " + limitCanInfo.MessageName);

            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.Channel = limitCanInfo.Channel;
            channelActionContainer.MainBoard = limitCanInfo.Board;
            channelAction.ChannelActions.Add(channelActionContainer);

            ProjectUtilsFunction.BuildCanTxStd(channelActionContainer, limitCanInfo.DefaultValue, limitCanInfo.MessageName, limitCanInfo.Variable);
        }

        protected float CreateCanTxFromSignalList(string ChannelName, FlowChartState state, bool resetChannel)
        {
            var actionEffectued = false;
            float sentValue = -1;
            var limitInfos = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && ChannelName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX && l.Priority == 0);
            if (limitInfos != null && limitInfos.Count() > 0)
            {
                if (limitInfos.FirstOrDefault().SignalName.Contains("dm"))
                {
                    actionEffectued = DmCanTxAction(ChannelName, limitInfos, state, resetChannel);
                }
                else
                {
                    if (!resetChannel)
                    {
                        actionEffectued = StandardCanTxAction(ChannelName, limitInfos.FirstOrDefault(), state, out sentValue);
                    }
                }
            }
            if (!actionEffectued && !resetChannel)
            {
                throw new ParserException(ParserExceptionType.CAN_CHANNEL_NOT_FOUND);
            }

            return sentValue;
        }

        private bool DmCanTxAction(string channelName, IEnumerable<LimitSignalInfo> limitInfo, FlowChartState state, bool resetChannel)
        {
            if (channelName.Contains("singleframe"))
            {
                BuildSingleFrameTest(channelName, state, limitInfo.ElementAt(0), resetChannel);
            } else if (channelName.Contains("multipleframe")) {
                BuildMultipleFramTEst(channelName, state, limitInfo, resetChannel);
            }

            return true;
        }

        public void BuildSingleFrameTest(string channelName, FlowChartState state, LimitSignalInfo frame, bool resetChannel)
        {
            Regex regexSPN = new Regex(@".*spn[=＝](\d+).*");
            Regex regexFMI = new Regex(@".*fmi[=＝](\d+).*");
            Regex regexLAMP = new Regex(@".*lamp[=＝](\d+).*");

            var data = "0X000000000001FFFF";

            string SPN = null;
            string FMI = null;
            string LAMP = null;

            Match matchSPN = regexSPN.Match(channelName);
            Match matchFMI = regexFMI.Match(channelName);
            Match matchLAMP = regexLAMP.Match(channelName);

            if (matchSPN.Success)
            {
                SPN = matchSPN.Groups[1].ToString();
            }
            if (matchFMI.Success)
            {
                FMI = matchFMI.Groups[1].ToString();
            }
            if (matchLAMP.Success)
            {
                LAMP = matchLAMP.Groups[1].ToString();
            }
            
            var channelAction = ProjectUtilsFunction.BuildChannelAction(state, resetChannel? Resources.lang.ResetDM1FrameForTest : Resources.lang.SendDM1FrameForTest + " " + channelName);

            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.Channel = frame.Channel;
            channelActionContainer.MainBoard = frame.Board;
            channelAction.ChannelActions.Add(channelActionContainer);

            BitArray ValueSend = ConvertHexToBitArray(data.Replace("0X", ""));
            BitArray valueToSet = new BitArray(1);

            if (SPN != null && !resetChannel) {
                valueToSet = new BitArray(new int[] { int.Parse(SPN) });
                data = SetCanData(16, 19, frame.Endian == "motorola", ValueSend, valueToSet, data, true);
            }

            if (FMI != null && !resetChannel)
            {
                valueToSet = new BitArray(new int[] { int.Parse(FMI) });
                data = SetCanData(32, 5, frame.Endian == "motorola", ValueSend, valueToSet, data);
            }

            if (LAMP != null && !resetChannel)
            {
                valueToSet = new BitArray(new int[] { int.Parse(LAMP) });
                data = SetCanData(0, 16, frame.Endian == "motorola", ValueSend, valueToSet, data);
            }

            ProjectUtilsFunction.BuildCanTxStd(channelActionContainer, data, frame.MessageName, frame.Variable);

        }

        public void BuildMultipleFramTEst(string channelName, FlowChartState state, IEnumerable<LimitSignalInfo> limitInfo, bool resetChannel)
        {
            LimitSignalInfo singleFrame = limitInfo.ElementAt(0);
            LimitSignalInfo headerFrame = limitInfo.ElementAt(2);
            LimitSignalInfo dataFrame = limitInfo.ElementAt(1);

            Regex regexSPN = new Regex(@".*spn[=＝](\d+).*");
            Regex regexFMI = new Regex(@".*fmi[=＝](\d+).*");
            Regex regexLAMP = new Regex(@".*lamp[=＝](\d+).*");

            var data = "0X0000000000000100000001FFFFFFFFF";
            var dataHeader = "0X200A0002FF" + singleFrame.CanId.Substring(4,2) + singleFrame.CanId.Substring(2, 2) + "00";

            string SPN = null;
            string FMI = null;
            string LAMP = null;

            Match matchSPN = regexSPN.Match(channelName);
            Match matchFMI = regexFMI.Match(channelName);
            Match matchLAMP = regexLAMP.Match(channelName);

            if (matchSPN.Success)
            {
                SPN = matchSPN.Groups[1].ToString();
            }
            if (matchFMI.Success)
            {
                FMI = matchFMI.Groups[1].ToString();
            }
            if (matchLAMP.Success)
            {
                LAMP = matchLAMP.Groups[1].ToString();
            }

            BitArray ValueSend = ConvertHexToBitArray(data.Replace("0X", ""));
            BitArray valueToSet = new BitArray(1);

            if (SPN != null && !resetChannel)
            {
                valueToSet = new BitArray(new int[] { int.Parse(SPN) });
                data = SetCanData(24, 19, dataFrame.Endian == "motorola", ValueSend, valueToSet, data, true);
                data = SetCanData(56, 19, dataFrame.Endian == "motorola", ValueSend, valueToSet, data, true);
            }

            if (FMI != null && !resetChannel)
            {
                valueToSet = new BitArray(new int[] { int.Parse(FMI) });
                data = SetCanData(40, 5, dataFrame.Endian == "motorola", ValueSend, valueToSet, data);
                data = SetCanData(72, 5, dataFrame.Endian == "motorola", ValueSend, valueToSet, data);
            }

            if (LAMP != null && !resetChannel)
            {
                valueToSet = new BitArray(new int[] { int.Parse(LAMP) });
                data = SetCanData(8, 16, dataFrame.Endian == "motorola", ValueSend, valueToSet, data);
            }



            var data1 = "0X01" + data.Replace("0X", "").Substring(2,14);
            var data2 = "0X02" + data.Replace("0X", "").Substring(16,14);


            var channelActionDataHeader = ProjectUtilsFunction.BuildChannelAction(state, resetChannel ? Resources.lang.ResetDM1HeaderFrameForTest : Resources.lang.SendDM1HeaderFrameForTest + " " + channelName);
            var channelActionContainerHeader = new ChannelActionContainer();
            channelActionContainerHeader.Channel = dataFrame.Channel;
            channelActionContainerHeader.MainBoard = dataFrame.Board;
            channelActionDataHeader.ChannelActions.Add(channelActionContainerHeader);

            ProjectUtilsFunction.BuildCanTxStd(channelActionContainerHeader, dataHeader, headerFrame.MessageName, headerFrame.Variable);

            var channelActionData1 = ProjectUtilsFunction.BuildChannelAction(state, resetChannel ? Resources.lang.ResetDM1Data1FrameForTest : Resources.lang.SendDM1Data1FrameForTest + " " + channelName);
            var channelActionContainerData1 = new ChannelActionContainer();
            channelActionContainerData1.Channel = dataFrame.Channel;
            channelActionContainerData1.MainBoard = dataFrame.Board;
            channelActionData1.ChannelActions.Add(channelActionContainerData1);

            ProjectUtilsFunction.BuildCanTxStd(channelActionContainerData1, data1, dataFrame.MessageName, dataFrame.Variable);

            var channelActionData2 = ProjectUtilsFunction.BuildChannelAction(state, resetChannel ? Resources.lang.ResetDM1Data2FrameForTest : Resources.lang.SendDM1Data2FrameForTest + " " + channelName);
            var channelActionContainerData2 = new ChannelActionContainer();
            channelActionContainerData2.Channel = dataFrame.Channel;
            channelActionContainerData2.MainBoard = dataFrame.Board;
            channelActionData2.ChannelActions.Add(channelActionContainerData2);

            ProjectUtilsFunction.BuildCanTxStd(channelActionContainerData2, data2, dataFrame.MessageName, dataFrame.Variable);
        }

        private bool StandardCanTxAction(string ChannelName, LimitSignalInfo limitInfo, FlowChartState state, out float sentValue)
        {
            Channel channel = null;
            var actionEffectued = false;
            var myChannelName = ChannelName;
            ChannelName = ChannelName.Replace(limitInfo.SignalName, "");
            sentValue = -1;

            var canId = limitInfo.CanId;
            var MessageName = limitInfo.MessageName;

            if (!ChannelName.ToLower().Contains("stop"))
            {
                float number;
                try
                {
                    number = float.Parse(Regex.Match(ChannelName, @"[-]?\d+[.]?\d*").Value);
                }
                catch
                {
                    throw new ParserException(ParserExceptionType.BAD_NUMBER_FORMAT);
                }

                int intCanId;
                var passed = int.TryParse(canId, NumberStyles.HexNumber, CultureInfo.CurrentCulture, out intCanId);

                if (passed)
                {
                    var CanMessage = product.Variables.Where(v => ((CanMessageVariableSettings)v.Settings).MessageId == intCanId).FirstOrDefault();
                    channel = ((CanMessageVariableSettings)CanMessage.Settings).Channel;
                    actionEffectued = true;
                    
                    ChannelName = ReplaceKeyword(ChannelName);

                    var keyword = KeywordsEnumList.GetKeyword(ChannelName);

                    if (keyword != KeywordsEnum.NOKEYWORD)
                    {
                        number = GetValueWithKeyWord(channel, number, keyword, limitInfo.A, true);
                    }
                    //Can Loop

                    var signalInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && myChannelName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
                    var alreadyStop = false;

                    //state.Actions.IndexOf
                    var lastCanLoopStart = state.Actions.LastOrDefault(a => a is CanLoopAction &&
                    ((CanLoopAction)a).Variable.Name == limitInfo.Variable.Name /*&&
                    !((CanLoopAction)a).Description.Contains("Default")*/ &&
                    ((CanLoopAction)a).CanLoopCmd == CanLoopCmd.LOOP) as CanLoopAction;
                    var lastCanLoopStop = state.Actions.LastOrDefault(a => a is CanLoopAction
                    && ((CanLoopAction)a).Variable.Name == limitInfo.Variable.Name /*&&
                    !((CanLoopAction)a).Description.Contains("Default")*/ &&
                    ((CanLoopAction)a).CanLoopCmd == CanLoopCmd.STOPLOOP) as CanLoopAction;

                    var indexofStart = (lastCanLoopStart != null) ? state.Actions.IndexOf(lastCanLoopStart) : -1;
                    var indexofStop = (lastCanLoopStop != null) ? state.Actions.IndexOf(lastCanLoopStop) : -1;

                    var checkMessageExist = canLoopSignalUsed.FirstOrDefault(si => si.Variable.Name == signalInfo.Variable.Name);

                    if (canLoopSignalUsed.Contains(signalInfo))
                    {
                        alreadyStop = true;
                        var stopLastCanLoop = new CanLoopAction();
                        stopLastCanLoop.CanLoopCmd = CanLoopCmd.STOPLOOP;
                        stopLastCanLoop.Description = Resources.lang.StopDefaultCanLoop + " " + limitInfo.MessageName;
                        if (ChannelName.Contains("b"))
                        {
                            SetCanLoop(stopLastCanLoop, signalInfo, Regex.Match(ChannelName, @"\d+").Value);
                        }
                        else
                        {
                            sentValue = SetCanLoop(stopLastCanLoop, myChannelName, number);
                        }
                        state.Actions.Add(stopLastCanLoop);
                    }
                    else
                    {
                        canLoopSignalUsed.Add(signalInfo);
                    }

                    CanLoopAction canLoopAction;
                    if (/*lastCanLoopStart != null && indexofStart > indexofStop*/checkMessageExist != null && !alreadyStop)
                    {
                        lastCanLoopStart.Description += " & " + myChannelName;
                        if (ChannelName.Contains("b"))
                        {
                            SetCanLoop(lastCanLoopStart, signalInfo, Regex.Match(ChannelName, @"\d+").Value, lastCanLoopStart.Data.Remove(0, 2));
                        }
                        else
                        {
                            sentValue = SetCanLoop(lastCanLoopStart, myChannelName, number, lastCanLoopStart.Data.Remove(0, 2));
                        }
                        canLoopAction = lastCanLoopStart;
                    }
                    else
                    {
                        if (!alreadyStop)
                        {
                            var canLoopActionStopDefault = new CanLoopAction();
                            canLoopActionStopDefault.CanLoopCmd = CanLoopCmd.STOPLOOP;
                            canLoopActionStopDefault.Description = Resources.lang.StopDefaultCanLoop + " " + limitInfo.MessageName;
                            if (ChannelName.Contains("b"))
                            {
                                SetCanLoop(canLoopActionStopDefault, signalInfo, Regex.Match(ChannelName, @"\d+").Value);
                            }
                            else
                            {
                                sentValue = SetCanLoop(canLoopActionStopDefault, myChannelName, number);
                            }
                            state.Actions.Add(canLoopActionStopDefault);
                        }

                        canLoopAction = new CanLoopAction();
                        canLoopAction.CanLoopCmd = CanLoopCmd.LOOP;
                        canLoopAction.Description = Resources.lang.StartCanLoop + " " + myChannelName;
                        canLoopAction.Variable = limitInfo.Variable;
                        if (ChannelName.Contains("b"))
                        {
                            SetCanLoop(canLoopAction, signalInfo, Regex.Match(ChannelName, @"\d+").Value);
                        }
                        else
                        {
                            SetCanLoop(canLoopAction, myChannelName, number);
                        }
                        state.Actions.Add(canLoopAction);
                        currentCanLoopAction.Add(new Tuple<string, object>(myChannelName, number));
                    }

                    if (limitInfo.ForwardList != null && limitInfo.ForwardList.Count > 0)
                    {

                        //Delay
                        ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayBeforeCheckingCanForward + " " + myChannelName, limitInfo.Frequency * 2);

                        var channelAction = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ReceiveCanForward + " " + myChannelName);
                        var limitFwInfos = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && myChannelName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.RX);
                        //ForwardCheck
                        var e = 0;
                        foreach (var fowardLimitInfo in limitInfo.ForwardList)
                        {
                            var limitFwInfo = limitFwInfos.ElementAt(e);
                            e++;
                            var channelActionContainer = new ChannelActionContainer();
                            channelActionContainer.Channel = fowardLimitInfo.Channel;
                            channelActionContainer.MainBoard = fowardLimitInfo.Board;
                            channelAction.ChannelActions.Add(channelActionContainer);

                            var CanOperation = new CanOperation();
                            CanOperation.Action = CanAction.RECVMSG;
                            CanOperation.ByteNumber = 8;
                            CanOperation.MessageName = limitFwInfo.MessageName;
                            CanOperation.FeedbackCtrl = CanFeedbackCtrl.NEWESTMSG;
                            channelActionContainer.Operation = CanOperation;

                            var data = new SignalData();
                            data.HighLimit = 0;
                            data.LowLimit = 0;
                            data.Offset = (double)fowardLimitInfo.B;
                            data.Scale = (double)fowardLimitInfo.A;

                            var bitInfo = new SignalBitInfo();
                            bitInfo.Length = fowardLimitInfo.BitLength;
                            bitInfo.Position = fowardLimitInfo.StartBit;

                            var limit = new Limit();
                            CanOperation.Limits.Add(limit);

                            var errorMessage = new ErrorMessage();
                            limit.ErrorMessage = errorMessage;
                            limit.Name = Resources.lang.ForwaredCheckError + " " + myChannelName;
                            errorMessage.Name = Resources.lang.ForwaredCheckError + " " + myChannelName;
                            errorMessage.Severity = Severity.Error;

                            var comparison = new SignalComparison();
                            var EqualComparison = new EqualTextComparison();
                            EqualComparison.Text = canLoopAction.Data;
                            var limitContainer = new ComparisonContainer(EqualComparison, ComparisonKind.EQUALTEXT);

                            comparison.SignalDataType = DataType.Text;
                            comparison.BaseComparison = limitContainer;

                            limit.Container.Comparison = comparison;
                            limit.Container.ComparisonKind = ComparisonKind.SIGNAL;
                            limit.Container.Id = 1;

                            var signal = new Signal();
                            signal.Alias = myChannelName;
                            data.BitInfo = bitInfo;
                            signal.Data = data;

                            comparison.Signal = signal;
                        }
                    }
                }
                else
                {
                    throw new Exception("Error in channel ID");
                }
            }
            else
            {
                var canLoopActionStopDefault = new CanLoopAction();
                canLoopActionStopDefault.CanLoopCmd = CanLoopCmd.STOPLOOP;
                canLoopActionStopDefault.Description = Resources.lang.StopDefaultCanLoop + " " + limitInfo.MessageName;
                SetCanLoop(canLoopActionStopDefault, myChannelName, 0);
                state.Actions.Add(canLoopActionStopDefault);

                currentCanLoopAction.Add(new Tuple<string, object>(myChannelName, null));
                actionEffectued = true;
            }
            return actionEffectued;
        }

        protected void SetResultChannel(string ChannelName, string function, ChannelAction channelAction, bool forceOff = false)
        {
            //container by action
            ChannelActionContainer channelActionContainer = new ChannelActionContainer();

            channelActionContainer.IsUsed = true;
            ChannelName = ChannelName.ToLower();

            var canChannel = GetChannelCanMessage(ChannelName);
            var channelInfo = GetChannelOutputInfo(ChannelName);

            if (canChannel != null)
            {
                channelActionContainer.Channel = canChannel;
                CreateOutputOperation(channelActionContainer, function, ChannelName, new ChannelOutputInfo());
            } else if (channelInfo != null)
            {
                var myChannelName = ChannelName.Replace(channelInfo.Function, "");
                //foreach the board to check if they contain the channel
                channelActionContainer.MainBoard = channelInfo.Board;

                channelActionContainer.Channel = channelInfo.Channel;
                CreateOutputOperation(channelActionContainer, function, myChannelName, channelInfo);
            } else
            {
                throw new ParserException(ParserExceptionType.RESULT_CHANNEL_NOT_FOUND);
            }
            channelAction.ChannelActions.Add(channelActionContainer);
        }

        protected Board GetBoardFromChannel(Channel channel)
        {
            foreach (var board in product.Boards)
            {
                if (board.Channels.Contains(channel))
                {
                    return board;
                }
            }
            throw new Exception(Resources.lang.NoBoardFoundForThisChannel);
        }

        protected Channel GetChannelCanMessage(string ChannelName)
        {
            Channel channel = null;
            //Analyse ChannelName to add Can Message
            if (ChannelName.Contains("0x"))
            {
                var canId = ChannelName.Substring(ChannelName.IndexOf("0x") + 2, 8);
                int intCanId;
                var passed = int.TryParse(canId, NumberStyles.HexNumber, CultureInfo.CurrentCulture, out intCanId);

                if (passed)
                {
                    var CanMessage = product.Variables.Where(v => ((CanMessageVariableSettings)v.Settings).MessageId == intCanId).FirstOrDefault();
                    channel = ((CanMessageVariableSettings)CanMessage.Settings).Channel;

                }
                else
                {
                    throw new Exception(Resources.lang.ErrorInChannelID);
                }
            }
            return channel;
        }

        protected ChannelOutputInfo GetChannelOutputInfo(string ChannelName)
        {
            ChannelOutputInfo channelOutputInfo = null;

            foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
            {
                if (ChannelName.Contains(key.ToLower()) && key != "0")
                {
                    channelOutputInfo = KeywordParser.FunctionChannelOutputMap.Where(ch => KeywordParser.KeywordFunctionMap[key] == ch.Function && ch.Load != "SELECT").FirstOrDefault(); //[KeywordParser.KeywordFunctionMap[key]];
                    break;
                }
            }

            return channelOutputInfo;
        }

        protected ChannelInputInfo GetChannelInputInfo(string ChannelName)
        {
            ChannelInputInfo channelInputInfo = null;

            foreach (var key in KeywordParser.KeywordFunctionMap.Keys)
            {
                if (ChannelName.Contains(key.ToLower()) && key != "0")
                {
                    channelInputInfo = KeywordParser.FunctionChannelInputMap.Where(ch => KeywordParser.KeywordFunctionMap[key] == ch.Function).FirstOrDefault(); //[KeywordParser.KeywordFunctionMap[key]];
                    break;
                }
            }

            return channelInputInfo;
        }

        protected void CleanCanLoops(FlowChartState state, string testInformation = "", bool resetDefault = true)
        {
            foreach (var canLoopInfo in currentCanLoopAction)
            {
                if (canLoopInfo.Item2 != null)
                {
                    var canLoopAction = new CanLoopAction();
                    canLoopAction.CanLoopCmd = CanLoopCmd.STOPLOOP;
                    canLoopAction.Description = Resources.lang.StopCanLoop + " " + canLoopInfo.Item1 + testInformation;
                    if (canLoopInfo.Item2.ToString().Contains("b"))
                    {
                        var signalInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && canLoopInfo.Item1.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
                        SetCanLoop(canLoopAction, signalInfo, Regex.Match(canLoopInfo.Item2.ToString(), @"\d+").Value);
                    }
                    else
                    {
                        SetCanLoop(canLoopAction, canLoopInfo.Item1, (float)canLoopInfo.Item2);
                    }
                    state.Actions.Add(canLoopAction);
                }

                if (resetDefault)
                {
                    StartDefaultCanLoop(state, canLoopInfo.Item1);
                }
            }

            currentCanLoopAction.Clear();
        }

        protected void CleanCanLoop(FlowChartState state, string name, bool resetDefault = true)
        {
            Tuple<string, object> canLoopInfo = null;
            foreach (var canLoop in currentCanLoopAction)
            {
                if (canLoop.Item1.Contains(name)) {
                    if (canLoop.Item2 != null)
                    {
                        var canLoopAction = new CanLoopAction();
                        canLoopAction.CanLoopCmd = CanLoopCmd.STOPLOOP;
                        canLoopAction.Description = Resources.lang.StopCanLoop + " " + canLoop.Item1;

                        if (canLoop.Item2.ToString().Contains("b"))
                        {
                            var signalInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && canLoop.Item1.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
                            SetCanLoop(canLoopAction, signalInfo, Regex.Match(canLoop.Item2.ToString(), @"\d+").Value);
                        }
                        else
                        {
                            SetCanLoop(canLoopAction, canLoop.Item1, (float)canLoop.Item2);
                        }

                        state.Actions.Add(canLoopAction);
                    }

                    if (resetDefault)
                    {
                        StartDefaultCanLoop(state, canLoop.Item1);
                    }
                    canLoopInfo = canLoop;
                    break;
                }
            }

            if (canLoopInfo != null) {
                currentCanLoopAction.Remove(canLoopInfo);
            }
        }

        protected void StartDefaultCanLoop(FlowChartState state, string signalName)
        {
            var limitCanInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && signalName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
            if (limitCanInfo.DefaultValue != null && limitCanInfo.DefaultValue != "")
            {
                if (limitCanInfo.ResetLoop)
                {
                    var canLoopActionStopDefault = new CanLoopAction();
                    canLoopActionStopDefault.CanLoopCmd = CanLoopCmd.LOOP;
                    canLoopActionStopDefault.Description = Resources.lang.StartDefaultCanLoop + " " + limitCanInfo.MessageName;
                    BuildCanLoop(canLoopActionStopDefault, signalName, limitCanInfo.DefaultValue, limitCanInfo);
                    state.Actions.Add(canLoopActionStopDefault);
                }
                else
                {
                    CreateCanTxFromSignalInfo(limitCanInfo, state);
                }
            }
        }


        protected void CreateInputOperation(ChannelActionContainer channelActionContainer, string function, string ChannelName, 
            ChannelInputInfo channelInputInfo, bool setDefaultValue, bool ResetInitialValue = false)
        {
            var channel = channelActionContainer.Channel;

            switch (channel.Kind)
            {
                case ChannelKind.AWGOUT:
                    ProjectUtilsFunction.BuildAwgOut(channelActionContainer, AwgAction.START, AwgFunction.SQUARE,
                                1, 1, (float)channelInputInfo.CurrentValue, (float)channelInputInfo.CurrentValue / 2);
                    break;
                case ChannelKind.CAN:
                    CanOperation canOperation = null;
                    if (channelActionContainer.Operation == null)
                    {
                        canOperation = new CanOperation();
                        channelActionContainer.Operation = canOperation;
                    }
                    else
                    {
                        canOperation = (CanOperation)channelActionContainer.Operation;
                    }

                    BuildCanChannelOperation(canOperation, ChannelName);
                    break;
                case ChannelKind.DCVOUT:
                    var VoltageOperation = new VoltageOutOperation();
                    channelActionContainer.Operation = VoltageOperation;
                    VoltageOperation.Value = (float)channelInputInfo.CurrentValue;

                    break;
                case ChannelKind.FRQOUT:
                    var FreqOutOperation = new FreqOutOperation();
                    FreqOutOperation.Frequency = (float)channelInputInfo.CurrentValue;
                    FreqOutOperation.DutyCycle = 0.5f;
                    FreqOutOperation.OffVoltage = 0;
                    FreqOutOperation.OnVoltage = 6;
                    channelActionContainer.Operation = FreqOutOperation;
                    break;
                case ChannelKind.RELAY:
                    var RelayOperation = new RelayOperation();
                    channelActionContainer.Operation = RelayOperation;

                    if (ResetInitialValue)
                    {
                        if (channelInputInfo.DefaultValue != null)
                        {
                            RelayOperation.State = (RelayState)channelInputInfo.DefaultValue;
                        }
                        else
                        {
                            RelayOperation.State = RelayState.OPEN;
                        }
                    }
                    else
                    {
                        //basic close
                        if (channelInputInfo.CurrentValue != null)
                        {
                            RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                        }
                        ChannelName = ReplaceKeyword(ChannelName);
                        //complexe reset
                        var keyword = KeywordsEnumList.GetKeyword(ChannelName);

                        if (keyword != KeywordsEnum.NOKEYWORD)
                        {
                            if ((keyword & KeywordsEnum.ON) == KeywordsEnum.ON)
                            {
                                if ((channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in") &&
                                    channelInputInfo.HwValue == "gnd")
                                {
                                    channelInputInfo.CurrentValue = RelayState.CLOSE;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (((channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in") &&
                                  channelInputInfo.HwValue == "vbat"))
                                {
                                    channelInputInfo.CurrentValue = RelayState.OPEN;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (((channelInputInfo.LogicalContact == "vbat" || channelInputInfo.LogicalContact == "hs_in") &&
                                  channelInputInfo.HwValue == "gnd"))
                                {
                                    channelInputInfo.CurrentValue = RelayState.OPEN;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (channelInputInfo.LogicalContact == "vbat" || channelInputInfo.LogicalContact == "hs_in")
                                {
                                    channelInputInfo.CurrentValue = RelayState.CLOSE;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in")
                                {
                                    channelInputInfo.CurrentValue = RelayState.OPEN;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else
                                {
                                    throw new ParserException(ParserExceptionType.RELAY_LOGIC_NOT_FOUND);
                                }
                            }
                            else if ((keyword & KeywordsEnum.OFF) == KeywordsEnum.OFF)
                            {
                                if (((channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in") &&
                                    channelInputInfo.HwValue == "vbat"))
                                {
                                    channelInputInfo.CurrentValue = RelayState.CLOSE;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if ((channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in") &&
                                    channelInputInfo.HwValue == "gnd")
                                {
                                    channelInputInfo.CurrentValue = RelayState.OPEN;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (((channelInputInfo.LogicalContact == "vbat" || channelInputInfo.LogicalContact == "hs_in") &&
                                  channelInputInfo.HwValue == "gnd"))
                                {
                                    channelInputInfo.CurrentValue = RelayState.CLOSE;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (channelInputInfo.LogicalContact == "vbat" || channelInputInfo.LogicalContact == "hs_in")
                                {
                                    channelInputInfo.CurrentValue = RelayState.OPEN;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else if (channelInputInfo.LogicalContact == "gnd" || channelInputInfo.LogicalContact == "ls_in")
                                {
                                    channelInputInfo.CurrentValue = RelayState.CLOSE;
                                    RelayOperation.State = (RelayState)channelInputInfo.CurrentValue;
                                }
                                else
                                {
                                    throw new ParserException(ParserExceptionType.RELAY_LOGIC_NOT_FOUND);
                                }
                            }

                            if (setDefaultValue)
                            {
                                channelInputInfo.BaseDefaultValue = channelInputInfo.DefaultValue;
                                channelInputInfo.DefaultValue = channelInputInfo.CurrentValue;
                            }
                        }
                    }

                    break;
                case ChannelKind.None:
                default:
                    break;
            }
        }

        private void CreateOutputOperation(ChannelActionContainer channelActionContainer, string function, string ChannelName, ChannelOutputInfo channelOutputInfo)
        {
            var channel = channelActionContainer.Channel;
            ChannelName = ReplaceKeyword(ChannelName);

            switch (channel.Kind)
            {
                case ChannelKind.VOLTMETER:
                    var keyword = KeywordsEnumList.GetKeyword(ChannelName);

                    if (keyword != KeywordsEnum.NOKEYWORD)
                    {
                        var VoltmeterOperation = new VoltmeterMeasureOperation();
                        channelActionContainer.Operation = VoltmeterOperation;

                        var limitInfo = GetLimitInfo(channel, function, ChannelName);
                        var limit = BuildLimit(limitInfo, ChannelName, channelOutputInfo);
                        product.ErrorMessages.Add(limit.ErrorMessage);
                        VoltmeterOperation.Limits.Add(limit);
                        break;
                    }
                    break;
                case ChannelKind.None:
                default:
                    break;
            }
        }

        protected string ReplaceKeyword(string testString)
        {
            foreach (var key in KeywordParser.StatusTestMap.Keys)
            {
                testString = testString.Replace(key, KeywordParser.StatusTestMap[key]);
            }

            foreach (var key in KeywordParser.Keywords.Keys)
            {
                testString = testString.Replace(key, KeywordParser.Keywords[key]);
            }

            return testString;
        }

        protected void BuildCanLoop(CanLoopAction canLoop, string name, string data, LimitSignalInfo signalInfo)
        {
            canLoop.Data = data;
            canLoop.Channel = signalInfo.Channel;
            canLoop.CurrentBoard = signalInfo.Board;
            canLoop.Variable = signalInfo.Variable;
            canLoop.Interval = signalInfo.Frequency;
            canLoop.ImagePath = null;
        }

        protected float SetCanLoop(CanLoopAction canLoop, string name, float number, string data = "0000000000000000")
        {
            var signalInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && name.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX && l.Priority == 0);
            return SetNumericCanLoop(canLoop, signalInfo, number, data);
        }

        protected float SetNumericCanLoop(CanLoopAction canLoop, LimitSignalInfo signalInfo, float number, string data)
        {
            int value = 0;
            float recalculatedValue = 0;
            var continousLimitInfos = MessageParser.MessageParser.LimitSignalInfoList.Where(l => l.SignalName != null && l.SignalName == signalInfo.SignalName && l.Direction == MessageFlowDirection.TX);
            if (signalInfo != null)
            {
                var canId = signalInfo.CanId;
                var MessageName = signalInfo.MessageName;
                canLoop.Variable = signalInfo.Variable;

                var a = signalInfo.A;
                var b = signalInfo.B;

                if (signalInfo.FormulaLogic == 1)
                {
                    value = (int)(((decimal)number - b) / a);
                } else if (signalInfo.FormulaLogic == 2)
                {
                    if (((decimal)number - b) == 0)
                    {
                        throw new ParserException(ParserExceptionType.CAN_CALCULATION_DIVISION_BY_ZERO);
                    }
                    value = (int)(a/((decimal)number - b));
                } else
                {
                    throw new ParserException(ParserExceptionType.FORMULAT_LOGIC_NOT_FOUND);
                }

                BitArray ValueSend = ConvertHexToBitArray(data);
                BitArray valueToSet = new BitArray(new int[] { value });

                if (value > (Math.Pow(2, signalInfo.BitLength) - 1))
                {
                    valueToSet.SetAll(true);
                    value = (int)Math.Pow(2, signalInfo.BitLength) - 1;
                }

                if (signalInfo.FormulaLogic == 1)
                {
                    recalculatedValue = (float)(a * value + b);
                }
                else if (signalInfo.FormulaLogic == 2) {
                    recalculatedValue = (float)(a / value + b);
                }

                data = SetCanData(signalInfo.StartBit, signalInfo.BitLength, signalInfo.Endian == "motorola", ValueSend, valueToSet, data);

                canLoop.Data = data;

                canLoop.Channel = signalInfo.Channel;
                canLoop.CurrentBoard = signalInfo.Board;
                canLoop.Interval = signalInfo.Frequency;
                canLoop.ImagePath = null;

                data = data.Remove(0, 2);

                if (continousLimitInfos.Count() > signalInfo.Priority + 1 && number - recalculatedValue > 0)
                {
                    var nextSignalInfo = continousLimitInfos.FirstOrDefault(si => si.Priority == signalInfo.Priority + 1);
                    recalculatedValue += SetNumericCanLoop(canLoop, nextSignalInfo, number - (float)((value*a) + b), data);
                }
            }
            return recalculatedValue;
        }

        protected void SetCanLoop(CanLoopAction canLoop, LimitSignalInfo signalInfo, string value, string data = "0000000000000000")
        {
            if (signalInfo != null)
            {
                var canId = signalInfo.CanId;
                var MessageName = signalInfo.MessageName;
                canLoop.Variable = signalInfo.Variable;

                BitArray ValueSend = ConvertHexToBitArray(data);
                BitArray valueToSet = ConvertBitStringToBitArray(StringUtils.Reverse(value));

                data = SetCanData(signalInfo.StartBit, signalInfo.BitLength, signalInfo.Endian == "motorola", ValueSend, valueToSet, data);

                canLoop.Data = data;

                canLoop.Channel = signalInfo.Channel;
                canLoop.CurrentBoard = signalInfo.Board;
                canLoop.Interval = signalInfo.Frequency;
                canLoop.ImagePath = null;
                data = data.Remove(0, 2);
            }
        }

        private string SetCanData(int i, int length, bool isMotorolla, BitArray ValueSend, BitArray valueToSet, string data, bool isSpN = false)
        {
            var bitShift = 0;
            for (int e = 0; e < length; e++)
            {
                ValueSend[i + bitShift] = valueToSet[e];
                i++;
                if (isMotorolla && i != 0 && i % 8 == 0)
                {
                    i -= 8 * 2;
                }

                if (isSpN && e == 15)
                {
                    bitShift = 5;
                }
            }

            Byte[] ByteArray = new byte[(int)Math.Ceiling(ValueSend.Count / 8.0)];
            ValueSend.CopyTo(ByteArray, 0);

            data = "0X" + BitConverter.ToString(ByteArray).Replace("-", "");

            return data;
        }

        protected void BuildCanChannelOperation(CanOperation canOperation, string channelName)
        {
            /*foreach (var SignalName in MessageParser.MessageParser.LimitSignalInfoList.Keys)
            {
                if (channelName.Contains(SignalName))
                {*/
            var signalInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && channelName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
            if (signalInfo != null)
            {

                if (signalInfo.Direction == MessageFlowDirection.RX)
                {
                    BuildCanChannelOperationRx(canOperation, channelName, signalInfo);
                }
                else
                {
                    BuildCanChannelOperationTx(canOperation, channelName, signalInfo);
                }
            }
                /*}
            }*/
        }

        public BitArray ConvertHexToBitArray(string data)
        {
            if (data == null)
                throw new Exception("error with CAN");

            string hexData = "";

            for (var i = 0; i < data.Length - 1; i = i + 2)
            {
                hexData += "" + data[i + 1] + "" + data[i];
            }


            BitArray ba = new BitArray(4 * hexData.Length);
            for (int i = 0; i < hexData.Length; i++)
            {
                byte b = byte.Parse(hexData[i].ToString(), NumberStyles.HexNumber);
                var byteString = new BitArray(new byte[] { b });
                //var f = 3;
                for (int j = 0; j < 4; j++)
                {
                    ba.Set(i * 4 + j, (b & (1 << j)) != 0/*byteString[j]*/);
                    // f--;
                }
            }
            return ba;
        }

        public BitArray ConvertBitStringToBitArray(string data)
        {
            if (data == null)
                throw new Exception("error with CAN");

            BitArray ba = new BitArray(data.Length);
            for (int i = 0; i < data.Length; i++)
            {
                if (data[i] == '1')
                {
                    ba.Set(i, true);
                }
                else if (data[i] == '0')
                {
                    ba.Set(i, false);
                }
                else
                {
                    throw new ParserException(ParserExceptionType.VALUE_NOT_FOUND);
                }
            }
            return ba;
        }

        private void BuildCanChannelOperationTx(CanOperation canOperation, string channelName, LimitSignalInfo signalInfo)
        {
            var nb = GetValueFromString(channelName);

            var data = "0000000000000000";

            if (canOperation.Data != null) {
                data = canOperation.Data.Replace("0x", "");
            }


            var dataToReplace = (((int)Math.Round(((nb - signalInfo.B) / signalInfo.A))));

            BitArray bitArray = ConvertHexToBitArray(data);

            BitArray valueToSet = new BitArray(new int[] { dataToReplace });

            var e = 0;
            for (int i = signalInfo.StartBit; i < (signalInfo.StartBit + signalInfo.BitLength); i++)
            {
                bitArray[i] = valueToSet[e];
                e++;
            }

            Byte[] ByteArray = new byte[(int)Math.Ceiling(bitArray.Count / 8.0)];
            bitArray.CopyTo(ByteArray, 0);

            data = "0x" + BitConverter.ToString(ByteArray).Replace("-", ""); ;

            canOperation.Action = CanAction.SENDMSGVALUE;
            canOperation.Data = data;
            canOperation.MessageName = signalInfo.MessageName;
        }

        private void BuildCanChannelOperationRx(CanOperation canOperation, string channelName, LimitSignalInfo signalInfo)
        {
            canOperation.Action = CanAction.RECVMSG;
            canOperation.ByteNumber = 8;
            canOperation.MessageName = signalInfo.MessageName;
            canOperation.Variable = signalInfo.Variable;

            var limit = new Limit();
            var comparison = new SignalComparison();

            var limitInfo = BuildCanValueLimitInfo(channelName);

            var comparisonContainer = CreateComparisonFromString(limitInfo, channelName, true);

            comparison.BaseComparison = comparisonContainer;
            var signal = new Signal();

            signal.Alias = channelName;

            var data = new SignalData();

            data.HighLimit = 0;
            data.LowLimit = 0;
            data.Offset = (double)signalInfo.B;
            data.Scale = (double)signalInfo.A;

            var bitInfo = new SignalBitInfo(); ;

            bitInfo.Length = signalInfo.BitLength;
            bitInfo.Position = signalInfo.StartBit;

            limit.Container.Comparison = comparison;
            limit.Container.ComparisonKind = ComparisonKind.SIGNAL;
            limit.Container.Id = 1;

            data.BitInfo = bitInfo;
            signal.Data = data;
            comparison.Signal = signal;

            canOperation.Limits.Add(limit);
        }

        protected LimitInfo BuildCanValueLimitInfo(string channelName)
        {
            var limitInfo = new LimitInfo();

            var gapString = IniFile.IniDataRaw["Channels"]["LIMITAMMETERVALUE"];
            limitInfo.ErrorMessage = channelName + " " + IniFile.IniDataRaw["Channels"]["LIMITAMMETERMESSAGE"];
            limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITAMMETERTYPE"], true);

            var percent = false;
            if (gapString.Contains("%"))
            {
                gapString = gapString.Replace("%", "");
                percent = true;
            }

            limitInfo.Gap = decimal.Parse(gapString);

            if (percent)
            {
                limitInfo.Gap = limitInfo.Gap / 100;
            }

            return limitInfo;
        }

        protected LimitInfo GetLimitInfo(Channel channel, string function, string op, bool isCapture = false)
        {
            LimitInfo limitInfo = new LimitInfo();
            bool percent = false;
            string gapString = "";

            if (isCapture)
            {
                gapString = IniFile.IniDataRaw["Channels"]["LIMITCAPTUREFLASHVALUE"];
                limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITCAPTUREFLASHMESSAGE"] + " : " + op;
                limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITCAPTUREFLASHTYPE"], true);
            }
            else
            {
                switch (channel.Kind)
                {
                    case ChannelKind.VOLTMETER:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERVALUE"];
                        limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERMESSAGE"] + " : " + op;
                        limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERTYPE"], true);
                        break;
                    case ChannelKind.FRQIN:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITFREQUENCYVALUE"];
                        limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITFREQUENCYMESSAGE"] + " : " + op;
                        limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITFREQUENCYTYPE"], true);
                        break;
                    case ChannelKind.DIN:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITDIGITVALUE"];
                        limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITDIGITMESSAGE"] + " : " + op;
                        limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITDIGITTYPE"], true);
                        break;
                    case ChannelKind.AMMETER:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITAMMETERVALUE"];
                        limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITAMMETERMESSAGE"] + " : " + op;
                        limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITAMMETERTYPE"], true);
                        break;
                    case ChannelKind.CAN:
                        gapString = IniFile.IniDataRaw["Channels"]["LIMITCANDECIMALVALUEVALUE"];
                        limitInfo.ErrorMessage = function + " " + IniFile.IniDataRaw["Channels"]["LIMITCANDECIMALVALUEMESSAGE"] + " : " + op;
                        limitInfo.Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITCANDECIMALVALUETYPE"], true);
                        break;
                    default:
                        throw new Exception("No match found");
                        break;
                }
            }
            if (gapString.Contains("%"))
            {
                gapString = gapString.Replace("%", "");
                percent = true;
            }

            limitInfo.Gap = decimal.Parse(gapString);

            if (percent)
            {
                limitInfo.Gap = limitInfo.Gap / 100;
            }

            return limitInfo;
        }

        protected Limit BuildLimit(LimitInfo limitInfo, string op, ChannelOutputInfo channelOutputInfo, bool gapFromValue = false, decimal value = 0)
        {
            var limit = new Limit();

            limit.Container = GetComparison(limitInfo, op, channelOutputInfo, gapFromValue, value);
            limit.ErrorMessage = new ErrorMessage();
            limit.ErrorMessage.Severity = limitInfo.Severity;
            limit.ErrorMessage.Name = limitInfo.ErrorMessage;
            limit.Name = limitInfo.ErrorMessage;

            return limit;
        }

        protected decimal GetValue(string op)
        {
            op = ReplaceKeyword(op);
            var keyword = KeywordsEnumList.GetKeyword(op);

            if (keyword != KeywordsEnum.NOKEYWORD)
            {
                if ((keyword & KeywordsEnum.ON) == KeywordsEnum.ON) {
                    return InputParser.VBat;
                } else if ((keyword & KeywordsEnum.OFF) == KeywordsEnum.OFF)
                {
                    return 0;
                }
            }
            return GetValueFromString(op);
        }

        protected Decimal GetValueFromString(string valueString)
        {
            decimal result;
            if (decimal.TryParse(valueString, out result))
            {
            }
            else
            {
                Regex regexp = new Regex(@".*[＝=><][ ]*([-]?\d+[.]?\d*)[ ]*.*");

                Match match = regexp.Match(valueString);
                if (match.Success)
                {
                    result = decimal.Parse(match.Groups[1].Value);
                }
            }
            return result;
        }

        protected ComparisonContainer GetComparison(LimitInfo limitInfo, string op, ChannelOutputInfo channelOutputInfo, bool gapFromValue = false, decimal valueGap = 0)
        {
            decimal value = 0;
            var containKeyword = false;
            op = ReplaceKeyword(op);
            var keyword = KeywordsEnumList.GetKeyword(op);

            if (keyword != KeywordsEnum.NOKEYWORD)
            {
                containKeyword = true;
                if ((keyword & KeywordsEnum.ON) == KeywordsEnum.ON) {
                    value = InputParser.VBat;
                    return BuildBetweenOrEqualComparison(value, limitInfo, gapFromValue, valueGap);
                } else if ((keyword & KeywordsEnum.OFF) == KeywordsEnum.OFF)
                {
                    value = channelOutputInfo.OffValue;
                    return BuildBetweenOrEqualComparison(value, limitInfo, gapFromValue, valueGap);
                }
                else if ((keyword & KeywordsEnum.EQUAL) == KeywordsEnum.EQUAL ||
                    (keyword & KeywordsEnum.SUP) == KeywordsEnum.SUP ||
                    (keyword & KeywordsEnum.INF) == KeywordsEnum.INF) {
                        return CreateComparisonFromString(limitInfo, op, gapFromValue, valueGap);
                }
            }

            if (!containKeyword)
            {
                CreateComparisonFromString(limitInfo, op);
            }


            throw new Exception("no key detected");
        }

        protected ComparisonContainer CreateComparisonFromString(LimitInfo limitInfo, string operation, bool gapFromValue = false, decimal valueGap = 0)
        {
            Comparison comparison = null;
            decimal value = 0;

            if (decimal.TryParse(operation, out value))
            {
                return BuildBetweenOrEqualComparison(value, limitInfo);
            }
            else
            {
                var valueString = GetValueString(operation);

                var regexpBinary = new Regex(@".*=b([-]?\d+)");
                var regexpHexa = new Regex(@".*=h([-]?[\dA-Fa-f]+)");

                Match matchBinary = regexpBinary.Match(operation);
                Match matchHexa = regexpHexa.Match(operation);

                Decimal.TryParse(valueString, out value);

                if (operation.Contains("!="))
                {
                    comparison = new NotEqualComparison();
                    ((NotEqualComparison)comparison).A = value;
                    return new ComparisonContainer(comparison, ComparisonKind.NOTEQUAL);
                } else if (/*operation.Contains("s=")*/matchHexa.Success)
                {
                    comparison = new EqualHexComparison();
                    ((EqualHexComparison)comparison).Text = valueString;
                    return new ComparisonContainer(comparison, ComparisonKind.EQUALTEXT);
                }
                else if (matchBinary.Success)
                {
                    value = Convert.ToInt64(valueString, 2);
                    comparison = new EqualComparison();
                    ((EqualComparison)comparison).A = value;
                    return new ComparisonContainer(comparison, ComparisonKind.EQUAL);
                }
                else if(operation.Contains("<"))
                {
                    comparison = new InfComparison();
                    ((InfComparison)comparison).A = value;
                    return new ComparisonContainer(comparison, ComparisonKind.INF);
                }
                else if (operation.Contains(">"))
                {
                    comparison = new SupComparison();
                    ((SupComparison)comparison).A = value;
                    return new ComparisonContainer(comparison, ComparisonKind.SUP);
                }
                else if (operation.Contains("=") || operation.Contains("＝"))
                {
                    return BuildBetweenOrEqualComparison(value, limitInfo, gapFromValue, valueGap);
                }
                throw new Exception("No possible comparision detected");
            }
        }

        private string GetValueString(string operation)
        {
            var regexpDecimal = new Regex(@".*[=＝><]([-]?\d+[.]?\d*)");

            Match matchDecimal = regexpDecimal.Match(operation);
            if (matchDecimal.Success)
            {
                return matchDecimal.Groups[1].Value;
            }


            var regexpBinary = new Regex(@".*[=＝]b([-]?\d+)");

            Match matchBinary = regexpBinary.Match(operation);
            if (matchBinary.Success)
            {
                return matchBinary.Groups[1].Value;
            }

            var regexpHexa = new Regex(@".*[=＝]h([-]?[\dA-Fa-f]+)");

            Match matchHexa = regexpHexa.Match(operation);
            if (matchHexa.Success)
            {
                return matchHexa.Groups[1].Value;
            }

            throw new ParserException(ParserExceptionType.BAD_NUMBER_FORMAT);
        }

        protected ComparisonContainer BuildBetweenOrEqualComparison(decimal value, LimitInfo limitInfo, bool gapFromValue = false, decimal valueGap=0)
        {
            var comparison = new BetweenOrEqualComparison();

            if (gapFromValue)
            {
                ((BetweenOrEqualComparison)comparison).A = value + (valueGap * limitInfo.Gap);
                ((BetweenOrEqualComparison)comparison).B = value - (valueGap * limitInfo.Gap);
            }
            else
            {
                ((BetweenOrEqualComparison)comparison).A = value + (InputParser.VBat * limitInfo.Gap);
                ((BetweenOrEqualComparison)comparison).B = value - (InputParser.VBat * limitInfo.Gap);
            }

            return new ComparisonContainer(comparison, ComparisonKind.BETWEENOREQUAL);
        }

        public void CreateReport(FlowChartState state, string projectName)
        {
            Directory.CreateDirectory("C:/ART logics/Report/" + projectName + "/");
            Directory.CreateDirectory("C:/ART logics/Template/" + projectName + "/");

            var templateFile = "template/reportTemplate.repx";
            var templateFilePath = "C:/ART logics/Template/" + projectName + "/reportTemplate.repx";

            System.IO.File.Copy(templateFile, templateFilePath, true);

            var reportAction = new ReportAction();
            reportAction.Description = "Excel Report";
            reportAction.ReportTemplate = templateFilePath;
            reportAction.OutputPath = "C:/ART logics/Report/" + projectName + "/";
            state.Actions.Add(reportAction);
        }

        protected string GetDelayTypeFromExcel(string operation, KeywordsEnum type)
        {
            var indexOfKeyword = operation.IndexOf(type.ToString().ToLower());
            testType = type;

            var firstNumber = operation.Substring(indexOfKeyword).IndexOfAny("0123456789".ToArray());
            var lastNumber = operation.Substring(indexOfKeyword).LastIndexOfAny("0123456789".ToArray()) + 1;
            var lenght = lastNumber - firstNumber;
            var value = operation.Substring(indexOfKeyword).Substring(firstNumber, lenght);

            specificTime = float.Parse(value);

            var myOperation = ReplaceKeyword(operation);
            specificTime = KeywordsEnumList.CalculateTIme(myOperation, specificTime);

            operation = operation.Remove(indexOfKeyword);

            return operation;
        }

        protected string BuildOperation(string operation, string result, out string newResult, bool removeDmFrame = true)
        {
            var myOperation = ReplaceKeyword(operation);
            newResult = result;

            if (myOperation.Contains(KeywordsEnum.AFTER.ToString().ToLower()))
            {
                testType = KeywordsEnum.AFTER;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains(KeywordsEnum.DURING.ToString().ToLower()))
            {
                testType = KeywordsEnum.DURING;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains(KeywordsEnum.HIGHLOW.ToString().ToLower()))
            {
                testType = KeywordsEnum.HIGHLOW;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains(KeywordsEnum.HIGH.ToString().ToLower()))
            {
                testType = KeywordsEnum.HIGH;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains(KeywordsEnum.LOW.ToString().ToLower()))
            {
                testType = KeywordsEnum.LOW;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains(KeywordsEnum.EVERY.ToString().ToLower()))
            {
                testType = KeywordsEnum.EVERY;
                var indexOfKeyword = myOperation.IndexOf(testType.ToString().ToLower());
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                operation = operation.Remove(indexOfKeyword);
            }
            else if (myOperation.Contains("singleframe"))
            {
                var indexOfKeyword = myOperation.IndexOf("singleframe");
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                if (removeDmFrame)
                {
                    operation = operation.Remove(indexOfKeyword);
                }
            }
            else if (myOperation.Contains("multipleframe"))
            {
                var indexOfKeyword = myOperation.IndexOf("multipleframe");
                newResult = result + " " + myOperation.Substring(indexOfKeyword);
                if (removeDmFrame)
                {
                    operation = operation.Remove(indexOfKeyword);
                }
            }

            return operation;
        }
    }
}
