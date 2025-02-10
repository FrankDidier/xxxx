using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.CanLoop;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public class StepperGaugesParser : ActionParserBase
    {
        private bool isFirstChannelCan = false;

        public StepperGaugesParser(ExcelWorkbook Workbook, Project project) : base(Workbook, project)
        {
            System.IO.Directory.CreateDirectory(ParserStaticVariable.GlobalPath);
        }

        public void ParseActionsStepperGauges()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.StepperGaugeSheet];

            var columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.StepperGaugesDefinitionColName, Worksheet);
            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.StepperGaugesFunctionColName, Worksheet);
            var columnNumberChannel = ExcelUtilsFunction.GetColumnNumber(3, ParserStaticVariable.StepperGaugesChannelName, Worksheet);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            
                for (var i = 4; Worksheet.Cells[i, columnNumberFunction].Value != null; i++)
                {
                try
                {
                    var channels = Worksheet.Cells[i, columnNumberChannel].Value.ToString().Trim().Split('\n');

                    var e = 0;
                    foreach (var channel in channels)
                    {
                        if (channel == "")
                            continue;

                        var name = Worksheet.Cells[i, columnNumberFunction].Value.ToString() + " => " + channel;
                        var defTab = Worksheet.Cells[i, columnNumberDefinition].Value.ToString().Split('\n');
                        var test = defTab[0];

                        test = test.Replace("指针显示:", "");

                        var testTab = test.Split(',');
                        float average = 0;
                        var excelPicture = Worksheet.Drawings["SG_" + i] as ExcelPicture;
                        
                        foreach (var currentTest in testTab)
                        {
                            canLoopSignalUsed.Clear();
                            var number = float.Parse(Regex.Match(currentTest, @"\d+(\.{0,1}\d+){0,1}").Value);
                            average += number;

                            if (channel.Substring(0, 2) == "0X" || channel.Substring(0, 2) == "0x")
                            {
                                CreateStepperGaugeActionCAN(state, excelPicture, name, number);
                            }
                            else
                            {
                                if (!IsCanChannel(channel))
                                {
                                    StopDefaultCanLoops(channels, state);
                                } else
                                {
                                    if (e == 0)
                                    {
                                        isFirstChannelCan = true;
                                    }
                                    foreach (var chn in channels)
                                    {
                                        if (!IsCanChannel(chn))
                                        {
                                            SetChannel(chn + "=oc", channel, state, Resources.lang.OpenResistiveChannels);
                                        }
                                    }
                                }
                                CreateStepperGaugeActionChannel(state, excelPicture, name, number, channel);
                                if (!IsCanChannel(channel))
                                {
                                    StartDefaultCanLoops(channels, state);
                                } else
                                {
                                    foreach (var chn in channels)
                                    {
                                        if (!IsCanChannel(chn))
                                        {
                                            SetChannel(chn + "=oc", channel, state, Resources.lang.OpenResistiveChannels, true);
                                        }
                                    }
                                }
                            }
                        }
                        canLoopSignalUsed.Clear();
                        average /= testTab.Count();

                        if (channel.Substring(0, 2) == "0X" || channel.Substring(0, 2) == "0x")
                        {
                            CreateStepperGaugeActionCAN(state, excelPicture, name, average);
                        }
                        else
                        {
                            if (!IsCanChannel(channel))
                            {
                                StopDefaultCanLoops(channels, state);
                            }
                            else
                            {
                                foreach (var chn in channels)
                                {
                                    if (!IsCanChannel(chn))
                                    {
                                        SetChannel(chn + "=oc", channel, state, Resources.lang.OpenResistiveChannels);
                                    }
                                }
                            }
                            CreateStepperGaugeActionChannel(state, excelPicture, name, average, channel);
                            if (!IsCanChannel(channel))
                            {
                                StartDefaultCanLoops(channels, state);
                            }
                            else
                            {
                                foreach (var chn in channels)
                                {
                                    if (!IsCanChannel(chn))
                                    {
                                        SetChannel(chn + "=oc", channel, state, Resources.lang.OpenResistiveChannels, true);
                                    }
                                }
                            }
                        }
                        e++;
                    }

                    if (channels.Count() > 1)
                    {
                        CreateStepperGaugeMultipleTest(Worksheet, columnNumberChannel, columnNumberFunction,
                            columnNumberDefinition, i, state, channels);
                    }
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
                }
            }
        }

        private void StartDefaultCanLoops(string[] signals, FlowChartState state)
        {
            foreach (var signalName in signals)
            {
                if (IsCanChannel(signalName))
                {
                    StartDefaultCanLoop(state, signalName.ToLower());
                }
            }
        }

        private void StopDefaultCanLoops(string[] signals, FlowChartState state)
        {
            foreach (var signalNameUpper in signals)
            {
                var signalName = signalNameUpper.ToLower();
                if (IsCanChannel(signalName))
                {
                    var limitInfo = MessageParser.MessageParser.LimitSignalInfoList.FirstOrDefault(l => l.SignalName != null && signalName.Contains(l.SignalName) && l.Direction == MessageFlowDirection.TX);
                    if (limitInfo != null)
                    {
                        var myChannelName = signalName.Replace(limitInfo.SignalName, "");
                        var canLoopActionStopDefault = new CanLoopAction();
                        canLoopActionStopDefault.ImagePath = null;
                        canLoopActionStopDefault.CanLoopCmd = CanLoopCmd.STOPLOOP;
                        canLoopActionStopDefault.Description = Resources.lang.StopDefaultCanLoop + " " + limitInfo.MessageName;
                        canLoopActionStopDefault.Variable = limitInfo.Variable;
                        SetCanLoop(canLoopActionStopDefault, signalName, 0);
                        state.Actions.Add(canLoopActionStopDefault);
                    }
                }
            }
        }

        private void CreateStepperGaugeMultipleTest(ExcelWorksheet Worksheet, int columnNumberChannel,
    int columnNumberFunction, int columnNumberDefinition,
    int i, FlowChartState state, string[] channels)
        {
            var defTab = Worksheet.Cells[i, columnNumberDefinition].Value.ToString().Split('\n');
            var test = defTab[0];
            test = test.Replace("指针显示:", "");
            var testTab = test.Split(',');
            float average = 0;
            var excelPicture = Worksheet.Drawings["SG_" + i] as ExcelPicture;

            var testNumber = 1;
            foreach (var currentTest in testTab)
            {
                canLoopSignalUsed.Clear();
                var number = float.Parse(Regex.Match(currentTest, @"\d+").Value);
                average += number;


                BuildStepperGaugeMultipleTest(Worksheet, columnNumberChannel, columnNumberFunction,
                        columnNumberDefinition, i, state, channels, excelPicture, number, testNumber, (testNumber == 1)? float.Parse(Regex.Match(testTab[1], @"\d+").Value):number/2);
                testNumber++;
            }
            canLoopSignalUsed.Clear();
            average /= testTab.Count();

            BuildStepperGaugeMultipleTest(Worksheet, columnNumberChannel, columnNumberFunction,
        columnNumberDefinition, i, state, channels, excelPicture, average, testNumber, average/2);
        }

        private void BuildStepperGaugeMultipleTest(ExcelWorksheet Worksheet, int columnNumberChannel,
            int columnNumberFunction, int columnNumberDefinition,
            int i, FlowChartState state, string[] channels,
            ExcelPicture excelPicture, float number, int testNumber, float wrongNumber)
        {
            string StepperGaugeName = "";
            List<string> analogNameList = new List<string>();

            var canValue = 0f;
            var resistiveValue = 0f;

            foreach (var channel in channels)
            {
                if (channel == "")
                    continue;

                var name = Worksheet.Cells[i, columnNumberFunction].Value.ToString() + " => " + channel;

                if (channel.Substring(0, 2) == "0X" || channel.Substring(0, 2) == "0x")
                {
                    //number
                    var channelActionStartCanLoop = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionStart + " " + name);
                    StepperGaugeName = name;
                }
                else
                {
                    //wrongNumber
                    if (IsCanChannel(channel) && isFirstChannelCan)
                    {
                        canValue = SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStart + " " + name);
                    } else if (IsCanChannel(channel) || isFirstChannelCan)
                    {
                        var tmpValue = 0f;
  
                        tmpValue = SetChannel(channel + "=" + wrongNumber, name, state, Resources.lang.ChannelActionStart + " " + name);

                        if (IsCanChannel(channel))
                        {
                            canValue = tmpValue;
                        } else
                        {
                            resistiveValue = tmpValue;
                        }

                    }  else
                    {
                        resistiveValue = SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStart + " " + name);
                    }

                    if (IsCanChannel(channel))
                    {
                        StepperGaugeName = name;
                    }
                    else
                    {
                        analogNameList.Add(name);
                    }
                }
            }

            //ask result
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".jpeg";

            string desc = "";

            if (isFirstChannelCan)
            {
                desc = Resources.lang.YouShouldSee + " : " + StepperGaugeName + " " + Resources.lang.At + " " + canValue;

                foreach (var analogName in analogNameList)
                {
                    desc += " (" + analogName + " at " + resistiveValue + ")";
                }
            }
            else
            {
                desc = Resources.lang.YouShouldSee + " : " + analogNameList[0] + " " + Resources.lang.At + " " + resistiveValue +
                " (" + StepperGaugeName + " at " + canValue + ")";

            }

            excelPicture.Image.Save(path);
            ProjectUtilsFunction.BuildUserAction(state, desc, UserActionButtons.Pass | UserActionButtons.Fail, path);

            foreach (var channel in channels)
            {
                if (channel == "")
                    continue;

                var name = Worksheet.Cells[i, columnNumberFunction].Value.ToString() + " => " + channel;
                if (channel.Substring(0, 2) == "0X" || channel.Substring(0, 2) == "0x")
                {
                    var channelActionStopCanLoop = ProjectUtilsFunction.BuildChannelAction(state, Resources.lang.ChannelActionStop + " " + name);
                }
                else
                {
                    if (IsCanChannel(channel) && isFirstChannelCan)
                    {
                        SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStop + " " + name, true);
                    }
                    else if (IsCanChannel(channel) || isFirstChannelCan)
                    {
                        if (number > 0)
                        {
                            if (testNumber > 1)
                            {
                                SetChannel(channel + "=" + number / 2, name, state, Resources.lang.ChannelActionStop + " " + name, true);
                            }
                            else
                            {
                                SetChannel(channel + "=" + number * 2, name, state, Resources.lang.ChannelActionStop + " " + name, true);
                            }
                        }
                        else
                        {
                            SetChannel(channel + "=100", name, state, Resources.lang.ChannelActionStop + " " + name, true);
                        }
                    }
                    else
                    {
                        SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStop + " " + name, true);
                    }

                    CleanCanLoops(state, " " + name);
                }
            }
        }

        private void CreateStepperGaugeActionCAN(FlowChartState state, ExcelPicture excelPicture, string name, float number)
        {
            //start send canloop
            var channelActionStart = new CanLoopAction();
            channelActionStart.CanLoopCmd = CanLoopCmd.LOOP;
            channelActionStart.Description = Resources.lang.ChannelActionStart + " " + name;
            state.Actions.Add(channelActionStart);

            SetCanLoop(channelActionStart, name, number);

            //ask result
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".jpeg";
            excelPicture.Image.Save(path);
            ProjectUtilsFunction.BuildUserAction(state, Resources.lang.YouShouldSee + " : " + name + " " + Resources.lang.At  + " " + number, UserActionButtons.Pass | UserActionButtons.Fail, path);

            //stop send canloop
            var channelActionStop = new CanLoopAction();
            channelActionStop.CanLoopCmd = CanLoopCmd.STOPLOOP;
            channelActionStop.Description = Resources.lang.ChannelActionStop + " " + name;
            state.Actions.Add(channelActionStop);

            SetCanLoop(channelActionStop, name, number);
        }

        private void CreateStepperGaugeActionChannel(FlowChartState state, ExcelPicture excelPicture, string name, float number, string channel)
        {
            //start send channel
            var valueSet = SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStart + " " + name);

            //ask result
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".jpeg";
            excelPicture.Image.Save(path);
            ProjectUtilsFunction.BuildUserAction(state, Resources.lang.YouShouldSee + " : " + name + " " + Resources.lang.At + " " + valueSet, UserActionButtons.Pass | UserActionButtons.Fail, path);

            //stop send channel
            SetChannel(channel + "=" + number, name, state, Resources.lang.ChannelActionStop + " " + name, true);

            CleanCanLoops(state, " " + name);
        }
    }
}
