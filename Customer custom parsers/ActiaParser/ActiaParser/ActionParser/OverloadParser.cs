using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Actions.Common;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.Limits.Comparisons.MultiRange;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.TestSuite.TestResults;
using ArtLogics.Translation.Parser.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ActionParser
{
    public struct OverLoadInfo
    {
        public float Load { get; set; }
        public string Function { get; set; }
        public string Type { get; set; }
        public string Product { get; set; }
    }

    public static class OverloadParser
    {
        public static Dictionary<float, ChannelOutputInfo> OverloadClusterRelays = new Dictionary<float, ChannelOutputInfo>();
        public static ChannelOutputInfo OverloadClusterVoltmeter;
        public static Dictionary<int, Dictionary<float, ChannelOutputInfo>> OverloadBCMRelay = new Dictionary<int, Dictionary<float, ChannelOutputInfo>>();
        public static Dictionary<int, ChannelOutputInfo> OverloadBCMVoltmeter = new Dictionary<int, ChannelOutputInfo>();
        public static Dictionary<string, ChannelOutputInfo> OverloadFunctionRelays = new Dictionary<string, ChannelOutputInfo>();

        public static void Reset()
        {
            OverloadClusterRelays.Clear();
            OverloadClusterVoltmeter = null;
            OverloadBCMRelay.Clear();
            OverloadBCMVoltmeter.Clear();
            OverloadFunctionRelays.Clear();
        }

        public static void BuildDelayAction(FlowChartState state, string desc)
        {
            var delay = new DelayAction();
            delay.Description = desc;
            delay.Duration = new TimeSpan(20000000);
            state.Actions.Add(delay);
        }

        public static ChannelActionContainer BuildGetVoltAction(ChannelOutputInfo ChannelInfo, Product product)
        {
            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.IsUsed = true;
            
            if (ChannelInfo.Product == "multic s")
            {
                channelActionContainer.Channel = OverloadClusterVoltmeter.Channel;
                channelActionContainer.MainBoard = OverloadClusterVoltmeter.Board;
            } else
            {
                var productName = ChannelInfo.Product;
                var productPos = int.Parse(productName.Replace("bcm ", ""));
                channelActionContainer.Channel = OverloadBCMVoltmeter[productPos].Channel;
                channelActionContainer.MainBoard = OverloadBCMVoltmeter[productPos].Board;
            }

            var voltmeteraction = new VoltmeterMeasureOperation();
            channelActionContainer.Operation = voltmeteraction;

            var gapString = IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERVALUE"];
            var ErrorMessage = IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERMESSAGE"];
            var Severity = (Severity)Enum.Parse(typeof(Severity), IniFile.IniDataRaw["Channels"]["LIMITVOLTMETERTYPE"], true);

            bool percent = false;
            if (gapString.Contains("%"))
            {
                gapString = gapString.Replace("%", "");
                percent = true;
            }

            var gap = decimal.Parse(gapString);

            if (percent)
            {
                gap = gap / 100;
            }

            var limit = new Limit();

            var comparison = new BetweenOrEqualComparison();

            ((BetweenOrEqualComparison)comparison).A = InputParser.VBat + (InputParser.VBat * gap);
            ((BetweenOrEqualComparison)comparison).B = InputParser.VBat - (InputParser.VBat * gap);

            var comparisonContainer = new ComparisonContainer(comparison, ComparisonKind.BETWEENOREQUAL);

            limit.Container = comparisonContainer;
            limit.ErrorMessage = new ArtLogics.TestSuite.TestResults.ErrorMessage();
            limit.ErrorMessage.Severity = Severity;
            limit.ErrorMessage.Name = ErrorMessage;
            limit.Name = ErrorMessage;
            product.ErrorMessages.Add(limit.ErrorMessage);
            voltmeteraction.Limits.Add(limit);

            return channelActionContainer;
        }

        public static ChannelActionContainer BuildRelayAction(ChannelOutputInfo channelInputInfo, RelayState state, Product product)
        {
            //container by action
            var channelActionContainer = new ChannelActionContainer();
            channelActionContainer.IsUsed = true;

            //foreach the board to check if they contain the channel
            channelActionContainer.MainBoard = channelInputInfo.Board;

            channelActionContainer.Channel = channelInputInfo.Channel;

            var RelayOperation = new RelayOperation();
            channelActionContainer.Operation = RelayOperation;

            RelayOperation.State = state;

            return channelActionContainer;
        }

        public static ChannelOutputInfo GetLoadChannel(ChannelOutputInfo channelOutputInfo, float specificLoad = -1)
        {
            float currentDiff = 0;
            var firstCheck = true;
            ChannelOutputInfo channel = null;
            float loadLooked = float.Parse(channelOutputInfo.Load);

            if (specificLoad > -1)
            {
                loadLooked = specificLoad;
            }

            if (channelOutputInfo.Product == "multic s")
            {
                foreach (var load in OverloadClusterRelays.Keys)
                {
                    if (Math.Abs(load - loadLooked) < currentDiff || firstCheck)
                    {
                        currentDiff = Math.Abs(load - loadLooked);
                        channel = OverloadClusterRelays[load];
                        firstCheck = false;
                    }
                }
            }
            else
            {
                var productName = channelOutputInfo.Product;
                var productPos = int.Parse(productName.Replace("bcm ", ""));
                foreach (var load in OverloadBCMRelay[productPos].Keys)
                {
                    if (Math.Abs(load - loadLooked) < currentDiff || firstCheck)
                    {
                        currentDiff = Math.Abs(load - loadLooked);
                        channel = OverloadBCMRelay[productPos][load];
                        firstCheck = false;
                    }
                }
            }

            return channel;
        }
    }
}
