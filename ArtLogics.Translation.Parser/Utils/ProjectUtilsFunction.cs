using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Actions.Common;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Environment.GlobalVariables;
using ArtLogics.TestSuite.Environment.Variables;
using ArtLogics.TestSuite.GlobalVariables;
using ArtLogics.TestSuite.GlobalVariables._Impl;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.CaptureSensor;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.Parser.Utils
{
    public static class ProjectUtilsFunction
    {
        public static void BuildAwgOut(ChannelActionContainer channelActionContainer, AwgAction action, AwgFunction function,
                                        float frequency, float dutyCycle, float amplitude, float offset)
        {
            var AwgOperation = new AwgOperation();
            channelActionContainer.Operation = AwgOperation;
            AwgOperation.Action = action;
            AwgOperation.Function = function;
            AwgOperation.Frequency = frequency;
            AwgOperation.DutyCycle = dutyCycle;
            AwgOperation.Amplitude = amplitude;
            AwgOperation.Offset = offset;
        }

        public static void BuildDCVout(ChannelActionContainer channelActionContainer, float value)
        {
            var VoltageOperation = new VoltageOutOperation();
            channelActionContainer.Operation = VoltageOperation;
            VoltageOperation.Value = value;
        }

        public static void BuildRelayOperation(ChannelActionContainer channelActionContainer, RelayState value)
        {
            var RelayOperation = new RelayOperation();
            channelActionContainer.Operation = RelayOperation;

            RelayOperation.State = value;
        }

        public static void BuildDelay(FlowChartState state, string desc, long time)
        {
            if (time < 0) {
                time = 0;
            }
            var delayAction = new DelayAction();
            delayAction.Duration = new TimeSpan(time * 10000);
            delayAction.Description = desc;
            delayAction.ImagePath = null;
            state.Actions.Add(delayAction);
        }

        public static ChannelAction BuildChannelAction(FlowChartState state, string desc)
        {
            var channelAction = new ChannelAction();
            channelAction.Description = desc;
            state.Actions.Add(channelAction);
            channelAction.ImagePath = null;

            return channelAction;
        }

        public static CaptureSensorAction BuildCapture(string desc, Channel channel, CaptureType captureType,
                                                    bool startCapture, Board board)
        {
            var channelActionCapture = new CaptureSensorAction();
            channelActionCapture.Description = desc;
            channelActionCapture.Channel = channel;
            channelActionCapture.ActionKind = captureType;
            channelActionCapture.StartCapture = startCapture;
            channelActionCapture.CurrentBoard = board;
            channelActionCapture.ImagePath = null;

            return channelActionCapture;
        }

        public static void BuildUserAction(FlowChartState state, string desc, UserActionButtons buttons, string imagePath = null, string soundPath = null, bool showTimer = false)
        {
            var userAction = new UserInputAction();
            userAction.Description = desc;
            userAction.Buttons = buttons;
            if (imagePath != null) {
                userAction.ImagePath = imagePath;
            }

            if (soundPath != null)
            {
                userAction.SoundFile = soundPath;
            }

            userAction.ShowTimer = showTimer;

            state.Actions.Add(userAction);
        }

        public static void BuildFrqOut(ChannelActionContainer channelActionContainer, float value)
        {
            var FreqOutOperation = new FreqOutOperation();
            channelActionContainer.Operation = FreqOutOperation;
            FreqOutOperation.DutyCycle = 0.5f;
            FreqOutOperation.OffVoltage = 0;
            FreqOutOperation.OnVoltage = 6;
            FreqOutOperation.Frequency = value;
        }

        public static void BuildCanTxStd(ChannelActionContainer channelActionContainer, string value, string messageName, 
                                        Variable variable)
        {
            var CanOperation = new CanOperation();
            CanOperation.Action = CanAction.SENDMSGVALUE;
            CanOperation.MessageName = messageName;
            CanOperation.Variable = variable;
            CanOperation.Value = value;
            CanOperation.Data = value;

            channelActionContainer.Operation = CanOperation;
        }

        public static NamedGlobalVariable BuildGlobalVariable(Product product, string name/*, object @decimal*/)
        {
            var descriptor = new DecimalVar(new DecimalVar._Descriptor());

            var globalVariable = new NamedGlobalVariable();
            globalVariable.Name = name;
            globalVariable.DescriptorId = descriptor.Descriptor.Guid;

            

            product.GlobalVariables.Add(globalVariable);

            return globalVariable;
        }
    }
}
