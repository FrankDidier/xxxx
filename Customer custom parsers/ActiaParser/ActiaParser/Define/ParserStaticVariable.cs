using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.Define
{
    public static class ParserStaticVariable
    {
        public static string InputsDefinitionSheet = "INPUTS_DEFINITION";
        public static string OutputDefinitionSheet = "OUTPUT_DEFINITION";
        public static string ReceiverSheet = "MultIC S J1939 Receiver";
        public static string SenderSheet = "MultIC S J1939 Sender";
        public static string OutputsSheet = "OutPuts";
        public static string LedSheet = "MultIC S_LED";
        public static string KeywordsFunctionSheet = "FUNCTION MAP";
        public static string KeywordsSheet = "KEYWORDS";
        public static string OutputCurrentDeratingSheet = "Output current derating ";
        public static string SymbolSheet = "MultIC S_TFT_Symbols";
        public static string StatusSheet = "Status";
        public static string AlarmsSheet = "Alarm";
        public static string StepperGaugeSheet = "Stepper Gauges";
        public static string StartStepsSheet = "START_STEPS";
        public static string EndStepsSheet = "END_STEPS";
        public static string AdditionnalTestSheet = "ADDITIONAL_TESTS";

        //DEFINITION COLUMN
        public static string PinMapColName = "LOADBOX PIN";
        public static string AliasColName = "ALIAS";
        public static string AliasLabelColName = "LABEL";
        public static string ResourceColName = "Resource";
        public static string CoeffAColName = "A";
        public static string CoeffBColName = "B";
        public static string FunctionColName = "Function";
        public static string HwValueColName = "HW Value";
        public static string ChildColName = "Child";
        public static string LogicalContactColName = "Logical contact";
        public static string LoadColName = "Load";
        public static string TypeColName = "Type";
        public static string ProductIdColName = "PRODUCT ID";
        public static string OffValueColName = "OFF Value";
        public static string RealValueColName = "Real Value";
        public static string InterpretedValueColName = "Interpreted Value";
        public static string PulseColName = "Pulse";
        public static string RatioColName = "Ratio";
        public static string OverCurrentColName = "Over Current";

        //CAN COLUMN
        //sender
        public static string CanSenderUsedColName = "有效";
        public static string CanSenderIdColName = "Identifier\n标识符";
        public static string CanSenderObjectColName = "Objects\n项目";
        public static string CanSenderBeforeStartColName = "Definition\n定义";
        public static string CanSenderAfterFinishColName = "Comments\n备注";
        public static string CanSenderLimitFrequencyColName = "Frequency\n周期";
        public static string CanSenderTestColName = "Test conditions";
        public static string CheckSumAlgoritmColName = "Checksum Algoritm";

        //receiver
        public static string CanReceiverUsedColName = "有 效";
        public static string CanReceiverIdColName = "Identifier";
        public static string CanReceiverObjectColName = "Objects";
        public static string CanReceiverBeforeStartColName = "Definition";
        public static string CanReceiverAfterFinishColName = "Comments";
        public static string CanReceiverLimitFrequencyColName = "Frequency";
        public static string CanReceiverPriorityColName = "PRIORITY";

        //Can Common
        public static string CanLimitStartBitColName = "START BIT (POSITION)";
        public static string CanLimitBitLengthColName = "BIT LENGTH (SIZE)";
        public static string CanLimitEndianColName = "ENDIAN";
        public static string CanLimitAColName = "A";
        public static string CanLimitBColName = "B";
        public static string CanEndianColName = "ENDIAN";
        public static string CanFormulaLogicColName = "Formula Logic";

        //Outputs
        public static string OutputsDefinitionColName = "定义";
        public static string OutputsUsedColName = "有效";
        public static string OutputsFunctionColName = "描述";
        public static string OutputsNoneValueColName = "NONE State";

        //Led
        public static string LedsDefinitionColName = "定义";
        public static string LedsProductColName = "有效";
        public static string LedsFunctionColName = "名称";

        //Symbols
        public static string SymbolsDefinitionColName = "定义";
        public static string SymbolsFunctionColName = "名称";
        public static string SymbolsProductColName = "有效";
        public static string SymbolsColorColName = "颜色";
        public static string SymbolsRemarksColName = "备注";
        public static string SymbolsPictureColName = "TFT\n符号";


        //Keywords
        public static string KeywordsColName = "KEYWORDS";
        public static string DescriptionColName = "Description";

        //InputValue
        public static string VbatCell = "G2";

        //Status
        public static string StatusFunctionColName = "名称";
        public static string StatusDefinitionColName = "定义";

        //Alarm
        public static string AlarmDefinitionColName = "定义";
        public static string AlarmLangageColName = "Language";

        //Stepper Gauge
        public static string StepperGaugesDefinitionColName = "定义";
        public static string StepperGaugesFunctionColName = "名称";
        public static string StepperGaugesChannelName = "来源";

        //Additional test
        public static string AdditionnalTestResultColName = "Result";
        public static string AdditionnalTestConditionColName = "Condition";

        public static string StepsDefinitionColName = "Resource name";
        public static string StepsValueColName = "Default value";

        public static int TitleRow = 2;
        public static int StartRow = 3;

        public static string GlobalPath = "C:/ART logics/ProjectResources/Project-" + DateTime.Now.ToString("yyyy-MM-dd-hh-mm-ss") + "/";
    }
}
