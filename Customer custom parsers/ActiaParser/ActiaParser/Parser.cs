using ActiaParser.ActionParser;
using ActiaParser.Define;
using ActiaParser.MessageParser;
using ActiaParser.ResourcesParser;
using ArtLogics.TestSuite.Missions;
using ArtLogics.TestSuite.Serialization.Formatters.Xml;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Shared.FlowChart;
using ArtLogics.TestSuite.Testing;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.TestSuite.TestResults;
using ArtLogics.Translation.Parser;
using ArtLogics.Translation.Parser.Model;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;

namespace CustomParser
{
    [Serializable]
    public class Parser : BaseParser
    {
        private ExcelPackage excelPackage;

        private ExcelWorkbook Workbook;

        private ResourcesParser ResourcesParser;
        private XmlFormatter XmlFormatter;
        private MessageParser MessageParser;
        private ActionParser ActionParser;
        private AlarmParser AlarmParser;
        private LedParser LedParser;
        private StepperGaugesParser StepperGaugesParser;
        private SymbolsParser SymbolsParser;
        private StateParser StateParser;
        private AdditionnalTestParser AdditionnalTestParser;

        public Parser () : base()
        {
            
        }

        public override void Parse(string input, string output, ProjectConfiguration projectConfig)
        {
            Console.WriteLine(Thread.CurrentThread.CurrentUICulture);
            Console.WriteLine(ActiaParser.Resources.lang.DelayAfterPrerquireNone);

            KeywordParser.Reset();
            OverloadParser.Reset();

            base.Parse(input, output, projectConfig);
            System.Console.WriteLine("Hello Actia");

            FileInfo inputInfo = new FileInfo(input);
            excelPackage = new ExcelPackage(inputInfo);

            this.Workbook = excelPackage.Workbook;
            var product = project.Products[0];
            var projectName = output.Split('/').ToList().LastOrDefault();

            //basic mission profil
            /*foreach (var board in product.Boards)
            {
                var mission = new Mission(board.Name + "Mission", 100, 10, Color.Blue);
                var missionBoard = new MissionBoard();
                missionBoard.Mission = mission;
                missionBoard.Board = board;
                product.MissionProfiles.Add(missionBoard);
            }*/
            //basic flowchart

            //baseic begin state
            var stateBegin = new FlowChartEntryState();
            stateBegin.EntryCondition = TestResult.None;
            stateBegin.IconUri = new Uri($"res://artlogics.common.ui.resources/ArtLogics.Common.Ui.Resources.Properties.Resources?enter");
            stateBegin.Name = "Start state";//LocalRes.EntryState;

            stateList.Add(stateBegin);

            //basic state appearance
            var groupAppearance = new GroupResultAppearance()
            {
                Name = "<Default>",
                Color = Color.Black
            };

            System.IO.Directory.CreateDirectory(ParserStaticVariable.GlobalPath);

            var stateBeginAppearance = new StateAppearance();
            stateBeginAppearance.State = stateBegin;
            stateBeginAppearance.GroupAppearance = groupAppearance;
            stateBeginAppearance.X = 50;
            stateBeginAppearance.Y = 200;
            stateBeginAppearance.Width = 200;
            stateBeginAppearance.Height = 100;
            stateBeginAppearance.TextSize = 25;
            stateBeginAppearance.BackgroundColor = Color.FromArgb(125, 28, 30, 28);
            stateBeginAppearance.TextColor = Color.White;
            product.StateAppearances.Add(stateBeginAppearance);

            product.GroupAppearances.Add(groupAppearance);

            //test case definition
            TestCaseDefinition testdef = new TestCaseDefinition();
            testdef.Name = "toto";

            TestCaseGraph testGraph = new TestCaseGraph();
            testGraph.States = stateList;
            testGraph.Transitions = transitionList;
            testGraph.Definition = testdef;

            product.TestCaseDefinitions.Add(testdef);
            product.TestCaseGraphs.Add(testGraph);

            //basic action
            /*var action = new PowerSupplyAction();
            action.Description = "test action xml";
            state.Actions.Add(action);*/

            //Build workspace
            BuildUnitTest(projectConfig, product);

            InputParser.ParseVBat(this.Workbook);
            KeywordParser.ParseFunctionMap(this.Workbook);
            KeywordParser.ParseKeywords(this.Workbook);
            KeywordParser.ParseStatus(this.Workbook);

            ResourcesParser = new ResourcesParser(this.Workbook, product);
            ResourcesParser.ParseResources();

            MessageParser = new MessageParser(this.Workbook, project, projectConfig.Resources);
            MessageParser.ParseMessages();

            StateParser = new StateParser(this.Workbook, project);

            LedParser = new LedParser(this.Workbook, project);
            SymbolsParser = new SymbolsParser(this.Workbook, project);
            AlarmParser = new AlarmParser(this.Workbook, project);
            StepperGaugesParser = new StepperGaugesParser(this.Workbook, project);
            ActionParser = new ActionParser(this.Workbook, project);
            AdditionnalTestParser = new AdditionnalTestParser(this.Workbook, project, stateList, transitionList);


            StateParser.ParseStartState();
            CreateState("Led test");
            LedParser.ParseActionsLed();
            CreateState("Symbols test");
            SymbolsParser.ParseActionsSymbols();
            CreateState("Alarm test");
            AlarmParser.ParseActionsAlarm();
            CreateState("Stepper Gauges test");
            StepperGaugesParser.ParseActionsStepperGauges();
            CreateState("Message Sender test");
            MessageParser.BuildMessageSenderTest();
            CreateState("Output test");
            ActionParser.ParseActionsOutput(ActionType.NORMAL);
            CreateState("Buzzer and Video test");
            ActionParser.ParseActionsOutput(ActionType.BUZZERANDVIDEO);
            CreateState("Overload Rating test");
            ActionParser.ParseActionsOutput(ActionType.OVERLOADRATING);
            CreateState("Overload Max test");
            ActionParser.ParseOverloadOutput();
            AdditionnalTestParser.ParseAdditionnalTest();
            CreateState("End test");
            StateParser.ParseEndState(projectName);


            Directory.CreateDirectory(output.Substring(0, output.LastIndexOf("/")));
            //save
            XmlFormatter = new XmlFormatter();
            XmlFormatter.Serialize(new Project[] { project }, output);

            Workbook.Worksheets.Dispose();
            Workbook.Dispose();
        }
    }
}
