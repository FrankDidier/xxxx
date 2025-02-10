using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.DevXlate.Units;
using ArtLogics.TestSuite.Environment;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Shared.FlowChart;
using ArtLogics.TestSuite.Shared.Services.Data;
using ArtLogics.TestSuite.Testing.Actions.Report;
using ArtLogics.TestSuite.Testing.Configuration;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Interfaces;
using ArtLogics.Translation.Parser.Model;
using NLog;
using NLog.Config;
using NLog.Targets;
using System.Drawing;
using System.IO;
using System.Text;

namespace ArtLogics.Translation.Parser
{
    public class BaseParser : MarshalByRefObject, IParser
    {
        protected Project project { get; set; }
        protected List<DeviceDescription> boardDescriptor;
        protected static ILogger _log = LogManager.GetCurrentClassLogger();
        protected List<FlowChartState> stateList = new List<FlowChartState>();
        protected List<Transition> transitionList = new List<Transition>();

        private string _logFileName = "${basedir}/Log/log-" + DateTime.Now.ToString("yyyy-MM-dd-HH-mm") + ".txt";
        public string LogFileName { get; set; }

        public BaseParser()
        {
            this.boardDescriptor = BoardService.LoadDescriptors();
        }

        public virtual void Parse(string input, string output, ProjectConfiguration projectConfig)
        {
            ConfigureLogger();
            project = new Project
            {
                Name = projectConfig.ProjectName,
                Id = Guid.NewGuid().ToString(),
                Products = new List<Product>()
            };

            var product = new Product
            {
                Name = projectConfig.ProductName,
                Guid = Guid.NewGuid().ToString(),
                Boards = new List<Board>()
            };

            BuildProduct(product, projectConfig);
            project.Products.Add(product);
        }

        protected void BuildUnitTest(ProjectConfiguration projectConfig, Product product)
        {
            for (var e = 1; IniFile.IniDataRaw["WorkSpace" + e].Count > 0; e++)
            {
                var workspace = new Workspace();
                workspace.Name = IniFile.IniDataRaw["WorkSpace" + e]["Name"];
                workspace.TestCaseDefinition = project.Products[0].TestCaseDefinitions[0];

                for (var i = 1; IniFile.IniDataRaw["WorkSpace" + e]["ART" + i] != null; i++)
                {
                    var ArtConfigName = IniFile.IniDataRaw["WorkSpace" + e]["ART" + i].ToString();
                    var ArtConfig = IniFile.IniDataRaw[ArtConfigName];

                    var workspaceUnit = new WorkspaceUnit();

                    workspaceUnit.Unit.Kind = (UnitType)Enum.Parse(typeof(UnitType), ArtConfig["Type"], true);
                    workspaceUnit.Unit.Name = ArtConfig["Name"];

                    for (var f = 1; ArtConfig["Ressource" + f] != null && ArtConfig["Ressource" + f] != ""; f++)
                    {
                        var workspaceUnitSlot = new WorkspaceUnitSlot();

                        workspaceUnitSlot.Board = product.Boards.Where(bo => bo.Name == ArtConfig["Ressource" + f]).FirstOrDefault();

                        workspaceUnit.Layout.Slots.Add(workspaceUnitSlot);
                    }

                    workspace.Units.Add(workspaceUnit);
                }
                project.Workspaces.Add(workspace);
            }
        }

        private void ConfigureLogger()
        {
            var config = new LoggingConfiguration();

            var fileTarget = new FileTarget("target")
            {
                FileName = _logFileName,
                Layout = "${message}",
                KeepFileOpen = false,
                ConcurrentWrites = true,
                Encoding = Encoding.UTF8,
            };

            config.AddTarget(fileTarget);

            config.AddRuleForOneLevel(LogLevel.Info, fileTarget); // only info to file

            LogManager.Configuration = config;

            var logEventInfo = new LogEventInfo { };
            LogFileName = fileTarget.FileName.Render(logEventInfo);
        }

        public void BuildProduct(Product product, ProjectConfiguration projectConfig)
        {
            var id = 0;
            foreach (var resource in projectConfig.Resources)
            {
                var newBoard = new Board
                {
                    Type = resource.RessourceType,
                    Name = resource.Alias,
                    Channels = new List<Channel>()
                };

                foreach (var ext in resource.Extensions)
                {
                    if (!string.IsNullOrEmpty(ext.ExtensionType))
                    {
                        var extensionBoard = new Board
                        {
                            Type = ext.ExtensionType,
                            Name = $"Extension-{id++}"
                        };
                        newBoard.Extensions.Add(extensionBoard);
                    }
                }

                product.Boards.Add(newBoard);
            }
        }

        public void ExportToXml(string outputFilePath)
        {
            if (project == null)
            {
                throw new InvalidOperationException("No project data to export.");
            }

            var xml = new XElement("Items",
                new XElement("Project",
                    new XAttribute("Name", project.Name),
                    new XAttribute("Id", project.Id),
                    new XElement("Products",
                        project.Products.Select(product =>
                            new XElement("Item",
                                new XAttribute("Name", product.Name),
                                new XAttribute("Guid", product.Guid),
                                new XElement("Boards",
                                    product.Boards.Select(board =>
                                        new XElement("Item",
                                            new XAttribute("Type", board.Type),
                                            new XAttribute("Name", board.Name),
                                            new XElement("Channels",
                                                board.Channels.Select(channel =>
                                                    new XElement("Item",
                                                        new XAttribute("Port", channel.Port),
                                                        new XAttribute("Kind", channel.Kind),
                                                        new XAttribute("Alias", channel.Alias),
                                                        new XAttribute("CoeffA", channel.CoeffA),
                                                        new XAttribute("CoeffB", channel.CoeffB)
                                                    )
                                                )
                                            ),
                                            new XElement("Extensions",
                                                board.Extensions?.Select(extension =>
                                                    new XElement("Item",
                                                        new XAttribute("Type", extension.Type),
                                                        new XAttribute("Name", extension.Name)
                                                    )
                                                )
                                            )
                                        )
                                    )
                                )
                            )
                        )
                    )
                )
            );

            xml.Save(outputFilePath);
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

        protected void HandleException(System.Exception err, string DefaultMsg)
        {
            if (err is ParserException)
            {
                var error = err as ParserException;
                error.Log();
            }
            else
            {
                _log.Info(DefaultMsg);
            }
        }

        protected void CreateState(string stateName)
        {
            var product = project.Products[0];

            //basic sate
            var state = new FlowChartState();
            state.IconUri = new Uri($"res://artlogics.common.ui.resources/ArtLogics.Common.Ui.Resources.Properties.Resources?condition_25x25");
            state.Name = stateName;
            stateList.Add(state);

            //basic state appearance
            var groupAppearance = new GroupResultAppearance()
            {
                Name = "<Default>",
                Color = Color.Black
            };

            //basic transition
            var transition = new Transition();
            transition.Name = "transition" + (stateList.Count - 1);
            transition.FromState = stateList[stateList.Count - 2];
            transition.ToState = state;
            transitionList.Add(transition);

            var stateAppearance = new StateAppearance();
            stateAppearance.State = state;
            stateAppearance.GroupAppearance = groupAppearance;
            stateAppearance.X = 800 * (stateList.Count - 1);
            stateAppearance.Y = 200;
            stateAppearance.Width = 200;
            stateAppearance.Height = 100;
            stateAppearance.TextSize = 25;
            stateAppearance.BackgroundColor = Color.FromArgb(125, 28, 30, 28);
            stateAppearance.TextColor = Color.White;
            product.StateAppearances.Add(stateAppearance);
        }

        public void Dispose()
        {
            this.project = null;
            this.boardDescriptor = null;
            LogManager.Shutdown();
        }
    }

    // Models for project representation
    public class Project
    {
        public string Name { get; set; }
        public string Id { get; set; }
        public List<Product> Products { get; set; } = new List<Product>();
    }

    public class Product
    {
        public string Name { get; set; }
        public string Guid { get; set; }
        public List<Board> Boards { get; set; } = new List<Board>();
    }

    public class Board
    {
        public string Type { get; set; }
        public string Name { get; set; }
        public List<Channel> Channels { get; set; } = new List<Channel>();
        public List<Board> Extensions { get; set; } = new List<Board>();
    }

    public class Channel
    {
        public int Port { get; set; }
        public string Kind { get; set; }
        public string Alias { get; set; }
        public double CoeffA { get; set; }
        public double CoeffB { get; set; }
    }

    public class ProjectConfiguration
    {
        public string ProjectName { get; set; }
        public string ProductName { get; set; }
        public List<Resource> Resources { get; set; } = new List<Resource>();
    }

    public class Resource
    {
        public string RessourceType { get; set; }
        public string Alias { get; set; }
        public List<Extension> Extensions { get; set; } = new List<Extension>();
    }

    public class Extension
    {
        public string ExtensionType { get; set; }
    }
}
