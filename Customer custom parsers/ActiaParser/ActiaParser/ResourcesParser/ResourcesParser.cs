using ActiaParser.ActionParser;
using ActiaParser.Define;
using ArtLogics.TestSuite.Actions.Common;
using ArtLogics.TestSuite.Boards;
using ArtLogics.TestSuite.Boards.Resources;
using ArtLogics.TestSuite.Environment;
using ArtLogics.TestSuite.Limits;
using ArtLogics.TestSuite.Limits.Comparisons;
using ArtLogics.TestSuite.Limits.Comparisons.MultiRange;
using ArtLogics.TestSuite.Operations;
using ArtLogics.TestSuite.Services;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Shared.Services.Data;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.Parser.Utils;
using ArtLogics.TestSuite.TestResults;
using DevExpress.XtraEditors;
using NLog;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ActiaParser.ResourcesParser
{
    public class ResourcesParser
    {
        private ExcelWorkbook Workbook;
        private Product product;
        private static ILogger _log = LogManager.GetCurrentClassLogger();

        public ResourcesParser(ExcelWorkbook Workbook, Product product)
        {
            this.Workbook = Workbook;
            this.product = product;
        }

        public void ParseResources()
        {
            ParseOutput();
            ParseInput();
        }

        private void ParseOutput()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.OutputDefinitionSheet];

            //ParseResource(Worksheet);

            var columnNumberResource = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.ResourceColName, Worksheet);
            var columnNumberAlias = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.AliasColName, Worksheet);
            var columnNumberCoeffA = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.CoeffAColName, Worksheet);
            var columnNumberCoeffB = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.CoeffBColName, Worksheet);
            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.FunctionColName, Worksheet);
            var columnNumberOffValue = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.OffValueColName, Worksheet);
            var columnNumberLoad = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.LoadColName, Worksheet);
            var columnNumberType = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.TypeColName, Worksheet);
            var columnNumberProductId = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.ProductIdColName, Worksheet);
            var columnNumberOverCurrent = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.OverCurrentColName, Worksheet);
            //ParseResource(Worksheet);

            for (var i = ParserStaticVariable.StartRow; (string)Worksheet.Cells[i, columnNumberAlias].Value != null; i++)
            {
                int CurrentBoardNumber = 0;
                string resource = "";
                try
                {
                    var value = Worksheet.Cells[i, columnNumberResource].Value.ToString().Split(',');
                    CurrentBoardNumber = int.Parse(value[0]) - 1;
                    resource = value[1];

                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell + " " + Worksheet.Cells[i, columnNumberResource]);
                    continue;
                }
                var Alias = Worksheet.Cells[i, columnNumberAlias].Value.ToString();

                float CoeffA = 1;
                if (columnNumberCoeffA > 0)
                {
                    try
                    {
                        CoeffA = float.Parse(Worksheet.Cells[i, columnNumberCoeffA].Value.ToString());
                    }
                    catch { }
                }

                float CoeffB = 0;
                if (columnNumberCoeffB > 0)
                {
                    try
                    {
                        CoeffB = float.Parse(Worksheet.Cells[i, columnNumberCoeffB].Value.ToString());
                    }
                    catch { }
                }

                Channel channel;
                try
                {
                    var RessourceKind = (ChannelKind)Enum.Parse(typeof(ChannelKind), resource.Remove(resource.IndexOf("-")));
                    var RessourceNumber = int.Parse(resource.Remove(0, resource.IndexOf("-") + 1)) - 1;

                    var channels = product.Boards[CurrentBoardNumber].Channels.Concat(product.Boards[CurrentBoardNumber].Extensions.SelectMany(f => f.Channels));

                    channel = channels.Where(e => e.Kind.Equals(RessourceKind)).ElementAt(RessourceNumber);
                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtLine + " " + i + ", " + Resources.lang.ThisRessourcesIsNotAvailableInYourProject);
                    continue;
                }

                if (channel != null)
                {
                    try
                    {
                        var functionNameObject = Worksheet.Cells[i, columnNumberFunction].Value;
                        var offValueObject = Worksheet.Cells[i, columnNumberOffValue].Value;
                        var loadObject = Worksheet.Cells[i, columnNumberLoad].Value;
                        var typeObject = Worksheet.Cells[i, columnNumberType].Value;
                        var productIdObject = Worksheet.Cells[i, columnNumberProductId].Value;
                        var overCurrentObject = Worksheet.Cells[i, columnNumberOverCurrent].Value;

                        var functionName = (functionNameObject != null) ? functionNameObject.ToString().Trim().ToLower() : "";
                        var offValue = (offValueObject != null) ? decimal.Parse(offValueObject.ToString()) : 0;
                        var load = (loadObject != null) ? loadObject.ToString().Trim().ToLower() : "";
                        var type = (typeObject != null) ? typeObject.ToString().Trim().ToLower() : "";
                        var productId = (productIdObject != null) ? productIdObject.ToString().Trim().ToLower() : "";
                        var overCurrent = (overCurrentObject != null) ? float.Parse(overCurrentObject.ToString().Trim()) : 0;

                        if (!type.Contains("in")) {

                            var channelInfo = new ChannelOutputInfo()
                            {
                                Alias = "",
                                Channel = channel,
                                Function = functionName,
                                Resource = resource,
                                OffValue = offValue,
                                Board = product.Boards[CurrentBoardNumber],
                                Load = load,
                                Product = productId,
                                OverCurrent = overCurrent,
                            };

                            if (Worksheet.Cells[i, columnNumberFunction].Value.ToString() == "overload_C")
                            {
                                if (channel.Kind == ChannelKind.RELAY)
                                {
                                    int loadKey;
                                    int.TryParse(Worksheet.Cells[i, columnNumberLoad].Value.ToString(), out loadKey);
                                    OverloadParser.OverloadClusterRelays[loadKey] = channelInfo;
                                }
                                else
                                {
                                    OverloadParser.OverloadClusterVoltmeter = channelInfo;
                                }
                            }
                            else if (Worksheet.Cells[i, columnNumberFunction].Value.ToString().Contains("overload_B"))
                            {
                                var productName = Worksheet.Cells[i, columnNumberProductId].Value.ToString();
                                var productPos = int.Parse(productName.Replace("BCM ", ""));
                                if (channel.Kind == ChannelKind.RELAY)
                                {
                                    float loadKey;
                                    float.TryParse(Worksheet.Cells[i, columnNumberLoad].Value.ToString(), out loadKey);

                                    if (!OverloadParser.OverloadBCMRelay.Keys.Contains(productPos))
                                    {
                                        OverloadParser.OverloadBCMRelay[productPos] = new Dictionary<float, ChannelOutputInfo>();
                                    }

                                    OverloadParser.OverloadBCMRelay[productPos][loadKey] = channelInfo;
                                }
                                else
                                {
                                    OverloadParser.OverloadBCMVoltmeter[productPos] = channelInfo;
                                }
                            }

                            else if (columnNumberLoad > 0 && Worksheet.Cells[i, columnNumberLoad].Value != null &&
                                Worksheet.Cells[i, columnNumberLoad].Value.ToString() == "SELECT")
                            {
                                OverloadParser.OverloadFunctionRelays[Worksheet.Cells[i, columnNumberFunction].Value.ToString().ToLower()] = channelInfo;
                            }
                            else if (Worksheet.Cells[i, columnNumberFunction].Value != null && Worksheet.Cells[i, columnNumberFunction].Value.ToString() != "")
                            {
                                KeywordParser.FunctionChannelOutputMap.Add(channelInfo);
                            }

                        } else
                        {
                            if (load == "in")
                            {
                                var channelInfo = new ChannelInputInfo()
                                {
                                    Alias = "",
                                    Channel = channel,
                                    Function = functionName,
                                    HwValue = "",
                                    LogicalContact = type,
                                    Board = product.Boards[CurrentBoardNumber],
                                    Resource = resource,
                                    Pulse = 0,
                                    Ratio = 0,
                                };

                                if (!KeywordParser.ChannelInputInterpretations.ContainsKey(functionName) || KeywordParser.ChannelInputInterpretations[functionName] == null)
                                {
                                    KeywordParser.ChannelInputInterpretations[functionName] = new ChannelInputInterpretations();
                                }

                                KeywordParser.FunctionChannelInputMap.Add(channelInfo);
                            }
                        }
                    }
                    catch
                    {
                        _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtLine + " " + i + ", " + Resources.lang.ThisRessourcesIsNotAvailableInYourProject);
                        continue;
                    }

                    channel.Alias = Alias;
                    if (columnNumberCoeffA > 0)
                    {
                        channel.CoeffA = CoeffA;
                    }
                    if (columnNumberCoeffB > 0)
                    {
                        channel.CoeffB = CoeffB;
                    }
                }
            }
        }

        private void ParseInput()
        {
            var Worksheet = Workbook.Worksheets[ParserStaticVariable.InputsDefinitionSheet];

            var columnNumberResource = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.ResourceColName, Worksheet);
            var columnNumberAlias = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.AliasColName, Worksheet);
            var columnNumberCoeffA = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.CoeffAColName, Worksheet);
            var columnNumberCoeffB = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.CoeffBColName, Worksheet);
            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.FunctionColName, Worksheet);
            var columnNumberHwValue = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.HwValueColName, Worksheet);
            var columnNumberLogicalContact = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.LogicalContactColName, Worksheet);
            var columnNumberRealValue = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.RealValueColName, Worksheet);
            var columnNumberInterpretedValue = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.InterpretedValueColName, Worksheet);
            var columnNumberPulse = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.PulseColName, Worksheet);
            var columnNumberRatio = ExcelUtilsFunction.GetColumnNumber(ParserStaticVariable.TitleRow, ParserStaticVariable.RatioColName, Worksheet);

            //ParseResource(Worksheet);

            for (var i = ParserStaticVariable.StartRow; (string)Worksheet.Cells[i, columnNumberAlias].Value != null; i++)
            {
                int CurrentBoardNumber = 0;
                string resource = "";
                try
                {
                    var value = Worksheet.Cells[i, columnNumberResource].Value.ToString().Split(',');
                    CurrentBoardNumber = int.Parse(value[0]) - 1;
                    resource = value[1];

                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell + " " + Worksheet.Cells[i, columnNumberResource]);
                    continue;
                }
                var Alias = Worksheet.Cells[i, columnNumberAlias].Value.ToString();

                float CoeffA = 1;
                if (columnNumberCoeffA > 0)
                {
                    try
                    {
                        CoeffA = float.Parse(Worksheet.Cells[i, columnNumberCoeffA].Value.ToString());
                    }
                    catch { }
                }

                float CoeffB = 0;
                if (columnNumberCoeffB > 0)
                {
                    try
                    {
                        CoeffB = float.Parse(Worksheet.Cells[i, columnNumberCoeffB].Value.ToString());
                    }
                    catch { }
                }

                Channel channel;
                try
                {
                    var RessourceKind = (ChannelKind)Enum.Parse(typeof(ChannelKind), resource.Remove(resource.IndexOf("-")));
                    var RessourceNumber = int.Parse(resource.Remove(0, resource.IndexOf("-") + 1)) - 1;

                    var channels = product.Boards[CurrentBoardNumber].Channels.Concat(product.Boards[CurrentBoardNumber].Extensions.SelectMany(f => f.Channels));

                    channel = channels.Where(e => e.Kind.Equals(RessourceKind)).ElementAt(RessourceNumber);
                }
                catch
                {
                    _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtLine + " " + i + ", " + Resources.lang.ThisRessourcesIsNotAvailableInYourProject);
                    continue;
                }

                if (channel != null)
                {
                    try
                    {
                        if (Worksheet.Cells[i, columnNumberFunction].Value != null && Worksheet.Cells[i, columnNumberFunction].Value.ToString() != "")
                        {
                            var functionNameObject = Worksheet.Cells[i, columnNumberFunction].Value;
                            var valueNameObject = Worksheet.Cells[i, columnNumberHwValue].Value;
                            var logicalContactObject = Worksheet.Cells[i, columnNumberLogicalContact].Value;
                            var interpretedValueObject = Worksheet.Cells[i, columnNumberInterpretedValue].Value;
                            var realValueObject = Worksheet.Cells[i, columnNumberRealValue].Value;
                            var pulseObject = Worksheet.Cells[i, columnNumberPulse].Value;
                            var ratioObject = Worksheet.Cells[i, columnNumberRatio].Value;

                            var functionName = (functionNameObject != null)?Worksheet.Cells[i, columnNumberFunction].Value.ToString().ToLower().Trim():"";
                            var valueName = (valueNameObject != null)?Worksheet.Cells[i, columnNumberHwValue].Value.ToString().ToLower() : "";
                            var logicalContact = (logicalContactObject != null)?Worksheet.Cells[i, columnNumberLogicalContact].Value.ToString().ToLower() : "";
                            var interpretedValue = (interpretedValueObject != null) ? decimal.Parse(interpretedValueObject.ToString()) : -1;
                            var realValue = (realValueObject != null) ? decimal.Parse(realValueObject.ToString()) : -1;
                            var pulse = (pulseObject != null) ? decimal.Parse(pulseObject.ToString()) : 0;
                            var ratio = (ratioObject != null) ? decimal.Parse(ratioObject.ToString()) : 0;

                            if (functionName == "燃油传感器")
                            {
                                var t = "";
                            }

                            var channelInfo = new ChannelInputInfo()
                            {
                                Alias = "",
                                Channel = channel,
                                Function = functionName,
                                HwValue = valueName,
                                LogicalContact = logicalContact,
                                Board = product.Boards[CurrentBoardNumber],
                                Resource = resource,
                                Pulse = pulse,
                                Ratio = ratio,
                            };

                            var channelInputInterpretation = new ChannelInputInterpretation()
                            {
                                InterpretedValue = interpretedValue,
                                RealValue = realValue,
                            };

                            if (!KeywordParser.ChannelInputInterpretations.ContainsKey(functionName) || KeywordParser.ChannelInputInterpretations[functionName] == null)
                            {
                                KeywordParser.ChannelInputInterpretations[functionName] = new ChannelInputInterpretations();
                            }

                                KeywordParser.ChannelInputInterpretations[functionName].Add(channelInputInterpretation);

                            KeywordParser.FunctionChannelInputMap.Add(channelInfo);
                        }
                    }
                    catch
                    {
                        _log.Info(Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtLine + " " + i + ", " + Resources.lang.ThisRessourcesIsNotAvailableInYourProject);
                        continue;
                    }

                    channel.Alias = Alias;
                    if (columnNumberCoeffA > 0)
                    {
                        channel.CoeffA = CoeffA;
                    }
                    if (columnNumberCoeffB > 0)
                    {
                        channel.CoeffB = CoeffB;
                    }
                }
            }
        }
    }
}
