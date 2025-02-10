using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
using ArtLogics.TestSuite.Testing.StateMachines;
using ArtLogics.Translation.Parser.Exception;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.Parser.Utils;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ActiaParser.ActionParser
{
    public class SymbolsParser : ActionParserBase
    {
        private ExcelWorksheet Worksheet { get; set; }

        public SymbolsParser(ExcelWorkbook Workbook, Project project) : base(Workbook, project)
        {
            System.IO.Directory.CreateDirectory(ParserStaticVariable.GlobalPath);
        }

        public void ParseActionsSymbols()
        {
            var titleLineIndex = 4;
            Worksheet = Workbook.Worksheets[ParserStaticVariable.SymbolSheet];

            var columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsDefinitionColName, Worksheet);
            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsFunctionColName, Worksheet);
            var colNumberProduct = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsProductColName, Worksheet);
            var colNumberColor = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsColorColName, Worksheet);
            var colNumberRemark = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsRemarksColName, Worksheet);
            var colNumberPicture = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.SymbolsPictureColName, Worksheet);
            
            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            for (var i = 5; Worksheet.Cells[i, colNumberProduct].Value != null; i++)
            {
                try
                {
                    if (Worksheet.Cells[i, columnNumberDefinition].Value == null || Worksheet.Cells[i, colNumberProduct].Value.ToString() != "Y")
                        continue;

                    var testParameters = new TestParameters(Worksheet, i, columnNumberFunction, columnNumberDefinition, state, colNumberColor);

                    if ((testParameters.ExcelPicture != null) || (Worksheet.Cells[i, colNumberPicture].Value != null && Worksheet.Cells[i, colNumberPicture].Value.ToString().Trim() != ""))
                    {
                        BuildBasicSymbols(testParameters, Worksheet.Cells[i, columnNumberDefinition].ToString(), i);
                    } else
                    {
                        BuildGaugesSymbols(testParameters, state, Worksheet.Cells[i, colNumberRemark]);
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
                    continue;
                }
            }
        }

        private void BuildGaugesSymbols(TestParameters testParameters, FlowChartState state, ExcelRange remarkCell)
        {
            if(remarkCell.Value == null)
            {
                var exp = new ParserException(ParserExceptionType.NULL_REMARK_FOR_SYMBOL_GAUGE);
                exp.Cell = remarkCell.ToString();
                throw exp;
            }

            var remarks = remarkCell.Value.ToString().Split('\n');

            foreach (var remarkWithComment in remarks)
            {
                if (remarkWithComment == "" || remarkWithComment.Substring(0, 2) == "//")
                    continue;

                var tableWithComment = remarkWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                var remarkList = tableWithComment[0].ToLower().Trim().Split('且'); ;

                for (var i = 0; i < remarkList.Count(); i++)
                {
                    var remark = remarkList[i];
                    if (i < remarkList.Count() - 1)
                    {
                        try
                        {
                            SetChannel(remark, testParameters.Function, state, Resources.lang.SetChannel + " " + remark);
                        }
                        catch (ParserException ex)
                        {
                            ex.Cell = remarkCell.ToString();
                            ex.Subcondition = remark;
                            ex.SubconditionLine = 1;
                            ex.Mode = ParserExceptionMode.SETTER;
                            ex.Sheet = Worksheet.Name;
                            throw ex;
                        }
                    }
                    else
                    {
                        Regex symbolGaugeOdoRegexp = new Regex(@"odo\((.*)\)");
                        Match matchOdo = symbolGaugeOdoRegexp.Match(remark);

                        Regex symbolGaugeTripOdoRegexp = new Regex(@"tripodo\((.*)\)");
                        Match matchTripOdo = symbolGaugeTripOdoRegexp.Match(remark);

                        Regex symbolGaugeTimeRegexp = new Regex(@"time\((.*)\)");
                        Match matchTime = symbolGaugeTimeRegexp.Match(remark);

                        if (matchOdo.Success || matchTripOdo.Success || matchTime.Success)
                        {
                            BuildSpecialSymboleGaugeTest(state, testParameters.Conditions, remark, testParameters.Function);
                        }
                        else
                        {
                            Regex symbolGaugeRegexp = new Regex(@".*: ([-]?\d+[.]?\d*)[ ]*~[ ]*([-]?\d+[.]?\d*).*");

                            Match match = symbolGaugeRegexp.Match(testParameters.Conditions[0].Condition);
                            if (match.Success)
                            {
                                BuilsNormalSymbolGaugeTest(remark, testParameters.Function, testParameters.Color, testParameters.Conditions,
                                                        state, remarkCell, match);
                            }
                            else
                            {
                                throw new ParserException(ParserExceptionType.LOGIC_NOT_FOUND);
                            }
                        }
                    }
                }

                CleanCanLoops(state);
                for (var i = 0; i < remarkList.Count(); i++)
                {
                    var remark = remarkList[i];
                    if (i < remarkList.Count() - 1)
                    {
                        try
                        {
                            SetChannel(remark, testParameters.Function, state, Resources.lang.SetChannel + " " + remark, true, true);
                        }
                        catch (ParserException ex)
                        {
                            ex.Cell = remarkCell.ToString();
                            ex.Subcondition = remark;
                            ex.SubconditionLine = 1;
                            ex.Mode = ParserExceptionMode.SETTER;
                            ex.Sheet = Worksheet.Name;
                            throw ex;
                        }
                    }
                }
            }
        }

        private void BuildSpecialSymboleGaugeTest(FlowChartState state, List<ConditionsParser> conditions, string remark, string function)
        {
            Regex symbolGaugeOdoRegexp = new Regex(@"odo\((.*)\)");
            Match matchOdo = symbolGaugeOdoRegexp.Match(remark);

            Regex symbolGaugeTimeRegexp = new Regex(@"time\((.*)\)");
            Match matchTime = symbolGaugeTimeRegexp.Match(remark);

            if (matchOdo.Success)
            {
                BuildOdo(state, matchOdo.Groups[1].Value.ToString(), conditions, function);
            }
            else if (matchTime.Success) {
                BuildTime(state, conditions, matchTime.Groups[1].Value);
            }
            else {
                throw new ParserException(ParserExceptionType.LOGIC_NOT_FOUND);
            }
        }

        private void BuildTime(FlowChartState state, List<ConditionsParser> conditions, string sentence)
        {
            if (sentence != null && sentence != "")
            {
                ProjectUtilsFunction.BuildUserAction(state, Resources.lang.PleaseCheckTheClusterTimeAndCompareItWith +" " + sentence + " " + Resources.lang.With + " " +
                                             conditions[1].Condition, UserActionButtons.Pass | UserActionButtons.Fail);
            }
            else
            {
                ProjectUtilsFunction.BuildUserAction(state, Resources.lang.PleaseCheckTheClusterTimeAndCompareItWithYourWatch + " " + Resources.lang.With + " " +
                                             conditions[1].Condition, UserActionButtons.Pass | UserActionButtons.Fail);
            }
        }

        private void BuildOdo(FlowChartState state, string channelName, List<ConditionsParser> conditions, string function)
        {
            var odometerDelay = long.Parse(IniFile.IniDataRaw["Odometer"]["Delay"]);

            canLoopSignalUsed.Clear();
            SetChannel(channelName + "=0", function, state, function + Resources.lang.ChannelAction);
            ProjectUtilsFunction.BuildUserAction(state, Resources.lang.PleaseCheckTheCluster + " " + function + "counter", UserActionButtons.Pass);
            SetChannel(channelName + "=0", function, state, function + Resources.lang.ChannelAction, true);
            CleanCanLoops(state);
            canLoopSignalUsed.Clear();
            SetChannel(channelName + "=120", function, state, function + Resources.lang.ChannelAction);
            ProjectUtilsFunction.BuildDelay(state, Resources.lang.DelayForOdometer, odometerDelay);
            SetChannel(channelName + "=120", function, state, function + Resources.lang.ChannelAction, true);
            CleanCanLoops(state);
            canLoopSignalUsed.Clear();
            SetChannel(channelName + "=0", function, state, function + Resources.lang.ChannelAction);
            ProjectUtilsFunction.BuildUserAction(state, Resources.lang.PleaseCheckTheCluster + " " + function + " " + Resources.lang.CounterItShouldHaveIncreaseBy1Km, UserActionButtons.Pass | UserActionButtons.Fail);
            SetChannel(channelName + "=120", function, state, function + Resources.lang.ChannelAction, true);
            CleanCanLoops(state);
        }

        private void BuilsNormalSymbolGaugeTest(string remark, string function, string color, List<ConditionsParser> conditions,
                                                FlowChartState state, ExcelRange remarkCell, Match match)
        {
            var minValue = float.Parse(match.Groups[1].Value);
            var maxValue = float.Parse(match.Groups[2].Value);
            var incr = (maxValue - minValue) / 4;

            for (var currentValue = minValue; currentValue <= maxValue;)
            {
                canLoopSignalUsed.Clear();
                float value;
                try
                {
                    value = SetChannel(remark + "=" + currentValue, function, state, Resources.lang.ChannelActionSetter + " " + function + "=" + currentValue + " /" + remark);
                }
                catch (ParserException ex)
                {
                    ex.Cell = remarkCell.ToString();
                    ex.Subcondition = remark + "=" + currentValue;
                    ex.SubconditionLine = 1;
                    ex.Mode = ParserExceptionMode.SETTER;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }

                ProjectUtilsFunction.BuildUserAction(state, Resources.lang.ResultOf + " " + function + ", " + Resources.lang.YouShouldSee  + " " + function + " " + Resources.lang.With + " " + color +
                                         " light at value " + ((value < 0) ? currentValue : value) + " " + Resources.lang.With + " " +
                                         conditions[1].Condition + " /" + remark, UserActionButtons.Pass | UserActionButtons.Fail);

                if (IsCanChannel(remark))
                {
                    var name = GetCanChannelName(remark);
                    CleanCanLoop(state, remark);
                }

                try
                {
                    SetChannel(remark + "=" + currentValue, function, state, Resources.lang.ChannelActionReset + " " + function + "=" + currentValue + " /" + remark, true);
                }
                catch (ParserException ex)
                {
                    ex.Cell = remarkCell.ToString();
                    ex.Subcondition = remark + "=" + currentValue;
                    ex.SubconditionLine = 1;
                    ex.Mode = ParserExceptionMode.SETTER;
                    ex.Sheet = Worksheet.Name;
                    throw ex;
                }
                currentValue = currentValue + incr;
            }
        }

        private void BuildBasicSymbols(TestParameters testParameters, string cell, int i)
        {
            var c = 1;
            foreach (var conditionInfo in testParameters.Conditions)
            {
                canLoopSignalUsed.Clear();
                var conditionWithComment = conditionInfo.Condition;

                if (conditionWithComment == "" || conditionWithComment.Substring(0, 2) == "//")
                    continue;

                var testInformation = " / " + c + " / " + conditionInfo.Condition;

                var tableWithComment = conditionWithComment.Split(new string[] { "//" }, StringSplitOptions.None);
                var condition = tableWithComment[0];

                var tab = condition.Split(new string[] { "-", "—" }, 2, StringSplitOptions.None);
                var result = tab[0].Trim().ToLower();
                var operation = tab[1].Trim().ToLower();

                operation = BuildOperation(operation, result, out result, false);

                var tabOperation = operation.Split('且');

                foreach (var op in tabOperation)
                {
                    if (op != "")
                    {
                        try
                        {
                            SetChannel(op, testParameters.Function, testParameters.State, Resources.lang.ChannelActionSetter + " " + testParameters.Function + testInformation);
                        }
                        catch (ParserException ex)
                        {
                            ex.Cell = cell;
                            ex.Subcondition = op;
                            ex.SubconditionLine = c;
                            ex.Mode = ParserExceptionMode.SETTER;
                            ex.Sheet = Worksheet.Name;
                            throw ex;
                        }
                    }
                }

                var myResult = ReplaceKeyword(result);
                var path = (myResult.Contains("flash")) ? CreateGif(i) : CreateImage(i);
                ProjectUtilsFunction.BuildUserAction(testParameters.State, Resources.lang.YouShouldSee + " " + testParameters.Function + " " + result + " " + Resources.lang.With + " " +
                    testParameters.Color + " " + Resources.lang.Light + testInformation, UserActionButtons.Pass | UserActionButtons.Fail, path);

                //reset
                CleanCanLoops(testParameters.State, testInformation);
                foreach (var op in tabOperation)
                {
                    if (op != "")
                    {
                        try
                        {
                            SetChannel(op, testParameters.Function, testParameters.State, Resources.lang.ChannelActionReset + " " + testParameters.Function + testInformation, true);
                        }
                        catch (ParserException ex)
                        {
                            ex.Cell = cell;
                            ex.Subcondition = op;
                            ex.SubconditionLine = c;
                            ex.Mode = ParserExceptionMode.RESET;
                            ex.Sheet = Worksheet.Name;
                            throw ex;
                        }
                    }
                }
                c++;
            }
        }

        public string CreateImage(int i)
        {
            var excelPicture = Worksheet.Drawings["SYM_" + i] as ExcelPicture;
            if (excelPicture == null)
                return null;
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".jpeg";
            excelPicture.Image.Save(path);

            return path;
        }

        public string CreateGif(int i)
        {
            var excelPicture = Worksheet.Drawings["SYM_" + i] as ExcelPicture;
            if (excelPicture == null)
                return null;
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".gif";
            //excelPicture.Image.Save(path);

            GifBitmapEncoder gEnc = new GifBitmapEncoder();

            Bitmap bmpImage = new Bitmap(excelPicture.Image);
            var bmp = bmpImage.GetHbitmap();
            var src = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
                bmp,
                IntPtr.Zero,
                Int32Rect.Empty,
                BitmapSizeOptions.FromEmptyOptions());

            // Try creating a new image with a custom palette.
            List<System.Windows.Media.Color> colors = new List<System.Windows.Media.Color>();
            colors.Add(System.Windows.Media.Color.FromRgb(235, 236, 239));
            BitmapPalette myPalette = new BitmapPalette(colors);

            var srcNull = BitmapSource.Create(bmpImage.Width, bmpImage.Height, 600, 600, PixelFormats.Indexed1, myPalette, new byte[128 * (128 / 8)], 128 / 8);
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(src));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));
            gEnc.Frames.Add(BitmapFrame.Create(srcNull));

            using (MemoryStream ms = new MemoryStream())
            {
                gEnc.Save(ms);

                //Loop
                var fileBytes = ms.ToArray();
                // This is the NETSCAPE2.0 Application Extension.
                var applicationExtension = new byte[] { 33, 255, 11, 78, 69, 84, 83, 67, 65, 80, 69, 50, 46, 48, 3, 1, 0, 0, 0 };
                var newBytes = new List<byte>();
                newBytes.AddRange(fileBytes.Take(13));
                newBytes.AddRange(applicationExtension);
                newBytes.AddRange(fileBytes.Skip(13));
                File.WriteAllBytes(path, newBytes.ToArray());
            }

            return path;
        }
    }
}
