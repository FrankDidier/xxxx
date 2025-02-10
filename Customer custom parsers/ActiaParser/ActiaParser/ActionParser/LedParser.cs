using ActiaParser.Define;
using ArtLogics.TestSuite.Actions;
using ArtLogics.TestSuite.Shared;
using ArtLogics.TestSuite.Testing.Actions.Channels;
using ArtLogics.TestSuite.Testing.Actions.User.UserInputAction;
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
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace ActiaParser.ActionParser
{
    public class LedParser : ActionParserBase
    {
        public ExcelWorksheet Worksheet { get; set; }

        public LedParser(ExcelWorkbook Workbook, Project project) : base (Workbook, project)
        {
        }

        public void ParseActionsLed()
        {
            var titleLineIndex = 4;

            Worksheet = Workbook.Worksheets[ParserStaticVariable.LedSheet];

            var columnNumberDefinition = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.LedsDefinitionColName, Worksheet);
            var columnNumberFunction = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.LedsFunctionColName, Worksheet);
            var colNumberProduct = ExcelUtilsFunction.GetColumnNumber(titleLineIndex, ParserStaticVariable.LedsProductColName, Worksheet);

            var delayTime = int.Parse(IniFile.IniDataRaw["Project"]["Delay"]);

            var state = product.TestCaseGraphs[0].States[product.TestCaseGraphs[0].States.Count - 1];

            for (var i = titleLineIndex + 1; Worksheet.Cells[i, colNumberProduct].Value != null; i++)
            {
                try
                {
                    if (Worksheet.Cells[i, columnNumberDefinition].Value == null || Worksheet.Cells[i, colNumberProduct].Value.ToString() != "Y")
                        continue;

                    var testParameters = new TestParameters(Worksheet, i, columnNumberFunction, columnNumberDefinition, state);

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
                                    SetChannel(op, testParameters.Function, state, Resources.lang.ChannelActionSetter + " " + testParameters.Function + testInformation);
                                }
                                catch (ParserException ex)
                                {
                                    ex.Cell = Worksheet.Cells[i, columnNumberDefinition].ToString();
                                    ex.Subcondition = op;
                                    ex.SubconditionLine = c;
                                    ex.Mode = ParserExceptionMode.SETTER;
                                    ex.Sheet = Worksheet.Name;
                                    throw ex;
                                }
                            }
                        }

                        //userAction

                        var myResult = ReplaceKeyword(result);
                        var path = (myResult.Contains("flash"))?CreateGif(i):CreateImage(i);

                        ProjectUtilsFunction.BuildUserAction(state, Resources.lang.YouShouldSee + " " + testParameters.Function + " " +
                            result + testInformation, UserActionButtons.Pass | UserActionButtons.Fail, path);

                        CleanCanLoops(state, testInformation);
                        foreach (var op in tabOperation)
                        {
                            if (op != "")
                            {
                                try
                                {
                                    SetChannel(op, testParameters.Function, state, Resources.lang.ChannelActionReset + " " + testParameters.Function + testInformation, true);
                                }
                                catch (ParserException ex)
                                {
                                    ex.Cell = Worksheet.Cells[i, columnNumberDefinition].ToString();
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
                catch (Exception err)
                {
                    HandleException(err, Resources.lang.WrongValueInSheet + " " + Worksheet.Name + " " + Resources.lang.AtCell  + " " + Worksheet.Cells[i, columnNumberDefinition]);
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

        public string CreateImage(int i)
        {
            var excelPicture = Worksheet.Drawings["LED_" + i] as ExcelPicture;
            if (excelPicture == null)
                return null;
            var path = ParserStaticVariable.GlobalPath + excelPicture.Name + ".jpeg";
            excelPicture.Image.Save(path);

            return path;
        }

        public string CreateGif(int i)
        {
            var excelPicture = Worksheet.Drawings["LED_" + i] as ExcelPicture;
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
            colors.Add(System.Windows.Media.Color.FromRgb(235,236,239));
            BitmapPalette myPalette = new BitmapPalette(colors);

            var srcNull = BitmapSource.Create(bmpImage.Width, bmpImage.Height, 600, 600, PixelFormats.Indexed1, myPalette, new byte[128 * (128 / 8)], 128/8);
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
