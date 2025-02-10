using ArtLogics.TestSuite.Common;
using ArtLogics.Translation.Parser;
using ArtLogics.Translation.Parser.Interfaces;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.View;
using DevExpress.XtraLayout.Utils;
using Mathtone.MIST;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArtLogics.Translation.ViewModel
{
    [Notifier]
    public class MainFromViewModel : ViewModelBase
    {
        public InputViewModel InputViewModel { get; set; }
        public ProjectConfiguration ProjectConfig { get; internal set; }

        

        public void StartTranslation(object sender, EventArgs e)
        {
            IParser parser = null;
            switch(InputViewModel.ParserTypeEnum)
            {
                case Translation.ViewModel.ParserType.Custom:
                    var assembly = Assembly.LoadFile(InputViewModel.LibraryFile);
                    var type = assembly.GetType("CustomParser.Parser");
                    parser = (IParser)Activator.CreateInstance(type);
                    parser.Parse(InputViewModel.InputFile, InputViewModel.OutputFile, ProjectConfig);
                    DisplayErrors(((BaseParser)parser).LogFileName);
                    break;
                case Translation.ViewModel.ParserType.Excel:
                    break;
                default:
                    break;
            }
            if (parser != null)
                parser.Dispose();
            GC.Collect();
        }

        private void DisplayErrors(string logfile)
        {
            var ErrorDisplay = new ErrorDisplayer();

            //System.Threading.Thread.CurrentThread.CurrentCulture

            ErrorDisplay.ViewModel.LogFile = logfile;

            ErrorDisplay.ShowDialog();
        }
    }
}
