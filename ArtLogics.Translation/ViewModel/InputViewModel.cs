using ArtLogics.TestSuite.Common;
using ArtLogics.Translation.Parser.Interfaces;
using DevExpress.XtraLayout.Utils;
using Mathtone.MIST;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArtLogics.Translation.ViewModel
{
    [Notifier]
    public class InputViewModel : ViewModelBase
    {
        [Notify]
        public string InputFile { get; set; }
        [Notify]
        public string OutputFile { get; set; }

        [Notify(nameof(ParserType), nameof(ParserLibraryVisibility))]
        public string ParserType
        {
            get => ParserTypeEnum.ToString();
            set
            {
                ParserTypeEnum = (ParserType)Enum.Parse(typeof(ParserType), value);
            }
        }

        public LayoutVisibility ParserLibraryVisibility => (this.ParserTypeEnum == Translation.ViewModel.ParserType.Custom) ? LayoutVisibility.Always : LayoutVisibility.Never;

        [Notify]
        public string LibraryFile { get; set; }

        private string projectPath = "C:/ART logics/Projects/";
        private string projectExtension = "apdb";
        public ParserType ParserTypeEnum { get; set; }

        public void SelectInputFile(object sender, EventArgs e)
        {
            var openDialog = new OpenFileDialog();
            openDialog.CheckFileExists = true;

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                InputFile = openDialog.FileName;
                var fileName = Path.GetFileNameWithoutExtension(openDialog.FileName);
                OutputFile = projectPath + fileName + "." + projectExtension;
            }
        }

        public void SelectOutputFile(object sender, EventArgs e)
        {
            var saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Project File|*.apdb";
            saveDialog.FileName = OutputFile;
            saveDialog.InitialDirectory = projectPath;

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                OutputFile = saveDialog.FileName;
            }
        }

        public void SelectParserLibrary(object sender, EventArgs e)
        {
            var openDialog = new OpenFileDialog();
            openDialog.CheckFileExists = true;
            openDialog.Filter = "Library File|*.dll";

            if (openDialog.ShowDialog() == DialogResult.OK)
            {
                LibraryFile = openDialog.FileName;
            }
        }
    }
}
