using ArtLogics.TestSuite.Common;
using DevExpress.XtraLayout.Utils;
using Mathtone.MIST;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.ViewModel
{
    [Notifier]
    public class ErrorDisplayerViewModel : ViewModelBase
    {
        [Notify(nameof(Comment))]
        public string Comment { get; set; }

        [Notify(nameof(ShowLog))]
        public LayoutVisibility ShowLog { get; set; }

        [Notify(nameof(FileContent))]
        public string FileContent { get; set; }

        private string _logFile;

        [Notify(nameof(FileContent), nameof(Comment), nameof(ShowLog))]
        public string LogFile {
            get => _logFile;
            set
            {
                _logFile = value.Replace('\\','/');
                if (File.Exists(_logFile))
                {
                    FileContent = File.ReadAllText(_logFile);
                    Comment = (FileContent.Split('\n').Count() - 1) + " error occured during the parsing, please check the below log and your excel file";
                    ShowLog = LayoutVisibility.Always;
                } else
                {
                    Comment = "Success";
                    ShowLog = LayoutVisibility.Never;
                }
            }
        }
    }
}
