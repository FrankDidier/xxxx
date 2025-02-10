using ArtLogics.TestSuite.Common;
using ArtLogics.Translation.Parser.Model;
using IniParser.Model;
using Mathtone.MIST;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArtLogics.Translation.ViewModel
{
    [Notifier]
    public class IniViewModel : ViewModelBase
    {
        [Notify]
        public BindingList<Tuple<string, dynamic>> IniData { get; set; }

        public IniViewModel()
        {
        }
    }
}
