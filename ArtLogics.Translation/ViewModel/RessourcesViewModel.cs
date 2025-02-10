using ArtLogics.TestSuite.Common;
using ArtLogics.Translation.Parser.Model;
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
    public class ResourcesViewModel : ViewModelBase
    {
        public BindingList<Parser.Model.Resources> Resources { get; set; }

        public ResourcesViewModel()
        {
            Resources = new BindingList<Parser.Model.Resources>();
        }
    }
}
