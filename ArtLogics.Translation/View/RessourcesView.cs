using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.Utils.MVVM;
using ArtLogics.Translation.ViewModel;

namespace ArtLogics.Translation.View
{
    public partial class RessourcesView : UserControl
    {
        private MVVMContextFluentAPI<ResourcesViewModel> fluent;
        public ResourcesViewModel ViewModel => fluent.ViewModel;

        public RessourcesView()
        {
            InitializeComponent();

            mvvmContext.ViewModelType = typeof(ResourcesViewModel);
            fluent = mvvmContext.OfType<ResourcesViewModel>();

            fluent.SetBinding(this.gridControlRessources, v => v.DataSource, vm => vm.Resources);
        }
    }
}
