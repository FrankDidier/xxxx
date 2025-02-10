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
using DevExpress.XtraGrid.Views.Grid;

namespace ArtLogics.Translation.View
{
    public partial class IniView : UserControl
    {
        private MVVMContextFluentAPI<IniViewModel> fluent;
        public IniViewModel ViewModel => fluent.ViewModel;

        public IniView()
        {
            InitializeComponent();

            mvvmContext.ViewModelType = typeof(IniViewModel);
            fluent = mvvmContext.OfType<IniViewModel>();

            fluent.SetBinding(this.gridControlIniValue, v => v.DataSource, vm => vm.IniData);
        }
    }
}
