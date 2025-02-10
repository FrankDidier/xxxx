using ArtLogics.Translation.ViewModel;
using DevExpress.Utils.MVVM;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ArtLogics.Translation.View
{
    public partial class ErrorDisplayer : Form
    {
        private MVVMContextFluentAPI<ErrorDisplayerViewModel> fluent;
        public ErrorDisplayerViewModel ViewModel => fluent.ViewModel;

        public ErrorDisplayer()
        {
            InitializeComponent();

            mvvmContext.ViewModelType = typeof(ErrorDisplayerViewModel);
            fluent = mvvmContext.OfType<ErrorDisplayerViewModel>();

            fluent.SetBinding(this.labelComment, v => v.Text, vm => vm.Comment);
            fluent.SetBinding(this.memoEditLogFile, v => v.Text, vm => vm.FileContent);
            fluent.SetBinding(this.layoutControlItemLogFile, v => v.Visibility, vm => vm.ShowLog);
        }
    }
}
