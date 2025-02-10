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
    public partial class InputsView : UserControl
    {
        private MVVMContextFluentAPI<InputViewModel> fluent;
        public InputViewModel ViewModel => fluent.ViewModel;

        public InputsView()
        {
            InitializeComponent();

            mvvmContext.ViewModelType = typeof(InputViewModel);
            fluent = mvvmContext.OfType<InputViewModel>();

            fluent.SetBinding(this.buttonEditInputFile, v => v.EditValue, vm => vm.InputFile);
            fluent.SetBinding(this.buttonEditOutputFile, v => v.EditValue, vm => vm.OutputFile);
            fluent.SetBinding(this.gridLookUpEditParser, v => v.EditValue, vm => vm.ParserType);
            fluent.SetBinding(this.buttonEditParserLibrary, v => v.EditValue, vm => vm.LibraryFile);
            fluent.SetBinding(this.layoutControlItemParserLibrary, v => v.Visibility, vm => vm.ParserLibraryVisibility);

            Init();
        }

        private void Init()
        {
            this.gridLookUpEditParser.Properties.DataSource = Enum.GetNames(typeof(ParserType));

            this.buttonEditInputFile.Click += ViewModel.SelectInputFile;
            this.buttonEditOutputFile.Click += ViewModel.SelectOutputFile;
            this.buttonEditParserLibrary.Click += ViewModel.SelectParserLibrary;
        }
    }
}
