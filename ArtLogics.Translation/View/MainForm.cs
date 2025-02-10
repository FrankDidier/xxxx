using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ArtLogics.Translation.Parser.Model;
using ArtLogics.Translation.ViewModel;
using DevExpress.Utils;
using DevExpress.Utils.MVVM;
using DevExpress.XtraBars;
using DevExpress.XtraBars.Ribbon;

namespace ArtLogics.Translation.View
{
    public partial class MainForm : RibbonForm
    {
        private MVVMContextFluentAPI<MainFromViewModel> fluent;
        private MainFromViewModel ViewModel => fluent.ViewModel;

        public string Language { get; set; } = "en";
        public bool Exit { get; set; } = true;

        public MainForm()
        {
            InitializeComponent();

            mvvmContext.ViewModelType = typeof(MainFromViewModel);
            fluent = mvvmContext.OfType<MainFromViewModel>();

            Init();
        }

        private void Init()
        {
            ViewModel.InputViewModel = this.inputsView.ViewModel;
            var IniFIle = new IniFile();
            this.iniView1.ViewModel.IniData = IniFIle.IniData;
            var projectConfig = new ProjectConfiguration();
            this.ressourcesView.ViewModel.Resources = projectConfig.Resources;
            ViewModel.ProjectConfig = projectConfig;
            this.simpleButtonGenerateDataBase.Click += ViewModel.StartTranslation;

            this.barButtonItemInput.ItemClick += ClickInput;
            this.barButtonItemConfiguration.ItemClick += ClickConfiguation;
            this.barButtonItemRessources.ItemClick += ClickRessources;

            this.barButtonItemEnglish.ItemClick += ModifyLanguage;
            this.barButtonItemChinese.ItemClick += ModifyLanguage;

            if (Thread.CurrentThread.CurrentCulture != null && Thread.CurrentThread.CurrentCulture.ToString().ToLower() == "zh-cn")
                this.barLinkContainerItemLanguageList.ImageOptions.Image = Properties.Resources.flag_china;
            else
                this.barLinkContainerItemLanguageList.ImageOptions.Image = Properties.Resources.Flag_UK;

        }

        private void ModifyLanguage(object sender, ItemClickEventArgs e)
        {
            Language = e.Item.Tag.ToString();
            Exit = false;
            this.Close();
        }

        private void ClickRessources(object sender, ItemClickEventArgs e)
        {
            this.xtraTabControlInput.SelectedTabPageIndex = 2;
        }

        private void ClickConfiguation(object sender, ItemClickEventArgs e)
        {
            this.xtraTabControlInput.SelectedTabPageIndex = 1;
        }

        private void ClickInput(object sender, ItemClickEventArgs e)
        {
            this.xtraTabControlInput.SelectedTabPageIndex = 0;
        }
    }
}
