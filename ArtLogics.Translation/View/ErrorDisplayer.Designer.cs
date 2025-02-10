namespace ArtLogics.Translation.View
{
    partial class ErrorDisplayer
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ErrorDisplayer));
            this.layoutControl1 = new DevExpress.XtraLayout.LayoutControl();
            this.memoEditLogFile = new DevExpress.XtraEditors.MemoEdit();
            this.labelComment = new System.Windows.Forms.Label();
            this.layoutControlGroup1 = new DevExpress.XtraLayout.LayoutControlGroup();
            this.layoutControlItem1 = new DevExpress.XtraLayout.LayoutControlItem();
            this.layoutControlItemLogFile = new DevExpress.XtraLayout.LayoutControlItem();
            this.mvvmContext = new DevExpress.Utils.MVVM.MVVMContext(this.components);
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).BeginInit();
            this.layoutControl1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.memoEditLogFile.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItemLogFile)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.mvvmContext)).BeginInit();
            this.SuspendLayout();
            // 
            // layoutControl1
            // 
            resources.ApplyResources(this.layoutControl1, "layoutControl1");
            this.layoutControl1.Appearance.DisabledLayoutGroupCaption.FontSizeDelta = ((int)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutGroupCaption.FontSizeDelta")));
            this.layoutControl1.Appearance.DisabledLayoutGroupCaption.FontStyleDelta = ((System.Drawing.FontStyle)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutGroupCaption.FontStyleDelta")));
            this.layoutControl1.Appearance.DisabledLayoutGroupCaption.GradientMode = ((System.Drawing.Drawing2D.LinearGradientMode)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutGroupCaption.GradientMode")));
            this.layoutControl1.Appearance.DisabledLayoutGroupCaption.Image = ((System.Drawing.Image)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutGroupCaption.Image")));
            this.layoutControl1.Appearance.DisabledLayoutItem.FontSizeDelta = ((int)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutItem.FontSizeDelta")));
            this.layoutControl1.Appearance.DisabledLayoutItem.FontStyleDelta = ((System.Drawing.FontStyle)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutItem.FontStyleDelta")));
            this.layoutControl1.Appearance.DisabledLayoutItem.GradientMode = ((System.Drawing.Drawing2D.LinearGradientMode)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutItem.GradientMode")));
            this.layoutControl1.Appearance.DisabledLayoutItem.Image = ((System.Drawing.Image)(resources.GetObject("layoutControl1.Appearance.DisabledLayoutItem.Image")));
            this.layoutControl1.Controls.Add(this.memoEditLogFile);
            this.layoutControl1.Controls.Add(this.labelComment);
            this.layoutControl1.Name = "layoutControl1";
            this.layoutControl1.OptionsPrint.AppearanceGroupCaption.FontSizeDelta = ((int)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceGroupCaption.FontSizeDelta")));
            this.layoutControl1.OptionsPrint.AppearanceGroupCaption.FontStyleDelta = ((System.Drawing.FontStyle)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceGroupCaption.FontStyleDelta")));
            this.layoutControl1.OptionsPrint.AppearanceGroupCaption.GradientMode = ((System.Drawing.Drawing2D.LinearGradientMode)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceGroupCaption.GradientMode")));
            this.layoutControl1.OptionsPrint.AppearanceGroupCaption.Image = ((System.Drawing.Image)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceGroupCaption.Image")));
            this.layoutControl1.OptionsPrint.AppearanceItemCaption.FontSizeDelta = ((int)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceItemCaption.FontSizeDelta")));
            this.layoutControl1.OptionsPrint.AppearanceItemCaption.FontStyleDelta = ((System.Drawing.FontStyle)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceItemCaption.FontStyleDelta")));
            this.layoutControl1.OptionsPrint.AppearanceItemCaption.GradientMode = ((System.Drawing.Drawing2D.LinearGradientMode)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceItemCaption.GradientMode")));
            this.layoutControl1.OptionsPrint.AppearanceItemCaption.Image = ((System.Drawing.Image)(resources.GetObject("layoutControl1.OptionsPrint.AppearanceItemCaption.Image")));
            this.layoutControl1.Root = this.layoutControlGroup1;
            // 
            // memoEditLogFile
            // 
            resources.ApplyResources(this.memoEditLogFile, "memoEditLogFile");
            this.memoEditLogFile.Name = "memoEditLogFile";
            this.memoEditLogFile.Properties.AccessibleDescription = resources.GetString("memoEditLogFile.Properties.AccessibleDescription");
            this.memoEditLogFile.Properties.AccessibleName = resources.GetString("memoEditLogFile.Properties.AccessibleName");
            this.memoEditLogFile.Properties.NullValuePrompt = resources.GetString("memoEditLogFile.Properties.NullValuePrompt");
            this.memoEditLogFile.Properties.NullValuePromptShowForEmptyValue = ((bool)(resources.GetObject("memoEditLogFile.Properties.NullValuePromptShowForEmptyValue")));
            this.memoEditLogFile.Properties.ReadOnly = true;
            this.memoEditLogFile.Properties.ScrollBars = System.Windows.Forms.ScrollBars.Both;
            this.memoEditLogFile.StyleController = this.layoutControl1;
            // 
            // labelComment
            // 
            resources.ApplyResources(this.labelComment, "labelComment");
            this.labelComment.Name = "labelComment";
            // 
            // layoutControlGroup1
            // 
            this.layoutControlGroup1.EnableIndentsWithoutBorders = DevExpress.Utils.DefaultBoolean.True;
            this.layoutControlGroup1.GroupBordersVisible = false;
            this.layoutControlGroup1.Items.AddRange(new DevExpress.XtraLayout.BaseLayoutItem[] {
            this.layoutControlItem1,
            this.layoutControlItemLogFile});
            this.layoutControlGroup1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlGroup1.Name = "layoutControlGroup1";
            this.layoutControlGroup1.Size = new System.Drawing.Size(593, 245);
            this.layoutControlGroup1.TextVisible = false;
            // 
            // layoutControlItem1
            // 
            resources.ApplyResources(this.layoutControlItem1, "layoutControlItem1");
            this.layoutControlItem1.Control = this.labelComment;
            this.layoutControlItem1.Location = new System.Drawing.Point(0, 0);
            this.layoutControlItem1.Name = "layoutControlItem1";
            this.layoutControlItem1.Size = new System.Drawing.Size(573, 34);
            this.layoutControlItem1.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItem1.TextVisible = false;
            // 
            // layoutControlItemLogFile
            // 
            resources.ApplyResources(this.layoutControlItemLogFile, "layoutControlItemLogFile");
            this.layoutControlItemLogFile.Control = this.memoEditLogFile;
            this.layoutControlItemLogFile.Location = new System.Drawing.Point(0, 34);
            this.layoutControlItemLogFile.Name = "layoutControlItemErrorLog";
            this.layoutControlItemLogFile.Size = new System.Drawing.Size(573, 191);
            this.layoutControlItemLogFile.TextSize = new System.Drawing.Size(0, 0);
            this.layoutControlItemLogFile.TextVisible = false;
            // 
            // mvvmContext
            // 
            this.mvvmContext.ContainerControl = this;
            // 
            // ErrorDisplayer
            // 
            resources.ApplyResources(this, "$this");
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.layoutControl1);
            this.Name = "ErrorDisplayer";
            ((System.ComponentModel.ISupportInitialize)(this.layoutControl1)).EndInit();
            this.layoutControl1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.memoEditLogFile.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlGroup1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItem1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.layoutControlItemLogFile)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.mvvmContext)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraLayout.LayoutControl layoutControl1;
        private DevExpress.XtraLayout.LayoutControlGroup layoutControlGroup1;
        private DevExpress.XtraEditors.MemoEdit memoEditLogFile;
        private System.Windows.Forms.Label labelComment;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItem1;
        private DevExpress.XtraLayout.LayoutControlItem layoutControlItemLogFile;
        private DevExpress.Utils.MVVM.MVVMContext mvvmContext;
    }
}