namespace WordAddIn
{
    partial class TDMS : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public TDMS()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(TDMS));
            this._tdms = this.Factory.CreateRibbonTab();
            this.Save = this.Factory.CreateRibbonGroup();
            this.btnSave = this.Factory.CreateRibbonButton();
            this.btnSaveAndClose = this.Factory.CreateRibbonButton();
            this.btnClose = this.Factory.CreateRibbonButton();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnUpdate = this.Factory.CreateRibbonButton();
            this.help = this.Factory.CreateRibbonGroup();
            this.btnHelp = this.Factory.CreateRibbonButton();
            this.btnSTP = this.Factory.CreateRibbonButton();
            this.btnFAQ = this.Factory.CreateRibbonButton();
            this.info = this.Factory.CreateRibbonGroup();
            this.btnOprogr = this.Factory.CreateRibbonButton();
            this._tdms.SuspendLayout();
            this.Save.SuspendLayout();
            this.group1.SuspendLayout();
            this.help.SuspendLayout();
            this.info.SuspendLayout();
            this.SuspendLayout();
            // 
            // _tdms
            // 
            this._tdms.Groups.Add(this.Save);
            this._tdms.Groups.Add(this.group1);
            this._tdms.Groups.Add(this.help);
            this._tdms.Groups.Add(this.info);
            this._tdms.Label = "TDMS";
            this._tdms.Name = "_tdms";
            // 
            // Save
            // 
            this.Save.Items.Add(this.btnSave);
            this.Save.Items.Add(this.btnSaveAndClose);
            this.Save.Items.Add(this.btnClose);
            this.Save.Label = "Сохранить";
            this.Save.Name = "Save";
            // 
            // btnSave
            // 
            this.btnSave.Image = ((System.Drawing.Image)(resources.GetObject("btnSave.Image")));
            this.btnSave.Label = "Сохранить";
            this.btnSave.Name = "btnSave";
            this.btnSave.ShowImage = true;
            this.btnSave.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSave_Click);
            // 
            // btnSaveAndClose
            // 
            this.btnSaveAndClose.Image = ((System.Drawing.Image)(resources.GetObject("btnSaveAndClose.Image")));
            this.btnSaveAndClose.Label = "Сохранить и закрыть";
            this.btnSaveAndClose.Name = "btnSaveAndClose";
            this.btnSaveAndClose.ShowImage = true;
            this.btnSaveAndClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSaveAndClose_Click);
            // 
            // btnClose
            // 
            this.btnClose.Image = ((System.Drawing.Image)(resources.GetObject("btnClose.Image")));
            this.btnClose.Label = "Закрыть без сохранения";
            this.btnClose.Name = "btnClose";
            this.btnClose.ShowImage = true;
            this.btnClose.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnClose_Click);
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnUpdate);
            this.group1.Label = "Обновление";
            this.group1.Name = "group1";
            // 
            // btnUpdate
            // 
            this.btnUpdate.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdate.Image")));
            this.btnUpdate.Label = "Обновить атрибуты";
            this.btnUpdate.Name = "btnUpdate";
            this.btnUpdate.ShowImage = true;
            this.btnUpdate.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdate_Click);
            // 
            // help
            // 
            this.help.Items.Add(this.btnHelp);
            this.help.Items.Add(this.btnSTP);
            this.help.Items.Add(this.btnFAQ);
            this.help.Label = "Помощь";
            this.help.Name = "help";
            // 
            // btnHelp
            // 
            this.btnHelp.Image = ((System.Drawing.Image)(resources.GetObject("btnHelp.Image")));
            this.btnHelp.Label = "Справочные материалы";
            this.btnHelp.Name = "btnHelp";
            this.btnHelp.ShowImage = true;
            this.btnHelp.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnHelp_Click);
            // 
            // btnSTP
            // 
            this.btnSTP.Image = ((System.Drawing.Image)(resources.GetObject("btnSTP.Image")));
            this.btnSTP.Label = "Стандарт предприятия";
            this.btnSTP.Name = "btnSTP";
            this.btnSTP.ShowImage = true;
            this.btnSTP.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSTP_Click);
            // 
            // btnFAQ
            // 
            this.btnFAQ.Image = ((System.Drawing.Image)(resources.GetObject("btnFAQ.Image")));
            this.btnFAQ.Label = "Часто задаваемые вопросы";
            this.btnFAQ.Name = "btnFAQ";
            this.btnFAQ.ShowImage = true;
            this.btnFAQ.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFAQ_Click);
            // 
            // info
            // 
            this.info.Items.Add(this.btnOprogr);
            this.info.Label = "О программе";
            this.info.Name = "info";
            // 
            // btnOprogr
            // 
            this.btnOprogr.Image = ((System.Drawing.Image)(resources.GetObject("btnOprogr.Image")));
            this.btnOprogr.Label = "О программе";
            this.btnOprogr.Name = "btnOprogr";
            this.btnOprogr.ShowImage = true;
            this.btnOprogr.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnOprogr_Click);
            // 
            // TDMS
            // 
            this.Name = "TDMS";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this._tdms);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.TDMS_Load);
            this._tdms.ResumeLayout(false);
            this._tdms.PerformLayout();
            this.Save.ResumeLayout(false);
            this.Save.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.help.ResumeLayout(false);
            this.help.PerformLayout();
            this.info.ResumeLayout(false);
            this.info.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab _tdms;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Save;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSave;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSaveAndClose;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnClose;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdate;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup help;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnHelp;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSTP;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFAQ;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup info;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnOprogr;
    }

    partial class ThisRibbonCollection
    {
        internal TDMS RibbonControl
        {
            get { return this.GetRibbon<TDMS>(); }
        }
    }
}
