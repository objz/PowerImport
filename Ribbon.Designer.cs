namespace PowerImport
{
    partial class Ribbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon()
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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.PowerImport = this.Factory.CreateRibbonGroup();
            this.import = this.Factory.CreateRibbonButton();
            this.refresh = this.Factory.CreateRibbonButton();
            this.refresh_all = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.PowerImport.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.ControlId.OfficeId = "TabData";
            this.tab1.Groups.Add(this.PowerImport);
            this.tab1.Label = "TabData";
            this.tab1.Name = "tab1";
            // 
            // PowerImport
            // 
            this.PowerImport.Items.Add(this.import);
            this.PowerImport.Items.Add(this.refresh);
            this.PowerImport.Items.Add(this.refresh_all);
            this.PowerImport.Label = "Power Import";
            this.PowerImport.Name = "PowerImport";
            this.PowerImport.Position = this.Factory.RibbonPosition.BeforeOfficeId("GroupDataQueriesAndConnections");
            // 
            // import
            // 
            this.import.Image = global::PowerImport.Properties.Resources.import;
            this.import.Label = "Import Queries";
            this.import.Name = "import";
            this.import.ShowImage = true;
            this.import.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.import_Click);
            // 
            // refresh
            // 
            this.refresh.Image = global::PowerImport.Properties.Resources.refresh;
            this.refresh.Label = "Refresh";
            this.refresh.Name = "refresh";
            this.refresh.ShowImage = true;
            this.refresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refresh_Click);
            // 
            // refresh_all
            // 
            this.refresh_all.Image = global::PowerImport.Properties.Resources.refresh_all;
            this.refresh_all.Label = "Refresh All";
            this.refresh_all.Name = "refresh_all";
            this.refresh_all.ShowImage = true;
            this.refresh_all.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.refresh_all_Click);
            // 
            // Ribbon
            // 
            this.Name = "Ribbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.PowerImport.ResumeLayout(false);
            this.PowerImport.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup PowerImport;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton import;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton refresh_all;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon Ribbon
        {
            get { return this.GetRibbon<Ribbon>(); }
        }
    }
}
