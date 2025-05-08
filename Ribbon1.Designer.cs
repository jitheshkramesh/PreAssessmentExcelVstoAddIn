namespace PreAssessmentExcelVstoAddIn
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.TextTools = this.Factory.CreateRibbonTab();
            this.Actions = this.Factory.CreateRibbonGroup();
            this.btnConvert = this.Factory.CreateRibbonButton();
            this.btnRevert = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.TextTools.SuspendLayout();
            this.Actions.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // TextTools
            // 
            this.TextTools.Groups.Add(this.Actions);
            this.TextTools.Label = "Text Tools";
            this.TextTools.Name = "TextTools";
            // 
            // Actions
            // 
            this.Actions.Items.Add(this.btnConvert);
            this.Actions.Items.Add(this.btnRevert);
            this.Actions.Label = "Actions";
            this.Actions.Name = "Actions";
            // 
            // btnConvert
            // 
            this.btnConvert.Label = "Convert to Alphanumeric";
            this.btnConvert.Name = "btnConvert";
            this.btnConvert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnConvert_Click);
            // 
            // btnRevert
            // 
            this.btnRevert.Label = "Revert to Original";
            this.btnRevert.Name = "btnRevert";
            this.btnRevert.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRevert_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Tabs.Add(this.TextTools);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.TextTools.ResumeLayout(false);
            this.TextTools.PerformLayout();
            this.Actions.ResumeLayout(false);
            this.Actions.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab TextTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup Actions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnConvert;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRevert;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}
