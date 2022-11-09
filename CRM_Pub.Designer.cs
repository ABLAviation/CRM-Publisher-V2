
namespace CRM_Publisher_V2
{
    partial class CRM_Pub : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public CRM_Pub()
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CRM_Pub));
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.btnUpdateCRM = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Label = "CRM Publisher";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.btnUpdateCRM);
            this.group1.Name = "group1";
            // 
            // btnUpdateCRM
            // 
            this.btnUpdateCRM.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnUpdateCRM.Image = ((System.Drawing.Image)(resources.GetObject("btnUpdateCRM.Image")));
            this.btnUpdateCRM.Label = "Update CRM";
            this.btnUpdateCRM.Name = "btnUpdateCRM";
            this.btnUpdateCRM.ShowImage = true;
            this.btnUpdateCRM.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnUpdateCRM_Click);
            // 
            // CRM_Pub
            // 
            this.Name = "CRM_Pub";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.CRM_Pub_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnUpdateCRM;
    }

    partial class ThisRibbonCollection
    {
        internal CRM_Pub CRM_Pub
        {
            get { return this.GetRibbon<CRM_Pub>(); }
        }
    }
}
