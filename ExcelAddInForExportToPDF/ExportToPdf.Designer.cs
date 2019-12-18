using AddInUtilities;

namespace ExcelAddInForExportToPDF
{
    partial class ExportToPdf : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ExportToPdf()
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
            this.VSTOButton = this.Factory.CreateRibbonButton();
            this.separator1 = this.Factory.CreateRibbonSeparator();
            this.asposeButton = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
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
            this.group1.Items.Add(this.VSTOButton);
            this.group1.Items.Add(this.separator1);
            this.group1.Items.Add(this.asposeButton);
            this.group1.Label = "EXPORT TO PDF";
            this.group1.Name = "group1";
            // 
            // VSTOButton
            // 
            this.VSTOButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.VSTOButton.Image = global::ExcelAddInForExportToPDF.Properties.Resources.VisualStudio;
            this.VSTOButton.Label = "VSTO";
            this.VSTOButton.Name = "VSTOButton";
            this.VSTOButton.ShowImage = true;
            this.VSTOButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportButton_Click);
            // 
            // separator1
            // 
            this.separator1.Name = "separator1";
            // 
            // asposeButton
            // 
            this.asposeButton.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.asposeButton.Image = global::ExcelAddInForExportToPDF.Properties.Resources.aspose1;
            this.asposeButton.Label = "Aspose";
            this.asposeButton.Name = "asposeButton";
            this.asposeButton.ShowImage = true;
            this.asposeButton.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.exportButton_Click);
            // 
            // ExportToPdf
            // 
            this.Name = "ExportToPdf";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ExportToPdf_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton VSTOButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton asposeButton;
        internal Microsoft.Office.Tools.Ribbon.RibbonSeparator separator1;
    }

    partial class ThisRibbonCollection
    {
        internal ExportToPdf ExportToPdf
        {
            get { return this.GetRibbon<ExportToPdf>(); }
        }
    }
}
