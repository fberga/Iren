namespace Iren.PSO.Applicazioni
{
    partial class ToolsExcelRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public ToolsExcelRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Liberare le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione componenti

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.FrontOffice = this.Factory.CreateRibbonTab();
            this.TabHome = this.Factory.CreateRibbonTab();
            this.TabInsert = this.Factory.CreateRibbonTab();
            this.TabPageLayoutExcel = this.Factory.CreateRibbonTab();
            this.TabFormulas = this.Factory.CreateRibbonTab();
            this.TabData = this.Factory.CreateRibbonTab();
            this.TabReview = this.Factory.CreateRibbonTab();
            this.TabView = this.Factory.CreateRibbonTab();
            this.TabDeveloper = this.Factory.CreateRibbonTab();
            this.TabAddIns = this.Factory.CreateRibbonTab();
            this.TabPrintPreview = this.Factory.CreateRibbonTab();
            this.TabBackgroundRemoval = this.Factory.CreateRibbonTab();
            this.TabSmartArtToolsDesign = this.Factory.CreateRibbonTab();
            this.FrontOffice.SuspendLayout();
            this.TabHome.SuspendLayout();
            this.TabInsert.SuspendLayout();
            this.TabPageLayoutExcel.SuspendLayout();
            this.TabFormulas.SuspendLayout();
            this.TabData.SuspendLayout();
            this.TabReview.SuspendLayout();
            this.TabView.SuspendLayout();
            this.TabDeveloper.SuspendLayout();
            this.TabAddIns.SuspendLayout();
            this.TabPrintPreview.SuspendLayout();
            this.TabBackgroundRemoval.SuspendLayout();
            this.TabSmartArtToolsDesign.SuspendLayout();

            InitializeComponent2();
            
            // 
            // FrontOffice
            // 
            this.FrontOffice.Label = "Front Office";
            this.FrontOffice.Name = "FrontOffice";
            // 
            // TabHome
            // 
            this.TabHome.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabHome.ControlId.OfficeId = "TabHome";
            this.TabHome.Label = "TabHome";
            this.TabHome.Name = "TabHome";
            // 
            // TabInsert
            // 
            this.TabInsert.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabInsert.ControlId.OfficeId = "TabInsert";
            this.TabInsert.Label = "TabInsert";
            this.TabInsert.Name = "TabInsert";
            // 
            // TabPageLayoutExcel
            // 
            this.TabPageLayoutExcel.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabPageLayoutExcel.ControlId.OfficeId = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Label = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Name = "TabPageLayoutExcel";
            this.TabPageLayoutExcel.Visible = false;
            // 
            // TabFormulas
            // 
            this.TabFormulas.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabFormulas.ControlId.OfficeId = "TabFormulas";
            this.TabFormulas.Label = "TabFormulas";
            this.TabFormulas.Name = "TabFormulas";
            // 
            // TabData
            // 
            this.TabData.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabData.ControlId.OfficeId = "TabData";
            this.TabData.Label = "TabData";
            this.TabData.Name = "TabData";
            this.TabData.Visible = false;
            // 
            // TabReview
            // 
            this.TabReview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabReview.ControlId.OfficeId = "TabReview";
            this.TabReview.Label = "TabReview";
            this.TabReview.Name = "TabReview";
            // 
            // TabView
            // 
            this.TabView.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabView.ControlId.OfficeId = "TabView";
            this.TabView.Label = "TabView";
            this.TabView.Name = "TabView";
            // 
            // TabDeveloper
            // 
            this.TabDeveloper.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabDeveloper.ControlId.OfficeId = "TabDeveloper";
            this.TabDeveloper.Label = "TabDeveloper";
            this.TabDeveloper.Name = "TabDeveloper";
            // 
            // TabAddIns
            // 
            this.TabAddIns.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabAddIns.Label = "TabAddIns";
            this.TabAddIns.Name = "TabAddIns";
            this.TabAddIns.Visible = false;
            // 
            // TabPrintPreview
            // 
            this.TabPrintPreview.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabPrintPreview.ControlId.OfficeId = "TabPrintPreview";
            this.TabPrintPreview.Label = "TabPrintPreview";
            this.TabPrintPreview.Name = "TabPrintPreview";
            this.TabPrintPreview.Visible = false;
            // 
            // TabBackgroundRemoval
            // 
            this.TabBackgroundRemoval.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabBackgroundRemoval.ControlId.OfficeId = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Label = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Name = "TabBackgroundRemoval";
            this.TabBackgroundRemoval.Visible = false;
            // 
            // TabSmartArtToolsDesign
            // 
            this.TabSmartArtToolsDesign.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.TabSmartArtToolsDesign.ControlId.OfficeId = "TabSmartArtToolsDesign";
            this.TabSmartArtToolsDesign.Label = "TabSmartArtToolsDesign";
            this.TabSmartArtToolsDesign.Name = "TabSmartArtToolsDesign";
            // 
            // ToolsExcelRibbon
            // 
            this.Name = "ToolsExcelRibbon";
            this.RibbonType = "Microsoft.Excel.Workbook";
            this.StartFromScratch = true;
            this.Tabs.Add(this.FrontOffice);
            this.Tabs.Add(this.TabHome);
            this.Tabs.Add(this.TabInsert);
            this.Tabs.Add(this.TabPageLayoutExcel);
            this.Tabs.Add(this.TabFormulas);
            this.Tabs.Add(this.TabData);
            this.Tabs.Add(this.TabReview);
            this.Tabs.Add(this.TabView);
            this.Tabs.Add(this.TabDeveloper);
            this.Tabs.Add(this.TabAddIns);
            this.Tabs.Add(this.TabPrintPreview);
            this.Tabs.Add(this.TabBackgroundRemoval);
            this.Tabs.Add(this.TabSmartArtToolsDesign);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.ToolsExcelRibbon_Load);
            this.FrontOffice.ResumeLayout(false);
            this.FrontOffice.PerformLayout();
            this.TabHome.ResumeLayout(false);
            this.TabHome.PerformLayout();
            this.TabInsert.ResumeLayout(false);
            this.TabInsert.PerformLayout();
            this.TabPageLayoutExcel.ResumeLayout(false);
            this.TabPageLayoutExcel.PerformLayout();
            this.TabFormulas.ResumeLayout(false);
            this.TabFormulas.PerformLayout();
            this.TabData.ResumeLayout(false);
            this.TabData.PerformLayout();
            this.TabReview.ResumeLayout(false);
            this.TabReview.PerformLayout();
            this.TabView.ResumeLayout(false);
            this.TabView.PerformLayout();
            this.TabDeveloper.ResumeLayout(false);
            this.TabDeveloper.PerformLayout();
            this.TabAddIns.ResumeLayout(false);
            this.TabAddIns.PerformLayout();
            this.TabPrintPreview.ResumeLayout(false);
            this.TabPrintPreview.PerformLayout();
            this.TabBackgroundRemoval.ResumeLayout(false);
            this.TabBackgroundRemoval.PerformLayout();
            this.TabSmartArtToolsDesign.ResumeLayout(false);
            this.TabSmartArtToolsDesign.PerformLayout();

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab TabHome;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabInsert;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabPageLayoutExcel;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabFormulas;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabData;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabReview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabView;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabDeveloper;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabAddIns;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabPrintPreview;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabBackgroundRemoval;
        public Microsoft.Office.Tools.Ribbon.RibbonTab FrontOffice;
        private Microsoft.Office.Tools.Ribbon.RibbonTab TabSmartArtToolsDesign;
    }

    partial class ThisRibbonCollection
    {
        internal ToolsExcelRibbon ToolsExcelRibbon
        {
            get { return this.GetRibbon<ToolsExcelRibbon>(); }
        }
    }
}
