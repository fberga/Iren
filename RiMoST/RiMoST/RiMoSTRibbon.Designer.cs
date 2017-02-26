namespace Iren.RiMoST
{
    partial class RiMoSTRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RiMoSTRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Pulire le risorse in uso.
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
            this.tbRichiestaModifica = this.Factory.CreateRibbonTab();
            this.groupChiudi = this.Factory.CreateRibbonGroup();
            this.btnChiudi = this.Factory.CreateRibbonButton();
            this.groupAzioni = this.Factory.CreateRibbonGroup();
            this.btnInvia = this.Factory.CreateRibbonButton();
            this.chkIsDraft = this.Factory.CreateRibbonToggleButton();
            this.groupModifica = this.Factory.CreateRibbonGroup();
            this.btnReset = this.Factory.CreateRibbonButton();
            this.btnRefresh = this.Factory.CreateRibbonButton();
            this.btnPrint = this.Factory.CreateRibbonButton();
            this.groupGestione = this.Factory.CreateRibbonGroup();
            this.lbAnniDisponibili = this.Factory.CreateRibbonLabel();
            this.cbAnniDisponibili = this.Factory.CreateRibbonComboBox();
            this.btnAnnulla = this.Factory.CreateRibbonButton();
            this.btnModifica = this.Factory.CreateRibbonButton();
            this.tabVersione = this.Factory.CreateRibbonTab();
            this.groupVersione = this.Factory.CreateRibbonGroup();
            this.lbVersioneLabel = this.Factory.CreateRibbonLabel();
            this.lbVersioneApp = this.Factory.CreateRibbonLabel();
            this.lbCoreV = this.Factory.CreateRibbonLabel();
            this.tbRichiestaModifica.SuspendLayout();
            this.groupChiudi.SuspendLayout();
            this.groupAzioni.SuspendLayout();
            this.groupModifica.SuspendLayout();
            this.groupGestione.SuspendLayout();
            this.tabVersione.SuspendLayout();
            this.groupVersione.SuspendLayout();
            // 
            // tbRichiestaModifica
            // 
            this.tbRichiestaModifica.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tbRichiestaModifica.Groups.Add(this.groupChiudi);
            this.tbRichiestaModifica.Groups.Add(this.groupAzioni);
            this.tbRichiestaModifica.Groups.Add(this.groupModifica);
            this.tbRichiestaModifica.Groups.Add(this.groupGestione);
            this.tbRichiestaModifica.Label = "Richiesta Modifica";
            this.tbRichiestaModifica.Name = "tbRichiestaModifica";
            // 
            // groupChiudi
            // 
            this.groupChiudi.Items.Add(this.btnChiudi);
            this.groupChiudi.Label = "Chiudi";
            this.groupChiudi.Name = "groupChiudi";
            // 
            // btnChiudi
            // 
            this.btnChiudi.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnChiudi.Image = global::Iren.RiMoST.Properties.Resources.Close;
            this.btnChiudi.Label = "Chiudi";
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.ShowImage = true;
            this.btnChiudi.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnChiudi_Click);
            // 
            // groupAzioni
            // 
            this.groupAzioni.Items.Add(this.btnInvia);
            this.groupAzioni.Items.Add(this.chkIsDraft);
            this.groupAzioni.Label = "Azioni";
            this.groupAzioni.Name = "groupAzioni";
            // 
            // btnInvia
            // 
            this.btnInvia.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnInvia.Image = global::Iren.RiMoST.Properties.Resources.Send;
            this.btnInvia.Label = "Conferma e invia";
            this.btnInvia.Name = "btnInvia";
            this.btnInvia.ShowImage = true;
            this.btnInvia.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnInvia_Click);
            // 
            // chkIsDraft
            // 
            this.chkIsDraft.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.chkIsDraft.Image = global::Iren.RiMoST.Properties.Resources.Draft;
            this.chkIsDraft.Label = "Bozza";
            this.chkIsDraft.Name = "chkIsDraft";
            this.chkIsDraft.ShowImage = true;
            this.chkIsDraft.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.chkIsDraft_Click);
            // 
            // groupModifica
            // 
            this.groupModifica.Items.Add(this.btnReset);
            this.groupModifica.Items.Add(this.btnRefresh);
            this.groupModifica.Items.Add(this.btnPrint);
            this.groupModifica.Label = "Modifica";
            this.groupModifica.Name = "groupModifica";
            // 
            // btnReset
            // 
            this.btnReset.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnReset.Image = global::Iren.RiMoST.Properties.Resources.New;
            this.btnReset.Label = "Nuova Modifica";
            this.btnReset.Name = "btnReset";
            this.btnReset.ShowImage = true;
            this.btnReset.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReset_Click);
            // 
            // btnRefresh
            // 
            this.btnRefresh.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnRefresh.Image = global::Iren.RiMoST.Properties.Resources.Refresh;
            this.btnRefresh.Label = "Aggiorna n°";
            this.btnRefresh.Name = "btnRefresh";
            this.btnRefresh.ShowImage = true;
            this.btnRefresh.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnRefresh_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnPrint.Image = global::Iren.RiMoST.Properties.Resources.Print_icon;
            this.btnPrint.Label = "Stampa";
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.ShowImage = true;
            this.btnPrint.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnPrint_Click);
            // 
            // groupGestione
            // 
            this.groupGestione.Items.Add(this.lbAnniDisponibili);
            this.groupGestione.Items.Add(this.cbAnniDisponibili);
            this.groupGestione.Items.Add(this.btnAnnulla);
            this.groupGestione.Items.Add(this.btnModifica);
            this.groupGestione.Label = "Gestione";
            this.groupGestione.Name = "groupGestione";
            // 
            // lbAnniDisponibili
            // 
            this.lbAnniDisponibili.Label = "Filtra per anno:";
            this.lbAnniDisponibili.Name = "lbAnniDisponibili";
            // 
            // cbAnniDisponibili
            // 
            this.cbAnniDisponibili.Image = global::Iren.RiMoST.Properties.Resources.Calendar;
            this.cbAnniDisponibili.Label = "comboBox1";
            this.cbAnniDisponibili.Name = "cbAnniDisponibili";
            this.cbAnniDisponibili.ScreenTip = "Seleziona l\'anno";
            this.cbAnniDisponibili.ShowImage = true;
            this.cbAnniDisponibili.ShowLabel = false;
            this.cbAnniDisponibili.Text = null;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnAnnulla.Image = global::Iren.RiMoST.Properties.Resources.Bin;
            this.btnAnnulla.Label = "Annulla una richiesta";
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.ShowImage = true;
            this.btnAnnulla.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAnnulla_Click);
            // 
            // btnModifica
            // 
            this.btnModifica.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnModifica.Image = global::Iren.RiMoST.Properties.Resources.Edit;
            this.btnModifica.Label = "Modifica una bozza";
            this.btnModifica.Name = "btnModifica";
            this.btnModifica.ShowImage = true;
            this.btnModifica.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnModifica_Click);
            // 
            // tabVersione
            // 
            this.tabVersione.Groups.Add(this.groupVersione);
            this.tabVersione.Label = "Versione";
            this.tabVersione.Name = "tabVersione";
            // 
            // groupVersione
            // 
            this.groupVersione.Items.Add(this.lbVersioneLabel);
            this.groupVersione.Items.Add(this.lbVersioneApp);
            this.groupVersione.Items.Add(this.lbCoreV);
            this.groupVersione.Label = "System Info";
            this.groupVersione.Name = "groupVersione";
            // 
            // lbVersioneLabel
            // 
            this.lbVersioneLabel.Label = "RiMoST                    ";
            this.lbVersioneLabel.Name = "lbVersioneLabel";
            // 
            // lbVersioneApp
            // 
            this.lbVersioneApp.Label = "label2";
            this.lbVersioneApp.Name = "lbVersioneApp";
            // 
            // lbCoreV
            // 
            this.lbCoreV.Label = "label3";
            this.lbCoreV.Name = "lbCoreV";
            // 
            // RiMoSTRibbon
            // 
            this.Name = "RiMoSTRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.StartFromScratch = true;
            this.Tabs.Add(this.tbRichiestaModifica);
            this.Tabs.Add(this.tabVersione);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.RiMoSTRibbon_Load);
            this.tbRichiestaModifica.ResumeLayout(false);
            this.tbRichiestaModifica.PerformLayout();
            this.groupChiudi.ResumeLayout(false);
            this.groupChiudi.PerformLayout();
            this.groupAzioni.ResumeLayout(false);
            this.groupAzioni.PerformLayout();
            this.groupModifica.ResumeLayout(false);
            this.groupModifica.PerformLayout();
            this.groupGestione.ResumeLayout(false);
            this.groupGestione.PerformLayout();
            this.tabVersione.ResumeLayout(false);
            this.tabVersione.PerformLayout();
            this.groupVersione.ResumeLayout(false);
            this.groupVersione.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tbRichiestaModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupAzioni;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupGestione;
        internal Microsoft.Office.Tools.Ribbon.RibbonTab tabVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnChiudi;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnInvia;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton chkIsDraft;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReset;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnRefresh;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnPrint;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbAnniDisponibili;
        internal Microsoft.Office.Tools.Ribbon.RibbonComboBox cbAnniDisponibili;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAnnulla;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnModifica;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupVersione;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbVersioneLabel;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbVersioneApp;
        internal Microsoft.Office.Tools.Ribbon.RibbonLabel lbCoreV;
    }

    partial class ThisRibbonCollection
    {
        internal RiMoSTRibbon RiMoSTRibbon
        {
            get { return this.GetRibbon<RiMoSTRibbon>(); }
        }
    }
}
