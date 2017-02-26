namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class ConfiguratoreRibbon
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ConfiguratoreRibbon));
            this.toolStripTopMenu = new System.Windows.Forms.ToolStrip();
            this.toolStripSeparator4 = new System.Windows.Forms.ToolStripSeparator();
            this.Salva = new System.Windows.Forms.ToolStripButton();
            this.copiaA = new System.Windows.Forms.ToolStripButton();
            this.Ricarica = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.addGroup = new System.Windows.Forms.ToolStripDropDownButton();
            this.nuovoGruppo = new System.Windows.Forms.ToolStripMenuItem();
            this.gruppoEsistente = new System.Windows.Forms.ToolStripMenuItem();
            this.addTasto = new System.Windows.Forms.ToolStripDropDownButton();
            this.nuovoTasto = new System.Windows.Forms.ToolStripMenuItem();
            this.tastoEsistente = new System.Windows.Forms.ToolStripMenuItem();
            this.addCombo = new System.Windows.Forms.ToolStripDropDownButton();
            this.nuovoCombo = new System.Windows.Forms.ToolStripMenuItem();
            this.comboEsistente = new System.Windows.Forms.ToolStripMenuItem();
            this.addEmptyContainer = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.ctrlLeftButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlDownButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlUpButton = new System.Windows.Forms.ToolStripButton();
            this.ctrlRightButton = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator3 = new System.Windows.Forms.ToolStripSeparator();
            this.btnTest = new System.Windows.Forms.ToolStripButton();
            this.btnProd = new System.Windows.Forms.ToolStripButton();
            this.panelRibbonLayout = new System.Windows.Forms.Panel();
            this.panelApplicazione = new System.Windows.Forms.Panel();
            this.drpUtenti = new System.Windows.Forms.ComboBox();
            this.drpApplicazioni = new System.Windows.Forms.ComboBox();
            this.lbUtenti = new System.Windows.Forms.Label();
            this.lbTitoloApplicazione = new System.Windows.Forms.Label();
            this.tableLayoutForm = new System.Windows.Forms.TableLayoutPanel();
            this.toolStripTopMenu.SuspendLayout();
            this.panelApplicazione.SuspendLayout();
            this.tableLayoutForm.SuspendLayout();
            this.SuspendLayout();
            // 
            // toolStripTopMenu
            // 
            this.toolStripTopMenu.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.toolStripTopMenu.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripTopMenu.Dock = System.Windows.Forms.DockStyle.None;
            this.toolStripTopMenu.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.toolStripTopMenu.ImageScalingSize = new System.Drawing.Size(32, 32);
            this.toolStripTopMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripSeparator4,
            this.Salva,
            this.copiaA,
            this.Ricarica,
            this.toolStripSeparator2,
            this.addGroup,
            this.addTasto,
            this.addCombo,
            this.addEmptyContainer,
            this.toolStripSeparator1,
            this.ctrlLeftButton,
            this.ctrlDownButton,
            this.ctrlUpButton,
            this.ctrlRightButton,
            this.toolStripSeparator3,
            this.btnTest,
            this.btnProd});
            this.toolStripTopMenu.Location = new System.Drawing.Point(485, 0);
            this.toolStripTopMenu.Name = "toolStripTopMenu";
            this.toolStripTopMenu.Size = new System.Drawing.Size(1035, 56);
            this.toolStripTopMenu.TabIndex = 2;
            this.toolStripTopMenu.TabStop = true;
            this.toolStripTopMenu.Text = "Drop down";
            // 
            // toolStripSeparator4
            // 
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new System.Drawing.Size(6, 56);
            // 
            // Salva
            // 
            this.Salva.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.save;
            this.Salva.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Salva.Name = "Salva";
            this.Salva.Size = new System.Drawing.Size(38, 53);
            this.Salva.Text = "Salva";
            this.Salva.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.Salva.Click += new System.EventHandler(this.ApplicaConfigurazione_Click);
            // 
            // copiaA
            // 
            this.copiaA.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.copia;
            this.copiaA.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.copiaA.Name = "copiaA";
            this.copiaA.Size = new System.Drawing.Size(51, 53);
            this.copiaA.Text = "Copia a";
            this.copiaA.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.copiaA.Click += new System.EventHandler(this.CopiaConfigurazione_Click);
            // 
            // Ricarica
            // 
            this.Ricarica.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.ricarica;
            this.Ricarica.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.Ricarica.Name = "Ricarica";
            this.Ricarica.Size = new System.Drawing.Size(52, 53);
            this.Ricarica.Text = "Ricarica";
            this.Ricarica.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.Ricarica.Click += new System.EventHandler(this.RicaricaRibbon_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 56);
            // 
            // addGroup
            // 
            this.addGroup.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nuovoGruppo,
            this.gruppoEsistente});
            this.addGroup.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addGroup;
            this.addGroup.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addGroup.Name = "addGroup";
            this.addGroup.Size = new System.Drawing.Size(60, 53);
            this.addGroup.Text = "Gruppo";
            this.addGroup.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // nuovoGruppo
            // 
            this.nuovoGruppo.Name = "nuovoGruppo";
            this.nuovoGruppo.Size = new System.Drawing.Size(168, 22);
            this.nuovoGruppo.Text = "Crea nuovo";
            this.nuovoGruppo.Click += new System.EventHandler(this.AggiungiGruppo_Click);
            // 
            // gruppoEsistente
            // 
            this.gruppoEsistente.Name = "gruppoEsistente";
            this.gruppoEsistente.Size = new System.Drawing.Size(168, 22);
            this.gruppoEsistente.Text = "Scegli tra esistenti";
            this.gruppoEsistente.Click += new System.EventHandler(this.ScegliGruppoEsistente_Click);
            // 
            // addTasto
            // 
            this.addTasto.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nuovoTasto,
            this.tastoEsistente});
            this.addTasto.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addButton;
            this.addTasto.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addTasto.Name = "addTasto";
            this.addTasto.Size = new System.Drawing.Size(49, 53);
            this.addTasto.Text = "Tasto";
            this.addTasto.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // nuovoTasto
            // 
            this.nuovoTasto.Name = "nuovoTasto";
            this.nuovoTasto.Size = new System.Drawing.Size(168, 22);
            this.nuovoTasto.Text = "Crea Nuovo";
            this.nuovoTasto.Click += new System.EventHandler(this.AggiungiNuovoTasto_Click);
            // 
            // tastoEsistente
            // 
            this.tastoEsistente.Name = "tastoEsistente";
            this.tastoEsistente.Size = new System.Drawing.Size(168, 22);
            this.tastoEsistente.Text = "Scegli tra esistenti";
            this.tastoEsistente.Click += new System.EventHandler(this.ScegliTastoEsistente_Click);
            // 
            // addCombo
            // 
            this.addCombo.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.nuovoCombo,
            this.comboEsistente});
            this.addCombo.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addDropDown;
            this.addCombo.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addCombo.Name = "addCombo";
            this.addCombo.Size = new System.Drawing.Size(60, 53);
            this.addCombo.Text = "Combo";
            this.addCombo.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // nuovoCombo
            // 
            this.nuovoCombo.Name = "nuovoCombo";
            this.nuovoCombo.Size = new System.Drawing.Size(168, 22);
            this.nuovoCombo.Text = "Crea Nuovo";
            this.nuovoCombo.Click += new System.EventHandler(this.AggiungiNuovoCombo_Click);
            // 
            // comboEsistente
            // 
            this.comboEsistente.Name = "comboEsistente";
            this.comboEsistente.Size = new System.Drawing.Size(168, 22);
            this.comboEsistente.Text = "Scegli tra esistenti";
            this.comboEsistente.Click += new System.EventHandler(this.ScegliComboEsistente_Click);
            // 
            // addEmptyContainer
            // 
            this.addEmptyContainer.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.addEmptySlot;
            this.addEmptyContainer.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.addEmptyContainer.Name = "addEmptyContainer";
            this.addEmptyContainer.Size = new System.Drawing.Size(74, 53);
            this.addEmptyContainer.Text = "Contenitore";
            this.addEmptyContainer.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.addEmptyContainer.ToolTipText = "Aggiungi contenitore vuoto";
            this.addEmptyContainer.Click += new System.EventHandler(this.AggiungiContenitoreVuoto_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 56);
            // 
            // ctrlLeftButton
            // 
            this.ctrlLeftButton.AutoSize = false;
            this.ctrlLeftButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlLeftButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.leftArrow;
            this.ctrlLeftButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlLeftButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlLeftButton.Name = "ctrlLeftButton";
            this.ctrlLeftButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlLeftButton.Text = "toolStripButton2";
            this.ctrlLeftButton.ToolTipText = "Sposta il controllo a sinistra";
            this.ctrlLeftButton.Click += new System.EventHandler(this.MoveLeft_Click);
            // 
            // ctrlDownButton
            // 
            this.ctrlDownButton.AutoSize = false;
            this.ctrlDownButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlDownButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.downArrow;
            this.ctrlDownButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlDownButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlDownButton.Name = "ctrlDownButton";
            this.ctrlDownButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlDownButton.Text = "toolStripButton1";
            this.ctrlDownButton.ToolTipText = "Sposta il controllo in basso";
            this.ctrlDownButton.Click += new System.EventHandler(this.MoveDown_Click);
            // 
            // ctrlUpButton
            // 
            this.ctrlUpButton.AutoSize = false;
            this.ctrlUpButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlUpButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.upArrow;
            this.ctrlUpButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlUpButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlUpButton.Name = "ctrlUpButton";
            this.ctrlUpButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlUpButton.Text = "toolStripButton4";
            this.ctrlUpButton.ToolTipText = "Sposta il controllo in alto";
            this.ctrlUpButton.Click += new System.EventHandler(this.MoveUp_Click);
            // 
            // ctrlRightButton
            // 
            this.ctrlRightButton.AutoSize = false;
            this.ctrlRightButton.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.ctrlRightButton.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.rightArrow;
            this.ctrlRightButton.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.ctrlRightButton.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.ctrlRightButton.Name = "ctrlRightButton";
            this.ctrlRightButton.Size = new System.Drawing.Size(24, 24);
            this.ctrlRightButton.Text = "toolStripButton3";
            this.ctrlRightButton.ToolTipText = "Sposta il controllo a destra";
            this.ctrlRightButton.Click += new System.EventHandler(this.MoveRight_Click);
            // 
            // toolStripSeparator3
            // 
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new System.Drawing.Size(6, 56);
            // 
            // btnTest
            // 
            this.btnTest.CheckOnClick = true;
            this.btnTest.Image = ((System.Drawing.Image)(resources.GetObject("btnTest.Image")));
            this.btnTest.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnTest.Name = "btnTest";
            this.btnTest.Size = new System.Drawing.Size(36, 53);
            this.btnTest.Text = "Test";
            this.btnTest.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            this.btnTest.CheckedChanged += new System.EventHandler(this.ChangeAmbiente);
            // 
            // btnProd
            // 
            this.btnProd.CheckOnClick = true;
            this.btnProd.Image = ((System.Drawing.Image)(resources.GetObject("btnProd.Image")));
            this.btnProd.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.btnProd.Name = "btnProd";
            this.btnProd.Size = new System.Drawing.Size(36, 53);
            this.btnProd.Text = "Prod";
            this.btnProd.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText;
            // 
            // panelRibbonLayout
            // 
            this.panelRibbonLayout.AutoScroll = true;
            this.panelRibbonLayout.AutoScrollMargin = new System.Drawing.Size(30, 0);
            this.panelRibbonLayout.AutoSize = true;
            this.panelRibbonLayout.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tableLayoutForm.SetColumnSpan(this.panelRibbonLayout, 100);
            this.panelRibbonLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRibbonLayout.Location = new System.Drawing.Point(2, 58);
            this.panelRibbonLayout.Margin = new System.Windows.Forms.Padding(2);
            this.panelRibbonLayout.Name = "panelRibbonLayout";
            this.panelRibbonLayout.Padding = new System.Windows.Forms.Padding(2);
            this.panelRibbonLayout.Size = new System.Drawing.Size(1516, 208);
            this.panelRibbonLayout.TabIndex = 3;
            // 
            // panelApplicazione
            // 
            this.panelApplicazione.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panelApplicazione.Controls.Add(this.drpUtenti);
            this.panelApplicazione.Controls.Add(this.drpApplicazioni);
            this.panelApplicazione.Controls.Add(this.lbUtenti);
            this.panelApplicazione.Controls.Add(this.lbTitoloApplicazione);
            this.panelApplicazione.Location = new System.Drawing.Point(3, 3);
            this.panelApplicazione.Name = "panelApplicazione";
            this.panelApplicazione.Size = new System.Drawing.Size(479, 50);
            this.panelApplicazione.TabIndex = 15;
            // 
            // drpUtenti
            // 
            this.drpUtenti.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.drpUtenti.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.drpUtenti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.drpUtenti.FormattingEnabled = true;
            this.drpUtenti.Location = new System.Drawing.Point(242, 23);
            this.drpUtenti.Name = "drpUtenti";
            this.drpUtenti.Size = new System.Drawing.Size(236, 24);
            this.drpUtenti.TabIndex = 1;
            this.drpUtenti.SelectedValueChanged += new System.EventHandler(this.CambioUtente);
            // 
            // drpApplicazioni
            // 
            this.drpApplicazioni.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.drpApplicazioni.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.drpApplicazioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.drpApplicazioni.FormattingEnabled = true;
            this.drpApplicazioni.Location = new System.Drawing.Point(0, 23);
            this.drpApplicazioni.Name = "drpApplicazioni";
            this.drpApplicazioni.Size = new System.Drawing.Size(236, 24);
            this.drpApplicazioni.TabIndex = 0;
            this.drpApplicazioni.SelectedValueChanged += new System.EventHandler(this.CambioApplicazione);
            // 
            // lbUtenti
            // 
            this.lbUtenti.AutoSize = true;
            this.lbUtenti.Location = new System.Drawing.Point(246, 4);
            this.lbUtenti.Name = "lbUtenti";
            this.lbUtenti.Size = new System.Drawing.Size(47, 16);
            this.lbUtenti.TabIndex = 13;
            this.lbUtenti.Text = "Utente";
            // 
            // lbTitoloApplicazione
            // 
            this.lbTitoloApplicazione.AutoSize = true;
            this.lbTitoloApplicazione.Location = new System.Drawing.Point(3, 4);
            this.lbTitoloApplicazione.Name = "lbTitoloApplicazione";
            this.lbTitoloApplicazione.Size = new System.Drawing.Size(86, 16);
            this.lbTitoloApplicazione.TabIndex = 13;
            this.lbTitoloApplicazione.Text = "Applicazione";
            // 
            // tableLayoutForm
            // 
            this.tableLayoutForm.AutoSize = true;
            this.tableLayoutForm.ColumnCount = 2;
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 485F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 1035F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 249F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100F));
            this.tableLayoutForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 183F));
            this.tableLayoutForm.Controls.Add(this.panelRibbonLayout, 0, 1);
            this.tableLayoutForm.Controls.Add(this.toolStripTopMenu, 1, 0);
            this.tableLayoutForm.Controls.Add(this.panelApplicazione, 0, 0);
            this.tableLayoutForm.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutForm.MinimumSize = new System.Drawing.Size(1520, 316);
            this.tableLayoutForm.Name = "tableLayoutForm";
            this.tableLayoutForm.RowCount = 3;
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 56F));
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 212F));
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.tableLayoutForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutForm.Size = new System.Drawing.Size(1520, 316);
            this.tableLayoutForm.TabIndex = 16;
            // 
            // ConfiguratoreRibbon
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1360, 383);
            this.Controls.Add(this.tableLayoutForm);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ConfiguratoreRibbon";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.toolStripTopMenu.ResumeLayout(false);
            this.toolStripTopMenu.PerformLayout();
            this.panelApplicazione.ResumeLayout(false);
            this.panelApplicazione.PerformLayout();
            this.tableLayoutForm.ResumeLayout(false);
            this.tableLayoutForm.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStrip toolStripTopMenu;
        private System.Windows.Forms.Panel panelRibbonLayout;
        private System.Windows.Forms.ToolStripButton ctrlDownButton;
        private System.Windows.Forms.ToolStripButton ctrlLeftButton;
        private System.Windows.Forms.ToolStripButton ctrlRightButton;
        private System.Windows.Forms.ToolStripButton ctrlUpButton;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton addEmptyContainer;
        private System.Windows.Forms.Label lbTitoloApplicazione;
        private System.Windows.Forms.ComboBox drpApplicazioni;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.Panel panelApplicazione;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator3;
        private System.Windows.Forms.ToolStripDropDownButton addTasto;
        private System.Windows.Forms.ToolStripMenuItem nuovoTasto;
        private System.Windows.Forms.ToolStripMenuItem tastoEsistente;
        private System.Windows.Forms.TableLayoutPanel tableLayoutForm;
        private System.Windows.Forms.ToolStripDropDownButton addCombo;
        private System.Windows.Forms.ToolStripMenuItem nuovoCombo;
        private System.Windows.Forms.ToolStripMenuItem comboEsistente;
        private System.Windows.Forms.ToolStripButton Salva;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator4;
        private System.Windows.Forms.ToolStripButton Ricarica;
        private System.Windows.Forms.ToolStripDropDownButton addGroup;
        private System.Windows.Forms.ToolStripMenuItem nuovoGruppo;
        private System.Windows.Forms.ToolStripMenuItem gruppoEsistente;
        private System.Windows.Forms.ToolStripButton copiaA;
        private System.Windows.Forms.ComboBox drpUtenti;
        private System.Windows.Forms.Label lbUtenti;
        private System.Windows.Forms.ToolStripButton btnTest;
        private System.Windows.Forms.ToolStripButton btnProd;
    }
}

