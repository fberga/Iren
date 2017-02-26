namespace Iren.PSO.Forms
{
    partial class ErrorPane
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

        #region Codice generato da Progettazione componenti

        /// <summary> 
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare 
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.panelDescrizione = new System.Windows.Forms.Panel();
            this.lbTesto = new System.Windows.Forms.Label();
            this.lbTitolo = new System.Windows.Forms.Label();
            this.panelContent = new System.Windows.Forms.Panel();
            this.treeViewErrori = new System.Windows.Forms.TreeView();
            this.panelPadding = new System.Windows.Forms.Panel();
            this.panelSeparator = new System.Windows.Forms.Panel();
            this.panelDescrizione.SuspendLayout();
            this.panelContent.SuspendLayout();
            this.panelPadding.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelDescrizione
            // 
            this.panelDescrizione.Controls.Add(this.lbTesto);
            this.panelDescrizione.Controls.Add(this.lbTitolo);
            this.panelDescrizione.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelDescrizione.Location = new System.Drawing.Point(0, 0);
            this.panelDescrizione.Name = "panelDescrizione";
            this.panelDescrizione.Padding = new System.Windows.Forms.Padding(4, 1, 2, 0);
            this.panelDescrizione.Size = new System.Drawing.Size(382, 86);
            this.panelDescrizione.TabIndex = 0;
            // 
            // lbTesto
            // 
            this.lbTesto.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lbTesto.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTesto.Location = new System.Drawing.Point(4, 19);
            this.lbTesto.Name = "lbTesto";
            this.lbTesto.Size = new System.Drawing.Size(376, 67);
            this.lbTesto.TabIndex = 3;
            this.lbTesto.Text = "Il pannello contiene la lista di errori suddivisi per UP e per giorno.";
            // 
            // lbTitolo
            // 
            this.lbTitolo.AutoSize = true;
            this.lbTitolo.Dock = System.Windows.Forms.DockStyle.Top;
            this.lbTitolo.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbTitolo.Location = new System.Drawing.Point(4, 1);
            this.lbTitolo.Margin = new System.Windows.Forms.Padding(5);
            this.lbTitolo.Name = "lbTitolo";
            this.lbTitolo.Size = new System.Drawing.Size(159, 18);
            this.lbTitolo.TabIndex = 2;
            this.lbTitolo.Text = "Pannello degli errori";
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.treeViewErrori);
            this.panelContent.Controls.Add(this.panelPadding);
            this.panelContent.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelContent.Location = new System.Drawing.Point(0, 86);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(382, 341);
            this.panelContent.TabIndex = 1;
            // 
            // treeViewErrori
            // 
            this.treeViewErrori.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.treeViewErrori.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewErrori.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewErrori.HideSelection = false;
            this.treeViewErrori.Location = new System.Drawing.Point(0, 10);
            this.treeViewErrori.Name = "treeViewErrori";
            this.treeViewErrori.Size = new System.Drawing.Size(382, 331);
            this.treeViewErrori.TabIndex = 0;
            this.treeViewErrori.BeforeSelect += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeViewErrori_BeforeSelect);
            this.treeViewErrori.NodeMouseClick += new System.Windows.Forms.TreeNodeMouseClickEventHandler(this.treeViewErrori_NodeMouseClick);
            // 
            // panelPadding
            // 
            this.panelPadding.Controls.Add(this.panelSeparator);
            this.panelPadding.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelPadding.Location = new System.Drawing.Point(0, 0);
            this.panelPadding.Name = "panelPadding";
            this.panelPadding.Padding = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.panelPadding.Size = new System.Drawing.Size(382, 10);
            this.panelPadding.TabIndex = 2;
            // 
            // panelSeparator
            // 
            this.panelSeparator.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelSeparator.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelSeparator.Location = new System.Drawing.Point(4, 0);
            this.panelSeparator.Margin = new System.Windows.Forms.Padding(0);
            this.panelSeparator.Name = "panelSeparator";
            this.panelSeparator.Size = new System.Drawing.Size(374, 1);
            this.panelSeparator.TabIndex = 1;
            // 
            // ErrorPane
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.panelContent);
            this.Controls.Add(this.panelDescrizione);
            this.Name = "ErrorPane";
            this.Size = new System.Drawing.Size(382, 575);
            this.SizeChanged += new System.EventHandler(this.ErrorPane_SizeChanged);
            this.panelDescrizione.ResumeLayout(false);
            this.panelDescrizione.PerformLayout();
            this.panelContent.ResumeLayout(false);
            this.panelPadding.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelDescrizione;
        private System.Windows.Forms.Label lbTesto;
        private System.Windows.Forms.Label lbTitolo;
        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.TreeView treeViewErrori;
        private System.Windows.Forms.Panel panelSeparator;
        private System.Windows.Forms.Panel panelPadding;


    }
}
