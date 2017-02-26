namespace XMLMerge
{
    partial class XMLMerge
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
            this.scegliCartella = new System.Windows.Forms.FolderBrowserDialog();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.txtPercorso = new System.Windows.Forms.TextBox();
            this.lbPercorso = new System.Windows.Forms.Label();
            this.btnApri = new System.Windows.Forms.Button();
            this.btnEsegui = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtOutput
            // 
            this.txtOutput.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.txtOutput.Location = new System.Drawing.Point(0, 84);
            this.txtOutput.Multiline = true;
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.txtOutput.Size = new System.Drawing.Size(601, 313);
            this.txtOutput.TabIndex = 0;
            // 
            // txtPercorso
            // 
            this.txtPercorso.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtPercorso.Location = new System.Drawing.Point(90, 6);
            this.txtPercorso.Name = "txtPercorso";
            this.txtPercorso.Size = new System.Drawing.Size(474, 26);
            this.txtPercorso.TabIndex = 1;
            // 
            // lbPercorso
            // 
            this.lbPercorso.AutoSize = true;
            this.lbPercorso.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbPercorso.Location = new System.Drawing.Point(12, 9);
            this.lbPercorso.Name = "lbPercorso";
            this.lbPercorso.Size = new System.Drawing.Size(72, 20);
            this.lbPercorso.TabIndex = 2;
            this.lbPercorso.Text = "Percorso";
            // 
            // btnApri
            // 
            this.btnApri.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApri.Location = new System.Drawing.Point(563, 6);
            this.btnApri.Name = "btnApri";
            this.btnApri.Size = new System.Drawing.Size(26, 26);
            this.btnApri.TabIndex = 3;
            this.btnApri.Text = "...";
            this.btnApri.UseVisualStyleBackColor = true;
            this.btnApri.Click += new System.EventHandler(this.btnApri_Click);
            // 
            // btnEsegui
            // 
            this.btnEsegui.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEsegui.Location = new System.Drawing.Point(504, 38);
            this.btnEsegui.Name = "btnEsegui";
            this.btnEsegui.Size = new System.Drawing.Size(85, 40);
            this.btnEsegui.TabIndex = 4;
            this.btnEsegui.Text = "Esegui";
            this.btnEsegui.UseVisualStyleBackColor = true;
            this.btnEsegui.Click += new System.EventHandler(this.btnEsegui_Click);
            // 
            // XMLMerge
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(601, 397);
            this.Controls.Add(this.btnEsegui);
            this.Controls.Add(this.btnApri);
            this.Controls.Add(this.lbPercorso);
            this.Controls.Add(this.txtPercorso);
            this.Controls.Add(this.txtOutput);
            this.Name = "XMLMerge";
            this.Text = "XML Merge";
            this.Shown += new System.EventHandler(this.XMLMerge_Shown);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.FolderBrowserDialog scegliCartella;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.TextBox txtPercorso;
        private System.Windows.Forms.Label lbPercorso;
        private System.Windows.Forms.Button btnApri;
        private System.Windows.Forms.Button btnEsegui;
    }
}

