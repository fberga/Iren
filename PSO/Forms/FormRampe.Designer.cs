namespace Iren.PSO.Forms
{
    partial class FormRampe
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.lbDesEntita = new System.Windows.Forms.Label();
            this.tableLayoutDesRampa = new System.Windows.Forms.TableLayoutPanel();
            this.panelValoriRampa = new System.Windows.Forms.Panel();
            this.panelLabelUP = new System.Windows.Forms.Panel();
            this.panelContent = new System.Windows.Forms.Panel();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.tableLayoutRampe = new System.Windows.Forms.TableLayoutPanel();
            this.panelLabelUP.SuspendLayout();
            this.panelContent.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(908, 5);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 48);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Applica";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(1021, 5);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // lbDesEntita
            // 
            this.lbDesEntita.AutoSize = true;
            this.lbDesEntita.Font = new System.Drawing.Font("Verdana", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbDesEntita.Location = new System.Drawing.Point(3, 9);
            this.lbDesEntita.Name = "lbDesEntita";
            this.lbDesEntita.Size = new System.Drawing.Size(107, 18);
            this.lbDesEntita.TabIndex = 6;
            this.lbDesEntita.Text = "lbDesEntita";
            // 
            // tableLayoutDesRampa
            // 
            this.tableLayoutDesRampa.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutDesRampa.ColumnCount = 3;
            this.tableLayoutDesRampa.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 60F));
            this.tableLayoutDesRampa.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutDesRampa.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 20F));
            this.tableLayoutDesRampa.Dock = System.Windows.Forms.DockStyle.Left;
            this.tableLayoutDesRampa.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutDesRampa.Name = "tableLayoutDesRampa";
            this.tableLayoutDesRampa.RowCount = 2;
            this.tableLayoutDesRampa.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutDesRampa.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutDesRampa.Size = new System.Drawing.Size(250, 195);
            this.tableLayoutDesRampa.TabIndex = 7;
            // 
            // panelValoriRampa
            // 
            this.panelValoriRampa.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelValoriRampa.Location = new System.Drawing.Point(250, 0);
            this.panelValoriRampa.Name = "panelValoriRampa";
            this.panelValoriRampa.Size = new System.Drawing.Size(884, 195);
            this.panelValoriRampa.TabIndex = 1;
            // 
            // panelLabelUP
            // 
            this.panelLabelUP.Controls.Add(this.lbDesEntita);
            this.panelLabelUP.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelLabelUP.Location = new System.Drawing.Point(5, 5);
            this.panelLabelUP.Name = "panelLabelUP";
            this.panelLabelUP.Size = new System.Drawing.Size(1134, 36);
            this.panelLabelUP.TabIndex = 10;
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.panelValoriRampa);
            this.panelContent.Controls.Add(this.tableLayoutDesRampa);
            this.panelContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelContent.Location = new System.Drawing.Point(5, 41);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(1134, 195);
            this.panelContent.TabIndex = 10;
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(5, 441);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(1134, 53);
            this.panelButtons.TabIndex = 11;
            // 
            // panel1
            // 
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(5, 236);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1134, 10);
            this.panel1.TabIndex = 12;
            // 
            // tableLayoutRampe
            // 
            this.tableLayoutRampe.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.Single;
            this.tableLayoutRampe.ColumnCount = 2;
            this.tableLayoutRampe.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutRampe.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutRampe.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.tableLayoutRampe.Location = new System.Drawing.Point(5, 246);
            this.tableLayoutRampe.Name = "tableLayoutRampe";
            this.tableLayoutRampe.RowCount = 2;
            this.tableLayoutRampe.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutRampe.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutRampe.Size = new System.Drawing.Size(1134, 195);
            this.tableLayoutRampe.TabIndex = 6;
            // 
            // frmRAMPE
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(1144, 499);
            this.Controls.Add(this.panelContent);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.tableLayoutRampe);
            this.Controls.Add(this.panelButtons);
            this.Controls.Add(this.panelLabelUP);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Name = "frmRAMPE";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Rampe";
            this.Load += new System.EventHandler(this.frmRAMPE_Load);
            this.panelLabelUP.ResumeLayout(false);
            this.panelLabelUP.PerformLayout();
            this.panelContent.ResumeLayout(false);
            this.panelButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Label lbDesEntita;
        private System.Windows.Forms.TableLayoutPanel tableLayoutDesRampa;
        private System.Windows.Forms.Panel panelValoriRampa;
        private System.Windows.Forms.Panel panelLabelUP;
        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutRampe;


    }
}