namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class GruppiEsistenti
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panelRibbonLayout = new System.Windows.Forms.Panel();
            this.listBoxGruppi = new System.Windows.Forms.ListBox();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnAggiungi = new System.Windows.Forms.Button();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.lbFunzioni = new System.Windows.Forms.Label();
            this.lbApplicazioni = new System.Windows.Forms.Label();
            this.listBoxFunzioni = new System.Windows.Forms.ListBox();
            this.listBoxApplicazioni = new System.Windows.Forms.ListBox();
            this.lbUtenti = new System.Windows.Forms.Label();
            this.listBoxUtenti = new System.Windows.Forms.ListBox();
            this.tableLayoutPanel1.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 4;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanel1.Controls.Add(this.listBoxUtenti, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.lbUtenti, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.panelRibbonLayout, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.listBoxGruppi, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.panelBottom, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.lbFunzioni, 3, 1);
            this.tableLayoutPanel1.Controls.Add(this.lbApplicazioni, 2, 1);
            this.tableLayoutPanel1.Controls.Add(this.listBoxFunzioni, 3, 2);
            this.tableLayoutPanel1.Controls.Add(this.listBoxApplicazioni, 2, 2);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 4;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 214F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(968, 586);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panelRibbonLayout
            // 
            this.panelRibbonLayout.BackColor = System.Drawing.SystemColors.ControlLight;
            this.tableLayoutPanel1.SetColumnSpan(this.panelRibbonLayout, 3);
            this.panelRibbonLayout.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelRibbonLayout.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panelRibbonLayout.Location = new System.Drawing.Point(245, 3);
            this.panelRibbonLayout.Name = "panelRibbonLayout";
            this.panelRibbonLayout.Padding = new System.Windows.Forms.Padding(2);
            this.panelRibbonLayout.Size = new System.Drawing.Size(720, 208);
            this.panelRibbonLayout.TabIndex = 4;
            // 
            // listBoxGruppi
            // 
            this.listBoxGruppi.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxGruppi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxGruppi.FormattingEnabled = true;
            this.listBoxGruppi.ItemHeight = 16;
            this.listBoxGruppi.Location = new System.Drawing.Point(3, 3);
            this.listBoxGruppi.Name = "listBoxGruppi";
            this.tableLayoutPanel1.SetRowSpan(this.listBoxGruppi, 4);
            this.listBoxGruppi.Size = new System.Drawing.Size(236, 580);
            this.listBoxGruppi.TabIndex = 0;
            this.listBoxGruppi.SelectedValueChanged += new System.EventHandler(this.CambioGruppo);
            // 
            // panelBottom
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panelBottom, 3);
            this.panelBottom.Controls.Add(this.btnAggiungi);
            this.panelBottom.Controls.Add(this.btnChiudi);
            this.panelBottom.Location = new System.Drawing.Point(245, 535);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(720, 48);
            this.panelBottom.TabIndex = 15;
            // 
            // btnAggiungi
            // 
            this.btnAggiungi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAggiungi.Location = new System.Drawing.Point(494, 0);
            this.btnAggiungi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAggiungi.Name = "btnAggiungi";
            this.btnAggiungi.Size = new System.Drawing.Size(113, 48);
            this.btnAggiungi.TabIndex = 3;
            this.btnAggiungi.Text = "Aggiungi";
            this.btnAggiungi.UseVisualStyleBackColor = true;
            this.btnAggiungi.Click += new System.EventHandler(this.AggiungiGruppo_Click);
            // 
            // btnChiudi
            // 
            this.btnChiudi.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnChiudi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnChiudi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChiudi.Location = new System.Drawing.Point(607, 0);
            this.btnChiudi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(113, 48);
            this.btnChiudi.TabIndex = 4;
            this.btnChiudi.Text = "Chiudi";
            this.btnChiudi.UseVisualStyleBackColor = true;
            // 
            // lbFunzioni
            // 
            this.lbFunzioni.AutoSize = true;
            this.lbFunzioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbFunzioni.Location = new System.Drawing.Point(729, 214);
            this.lbFunzioni.Name = "lbFunzioni";
            this.lbFunzioni.Size = new System.Drawing.Size(57, 16);
            this.lbFunzioni.TabIndex = 23;
            this.lbFunzioni.Text = "Funzioni";
            // 
            // lbApplicazioni
            // 
            this.lbApplicazioni.AutoSize = true;
            this.lbApplicazioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbApplicazioni.Location = new System.Drawing.Point(487, 214);
            this.lbApplicazioni.Name = "lbApplicazioni";
            this.lbApplicazioni.Size = new System.Drawing.Size(81, 16);
            this.lbApplicazioni.TabIndex = 21;
            this.lbApplicazioni.Text = "Applicazioni";
            // 
            // listBoxFunzioni
            // 
            this.listBoxFunzioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxFunzioni.FormattingEnabled = true;
            this.listBoxFunzioni.ItemHeight = 16;
            this.listBoxFunzioni.Location = new System.Drawing.Point(729, 235);
            this.listBoxFunzioni.Name = "listBoxFunzioni";
            this.listBoxFunzioni.Size = new System.Drawing.Size(236, 292);
            this.listBoxFunzioni.TabIndex = 2;
            // 
            // listBoxApplicazioni
            // 
            this.listBoxApplicazioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxApplicazioni.FormattingEnabled = true;
            this.listBoxApplicazioni.ItemHeight = 16;
            this.listBoxApplicazioni.Location = new System.Drawing.Point(487, 235);
            this.listBoxApplicazioni.Name = "listBoxApplicazioni";
            this.listBoxApplicazioni.Size = new System.Drawing.Size(236, 292);
            this.listBoxApplicazioni.TabIndex = 1;
            this.listBoxApplicazioni.SelectedValueChanged += new System.EventHandler(this.CambioApplicazione);
            // 
            // lbUtenti
            // 
            this.lbUtenti.AutoSize = true;
            this.lbUtenti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbUtenti.Location = new System.Drawing.Point(245, 214);
            this.lbUtenti.Name = "lbUtenti";
            this.lbUtenti.Size = new System.Drawing.Size(42, 16);
            this.lbUtenti.TabIndex = 22;
            this.lbUtenti.Text = "Utenti";
            // 
            // listBoxUtenti
            // 
            this.listBoxUtenti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.listBoxUtenti.FormattingEnabled = true;
            this.listBoxUtenti.ItemHeight = 16;
            this.listBoxUtenti.Location = new System.Drawing.Point(245, 235);
            this.listBoxUtenti.Name = "listBoxUtenti";
            this.listBoxUtenti.Size = new System.Drawing.Size(236, 292);
            this.listBoxUtenti.TabIndex = 2;
            this.listBoxUtenti.SelectedIndexChanged += new System.EventHandler(this.CambioUtente);
            // 
            // GruppiEsistenti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1047, 611);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "GruppiEsistenti";
            this.Text = "GruppiEsistenti";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListBox listBoxGruppi;
        private System.Windows.Forms.Panel panelRibbonLayout;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnAggiungi;
        private System.Windows.Forms.Button btnChiudi;
        private System.Windows.Forms.ListBox listBoxApplicazioni;
        private System.Windows.Forms.Label lbApplicazioni;
        private System.Windows.Forms.ListBox listBoxFunzioni;
        private System.Windows.Forms.Label lbFunzioni;
        private System.Windows.Forms.ListBox listBoxUtenti;
        private System.Windows.Forms.Label lbUtenti;
    }
}