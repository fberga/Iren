namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class AssegnaFunzioni
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
            this.treeViewNotUtilized = new System.Windows.Forms.TreeView();
            this.treeViewUtilized = new System.Windows.Forms.TreeView();
            this.btnRimuovi = new System.Windows.Forms.Button();
            this.btnAggiungi = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.toolStrip1 = new System.Windows.Forms.ToolStrip();
            this.toolStripSotto = new System.Windows.Forms.ToolStripButton();
            this.toolStripSopra = new System.Windows.Forms.ToolStripButton();
            this.panelBottom.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            this.toolStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeViewNotUtilized
            // 
            this.treeViewNotUtilized.Location = new System.Drawing.Point(0, 45);
            this.treeViewNotUtilized.Margin = new System.Windows.Forms.Padding(0, 0, 0, 3);
            this.treeViewNotUtilized.Name = "treeViewNotUtilized";
            this.tableLayoutPanel1.SetRowSpan(this.treeViewNotUtilized, 2);
            this.treeViewNotUtilized.Size = new System.Drawing.Size(282, 453);
            this.treeViewNotUtilized.TabIndex = 0;
            // 
            // treeViewUtilized
            // 
            this.treeViewUtilized.Location = new System.Drawing.Point(318, 45);
            this.treeViewUtilized.Margin = new System.Windows.Forms.Padding(0, 0, 0, 3);
            this.treeViewUtilized.Name = "treeViewUtilized";
            this.tableLayoutPanel1.SetRowSpan(this.treeViewUtilized, 2);
            this.treeViewUtilized.Size = new System.Drawing.Size(282, 453);
            this.treeViewUtilized.TabIndex = 3;
            // 
            // btnRimuovi
            // 
            this.btnRimuovi.BackgroundImage = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.leftArrow;
            this.btnRimuovi.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnRimuovi.FlatAppearance.BorderSize = 0;
            this.btnRimuovi.Location = new System.Drawing.Point(286, 277);
            this.btnRimuovi.Margin = new System.Windows.Forms.Padding(4);
            this.btnRimuovi.Name = "btnRimuovi";
            this.btnRimuovi.Size = new System.Drawing.Size(28, 30);
            this.btnRimuovi.TabIndex = 2;
            this.btnRimuovi.UseVisualStyleBackColor = true;
            this.btnRimuovi.Click += new System.EventHandler(this.RimuoviFunzione_Click);
            // 
            // btnAggiungi
            // 
            this.btnAggiungi.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAggiungi.BackgroundImage = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.rightArrow;
            this.btnAggiungi.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.btnAggiungi.FlatAppearance.BorderSize = 0;
            this.btnAggiungi.Location = new System.Drawing.Point(286, 239);
            this.btnAggiungi.Margin = new System.Windows.Forms.Padding(4);
            this.btnAggiungi.Name = "btnAggiungi";
            this.btnAggiungi.Size = new System.Drawing.Size(28, 30);
            this.btnAggiungi.TabIndex = 1;
            this.btnAggiungi.UseVisualStyleBackColor = true;
            this.btnAggiungi.Click += new System.EventHandler(this.AggiungiFunzione_Click);
            // 
            // label1
            // 
            this.label1.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(4, 2);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(121, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Funzioni disponibili";
            // 
            // label2
            // 
            this.label2.Anchor = System.Windows.Forms.AnchorStyles.Left;
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(322, 2);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(124, 16);
            this.label2.TabIndex = 7;
            this.label2.Text = "Funzioni assegnate";
            // 
            // panelBottom
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panelBottom, 3);
            this.panelBottom.Controls.Add(this.btnOk);
            this.panelBottom.Controls.Add(this.btnAnnulla);
            this.panelBottom.Location = new System.Drawing.Point(0, 501);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(600, 48);
            this.panelBottom.TabIndex = 15;
            // 
            // btnOk
            // 
            this.btnOk.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(374, 0);
            this.btnOk.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(113, 48);
            this.btnOk.TabIndex = 4;
            this.btnOk.Text = "Ok";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.AssegnaFunzioni_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(487, 0);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.AnnullaCambiamenti_Click);
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 3;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 36F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.label2, 2, 0);
            this.tableLayoutPanel1.Controls.Add(this.panelBottom, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.btnAggiungi, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.treeViewUtilized, 2, 2);
            this.tableLayoutPanel1.Controls.Add(this.treeViewNotUtilized, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.btnRimuovi, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.toolStrip1, 2, 1);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 25F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(600, 550);
            this.tableLayoutPanel1.TabIndex = 8;
            // 
            // toolStrip1
            // 
            this.toolStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripSotto,
            this.toolStripSopra});
            this.toolStrip1.Location = new System.Drawing.Point(318, 20);
            this.toolStrip1.Name = "toolStrip1";
            this.toolStrip1.Size = new System.Drawing.Size(282, 25);
            this.toolStrip1.TabIndex = 16;
            this.toolStrip1.Text = "toolStrip1";
            // 
            // toolStripSotto
            // 
            this.toolStripSotto.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.downArrow;
            this.toolStripSotto.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSotto.Name = "toolStripSotto";
            this.toolStripSotto.Size = new System.Drawing.Size(92, 22);
            this.toolStripSotto.Text = "Sposta sotto";
            this.toolStripSotto.Click += new System.EventHandler(this.SpostaSotto_Click);
            // 
            // toolStripSopra
            // 
            this.toolStripSopra.Image = global::Iren.ToolsExcel.ConfiguratoreRibbon.Properties.Resources.upArrow;
            this.toolStripSopra.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripSopra.Name = "toolStripSopra";
            this.toolStripSopra.Size = new System.Drawing.Size(94, 22);
            this.toolStripSopra.Text = "Sposta sopra";
            this.toolStripSopra.Click += new System.EventHandler(this.SpostaSopra_Click);
            // 
            // AssegnaFunzioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(688, 609);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "AssegnaFunzioni";
            this.Text = "AssegnaFunzioni";
            this.panelBottom.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.toolStrip1.ResumeLayout(false);
            this.toolStrip1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TreeView treeViewNotUtilized;
        private System.Windows.Forms.TreeView treeViewUtilized;
        private System.Windows.Forms.Button btnRimuovi;
        private System.Windows.Forms.Button btnAggiungi;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ToolStrip toolStrip1;
        private System.Windows.Forms.ToolStripButton toolStripSotto;
        private System.Windows.Forms.ToolStripButton toolStripSopra;
    }
}