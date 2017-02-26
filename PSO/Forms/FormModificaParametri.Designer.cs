namespace Iren.PSO.Forms
{
    partial class FormModificaParametri
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormModificaParametri));
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.cmbEntita = new System.Windows.Forms.ComboBox();
            this.contextMenuPar = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.modificareValoreContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaParametroContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.inserisciSopraContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.inserisciSottoContextMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.panelTop = new System.Windows.Forms.Panel();
            this.labelSelPar = new System.Windows.Forms.Label();
            this.labelSelEntita = new System.Windows.Forms.Label();
            this.cmbParametri = new System.Windows.Forms.ComboBox();
            this.panelTopMenu = new System.Windows.Forms.Panel();
            this.menuPar = new System.Windows.Forms.ToolStrip();
            this.elimiaTopMenu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.inserisciSopraTopMenu = new System.Windows.Forms.ToolStripButton();
            this.inserisciSottoTopMenu = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.modificaTopMenu = new System.Windows.Forms.ToolStripButton();
            this.panelMiddle = new System.Windows.Forms.Panel();
            this.dataGridParametri = new System.Windows.Forms.DataGridView();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.contextMenuPar.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.panelTopMenu.SuspendLayout();
            this.menuPar.SuspendLayout();
            this.panelMiddle.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametri)).BeginInit();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbEntita
            // 
            this.cmbEntita.FormattingEnabled = true;
            this.cmbEntita.Location = new System.Drawing.Point(179, 1);
            this.cmbEntita.Name = "cmbEntita";
            this.cmbEntita.Size = new System.Drawing.Size(518, 28);
            this.cmbEntita.TabIndex = 0;
            this.cmbEntita.SelectedIndexChanged += new System.EventHandler(this.CambiaEntita);
            // 
            // contextMenuPar
            // 
            this.contextMenuPar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.modificareValoreContextMenu,
            this.cancellaParametroContextMenu,
            this.inserisciSopraContextMenu,
            this.inserisciSottoContextMenu});
            this.contextMenuPar.Name = "contextMenuDataGrid";
            this.contextMenuPar.Size = new System.Drawing.Size(233, 114);
            // 
            // modificareValoreContextMenu
            // 
            this.modificareValoreContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("modificareValoreContextMenu.Image")));
            this.modificareValoreContextMenu.Name = "modificareValoreContextMenu";
            this.modificareValoreContextMenu.Size = new System.Drawing.Size(232, 22);
            this.modificareValoreContextMenu.Text = "Modificare valore";
            this.modificareValoreContextMenu.ToolTipText = "Modifica il valore del parametro se il parametro non è ancora attivo";
            this.modificareValoreContextMenu.Click += new System.EventHandler(this.ModificaParametro);
            // 
            // cancellaParametroContextMenu
            // 
            this.cancellaParametroContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("cancellaParametroContextMenu.Image")));
            this.cancellaParametroContextMenu.Name = "cancellaParametroContextMenu";
            this.cancellaParametroContextMenu.Size = new System.Drawing.Size(232, 22);
            this.cancellaParametroContextMenu.Text = "Elimina parametro";
            this.cancellaParametroContextMenu.ToolTipText = "Cancella il parametro se il parametro non è ancora attivo";
            this.cancellaParametroContextMenu.Click += new System.EventHandler(this.CancellaParametro);
            // 
            // inserisciSopraContextMenu
            // 
            this.inserisciSopraContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSopraContextMenu.Image")));
            this.inserisciSopraContextMenu.Name = "inserisciSopraContextMenu";
            this.inserisciSopraContextMenu.Size = new System.Drawing.Size(232, 22);
            this.inserisciSopraContextMenu.Text = "Inserisci sopra riga selezionata";
            this.inserisciSopraContextMenu.Click += new System.EventHandler(this.InserisciSopra);
            // 
            // inserisciSottoContextMenu
            // 
            this.inserisciSottoContextMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSottoContextMenu.Image")));
            this.inserisciSottoContextMenu.Name = "inserisciSottoContextMenu";
            this.inserisciSottoContextMenu.Size = new System.Drawing.Size(232, 22);
            this.inserisciSottoContextMenu.Text = "Inserisci sotto riga selezionata";
            this.inserisciSottoContextMenu.Click += new System.EventHandler(this.InserisciSotto);
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.labelSelPar);
            this.panelTop.Controls.Add(this.labelSelEntita);
            this.panelTop.Controls.Add(this.cmbParametri);
            this.panelTop.Controls.Add(this.cmbEntita);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(5, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(700, 62);
            this.panelTop.TabIndex = 2;
            // 
            // labelSelPar
            // 
            this.labelSelPar.AutoSize = true;
            this.labelSelPar.Location = new System.Drawing.Point(3, 35);
            this.labelSelPar.Name = "labelSelPar";
            this.labelSelPar.Size = new System.Drawing.Size(170, 20);
            this.labelSelPar.TabIndex = 6;
            this.labelSelPar.Text = "Seleziona il parametro:";
            // 
            // labelSelEntita
            // 
            this.labelSelEntita.AutoSize = true;
            this.labelSelEntita.Location = new System.Drawing.Point(3, 4);
            this.labelSelEntita.Name = "labelSelEntita";
            this.labelSelEntita.Size = new System.Drawing.Size(115, 20);
            this.labelSelEntita.TabIndex = 1;
            this.labelSelEntita.Text = "Seleziona l\'UP:";
            // 
            // cmbParametri
            // 
            this.cmbParametri.FormattingEnabled = true;
            this.cmbParametri.Location = new System.Drawing.Point(179, 32);
            this.cmbParametri.Name = "cmbParametri";
            this.cmbParametri.Size = new System.Drawing.Size(518, 28);
            this.cmbParametri.TabIndex = 5;
            this.cmbParametri.SelectedIndexChanged += new System.EventHandler(this.CambiaParametro);
            // 
            // panelTopMenu
            // 
            this.panelTopMenu.Controls.Add(this.menuPar);
            this.panelTopMenu.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTopMenu.Location = new System.Drawing.Point(0, 0);
            this.panelTopMenu.Name = "panelTopMenu";
            this.panelTopMenu.Size = new System.Drawing.Size(700, 37);
            this.panelTopMenu.TabIndex = 11;
            // 
            // menuPar
            // 
            this.menuPar.BackColor = System.Drawing.SystemColors.Control;
            this.menuPar.Dock = System.Windows.Forms.DockStyle.Fill;
            this.menuPar.GripMargin = new System.Windows.Forms.Padding(3);
            this.menuPar.GripStyle = System.Windows.Forms.ToolStripGripStyle.Hidden;
            this.menuPar.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.elimiaTopMenu,
            this.toolStripSeparator1,
            this.inserisciSopraTopMenu,
            this.inserisciSottoTopMenu,
            this.toolStripSeparator2,
            this.modificaTopMenu});
            this.menuPar.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow;
            this.menuPar.Location = new System.Drawing.Point(0, 0);
            this.menuPar.Name = "menuPar";
            this.menuPar.Padding = new System.Windows.Forms.Padding(3);
            this.menuPar.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
            this.menuPar.Size = new System.Drawing.Size(700, 37);
            this.menuPar.TabIndex = 10;
            this.menuPar.Text = "Strumenti Parametri Giornalieri";
            // 
            // elimiaTopMenu
            // 
            this.elimiaTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("elimiaTopMenu.Image")));
            this.elimiaTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.elimiaTopMenu.Name = "elimiaTopMenu";
            this.elimiaTopMenu.Size = new System.Drawing.Size(66, 28);
            this.elimiaTopMenu.Text = "Elimina";
            this.elimiaTopMenu.ToolTipText = "Elimina la riga selezionata";
            this.elimiaTopMenu.Click += new System.EventHandler(this.CancellaParametro);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 31);
            // 
            // inserisciSopraTopMenu
            // 
            this.inserisciSopraTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSopraTopMenu.Image")));
            this.inserisciSopraTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.inserisciSopraTopMenu.Name = "inserisciSopraTopMenu";
            this.inserisciSopraTopMenu.Size = new System.Drawing.Size(101, 28);
            this.inserisciSopraTopMenu.Text = "Inserisci sopra";
            this.inserisciSopraTopMenu.ToolTipText = "Inserisci sopra la riga corrente";
            this.inserisciSopraTopMenu.Click += new System.EventHandler(this.InserisciSopra);
            // 
            // inserisciSottoTopMenu
            // 
            this.inserisciSottoTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("inserisciSottoTopMenu.Image")));
            this.inserisciSottoTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.inserisciSottoTopMenu.Name = "inserisciSottoTopMenu";
            this.inserisciSottoTopMenu.Size = new System.Drawing.Size(99, 28);
            this.inserisciSottoTopMenu.Text = "Inserisci sotto";
            this.inserisciSottoTopMenu.ToolTipText = "Inserisci sotto la riga corrente";
            this.inserisciSottoTopMenu.Click += new System.EventHandler(this.InserisciSotto);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 31);
            // 
            // modificaTopMenu
            // 
            this.modificaTopMenu.Image = ((System.Drawing.Image)(resources.GetObject("modificaTopMenu.Image")));
            this.modificaTopMenu.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.modificaTopMenu.Name = "modificaTopMenu";
            this.modificaTopMenu.Size = new System.Drawing.Size(74, 28);
            this.modificaTopMenu.Text = "Modifica";
            this.modificaTopMenu.ToolTipText = "Modifica la riga selezionata";
            this.modificaTopMenu.Click += new System.EventHandler(this.ModificaParametro);
            // 
            // panelMiddle
            // 
            this.panelMiddle.Controls.Add(this.dataGridParametri);
            this.panelMiddle.Controls.Add(this.panelTopMenu);
            this.panelMiddle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelMiddle.Location = new System.Drawing.Point(5, 67);
            this.panelMiddle.Name = "panelMiddle";
            this.panelMiddle.Size = new System.Drawing.Size(700, 344);
            this.panelMiddle.TabIndex = 2;
            // 
            // dataGridParametri
            // 
            this.dataGridParametri.AllowUserToAddRows = false;
            this.dataGridParametri.AllowUserToDeleteRows = false;
            this.dataGridParametri.AllowUserToResizeColumns = false;
            this.dataGridParametri.AllowUserToResizeRows = false;
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(216)))), ((int)(((byte)(251)))), ((int)(((byte)(252)))));
            this.dataGridParametri.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dataGridParametri.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dataGridParametri.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridParametri.ContextMenuStrip = this.contextMenuPar;
            this.dataGridParametri.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridParametri.EditMode = System.Windows.Forms.DataGridViewEditMode.EditProgrammatically;
            this.dataGridParametri.Location = new System.Drawing.Point(0, 37);
            this.dataGridParametri.MultiSelect = false;
            this.dataGridParametri.Name = "dataGridParametri";
            this.dataGridParametri.Size = new System.Drawing.Size(700, 307);
            this.dataGridParametri.TabIndex = 6;
            this.dataGridParametri.VirtualMode = true;
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.btnAnnulla);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(5, 411);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelBottom.Size = new System.Drawing.Size(700, 53);
            this.panelBottom.TabIndex = 12;
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(587, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 6;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.FormParametriClose);
            // 
            // FormModificaParametri
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 469);
            this.Controls.Add(this.panelMiddle);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.panelTop);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormModificaParametri";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.contextMenuPar.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.panelTopMenu.ResumeLayout(false);
            this.panelTopMenu.PerformLayout();
            this.menuPar.ResumeLayout(false);
            this.menuPar.PerformLayout();
            this.panelMiddle.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridParametri)).EndInit();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbEntita;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.ContextMenuStrip contextMenuPar;
        private System.Windows.Forms.ToolStripMenuItem modificareValoreContextMenu;
        private System.Windows.Forms.ToolStripMenuItem cancellaParametroContextMenu;
        private System.Windows.Forms.Label labelSelEntita;
        private System.Windows.Forms.ToolStripMenuItem inserisciSopraContextMenu;
        private System.Windows.Forms.ToolStripMenuItem inserisciSottoContextMenu;
        private System.Windows.Forms.Panel panelTopMenu;
        private System.Windows.Forms.Label labelSelPar;
        private System.Windows.Forms.ComboBox cmbParametri;
        private System.Windows.Forms.ToolStrip menuPar;
        private System.Windows.Forms.ToolStripButton elimiaTopMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripButton inserisciSopraTopMenu;
        private System.Windows.Forms.ToolStripButton inserisciSottoTopMenu;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton modificaTopMenu;
        private System.Windows.Forms.Panel panelMiddle;
        private System.Windows.Forms.DataGridView dataGridParametri;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnAnnulla;
    }
}