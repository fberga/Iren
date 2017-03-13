namespace Iren.PSO.Forms
{
    partial class FormBilanciamento
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
            this.dgvBilanciamento = new System.Windows.Forms.DataGridView();
            this.lbUPPrinc = new System.Windows.Forms.Label();
            this.lbZonaUP = new System.Windows.Forms.Label();
            this.panelCentrale = new System.Windows.Forms.Panel();
            this.lbZonaUPprinc = new System.Windows.Forms.Label();
            this.lbSiglaEntitaPrinc = new System.Windows.Forms.Label();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelError = new System.Windows.Forms.Label();
            this.btnModifica = new System.Windows.Forms.Button();
            this.btnRipristina = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dgvBilanciamento)).BeginInit();
            this.panelCentrale.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dgvBilanciamento
            // 
            this.dgvBilanciamento.AllowUserToAddRows = false;
            this.dgvBilanciamento.AllowUserToDeleteRows = false;
            this.dgvBilanciamento.AllowUserToResizeRows = false;
            this.dgvBilanciamento.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvBilanciamento.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dgvBilanciamento.Location = new System.Drawing.Point(0, 39);
            this.dgvBilanciamento.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dgvBilanciamento.Name = "dgvBilanciamento";
            this.dgvBilanciamento.Size = new System.Drawing.Size(979, 552);
            this.dgvBilanciamento.TabIndex = 0;
            this.dgvBilanciamento.CellEnter += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBilanciamento_CellEnter);
            this.dgvBilanciamento.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvBilanciamento_CellValueChanged);
            this.dgvBilanciamento.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgvBilanciamento_CurrentCellDirtyStateChanged);
            // 
            // lbUPPrinc
            // 
            this.lbUPPrinc.AutoSize = true;
            this.lbUPPrinc.Location = new System.Drawing.Point(3, 3);
            this.lbUPPrinc.Name = "lbUPPrinc";
            this.lbUPPrinc.Size = new System.Drawing.Size(93, 20);
            this.lbUPPrinc.TabIndex = 3;
            this.lbUPPrinc.Text = "UP principale:";
            // 
            // lbZonaUP
            // 
            this.lbZonaUP.AutoSize = true;
            this.lbZonaUP.Location = new System.Drawing.Point(352, 3);
            this.lbZonaUP.Name = "lbZonaUP";
            this.lbZonaUP.Size = new System.Drawing.Size(42, 20);
            this.lbZonaUP.TabIndex = 5;
            this.lbZonaUP.Text = "Zona:";
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.lbZonaUPprinc);
            this.panelCentrale.Controls.Add(this.lbSiglaEntitaPrinc);
            this.panelCentrale.Controls.Add(this.lbZonaUP);
            this.panelCentrale.Controls.Add(this.lbUPPrinc);
            this.panelCentrale.Controls.Add(this.dgvBilanciamento);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(3, 3);
            this.panelCentrale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Size = new System.Drawing.Size(979, 591);
            this.panelCentrale.TabIndex = 14;
            // 
            // lbZonaUPprinc
            // 
            this.lbZonaUPprinc.AutoSize = true;
            this.lbZonaUPprinc.Location = new System.Drawing.Point(400, 3);
            this.lbZonaUPprinc.Name = "lbZonaUPprinc";
            this.lbZonaUPprinc.Size = new System.Drawing.Size(45, 20);
            this.lbZonaUPprinc.TabIndex = 8;
            this.lbZonaUPprinc.Text = "label2";
            // 
            // lbSiglaEntitaPrinc
            // 
            this.lbSiglaEntitaPrinc.AutoSize = true;
            this.lbSiglaEntitaPrinc.Location = new System.Drawing.Point(130, 3);
            this.lbSiglaEntitaPrinc.Name = "lbSiglaEntitaPrinc";
            this.lbSiglaEntitaPrinc.Size = new System.Drawing.Size(45, 20);
            this.lbSiglaEntitaPrinc.TabIndex = 7;
            this.lbSiglaEntitaPrinc.Text = "label1";
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.panel1);
            this.panelButtons.Controls.Add(this.btnModifica);
            this.panelButtons.Controls.Add(this.btnRipristina);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(3, 594);
            this.panelButtons.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(979, 53);
            this.panelButtons.TabIndex = 15;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.labelError);
            this.panel1.Location = new System.Drawing.Point(3, 4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(635, 46);
            this.panel1.TabIndex = 8;
            // 
            // labelError
            // 
            this.labelError.AutoSize = true;
            this.labelError.ForeColor = System.Drawing.Color.Red;
            this.labelError.Location = new System.Drawing.Point(3, 0);
            this.labelError.MaximumSize = new System.Drawing.Size(635, 0);
            this.labelError.Name = "labelError";
            this.labelError.Size = new System.Drawing.Size(0, 20);
            this.labelError.TabIndex = 7;
            // 
            // btnModifica
            // 
            this.btnModifica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnModifica.Enabled = false;
            this.btnModifica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnModifica.Location = new System.Drawing.Point(640, 5);
            this.btnModifica.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnModifica.Name = "btnModifica";
            this.btnModifica.Size = new System.Drawing.Size(113, 48);
            this.btnModifica.TabIndex = 4;
            this.btnModifica.Text = "Modifica";
            this.btnModifica.UseVisualStyleBackColor = true;
            this.btnModifica.Click += new System.EventHandler(this.Btn_Esegui_Click);
            // 
            // btnRipristina
            // 
            this.btnRipristina.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnRipristina.Enabled = false;
            this.btnRipristina.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRipristina.Location = new System.Drawing.Point(753, 5);
            this.btnRipristina.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnRipristina.Name = "btnRipristina";
            this.btnRipristina.Size = new System.Drawing.Size(113, 48);
            this.btnRipristina.TabIndex = 6;
            this.btnRipristina.Text = "Ripristina";
            this.btnRipristina.UseVisualStyleBackColor = true;
            this.btnRipristina.Click += new System.EventHandler(this.RipristinaValori_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(866, 5);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // FormBilanciamento
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(985, 650);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Arial Narrow", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Location = new System.Drawing.Point(10, 10);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormBilanciamento";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "FormBilanciamento";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormBilanciamento_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormBilanciamento_FormClosed);
            this.Load += new System.EventHandler(this.FormBilanciamento_Load);
            this.Shown += new System.EventHandler(this.FormBilanciamento_Shown);
            ((System.ComponentModel.ISupportInitialize)(this.dgvBilanciamento)).EndInit();
            this.panelCentrale.ResumeLayout(false);
            this.panelCentrale.PerformLayout();
            this.panelButtons.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dgvBilanciamento;
        private System.Windows.Forms.Label lbUPPrinc;
        private System.Windows.Forms.Label lbZonaUP;
        private System.Windows.Forms.Panel panelCentrale;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnModifica;
        private System.Windows.Forms.Button btnRipristina;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Label lbZonaUPprinc;
        private System.Windows.Forms.Label lbSiglaEntitaPrinc;
        private System.Windows.Forms.Label labelError;
        private System.Windows.Forms.Panel panel1;
    }
}