namespace GeneraXls
{
    partial class FrmGeneraXls
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
            this.panelContent = new System.Windows.Forms.Panel();
            this.lblXmlVilla = new System.Windows.Forms.Label();
            this.lblXmlTelessio = new System.Windows.Forms.Label();
            this.lblXmlOrco = new System.Windows.Forms.Label();
            this.lblXmlRosone = new System.Windows.Forms.Label();
            this.btnSfogliaPathOutput = new System.Windows.Forms.Button();
            this.btnSfogliaPathInput = new System.Windows.Forms.Button();
            this.lbPathOutput = new System.Windows.Forms.Label();
            this.lbPathInput = new System.Windows.Forms.Label();
            this.txtPathOutput = new System.Windows.Forms.TextBox();
            this.txtPathInput = new System.Windows.Forms.TextBox();
            this.lbMercato = new System.Windows.Forms.Label();
            this.lbData = new System.Windows.Forms.Label();
            this.cmbMercato = new System.Windows.Forms.ComboBox();
            this.dtpData = new System.Windows.Forms.DateTimePicker();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnGenera = new System.Windows.Forms.Button();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.folderBrowserDialogOutput = new System.Windows.Forms.FolderBrowserDialog();
            this.folderBrowserDialogInput = new System.Windows.Forms.FolderBrowserDialog();
            this.panelContent.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.lblXmlVilla);
            this.panelContent.Controls.Add(this.lblXmlTelessio);
            this.panelContent.Controls.Add(this.lblXmlOrco);
            this.panelContent.Controls.Add(this.lblXmlRosone);
            this.panelContent.Controls.Add(this.btnSfogliaPathOutput);
            this.panelContent.Controls.Add(this.btnSfogliaPathInput);
            this.panelContent.Controls.Add(this.lbPathOutput);
            this.panelContent.Controls.Add(this.lbPathInput);
            this.panelContent.Controls.Add(this.txtPathOutput);
            this.panelContent.Controls.Add(this.txtPathInput);
            this.panelContent.Controls.Add(this.lbMercato);
            this.panelContent.Controls.Add(this.lbData);
            this.panelContent.Controls.Add(this.cmbMercato);
            this.panelContent.Controls.Add(this.dtpData);
            this.panelContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelContent.Location = new System.Drawing.Point(4, 5);
            this.panelContent.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(668, 223);
            this.panelContent.TabIndex = 14;
            // 
            // lblXmlVilla
            // 
            this.lblXmlVilla.AutoSize = true;
            this.lblXmlVilla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblXmlVilla.Location = new System.Drawing.Point(105, 132);
            this.lblXmlVilla.Name = "lblXmlVilla";
            this.lblXmlVilla.Size = new System.Drawing.Size(78, 20);
            this.lblXmlVilla.TabIndex = 13;
            this.lblXmlVilla.Text = "Xml Villa";
            // 
            // lblXmlTelessio
            // 
            this.lblXmlTelessio.AutoSize = true;
            this.lblXmlTelessio.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblXmlTelessio.Location = new System.Drawing.Point(105, 112);
            this.lblXmlTelessio.Name = "lblXmlTelessio";
            this.lblXmlTelessio.Size = new System.Drawing.Size(110, 20);
            this.lblXmlTelessio.TabIndex = 12;
            this.lblXmlTelessio.Text = "Xml Telessio";
            // 
            // lblXmlOrco
            // 
            this.lblXmlOrco.AutoSize = true;
            this.lblXmlOrco.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblXmlOrco.Location = new System.Drawing.Point(105, 92);
            this.lblXmlOrco.Name = "lblXmlOrco";
            this.lblXmlOrco.Size = new System.Drawing.Size(82, 20);
            this.lblXmlOrco.TabIndex = 11;
            this.lblXmlOrco.Text = "Xml Orco";
            // 
            // lblXmlRosone
            // 
            this.lblXmlRosone.AutoSize = true;
            this.lblXmlRosone.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblXmlRosone.Location = new System.Drawing.Point(105, 72);
            this.lblXmlRosone.Name = "lblXmlRosone";
            this.lblXmlRosone.Size = new System.Drawing.Size(106, 20);
            this.lblXmlRosone.TabIndex = 10;
            this.lblXmlRosone.Text = "Xml Rosone";
            // 
            // btnSfogliaPathOutput
            // 
            this.btnSfogliaPathOutput.Location = new System.Drawing.Point(550, 171);
            this.btnSfogliaPathOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSfogliaPathOutput.Name = "btnSfogliaPathOutput";
            this.btnSfogliaPathOutput.Size = new System.Drawing.Size(112, 35);
            this.btnSfogliaPathOutput.TabIndex = 9;
            this.btnSfogliaPathOutput.Text = "Sfoglia";
            this.btnSfogliaPathOutput.UseVisualStyleBackColor = true;
            this.btnSfogliaPathOutput.Click += new System.EventHandler(this.btnSfogliaPathOutput_Click);
            // 
            // btnSfogliaPathInput
            // 
            this.btnSfogliaPathInput.Location = new System.Drawing.Point(550, 37);
            this.btnSfogliaPathInput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSfogliaPathInput.Name = "btnSfogliaPathInput";
            this.btnSfogliaPathInput.Size = new System.Drawing.Size(112, 35);
            this.btnSfogliaPathInput.TabIndex = 8;
            this.btnSfogliaPathInput.Text = "Sfoglia";
            this.btnSfogliaPathInput.UseVisualStyleBackColor = true;
            this.btnSfogliaPathInput.Click += new System.EventHandler(this.btnSfogliaPathInput_Click);
            // 
            // lbPathOutput
            // 
            this.lbPathOutput.AutoSize = true;
            this.lbPathOutput.Location = new System.Drawing.Point(4, 186);
            this.lbPathOutput.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbPathOutput.Name = "lbPathOutput";
            this.lbPathOutput.Size = new System.Drawing.Size(95, 20);
            this.lbPathOutput.TabIndex = 7;
            this.lbPathOutput.Text = "Path Output";
            // 
            // lbPathInput
            // 
            this.lbPathInput.AutoSize = true;
            this.lbPathInput.Location = new System.Drawing.Point(4, 44);
            this.lbPathInput.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbPathInput.Name = "lbPathInput";
            this.lbPathInput.Size = new System.Drawing.Size(83, 20);
            this.lbPathInput.TabIndex = 6;
            this.lbPathInput.Text = "Path Input";
            // 
            // txtPathOutput
            // 
            this.txtPathOutput.Location = new System.Drawing.Point(109, 180);
            this.txtPathOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPathOutput.Name = "txtPathOutput";
            this.txtPathOutput.Size = new System.Drawing.Size(433, 26);
            this.txtPathOutput.TabIndex = 5;
            // 
            // txtPathInput
            // 
            this.txtPathInput.Location = new System.Drawing.Point(109, 41);
            this.txtPathInput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPathInput.Name = "txtPathInput";
            this.txtPathInput.Size = new System.Drawing.Size(433, 26);
            this.txtPathInput.TabIndex = 4;
            // 
            // lbMercato
            // 
            this.lbMercato.AutoSize = true;
            this.lbMercato.Location = new System.Drawing.Point(392, 10);
            this.lbMercato.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbMercato.Name = "lbMercato";
            this.lbMercato.Size = new System.Drawing.Size(67, 20);
            this.lbMercato.TabIndex = 3;
            this.lbMercato.Text = "Mercato";
            // 
            // lbData
            // 
            this.lbData.AutoSize = true;
            this.lbData.Location = new System.Drawing.Point(4, 10);
            this.lbData.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbData.Name = "lbData";
            this.lbData.Size = new System.Drawing.Size(44, 20);
            this.lbData.TabIndex = 2;
            this.lbData.Text = "Data";
            // 
            // cmbMercato
            // 
            this.cmbMercato.FormattingEnabled = true;
            this.cmbMercato.Location = new System.Drawing.Point(482, 7);
            this.cmbMercato.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.cmbMercato.Name = "cmbMercato";
            this.cmbMercato.Size = new System.Drawing.Size(180, 28);
            this.cmbMercato.TabIndex = 1;
            this.cmbMercato.SelectedIndexChanged += new System.EventHandler(this.cmbMercato_SelectedIndexChanged);
            // 
            // dtpData
            // 
            this.dtpData.Location = new System.Drawing.Point(109, 5);
            this.dtpData.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.dtpData.Name = "dtpData";
            this.dtpData.Size = new System.Drawing.Size(261, 26);
            this.dtpData.TabIndex = 0;
            this.dtpData.ValueChanged += new System.EventHandler(this.dtpData_ValueChanged);
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnGenera);
            this.panelButtons.Controls.Add(this.btnChiudi);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(4, 228);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(668, 53);
            this.panelButtons.TabIndex = 15;
            // 
            // btnGenera
            // 
            this.btnGenera.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnGenera.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenera.Location = new System.Drawing.Point(442, 3);
            this.btnGenera.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnGenera.Name = "btnGenera";
            this.btnGenera.Size = new System.Drawing.Size(113, 50);
            this.btnGenera.TabIndex = 4;
            this.btnGenera.Text = "Genera";
            this.btnGenera.UseVisualStyleBackColor = true;
            this.btnGenera.Click += new System.EventHandler(this.btnGenera_Click);
            // 
            // btnChiudi
            // 
            this.btnChiudi.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnChiudi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnChiudi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChiudi.Location = new System.Drawing.Point(555, 3);
            this.btnChiudi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(113, 50);
            this.btnChiudi.TabIndex = 8;
            this.btnChiudi.Text = "Chiudi";
            this.btnChiudi.UseVisualStyleBackColor = true;
            this.btnChiudi.Click += new System.EventHandler(this.btnChiudi_Click);
            // 
            // FrmGeneraXls
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(676, 286);
            this.Controls.Add(this.panelContent);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FrmGeneraXls";
            this.Padding = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Text = "FrmGeneraXls";
            this.panelContent.ResumeLayout(false);
            this.panelContent.PerformLayout();
            this.panelButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.Button btnSfogliaPathOutput;
        private System.Windows.Forms.Button btnSfogliaPathInput;
        private System.Windows.Forms.Label lbPathOutput;
        private System.Windows.Forms.Label lbPathInput;
        private System.Windows.Forms.TextBox txtPathOutput;
        private System.Windows.Forms.TextBox txtPathInput;
        private System.Windows.Forms.Label lbMercato;
        private System.Windows.Forms.Label lbData;
        private System.Windows.Forms.ComboBox cmbMercato;
        private System.Windows.Forms.DateTimePicker dtpData;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnGenera;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogOutput;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialogInput;
        private System.Windows.Forms.Label lblXmlVilla;
        private System.Windows.Forms.Label lblXmlTelessio;
        private System.Windows.Forms.Label lblXmlOrco;
        private System.Windows.Forms.Label lblXmlRosone;
        private System.Windows.Forms.Button btnChiudi;

    }
}

