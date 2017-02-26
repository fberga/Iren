namespace GeneraXls
{
    partial class LoadForm
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
            this.panelTop = new System.Windows.Forms.Panel();
            this.lbData = new System.Windows.Forms.Label();
            this.cmbMercato = new System.Windows.Forms.ComboBox();
            this.lbMercato = new System.Windows.Forms.Label();
            this.dtpData = new System.Windows.Forms.DateTimePicker();
            this.panelContent = new System.Windows.Forms.Panel();
            this.dataGridCentrali = new System.Windows.Forms.DataGridView();
            this.lbPathInput = new System.Windows.Forms.Label();
            this.btnSfogliaPathOutput = new System.Windows.Forms.Button();
            this.lbPathOutput = new System.Windows.Forms.Label();
            this.txtPathOutput = new System.Windows.Forms.TextBox();
            this.btnSfogliaPathInput = new System.Windows.Forms.Button();
            this.txtPathInput = new System.Windows.Forms.TextBox();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnGenera = new System.Windows.Forms.Button();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.chooseFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.panelTop.SuspendLayout();
            this.panelContent.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridCentrali)).BeginInit();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.lbData);
            this.panelTop.Controls.Add(this.cmbMercato);
            this.panelTop.Controls.Add(this.lbMercato);
            this.panelTop.Controls.Add(this.dtpData);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(3, 3);
            this.panelTop.Margin = new System.Windows.Forms.Padding(4);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(563, 53);
            this.panelTop.TabIndex = 0;
            // 
            // lbData
            // 
            this.lbData.AutoSize = true;
            this.lbData.Location = new System.Drawing.Point(48, 13);
            this.lbData.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lbData.Name = "lbData";
            this.lbData.Size = new System.Drawing.Size(42, 18);
            this.lbData.TabIndex = 6;
            this.lbData.Text = "Data";
            // 
            // cmbMercato
            // 
            this.cmbMercato.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMercato.FormattingEnabled = true;
            this.cmbMercato.Location = new System.Drawing.Point(388, 10);
            this.cmbMercato.Margin = new System.Windows.Forms.Padding(6, 7, 6, 7);
            this.cmbMercato.Name = "cmbMercato";
            this.cmbMercato.Size = new System.Drawing.Size(165, 26);
            this.cmbMercato.TabIndex = 5;
            // 
            // lbMercato
            // 
            this.lbMercato.AutoSize = true;
            this.lbMercato.Location = new System.Drawing.Point(311, 13);
            this.lbMercato.Margin = new System.Windows.Forms.Padding(6, 0, 6, 0);
            this.lbMercato.Name = "lbMercato";
            this.lbMercato.Size = new System.Drawing.Size(65, 18);
            this.lbMercato.TabIndex = 7;
            this.lbMercato.Text = "Mercato";
            // 
            // dtpData
            // 
            this.dtpData.Format = System.Windows.Forms.DateTimePickerFormat.Short;
            this.dtpData.Location = new System.Drawing.Point(102, 10);
            this.dtpData.Margin = new System.Windows.Forms.Padding(6, 7, 6, 7);
            this.dtpData.Name = "dtpData";
            this.dtpData.Size = new System.Drawing.Size(165, 26);
            this.dtpData.TabIndex = 4;
            // 
            // panelContent
            // 
            this.panelContent.Controls.Add(this.dataGridCentrali);
            this.panelContent.Controls.Add(this.lbPathInput);
            this.panelContent.Controls.Add(this.btnSfogliaPathOutput);
            this.panelContent.Controls.Add(this.lbPathOutput);
            this.panelContent.Controls.Add(this.txtPathOutput);
            this.panelContent.Controls.Add(this.btnSfogliaPathInput);
            this.panelContent.Controls.Add(this.txtPathInput);
            this.panelContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelContent.Location = new System.Drawing.Point(3, 56);
            this.panelContent.Name = "panelContent";
            this.panelContent.Size = new System.Drawing.Size(563, 271);
            this.panelContent.TabIndex = 1;
            // 
            // dataGridCentrali
            // 
            this.dataGridCentrali.AllowUserToAddRows = false;
            this.dataGridCentrali.AllowUserToDeleteRows = false;
            this.dataGridCentrali.AllowUserToResizeColumns = false;
            this.dataGridCentrali.AllowUserToResizeRows = false;
            this.dataGridCentrali.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridCentrali.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.dataGridCentrali.Location = new System.Drawing.Point(0, 86);
            this.dataGridCentrali.Name = "dataGridCentrali";
            this.dataGridCentrali.ReadOnly = true;
            this.dataGridCentrali.Size = new System.Drawing.Size(563, 185);
            this.dataGridCentrali.TabIndex = 21;
            // 
            // lbPathInput
            // 
            this.lbPathInput.AutoSize = true;
            this.lbPathInput.Location = new System.Drawing.Point(18, 14);
            this.lbPathInput.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbPathInput.Name = "lbPathInput";
            this.lbPathInput.Size = new System.Drawing.Size(76, 18);
            this.lbPathInput.TabIndex = 16;
            this.lbPathInput.Text = "Path Input";
            // 
            // btnSfogliaPathOutput
            // 
            this.btnSfogliaPathOutput.Location = new System.Drawing.Point(485, 46);
            this.btnSfogliaPathOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSfogliaPathOutput.Name = "btnSfogliaPathOutput";
            this.btnSfogliaPathOutput.Size = new System.Drawing.Size(68, 26);
            this.btnSfogliaPathOutput.TabIndex = 20;
            this.btnSfogliaPathOutput.Text = "Sfoglia";
            this.btnSfogliaPathOutput.UseVisualStyleBackColor = true;
            // 
            // lbPathOutput
            // 
            this.lbPathOutput.AutoSize = true;
            this.lbPathOutput.Location = new System.Drawing.Point(5, 50);
            this.lbPathOutput.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.lbPathOutput.Name = "lbPathOutput";
            this.lbPathOutput.Size = new System.Drawing.Size(89, 18);
            this.lbPathOutput.TabIndex = 19;
            this.lbPathOutput.Text = "Path Output";
            // 
            // txtPathOutput
            // 
            this.txtPathOutput.Location = new System.Drawing.Point(102, 46);
            this.txtPathOutput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPathOutput.Name = "txtPathOutput";
            this.txtPathOutput.Size = new System.Drawing.Size(385, 26);
            this.txtPathOutput.TabIndex = 18;
            // 
            // btnSfogliaPathInput
            // 
            this.btnSfogliaPathInput.Location = new System.Drawing.Point(485, 10);
            this.btnSfogliaPathInput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnSfogliaPathInput.Name = "btnSfogliaPathInput";
            this.btnSfogliaPathInput.Size = new System.Drawing.Size(68, 26);
            this.btnSfogliaPathInput.TabIndex = 17;
            this.btnSfogliaPathInput.Text = "Sfoglia";
            this.btnSfogliaPathInput.UseVisualStyleBackColor = true;
            // 
            // txtPathInput
            // 
            this.txtPathInput.Location = new System.Drawing.Point(102, 10);
            this.txtPathInput.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPathInput.Name = "txtPathInput";
            this.txtPathInput.Size = new System.Drawing.Size(385, 26);
            this.txtPathInput.TabIndex = 15;
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.btnGenera);
            this.panelBottom.Controls.Add(this.btnChiudi);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelBottom.Location = new System.Drawing.Point(3, 327);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelBottom.Size = new System.Drawing.Size(563, 53);
            this.panelBottom.TabIndex = 21;
            // 
            // btnGenera
            // 
            this.btnGenera.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnGenera.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnGenera.Location = new System.Drawing.Point(337, 3);
            this.btnGenera.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnGenera.Name = "btnGenera";
            this.btnGenera.Size = new System.Drawing.Size(113, 50);
            this.btnGenera.TabIndex = 4;
            this.btnGenera.Text = "Genera";
            this.btnGenera.UseVisualStyleBackColor = true;
            // 
            // btnChiudi
            // 
            this.btnChiudi.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnChiudi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnChiudi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChiudi.Location = new System.Drawing.Point(450, 3);
            this.btnChiudi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(113, 50);
            this.btnChiudi.TabIndex = 8;
            this.btnChiudi.Text = "Chiudi";
            this.btnChiudi.UseVisualStyleBackColor = true;
            this.btnChiudi.Click += new System.EventHandler(this.btnChiudi_Click);
            // 
            // LoadForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(569, 383);
            this.Controls.Add(this.panelContent);
            this.Controls.Add(this.panelBottom);
            this.Controls.Add(this.panelTop);
            this.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "LoadForm";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.Text = "Genera XLS";
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.panelContent.ResumeLayout(false);
            this.panelContent.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridCentrali)).EndInit();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Label lbMercato;
        private System.Windows.Forms.Label lbData;
        private System.Windows.Forms.ComboBox cmbMercato;
        private System.Windows.Forms.DateTimePicker dtpData;
        private System.Windows.Forms.Panel panelContent;
        private System.Windows.Forms.Label lbPathInput;
        private System.Windows.Forms.Button btnSfogliaPathOutput;
        private System.Windows.Forms.Label lbPathOutput;
        private System.Windows.Forms.TextBox txtPathOutput;
        private System.Windows.Forms.Button btnSfogliaPathInput;
        private System.Windows.Forms.TextBox txtPathInput;
        private System.Windows.Forms.DataGridView dataGridCentrali;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnGenera;
        private System.Windows.Forms.Button btnChiudi;
        private System.Windows.Forms.FolderBrowserDialog chooseFolder;
    }
}