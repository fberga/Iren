namespace Iren.FrontOffice.Grafici_IPEX_2014
{
    partial class MainForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.wbNaviga = new System.Windows.Forms.WebBrowser();
            this.bdFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.panel1 = new System.Windows.Forms.Panel();
            this.txtOutputXML = new System.Windows.Forms.TextBox();
            this.btnOpenXML = new System.Windows.Forms.Button();
            this.label5 = new System.Windows.Forms.Label();
            this.btnMSD = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.chkTutte = new System.Windows.Forms.CheckBox();
            this.txtIdZone = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOpen = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.txtOutput = new System.Windows.Forms.TextBox();
            this.btnSalva = new System.Windows.Forms.Button();
            this.btnLogin = new System.Windows.Forms.Button();
            this.cboZona = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dtData = new System.Windows.Forms.DateTimePicker();
            this.loadingPanel = new System.Windows.Forms.Panel();
            this.lbFileName = new System.Windows.Forms.Label();
            this.lbDownload = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.panel1.SuspendLayout();
            this.loadingPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // wbNaviga
            // 
            this.wbNaviga.Dock = System.Windows.Forms.DockStyle.Fill;
            this.wbNaviga.Location = new System.Drawing.Point(0, 95);
            this.wbNaviga.MinimumSize = new System.Drawing.Size(20, 20);
            this.wbNaviga.Name = "wbNaviga";
            this.wbNaviga.Size = new System.Drawing.Size(733, 521);
            this.wbNaviga.TabIndex = 3;
            this.wbNaviga.Url = new System.Uri("", System.UriKind.Relative);
            // 
            // bdFolder
            // 
            this.bdFolder.SelectedPath = "C:\\";
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.LightGray;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.txtOutputXML);
            this.panel1.Controls.Add(this.btnOpenXML);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.btnMSD);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.chkTutte);
            this.panel1.Controls.Add(this.txtIdZone);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.btnOpen);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.txtOutput);
            this.panel1.Controls.Add(this.btnSalva);
            this.panel1.Controls.Add(this.btnLogin);
            this.panel1.Controls.Add(this.cboZona);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.dtData);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(733, 95);
            this.panel1.TabIndex = 17;
            // 
            // txtOutputXML
            // 
            this.txtOutputXML.Enabled = false;
            this.txtOutputXML.Location = new System.Drawing.Point(77, 64);
            this.txtOutputXML.Name = "txtOutputXML";
            this.txtOutputXML.Size = new System.Drawing.Size(435, 20);
            this.txtOutputXML.TabIndex = 33;
            // 
            // btnOpenXML
            // 
            this.btnOpenXML.Location = new System.Drawing.Point(514, 64);
            this.btnOpenXML.Name = "btnOpenXML";
            this.btnOpenXML.Size = new System.Drawing.Size(26, 21);
            this.btnOpenXML.TabIndex = 32;
            this.btnOpenXML.Text = "...";
            this.btnOpenXML.UseVisualStyleBackColor = true;
            this.btnOpenXML.Click += new System.EventHandler(this.btnOpenXML_Click);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(12, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(64, 13);
            this.label5.TabIndex = 31;
            this.label5.Text = "Output XML";
            // 
            // btnMSD
            // 
            this.btnMSD.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnMSD.Location = new System.Drawing.Point(601, 13);
            this.btnMSD.Name = "btnMSD";
            this.btnMSD.Size = new System.Drawing.Size(119, 21);
            this.btnMSD.TabIndex = 29;
            this.btnMSD.Text = "Salva solo MSD";
            this.btnMSD.UseVisualStyleBackColor = true;
            this.btnMSD.Click += new System.EventHandler(this.btnMSD_Click);
            // 
            // label4
            // 
            this.label4.Location = new System.Drawing.Point(12, 16);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 13);
            this.label4.TabIndex = 28;
            this.label4.Text = "Data";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // chkTutte
            // 
            this.chkTutte.AutoSize = true;
            this.chkTutte.Location = new System.Drawing.Point(427, 16);
            this.chkTutte.Name = "chkTutte";
            this.chkTutte.Size = new System.Drawing.Size(51, 17);
            this.chkTutte.TabIndex = 27;
            this.chkTutte.Text = "Tutte";
            this.chkTutte.UseVisualStyleBackColor = true;
            // 
            // txtIdZone
            // 
            this.txtIdZone.Enabled = false;
            this.txtIdZone.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtIdZone.Location = new System.Drawing.Point(589, 12);
            this.txtIdZone.Name = "txtIdZone";
            this.txtIdZone.Size = new System.Drawing.Size(55, 20);
            this.txtIdZone.TabIndex = 26;
            this.txtIdZone.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.txtIdZone.Visible = false;
            this.txtIdZone.TextChanged += new System.EventHandler(this.txtIdZone_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(511, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(79, 13);
            this.label2.TabIndex = 25;
            this.label2.Text = "Zone Result ID";
            this.label2.Visible = false;
            // 
            // btnOpen
            // 
            this.btnOpen.Location = new System.Drawing.Point(514, 38);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(26, 21);
            this.btnOpen.TabIndex = 24;
            this.btnOpen.Text = "...";
            this.btnOpen.UseVisualStyleBackColor = true;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(55, 13);
            this.label3.TabIndex = 23;
            this.label3.Text = "Output file";
            // 
            // txtOutput
            // 
            this.txtOutput.Enabled = false;
            this.txtOutput.Location = new System.Drawing.Point(77, 38);
            this.txtOutput.Name = "txtOutput";
            this.txtOutput.Size = new System.Drawing.Size(435, 20);
            this.txtOutput.TabIndex = 22;
            this.txtOutput.TextChanged += new System.EventHandler(this.txtOutput_TextChanged);
            // 
            // btnSalva
            // 
            this.btnSalva.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSalva.Enabled = false;
            this.btnSalva.Location = new System.Drawing.Point(601, 38);
            this.btnSalva.Name = "btnSalva";
            this.btnSalva.Size = new System.Drawing.Size(119, 21);
            this.btnSalva.TabIndex = 21;
            this.btnSalva.Text = "Salva grafici e MSD";
            this.btnSalva.UseVisualStyleBackColor = true;
            this.btnSalva.Click += new System.EventHandler(this.btnSalva_Click);
            // 
            // btnLogin
            // 
            this.btnLogin.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLogin.Enabled = false;
            this.btnLogin.Location = new System.Drawing.Point(640, 12);
            this.btnLogin.Name = "btnLogin";
            this.btnLogin.Size = new System.Drawing.Size(80, 21);
            this.btnLogin.TabIndex = 20;
            this.btnLogin.Text = "Login";
            this.btnLogin.UseVisualStyleBackColor = true;
            this.btnLogin.Visible = false;
            this.btnLogin.Click += new System.EventHandler(this.btnLogin_Click);
            // 
            // cboZona
            // 
            this.cboZona.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cboZona.FormattingEnabled = true;
            this.cboZona.Items.AddRange(new object[] {
            "NORD",
            "CNOR",
            "SICI"});
            this.cboZona.Location = new System.Drawing.Point(331, 12);
            this.cboZona.Name = "cboZona";
            this.cboZona.Size = new System.Drawing.Size(90, 21);
            this.cboZona.TabIndex = 19;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(293, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(32, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Zona";
            // 
            // dtData
            // 
            this.dtData.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtData.Location = new System.Drawing.Point(77, 12);
            this.dtData.MinDate = new System.DateTime(2000, 1, 1, 0, 0, 0, 0);
            this.dtData.Name = "dtData";
            this.dtData.Size = new System.Drawing.Size(197, 20);
            this.dtData.TabIndex = 17;
            this.dtData.ValueChanged += new System.EventHandler(this.dtData_ValueChanged);
            // 
            // loadingPanel
            // 
            this.loadingPanel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.loadingPanel.Controls.Add(this.lbFileName);
            this.loadingPanel.Controls.Add(this.lbDownload);
            this.loadingPanel.Controls.Add(this.progressBar1);
            this.loadingPanel.Location = new System.Drawing.Point(272, 304);
            this.loadingPanel.Name = "loadingPanel";
            this.loadingPanel.Size = new System.Drawing.Size(420, 102);
            this.loadingPanel.TabIndex = 18;
            this.loadingPanel.Visible = false;
            // 
            // lbFileName
            // 
            this.lbFileName.AutoSize = true;
            this.lbFileName.Location = new System.Drawing.Point(9, 43);
            this.lbFileName.Name = "lbFileName";
            this.lbFileName.Size = new System.Drawing.Size(20, 13);
            this.lbFileName.TabIndex = 2;
            this.lbFileName.Text = "file";
            // 
            // lbDownload
            // 
            this.lbDownload.AutoSize = true;
            this.lbDownload.Location = new System.Drawing.Point(9, 23);
            this.lbDownload.Name = "lbDownload";
            this.lbDownload.Size = new System.Drawing.Size(137, 13);
            this.lbDownload.TabIndex = 1;
            this.lbDownload.Text = "Download dei file in corso...";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(9, 64);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(401, 26);
            this.progressBar1.TabIndex = 0;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(733, 616);
            this.Controls.Add(this.loadingPanel);
            this.Controls.Add(this.wbNaviga);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MinimumSize = new System.Drawing.Size(650, 650);
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Grafici IPEX";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.loadingPanel.ResumeLayout(false);
            this.loadingPanel.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.WebBrowser wbNaviga;
        private System.Windows.Forms.FolderBrowserDialog bdFolder;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.CheckBox chkTutte;
        private System.Windows.Forms.TextBox txtIdZone;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtOutput;
        private System.Windows.Forms.Button btnSalva;
        private System.Windows.Forms.Button btnLogin;
        private System.Windows.Forms.ComboBox cboZona;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtData;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnMSD;
        private System.Windows.Forms.Panel loadingPanel;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lbDownload;
        private System.Windows.Forms.Label lbFileName;
        private System.Windows.Forms.Button btnOpenXML;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox txtOutputXML;
    }
}

