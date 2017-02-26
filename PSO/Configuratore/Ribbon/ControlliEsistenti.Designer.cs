namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class ControlliEsistenti
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
            this.tableLayoutPanelForm = new System.Windows.Forms.TableLayoutPanel();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnAggiungi = new System.Windows.Forms.Button();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.treeViewControlli = new System.Windows.Forms.TreeView();
            this.panelTopContent = new System.Windows.Forms.Panel();
            this.lbLabel = new System.Windows.Forms.Label();
            this.txtLabel = new System.Windows.Forms.TextBox();
            this.panelDimensione = new System.Windows.Forms.Panel();
            this.radioDimLarge = new System.Windows.Forms.RadioButton();
            this.radioDimSmall = new System.Windows.Forms.RadioButton();
            this.lbDimension = new System.Windows.Forms.Label();
            this.imgButton = new System.Windows.Forms.PictureBox();
            this.lbScreenTip = new System.Windows.Forms.Label();
            this.lbDesc = new System.Windows.Forms.Label();
            this.txtScreenTip = new System.Windows.Forms.TextBox();
            this.txtDesc = new System.Windows.Forms.TextBox();
            this.lbGruppi = new System.Windows.Forms.Label();
            this.lbApplicazioni = new System.Windows.Forms.Label();
            this.lbFunzioni = new System.Windows.Forms.Label();
            this.listBoxGruppi = new System.Windows.Forms.ListBox();
            this.listBoxApplicazioni = new System.Windows.Forms.ListBox();
            this.listBoxFunzioni = new System.Windows.Forms.ListBox();
            this.tableLayoutPanelForm.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.panelTopContent.SuspendLayout();
            this.panelDimensione.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.imgButton)).BeginInit();
            this.SuspendLayout();
            // 
            // tableLayoutPanelForm
            // 
            this.tableLayoutPanelForm.ColumnCount = 4;
            this.tableLayoutPanelForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanelForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanelForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanelForm.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 25F));
            this.tableLayoutPanelForm.Controls.Add(this.panelBottom, 0, 3);
            this.tableLayoutPanelForm.Controls.Add(this.treeViewControlli, 0, 0);
            this.tableLayoutPanelForm.Controls.Add(this.panelTopContent, 1, 0);
            this.tableLayoutPanelForm.Controls.Add(this.lbGruppi, 1, 1);
            this.tableLayoutPanelForm.Controls.Add(this.lbApplicazioni, 2, 1);
            this.tableLayoutPanelForm.Controls.Add(this.lbFunzioni, 3, 1);
            this.tableLayoutPanelForm.Controls.Add(this.listBoxGruppi, 1, 2);
            this.tableLayoutPanelForm.Controls.Add(this.listBoxApplicazioni, 2, 2);
            this.tableLayoutPanelForm.Controls.Add(this.listBoxFunzioni, 3, 2);
            this.tableLayoutPanelForm.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanelForm.Name = "tableLayoutPanelForm";
            this.tableLayoutPanelForm.RowCount = 4;
            this.tableLayoutPanelForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 214F));
            this.tableLayoutPanelForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 18F));
            this.tableLayoutPanelForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanelForm.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.tableLayoutPanelForm.Size = new System.Drawing.Size(968, 556);
            this.tableLayoutPanelForm.TabIndex = 0;
            // 
            // panelBottom
            // 
            this.tableLayoutPanelForm.SetColumnSpan(this.panelBottom, 3);
            this.panelBottom.Controls.Add(this.btnAggiungi);
            this.panelBottom.Controls.Add(this.btnChiudi);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBottom.Location = new System.Drawing.Point(245, 505);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(720, 48);
            this.panelBottom.TabIndex = 14;
            // 
            // btnAggiungi
            // 
            this.btnAggiungi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAggiungi.Location = new System.Drawing.Point(494, 0);
            this.btnAggiungi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAggiungi.Name = "btnAggiungi";
            this.btnAggiungi.Size = new System.Drawing.Size(113, 48);
            this.btnAggiungi.TabIndex = 4;
            this.btnAggiungi.Text = "Aggiungi";
            this.btnAggiungi.UseVisualStyleBackColor = true;
            this.btnAggiungi.Click += new System.EventHandler(this.AggiungiControllo_Click);
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
            this.btnChiudi.TabIndex = 5;
            this.btnChiudi.Text = "Chiudi";
            this.btnChiudi.UseVisualStyleBackColor = true;
            // 
            // treeViewControlli
            // 
            this.treeViewControlli.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewControlli.Location = new System.Drawing.Point(3, 3);
            this.treeViewControlli.Name = "treeViewControlli";
            this.tableLayoutPanelForm.SetRowSpan(this.treeViewControlli, 4);
            this.treeViewControlli.Size = new System.Drawing.Size(236, 550);
            this.treeViewControlli.TabIndex = 0;
            this.treeViewControlli.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.AfterSelectNode);
            // 
            // panelTopContent
            // 
            this.tableLayoutPanelForm.SetColumnSpan(this.panelTopContent, 3);
            this.panelTopContent.Controls.Add(this.lbLabel);
            this.panelTopContent.Controls.Add(this.txtLabel);
            this.panelTopContent.Controls.Add(this.panelDimensione);
            this.panelTopContent.Controls.Add(this.lbDimension);
            this.panelTopContent.Controls.Add(this.imgButton);
            this.panelTopContent.Controls.Add(this.lbScreenTip);
            this.panelTopContent.Controls.Add(this.lbDesc);
            this.panelTopContent.Controls.Add(this.txtScreenTip);
            this.panelTopContent.Controls.Add(this.txtDesc);
            this.panelTopContent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTopContent.Location = new System.Drawing.Point(245, 3);
            this.panelTopContent.Name = "panelTopContent";
            this.panelTopContent.Size = new System.Drawing.Size(720, 208);
            this.panelTopContent.TabIndex = 1;
            // 
            // lbLabel
            // 
            this.lbLabel.AutoSize = true;
            this.lbLabel.Location = new System.Drawing.Point(93, 5);
            this.lbLabel.Name = "lbLabel";
            this.lbLabel.Size = new System.Drawing.Size(42, 16);
            this.lbLabel.TabIndex = 19;
            this.lbLabel.Text = "Label";
            // 
            // txtLabel
            // 
            this.txtLabel.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtLabel.Location = new System.Drawing.Point(141, 3);
            this.txtLabel.Name = "txtLabel";
            this.txtLabel.ReadOnly = true;
            this.txtLabel.Size = new System.Drawing.Size(579, 22);
            this.txtLabel.TabIndex = 18;
            this.txtLabel.TabStop = false;
            // 
            // panelDimensione
            // 
            this.panelDimensione.BackColor = System.Drawing.SystemColors.Control;
            this.panelDimensione.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelDimensione.Controls.Add(this.radioDimLarge);
            this.panelDimensione.Controls.Add(this.radioDimSmall);
            this.panelDimensione.Location = new System.Drawing.Point(141, 163);
            this.panelDimensione.Name = "panelDimensione";
            this.panelDimensione.Size = new System.Drawing.Size(170, 45);
            this.panelDimensione.TabIndex = 17;
            // 
            // radioDimLarge
            // 
            this.radioDimLarge.AutoSize = true;
            this.radioDimLarge.Checked = true;
            this.radioDimLarge.Enabled = false;
            this.radioDimLarge.Location = new System.Drawing.Point(2, 22);
            this.radioDimLarge.Name = "radioDimLarge";
            this.radioDimLarge.Size = new System.Drawing.Size(71, 20);
            this.radioDimLarge.TabIndex = 0;
            this.radioDimLarge.Text = "Grande";
            this.radioDimLarge.UseVisualStyleBackColor = true;
            // 
            // radioDimSmall
            // 
            this.radioDimSmall.AutoSize = true;
            this.radioDimSmall.Enabled = false;
            this.radioDimSmall.Location = new System.Drawing.Point(2, 2);
            this.radioDimSmall.Name = "radioDimSmall";
            this.radioDimSmall.Size = new System.Drawing.Size(71, 20);
            this.radioDimSmall.TabIndex = 1;
            this.radioDimSmall.Text = "Piccolo";
            this.radioDimSmall.UseVisualStyleBackColor = true;
            // 
            // lbDimension
            // 
            this.lbDimension.AutoSize = true;
            this.lbDimension.Location = new System.Drawing.Point(55, 168);
            this.lbDimension.Name = "lbDimension";
            this.lbDimension.Size = new System.Drawing.Size(80, 16);
            this.lbDimension.TabIndex = 16;
            this.lbDimension.Text = "Dimensione";
            // 
            // imgButton
            // 
            this.imgButton.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.imgButton.Location = new System.Drawing.Point(0, 3);
            this.imgButton.Name = "imgButton";
            this.imgButton.Size = new System.Drawing.Size(50, 50);
            this.imgButton.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.imgButton.TabIndex = 15;
            this.imgButton.TabStop = false;
            // 
            // lbScreenTip
            // 
            this.lbScreenTip.AutoSize = true;
            this.lbScreenTip.Location = new System.Drawing.Point(61, 99);
            this.lbScreenTip.Name = "lbScreenTip";
            this.lbScreenTip.Size = new System.Drawing.Size(74, 16);
            this.lbScreenTip.TabIndex = 14;
            this.lbScreenTip.Text = "Screen Tip";
            // 
            // lbDesc
            // 
            this.lbDesc.AutoSize = true;
            this.lbDesc.Location = new System.Drawing.Point(56, 33);
            this.lbDesc.Name = "lbDesc";
            this.lbDesc.Size = new System.Drawing.Size(79, 16);
            this.lbDesc.TabIndex = 13;
            this.lbDesc.Text = "Descrizione";
            // 
            // txtScreenTip
            // 
            this.txtScreenTip.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtScreenTip.Location = new System.Drawing.Point(141, 97);
            this.txtScreenTip.Multiline = true;
            this.txtScreenTip.Name = "txtScreenTip";
            this.txtScreenTip.ReadOnly = true;
            this.txtScreenTip.Size = new System.Drawing.Size(579, 60);
            this.txtScreenTip.TabIndex = 12;
            this.txtScreenTip.TabStop = false;
            // 
            // txtDesc
            // 
            this.txtDesc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.txtDesc.Location = new System.Drawing.Point(141, 31);
            this.txtDesc.Multiline = true;
            this.txtDesc.Name = "txtDesc";
            this.txtDesc.ReadOnly = true;
            this.txtDesc.Size = new System.Drawing.Size(579, 60);
            this.txtDesc.TabIndex = 11;
            this.txtDesc.TabStop = false;
            // 
            // lbGruppi
            // 
            this.lbGruppi.AutoSize = true;
            this.lbGruppi.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbGruppi.Location = new System.Drawing.Point(245, 216);
            this.lbGruppi.Name = "lbGruppi";
            this.lbGruppi.Size = new System.Drawing.Size(236, 16);
            this.lbGruppi.TabIndex = 15;
            this.lbGruppi.Text = "Gruppi";
            // 
            // lbApplicazioni
            // 
            this.lbApplicazioni.AutoSize = true;
            this.lbApplicazioni.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbApplicazioni.Location = new System.Drawing.Point(487, 216);
            this.lbApplicazioni.Name = "lbApplicazioni";
            this.lbApplicazioni.Size = new System.Drawing.Size(236, 16);
            this.lbApplicazioni.TabIndex = 16;
            this.lbApplicazioni.Text = "Applicazioni";
            // 
            // lbFunzioni
            // 
            this.lbFunzioni.AutoSize = true;
            this.lbFunzioni.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbFunzioni.Location = new System.Drawing.Point(729, 216);
            this.lbFunzioni.Name = "lbFunzioni";
            this.lbFunzioni.Size = new System.Drawing.Size(236, 16);
            this.lbFunzioni.TabIndex = 17;
            this.lbFunzioni.Text = "Funzioni";
            // 
            // listBoxGruppi
            // 
            this.listBoxGruppi.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxGruppi.FormattingEnabled = true;
            this.listBoxGruppi.ItemHeight = 16;
            this.listBoxGruppi.Location = new System.Drawing.Point(245, 235);
            this.listBoxGruppi.Name = "listBoxGruppi";
            this.listBoxGruppi.Size = new System.Drawing.Size(236, 264);
            this.listBoxGruppi.TabIndex = 1;
            this.listBoxGruppi.SelectedIndexChanged += new System.EventHandler(this.SelectedGroupChanged);
            // 
            // listBoxApplicazioni
            // 
            this.listBoxApplicazioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxApplicazioni.FormattingEnabled = true;
            this.listBoxApplicazioni.ItemHeight = 16;
            this.listBoxApplicazioni.Location = new System.Drawing.Point(487, 235);
            this.listBoxApplicazioni.Name = "listBoxApplicazioni";
            this.listBoxApplicazioni.Size = new System.Drawing.Size(236, 264);
            this.listBoxApplicazioni.TabIndex = 2;
            this.listBoxApplicazioni.SelectedIndexChanged += new System.EventHandler(this.SelectedApplicationChanged);
            // 
            // listBoxFunzioni
            // 
            this.listBoxFunzioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxFunzioni.FormattingEnabled = true;
            this.listBoxFunzioni.ItemHeight = 16;
            this.listBoxFunzioni.Location = new System.Drawing.Point(729, 235);
            this.listBoxFunzioni.Name = "listBoxFunzioni";
            this.listBoxFunzioni.Size = new System.Drawing.Size(236, 264);
            this.listBoxFunzioni.TabIndex = 3;
            // 
            // ControlliEsistenti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(1045, 625);
            this.Controls.Add(this.tableLayoutPanelForm);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "ControlliEsistenti";
            this.Text = "Controlli Esistenti";
            this.tableLayoutPanelForm.ResumeLayout(false);
            this.tableLayoutPanelForm.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.panelTopContent.ResumeLayout(false);
            this.panelTopContent.PerformLayout();
            this.panelDimensione.ResumeLayout(false);
            this.panelDimensione.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.imgButton)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanelForm;
        private System.Windows.Forms.TreeView treeViewControlli;
        private System.Windows.Forms.Panel panelTopContent;
        private System.Windows.Forms.Label lbScreenTip;
        private System.Windows.Forms.Label lbDesc;
        private System.Windows.Forms.TextBox txtScreenTip;
        private System.Windows.Forms.TextBox txtDesc;
        private System.Windows.Forms.PictureBox imgButton;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnAggiungi;
        private System.Windows.Forms.Button btnChiudi;
        private System.Windows.Forms.Label lbGruppi;
        private System.Windows.Forms.Label lbApplicazioni;
        private System.Windows.Forms.Label lbFunzioni;
        private System.Windows.Forms.ListBox listBoxGruppi;
        private System.Windows.Forms.ListBox listBoxApplicazioni;
        private System.Windows.Forms.ListBox listBoxFunzioni;
        private System.Windows.Forms.Panel panelDimensione;
        private System.Windows.Forms.RadioButton radioDimLarge;
        private System.Windows.Forms.RadioButton radioDimSmall;
        private System.Windows.Forms.Label lbDimension;
        private System.Windows.Forms.Label lbLabel;
        private System.Windows.Forms.TextBox txtLabel;

    }
}