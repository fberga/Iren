namespace Iren.PSO.Forms
{
    partial class FormImportXML
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
            this.generalContainer = new System.Windows.Forms.TableLayoutPanel();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnApri = new System.Windows.Forms.Button();
            this.btnImporta = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.chkTutte = new System.Windows.Forms.CheckBox();
            this.richTextInfoTop = new System.Windows.Forms.RichTextBox();
            this.tvEntitaInformazioni = new Iren.PSO.Forms.BugFixedTreeView();
            this.openFileXMLImport = new System.Windows.Forms.OpenFileDialog();
            this.generalContainer.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // generalContainer
            // 
            this.generalContainer.ColumnCount = 1;
            this.generalContainer.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.generalContainer.Controls.Add(this.panelBottom, 0, 2);
            this.generalContainer.Controls.Add(this.panelTop, 0, 0);
            this.generalContainer.Controls.Add(this.tvEntitaInformazioni, 0, 1);
            this.generalContainer.Dock = System.Windows.Forms.DockStyle.Fill;
            this.generalContainer.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.generalContainer.Location = new System.Drawing.Point(0, 0);
            this.generalContainer.Name = "generalContainer";
            this.generalContainer.RowCount = 4;
            this.generalContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 71F));
            this.generalContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 349F));
            this.generalContainer.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 54F));
            this.generalContainer.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.generalContainer.Size = new System.Drawing.Size(872, 474);
            this.generalContainer.TabIndex = 0;
            // 
            // panelBottom
            // 
            this.panelBottom.Controls.Add(this.btnApri);
            this.panelBottom.Controls.Add(this.btnImporta);
            this.panelBottom.Controls.Add(this.btnAnnulla);
            this.panelBottom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelBottom.Location = new System.Drawing.Point(3, 423);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(866, 48);
            this.panelBottom.TabIndex = 2;
            // 
            // btnApri
            // 
            this.btnApri.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApri.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApri.Location = new System.Drawing.Point(527, 0);
            this.btnApri.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApri.Name = "btnApri";
            this.btnApri.Size = new System.Drawing.Size(113, 48);
            this.btnApri.TabIndex = 8;
            this.btnApri.Text = "Apri file";
            this.btnApri.UseVisualStyleBackColor = true;
            this.btnApri.Click += new System.EventHandler(this.btnApri_Click);
            // 
            // btnImporta
            // 
            this.btnImporta.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnImporta.Enabled = false;
            this.btnImporta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnImporta.Location = new System.Drawing.Point(640, 0);
            this.btnImporta.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnImporta.Name = "btnImporta";
            this.btnImporta.Size = new System.Drawing.Size(113, 48);
            this.btnImporta.TabIndex = 6;
            this.btnImporta.Text = "Importa dati";
            this.btnImporta.UseVisualStyleBackColor = true;
            this.btnImporta.Click += new System.EventHandler(this.btnImporta_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(753, 0);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 7;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.chkTutte);
            this.panelTop.Controls.Add(this.richTextInfoTop);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTop.Location = new System.Drawing.Point(3, 3);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(866, 65);
            this.panelTop.TabIndex = 4;
            // 
            // chkTutte
            // 
            this.chkTutte.AutoSize = true;
            this.chkTutte.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkTutte.Location = new System.Drawing.Point(798, 41);
            this.chkTutte.Name = "chkTutte";
            this.chkTutte.Size = new System.Drawing.Size(65, 24);
            this.chkTutte.TabIndex = 7;
            this.chkTutte.Text = "Tutte";
            this.chkTutte.UseVisualStyleBackColor = true;
            this.chkTutte.CheckedChanged += new System.EventHandler(this.chkTutte_CheckedChanged);
            // 
            // richTextInfoTop
            // 
            this.richTextInfoTop.BackColor = System.Drawing.SystemColors.Control;
            this.richTextInfoTop.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.richTextInfoTop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.richTextInfoTop.Location = new System.Drawing.Point(0, 0);
            this.richTextInfoTop.Name = "richTextInfoTop";
            this.richTextInfoTop.Size = new System.Drawing.Size(866, 65);
            this.richTextInfoTop.TabIndex = 6;
            this.richTextInfoTop.Text = "Nessun file XML caricato...";
            // 
            // tvEntitaInformazioni
            // 
            this.tvEntitaInformazioni.CheckBoxes = true;
            this.tvEntitaInformazioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tvEntitaInformazioni.Location = new System.Drawing.Point(3, 74);
            this.tvEntitaInformazioni.Name = "tvEntitaInformazioni";
            this.tvEntitaInformazioni.Size = new System.Drawing.Size(866, 343);
            this.tvEntitaInformazioni.TabIndex = 5;
            this.tvEntitaInformazioni.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.tvEntitaInformazioni_AfterCheck);
            // 
            // openFileXMLImport
            // 
            this.openFileXMLImport.FileName = "openFileDialog1";
            this.openFileXMLImport.FileOk += new System.ComponentModel.CancelEventHandler(this.openFileXMLImport_FileOk);
            // 
            // FormImportXML
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(872, 474);
            this.Controls.Add(this.generalContainer);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "FormImportXML";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormImportXML";
            this.generalContainer.ResumeLayout(false);
            this.panelBottom.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel generalContainer;
        private System.Windows.Forms.OpenFileDialog openFileXMLImport;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnImporta;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Button btnApri;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.RichTextBox richTextInfoTop;
        private BugFixedTreeView tvEntitaInformazioni;
        private System.Windows.Forms.CheckBox chkTutte;
    }
}