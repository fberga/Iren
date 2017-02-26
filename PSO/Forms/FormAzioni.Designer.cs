namespace Iren.PSO.Forms
{
    partial class FormAzioni
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
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnMeteo = new System.Windows.Forms.Button();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.groupMercati = new System.Windows.Forms.GroupBox();
            this.checkTutte = new System.Windows.Forms.CheckBox();
            this.groupDate = new System.Windows.Forms.GroupBox();
            this.comboGiorni = new System.Windows.Forms.Button();
            this.panelCentrale = new System.Windows.Forms.Panel();
            this.panelUP = new System.Windows.Forms.Panel();
            this.treeViewUP = new Iren.PSO.Forms.BugFixedTreeView();
            this.panelCategorie = new System.Windows.Forms.Panel();
            this.treeViewCategorie = new Iren.PSO.Forms.BugFixedTreeView();
            this.panelAzioni = new System.Windows.Forms.Panel();
            this.treeViewAzioni = new Iren.PSO.Forms.BugFixedTreeView();
            this.panelButtons.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.groupDate.SuspendLayout();
            this.panelCentrale.SuspendLayout();
            this.panelUP.SuspendLayout();
            this.panelCategorie.SuspendLayout();
            this.panelAzioni.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnMeteo);
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(5, 503);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(1050, 53);
            this.panelButtons.TabIndex = 12;
            // 
            // btnMeteo
            // 
            this.btnMeteo.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnMeteo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnMeteo.Location = new System.Drawing.Point(711, 3);
            this.btnMeteo.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnMeteo.Name = "btnMeteo";
            this.btnMeteo.Size = new System.Drawing.Size(113, 50);
            this.btnMeteo.TabIndex = 6;
            this.btnMeteo.Text = "Meteo";
            this.btnMeteo.UseVisualStyleBackColor = true;
            this.btnMeteo.Click += new System.EventHandler(this.btnMeteo_Click);
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(824, 3);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 50);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Esegui";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(937, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.groupMercati);
            this.panelTop.Controls.Add(this.checkTutte);
            this.panelTop.Controls.Add(this.groupDate);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(5, 5);
            this.panelTop.Name = "panelTop";
            this.panelTop.Size = new System.Drawing.Size(1050, 53);
            this.panelTop.TabIndex = 13;
            // 
            // groupMercati
            // 
            this.groupMercati.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupMercati.Location = new System.Drawing.Point(350, 0);
            this.groupMercati.Name = "groupMercati";
            this.groupMercati.Size = new System.Drawing.Size(350, 53);
            this.groupMercati.TabIndex = 3;
            this.groupMercati.TabStop = false;
            this.groupMercati.Text = "Mercati";
            // 
            // checkTutte
            // 
            this.checkTutte.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.checkTutte.AutoSize = true;
            this.checkTutte.BackColor = System.Drawing.SystemColors.Control;
            this.checkTutte.CausesValidation = false;
            this.checkTutte.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkTutte.Location = new System.Drawing.Point(982, 28);
            this.checkTutte.Name = "checkTutte";
            this.checkTutte.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.checkTutte.Size = new System.Drawing.Size(65, 24);
            this.checkTutte.TabIndex = 2;
            this.checkTutte.Text = "Tutte";
            this.checkTutte.UseVisualStyleBackColor = false;
            this.checkTutte.CheckedChanged += new System.EventHandler(this.checkTutte_CheckedChanged);
            // 
            // groupDate
            // 
            this.groupDate.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupDate.Controls.Add(this.comboGiorni);
            this.groupDate.Dock = System.Windows.Forms.DockStyle.Left;
            this.groupDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupDate.Location = new System.Drawing.Point(0, 0);
            this.groupDate.Name = "groupDate";
            this.groupDate.Size = new System.Drawing.Size(350, 53);
            this.groupDate.TabIndex = 1;
            this.groupDate.TabStop = false;
            this.groupDate.Text = "Giorni";
            // 
            // comboGiorni
            // 
            this.comboGiorni.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboGiorni.ForeColor = System.Drawing.SystemColors.ControlDarkDark;
            this.comboGiorni.Location = new System.Drawing.Point(6, 19);
            this.comboGiorni.Name = "comboGiorni";
            this.comboGiorni.Size = new System.Drawing.Size(338, 28);
            this.comboGiorni.TabIndex = 0;
            this.comboGiorni.Text = "- Click per selezionare le date -";
            this.comboGiorni.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.comboGiorni.UseVisualStyleBackColor = true;
            this.comboGiorni.TextChanged += new System.EventHandler(this.comboGiorni_TextChanged);
            this.comboGiorni.Click += new System.EventHandler(this.comboGiorni_MouseClick);
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.panelUP);
            this.panelCentrale.Controls.Add(this.panelCategorie);
            this.panelCentrale.Controls.Add(this.panelAzioni);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(5, 58);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Padding = new System.Windows.Forms.Padding(0, 5, 0, 5);
            this.panelCentrale.Size = new System.Drawing.Size(1050, 445);
            this.panelCentrale.TabIndex = 14;
            // 
            // panelUP
            // 
            this.panelUP.Controls.Add(this.treeViewUP);
            this.panelUP.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelUP.Location = new System.Drawing.Point(700, 5);
            this.panelUP.Name = "panelUP";
            this.panelUP.Padding = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.panelUP.Size = new System.Drawing.Size(350, 435);
            this.panelUP.TabIndex = 6;
            // 
            // treeViewUP
            // 
            this.treeViewUP.CheckBoxes = true;
            this.treeViewUP.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewUP.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewUP.Location = new System.Drawing.Point(3, 0);
            this.treeViewUP.Name = "treeViewUP";
            this.treeViewUP.ShowNodeToolTips = true;
            this.treeViewUP.ShowPlusMinus = false;
            this.treeViewUP.ShowRootLines = false;
            this.treeViewUP.Size = new System.Drawing.Size(344, 435);
            this.treeViewUP.TabIndex = 1;
            this.treeViewUP.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeViewUP_AfterCheck);
            this.treeViewUP.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView_BeforeCollapse);
            // 
            // panelCategorie
            // 
            this.panelCategorie.Controls.Add(this.treeViewCategorie);
            this.panelCategorie.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelCategorie.Location = new System.Drawing.Point(350, 5);
            this.panelCategorie.Name = "panelCategorie";
            this.panelCategorie.Padding = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.panelCategorie.Size = new System.Drawing.Size(350, 435);
            this.panelCategorie.TabIndex = 5;
            // 
            // treeViewCategorie
            // 
            this.treeViewCategorie.CheckBoxes = true;
            this.treeViewCategorie.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewCategorie.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewCategorie.Location = new System.Drawing.Point(3, 0);
            this.treeViewCategorie.Name = "treeViewCategorie";
            this.treeViewCategorie.ShowNodeToolTips = true;
            this.treeViewCategorie.ShowPlusMinus = false;
            this.treeViewCategorie.ShowRootLines = false;
            this.treeViewCategorie.Size = new System.Drawing.Size(344, 435);
            this.treeViewCategorie.TabIndex = 1;
            this.treeViewCategorie.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterCheck);
            this.treeViewCategorie.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView_BeforeCollapse);
            // 
            // panelAzioni
            // 
            this.panelAzioni.Controls.Add(this.treeViewAzioni);
            this.panelAzioni.Dock = System.Windows.Forms.DockStyle.Left;
            this.panelAzioni.Location = new System.Drawing.Point(0, 5);
            this.panelAzioni.Name = "panelAzioni";
            this.panelAzioni.Padding = new System.Windows.Forms.Padding(3, 0, 3, 0);
            this.panelAzioni.Size = new System.Drawing.Size(350, 435);
            this.panelAzioni.TabIndex = 4;
            // 
            // treeViewAzioni
            // 
            this.treeViewAzioni.CheckBoxes = true;
            this.treeViewAzioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewAzioni.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.treeViewAzioni.Location = new System.Drawing.Point(3, 0);
            this.treeViewAzioni.Name = "treeViewAzioni";
            this.treeViewAzioni.ShowPlusMinus = false;
            this.treeViewAzioni.ShowRootLines = false;
            this.treeViewAzioni.Size = new System.Drawing.Size(344, 435);
            this.treeViewAzioni.TabIndex = 0;
            this.treeViewAzioni.AfterCheck += new System.Windows.Forms.TreeViewEventHandler(this.treeView_AfterCheck);
            this.treeViewAzioni.BeforeCollapse += new System.Windows.Forms.TreeViewCancelEventHandler(this.treeView_BeforeCollapse);
            // 
            // FormAzioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1060, 561);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelTop);
            this.Controls.Add(this.panelButtons);
            this.Name = "FormAzioni";
            this.Padding = new System.Windows.Forms.Padding(5);
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Azioni";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormAzioni_FormClosing);
            this.Load += new System.EventHandler(this.frmAZIONI_Load);
            this.panelButtons.ResumeLayout(false);
            this.panelTop.ResumeLayout(false);
            this.panelTop.PerformLayout();
            this.groupDate.ResumeLayout(false);
            this.panelCentrale.ResumeLayout(false);
            this.panelUP.ResumeLayout(false);
            this.panelCategorie.ResumeLayout(false);
            this.panelAzioni.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Panel panelTop;
        private System.Windows.Forms.Panel panelCentrale;
        private System.Windows.Forms.Panel panelUP;
        private System.Windows.Forms.Panel panelCategorie;
        private System.Windows.Forms.Panel panelAzioni;
        private BugFixedTreeView treeViewUP;
        private BugFixedTreeView treeViewCategorie;
        private BugFixedTreeView treeViewAzioni;
        private System.Windows.Forms.GroupBox groupDate;
        private System.Windows.Forms.Button btnMeteo;
        private System.Windows.Forms.CheckBox checkTutte;
        private System.Windows.Forms.Button comboGiorni;
        private System.Windows.Forms.GroupBox groupMercati;

        /******************** Modifica nuovi mercati MB  BEGIN ********************/
        /******************** Modifica nuovi mercati MB  END ********************/
    }
}