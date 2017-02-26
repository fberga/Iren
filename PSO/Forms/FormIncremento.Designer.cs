namespace Iren.PSO.Forms
{
    partial class FormIncremento
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
            this.panelCentrale = new System.Windows.Forms.Panel();
            this.lbErrore = new System.Windows.Forms.Label();
            this.lbScegli = new System.Windows.Forms.Label();
            this.txtValore = new System.Windows.Forms.TextBox();
            this.txtPercentuale = new System.Windows.Forms.TextBox();
            this.rdbIncremento = new System.Windows.Forms.RadioButton();
            this.rdbPercentuale = new System.Windows.Forms.RadioButton();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnRipristina = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelCentrale.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.lbErrore);
            this.panelCentrale.Controls.Add(this.lbScegli);
            this.panelCentrale.Controls.Add(this.txtValore);
            this.panelCentrale.Controls.Add(this.txtPercentuale);
            this.panelCentrale.Controls.Add(this.rdbIncremento);
            this.panelCentrale.Controls.Add(this.rdbPercentuale);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(0, 0);
            this.panelCentrale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Size = new System.Drawing.Size(552, 130);
            this.panelCentrale.TabIndex = 5;
            // 
            // lbErrore
            // 
            this.lbErrore.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbErrore.ForeColor = System.Drawing.Color.Red;
            this.lbErrore.Location = new System.Drawing.Point(0, 99);
            this.lbErrore.Name = "lbErrore";
            this.lbErrore.Size = new System.Drawing.Size(552, 31);
            this.lbErrore.TabIndex = 9;
            this.lbErrore.Text = "label1";
            this.lbErrore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbScegli
            // 
            this.lbScegli.AutoSize = true;
            this.lbScegli.Location = new System.Drawing.Point(12, 5);
            this.lbScegli.Name = "lbScegli";
            this.lbScegli.Size = new System.Drawing.Size(195, 20);
            this.lbScegli.TabIndex = 8;
            this.lbScegli.Text = "Scegli il tipo di incremento:";
            // 
            // txtValore
            // 
            this.txtValore.Enabled = false;
            this.txtValore.Location = new System.Drawing.Point(166, 68);
            this.txtValore.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtValore.Name = "txtValore";
            this.txtValore.Size = new System.Drawing.Size(373, 26);
            this.txtValore.TabIndex = 7;
            this.txtValore.EnabledChanged += new System.EventHandler(this.TextElements_EnabledChanged);
            this.txtValore.TextChanged += new System.EventHandler(this.TextElements_TextChanged);
            // 
            // txtPercentuale
            // 
            this.txtPercentuale.Location = new System.Drawing.Point(166, 36);
            this.txtPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPercentuale.Name = "txtPercentuale";
            this.txtPercentuale.Size = new System.Drawing.Size(373, 26);
            this.txtPercentuale.TabIndex = 6;
            this.txtPercentuale.EnabledChanged += new System.EventHandler(this.TextElements_EnabledChanged);
            this.txtPercentuale.TextChanged += new System.EventHandler(this.TextElements_TextChanged);
            // 
            // rdbIncremento
            // 
            this.rdbIncremento.AutoSize = true;
            this.rdbIncremento.Location = new System.Drawing.Point(30, 70);
            this.rdbIncremento.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbIncremento.Name = "rdbIncremento";
            this.rdbIncremento.Size = new System.Drawing.Size(77, 24);
            this.rdbIncremento.TabIndex = 5;
            this.rdbIncremento.Text = "Valore:";
            this.rdbIncremento.UseVisualStyleBackColor = true;
            this.rdbIncremento.CheckedChanged += new System.EventHandler(this.TipoIncermento_checkedChanged);
            // 
            // rdbPercentuale
            // 
            this.rdbPercentuale.AutoSize = true;
            this.rdbPercentuale.Checked = true;
            this.rdbPercentuale.Location = new System.Drawing.Point(30, 36);
            this.rdbPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbPercentuale.Name = "rdbPercentuale";
            this.rdbPercentuale.Size = new System.Drawing.Size(116, 24);
            this.rdbPercentuale.TabIndex = 4;
            this.rdbPercentuale.TabStop = true;
            this.rdbPercentuale.Text = "Percentuale:";
            this.rdbPercentuale.UseVisualStyleBackColor = true;
            this.rdbPercentuale.CheckedChanged += new System.EventHandler(this.TipoIncermento_checkedChanged);
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnRipristina);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(0, 130);
            this.panelButtons.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 5, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(552, 53);
            this.panelButtons.TabIndex = 13;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(213, 5);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 48);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Modifica";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnRipristina
            // 
            this.btnRipristina.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnRipristina.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRipristina.Location = new System.Drawing.Point(326, 5);
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
            this.btnAnnulla.Location = new System.Drawing.Point(439, 5);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // FormIncremento
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(552, 183);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormIncremento";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormIncremento";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormIncremento_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormIncremento_FormClosed);
            this.panelCentrale.ResumeLayout(false);
            this.panelCentrale.PerformLayout();
            this.panelButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelCentrale;
        private System.Windows.Forms.TextBox txtValore;
        private System.Windows.Forms.TextBox txtPercentuale;
        private System.Windows.Forms.RadioButton rdbIncremento;
        private System.Windows.Forms.RadioButton rdbPercentuale;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Label lbScegli;
        private System.Windows.Forms.Button btnRipristina;
        private System.Windows.Forms.Label lbErrore;
    }
}