namespace Iren.PSO.Forms
{
    partial class FormIncrementoMI
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
            this.groupQuantità = new System.Windows.Forms.GroupBox();
            this.Vai_a = new System.Windows.Forms.Label();
            this.comboBox_VaiA = new System.Windows.Forms.ComboBox();
            this.groupPrezzo = new System.Windows.Forms.GroupBox();
            this.rdbPercentuale = new System.Windows.Forms.RadioButton();
            this.rdbIncremento = new System.Windows.Forms.RadioButton();
            this.comboBox_applicaA = new System.Windows.Forms.ComboBox();
            this.txtPercentuale = new System.Windows.Forms.TextBox();
            this.txtValore = new System.Windows.Forms.TextBox();
            this.Operazione_txt = new System.Windows.Forms.Label();
            this.lbErrore = new System.Windows.Forms.Label();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnRipristina = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelCentrale.SuspendLayout();
            this.groupQuantità.SuspendLayout();
            this.groupPrezzo.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // panelCentrale
            // 
            this.panelCentrale.Controls.Add(this.groupQuantità);
            this.panelCentrale.Controls.Add(this.groupPrezzo);
            this.panelCentrale.Controls.Add(this.lbErrore);
            this.panelCentrale.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelCentrale.Location = new System.Drawing.Point(0, 0);
            this.panelCentrale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.panelCentrale.Name = "panelCentrale";
            this.panelCentrale.Size = new System.Drawing.Size(552, 181);
            this.panelCentrale.TabIndex = 5;
            // 
            // groupQuantità
            // 
            this.groupQuantità.Controls.Add(this.Vai_a);
            this.groupQuantità.Controls.Add(this.comboBox_VaiA);
            this.groupQuantità.Location = new System.Drawing.Point(12, 12);
            this.groupQuantità.Name = "groupQuantità";
            this.groupQuantità.Size = new System.Drawing.Size(528, 132);
            this.groupQuantità.TabIndex = 13;
            this.groupQuantità.TabStop = false;
            this.groupQuantità.Text = "Incremento quantità";
            this.groupQuantità.Visible = false;
            // 
            // Vai_a
            // 
            this.Vai_a.AutoSize = true;
            this.Vai_a.Location = new System.Drawing.Point(7, 41);
            this.Vai_a.Name = "Vai_a";
            this.Vai_a.Size = new System.Drawing.Size(49, 20);
            this.Vai_a.TabIndex = 1;
            this.Vai_a.Text = "Vai a:";
            // 
            // comboBox_VaiA
            // 
            this.comboBox_VaiA.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_VaiA.FormattingEnabled = true;
            this.comboBox_VaiA.Location = new System.Drawing.Point(6, 71);
            this.comboBox_VaiA.Name = "comboBox_VaiA";
            this.comboBox_VaiA.Size = new System.Drawing.Size(364, 28);
            this.comboBox_VaiA.TabIndex = 0;
            this.comboBox_VaiA.SelectedValueChanged += new System.EventHandler(this.StateChanged);
            // 
            // groupPrezzo
            // 
            this.groupPrezzo.Controls.Add(this.rdbPercentuale);
            this.groupPrezzo.Controls.Add(this.rdbIncremento);
            this.groupPrezzo.Controls.Add(this.comboBox_applicaA);
            this.groupPrezzo.Controls.Add(this.txtPercentuale);
            this.groupPrezzo.Controls.Add(this.txtValore);
            this.groupPrezzo.Controls.Add(this.Operazione_txt);
            this.groupPrezzo.Location = new System.Drawing.Point(12, 12);
            this.groupPrezzo.Name = "groupPrezzo";
            this.groupPrezzo.Size = new System.Drawing.Size(528, 132);
            this.groupPrezzo.TabIndex = 12;
            this.groupPrezzo.TabStop = false;
            this.groupPrezzo.Text = "Incremento prezzo";
            this.groupPrezzo.Visible = false;
            // 
            // rdbPercentuale
            // 
            this.rdbPercentuale.AutoSize = true;
            this.rdbPercentuale.Checked = true;
            this.rdbPercentuale.Location = new System.Drawing.Point(23, 27);
            this.rdbPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbPercentuale.Name = "rdbPercentuale";
            this.rdbPercentuale.Size = new System.Drawing.Size(116, 24);
            this.rdbPercentuale.TabIndex = 4;
            this.rdbPercentuale.TabStop = true;
            this.rdbPercentuale.Text = "Percentuale:";
            this.rdbPercentuale.UseVisualStyleBackColor = true;
            this.rdbPercentuale.CheckedChanged += new System.EventHandler(this.TipoIncermento_checkedChanged);
            // 
            // rdbIncremento
            // 
            this.rdbIncremento.AutoSize = true;
            this.rdbIncremento.Location = new System.Drawing.Point(23, 61);
            this.rdbIncremento.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.rdbIncremento.Name = "rdbIncremento";
            this.rdbIncremento.Size = new System.Drawing.Size(77, 24);
            this.rdbIncremento.TabIndex = 5;
            this.rdbIncremento.Text = "Valore:";
            this.rdbIncremento.UseVisualStyleBackColor = true;
            this.rdbIncremento.CheckedChanged += new System.EventHandler(this.TipoIncermento_checkedChanged);
            // 
            // comboBox_applicaA
            // 
            this.comboBox_applicaA.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox_applicaA.Location = new System.Drawing.Point(159, 95);
            this.comboBox_applicaA.Name = "comboBox_applicaA";
            this.comboBox_applicaA.Size = new System.Drawing.Size(351, 28);
            this.comboBox_applicaA.TabIndex = 11;
            this.comboBox_applicaA.SelectedValueChanged += new System.EventHandler(this.StateChanged);
            // 
            // txtPercentuale
            // 
            this.txtPercentuale.Location = new System.Drawing.Point(159, 27);
            this.txtPercentuale.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtPercentuale.Name = "txtPercentuale";
            this.txtPercentuale.Size = new System.Drawing.Size(177, 26);
            this.txtPercentuale.TabIndex = 6;
            this.txtPercentuale.EnabledChanged += new System.EventHandler(this.TextElements_EnabledChanged);
            this.txtPercentuale.TextChanged += new System.EventHandler(this.TextElements_TextChanged);
            // 
            // txtValore
            // 
            this.txtValore.Enabled = false;
            this.txtValore.Location = new System.Drawing.Point(159, 59);
            this.txtValore.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.txtValore.Name = "txtValore";
            this.txtValore.Size = new System.Drawing.Size(177, 26);
            this.txtValore.TabIndex = 7;
            this.txtValore.EnabledChanged += new System.EventHandler(this.TextElements_EnabledChanged);
            this.txtValore.TextChanged += new System.EventHandler(this.TextElements_TextChanged);
            // 
            // Operazione_txt
            // 
            this.Operazione_txt.AutoSize = true;
            this.Operazione_txt.Location = new System.Drawing.Point(19, 98);
            this.Operazione_txt.Name = "Operazione_txt";
            this.Operazione_txt.Size = new System.Drawing.Size(78, 20);
            this.Operazione_txt.TabIndex = 10;
            this.Operazione_txt.Text = "Applica a:";
            // 
            // lbErrore
            // 
            this.lbErrore.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.lbErrore.ForeColor = System.Drawing.Color.Red;
            this.lbErrore.Location = new System.Drawing.Point(0, 150);
            this.lbErrore.Name = "lbErrore";
            this.lbErrore.Size = new System.Drawing.Size(552, 31);
            this.lbErrore.TabIndex = 9;
            this.lbErrore.Text = "label1";
            this.lbErrore.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnRipristina);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(0, 181);
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
            // FormIncrementoMI
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(552, 234);
            this.Controls.Add(this.panelCentrale);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "FormIncrementoMI";
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormIncremento";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FormIncremento_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormIncremento_FormClosed);
            this.panelCentrale.ResumeLayout(false);
            this.groupQuantità.ResumeLayout(false);
            this.groupQuantità.PerformLayout();
            this.groupPrezzo.ResumeLayout(false);
            this.groupPrezzo.PerformLayout();
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
        private System.Windows.Forms.Button btnRipristina;
        private System.Windows.Forms.Label lbErrore;
        private System.Windows.Forms.Label Operazione_txt;
        private System.Windows.Forms.GroupBox groupQuantità;
        private System.Windows.Forms.Label Vai_a;
        private System.Windows.Forms.ComboBox comboBox_VaiA;
        private System.Windows.Forms.GroupBox groupPrezzo;
        private System.Windows.Forms.ComboBox comboBox_applicaA;
    }
}