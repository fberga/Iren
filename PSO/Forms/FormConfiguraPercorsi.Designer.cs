namespace Iren.PSO.Forms
{
    partial class FormConfiguraPercorsi
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
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelDati = new System.Windows.Forms.Panel();
            this.dataGridConfigurazioni = new System.Windows.Forms.DataGridView();
            this.panelButtons.SuspendLayout();
            this.panelDati.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridConfigurazioni)).BeginInit();
            this.SuspendLayout();
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(3, 281);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(1031, 53);
            this.panelButtons.TabIndex = 14;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(805, 3);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 50);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "Applica";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(918, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Chiudi";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // panelDati
            // 
            this.panelDati.Controls.Add(this.dataGridConfigurazioni);
            this.panelDati.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelDati.Location = new System.Drawing.Point(3, 3);
            this.panelDati.Name = "panelDati";
            this.panelDati.Size = new System.Drawing.Size(1031, 278);
            this.panelDati.TabIndex = 15;
            // 
            // dataGridConfigurazioni
            // 
            this.dataGridConfigurazioni.AllowUserToAddRows = false;
            this.dataGridConfigurazioni.AllowUserToDeleteRows = false;
            this.dataGridConfigurazioni.AllowUserToResizeColumns = false;
            this.dataGridConfigurazioni.AllowUserToResizeRows = false;
            this.dataGridConfigurazioni.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridConfigurazioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridConfigurazioni.Location = new System.Drawing.Point(0, 0);
            this.dataGridConfigurazioni.Name = "dataGridConfigurazioni";
            this.dataGridConfigurazioni.Size = new System.Drawing.Size(1031, 278);
            this.dataGridConfigurazioni.TabIndex = 0;
            // 
            // FormConfiguraPercorsi
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1037, 337);
            this.Controls.Add(this.panelDati);
            this.Controls.Add(this.panelButtons);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "FormConfiguraPercorsi";
            this.Padding = new System.Windows.Forms.Padding(3);
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "FormConfig";
            this.Load += new System.EventHandler(this.FormConfig_Load);
            this.panelButtons.ResumeLayout(false);
            this.panelDati.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridConfigurazioni)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Panel panelDati;
        private System.Windows.Forms.DataGridView dataGridConfigurazioni;
    }
}