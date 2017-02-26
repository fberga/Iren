namespace Iren.PSO.Forms
{
    partial class FormCalendar
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Liberare le risorse in uso.
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
            this.calObj = new System.Windows.Forms.MonthCalendar();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnANNULLA = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // calObj
            // 
            this.calObj.Dock = System.Windows.Forms.DockStyle.Fill;
            this.calObj.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.calObj.Location = new System.Drawing.Point(6, 6);
            this.calObj.Margin = new System.Windows.Forms.Padding(14);
            this.calObj.MaxSelectionCount = 1;
            this.calObj.Name = "calObj";
            this.calObj.TabIndex = 0;
            // 
            // btnOK
            // 
            this.btnOK.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOK.Location = new System.Drawing.Point(6, 174);
            this.btnOK.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(113, 49);
            this.btnOK.TabIndex = 1;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnANNULLA
            // 
            this.btnANNULLA.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnANNULLA.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnANNULLA.Location = new System.Drawing.Point(120, 174);
            this.btnANNULLA.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnANNULLA.Name = "btnANNULLA";
            this.btnANNULLA.Size = new System.Drawing.Size(113, 49);
            this.btnANNULLA.TabIndex = 2;
            this.btnANNULLA.Text = "Annulla";
            this.btnANNULLA.UseVisualStyleBackColor = true;
            this.btnANNULLA.Click += new System.EventHandler(this.btnANNULLA_Click);
            // 
            // FormCalendar
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.CancelButton = this.btnANNULLA;
            this.ClientSize = new System.Drawing.Size(240, 230);
            this.Controls.Add(this.btnANNULLA);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.calObj);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormCalendar";
            this.Padding = new System.Windows.Forms.Padding(6);
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Calendario";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.MonthCalendar calObj;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnANNULLA;
    }
}

