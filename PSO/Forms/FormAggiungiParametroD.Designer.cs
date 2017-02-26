namespace Iren.ToolsExcel.Forms
{
    partial class FormAggiungiParametroD
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
            this.dateTimeIV = new System.Windows.Forms.DateTimePicker();
            this.groupDatiParD = new System.Windows.Forms.GroupBox();
            this.txtValore = new System.Windows.Forms.TextBox();
            this.labelValore = new System.Windows.Forms.Label();
            this.labelDataIV = new System.Windows.Forms.Label();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnAggiungi = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.groupDatiParD.SuspendLayout();
            this.panelButtons.SuspendLayout();
            this.SuspendLayout();
            // 
            // dateTimeIV
            // 
            this.dateTimeIV.Location = new System.Drawing.Point(118, 33);
            this.dateTimeIV.Name = "dateTimeIV";
            this.dateTimeIV.Size = new System.Drawing.Size(247, 26);
            this.dateTimeIV.TabIndex = 0;
            this.dateTimeIV.ValueChanged += new System.EventHandler(this.dateTimeIV_ValueChanged);
            this.dateTimeIV.Validating += new System.ComponentModel.CancelEventHandler(this.dateTimeIV_Validating);
            // 
            // groupDatiParD
            // 
            this.groupDatiParD.Controls.Add(this.txtValore);
            this.groupDatiParD.Controls.Add(this.labelValore);
            this.groupDatiParD.Controls.Add(this.labelDataIV);
            this.groupDatiParD.Controls.Add(this.dateTimeIV);
            this.groupDatiParD.Dock = System.Windows.Forms.DockStyle.Fill;
            this.groupDatiParD.Location = new System.Drawing.Point(3, 0);
            this.groupDatiParD.Name = "groupDatiParD";
            this.groupDatiParD.Size = new System.Drawing.Size(393, 109);
            this.groupDatiParD.TabIndex = 1;
            this.groupDatiParD.TabStop = false;
            this.groupDatiParD.Text = "XXX";
            // 
            // txtValore
            // 
            this.txtValore.Location = new System.Drawing.Point(118, 66);
            this.txtValore.Name = "txtValore";
            this.txtValore.Size = new System.Drawing.Size(247, 26);
            this.txtValore.TabIndex = 3;
            this.txtValore.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtValore_KeyDown);
            this.txtValore.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtValore_KeyPress);
            // 
            // labelValore
            // 
            this.labelValore.AutoSize = true;
            this.labelValore.Location = new System.Drawing.Point(3, 69);
            this.labelValore.Name = "labelValore";
            this.labelValore.Size = new System.Drawing.Size(59, 20);
            this.labelValore.TabIndex = 2;
            this.labelValore.Text = "Valore:";
            // 
            // labelDataIV
            // 
            this.labelDataIV.AutoSize = true;
            this.labelDataIV.Location = new System.Drawing.Point(3, 38);
            this.labelDataIV.Name = "labelDataIV";
            this.labelDataIV.Size = new System.Drawing.Size(106, 20);
            this.labelDataIV.TabIndex = 1;
            this.labelDataIV.Text = "Inizio Validità:";
            // 
            // panelButtons
            // 
            this.panelButtons.Controls.Add(this.btnAggiungi);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(3, 109);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.panelButtons.Size = new System.Drawing.Size(393, 53);
            this.panelButtons.TabIndex = 14;
            // 
            // btnAggiungi
            // 
            this.btnAggiungi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAggiungi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAggiungi.Location = new System.Drawing.Point(167, 3);
            this.btnAggiungi.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAggiungi.Name = "btnAggiungi";
            this.btnAggiungi.Size = new System.Drawing.Size(113, 50);
            this.btnAggiungi.TabIndex = 4;
            this.btnAggiungi.Text = "Aggiungi";
            this.btnAggiungi.UseVisualStyleBackColor = true;
            this.btnAggiungi.Click += new System.EventHandler(this.btnAggiungi_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(280, 3);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // FormAggiungiParametroD
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(399, 165);
            this.Controls.Add(this.groupDatiParD);
            this.Controls.Add(this.panelButtons);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "FormAggiungiParametroD";
            this.Padding = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.groupDatiParD.ResumeLayout(false);
            this.groupDatiParD.PerformLayout();
            this.panelButtons.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DateTimePicker dateTimeIV;
        private System.Windows.Forms.GroupBox groupDatiParD;
        private System.Windows.Forms.Label labelDataIV;
        private System.Windows.Forms.TextBox txtValore;
        private System.Windows.Forms.Label labelValore;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnAggiungi;
        private System.Windows.Forms.Button btnAnnulla;
    }
}