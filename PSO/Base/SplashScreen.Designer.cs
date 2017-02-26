namespace Iren.PSO.Base
{
    partial class SplashScreen
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SplashScreen));
            this.panelAll = new System.Windows.Forms.Panel();
            this.lbText = new System.Windows.Forms.Label();
            this.lbCaricamento = new System.Windows.Forms.Label();
            this.loader = new System.Windows.Forms.PictureBox();
            this.panelAll.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.loader)).BeginInit();
            this.SuspendLayout();
            // 
            // panelAll
            // 
            this.panelAll.BackColor = System.Drawing.Color.Snow;
            this.panelAll.Controls.Add(this.lbText);
            this.panelAll.Controls.Add(this.lbCaricamento);
            this.panelAll.Controls.Add(this.loader);
            this.panelAll.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelAll.Location = new System.Drawing.Point(2, 2);
            this.panelAll.Name = "panelAll";
            this.panelAll.Size = new System.Drawing.Size(529, 120);
            this.panelAll.TabIndex = 3;
            // 
            // lbText
            // 
            this.lbText.AutoSize = true;
            this.lbText.BackColor = System.Drawing.Color.Transparent;
            this.lbText.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbText.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbText.Location = new System.Drawing.Point(207, 79);
            this.lbText.Name = "lbText";
            this.lbText.Size = new System.Drawing.Size(119, 18);
            this.lbText.TabIndex = 5;
            this.lbText.Text = "Inizializzazione...";
            this.lbText.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lbText.SizeChanged += new System.EventHandler(this.lbText_SizeChanged);
            // 
            // lbCaricamento
            // 
            this.lbCaricamento.AutoSize = true;
            this.lbCaricamento.BackColor = System.Drawing.Color.Transparent;
            this.lbCaricamento.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbCaricamento.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.lbCaricamento.Location = new System.Drawing.Point(227, 51);
            this.lbCaricamento.Name = "lbCaricamento";
            this.lbCaricamento.Size = new System.Drawing.Size(79, 18);
            this.lbCaricamento.TabIndex = 3;
            this.lbCaricamento.Text = "Attendere";
            this.lbCaricamento.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // loader
            // 
            this.loader.BackColor = System.Drawing.Color.Transparent;
            this.loader.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.loader.Image = ((System.Drawing.Image)(resources.GetObject("loader.Image")));
            this.loader.Location = new System.Drawing.Point(251, 17);
            this.loader.Name = "loader";
            this.loader.Size = new System.Drawing.Size(30, 31);
            this.loader.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
            this.loader.TabIndex = 4;
            this.loader.TabStop = false;
            // 
            // SplashForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(64)))), ((int)(((byte)(64)))), ((int)(((byte)(64)))));
            this.ClientSize = new System.Drawing.Size(533, 124);
            this.Controls.Add(this.panelAll);
            this.ForeColor = System.Drawing.SystemColors.Control;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "SplashForm";
            this.Padding = new System.Windows.Forms.Padding(2);
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "LoaderScreen";
            this.panelAll.ResumeLayout(false);
            this.panelAll.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.loader)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panelAll;
        private System.Windows.Forms.Label lbText;
        private System.Windows.Forms.Label lbCaricamento;
        private System.Windows.Forms.PictureBox loader;


    }
}