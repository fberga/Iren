namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class CambiaImmagine
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
            this.mainLayout = new System.Windows.Forms.TableLayoutPanel();
            this.imageListView = new System.Windows.Forms.ListView();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.lbSelImmagine = new System.Windows.Forms.Label();
            this.mainLayout.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainLayout
            // 
            this.mainLayout.ColumnCount = 2;
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainLayout.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 539F));
            this.mainLayout.Controls.Add(this.imageListView, 0, 1);
            this.mainLayout.Controls.Add(this.panelBottom, 0, 2);
            this.mainLayout.Controls.Add(this.lbSelImmagine, 0, 0);
            this.mainLayout.Location = new System.Drawing.Point(0, 0);
            this.mainLayout.Name = "mainLayout";
            this.mainLayout.RowCount = 3;
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 24F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 52F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.mainLayout.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.mainLayout.Size = new System.Drawing.Size(618, 506);
            this.mainLayout.TabIndex = 17;
            // 
            // imageListView
            // 
            this.mainLayout.SetColumnSpan(this.imageListView, 2);
            this.imageListView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.imageListView.Location = new System.Drawing.Point(2, 26);
            this.imageListView.Margin = new System.Windows.Forms.Padding(2);
            this.imageListView.MultiSelect = false;
            this.imageListView.Name = "imageListView";
            this.imageListView.ShowItemToolTips = true;
            this.imageListView.Size = new System.Drawing.Size(614, 426);
            this.imageListView.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.imageListView.TabIndex = 0;
            this.imageListView.UseCompatibleStateImageBehavior = false;
            this.imageListView.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.SelectItemByDoubleClick);
            // 
            // panelBottom
            // 
            this.mainLayout.SetColumnSpan(this.panelBottom, 2);
            this.panelBottom.Controls.Add(this.btnApplica);
            this.panelBottom.Controls.Add(this.btnAnnulla);
            this.panelBottom.Location = new System.Drawing.Point(2, 456);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(2);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(614, 48);
            this.panelBottom.TabIndex = 14;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(388, 0);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 48);
            this.btnApplica.TabIndex = 8;
            this.btnApplica.Text = "Applica";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.Applica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(501, 0);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(6, 8, 6, 8);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 48);
            this.btnAnnulla.TabIndex = 7;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            // 
            // lbSelImmagine
            // 
            this.lbSelImmagine.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lbSelImmagine.AutoSize = true;
            this.mainLayout.SetColumnSpan(this.lbSelImmagine, 2);
            this.lbSelImmagine.Location = new System.Drawing.Point(3, 3);
            this.lbSelImmagine.Margin = new System.Windows.Forms.Padding(3);
            this.lbSelImmagine.Name = "lbSelImmagine";
            this.lbSelImmagine.Size = new System.Drawing.Size(303, 18);
            this.lbSelImmagine.TabIndex = 15;
            this.lbSelImmagine.Text = "Seleziona l\'immagine da applicare al tasto";
            // 
            // CambiaImmagine
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(663, 532);
            this.Controls.Add(this.mainLayout);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "CambiaImmagine";
            this.Text = "Cambia immagine";
            this.mainLayout.ResumeLayout(false);
            this.mainLayout.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel mainLayout;
        private System.Windows.Forms.ListView imageListView;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Label lbSelImmagine;
    }
}