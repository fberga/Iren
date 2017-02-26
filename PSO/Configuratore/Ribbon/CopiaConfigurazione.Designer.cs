namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    partial class CopiaConfigurazione
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.panelBottom = new System.Windows.Forms.Panel();
            this.btnCopia = new System.Windows.Forms.Button();
            this.btnChiudi = new System.Windows.Forms.Button();
            this.listBoxUtentiFrom = new System.Windows.Forms.ListBox();
            this.listBoxUtentiTo = new System.Windows.Forms.ListBox();
            this.listBoxApplicazioni = new System.Windows.Forms.ListBox();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnEmpty = new System.Windows.Forms.Button();
            this.lbDa = new System.Windows.Forms.Label();
            this.lbA = new System.Windows.Forms.Label();
            this.tableLayoutPanel1.SuspendLayout();
            this.panelBottom.SuspendLayout();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.Controls.Add(this.lbA, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.panelBottom, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.listBoxUtentiFrom, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.listBoxUtentiTo, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.listBoxApplicazioni, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.panel1, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.lbDa, 0, 0);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Margin = new System.Windows.Forms.Padding(4);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 31F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 48F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(561, 574);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // panelBottom
            // 
            this.tableLayoutPanel1.SetColumnSpan(this.panelBottom, 3);
            this.panelBottom.Controls.Add(this.btnCopia);
            this.panelBottom.Controls.Add(this.btnChiudi);
            this.panelBottom.Location = new System.Drawing.Point(0, 526);
            this.panelBottom.Margin = new System.Windows.Forms.Padding(0);
            this.panelBottom.Name = "panelBottom";
            this.panelBottom.Size = new System.Drawing.Size(561, 48);
            this.panelBottom.TabIndex = 16;
            // 
            // btnCopia
            // 
            this.btnCopia.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnCopia.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCopia.Location = new System.Drawing.Point(335, 0);
            this.btnCopia.Margin = new System.Windows.Forms.Padding(0);
            this.btnCopia.Name = "btnCopia";
            this.btnCopia.Size = new System.Drawing.Size(113, 48);
            this.btnCopia.TabIndex = 4;
            this.btnCopia.Text = "Copia";
            this.btnCopia.UseVisualStyleBackColor = true;
            this.btnCopia.Click += new System.EventHandler(this.btnCopia_Click);
            // 
            // btnChiudi
            // 
            this.btnChiudi.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnChiudi.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnChiudi.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnChiudi.Location = new System.Drawing.Point(448, 0);
            this.btnChiudi.Margin = new System.Windows.Forms.Padding(0);
            this.btnChiudi.Name = "btnChiudi";
            this.btnChiudi.Size = new System.Drawing.Size(113, 48);
            this.btnChiudi.TabIndex = 5;
            this.btnChiudi.Text = "Chiudi";
            this.btnChiudi.UseVisualStyleBackColor = true;
            // 
            // listBoxUtentiFrom
            // 
            this.listBoxUtentiFrom.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxUtentiFrom.FormattingEnabled = true;
            this.listBoxUtentiFrom.ItemHeight = 20;
            this.listBoxUtentiFrom.Location = new System.Drawing.Point(2, 33);
            this.listBoxUtentiFrom.Margin = new System.Windows.Forms.Padding(2);
            this.listBoxUtentiFrom.Name = "listBoxUtentiFrom";
            this.listBoxUtentiFrom.Size = new System.Drawing.Size(276, 228);
            this.listBoxUtentiFrom.TabIndex = 0;
            this.listBoxUtentiFrom.SelectedIndexChanged += new System.EventHandler(this.listBoxUtentiFrom_SelectedIndexChanged);
            // 
            // listBoxUtentiTo
            // 
            this.listBoxUtentiTo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxUtentiTo.FormattingEnabled = true;
            this.listBoxUtentiTo.ItemHeight = 20;
            this.listBoxUtentiTo.Location = new System.Drawing.Point(282, 33);
            this.listBoxUtentiTo.Margin = new System.Windows.Forms.Padding(2);
            this.listBoxUtentiTo.Name = "listBoxUtentiTo";
            this.tableLayoutPanel1.SetRowSpan(this.listBoxUtentiTo, 3);
            this.listBoxUtentiTo.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxUtentiTo.Size = new System.Drawing.Size(277, 491);
            this.listBoxUtentiTo.TabIndex = 3;
            // 
            // listBoxApplicazioni
            // 
            this.listBoxApplicazioni.Dock = System.Windows.Forms.DockStyle.Fill;
            this.listBoxApplicazioni.FormattingEnabled = true;
            this.listBoxApplicazioni.ItemHeight = 20;
            this.listBoxApplicazioni.Location = new System.Drawing.Point(2, 296);
            this.listBoxApplicazioni.Margin = new System.Windows.Forms.Padding(2);
            this.listBoxApplicazioni.Name = "listBoxApplicazioni";
            this.listBoxApplicazioni.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.listBoxApplicazioni.Size = new System.Drawing.Size(276, 228);
            this.listBoxApplicazioni.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.btnEmpty);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 263);
            this.panel1.Margin = new System.Windows.Forms.Padding(0, 0, 2, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(278, 31);
            this.panel1.TabIndex = 17;
            // 
            // btnEmpty
            // 
            this.btnEmpty.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnEmpty.Location = new System.Drawing.Point(146, 0);
            this.btnEmpty.Margin = new System.Windows.Forms.Padding(0);
            this.btnEmpty.Name = "btnEmpty";
            this.btnEmpty.Size = new System.Drawing.Size(132, 31);
            this.btnEmpty.TabIndex = 1;
            this.btnEmpty.Text = "Vuota selezione";
            this.btnEmpty.UseVisualStyleBackColor = true;
            this.btnEmpty.Click += new System.EventHandler(this.btnEmpty_Click);
            // 
            // lbDa
            // 
            this.lbDa.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lbDa.AutoSize = true;
            this.lbDa.Location = new System.Drawing.Point(3, 11);
            this.lbDa.Name = "lbDa";
            this.lbDa.Size = new System.Drawing.Size(30, 20);
            this.lbDa.TabIndex = 18;
            this.lbDa.Text = "Da";
            // 
            // lbA
            // 
            this.lbA.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lbA.AutoSize = true;
            this.lbA.Location = new System.Drawing.Point(283, 11);
            this.lbA.Name = "lbA";
            this.lbA.Size = new System.Drawing.Size(20, 20);
            this.lbA.TabIndex = 19;
            this.lbA.Text = "A";
            // 
            // CopiaConfigurazione
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.ClientSize = new System.Drawing.Size(779, 602);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.Name = "CopiaConfigurazione";
            this.Text = "CopiaConfigurazione";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.panelBottom.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.ListBox listBoxUtentiFrom;
        private System.Windows.Forms.ListBox listBoxUtentiTo;
        private System.Windows.Forms.ListBox listBoxApplicazioni;
        private System.Windows.Forms.Panel panelBottom;
        private System.Windows.Forms.Button btnCopia;
        private System.Windows.Forms.Button btnChiudi;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button btnEmpty;
        private System.Windows.Forms.Label lbA;
        private System.Windows.Forms.Label lbDa;
    }
}