namespace Iren.PSO.Forms
{
    partial class FormSelezioneDate
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
            this.checkDate = new System.Windows.Forms.CheckedListBox();
            this.panelButtons = new System.Windows.Forms.Panel();
            this.btnApplica = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.panelAll = new System.Windows.Forms.Panel();
            this.panelTop = new System.Windows.Forms.SplitContainer();
            this.checkClusterDate = new System.Windows.Forms.CheckedListBox();
            this.panelButtons.SuspendLayout();
            this.panelAll.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.panelTop)).BeginInit();
            this.panelTop.Panel1.SuspendLayout();
            this.panelTop.Panel2.SuspendLayout();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // checkDate
            // 
            this.checkDate.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.checkDate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkDate.FormattingEnabled = true;
            this.checkDate.Location = new System.Drawing.Point(3, 3);
            this.checkDate.Name = "checkDate";
            this.checkDate.Size = new System.Drawing.Size(428, 225);
            this.checkDate.TabIndex = 0;
            this.checkDate.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkDate_ItemCheck);
            this.checkDate.MouseLeave += new System.EventHandler(this.CheckedListBox_MouseLeave);
            this.checkDate.MouseMove += new System.Windows.Forms.MouseEventHandler(this.CheckedListBox_MouseMove);
            // 
            // panelButtons
            // 
            this.panelButtons.BackColor = System.Drawing.SystemColors.Window;
            this.panelButtons.Controls.Add(this.btnApplica);
            this.panelButtons.Controls.Add(this.btnAnnulla);
            this.panelButtons.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panelButtons.Location = new System.Drawing.Point(0, 335);
            this.panelButtons.Name = "panelButtons";
            this.panelButtons.Padding = new System.Windows.Forms.Padding(0, 0, 3, 3);
            this.panelButtons.Size = new System.Drawing.Size(434, 53);
            this.panelButtons.TabIndex = 13;
            // 
            // btnApplica
            // 
            this.btnApplica.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnApplica.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnApplica.Location = new System.Drawing.Point(205, 0);
            this.btnApplica.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnApplica.Name = "btnApplica";
            this.btnApplica.Size = new System.Drawing.Size(113, 50);
            this.btnApplica.TabIndex = 4;
            this.btnApplica.Text = "OK";
            this.btnApplica.UseVisualStyleBackColor = true;
            this.btnApplica.Click += new System.EventHandler(this.btnApplica_Click);
            // 
            // btnAnnulla
            // 
            this.btnAnnulla.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnAnnulla.Dock = System.Windows.Forms.DockStyle.Right;
            this.btnAnnulla.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAnnulla.Location = new System.Drawing.Point(318, 0);
            this.btnAnnulla.Margin = new System.Windows.Forms.Padding(4, 5, 4, 5);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(113, 50);
            this.btnAnnulla.TabIndex = 5;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);
            // 
            // panelAll
            // 
            this.panelAll.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panelAll.Controls.Add(this.panelTop);
            this.panelAll.Controls.Add(this.panelButtons);
            this.panelAll.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelAll.Location = new System.Drawing.Point(0, 0);
            this.panelAll.Name = "panelAll";
            this.panelAll.Size = new System.Drawing.Size(436, 390);
            this.panelAll.TabIndex = 1;
            // 
            // panelTop
            // 
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelTop.FixedPanel = System.Windows.Forms.FixedPanel.Panel1;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // panelTop.Panel1
            // 
            this.panelTop.Panel1.Controls.Add(this.checkClusterDate);
            this.panelTop.Panel1.Padding = new System.Windows.Forms.Padding(3);
            // 
            // panelTop.Panel2
            // 
            this.panelTop.Panel2.Controls.Add(this.checkDate);
            this.panelTop.Panel2.Padding = new System.Windows.Forms.Padding(3);
            this.panelTop.Size = new System.Drawing.Size(434, 335);
            this.panelTop.SplitterDistance = 100;
            this.panelTop.TabIndex = 2;
            // 
            // checkClusterDate
            // 
            this.checkClusterDate.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.checkClusterDate.Dock = System.Windows.Forms.DockStyle.Fill;
            this.checkClusterDate.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.checkClusterDate.FormattingEnabled = true;
            this.checkClusterDate.Location = new System.Drawing.Point(3, 3);
            this.checkClusterDate.Name = "checkClusterDate";
            this.checkClusterDate.Size = new System.Drawing.Size(428, 94);
            this.checkClusterDate.TabIndex = 1;
            this.checkClusterDate.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.checkClusterDate_ItemCheck);
            this.checkClusterDate.MouseLeave += new System.EventHandler(this.CheckedListBox_MouseLeave);
            this.checkClusterDate.MouseMove += new System.Windows.Forms.MouseEventHandler(this.CheckedListBox_MouseMove);
            // 
            // FormSelezioneDate
            // 
            this.AcceptButton = this.btnApplica;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(436, 390);
            this.Controls.Add(this.panelAll);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Name = "FormSelezioneDate";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "FormSelezioneDate";
            this.VisibleChanged += new System.EventHandler(this.FormSelezioneDate_VisibleChanged);
            this.panelButtons.ResumeLayout(false);
            this.panelAll.ResumeLayout(false);
            this.panelTop.Panel1.ResumeLayout(false);
            this.panelTop.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.panelTop)).EndInit();
            this.panelTop.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.CheckedListBox checkDate;
        private System.Windows.Forms.Panel panelButtons;
        private System.Windows.Forms.Button btnApplica;
        private System.Windows.Forms.Button btnAnnulla;
        private System.Windows.Forms.Panel panelAll;
        private System.Windows.Forms.SplitContainer panelTop;
        private System.Windows.Forms.CheckedListBox checkClusterDate;
    }
}