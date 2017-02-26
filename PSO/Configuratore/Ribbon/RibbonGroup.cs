using System;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class RibbonGroup : SelectablePanel
    {
        public const string NEW_GROUP_PREFIX = "New Group";

        private TextBox Label { get; set; }
        public new string Text { get { return base.Text; } set { base.Text = Label.Text = value; } }
        //public string Name { get; set; }

        public int IdGruppo { get; private set; }
        
        public RibbonGroup() 
            : base()
        {
            this.Padding = new Padding(4, 4, 4, 4);

            this.Label = new TextBox();

            this.Label.ReadOnly = true;
            this.Controls.Add(this.Label);
            this.Label.Dock = DockStyle.Bottom;
            this.Label.AutoSize = false;
            this.Label.TextAlign = HorizontalAlignment.Center;
            this.Label.Leave += CheckTextChanged;

            this.Label.BorderStyle = BorderStyle.None;
        }
        public RibbonGroup(Control ribbon, int id)
            : this()
        {
            this.IdGruppo = id;

            BackColor = ControlPaint.LightLight(ribbon.BackColor);

            this.Top = ribbon.Padding.Top;
            this.Width = (int)(Utility.MeasureTextSize(this.Label).Width + 20);
            this.Height = ribbon.Height - ribbon.Padding.Top - ribbon.Padding.Bottom - 20;
            this.Label.BackColor = ControlPaint.LightLight(ribbon.BackColor);
        }
        public RibbonGroup(Control ribbon)
            : this(ribbon, 0)
        {
            using (ConfiguratoreControllo cfgCtrl = new ConfiguratoreControllo(ribbon, typeof(RibbonGroup)))
            {
                if (cfgCtrl.ShowDialog() == DialogResult.OK)
                {
                    Name = cfgCtrl.CtrlName;
                    Text = cfgCtrl.CtrlText;
                }
                else
                {
                    Dispose();
                    return;
                }
            }

            //BackColor = ControlPaint.LightLight(ribbon.BackColor);

            //this.Top = ribbon.Padding.Top;
            this.Width = (int)(Utility.MeasureTextSize(this.Label).Width + 20);
            //this.Height = ribbon.Height - ribbon.Padding.Top - ribbon.Padding.Bottom - 20;
            //this.Label.BackColor = ControlPaint.LightLight(ribbon.BackColor);
        }
        

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            var rc = this.ClientRectangle;
            ControlPaint.DrawBorder3D(pe.Graphics, rc, Border3DStyle.Etched, Border3DSide.Right);
        }
        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if(e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                if (IdGruppo == 0)
                {
                    using (ConfiguratoreControllo cfgCtrl = new ConfiguratoreControllo(this, Parent, GetType()))
                    {
                        if (cfgCtrl.ShowDialog() == DialogResult.OK)
                        {
                            Text = cfgCtrl.CtrlText;
                            Name = cfgCtrl.CtrlName;
                        }
                    }
                }
            }

            base.OnDoubleClick(e);
        }
        protected override void OnControlAdded(ControlEventArgs e)
        {
            base.OnControlAdded(e);
            Utility.UpdateGroupDimension(this);
        }
        protected override void OnControlRemoved(ControlEventArgs e)
        {
            this.CompactCtrls();
            Utility.UpdateGroupDimension(this);

            base.OnControlRemoved(e);
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            if(Parent != null)
                Utility.GroupsDisplacement(Parent);
            
            base.OnSizeChanged(e);
        }


        //private void SelectAllText(object sender, EventArgs e)
        //{
        //    this.Label.SelectAll();
        //}
        private void CheckTextChanged(object sender, EventArgs e)
        {
            if (this.Label.Name != this.Text)
            {
                this.Label.Name = this.Text;
                Utility.UpdateGroupDimension(this);
            }
        }
        private void CompactCtrls()
        {
            var ctrls = this.Controls
                .OfType<ControlContainer>()
                .OrderBy(c => c.Left)
                .ToList();

            if (ctrls.Count > 0)
            {
                ctrls[0].Left = this.Padding.Left;
                for (int i = 1; i < ctrls.Count; i++)
                    ctrls[i].Left = ctrls[i - 1].Right;
            }
        }
    }
}