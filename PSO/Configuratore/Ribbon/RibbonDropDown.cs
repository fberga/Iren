using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class RibbonDropDown : SelectablePanel, IRibbonControl
    {
        public const string NEW_COMBO_PREFIX = "New Combo";

        private TextBox _label = new TextBox();
        private ComboBox _drp = new ComboBox();
        private Point _startPt = new Point(int.MaxValue, int.MaxValue);


        public int IdTipologia { get { return 3; } set { IdTipologia = value; } }
        public string Description { get; set; }
        public string ScreenTip { get; set; }
        //public new string Name { get; set; }
        public new string Text { get { return base.Text; } set { base.Text = _label.Text = value; } }
        public int Slot { get { return 2; } }
        public int Dimension { get { return -1; } }
        public bool ToggleButton { get { return false; } }
        public string ImageKey { get { return ""; } }
        public int IdControllo { get; private set; }
        public List<int> Functions { get; set; }
        public new bool Enabled { get { return base.Enabled; } set { base.Enabled = _drp.Enabled = _label.Enabled = value; } }

        public RibbonDropDown()
        {            
            this.Padding = new Padding(4, (33 - _label.Height) / 2, 4, 4);
            this.Controls.Add(_drp);
            this.Controls.Add(_label);

            Height = 66;

            _label.Click += SelectAllText;
            _label.KeyDown += AvoidNewLine;
            _label.Leave += CheckTextChanged;
            _label.MouseMove += ControlMouseMove;
            _drp.MouseMove += ControlMouseMove;
            _label.MouseLeave += ControlMouseLeave;
            _drp.MouseLeave += ControlMouseLeave;

            _label.Top = Padding.Top;
            _label.Left = Padding.Left;
            _label.Multiline = true;
            _label.Height = 25;            
            
            _label.BorderStyle = BorderStyle.None;
            _label.BackColor = BackColor;
            _drp.Top = Height - _drp.Height - 10;
            _drp.Left = Padding.Left;

            _drp.Width = 40;
            Font = Utility.StdFont;
            Functions = new List<int>();
        }

        private void AvoidNewLine(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                Parent.Focus();
            }
        }
        public RibbonDropDown(int id)
            : this()
        {
            IdControllo = id;
        }
        public RibbonDropDown(Control ribbon)
            : this()
        {
            using (ConfiguratoreControllo cc = new ConfiguratoreControllo(ribbon, typeof(RibbonDropDown)))
            {
                if (cc.ShowDialog() == DialogResult.OK)
                {
                    Name = cc.CtrlName;
                    Text = cc.CtrlText;
                }
                else
                {
                    Dispose();
                    return;
                }
            }
            _label.TextAlign = HorizontalAlignment.Left;
            
            SetWidth();
        }

        public void SetWidth()
        {
            //SizeF s = Utility.MeasureTextSize(_label);
            int width = Math.Max(_label.GetPreferredSize(_label.Size).Width, _drp.Width);
            _label.Width = width;
            this.Width = width + 2 * Padding.Left;
        }

        private void SelectAllText(object sender, EventArgs e)
        {
            _label.SelectAll();
        }

        private void CheckTextChanged(object sender, EventArgs e)
        {
            if (Name != _label.Text.Replace(" ", ""))
            {
                Name = _label.Text.Replace(" ", "");
                SetWidth();
                //_label.SelectAll();
                _label.SelectionStart = 0;
            }
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            if (Parent != null)
            {
                ControlContainer parent = Parent as ControlContainer;
                parent.SetContainerWidth();
            }
            base.OnSizeChanged(e);
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                if (IdControllo == 0)
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
                else
                {
                    RibbonGroup grp = Parent.Parent as RibbonGroup;

                    using (AssegnaFunzioni afForm = new AssegnaFunzioni(this, grp, 1, 62))
                    {
                        afForm.ShowDialog();
                    }
                }
            }
            base.OnDoubleClick(e);
        }
        protected void ControlMouseMove(object sender, MouseEventArgs e)
        {
            base.OnMouseEnter(e);
            BackColor = Color.FromKnownColor(KnownColor.ControlDark);
            _label.BackColor = BackColor;
        }
        protected override void OnMouseMove(MouseEventArgs mevent)
        {
            if (mevent.Button == System.Windows.Forms.MouseButtons.Left && Math.Pow(mevent.Location.X - _startPt.X, 2) + Math.Pow(mevent.Location.Y - _startPt.Y, 2) > Math.Pow(SystemInformation.DragSize.Height, 2))
                DoDragDrop(this, DragDropEffects.Move);

            ControlMouseMove(this, mevent);
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            ControlMouseLeave(this, e);
        }
        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            _startPt = mevent.Location;
            Select();
            if (mevent.Clicks == 2)
                OnDoubleClick(mevent);

            //base.OnMouseMove(mevent);
        }
        private void ControlMouseLeave(object sender, EventArgs e)
        {
            BackColor = Color.FromKnownColor(KnownColor.Control);
            _label.BackColor = BackColor;
        }

        protected override void Dispose(bool disposing)
        {
            if (!Utility.Refreshing && ConfiguratoreRibbon.ControlliUtilizzati.Contains(IdControllo))
                ConfiguratoreRibbon.GruppoControlloCancellati.Add(ConfiguratoreRibbon.GruppoControlloUtilizzati[ConfiguratoreRibbon.ControlliUtilizzati.IndexOf(IdControllo)]);
            base.Dispose(disposing);
        }
    }
}
