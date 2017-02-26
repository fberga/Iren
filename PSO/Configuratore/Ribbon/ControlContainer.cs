using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    class ControlContainer : SelectablePanel
    {
        public int FreeSlot { get; private set; }
        public int CtrlCount { get; private set; }

        //private Dictionary<int, Control> _slots = new Dictionary<int, Control>();

        public ControlContainer()
        {
            //BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            FreeSlot = 3;
            CtrlCount = 0;
            AllowDrop = true;
            Padding = new Padding(1, 1, 0, 0);
            Font = Utility.StdFont;
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            //base.OnPaint(pe);
            if (this.Focused)
            {
                var rc = this.ClientRectangle;
                //rc.Inflate(1, 1);
                ControlPaint.DrawFocusRectangle(pe.Graphics, rc);
            }
        }

        public void SetContainerWidth()
        {
            int width =
                Utility.GetAll(this)
                .Select(c => c.Width)
                .DefaultIfEmpty()
                .Max();

            Width = width == 0 ? 50 : width + 2;
        }

        protected override void OnControlAdded(ControlEventArgs e)
        {
            SetContainerWidth();

            if (CtrlCount == 0 && Parent != null)
                BorderStyle = System.Windows.Forms.BorderStyle.None;

            CtrlCount += 1;

            FreeSlot -= ((IRibbonControl)e.Control).Slot;

            if (e.Control.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = (RibbonButton)e.Control;
                btn.PropertyChanged += ButtonPropertyChanged;
            }

            base.OnControlAdded(e);
        }
        protected override void OnControlRemoved(ControlEventArgs e)
        {
            SetContainerWidth();
            CtrlCount -= 1;

            if (CtrlCount == 0)
                BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

            FreeSlot += ((IRibbonControl)e.Control).Slot;

            if (e.Control.GetType() == typeof(RibbonButton))
            {
                RibbonButton btn = (RibbonButton)e.Control;
                btn.PropertyChanged += ButtonPropertyChanged;
            }

            CompactCtrls();

            base.OnControlRemoved(e);
        }

        private void CompactCtrls()
        {
            var ctrls = Controls;
            if (ctrls.Count > 0)
            {
                ctrls[0].Top = Padding.Top;
                for (int i = 1; i < ctrls.Count; i++)
                    ctrls[i].Top = ctrls[i - 1].Bottom;
            }
        }

        private void ButtonPropertyChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            RibbonButton btn = sender as RibbonButton;
            if (btn.Parent == this && e.PropertyName == "Dimensione")
            {
                //non può essere diverso: o va a 1 e occupa tutto lo spazio, o va a 0 e occupa uno solo dei 3 slot
                FreeSlot = 3;
                FreeSlot -= btn.Slot;
            }
        }

        protected override void OnDragEnter(DragEventArgs drgevent)
        {
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;
            if (ctrl.Parent != this)
            {
                int slot = ((IRibbonControl)ctrl).Slot;
                if(slot <= FreeSlot)
                    drgevent.Effect = DragDropEffects.Move;
                else
                    drgevent.Effect = DragDropEffects.None;
            }

            base.OnDragEnter(drgevent);
        }
        protected override void OnDragOver(DragEventArgs drgevent)
        {
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;
            
            //DisplaceObjects(drgevent, ctrl.Height);

            base.OnDragOver(drgevent);
        }
        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            if (Parent != null)
            {
                Utility.UpdateGroupDimension(Parent);
            }
        }
        protected override void OnDragDrop(DragEventArgs drgevent)
        {            
            Control ctrl = drgevent.Data.GetData(drgevent.Data.GetFormats()[0]) as Control;

            int top = 0;
            if (ctrl != null)
            {
                
                int slot = ((IRibbonControl)ctrl).Slot;
                if (slot < 3)
                {
                    top =
                        Utility.GetAll(this, typeof(IRibbonControl))
                        .Select(b => b.Bottom)
                        .DefaultIfEmpty()
                        .Max();
                }

                Controls.Add(ctrl);
                ctrl.Left = Padding.Left;
                ctrl.Top = top == 0 ? Padding.Top : top;

            }
            base.OnDragDrop(drgevent);
        }

        protected override void OnMouseClick(MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Right)
            {
                Control c = null;
                foreach (Control ctrl in Controls)
                {
                    if (ctrl.DisplayRectangle.IntersectsWith(new Rectangle(e.Location, new Size(1, 1)))) 
                    {
                        c = ctrl;
                        break;
                    }
                }


            }
        }
    }
}
