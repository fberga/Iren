using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.Windows.Forms;

namespace Iren.ToolsExcel.ConfiguratoreRibbon
{
    public class RibbonButton : SelectableButton, INotifyPropertyChanged, IRibbonControl
    {
        public const string NEW_BUTTON_PREFIX = "New Button";

        private Size largeBtnMinSize = new Size(50, 100);
        private Size smallBtnMaxSize = new Size(250, 33);

        private Point _startPt = new Point(int.MaxValue, int.MaxValue);
        private bool _enabled;

        public int IdTipologia { get { return ToggleButton ? 2 : 1; } }
        public int Slot { get { return Dimension == 1 ? 3 : 1; } }

        private int _dimensione = 1;
        public int Dimension {
            get
            {
                return _dimensione;
            }
            set 
            {
                _dimensione = value;
                if (_dimensione == 1)
                {
                    SetUpLargeButton();
                    SetLargeButtonDimension();
                }
                else if (_dimensione == 0)
                {
                    SetUpSmallButton();
                    SetSmallButtonDimension();
                }
            } 
        }
        public string Description { get; set; }
        public string ScreenTip { get; set; }
        public bool ToggleButton { get; set; }
        public int IdControllo { get; private set; }
        public List<int> Functions { get; set; }
        public new bool Enabled
        {
            get
            {
                return _enabled;
            }
            set
            {                
                _enabled = value;

                if (!value) 
                {
                    Image = MakeGrayscale(Image);
                    ForeColor = System.Drawing.Color.Gray;
                }
                else
                {
                    Image = Utility.ImageListNormal.Images[ImageKey];
                    if(Parent != null)
                        ForeColor = Parent.ForeColor;
                }
                Invalidate();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        public RibbonButton(string imageKey, int id)
        {
            Enabled = true;
            ImageKey = imageKey;
            IdControllo = id;
            Font = Utility.StdFont;
            Functions = new List<int>();
        }
        public RibbonButton(Control ribbon)
        {
            Enabled = true;
            Font = Utility.StdFont;
            Functions = new List<int>();
            
            SetUpLargeButton();
            Dimension = 1;

            using (ConfiguratoreTasto configuraTasto = new ConfiguratoreTasto(this, ribbon))
            {
                if (configuraTasto.ShowDialog() != DialogResult.OK)
                {
                    this.Dispose();
                    return;
                }
            }

            SetLargeButtonDimension();
        }

        public void SetUpLargeButton()
        {
            ImageList = Utility.ImageListNormal;
            MaximumSize = new Size(int.MaxValue, int.MaxValue);
            MinimumSize = largeBtnMinSize;
            ImageAlign = ContentAlignment.TopCenter;
            TextImageRelation = TextImageRelation.ImageAboveText;
            TextAlign = ContentAlignment.MiddleCenter;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
        }
        public void SetUpSmallButton()
        {
            ImageList = Utility.ImageListSmall;
            MinimumSize = new Size(0, 0);
            MaximumSize = smallBtnMaxSize;
            ImageAlign = ContentAlignment.MiddleLeft;
            TextImageRelation = TextImageRelation.ImageBeforeText;
            TextAlign = ContentAlignment.MiddleLeft;
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            AutoEllipsis = true;
        }

        public void SetLargeButtonDimension()
        {
            Width = Math.Min((int)(Utility.MeasureTextSize(this).Width + 15), 250);
            Height = MinimumSize.Height;
        }
        public void SetSmallButtonDimension()
        {
            Width = Math.Min((int)(Utility.MeasureTextSize(this).Width + 30), 250);
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

        protected void OnPropertyChanged(string name)
        {
            PropertyChangedEventHandler handler = PropertyChanged;
            if (handler != null)
            {
                handler(this, new PropertyChangedEventArgs(name));
            }
        }

        protected override void OnMouseDoubleClick(MouseEventArgs e)
        {
            if (e.Button == System.Windows.Forms.MouseButtons.Left)
            {
                if (IdControllo == 0)
                {
                    int dim = Dimension;

                    using (ConfiguratoreTasto cfg = new ConfiguratoreTasto(this))
                    {
                        cfg.ShowDialog();

                        if (dim != Dimension)
                        {
                            OnPropertyChanged("Dimensione");
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
        protected override void OnMouseDown(MouseEventArgs mevent)
        {
            _startPt = mevent.Location;
            if (mevent.Clicks == 2)
                OnMouseDoubleClick(mevent);
        }
        protected override void OnMouseMove(MouseEventArgs mevent)
        {
            if (mevent.Button == System.Windows.Forms.MouseButtons.Left && Math.Pow(mevent.Location.X - _startPt.X, 2) + Math.Pow(mevent.Location.Y - _startPt.Y, 2) > Math.Pow(SystemInformation.DragSize.Height, 2))
            {
                DoDragDrop(this, DragDropEffects.Move);
            }
        }
        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            BackColor = Color.FromKnownColor(KnownColor.ControlDark);
        }
        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            BackColor = Color.FromKnownColor(KnownColor.Control);
        }

        protected override void OnPaint(PaintEventArgs pe)
        {
            base.OnPaint(pe);
            if (!Enabled)
            {
                

                //ForeColor = System.Drawing.Color.Gray;

                //ColorMatrix matrix = new ColorMatrix(new float[][]{
                //    new float[] {0.299f, 0.299f, 0.299f, 0, 0},
                //    new float[] {0.587f, 0.587f, 0.587f, 0, 0},
                //    new float[] {0.114f, 0.114f, 0.114f, 0, 0},
                //    new float[] {     0,      0,      0, 1, 0},
                //    new float[] {     0,      0,      0, 0, 0}
                //});

                //Image image = (Bitmap)Image.Clone();
                //ImageAttributes attributes = new ImageAttributes();
                //attributes.SetColorMatrix(matrix);
                //using (Graphics graphics = Graphics.FromImage(image))
                //{
                //    graphics.DrawImage(
                //        image,
                //        new Rectangle(0, 0, image.Width, image.Height),
                //        0,
                //        0,
                //        image.Width,
                //        image.Height,
                //        GraphicsUnit.Pixel,
                //        attributes
                //    );
                //}
                //Image = image;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (!Utility.Refreshing && ConfiguratoreRibbon.ControlliUtilizzati.Contains(IdControllo))
                ConfiguratoreRibbon.GruppoControlloCancellati.Add(ConfiguratoreRibbon.GruppoControlloUtilizzati[ConfiguratoreRibbon.ControlliUtilizzati.IndexOf(IdControllo)]);

            base.Dispose(disposing);
        }


        private Image MakeGrayscale(Image original)
        {
            Image newBitmap = new Bitmap(original.Width, original.Height);
            Graphics g = Graphics.FromImage(newBitmap);
            ColorMatrix colorMatrix = new ColorMatrix(
                new float[][] 
            {
                new float[] {.299f, .299f, .299f, 0, 0},
                new float[] {.587f, .587f, .587f, 0, 0},
                new float[] {.114f, .114f, .114f, 0, 0},
                new float[] {0, 0, 0, 1, 0},
                new float[] {0, 0, 0, 0, 1}
            });

            ImageAttributes attributes = new ImageAttributes();
            attributes.SetColorMatrix(colorMatrix);
            g.DrawImage(
                original,
                new Rectangle(0, 0, original.Width, original.Height),
                0, 0, original.Width, original.Height,
                GraphicsUnit.Pixel, attributes);

            g.Dispose();
            return newBitmap;
        }
    }
}
