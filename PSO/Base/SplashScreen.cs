using System;
using System.Threading;
using System.Windows.Forms;

namespace Iren.PSO.Base
{
    public partial class SplashScreen : Form
    {
        private const int CS_DROPSHADOW = 0x20000;
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams cp = base.CreateParams;
                cp.ClassStyle |= CS_DROPSHADOW;
                return cp;
            }
        }

        private delegate void ShowDelegate();
        private delegate void CloseDelegate();
        private delegate void UpdateStatusDelegate(string status);

        public SplashScreen()
        {
            InitializeComponent();
        }

        public void ShowSplashScreen()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new ShowDelegate(ShowSplashScreen));
                return;
            }
            base.Show();
            if (!this.IsDisposed)
                Application.Run(this);
        }
        public void CloseSplashScreen()
        {
            if (InvokeRequired)
            {
                BeginInvoke(new CloseDelegate(CloseSplashScreen));
                return;
            }

            base.Close();
        }

        public void UdpateStatusText(string status)
        {
            if (InvokeRequired)
            {
                BeginInvoke(new UpdateStatusDelegate(UdpateStatusText), status);
                return;
            }

            if (status.Length > 70)
                status = status.Substring(0, 67) + " ...";
            if(status != lbText.Text)
                lbText.Text = status;
        }

        private void lbText_SizeChanged(object sender, EventArgs e)
        {
            double width = lbText.Width;
            lbText.Left = (int)Math.Round((panelAll.Width / 2) - (width / 2));
        }

        static SplashScreen sf = null;
        static Thread _splashthread;

        public new static void Show()
        {
            if (sf == null)
            {
                sf = new SplashScreen();
                _splashthread = new Thread(new ThreadStart(sf.ShowSplashScreen));
                _splashthread.IsBackground = true;

                _splashthread.Start();
            }
        }
        public new static void Close()
        {
            if (sf != null)
            {
                sf.CloseSplashScreen();
                sf = null;
            }
        }
        public static void UpdateStatus(string status)
        {
            if (sf != null)
            {
                //evito cross thread exception
                try { sf.UdpateStatusText(status); }
                catch { }
            }
        }
    }
}
