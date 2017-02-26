using Iren.PSO.Base;
using System;
using System.Collections;
using System.Configuration;
using System.Data;
using System.Deployment.Application;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Iren.PSO.Launcher
{
    public class LDaemon : ApplicationContext, IDisposable
    {
        public int IdUtente 
        {
            get;
            private set;
        }

        private NotifyIcon trayIcon;
        private ContextMenuStrip trayIconContextMenu;
        private ImageList imageList;
        private LForm launcherForm;

        private static Excel.Application _xlApp;

        private bool disposed = false;

        public LDaemon()
        {
#if !DEBUG
            if (ApplicationDeployment.IsNetworkDeployed && ApplicationDeployment.CurrentDeployment.IsFirstRun)
            {
                string startupPath = Environment.GetFolderPath(Environment.SpecialFolder.Startup);
                startupPath = Path.Combine(startupPath, Process.GetCurrentProcess().ProcessName) + ".appref-ms";

                if (!File.Exists(startupPath))
                {
                    string allProgramsPath = Environment.GetFolderPath(Environment.SpecialFolder.Programs);
                    string shortcutPath = Path.Combine(allProgramsPath, Process.GetCurrentProcess().ProcessName, Process.GetCurrentProcess().ProcessName) + ".appref-ms";
                    File.Copy(shortcutPath, startupPath);
                }
            }
#endif
            InitializeComponent();
            Initialize();

            trayIcon.Visible = true;
        }

        ~LDaemon()
        {
            Dispose(false);
        }

        private void Initialize()
        {
            DataBase.CreateNew(ConfigurationManager.AppSettings["DB"], false);
            bool error = false;
            if (!InitializeUser())
                error = true;
            else if (!InitializeApplication())
                error = true;

            if (error) 
            {
                InitializeErrorMenu();
            }
            else
            {
                trayIcon.DoubleClick += OpenLauncherForm;
            }

            DataBase.Close();
        }
        
        private void InitializeComponent()
        {
            imageList = new ImageList();
            imageList.ImageSize = new System.Drawing.Size(32, 32);

            var resourceSet = PSO.Base.Properties.Resources.ResourceManager.GetResourceSet(CultureInfo.InstalledUICulture, true, true);
            var imgs =
                from r in resourceSet.Cast<DictionaryEntry>()
                where r.Value is Image
                select r;

            foreach (var img in imgs)
                imageList.Images.Add(img.Key as string, img.Value as Image);

            trayIcon = new NotifyIcon();

            trayIcon.Text = "PSO - Pianificazione Strategica delle Offerte";

            trayIcon.Icon = Properties.Resources.PSO;
            

            trayIconContextMenu = new ContextMenuStrip();
            trayIconContextMenu.Name = "trayIconContextMenu";
            trayIconContextMenu.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            trayIconContextMenu.ImageScalingSize = new System.Drawing.Size(28, 28);
            
            trayIcon.ContextMenuStrip = trayIconContextMenu;
        }

        private void OpenLauncherForm(object sender, EventArgs e)
        {
            if (launcherForm == null)
             launcherForm = new LForm(trayIconContextMenu);

            launcherForm.Show();            
        }

        private bool InitializeUser()
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, "@CodUtenteWindows=" + Environment.UserName);

            if (dtUtente != null && dtUtente.Rows.Count > 0)
            {
                IdUtente = (int)dtUtente.Rows[0]["IdUtente"];
                return true;
            }
            else
            {
                trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                trayIcon.BalloonTipText = "Sembra che l'utente non sia configurato per PSO. Contattare l'amministratore.";
                trayIcon.BalloonTipTitle = "Attenzione";
                return false;
            }

            
        }
        private bool InitializeApplication()
        {
            trayIconContextMenu.SuspendLayout();
            DataTable dtApplicazioni = DataBase.Select(DataBase.SP.APPLICAZIONE, "@IdApplicazione=0;@IdUtente="+IdUtente);
            if (dtApplicazioni != null && dtApplicazioni.Rows.Count > 0)
            {
                trayIconContextMenu.Items.Clear();
                trayIconContextMenu.ImageList = imageList;

                for (int i = 0; i < dtApplicazioni.Rows.Count; i++)
                {
                    DataTable dtControllo = DataBase.Select(DataBase.SP.RIBBON.CONTROLLO_APPLICAZIONE, "@IdApplicazione=" + dtApplicazioni.Rows[i]["IdApplicazione"]);

                    if (dtControllo != null && dtControllo.Rows.Count > 0)
                    {
                        if (!trayIconContextMenu.Items.ContainsKey(dtControllo.Rows[0]["Nome"].ToString()))
                        {
                            ToolStripItem item = trayIconContextMenu.Items.Add(dtControllo.Rows[0]["Label"].ToString());
                            item.ImageKey = dtControllo.Rows[0]["Immagine"].ToString();
                            item.Name = dtControllo.Rows[0]["Nome"].ToString();
                            item.Tag = dtControllo.Rows[0]["IdApplicazione"];
                            item.Margin = new Padding(0, 0, 0, 0);
                            item.Padding = new Padding(0, 0, 0, 0);
                            item.Click += StartApplication;
                        }
                    }
                }

                trayIconContextMenu.ResumeLayout(false);
                return true;
            }
            else
            {
                trayIconContextMenu.ResumeLayout(false);
                
                trayIcon.BalloonTipIcon = ToolTipIcon.Error;
                trayIcon.BalloonTipText = "In seguito ad un errore, non è stato possibile caricare la lista delle applicazioni configurate. Contattare l'amministratore.";
                trayIcon.BalloonTipTitle = "Attenzione";
                
                return false;
            }
        }
        private void InitializeErrorMenu()
        {
            if (!trayIconContextMenu.Items.ContainsKey("Ricarica"))
            {
                trayIcon.DoubleClick -= OpenLauncherForm;
                trayIconContextMenu.Items.Clear();

                ToolStripItem ricarica = trayIconContextMenu.Items.Add("Ricarica launcher");
                ricarica.Name = "Ricarica";
                ricarica.Margin = new Padding(0, 0, 0, 0);
                ricarica.Padding = new Padding(0, 0, 0, 0);
                ricarica.Click += RicaricaApplicazione;
            }
        }

        private void RicaricaApplicazione(object sender, EventArgs e)
        {
            Initialize();
        }

        public static void StartApplication(object sender, EventArgs e)
        {
            int idApplicazione = (int)GetTag(sender);

            Excel.Workbooks wbs = null;
            try
            {
                wbs = _xlApp.Workbooks;
            }
            catch
            {
                _xlApp = new Excel.Application();
                _xlApp.WorkbookBeforeClose += CheckIfLast;
            }
            finally
            {
                if (wbs != null) Marshal.ReleaseComObject(wbs);
                wbs = null;
            }

            Workbook.AvviaApplicazione(_xlApp, idApplicazione);
        }

        private static void CheckIfLast(Excel.Workbook Wb, ref bool Cancel)
        {
            try
            {
                Excel.Workbooks wbs = _xlApp.Workbooks;

                int count = wbs.OfType<Excel.Workbook>()
                    .Count(wb => !wb.Equals(Wb) && wb.Windows[1].Visible);

                if (count <= 1)
                {
                    _xlApp.Quit();
                    Marshal.ReleaseComObject(_xlApp);
                    _xlApp = null;
                }
            }
            catch { }

        }

        private static object GetTag(object sender)
        {
            Button button = sender as Button;
            ToolStripItem tsi = sender as ToolStripItem;

            if (button != null)
                return button.Tag;
            if (tsi != null)
                return tsi.Tag;

            throw new ArgumentException("Unexpected sender");
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposed)
                return;
            
            if (disposing)
            {

            }

            try
            {
                foreach (Excel.Workbook wb in _xlApp.Workbooks)
                {
                    wb.Close(false);
                    Marshal.ReleaseComObject(wb);
                }
                _xlApp.Quit();
                Marshal.ReleaseComObject(_xlApp);
            }
            catch { }
            GC.WaitForPendingFinalizers();
            GC.Collect();

            disposed = true;
        }
    }
}
