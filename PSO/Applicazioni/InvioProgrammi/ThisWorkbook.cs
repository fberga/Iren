using Iren.PSO.Base;
using Microsoft.Office.Tools.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System;
using System.Data;
using System.Deployment.Application;
using System.Reflection;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

// ************************************************************* INVIO PROGRAMMI ************************************************************* //

namespace Iren.PSO.Applicazioni
{
    public partial class ThisWorkbook : IPSOThisWorkbook
    {
        #region Proprietà

        public System.Version Version
        {
            get
            {
                try
                {
                    return ApplicationDeployment.CurrentDeployment.CurrentVersion;
                }
                catch (Exception)
                {
                    return Assembly.GetExecutingAssembly().GetName().Version;
                }
            }
        }

        public Worksheet Main { get { return Globals.Main.Base; } }
        public Worksheet Log { get { return Globals.Log.Base; } }
        public new Excel.Worksheet ActiveSheet { get { return (Excel.Worksheet)base.ActiveSheet; } }
        public new Excel.Sheets Sheets { get { return base.Sheets; } }

        public int IdApplicazione { get { return idApplicazione; } set { idApplicazione = value; } }
        public int IdUtente { get { return idUtente; } set { idUtente = value; } }
        public string NomeUtente { get { return nomeUtente; } set { nomeUtente = value; } }
        public DateTime DataAttiva { get { return dataAttiva; } set { dataAttiva = value; } }
        public string Ambiente { get { return ambiente; } set { ambiente = value; } }
        public string Pwd { get { return password; } }
        public int IdStagione { get { return idStagione; } set { idStagione = value; } }

        public DataSet RepositoryDataSet { get { return repositoryDataSet; } }
        public DataTable LogDataTable { get { return logDataTable; } set { logDataTable = value; } }
        public DataSet RibbonDataSet { get { return ribbonDataSet; } }

        #endregion

        #region Cached Attribute

        [CachedAttribute()]
        public int idApplicazione = 2;
        [CachedAttribute()]
        public int idUtente = -1;
        [CachedAttribute()]
        public string nomeUtente = string.Empty;
        [CachedAttribute()]
        public DateTime dataAttiva = DateTime.Today;
        [CachedAttribute()]
        public string ambiente = Simboli.PROD;
        [CachedAttribute()]
        public DataSet repositoryDataSet = new DataSet();
        [CachedAttribute()]
        public DataTable logDataTable = new DataTable();
        [CachedAttribute()]
        public DataSet ribbonDataSet = new DataSet();
        [CachedAttribute()]
        public string password = "8176";
        [CachedAttribute()]
        public int idStagione = -1;

        #endregion

        #region Codice generato dalla finestra di progettazione di VSTO

        /// <summary>
        /// Metodo richiesto per il supporto della finestra di progettazione - non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InternalStartup()
        {
            this.BeforeClose += new Microsoft.Office.Interop.Excel.WorkbookEvents_BeforeCloseEventHandler(this.ThisWorkbook_BeforeClose);
            this.SheetSelectionChange += new Microsoft.Office.Interop.Excel.WorkbookEvents_SheetSelectionChangeEventHandler(Handler.CellClick);
            this.Startup += new System.EventHandler(this.ThisWorkbook_Startup);
        }

        #endregion

        private void ThisWorkbook_Startup(object sender, System.EventArgs e)
        {
            Application.ScreenUpdating = false;
#if DEBUG
            //TODO Ripristinare DEV
            ambiente = Simboli.TEST;
#else
            /*********************** Modifica per ambient di Test *********************/
            //ambiente = Simboli.TEST;  //TODO Commentare per passaggio in produzione
#endif
            PSO.Base.Workbook.StartUp(this);
            Globals.Ribbons.GetRibbon<ToolsExcelRibbon>().InitRibbon();
            Application.ScreenUpdating = true;
        }
        private void ThisWorkbook_BeforeClose(ref bool Cancel)
        {
            PSO.Base.Workbook.Close();
            Saved = true;
        }
    }
}