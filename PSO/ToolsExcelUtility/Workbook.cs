using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using Iren.ToolsExcel.UserConfig;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace Iren.ToolsExcel.Utility
{
    public class Workbook
    {
        #region Variabili

        /// <summary>
        /// Il workbook.
        /// </summary>
        protected static IToolsExcelThisWorkbook _wb;
        /// <summary>
        /// Flag che viene utilizzato per bloccare l'evento SheetSelectionChange quando la selezione è cambiata dal pannello laterale dei check.
        /// </summary>
        public static bool FromErrorPane = false;

        public static IWin32Window Window;

        #endregion

        #region Proprietà

        /// <summary>
        /// L'oggetto Excel del Workbook per accedere a tutti gli handler e proprietà. (Read only)
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Workbook WB { get { return _wb.Base; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio Main (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Main { get { return (Excel.Worksheet)_wb.Sheets["Main"]; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio di Log (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Log { get { return (Excel.Worksheet)_wb.Sheets["Log"]; } }
        /// <summary>
        /// Scorciatoia per accedere al foglio attivo.
        /// </summary>
        public static Excel.Worksheet ActiveSheet { get { return (Excel.Worksheet)_wb.ActiveSheet; } }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Application di Excel.
        /// </summary>
        public static Excel.Application Application { get { return _wb.Application; } }
        /// <summary>
        /// Lista di tutti i fogli che rappresentano una Categoria sul DB (non fanno parte i fogli Log, Main, MSDx). I fogli non sono indicizzati per nome, solo per indice.
        /// </summary>
        public static IList<Excel.Worksheet> CategorySheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name != "Log" && ws.Name != "Main" && !ws.Name.StartsWith("MSD")).ToList(); } }
        /// <summary>
        /// Lista di tutti fogli indicizzati per nome.
        /// </summary>
        public static Excel.Sheets Sheets { get { return WB.Sheets; } }
        /// <summary>
        /// Lista dei folgi MSDx utile solo in Invio Programmi.
        /// </summary>
        public static IList<Excel.Worksheet> MSDSheets { get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name.StartsWith("MSD")).ToList(); } }
        /// <summary>
        /// La versione dell'applicazione.
        /// </summary>
        public static System.Version WorkbookVersion { get { return _wb.Version; } }
        /// <summary>
        /// La versione della classe Core
        /// </summary>
        public static System.Version CoreVersion { get { return DataBase.Versione; } }
        /// <summary>
        /// La versione della classe Base.
        /// </summary>
        public static System.Version BaseVersion { get { return Assembly.GetExecutingAssembly().GetName().Version; } }
        /// <summary>
        /// Flag per attivare/disattivare il refresh dello schermo.
        /// </summary>
        public static bool ScreenUpdating { get { return Application.ScreenUpdating; } set { Application.ScreenUpdating = value; } }

        public static string Password { get { return _wb.Pwd; } }
        public static string NomeUtente { get { return _wb.NomeUtente; } }
        public static string Ambiente 
        { 
            get { return _wb.Ambiente; } 
            set 
            {
                _wb.Ambiente = value; 
                Handler.ChangeAmbiente(value); 
            } 
        }
        public static string Stagione 
        {
            get
            {
                return Repository[DataBase.TAB.STAGIONE].AsEnumerable()
                    .Where(r => r["IdTipologiaStagione"].Equals(IdStagione))
                    .Select(r => r["DesTipologiaStagione"].ToString())
                    .FirstOrDefault();
            }
        }
        public static string Mercato 
        {
            get
            {
                return Repository[DataBase.TAB.MERCATI].AsEnumerable()
                    .Where(r => r["IdApplicazioneMercato"].Equals(IdApplicazione))
                    .Select(r => r["DesMercato"].ToString())
                    .FirstOrDefault();
            }
        }

        public static int IdApplicazione 
        { 
            get 
            { 
                return _wb.IdApplicazione; 
            } 
            set 
            { 
                _wb.IdApplicazione = value;
                if (DataBase.IsInitialized) DataBase.IdApplicazione = value;
            } 
        }
        public static int IdUtente 
        { 
            get 
            { 
                return _wb.IdUtente; 
            }
            set
            {
                _wb.IdUtente = value;
                if (DataBase.IsInitialized) DataBase.IdUtente = value;
            }
        }
        public static DateTime DataAttiva
        {
            get
            {
                return _wb.DataAttiva;
            }
            set
            {
                _wb.DataAttiva = value;
                if (DataBase.IsInitialized) DataBase.DataAttiva = value;
            }
        }
        
        public static int IdStagione { 
            get { return _wb.IdStagione; } 
            set 
            { 
                _wb.IdStagione = value;
                
                var sheetPrevisione = Sheets.Cast<Excel.Worksheet>()
                    .Where(s => s.Name == "Previsione")
                    .FirstOrDefault();

                if (sheetPrevisione != null)
                {
                    DefinedNames definedNames = new DefinedNames("Previsione");
                    DateTime dataFine = Utility.Workbook.DataAttiva.AddDays(Struct.intervalloGiorni);
                    Range rng = definedNames.Get("CT_TORINO", "STAGIONE", Utility.Date.SuffissoDATA1, Utility.Date.GetSuffissoOra(1)).Extend(colOffset: Utility.Date.GetOreIntervallo(dataFine));
                    sheetPrevisione.Range[rng.ToString()].Value = value;
                }
            }
        }

        public static Utility.Repository Repository { get; private set; }
        public static DataTable LogDataTable { get { return _wb.LogDataTable; } }

        #endregion

        #region Metodi

        /// <summary>
        /// Carica dal DB i dati riguardanti le proprietà dell'applicazione che si trovano nella tabella APPLICAZIONE. Assegna alle variabili globali di applicazione i valori.
        /// </summary>
        public static void AggiornaParametriApplicazione()
        {
            DataRow r = Workbook.Repository.CaricaApplicazione(_wb.IdApplicazione);
            if (r == null)
                throw new ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");

            Simboli.nomeApplicazione = r["DesApplicazione"].ToString();
            Struct.intervalloGiorni = (r["IntervalloGiorniEntita"] is DBNull ? 0 : (int)r["IntervalloGiorniEntita"]);
            Struct.tipoVisualizzazione = r["TipoVisualizzazione"] is DBNull ? "O" : r["TipoVisualizzazione"].ToString();
            Struct.visualizzaRiepilogo = r["VisRiepilogo"] is DBNull ? true : r["VisRiepilogo"].Equals("1");
            Struct.cell.width.empty = double.Parse(r["ColVuotaWidth"].ToString());
            Struct.cell.width.dato = double.Parse(r["ColDatoWidth"].ToString());
            Struct.cell.width.entita = double.Parse(r["ColEntitaWidth"].ToString());
            Struct.cell.width.informazione = double.Parse(r["ColInformazioneWidth"].ToString());
            Struct.cell.width.unitaMisura = double.Parse(r["ColUMWidth"].ToString());
            Struct.cell.width.parametro = double.Parse(r["ColParametroWidth"].ToString());
            Struct.cell.width.jolly1 = double.Parse(r["ColJolly1Width"].ToString());
            Struct.cell.height.normal = double.Parse(r["RowHeight"].ToString());
            Struct.cell.height.empty = double.Parse(r["RowVuotaHeight"].ToString());
        }
        /// <summary>
        /// Aggiorna i label indicanti lo stato dei Database in seguito ad un cambio di stato.
        /// </summary>
        public static void AggiornaLabelStatoDB()
        {
            //disabilito l'aggiornamento in caso di modifica dati... lo ripeto alla chiusura in caso
            if (!Simboli.ModificaDati)
            {
                bool isProtected = true;
                try
                {
                    Workbook.WB.Application.ScreenUpdating = false;
                    isProtected = Main.ProtectContents;

                    if (isProtected)
                        Main.Unprotect(Utility.Workbook.Password);


                    Riepilogo main = new Riepilogo(Utility.Workbook.Main);

                    if (DataBase.OpenConnection())
                    {
                        Dictionary<Core.DataBase.NomiDB, ConnectionState> stato = DataBase.StatoDB;
                        Simboli.SQLServerOnline = stato[Core.DataBase.NomiDB.SQLSERVER] == ConnectionState.Open;
                        Simboli.ImpiantiOnline = stato[Core.DataBase.NomiDB.IMP] == ConnectionState.Open;
                        Simboli.ElsagOnline = stato[Core.DataBase.NomiDB.ELSAG] == ConnectionState.Open;

                        main.UpdateData();

                        DataBase.CloseConnection();
                    }
                    else
                    {
                        Simboli.SQLServerOnline = false;
                        Simboli.ImpiantiOnline = false;
                        Simboli.ElsagOnline = false;

                        main.RiepilogoInEmergenza();
                    }

                    if (isProtected)
                        Main.Protect(Utility.Workbook.Password);
                }
                catch { }

                //lo faccio a parte perché se andasse in errore prima deve almeno provare a riattivare lo screen updating!!!
                try { Workbook.WB.Application.ScreenUpdating = true; }
                catch { }
            }
        }
        /// <summary>
        /// Handler per il PropertyChanged della classe Core.DataBase. Attiva l'aggiornamento dei label.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void StatoDBChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e)
        {
            AggiornaLabelStatoDB();
        }        
        /// <summary>
        /// Restituisce lo UserConfigElement collegato alla chiave configKey nella sezione usrConfig (da non confondere con appSettings).
        /// </summary>
        /// <param name="configKey">Chiave.</param>
        /// <returns>Restituisce l'elemento ricercato.</returns>
        public static UserConfiguration GetUsrConfiguration()
        {
            return (UserConfiguration)ConfigurationManager.GetSection("usrConfig");
        }
        public static UserConfigElement GetUsrConfigElement(string configKey)
        {
            var settings = GetUsrConfiguration();
            return (UserConfigElement)settings.Items[configKey];
        }

        /// <summary>
        /// Restituisce un array con le tre componenti intere Red Green Blue a partire da una stringa suddivisa con un separatore sep. Non ha una gestione di errore, se il parser non riesce ad interpretare la stringa, va in errore.
        /// </summary>
        /// <param name="rgb">Stringa nel formato RRR[sep]GGG[sep]BBB.</param>
        /// <param name="sep">Separatore.</param>
        /// <returns>Restituisce le tre componenti trovate.</returns>
        public static int[] GetRGBFromString(string rgb, char sep = ';')
        {
            string[] rgbComp = rgb.Split(sep);

            return new int[] { int.Parse(rgbComp[0]), int.Parse(rgbComp[1]), int.Parse(rgbComp[2]) };
        }

        #region AppSettings
        /// <summary>
        /// Quando il file è criptato capita che senza il refresh vada in errore.
        /// </summary>
        /// <param name="key">La chiave da ricercare nella sezione appSettings</param>
        /// <returns>Restituisce la stringa del Value.</returns>
        public static string AppSettings(string key)
        {
            try
            {
                return ConfigurationManager.AppSettings[key];
            }
            catch
            {
                ConfigurationManager.RefreshSection("appSettings");
                return ConfigurationManager.AppSettings[key];
            }
        }
        /// <summary>
        /// Assegna value al valore della chiave key della sesione appSettings del file di configurazione. Alla fine dell'operazione esegue il refresh della sezione in modo da forzare la riscrittura su disco dei nuovi valori.
        /// </summary>
        /// <param name="key">Chiave da modificare.</param>
        /// <param name="value">Nuovo valore da assegnare.</param>
        public static void ChangeAppSettings(string key, string value)
        {
            var config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            config.AppSettings.Settings[key].Value = value;
            config.Save(ConfigurationSaveMode.Minimal);
            ConfigurationManager.RefreshSection("appSettings");
        }
        #endregion

        #region Init
        private static void InitLog()
        {
            DataTable dtLog = DataBase.Select(DataBase.SP.APPLICAZIONE_LOG);
            if (dtLog != null)
            {
                dtLog.TableName = DataBase.TAB.LOG;
                if (_wb.LogDataTable != null)
                    _wb.LogDataTable.Merge(dtLog);
                else
                    _wb.LogDataTable = dtLog;

                //_wb.LogDataTable.Tables[DataBase.TAB.LOG].DefaultView.Sort = "Data DESC";

            }
        }

        public static void GetUtente(out int idUtente, out string nomeUtente)
        {
            DataTable dtUtente = DataBase.Select(DataBase.SP.UTENTE, "@CodUtenteWindows=" + Environment.UserName);
            if (dtUtente != null && dtUtente.Rows.Count > 0)
            {
                idUtente = (int)dtUtente.Rows[0]["IdUtente"];
                nomeUtente = dtUtente.Rows[0]["Nome"].ToString();
            } 
            else 
            {
                idUtente = 0;
                nomeUtente = "NON CONFIGURATO";
            }
        }
        private static void InitUtente()
        {
            int idUtente;
            string nomeUtente;

            GetUtente(out idUtente, out nomeUtente);

            _wb.IdUtente = idUtente;
            _wb.NomeUtente = nomeUtente;
        }

        private static bool Initialize()
        {
            if (DataBase.OpenConnection())
            {
                InitUtente();

                DataBase.SetParameters(_wb.DataAttiva, _wb.IdUtente, _wb.IdApplicazione);
                //Se non ci sono tabelle, inizializzo il repository allo stato attuale
                if (Workbook.Repository.TablesCount == 0)
                {
                    Aggiorna aggiorna = new Aggiorna();
                    aggiorna.Struttura(avoidRepositoryUpdate: false);
                }
                else
                    Workbook.AggiornaParametriApplicazione();

                if (Workbook.Repository.Applicazione != null)
                {
                    Simboli.rgbSfondo = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BackColorApp"].ToString());
                    Simboli.rgbTitolo = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BackColorFrameApp"].ToString());
                    Simboli.rgbLinee = Workbook.GetRGBFromString(Workbook.Repository.Applicazione["BorderColorApp"].ToString());
                }

                InitLog();

                return false;
            }
            else //Emergenza
            {
                if (Repository.TablesCount == 0)
                {
                    System.Windows.Forms.MessageBox.Show("Il foglio non è inizializzato e non c'è connessione ad DB... Impossibile procedere! L'applicazione verrà chiusa.", "ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);

                    _wb.Base.Close();
                    return false;
                }

                DataBase.DB.SetParameters(dataAttiva: _wb.DataAttiva);
                Simboli.nomeApplicazione = Workbook.Repository.Applicazione["DesApplicazione"].ToString();
                Struct.intervalloGiorni = Workbook.Repository.Applicazione["IntervalloGiorniEntita"] is DBNull ? 0 : (int)Workbook.Repository.Applicazione["IntervalloGiorniEntita"];
                Struct.visualizzaRiepilogo = Workbook.Repository.Applicazione["VisRiepilogo"] is DBNull ? true : Workbook.Repository.Applicazione["VisRiepilogo"].Equals("1");

                return true;
            }
        }


        public static bool Update()
        {
            //UPDATE
            string updatePath = Path.Combine(_wb.Path, "UPDATE");
            if (Directory.Exists(updatePath) && Directory.GetFiles(updatePath, _wb.Name).Any())
            {
                string name = _wb.Name;
                string fullName = _wb.FullName;
                _wb.Base.SaveAs(Path.Combine(_wb.Path, "old_" + _wb.Name));
                File.Copy(Path.Combine(updatePath, name), fullName, true);
                File.Delete(Path.Combine(updatePath, name));
                Application.Workbooks.Open(fullName);
                _wb.Base.Windows[1].Visible = false;
                return true;
            }
            else
            {
                Window = new Win32Window(new IntPtr(Workbook.Application.Hwnd));
                try { if (File.Exists(Path.Combine(_wb.Path, "old_" + _wb.Name))) File.Delete(Path.Combine(_wb.Path, "old_" + _wb.Name)); }
                catch { }
            }

            return false;
        }
        public static void ControlloAreeDiRete()
        {
            //controllo le aree di rete (se presenti)
            var usrConfig = GetUsrConfiguration();
            Dictionary<string, string> pathNonDisponibili = new Dictionary<string, string>();
            foreach (UserConfigElement ele in usrConfig.Items)
            {
                if (ele.Type == UserConfigElement.ElementType.path)
                {
                    string pathStr = ele.Value;

                    try { System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(pathStr); }
                    catch { pathNonDisponibili.Add(ele.Desc, pathStr); }
                }
            }
            //segnalo all'utente l'impossibilità di accedere alle aree di rete
            if (pathNonDisponibili.Count > 0)
            {
                string paths = "\n";
                foreach (var kv in pathNonDisponibili)
                    paths += " - " + kv.Key + " : '" + kv.Value + "'\n";

                System.Windows.Forms.MessageBox.Show("I path seguenti non sono raggiungibili o non presentano privilegi di scrittura:" + paths, Simboli.nomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        public static void StartUp(IToolsExcelThisWorkbook wb)
        {
            _wb = wb;
            Repository = new Utility.Repository(wb);

            DataBase.CreateNew(Ambiente);
            if (!Update())
            {
                Application.Iteration = true;
                Application.MaxIterations = 100;
                Application.EnableEvents = false;

                Style.StdStyles();

                ControlloAreeDiRete();

                foreach (Excel.Worksheet ws in CategorySheets)
                {
                    ws.Activate();
                    ws.Range["A1"].Select();
                    Application.ActiveWindow.ScrollRow = 1;
                }

                Main.Select();
                Application.WindowState = Excel.XlWindowState.xlMaximized;
                
                ScreenUpdating = false;
                
                Initialize();
                Application.EnableEvents = true;
            }
        }
        #endregion

        #region Log
        public static void InsertLog(Core.DataBase.TipologiaLOG logType, string message)
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            if (prot) log.Unprotect(Password);
            DataBase db = new DataBase();
            db.InsertLog(logType, message);
            if (prot) log.Protect(Password);
        }
        public static void RefreshLog()
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            if (prot) log.Unprotect(Password);
            DataBase db = new DataBase();
            db.RefreshLog();
            if (prot) log.Protect(Password);
        }
        #endregion

        #region Close

        public static void Close()
        {
            if (Workbook.Repository != null)
            {
                Simboli.EmergenzaForzata = false;
                Application.ScreenUpdating = false;
                if (WB.Application.DisplayDocumentActionTaskPane)
                    WB.Application.DisplayDocumentActionTaskPane = false;

                Main.Select();
                if (Simboli.ModificaDati)
                {
                    Sheet.Protected = false;
                    Simboli.ModificaDati = false;
                    Sheet.AbilitaModifica(false);
                    Sheet.SalvaModifiche();
                    Sheet.Protected = true;
                }
                DataBase.SalvaModificheDB();
                InsertLog(Core.DataBase.TipologiaLOG.LogAccesso, "Log off - " + Environment.UserName + " - " + Environment.MachineName);

                Application.ScreenUpdating = true;
                
                _wb.Base.Save();
                Application.DisplayAlerts = false;
            }
        }
        #endregion

        #endregion
    }
}

