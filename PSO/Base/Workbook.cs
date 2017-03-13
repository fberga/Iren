using Iren.PSO.UserConfig;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;



namespace Iren.PSO.Base
{
    public class Workbook
    {
        #region Variabili

        /// <summary>
        /// Il workbook.
        /// </summary>
        protected static IPSOThisWorkbook _wb;
        /// <summary>
        /// Flag che viene utilizzato per bloccare l'evento SheetSelectionChange quando la selezione è cambiata dal pannello laterale dei check.
        /// </summary>
        public static bool FromErrorPane = false;

        private static bool _isStoreEditEnabled = false;

        private static bool _wasInEmergency = false;

        #endregion

        #region Proprietà

        /// <summary>
        /// L'oggetto Excel del Workbook per accedere a tutti gli handler e proprietà. (Read only)
        /// </summary>
        public static Microsoft.Office.Tools.Excel.Workbook WB 
        { 
            get { return _wb.Base; } 
        }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio Main (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Main 
        { 
            get { return (Excel.Worksheet)_wb.Sheets["Main"]; } 
        }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Excel del foglio di Log (sempre presente in tutti i fogli).
        /// </summary>
        public static Excel.Worksheet Log 
        { 
            get { return (Excel.Worksheet)_wb.Sheets["Log"]; } 
        }
        /// <summary>
        /// Scorciatoia per accedere al foglio attivo.
        /// </summary>
        public static Excel.Worksheet ActiveSheet 
        { 
            get { return (Excel.Worksheet)_wb.ActiveSheet; } 
        }
        /// <summary>
        /// Scorciatoia per accedere all'oggetto Application di Excel.
        /// </summary>
        public static Excel.Application Application 
        { 
            get { return _wb.Application; } 
        }
        /// <summary>
        /// Lista di tutti i fogli che rappresentano una Categoria sul DB (non fanno parte i fogli Log, Main, MSDx). I fogli non sono indicizzati per nome, solo per indice.
        /// </summary>
        public static IList<Excel.Worksheet> CategorySheets 
        { 
            get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name != "Log" && ws.Name != "Main" && !ws.Name.StartsWith("MSD")).ToList(); } 
        }
        /// <summary>
        /// Lista di tutti fogli indicizzati per nome.
        /// </summary>
        public static Excel.Sheets Sheets 
        { 
            get { return WB.Sheets; } 
        }
        /// <summary>
        /// Lista dei folgi MSDx utile solo in Invio Programmi.
        /// </summary>
        public static IList<Excel.Worksheet> MSDSheets 
        { 
            get { return _wb.Sheets.Cast<Excel.Worksheet>().Where(ws => ws.Name.StartsWith("MSD")).ToList(); } 
        }
        /// <summary>
        /// La versione dell'applicazione.
        /// </summary>
        public static System.Version WorkbookVersion 
        { 
            get { return _wb.Version; } 
        }
        /// <summary>
        /// La versione della classe Core
        /// </summary>
        public static System.Version CoreVersion 
        { 
            get { return Iren.PSO.Base.DataBase.Versione; } 
        }
        /// <summary>
        /// La versione della classe Base.
        /// </summary>
        public static System.Version BaseVersion 
        { 
            get { return Assembly.GetExecutingAssembly().GetName().Version; } 
        }
        /// <summary>
        /// Flag per attivare/disattivare il refresh dello schermo.
        /// </summary>
        public static bool ScreenUpdating 
        { 
            get { return Application.ScreenUpdating; } set { Application.ScreenUpdating = value; } 
        }
        /// <summary>
        /// Cached password per bloccare i fogli.
        /// </summary>
        public static string Password 
        { get { return _wb.Pwd; } }
        /// <summary>
        /// Cached nome utente.
        /// </summary>
        public static string NomeUtente 
        { get { return _wb.NomeUtente; } }
        /// <summary>
        /// Cached ambiente. Simboli.[DEV|TEST|PROD]
        /// </summary>
        public static string Ambiente 
        { 
            get { return _wb.Ambiente; } 
            set 
            {
                _wb.Ambiente = value; 
                Handler.ChangeAmbiente(value); 
            } 
        }
        /// <summary>
        /// Sigla tipologia stagione.
        /// </summary>
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
        /// <summary>
        /// Sigla tipologia mercato.
        /// </summary>
        
        //09/02/2017 MOD gestione mercati MI
        private static string _mercatoMI;
        public static string Mercato 
        {
            get
            {
                return _wb.Mercato;
            }
            set
            {
                _wb.Mercato = value;
            }
        }
        /// <summary>
        /// Cached IdApplicazione.
        /// </summary>
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
        /// <summary>
        /// Cached IdUtente.
        /// </summary>
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
        /// <summary>
        /// Cached data attiva.
        /// </summary>
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
        /// <summary>
        /// Cached id stagione.
        /// </summary>
        public static int IdStagione { 
            get { return _wb.IdStagione; } 
            set 
            { 
                _wb.IdStagione = value;

                Handler.ScriviStagione(value);
            }
        }
        /// <summary>
        /// Repository contentente i dati salvati localmente nel foglio. Si basa sul Cached Attribute DataSet repositoryDataSet.
        /// </summary>
        public static Repository Repository 
        { get; private set; }
        /// <summary>
        /// Cached tabella con i dati di log.
        /// </summary>
        public static DataTable LogDataTable 
        { get { return _wb.LogDataTable; } }

        public static bool AbortedLoading { get; private set; }


        //parametri console
        public static bool DaConsole
        { get; private set; }
        /// <summary>
        /// Serve a stabilire se il foglio sia o no inizializzato dopo un aggiornamento o l'aggiornamento sia stato forzato dalla console
        /// </summary>
        public static bool DaAggiornare 
        { get; private set; }
        public static bool AccettaCambioData
        { get; private set; }
        public static bool RifiutaCambioData
        { get; private set; }
        public static bool AggiornaDati
        { get; private set; }

        public static bool HaAzioni
        {
            get;
            private set;
        }
        public static string[] ListaAzioni
        { get; private set; }
        public static bool HaEntita
        {
            get;
            private set;
        }
        public static string[] ListaEntita
        { get; private set; }

        #endregion

        #region Metodi

        /// <summary>
        /// Carica dal DB i dati riguardanti le proprietà dell'applicazione che si trovano nella tabella APPLICAZIONE. Assegna alle variabili globali di applicazione i valori.
        /// </summary>
        public static void AggiornaParametriApplicazione(bool avoidRepositoryUpdate) 
        {
            DataRow r;
            if (!avoidRepositoryUpdate)
                r = Workbook.Repository.CaricaApplicazione(_wb.IdApplicazione);
            else
                r = Workbook.Repository.CambiaApplicazione(_wb.IdApplicazione);

            if (r == null)
                throw new Core.ApplicationNotFoundException("L'appID inserito non ha restituito risultati.");

            Simboli.NomeApplicazione = r["DesApplicazione"].ToString();
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
                        Main.Unprotect(Workbook.Password);

                    //refresh nel caso stia passando da emergenza a normale e viceversa
                    Riepilogo main = new Riepilogo(Workbook.Main);

                    if (DataBase.OpenConnection())
                    {
                        Dictionary<PSO.Core.DataBase.NomiDB, ConnectionState> stato = DataBase.StatoDB;
                        Simboli.SQLServerOnline = stato[PSO.Core.DataBase.NomiDB.SQLSERVER] == ConnectionState.Open;
                        Simboli.ImpiantiOnline = stato[PSO.Core.DataBase.NomiDB.IMP] == ConnectionState.Open;
                        Simboli.ElsagOnline = stato[PSO.Core.DataBase.NomiDB.ELSAG] == ConnectionState.Open;

                        if (_wasInEmergency)
                            main.UpdateData();
                        
                        _wasInEmergency = false;

                        DataBase.CloseConnection();
                    }
                    else
                    {
                        Simboli.SQLServerOnline = false;
                        Simboli.ImpiantiOnline = false;
                        Simboli.ElsagOnline = false;
                        
                        if(!_wasInEmergency)
                            main.RiepilogoInEmergenza();
                        
                        _wasInEmergency = true;
                    }

                    if (isProtected)
                        Main.Protect(Workbook.Password);
                }
                catch { }

                //lo faccio a parte perché se andasse in errore prima deve almeno provare a riattivare lo screen updating!!!
                try { Workbook.WB.Application.ScreenUpdating = true; }
                catch { }
            }
        }
        /// <summary>
        /// Handler per il PropertyChanged della classe PSO.Core.DataBase. Attiva l'aggiornamento dei label.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        public static void StatoDBChanged(object sender, System.ComponentModel.PropertyChangedEventArgs e) 
        {
            AggiornaLabelStatoDB();
        }        
        /// <summary>
        /// Restituisce la sezione "UserConfig" del file di configurazione.
        /// </summary>
        /// <returns>La sezione  "UserConfig" del file di configurazione.</returns>
        public static UserConfiguration GetUsrConfiguration() 
        {
            return (UserConfiguration)ConfigurationManager.GetSection("usrConfig");
        }
        /// <summary>
        /// Restituisce lo UserConfigElement collegato alla chiave configKey nella sezione usrConfig (da non confondere con appSettings).
        /// </summary>
        /// <param name="configKey">Chiave.</param>
        /// <returns>Restituisce l'elemento ricercato.</returns>
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
        /// <summary>
        /// Funzione che lancia l'applicazione o la sua installazione nel caso in cui non sia ancora installata.
        /// </summary>
        /// <param name="xlApp">Applicazione Excel su cui lanciare il programma.</param>
        /// <param name="idApplicazione">Id dell'applicazione da avviare.</param>
        /// 
        public static void AvviaApplicazione(Excel.Application xlApp, int idApplicazione)
        {
            string file = Simboli.FileApplicazione[idApplicazione];

            //controllo se è già aperta
            bool opened = false;


            Excel._Workbook workbook = xlApp.Workbooks.OfType<Excel.Workbook>().Where(wb => wb.Name == file + ".xlsm").FirstOrDefault();

            if (workbook != null)
            {
                xlApp.Visible = true;
                xlApp.WindowState = Excel.XlWindowState.xlMaximized;
                workbook.Activate();
                opened = true;
            }

            if (!opened)
            {
#if DEBUG
                string path = @"D:\Repository\Iren\PSO\Applicazioni\" + file + @"\bin\Debug\" + file + ".xlsm";
#else
                string path = Path.Combine(Environment.ExpandEnvironmentVariables(Simboli.LocalBasePath), file + ".xlsm");
#endif
                if (!File.Exists(path))
                {                    
                    //Installazione da remoto
                    string installPath = Path.Combine(Simboli.RemoteBasePath, file, file + ".vsto");

                    var process = new System.Diagnostics.Process
                    {
                        StartInfo = new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = installPath
                            //Arguments = "dfshim.dll,ShOpenVerbApplication " + installPath
                        },
                        EnableRaisingEvents = true,
                    };
                    process.Start();
                }
                else
                {
                    xlApp.Visible = true;
                    xlApp.WindowState = Excel.XlWindowState.xlMaximized;
                    xlApp.Workbooks.Open(path);
                }
            }
            
        }


        public static void AddStdStoreEdit()
        {
            if (!_isStoreEditEnabled)
            {
                WB.SheetChange += Handler.StoreEdit;
                _isStoreEditEnabled = true;
            }
        }
        public static void RemoveStdStoreEdit()
        {
            if (_isStoreEditEnabled)
            {
                WB.SheetChange -= Handler.StoreEdit;
                _isStoreEditEnabled = false;
            }
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
        #endregion

        #region Init
        
        /// <summary>
        /// Recupera da DB i dati del log.
        /// </summary>
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
        /// <summary>
        /// Recupera dal DB l'utente in base ad Environment.UserName.
        /// </summary>
        /// <param name="idUtente">Parametro di output contenente l'id dell'utente.</param>
        /// <param name="nomeUtente">Parametro di output contenente il nome utente.</param>
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
        /// <summary>
        /// Utilizza GetUtente per settare i CachedAttribute del fogio.
        /// </summary>
        private static void InitUtente()
        {
            int idUtente;
            string nomeUtente;

            GetUtente(out idUtente, out nomeUtente);

            _wb.IdUtente = idUtente;
            _wb.NomeUtente = nomeUtente;
        }
        /// <summary>
        /// Inizializza il foglio dopo l'apertura. Restituisce false se il foglio è in emergenza, true se la condizione è normale.
        /// </summary>
        /// <returns>Restituisce false se il foglio è in emergenza, true se la condizione è normale.</returns>
        private static bool Initialize()
        {
            if (DataBase.OpenConnection())
            {
                InitUtente();

                //forzo aggiornamento dei parametri del DB che altrimenti sono sicronizzati con quelli del workbook
                DataBase.IdApplicazione = Workbook.IdApplicazione;
                DataBase.DataAttiva = Workbook.DataAttiva;
                DataBase.IdUtente = Workbook.IdUtente;
                
                //Se non ci sono tabelle, inizializzo il repository allo stato attuale
                if (Workbook.Repository.TablesCount == 0)
                {
                    SplashScreen.Show();
                    Workbook.Repository.Aggiorna();
                    DaAggiornare = true;
                }
                
                Workbook.AggiornaParametriApplicazione(false);

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
                
                DataBase.DataAttiva = Workbook.DataAttiva;
                Simboli.NomeApplicazione = Workbook.Repository.Applicazione["DesApplicazione"].ToString();
                Struct.intervalloGiorni = Workbook.Repository.Applicazione["IntervalloGiorniEntita"] is DBNull ? 0 : (int)Workbook.Repository.Applicazione["IntervalloGiorniEntita"];
                Struct.visualizzaRiepilogo = Workbook.Repository.Applicazione["VisRiepilogo"] is DBNull ? true : Workbook.Repository.Applicazione["VisRiepilogo"].Equals("1");

                return true;
            }
        }
        /// <summary>
        /// Procedura per l'update del foglio che apre la nuova copia e cancella la vecchia.
        /// </summary>
        /// <returns></returns>
        public static bool Update()
        {
            try
            {
                //UPDATE
                string updatePath = Path.Combine(_wb.Path, "UPDATE");
                if (Directory.Exists(updatePath) && Directory.GetFiles(updatePath, _wb.Name).Any())
                {
                    string name = _wb.Name;
                    string fullName = _wb.FullName;
                    Application.DisplayAlerts = false;
                    _wb.Base.SaveAs(Filename: Path.Combine(_wb.Path, "old_" + _wb.Name), ConflictResolution: Excel.XlSaveConflictResolution.xlLocalSessionChanges );
                    Application.DisplayAlerts = true;
                    File.Copy(Path.Combine(updatePath, name), fullName, true);
                    File.Delete(Path.Combine(updatePath, name));
                    Application.Workbooks.Open(fullName);
                    
                    
                    _wb.Base.Windows[1].Visible = false;
                    return true;
                }
                else
                {
                    try { if (File.Exists(Path.Combine(_wb.Path, "old_" + _wb.Name))) File.Delete(Path.Combine(_wb.Path, "old_" + _wb.Name)); }
                    catch { }
                }
            }
            catch
            {
                Application.DisplayAlerts = true;
            }

            return false;
        }
        /// <summary>
        /// Funzione per il controllo delle aree di rete utilizzate dall'applicativo: segnala all'utente l'impossibilità di raggiungerle.
        /// </summary>
        public static void ControlloAreeDiRete()
        {
            //controllo le aree di rete (se presenti)
            var usrConfig = GetUsrConfiguration();
            Dictionary<string, string> pathNonDisponibili = new Dictionary<string, string>();
            foreach (UserConfigElement ele in usrConfig.Items)
            {
                if (ele.Type == UserConfigElement.ElementType.path)
                {
                    string pathStr = Esporta.PreparePath(ele);

                    try { System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(pathStr); }
                    catch { pathNonDisponibili.Add(ele.Desc, pathStr); }
                }
            }

            if (Workbook.Repository.Applicazione != null)
            {
                //controllo il percorso di backup che ora è remoto
                try { System.Security.AccessControl.DirectorySecurity ds = Directory.GetAccessControl(Workbook.Repository.Applicazione["PathBackup"].ToString()); }
                catch { pathNonDisponibili.Add("Percorso di backup", Workbook.Repository.Applicazione["PathBackup"].ToString()); }
            }

            //segnalo all'utente l'impossibilità di accedere alle aree di rete
            if (pathNonDisponibili.Count > 0)
            {
                string paths = "\n";
                foreach (var kv in pathNonDisponibili)
                    paths += " - " + kv.Key + " : '" + kv.Value + "'\n";

                SplashScreen.Close();
                System.Windows.Forms.MessageBox.Show("I path seguenti non sono raggiungibili o non presentano privilegi di scrittura:" + paths, Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }
        /// <summary>
        /// Funzione che prepara l'ambiente in seguito all'avvio dell'applicativo.
        /// </summary>
        /// <param name="wb">Il workbook attivo.</param>
        public static void StartUp(IPSOThisWorkbook wb)
        {
            DaConsole = CheckAvvioAutomatico();

            _wb = wb;
            Application.DisplayAlerts = true;
            Application.CellDragAndDrop = false;
            Application.EnableAutoComplete = false;

            Repository = new Repository(wb);

            if(!DaConsole)
                DaAggiornare = false;

            //TODO ripristinare 
            DataBase.CreateNew(Ambiente);
            //DataBase.CreateNew(Simboli.DEV);

            DataBase.AddPropertyChanged(Workbook.StatoDBChanged);

            AbortedLoading = Update();

            if (!AbortedLoading)
            {
                Application.Iteration = true;
                Application.MaxIterations = 100;
                Application.EnableEvents = false;

                Style.StdStyles();

                ControlloAreeDiRete();

                foreach (Excel._Worksheet ws in CategorySheets)
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
        
        private static bool CheckAvvioAutomatico() 
        {
            string path = @"C:\Emergenza\AvvioAutomatico.xml";
            if (File.Exists(path))
            {
                XDocument doc = XDocument.Load(path);
                HaAzioni = false;
                HaEntita = false;

                foreach (XElement ele in doc.Element("AvvioAutomatico").Elements())
                {
                    switch (ele.Name.ToString())
                    {
                        case "AccettaCambioData":
                            AccettaCambioData = bool.Parse(ele.Value);
                            break;
                        case "RifiutaCambioData":
                            RifiutaCambioData = bool.Parse(ele.Value);
                            break;
                        case "AggiornaStruttura":
                            DaAggiornare = bool.Parse(ele.Value);
                            break;
                        case "AggiornaDati":
                            AggiornaDati = bool.Parse(ele.Value);
                            break;
                        case "ListaAzioni":
                            HaAzioni = true;
                            if (ele.Value.ToString().Length > 0)
                                ListaAzioni = ele.Value.ToString().Split(';');
                            else
                                ListaAzioni = null;
                            break;
                        case "ListaEntita":
                            HaEntita = true;
                            if (ele.Value.ToString().Length > 0)
                                ListaEntita = ele.Value.ToString().Split(';');
                            else
                                ListaEntita = null;
                            break;
                    }
                }
# if !DEBUG
                File.Delete(path);
# endif
                return true;
            }

            return false;

        }

        #endregion

        #region Log
        public static void InsertLog(PSO.Core.DataBase.TipologiaLOG logType, string message)
        {
            //si verifica quando aggiorno il workbook alla nuova versione. In ogni caso impedisco che il log vada in errore
            if (DataBase.IdUtente != -1)
            {
                Excel.Worksheet log = _wb.Sheets["Log"];
                bool prot = log.ProtectContents;
                if (prot) log.Unprotect(Password);
                DataBase db = new DataBase();
                db.InsertLog(logType, message);
                if (prot) log.Protect(Password);
            }
        }
        public static void RefreshLog()
        {
            Excel.Worksheet log = _wb.Sheets["Log"];
            bool prot = log.ProtectContents;
            bool events = Workbook.Application.EnableEvents;
            if (prot) log.Unprotect(Password);
            if (events) Workbook.Application.EnableEvents = false;
            DataBase db = new DataBase();
            db.RefreshLog();
            if (events) Workbook.Application.EnableEvents = true;
            if (prot) log.Protect(Password);
        }
        #endregion

        #region Close

        public static void Close() 
        {
            if (!AbortedLoading && Workbook.Repository != null)
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
                InsertLog(PSO.Core.DataBase.TipologiaLOG.LogAccesso, "Log off - " + Environment.UserName + " - " + Environment.MachineName);
                DataBase.Close();

                Application.ScreenUpdating = true;
                Application.CellDragAndDrop = true;
                Application.EnableAutoComplete = true;
                _wb.Base.Save();



                //try
                //{
                //    Application.Workbooks["old_" + _wb.Name].Close(SaveChanges: false);
                //    File.Delete(Path.Combine(_wb.Path, "old_" + _wb.Name));
                //}
                //catch { }
                    
                //var visibleConut = Application.Workbooks.OfType<Excel.Workbook>().Count(wb => wb.Windows[1].Visible == true);
                //if (visibleConut <= 1)
                //{
                //    Application.Quit();
                //    Marshal.FinalReleaseComObject(Application);
                //}

                //GC.Collect();
                //non metterlo mai, blocca l'applicazione ad un eventuale secondo avvio!!
                //Application.DisplayAlerts = false;
            }
        }

        #endregion

        #endregion
    }
}

