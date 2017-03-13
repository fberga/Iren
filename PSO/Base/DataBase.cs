using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Iren.PSO.Base
{
    public class DataBase
    {
        #region Costanti

        public const string NAME = "LocalDB";
        public struct SP 
        {
            public const string APPLICAZIONE = "spApplicazioneProprieta",
                APPLICAZIONE_INFORMAZIONE = "spApplicazioneInformazione",
                //APPLICAZIONE_INFORMAZIONE_H = "spApplicazioneInformazioneH",
                APPLICAZIONE_INFORMAZIONE_EXPORT = "spApplicazioneInformazione_Export",
                APPLICAZIONE_INFORMAZIONE_COMMENTO = "spApplicazioneInformazioneCommento",
                APPLICAZIONE_INIT = "spApplicazioneInit",
                APPLICAZIONE_LOG = "spApplicazioneLog",
                APPLICAZIONE_NOTE = "spApplicazioneNote",
                APPLICAZIONE_RIEPILOGO = "spApplicazioneRiepilogo",
                AZIONE = "spAzione",
                AZIONE_CATEGORIA = "spAzioneCategoria",
                CALCOLO = "spCalcolo",
                CALCOLO_INFORMAZIONE = "spCalcoloInformazione",
                CARICA_AZIONE_INFORMAZIONE = "spCaricaAzioneInformazione",
                CATEGORIA = "spCategoria",
                CATEGORIA_ENTITA = "spCategoriaEntita",
                CHECK_FONTE_METEO = "spCheckFonteMeteo",
                CHECKMODIFICASTRUTTURA = "spCheckModificaStruttura",
                ENTITA_ASSETTO = "spEntitaAssetto",
                ENTITA_AZIONE = "spEntitaAzione",
                ENTITA_AZIONE_CALCOLO = "spEntitaAzioneCalcolo",
                ENTITA_AZIONE_INFORMAZIONE = "spEntitaAzioneInformazione",
                ENTITA_CALCOLO = "spEntitaCalcolo",
                ENTITA_COMMITMENT = "spEntitaCommitment",
                ENTITA_GRAFICO = "spEntitaGrafico",
                ENTITA_GRAFICO_INFORMAZIONE = "spEntitaGraficoInformazione",
                ENTITA_INFORMAZIONE = "spEntitaInformazione",
                ENTITA_INFORMAZIONE_FORMATTAZIONE = "spEntitaInformazioneFormattazione",
                ENTITA_PARAMETRO_D = "spEntitaParametroD",
                ENTITA_PARAMETRO_H = "spEntitaParametroH",
                ENTITA_PARAMETRO = "spEntitaParametro",
                ENTITA_PROPRIETA = "spEntitaProprieta",
                ENTITA_RAMPA = "spEntitaRampa",
                GET_LAST_DATA_VALIDATA_GAS = "spGetLastDataValidataGas",
                GET_ORE_FERMATA = "spGetOreFermata",
                GET_VERSIONE = "spGetVersione",
                INSERT_APPLICAZIONE_INFORMAZIONE_XML = "spInsertApplicazioneInformazioneXML",
                INSERT_APPLICAZIONE_RIEPILOGO = "spInsertApplicazioneRiepilogo",
                INSERT_LOG = "spInsertLog",
                INSERT_PROGRAMMAZIONE_PARAMETRO = "spInsertProgrammazione_Parametro",
                STAGIONE = "spTipologiaStagione",
                TIPOLOGIA_CHECK = "spTipologiaCheck",
                UPDATE_METEO = "spUpdateFonteMeteo",
                UTENTE = "spUtente",
                UTENTE_GRUPPO = "spUtenteGruppo",
                DEFINIZIONE_OFFERTA = "spFormDefinizioneOfferta";
                

            public struct PAR
            {
                public const string DELETE_PARAMETRO = "PAR.spDeleteParametro",
                ELENCO_PARAMETRI = "PAR.spElencoParametri",
                INSERT_PARAMETRO = "PAR.spInsertParametro",
                VALORI_PARAMETRI = "PAR.spValoriParametri",
                UPDATE_PARAMETRO = "PAR.spUpdateParametro";
            }

            public struct  RIBBON
            {
                public const string GRUPPO_CONTROLLO = "RIBBON.spGruppoControllo",
                    CONTROLLO_APPLICAZIONE = "RIBBON.spControlloApplicazione",
                    CONTROLLO_FUNZIONE = "RIBBON.spControlloFunzione",

                    CONTROLLO = "RIBBON.spControllo",
                    INSERT_CONTROLLO = "RIBBON.spInsertControllo",
                    INSERT_GRUPPO = "RIBBON.spInsertGruppo",
                    INSERT_GRUPPO_CONTROLLO = "RIBBON.spInsertGruppoControllo",
                    DELETE_GRUPPO_CONTROLLO = "RIBBON.spDeleteGruppoControllo",
                    DELETE_FUNZIONI_CONTROLLO = "RIBBON.spDeleteControlloFunzione",
                    INSERT_CONTROLLO_FUNZIONE = "RIBBON.spInsertControlloFunzione",
                    FUNZIONE = "RIBBON.spFunzione",
                    COPIA_CONFIGURAZIONE = "RIBBON.spCopiaConfigurazione";
            }

        }
        public struct TAB 
        {
            public const string ADDRESS_FROM = "AddressFrom",
                ADDRESS_TO = "AddressTo",
                ANNOTA = "AnnotaModifica",
                //APPLICAZIONE_RIBBON = "ApplicazioneRibbon",
                AZIONE = "Azione",
                AZIONE_CATEGORIA = "AzioneCategoria",
                CALCOLO = "Calcolo",
                CALCOLO_INFORMAZIONE = "CalcoloInformazione",
                CATEGORIA = "Categoria",
                CATEGORIA_ENTITA = "CategoriaEntita",
                CHECK = "Check",
                DATE_DEFINITE = "DefinedDates",
                DATI_APPLICAZIONE = "DatiApplicazione",
                //DATI_APPLICAZIONE_D = "DatiApplicazioneD",
                DATI_APPLICAZIONE_COMMENTO = "DatiApplicazioneCommento",
                DEFINIZIONE_OFFERTA = "DefinizioneOfferta",
                EDITABILI = "Editabili",
                ENTITA_ASSETTO = "EntitaAssetto",
                ENTITA_AZIONE = "EntitaAzione",
                ENTITA_AZIONE_CALCOLO = "EntitaAzioneCalcolo",
                ENTITA_AZIONE_INFORMAZIONE = "EntitaAzioneInformazione",
                ENTITA_CALCOLO = "EntitaCalcolo",
                ENTITA_COMMITMENT = "EntitaCommitment",
                ENTITA_GRAFICO = "EntitaGrafico",
                ENTITA_GRAFICO_INFORMAZIONE = "EntitaGraficoInformazione",
                ENTITA_INFORMAZIONE = "EntitaInformazione",
                ENTITA_INFORMAZIONE_FORMATTAZIONE = "EntitaInformazioneFormattazione",
                //ENTITA_PARAMETRO_D = "EntitaParametroD",
                //ENTITA_PARAMETRO_H = "EntitaParametroH",
                ENTITA_PARAMETRO = "EntitaParametro",
                ENTITA_PROPRIETA = "EntitaProprieta",
                ENTITA_RAMPA = "EntitaRampa",
                EXPORT_XML = "ExportXML",
                LISTA_APPLICAZIONI = "ListaApplicazioni",
                LOG = "Log",
                MERCATI = "Mercati",
                MODIFICA = "Modifica",
                MODIFICA_CANCEL = "ModificaCancel",
                NOMI_DEFINITI = "DefinedNames",
                SALVADB = "SaveDB",
                SELECTION = "Selection",
                STAGIONE = "Stagione",
                TIPOLOGIA_CHECK = "TipologiaCheck";

            public struct RIBBON
            {
                public const string GRUPPO_CONTROLLO = "GruppoControllo",
                    CONTROLLO_APPLICAZIONE = "ControlloApplicazione",
                    CONTROLLO_FUNZIONE = "ControlloFunzione";
            }
        }

        #endregion

        #region Variabili
        
        protected static Core.DataBase _db = null;

        #endregion

        #region Proprietà statiche
        /// <summary>
        /// Restituisce lo stato dei database utilizzati dall'applicazione.
        /// </summary>
        public static Dictionary<Core.DataBase.NomiDB, ConnectionState> StatoDB 
        {
            get
            {
                if (Simboli.EmergenzaForzata)
                {
                    return new Dictionary<Core.DataBase.NomiDB, ConnectionState>() 
                    {
                        {Core.DataBase.NomiDB.SQLSERVER, ConnectionState.Closed},
                        {Core.DataBase.NomiDB.IMP, ConnectionState.Closed},
                        {Core.DataBase.NomiDB.ELSAG, ConnectionState.Closed}
                    };
                }

                return _db.StatoDB;
            }
        }
        /// <summary>
        /// Restituisce la versione della classe Iren.PSO.Core.DataBase.
        /// </summary>
        public static System.Version Versione 
        { get { return _db.GetCurrentV(); } }
        /// <summary>
        /// True se il database è stato inizializzato.
        /// </summary>
        public static bool IsInitialized 
        { get; private set; }
        /// <summary>
        /// Id dell'utente configurato.
        /// </summary>
        public static int IdUtente 
        { get { return _db.IdUtente; } set { _db.IdUtente = value; } }
        /// <summary>
        /// Id dell'applicazione utilizzata.
        /// </summary>
        public static int IdApplicazione 
        { get { return _db.IdApplicazione; } set { _db.IdApplicazione = value; } }
        /// <summary>
        /// Data attiva.
        /// </summary>
        public static DateTime DataAttiva 
        { get { return _db.DataAttiva; } set { _db.DataAttiva = value; } }

        #endregion

        #region Metodi Statici
        
        /// <summary>
        /// Metodo che aggiunge un handler all'evento PropertyChanged del DataBase.
        /// </summary>
        /// <param name="handler">Handler per l'evento PropertyChanged.</param>
        public static void AddPropertyChanged(System.ComponentModel.PropertyChangedEventHandler handler)
        {
            _db.PropertyChanged += handler;
        }
        /// <summary>
        /// Inizializza il nuovo Core.DataBase collegato al dbName che rappresenta l'ambiente Prod|Test|Dev.
        /// </summary>
        /// <param name="dbName">Nome (corrisponde all'ambiente) del Database.</param>
        public static void CreateNew(string ambiente, bool checkDB = true)
        {
            if (_db == null || _db.Ambiente != ambiente)
                _db = new Core.DataBase(ambiente, checkDB);

            IsInitialized = true;
        }
        /// <summary>
        /// Disattiva l'oggetto DataBase.
        /// </summary>
        public static void Close()
        {
            IsInitialized = false;
            _db.Dispose();
        }
        /// <summary>
        /// Cambio ambiente tra Prod|Test|Prod.
        /// </summary>
        /// <param name="ambiente">Nuovo ambiente.</param>
        public static bool SwitchEnvironment(string ambiente)
        {
            if (_db.Ambiente != ambiente)
            {
                Workbook.Ambiente = ambiente;

                Delegate[] list = _db.GetPropertyChangedInvocationList();

                _db = new Core.DataBase(ambiente);
                _db.SetParameters(Workbook.DataAttiva, Workbook.IdUtente, Workbook.IdApplicazione);
                foreach (Delegate d in list)
                {
                    EventInfo ei = _db.GetType().GetEvent("PropertyChanged");
                    ei.AddEventHandler(_db, d);
                }
                return true;
            }
            return false;
        }
        /// <summary>
        /// Salva le modifiche effettuate ai fogli sul DataBase. Il processo consiste nella creazione di un file XML contenente tutte le righe della tabella di Modifica e successivo svuotamento della tabella stessa. Il processo richiede una connessione aperta. Diversamente le modifiche vengono salvate nella cartella di Emergenza dove, al momento della successiva chiamata al metodo, vengono reinviati al server in ordine cronologico.
        /// </summary>
        public static void SalvaModificheDB(string nomeTabella = TAB.MODIFICA)
        {
            if (Workbook.Repository != null)
            {
                //prendo la tabella di modifica e controllo se è nulla
                DataTable modifiche = Workbook.Repository[nomeTabella];
                if (modifiche != null && Workbook.IdUtente != 0)   //non invia se l'utente non è configurato... in ogni caso la tabella è vuota!!
                {
                    //tolgo il namespace che altrimenti aggiunge informazioni inutili al file da mandare al server
                    DataTable dt = modifiche.Copy();
                    dt.TableName = TAB.MODIFICA;
                    dt.Namespace = "";

                    //path del caricatore sul server
                    string cartellaRemota = Workbook.Repository.Applicazione["PathExportModifiche"].ToString();
                    //path della cartella di emergenza
                    string cartellaEmergenza = Workbook.Repository.Applicazione["PathExportModificheEmergenza"].ToString();
                    //path della cartella di archivio in cui copiare i file in caso di esito positivo nel saltavaggio
                    string cartellaArchivio = Workbook.Repository.Applicazione["PathExportModificheArchivio"].ToString();

                    string fileName = "";
                    //se la connessione è aperta (in emergenza forzata sarà sempre false) ed esiste la cartella del caricatore
                    if (OpenConnection() && Directory.Exists(cartellaRemota))
                    {
                        //metto in lavorazione i file nella cartella di emergenza
                        string[] fileEmergenza = Directory.GetFiles(cartellaEmergenza);
                        bool executed = false;
                        if (fileEmergenza.Length > 0)
                        {
                            if (System.Windows.Forms.MessageBox.Show("Sono presenti delle modifiche non ancora salvate sul DB. Procedere con il salvataggio? \n\nPremere Sì per inviare i dati al server, No per cancellare definitivamente le modifiche.", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.Yes)
                            {
                                //il nome file contiene la data, quindi li metto in ordine cronologico
                                Array.Sort<string>(fileEmergenza);
                                foreach (string file in fileEmergenza)
                                {
                                    File.Copy(file, Path.Combine(cartellaRemota, file.Split('\\').Last()), true);

                                    executed = DataBase.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new Core.QryParams() { { "@NomeFile", file.Split('\\').Last() } });
                                    if (executed)
                                    {
                                        if (!Directory.Exists(cartellaArchivio))
                                            Directory.CreateDirectory(cartellaArchivio);

                                        File.Move(Path.Combine(cartellaRemota, file.Split('\\').Last()), Path.Combine(cartellaArchivio, file.Split('\\').Last()));
                                        File.Delete(file);
                                    }
                                    else
                                    {
                                        System.Windows.Forms.MessageBox.Show("Il server ha restituito un errore nel salvataggio. Le modifiche rimarranno comunque salvate in locale.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    }
                                }
                            }
                            else
                            {
                                foreach (string file in fileEmergenza)
                                    File.Delete(file);
                            }
                        }

                        //controllo se la tabella è vuota
                        if (dt.Rows.Count == 0)
                            return;

                        //salvo le modifiche appena effettuate
                        fileName = Path.Combine(cartellaRemota, Simboli.NomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + (Workbook.Ambiente != Simboli.PROD ? "_" + Workbook.Ambiente : "") + ".xml");
                        dt.WriteXml(fileName);//, XmlWriteMode.WriteSchema);

                        //se la query indica che il processo è andato a buon fine, sposto in archivio
                        executed = DataBase.Insert(SP.INSERT_APPLICAZIONE_INFORMAZIONE_XML, new Core.QryParams() { { "@NomeFile", fileName.Split('\\').Last() } });
                        if (executed)
                        {
                            if (!Directory.Exists(cartellaArchivio))
                                Directory.CreateDirectory(cartellaArchivio);

                            File.Move(fileName, Path.Combine(cartellaArchivio, fileName.Split('\\').Last()));
                        }
                        else
                        {
                            if (!Directory.Exists(cartellaEmergenza))
                                Directory.CreateDirectory(cartellaEmergenza);

                            fileName = Path.Combine(cartellaEmergenza, Simboli.NomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + (Workbook.Ambiente != Simboli.PROD ? "_" + Workbook.Ambiente : "") + ".xml");
                            dt.WriteXml(fileName);//, XmlWriteMode.WriteSchema);

                            Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "Errore nel salvataggio delle modifiche. '" + fileName + "' si trova in " + Environment.MachineName);

                            System.Windows.Forms.MessageBox.Show("Il server ha restituito un errore nel salvataggio. Le modifiche rimarranno comunque salvate in locale.", Simboli.NomeApplicazione + " - ERRORE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        }
                    }
                    else
                    {
                        if (dt.Rows.Count == 0)
                            return;

                        if (!Directory.Exists(cartellaEmergenza))
                            Directory.CreateDirectory(cartellaEmergenza);

                        //metto le modifiche nella cartella emergenza
                        fileName = Path.Combine(cartellaEmergenza, Simboli.NomeApplicazione.Replace(" ", "").ToUpperInvariant() + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + ".xml");
                        
                        dt.WriteXml(fileName, XmlWriteMode.WriteSchema);

                        System.Windows.Forms.MessageBox.Show("A causa di problemi di rete le modifiche sono state salvate in locale", Simboli.NomeApplicazione + " - ATTENZIONE!!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
                    }

                    //svuoto la tabella di modifiche
                    modifiche.Clear();
                }
            }
        }
        /// <summary>
        /// Aggiunge la riga di riepilogo al DB in modo da far evidenziare la casella nel foglio Main del Workbook.
        /// </summary>
        /// <param name="siglaEntita">La sigla dell'entità di cui aggiungere il riepilogo.</param>
        /// <param name="siglaAzione">La sigla dell'azione di cui aggiungere il riepilogo.</param>
        /// <param name="giorno">Il giorno in cui aggiungere il riepilogo.</param>
        /// <param name="presente">Se il dato collegato alla coppia Entità - Azione è presente o no nel DB.</param>
        public static void InsertApplicazioneRiepilogo(object siglaEntita, object siglaAzione, DateTime giorno, bool presente = true, String parametro = null)
        {
            bool visible = Workbook.Repository[DataBase.TAB.AZIONE]
                .AsEnumerable()
                .Where(r => r["SiglaAzione"].Equals(siglaAzione))
                .Select(r => r["Visibile"].Equals("1"))
                .FirstOrDefault();

            if (visible)
            {
                try
                {
                    if (OpenConnection())
                    {
                        Core.QryParams parameters = new Core.QryParams() {
                            {"@SiglaEntita", siglaEntita},
                            {"@SiglaAzione", siglaAzione},
                            {"@Data", giorno.ToString("yyyyMMdd")},
                            {"@Presente", presente ? "1" : "0"}
                        };

                        /* 13/3/2017 Così riesco per MI a differenziare il riepilogo per mercato valorizzando 'parametro'  */
                        if (parametro != null)
                        {
                            parametro = parametro.Substring(parametro.Length - 1, 1);
                            parameters.Add("@Parametro", parametro);
                        }

                        _db.Insert(DataBase.SP.INSERT_APPLICAZIONE_RIEPILOGO, parameters);
                    }
                }
                catch (Exception e)
                {
                    Workbook.InsertLog(Core.DataBase.TipologiaLOG.LogErrore, "InsertApplicazioneRiepilogo [" + giorno + ", " + siglaEntita + ", " + siglaAzione + "]: " + e.Message);

                    System.Windows.Forms.MessageBox.Show(e.Message, Simboli.NomeApplicazione + " - ERRORE!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                }
            }
        }
        /// <summary>
        /// Inizializza i valori di default e imposta tutte le informazioni che devono essere "trascinate" dai giorni precedenti
        /// </summary>
        public static void ExecuteSPApplicazioneInit(DateTime giorno)
        {
            SplashScreen.UpdateStatus("Inizializzazione valori di default");
            Select(SP.APPLICAZIONE_INIT, "@Data=" + giorno.ToString("yyyyMMdd"));
        }
        /// <summary>
        /// Funzione per l'apertura della connessione che considera anche la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <returns>True se la connessione viene aperta, false altrimenti.</returns>
        public static bool OpenConnection()
        {
            if (!Simboli.EmergenzaForzata)
                return _db.OpenConnection();

            return false;
        }
        /// <summary>
        /// Funzione per chiudere la connessione che considera anche la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <returns>True se la connessione viene chiusa, false altrimenti.</returns>
        public static bool CloseConnection()
        {
            if (!Simboli.EmergenzaForzata)
                return _db.CloseConnection();

            return false;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, Core.QryParams parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();

                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, String parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, parameters, timeout);
                CloseConnection();

                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="timeout">Timeout del comando.</param>
        /// <returns>DataTable contenente il risultato della storedProcedure.</returns>
        public static DataTable Select(string storedProcedure, int timeout = 300)
        {
            if (OpenConnection())
            {
                DataTable dt = _db.Select(storedProcedure, timeout);
                CloseConnection();

                return dt;
            }

            return null;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout per la query.</param>
        /// <returns>True se il comando è andato a buon fine, false altrimenti.</returns>
        public static bool Insert(string storedProcedure, Core.QryParams parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                bool o = _db.Insert(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }
        /// <summary>
        /// Funzione per eseguire una stored procedure. Fa un "override" della funzione fornita da Core.DataBase che considera la presenza del flag di Emergenza Forzata. Restituisce i parametri di output che vengono valorizzati dalla stored procedure.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="outParams">Lista dei parametri in uscita indicizzati per nome.</param>
        /// <param name="timeout">Timeout per la query.</param>
        /// <returns>True se il comando è andato a buon fine, false altrimenti.</returns>
        public static bool Insert(string storedProcedure, Core.QryParams parameters, out Dictionary<string, object> outParams, int timeout = 300)
        {
            if (OpenConnection())
            {
                bool o = _db.Insert(storedProcedure, parameters, out outParams, timeout);
                CloseConnection();
                return o;
            }
            outParams = null;
            return false;
        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di cancellazione (principalmente una questione mnemonica). Restituisce true se il comando è andato a buon fine, false altrimenti.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout per la query.</param>
        /// <returns>Restituisce true se il comando è andato a buon fine, false altrimenti.</returns>
        public static bool Delete(string storedProcedure, Core.QryParams parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                bool o = _db.Delete(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di cancellazione (principalmente una questione mnemonica). Restituisce true se il comando è andato a buon fine, false altrimenti.
        /// </summary>
        /// <param name="storedProcedure">Stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout per la query.</param>
        /// <returns>Restituisce true se il comando è andato a buon fine, false altrimenti.</returns>
        public static bool Delete(string storedProcedure, string parameters, int timeout = 300)
        {
            if (OpenConnection())
            {
                bool o = _db.Delete(storedProcedure, parameters, timeout);
                CloseConnection();
                return o;
            }
            return false;
        }

        #endregion

        #region Metodi

        /// <summary>
        /// Inserisce una riga di Log.
        /// </summary>
        /// <param name="logType">Tipologia di Log.</param>
        /// <param name="message">Messaggio del Log.</param>
        public void InsertLog(Core.DataBase.TipologiaLOG logType, string message)
        {
#if (!DEBUG || DEBUG)
            if (OpenConnection())
            {
                Insert(SP.INSERT_LOG, new Core.QryParams() { { "@IdTipologia", logType }, { "@Messaggio", message } });
            }
#endif
            RefreshLog();
        }
        /// <summary>
        /// Aggiorna il foglio di Log.
        /// </summary>
        public void RefreshLog()
        {
            if (OpenConnection())
            {
                DataTable dt = Select(SP.APPLICAZIONE_LOG);
                if (dt != null)
                {
                    //Workbook.RemoveStdStoreEdit();
                    Workbook.LogDataTable.Clear();
                    Workbook.LogDataTable.Merge(dt);

                    if (Workbook.Log.ListObjects.Count > 0)
                        Workbook.Log.ListObjects[1].Range.EntireColumn.AutoFit();

                    //Workbook.AddStdStoreEdit();
                }
            }
        }

        #endregion
    }
}
