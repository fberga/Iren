using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Iren.PSO.Core
{
    public class DataBase : INotifyPropertyChanged, IDisposable
    {
        #region Nomi di Sistema

        public enum TipologiaLOG
        {
            LogErrore = 1,
            LogCarica = 2,
            LogGenera = 3,
            LogEsporta = 4,
            LogModifica = 5,
            LogAccesso = 6
        }

        public enum NomiDB
        {
            SQLSERVER = 1,
            IMP = 2,
            ELSAG = 3
        }

        public const string ALL = "ALL";

        #endregion

        #region Variabili

        private Command _cmd;
        private Command _internalCmd;

        private System.Threading.Timer _checkDBTrhead;

        private SqlConnection _sqlConn;
        private SqlConnection _internalsqlConn;
        private string _connStr = "";

        //private string _dataAttiva = DateTime.Now.ToString("yyyyMMdd");
        //private int _idUtenteAttivo = -1;
        //private int _idApplicazione = -1;
        private Dictionary<NomiDB, ConnectionState> _statoDB = new Dictionary<NomiDB, ConnectionState>() { 
            {NomiDB.SQLSERVER, ConnectionState.Closed},
            {NomiDB.IMP, ConnectionState.Closed},
            {NomiDB.ELSAG, ConnectionState.Closed}
        };

        private string _ambiente;
        public event PropertyChangedEventHandler PropertyChanged;

        private int _chechDBCount = 4;

        private bool disposed = false;

        #endregion

        #region Proprietà

        public DateTime DataAttiva { get; set; }
        public int IdUtente { get; set; }
        public int IdApplicazione { get; set; }
        public string Ambiente { get { return _ambiente; } }

        public Dictionary<NomiDB, ConnectionState> StatoDB { get { return _statoDB; } }

        #endregion

        #region Costruttori

        public DataBase(string dbName, bool checkDB = true) 
        {
            _ambiente = dbName;
            try
            {
                _connStr = ConfigurationManager.ConnectionStrings[dbName].ConnectionString;
                _sqlConn = new SqlConnection(_connStr);
                _internalsqlConn = new SqlConnection(_connStr);

                _cmd = new Command(_sqlConn);
//#if !DEBUG
                if (checkDB)
                {
                    _internalCmd = new Command(_internalsqlConn);
                    _checkDBTrhead = new System.Threading.Timer(CheckDB, null, 0, 250 * 60);
                }
//#endif
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "Core.DataBase - ERROR!!", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }

            IdApplicazione = -1;
            IdUtente = -1;
            DataAttiva = DateTime.MinValue;

        }

        #endregion

        #region Metodi Pubblici

        /// <summary>
        /// Apre la connessione al DB.
        /// </summary>
        /// <returns>True se la connessione è aperta, false altrimenti.</returns>
        public bool OpenConnection()
        {
            return OpenConnection(_sqlConn);
        }
        /// <summary>
        /// Chiude la connessione al DB.
        /// </summary>
        /// <returns>True se la connessione viene chiusa o era già chiusa, false altrimenti.</returns>
        public bool CloseConnection()
        {
            return CloseConnection(_sqlConn);
        }
        /// <summary>
        /// Imposta i parametri principali da utilizzare in quasi tutti i comandi.
        /// </summary>
        /// <param name="dataAttiva">La data di riferimento.</param>
        /// <param name="idUtente">Id dell'utente che ha eseguito il login.</param>
        /// <param name="idApplicazione">Id dell'applicazione aperta.</param>
        public void SetParameters(DateTime dataAttiva, int idUtente, int idApplicazione)
        {
            DataAttiva = dataAttiva;
            IdUtente = idUtente;
            IdApplicazione = idApplicazione;
        }
        /// <summary>
        /// Cambia la data di riferimento.
        /// </summary>
        /// <param name="dataAttiva">Nuova data di riferimento.</param>
        //public void ChangeDate(string dataAttiva)
        //{
        //    _dataAttiva = dataAttiva;
        //}
        /// <summary>
        /// Cambia l'id applicazione (utilizzato solo in invio programmi quando viene cambiato il programma).
        /// </summary>
        /// <param name="appID">Nuovo id applicazione.</param>
        //public void ChangeAppID(int appID)
        //{
        //    _idApplicazione = appID;
        //}

        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure che inserisce o modifica valori sul db senza restituire nessun risultato.
        /// </summary>
        /// <param name="storedProcedure">Nome della stored procedure.</param>
        /// <param name="parameters">Parametri richiesti dalla stored procedure.</param>
        /// <returns>True se il comando è andato a buon fine, false altrimenti.</returns>
        public bool Insert(string storedProcedure, QryParams parameters, int timeout = 300)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && IdApplicazione != -1)
                parameters.Add("@IdApplicazione", IdApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && IdUtente != -1)
                parameters.Add("@IdUtente", IdUtente);
            if (!parameters.ContainsKey("@Data") && DataAttiva != DateTime.MinValue)
                parameters.Add("@Data", DataAttiva.ToString("yyyyMMdd"));

            try
            {
                SqlCommand cmd = _cmd.SqlCmd(storedProcedure, parameters, timeout);
                cmd.ExecuteNonQuery();
                return cmd.Parameters[0].Value.Equals(0);
            }
            catch (TimeoutException) 
            {
                return false;
            }
            catch (SqlException e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message);
                return false;
            }
        }
        public bool Insert(string storedProcedure, string parameters, int timeout = 300)
        {
            return Insert(storedProcedure, getParamsFromString(parameters), timeout);
        }
        public bool Insert(string storedProcedure, QryParams parameters, out Dictionary<string, object> outParams, int timeout = 300)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && IdApplicazione != -1)
                parameters.Add("@IdApplicazione", IdApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && IdUtente != -1)
                parameters.Add("@IdUtente", IdUtente);
            if (!parameters.ContainsKey("@Data") && DataAttiva != DateTime.MinValue)
                parameters.Add("@Data", DataAttiva.ToString("yyyyMMdd"));

            try
            {
                SqlCommand cmd = _cmd.SqlCmd(storedProcedure, parameters, timeout);
                cmd.ExecuteNonQuery();
                outParams = new Dictionary<string, object>();
                
                foreach (SqlParameter par in cmd.Parameters)
                {
                    if (par.Direction == ParameterDirection.InputOutput || par.Direction == ParameterDirection.Output || par.Direction == ParameterDirection.ReturnValue)
                        outParams.Add(par.ParameterName, par.Value);
                }
                return true;
            }
            catch (TimeoutException)
            {
                outParams = null;
                return false;
            }
        }
        public bool Insert(string storedProcedure, string parameters, out Dictionary<string, object> outParams, int timeout = 300)
        {
            return Insert(storedProcedure, getParamsFromString(parameters), out outParams, timeout);
        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di selezone di valori. Restituisce una tabella contenente i record restituiti dal comando.
        /// </summary>
        /// <param name="storedProcedure">Nome della stored procedure.</param>
        /// <param name="parameters">Parametri richiesti dalla stored procedure.</param>
        /// <param name="timeout">Time out di esecuzione della stored procedure.</param>
        /// <returns>Tabella contenente i record restituiti dalla stored procedure.</returns>
        public DataTable Select(string storedProcedure, QryParams parameters, int timeout = 300)
        {
            return Select(_cmd, storedProcedure, parameters, timeout);
        }
        public DataTable Select(string storedProcedure, String parameters, int timeout = 300)
        {
            return Select(storedProcedure, getParamsFromString(parameters), timeout);
        }
        public DataTable Select(string storedProcedure, int timeout = 300)
        {
            QryParams parameters = new QryParams();
            return Select(storedProcedure, parameters, timeout);
        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di cancellazione (solo mnemonico, esegue funzione insert).
        /// </summary>
        /// <param name="storedProcedure">Nome della stored procedure.</param>
        /// <param name="parameters">Parametri richiesti dalla stored procedure.</param>
        /// <param name="timeout">Time out di esecuzione della stored procedure.</param>
        /// <returns>Restituisce true se la query è andata a buon fine, false se è andata in timeout.</returns>
        public bool Delete(string storedProcedure, QryParams parameters, int timeout = 300)
        {
            return Insert(storedProcedure, parameters, timeout);
        }
        public bool Delete(string storedProcedure, string parameters, int timeout = 300)
        {
            return Delete(storedProcedure, getParamsFromString(parameters), timeout);
        }

        /// <summary>
        /// Restituisce la versione attuale della libreria core.
        /// </summary>
        /// <returns>Versione.</returns>
        public System.Version GetCurrentV()
        {
            return Assembly.GetExecutingAssembly().GetName().Version;
        }

        public Delegate[] GetPropertyChangedInvocationList()
        {
            return PropertyChanged.GetInvocationList();
        }

        public void Dispose() 
        {
            Dispose(true);
            GC.SuppressFinalize(this); 
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                if (_sqlConn.State == ConnectionState.Open)
                    _sqlConn.Close();
                if (_internalsqlConn != null && _internalsqlConn.State == ConnectionState.Open)
                    _internalsqlConn.Close();

                PropertyChanged = null;

                _cmd.Dispose();
                _sqlConn.Dispose();

                if (_internalCmd != null)
                    _internalCmd.Dispose();
                if (_internalsqlConn != null)
                    _internalsqlConn.Dispose();
                if (_checkDBTrhead != null)
                    _checkDBTrhead.Dispose();
            }

            // Free any unmanaged objects here.
            //
            disposed = true;
        }

        #endregion

        #region Metodi Privati

        /// <summary>
        /// Dalla stringa di parametri del tipo "@param1=val1;@param2=val2;..." isola i singoli parametri e restituisce la lista QryParams.
        /// </summary>
        /// <param name="parameters">Stringa di parametri.</param>
        /// <returns>Lista di parametri.</returns>
        private QryParams getParamsFromString(string parameters)
        {
            Regex regex = new Regex(@"@\w+[=:][^;:=]+");
            MatchCollection paramsList = regex.Matches(parameters);
            Regex split = new Regex("[=:]");
            QryParams o = new QryParams();
            foreach (Match par in paramsList)
            {
                string[] keyVal = split.Split(par.Value);
                if (keyVal.Length != 2)
                    throw new InvalidExpressionException("The provided list of parameters is invalid.");
                o.Add(keyVal[0], keyVal[1]);
            }
            return o;
        }

        /// <summary>
        /// Apre la connessione conn.
        /// </summary>
        /// <param name="conn">Connessione da aprire.</param>
        /// <returns>True se la connessione viene aperta o è già aperta, false altrimenti.</returns>
        private bool OpenConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Closed)
                    conn.Open();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
        /// <summary>
        /// Chiude la connessione conn.
        /// </summary>
        /// <param name="conn">Connessione da chiudere.</param>
        /// <returns>True se la connessione viene chiusa o è chiusa, false altrimenti.</returns>
        private bool CloseConnection(SqlConnection conn)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
            catch (Exception)
            {
                return false;
            }

            return true;
        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di selezone di valori. Restituisce una tabella contenente i record restituiti dal comando.
        /// </summary>
        /// <param name="cmd">Comando con cui eseguire la stored procedure.</param>
        /// <param name="storedProcedure">Nome della stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <param name="logEnabled">Flag per attivare disattivare il log (usa spInsertLog che deve essere definita nello schema in uso!!).</param>
        /// <returns>Tabella contenente i valori restituiti dalla stored procedure.</returns>
        private DataTable Select(Command cmd, string storedProcedure, QryParams parameters, int timeout = 300, bool logEnabled = true)
        {
            if (!parameters.ContainsKey("@IdApplicazione") && IdApplicazione != -1)
                parameters.Add("@IdApplicazione", IdApplicazione);
            if (!parameters.ContainsKey("@IdUtente") && IdUtente != -1)
                parameters.Add("@IdUtente", IdUtente);
            if (!parameters.ContainsKey("@Data") && DataAttiva != DateTime.MinValue)
                parameters.Add("@Data", DataAttiva.ToString("yyyyMMdd"));

            try
            {
                DataTable dt = new DataTable();

                using (SqlDataReader dr = cmd.SqlCmd(storedProcedure, parameters, timeout).ExecuteReader())
                {
                    dt.Load(dr);
                }
                return dt;
            }
            catch (SqlException)
            {
                if (logEnabled && OpenConnection())
                {
                    //nel caso spInsertLog (definita solo per PSO ma non per RiMoST ad esempio) non sia definita, va in errore
                    try { Insert("spInsertLog", new QryParams() { { "@IdTipologia", TipologiaLOG.LogErrore }, { "@Messaggio", "Core.DataBase.Select[" + storedProcedure + "," + parameters + "]" } }); } 
                    catch {}
                    CloseConnection();
                }
                return null;
            }
            catch (InvalidOperationException) { return null; }

        }
        /// <summary>
        /// Funzione per l'esecuzione di una stored procedure di selezone di valori. Restituisce una tabella contenente i record restituiti dal comando.
        /// </summary>
        /// <param name="cmd">Comando con cui eseguire la stored procedure.</param>
        /// <param name="storedProcedure">Nome della stored procedure.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <param name="timeout">Timeout di esecuzione.</param>
        /// <param name="logEnabled">Flag per attivare disattivare il log (usa spInsertLog che deve essere definita nello schema in uso!!).</param>
        /// <returns>Tabella contenente i valori restituiti dalla stored procedure.</returns>
        private DataTable Select(Command cmd, string storedProcedure, String parameters, int timeout = 300, bool logEnabled = true)
        {
            return Select(cmd, storedProcedure, getParamsFromString(parameters), timeout);
        }
        /// <summary>
        /// Controlla lo stato del DB e notifica se ci sono stati dei cambi.
        /// </summary>
        /// <param name="state">Stato.</param>
        private void CheckDB(object state)
        {
            Dictionary<NomiDB, ConnectionState> oldStatoDB = new Dictionary<NomiDB, ConnectionState>(_statoDB);
            
            OpenConnection(_internalsqlConn);
            
            _statoDB[NomiDB.SQLSERVER] = _internalsqlConn.State;

            if (_statoDB[NomiDB.SQLSERVER] == ConnectionState.Open)
            {
                if (_chechDBCount == 4)
                {
                    _chechDBCount = 0;
                    if (OpenConnection(_internalsqlConn))
                    {
                        DataTable imp = Select(_internalCmd, "spCheckDB", "@Nome=IMP", 5, false);

                        if (imp != null && imp.Rows.Count > 0 && imp.Rows[0]["Stato"].Equals(0))
                            _statoDB[NomiDB.IMP] = ConnectionState.Open;
                        else
                            _statoDB[NomiDB.IMP] = ConnectionState.Closed;
                        CloseConnection(_internalsqlConn);
                    }

                    if (OpenConnection(_internalsqlConn))
                    {
                        DataTable elsag = Select(_internalCmd, "spCheckDB", "@Nome=ELSAG", 5, false);

                        if (elsag != null && elsag.Rows.Count > 0 && elsag.Rows[0]["Stato"].Equals(0))
                            _statoDB[NomiDB.ELSAG] = ConnectionState.Open;
                        else
                            _statoDB[NomiDB.ELSAG] = ConnectionState.Closed;
                        CloseConnection(_internalsqlConn);
                    }
                }
                _chechDBCount++;
            }
            else
            {
                _chechDBCount = 4;
                _statoDB[NomiDB.IMP] = ConnectionState.Closed;
                _statoDB[NomiDB.ELSAG] = ConnectionState.Closed;
            }

            if (_statoDB[NomiDB.SQLSERVER] != oldStatoDB[NomiDB.SQLSERVER]
                || _statoDB[NomiDB.IMP] != oldStatoDB[NomiDB.IMP]
                || _statoDB[NomiDB.ELSAG] != oldStatoDB[NomiDB.ELSAG])
            {
                NotifyPropertyChanged("StatoDB");
            }
            
            CloseConnection(_internalsqlConn);            
        }
        /// <summary>
        /// Metodo di notifica di un cambio di valore della proprietà propertyName.
        /// </summary>
        /// <param name="propertyName">Nome della proprietà di cui notificare il cambiamento.</param>
        private void NotifyPropertyChanged(String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }

        #endregion        
    }
}
