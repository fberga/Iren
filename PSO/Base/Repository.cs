using System.Data;
using System.Linq;

namespace Iren.PSO.Base
{
    public class Repository
    {
        #region Variabili

        private IPSOThisWorkbook _wb;
        private static bool _isMultiApplication = false;
        private static int[] _appIDs = null;

        #endregion

        #region Proprietà

        public DataTable this[string tableName]
        {
            get
            {
                if (_wb.RepositoryDataSet.Tables.Contains(tableName))
                    return _wb.RepositoryDataSet.Tables[tableName];

                return null;
            }
            private set
            {
                if (_wb.RepositoryDataSet.Tables.Contains(tableName))
                    _wb.RepositoryDataSet.Tables.Remove(tableName);
                if (value.TableName != tableName)
                    value.TableName = tableName;
                _wb.RepositoryDataSet.Tables.Add(value);
            }
        }
        public DataTable this[int index]
        {
            get
            {
                if (_wb.RepositoryDataSet.Tables.Count > index)
                    return _wb.RepositoryDataSet.Tables[index];

                return null;
            }
        }
        public DataRow Applicazione { get; private set; }
        public int TablesCount { get { return _wb.RepositoryDataSet.Tables.Count; } }
        public DataSet DataSet { get { return _wb.RepositoryDataSet; } }

        #endregion

        #region Costruttore

        public Repository(IPSOThisWorkbook wb)
        {
            _wb = wb;

            if (Contains(DataBase.TAB.LISTA_APPLICAZIONI))
                Applicazione = this[DataBase.TAB.LISTA_APPLICAZIONI].AsEnumerable()
                    .Where(r => r["IdApplicazione"].Equals(wb.IdApplicazione))
                    .FirstOrDefault();
        }

        #endregion

        #region Metodi

        public void Aggiorna()
        {
            SplashScreen.UpdateStatus("Aggiornamento repository interno");
            //_isMultiApplication = appIDs != null;
            //_appIDs = appIDs;

            InitStrutturaNomi();
            //CaricaApplicazioni();
            CaricaApplicazione(_wb.IdApplicazione);
            CaricaMercati();

            //decido se è necessario caricare più applicazioni o solo una
            _isMultiApplication = this[DataBase.TAB.MERCATI].AsEnumerable()
                .Select(r => (int)r["IdApplicazioneMercato"])
                .Contains(Workbook.IdApplicazione);

            _appIDs = this[DataBase.TAB.MERCATI].AsEnumerable()
                .Select(r => (int)r["IdApplicazioneMercato"])
                .ToArray();

            CaricaAzioni();
            CaricaCategorie();
            CaricaAzioneCategoria();
            CaricaCategoriaEntita();
            CaricaEntitaAzione();
            CaricaEntitaAzioneCalcolo();
            CaricaEntitaInformazione();
            CaricaEntitaAzioneInformazione();
            CaricaCalcolo();
            CaricaCalcoloInformazione();
            //CaricaEntitaCalcolo();
            CaricaEntitaGrafico();
            CaricaEntitaGraficoInformazione();
            CaricaEntitaCommitment();
            CaricaEntitaRampa();
            CaricaEntitaAssetto();
            CaricaEntitaProprieta();
            CaricaEntitaInformazioneFormattazione();
            CaricaParametri();
            CaricaStagioni();
            CaricaDefinzioneOfferteMI();

            //_wb.RepositoryDataSet.AcceptChanges();
        }

        public void CaricaParametri()
        {
            //CaricaEntitaParametroD();
            //CaricaEntitaParametroH();
            CaricaEntitaParametro();
        }
        
        #region Aggiorna Struttura Dati

        #region Init Struttura Nomi

        public void InitStrutturaNomi()
        {

            this[DataBase.TAB.NOMI_DEFINITI] = DefinedNames.GetDefaultNameTable(DataBase.TAB.NOMI_DEFINITI);
            this[DataBase.TAB.DATE_DEFINITE] = DefinedNames.GetDefaultDateTable(DataBase.TAB.DATE_DEFINITE);
            this[DataBase.TAB.ADDRESS_FROM] = DefinedNames.GetDefaultAddressFromTable(DataBase.TAB.ADDRESS_FROM);
            this[DataBase.TAB.ADDRESS_TO] = DefinedNames.GetDefaultAddressToTable(DataBase.TAB.ADDRESS_TO);
            this[DataBase.TAB.EDITABILI] = DefinedNames.GetDefaultEditableTable(DataBase.TAB.EDITABILI);
            this[DataBase.TAB.SALVADB] = DefinedNames.GetDefaultSaveTable(DataBase.TAB.SALVADB);
            this[DataBase.TAB.ANNOTA] = DefinedNames.GetDefaultToNoteTable(DataBase.TAB.ANNOTA);
            this[DataBase.TAB.CHECK] = DefinedNames.GetDefaultCheckTable(DataBase.TAB.CHECK);
            this[DataBase.TAB.SELECTION] = DefinedNames.GetDefaultSelectionTable(DataBase.TAB.SELECTION);
            this[DataBase.TAB.MODIFICA] = CreaTabellaModifica(DataBase.TAB.MODIFICA);
            this[DataBase.TAB.EXPORT_XML] = CreaTabellaExportXML(DataBase.TAB.EXPORT_XML);
        }
        public DataTable CreaTabellaModifica(string name)
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(string)},
                    {"AnnotaModifica", typeof(string)},
                    {"IdApplicazione", typeof(string)},
                    {"IdUtente", typeof(string)}
                }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
            return dt;
        }

        public DataTable CreaTabellaRipristinaIncremento(string name)
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(string)},
                    {"Commento", typeof(string)}
                }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
            return dt;
        }

        private DataTable CreaTabellaExportXML(string name)
        {
            DataTable dt = new DataTable(name)
            {
                Columns =
                {
                    {"SiglaEntita", typeof(string)},
                    {"SiglaInformazione", typeof(string)},
                    {"Data", typeof(string)},
                    {"Valore", typeof(string)},
                    {"AnnotaModifica", typeof(string)},
                    {"IdApplicazione", typeof(string)},
                    {"IdUtente", typeof(string)}
                }
            };

            dt.PrimaryKey = new DataColumn[] { dt.Columns["SiglaEntita"], dt.Columns["SiglaInformazione"], dt.Columns["Data"] };
            return dt;
        }

        #endregion

        /// <summary>
        /// Metodo richiamato da tutte le routine sottostanti che effettua la chiamata alla stored procedure sul server e aggiunge la tabella al DataSet locale. Restituisce true se l'operazione è andata a buon fine, lancia un'eccezione RepositoryUpdateException se fallisce.
        /// </summary>
        /// <param name="tableName">Nome della tabella da aggiornare.</param>
        /// <param name="spName">Nome della stored procedure da eseguire.</param>
        /// <param name="parameters">Parametri della stored procedure.</param>
        /// <returns>True se l'operazione è andata a buon fine.</returns>
        private void CaricaDati(string tableName, string spName, Core.QryParams parameters)
        {
            DataTable dt = new DataTable();
            if (_isMultiApplication)
            {
                foreach (int id in _appIDs)
                {
                    parameters["@IdApplicazione"] = id;
                    dt.Merge(DataBase.Select(spName, parameters) ?? new DataTable());
                }
                if (dt.Columns.Count == 0)
                    dt = null;
            }
            else
            {
                dt = DataBase.Select(spName, parameters);
            }

            if (dt != null)
                this[tableName] = dt;
        }
        /// <summary>
        /// Carica la lista di tutte le applicazioni disponibili.
        /// </summary>
        private void CaricaApplicazioni()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@IdApplicazione", 0}
                };

            CaricaDati(DataBase.TAB.LISTA_APPLICAZIONI, DataBase.SP.APPLICAZIONE, parameters);
        }
        /// <summary>
        /// Carica le azioni.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioni()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaAzione", PSO.Core.DataBase.ALL},
                    {"@Operativa", PSO.Core.DataBase.ALL},
                    {"@Visibile", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.AZIONE, DataBase.SP.AZIONE, parameters);
        }
        /// <summary>
        /// Carica le categorie.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategorie()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaCategoria", PSO.Core.DataBase.ALL},
                    {"@Operativa", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CATEGORIA, DataBase.SP.CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione azione categoria.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioneCategoria()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaAzione", PSO.Core.DataBase.ALL},
                    {"@SiglaCategoria", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.AZIONE_CATEGORIA, DataBase.SP.AZIONE_CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione categoria entita.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategoriaEntita()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaCategoria", PSO.Core.DataBase.ALL},
                    {"@SiglaEntita", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CATEGORIA_ENTITA, DataBase.SP.CATEGORIA_ENTITA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaAzione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE, DataBase.SP.ENTITA_AZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione calcolo.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneCalcolo()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaAzione", PSO.Core.DataBase.ALL},
                    {"@SiglaCalcolo", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE_CALCOLO, DataBase.SP.ENTITA_AZIONE_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaInformazione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_INFORMAZIONE, DataBase.SP.ENTITA_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneInformazione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaAzione", PSO.Core.DataBase.ALL},
                    {"@SiglaInformazione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE, DataBase.SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica i calcoli.
        /// </summary>
        /// <returns></returns>
        private void CaricaCalcolo()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaCalcolo", PSO.Core.DataBase.ALL},
                    {"@IdTipologiaCalcolo", 0}
                };

            CaricaDati(DataBase.TAB.CALCOLO, DataBase.SP.CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione calcolo informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaCalcoloInformazione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaCalcolo", PSO.Core.DataBase.ALL},
                    {"@SiglaInformazione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CALCOLO_INFORMAZIONE, DataBase.SP.CALCOLO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità calcolo.
        /// </summary>
        /// <returns></returns>
        //private void CaricaEntitaCalcolo()
        //{
        //    Core.QryParams parameters = new Core.QryParams() 
        //        {
        //            {"@SiglaEntita", PSO.Core.DataBase.ALL},
        //            {"@SiglaCalcolo", PSO.Core.DataBase.ALL}
        //        };

        //    CaricaDati(DataBase.TAB.ENTITA_CALCOLO, DataBase.SP.ENTITA_CALCOLO, parameters);
        //}
        /// <summary>
        /// Carica la relazione entità grafico.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGrafico()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaGrafico", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_GRAFICO, DataBase.SP.ENTITA_GRAFICO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGraficoInformazione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaGrafico", PSO.Core.DataBase.ALL},
                    {"@SiglaInformazione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_GRAFICO_INFORMAZIONE, DataBase.SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità commitment.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaCommitment()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_COMMITMENT, DataBase.SP.ENTITA_COMMITMENT, parameters);
        }
        /// <summary>
        /// Carica la relazione entità rampa.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaRampa()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_RAMPA, DataBase.SP.ENTITA_RAMPA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità assetto.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAssetto()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_ASSETTO, DataBase.SP.ENTITA_ASSETTO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità proprietà.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaProprieta()
        {
            CaricaDati(DataBase.TAB.ENTITA_PROPRIETA, DataBase.SP.ENTITA_PROPRIETA, new Core.QryParams());
        }
        /// <summary>
        /// Carica la relazione entità informazione formattazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazioneFormattazione()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@SiglaEntita", PSO.Core.DataBase.ALL},
                    {"@SiglaInformazione", PSO.Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_INFORMAZIONE_FORMATTAZIONE, DataBase.SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
        }
        /// <summary>
        /// Carica la tipologia check.
        /// </summary>
        /// <returns></returns>
        private void CaricaTipologiaCheck()
        {
            CaricaDati(DataBase.TAB.TIPOLOGIA_CHECK, DataBase.SP.TIPOLOGIA_CHECK, new Core.QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro giornaliero.
        /// </summary>
        /// <returns></returns>
        //private void CaricaEntitaParametroD()
        //{
        //    CaricaDati(DataBase.TAB.ENTITA_PARAMETRO_D, DataBase.SP.ENTITA_PARAMETRO_D, new Core.QryParams());
        //}
        /// <summary>
        /// Carica la relazione entità parametro orario.
        /// </summary>
        /// <returns></returns>
        //private void CaricaEntitaParametroH()
        //{
        //    CaricaDati(DataBase.TAB.ENTITA_PARAMETRO_H, DataBase.SP.ENTITA_PARAMETRO_H, new Core.QryParams());
        //}

        private void CaricaEntitaParametro()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@Data", Workbook.DataAttiva.ToString("yyyyMMdd")}
                };

            CaricaDati(DataBase.TAB.ENTITA_PARAMETRO, DataBase.SP.ENTITA_PARAMETRO, parameters);
        }

        private void CaricaStagioni()
        {
            CaricaDati(DataBase.TAB.STAGIONE, DataBase.SP.STAGIONE, new Core.QryParams());
        }

        private void CaricaMercati()
        {
            DataTable dt = new DataTable(DataBase.TAB.MERCATI)
            {
                Columns = 
                {
                    {"IdApplicazioneMercato", typeof(int)},
                    {"DesMercato", typeof(string)}
                }
            };

            var mercati = this[DataBase.TAB.LISTA_APPLICAZIONI].AsEnumerable()
                .Where(r => r["SiglaApplicazione"].ToString().StartsWith("INVIA_PROGRAMMA_"))
                .OrderBy(r => r["SiglaApplicazione"].ToString().Replace("INVIA_PROGRAMMA_", ""))
                .Select(r =>
                {
                    DataRow o = dt.NewRow();
                    o["IdApplicazioneMercato"] = r["IdApplicazione"];
                    o["DesMercato"] = r["SiglaApplicazione"].ToString().Replace("INVIA_PROGRAMMA_", "");

                    return o;
                });

            foreach(DataRow r in mercati)
            {            
                // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari  ***** BEGIN *****
                //TODO scommentare per passare in produzione prima del 01/02
                // if (!r["DesMercato"].Equals("MSD5") && !r["DesMercato"].Equals("MSD6"))
                // Modifica per InvioProgrammi MSD5 e MSD6 e nuovi orari ***** END ****
                    dt.Rows.Add(r);
            }
            this[DataBase.TAB.MERCATI] = dt;
        }

        private void CaricaDefinzioneOfferteMI()
        {
            Core.QryParams parameters = new Core.QryParams() 
                {
                    {"@IdMercato", 0},
                    {"@SiglaEntita", "all"},
                    {"@SiglaInformazione", "all"},
                };

            CaricaDati(DataBase.TAB.DEFINIZIONE_OFFERTA, DataBase.SP.DEFINIZIONE_OFFERTA, parameters);
        }

        #endregion

        public DataRow CaricaApplicazione(object IdApplicazione)
        {
            CaricaApplicazioni();

            return CambiaApplicazione(IdApplicazione);
        }
        public DataRow CambiaApplicazione(object IdApplicazione)
        {
            Applicazione = this[DataBase.TAB.LISTA_APPLICAZIONI].AsEnumerable()
                .Where(r => r["IdApplicazione"].Equals(IdApplicazione))
                .FirstOrDefault();

            return Applicazione;
        }

        public void Add(DataTable table)
        {
            _wb.RepositoryDataSet.Tables.Add(table);
        }

        public bool Contains(string name)
        {
            return _wb.RepositoryDataSet.Tables.Contains(name);
        }

        public void Remove(string name)
        {
            _wb.RepositoryDataSet.Tables.Remove(name);
        }

        #endregion
    }
}
