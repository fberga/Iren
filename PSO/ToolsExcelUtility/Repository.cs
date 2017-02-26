using Iren.ToolsExcel.Base;
using Iren.ToolsExcel.Core;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace Iren.ToolsExcel.Utility
{
    public class Repository
    {
        #region Variabili

        private IToolsExcelThisWorkbook _wb;
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
        public bool DaAggiornare { get; set; }
        public DataSet DataSet { get { return _wb.RepositoryDataSet; } }

        #endregion

        #region Costruttore

        public Repository(IToolsExcelThisWorkbook wb)
        {
            _wb = wb;
            DaAggiornare = false;

            if (Contains(DataBase.TAB.LISTA_APPLICAZIONI))
                Applicazione = this[DataBase.TAB.LISTA_APPLICAZIONI].AsEnumerable()
                    .Where(r => r["IdApplicazione"].Equals(wb.IdApplicazione))
                    .FirstOrDefault();
        }

        #endregion

        #region Metodi

        public void Aggiorna()
        {
            //_isMultiApplication = appIDs != null;
            //_appIDs = appIDs;

            InitStrutturaNomi();
            CaricaApplicazioni();
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
            CaricaApplicazioneRibbon();
            CaricaAzioneCategoria();
            CaricaCategoriaEntita();
            CaricaEntitaAzione();
            CaricaEntitaAzioneCalcolo();
            CaricaEntitaInformazione();
            CaricaEntitaAzioneInformazione();
            CaricaCalcolo();
            CaricaCalcoloInformazione();
            CaricaEntitaCalcolo();
            CaricaEntitaGrafico();
            CaricaEntitaGraficoInformazione();
            CaricaEntitaCommitment();
            CaricaEntitaRampa();
            CaricaEntitaAssetto();
            CaricaEntitaProprieta();
            CaricaEntitaInformazioneFormattazione();
            CaricaEntitaParametroD();
            CaricaEntitaParametroH();
            CaricaStagioni();

            //_wb.RepositoryDataSet.AcceptChanges();
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
        private DataTable CreaTabellaModifica(string name)
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
        private void CaricaDati(string tableName, string spName, QryParams parameters)
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
            QryParams parameters = new QryParams() 
                {
                    {"@IdApplicazione", 0}
                };

            CaricaDati(DataBase.TAB.LISTA_APPLICAZIONI, DataBase.SP.APPLICAZIONE, parameters);
        }
        /// <summary>
        /// Carica i dati necessari alla creazione del menu ribbon.
        /// </summary>
        /// <returns></returns>
        private void CaricaApplicazioneRibbon()
        {
            CaricaDati(DataBase.TAB.APPLICAZIONE_RIBBON, DataBase.SP.APPLICAZIONE_RIBBON, new QryParams());
        }
        /// <summary>
        /// Carica le azioni.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioni()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL},
                    {"@Visibile", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.AZIONE, DataBase.SP.AZIONE, parameters);
        }
        /// <summary>
        /// Carica le categorie.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategorie()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@Operativa", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CATEGORIA, DataBase.SP.CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione azione categoria.
        /// </summary>
        /// <returns></returns>
        private void CaricaAzioneCategoria()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCategoria", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.AZIONE_CATEGORIA, DataBase.SP.AZIONE_CATEGORIA, parameters);
        }
        /// <summary>
        /// Carica la relazione categoria entita.
        /// </summary>
        /// <returns></returns>
        private void CaricaCategoriaEntita()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCategoria", Core.DataBase.ALL},
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CATEGORIA_ENTITA, DataBase.SP.CATEGORIA_ENTITA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE, DataBase.SP.ENTITA_AZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione calcolo.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE_CALCOLO, DataBase.SP.ENTITA_AZIONE_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_INFORMAZIONE, DataBase.SP.ENTITA_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità azione informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAzioneInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaAzione", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_AZIONE_INFORMAZIONE, DataBase.SP.ENTITA_AZIONE_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica i calcoli.
        /// </summary>
        /// <returns></returns>
        private void CaricaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
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
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaCalcolo", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.CALCOLO_INFORMAZIONE, DataBase.SP.CALCOLO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità calcolo.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaCalcolo()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaCalcolo", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_CALCOLO, DataBase.SP.ENTITA_CALCOLO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGrafico()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_GRAFICO, DataBase.SP.ENTITA_GRAFICO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità grafico informazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaGraficoInformazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaGrafico", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_GRAFICO_INFORMAZIONE, DataBase.SP.ENTITA_GRAFICO_INFORMAZIONE, parameters);
        }
        /// <summary>
        /// Carica la relazione entità commitment.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaCommitment()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_COMMITMENT, DataBase.SP.ENTITA_COMMITMENT, parameters);
        }
        /// <summary>
        /// Carica la relazione entità rampa.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaRampa()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_RAMPA, DataBase.SP.ENTITA_RAMPA, parameters);
        }
        /// <summary>
        /// Carica la relazione entità assetto.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaAssetto()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_ASSETTO, DataBase.SP.ENTITA_ASSETTO, parameters);
        }
        /// <summary>
        /// Carica la relazione entità proprietà.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaProprieta()
        {
            CaricaDati(DataBase.TAB.ENTITA_PROPRIETA, DataBase.SP.ENTITA_PROPRIETA, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità informazione formattazione.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaInformazioneFormattazione()
        {
            QryParams parameters = new QryParams() 
                {
                    {"@SiglaEntita", Core.DataBase.ALL},
                    {"@SiglaInformazione", Core.DataBase.ALL}
                };

            CaricaDati(DataBase.TAB.ENTITA_INFORMAZIONE_FORMATTAZIONE, DataBase.SP.ENTITA_INFORMAZIONE_FORMATTAZIONE, parameters);
        }
        /// <summary>
        /// Carica la tipologia check.
        /// </summary>
        /// <returns></returns>
        private void CaricaTipologiaCheck()
        {
            CaricaDati(DataBase.TAB.TIPOLOGIA_CHECK, DataBase.SP.TIPOLOGIA_CHECK, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro giornaliero.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaParametroD()
        {
            CaricaDati(DataBase.TAB.ENTITA_PARAMETRO_D, DataBase.SP.ENTITA_PARAMETRO_D, new QryParams());
        }
        /// <summary>
        /// Carica la relazione entità parametro orario.
        /// </summary>
        /// <returns></returns>
        private void CaricaEntitaParametroH()
        {
            CaricaDati(DataBase.TAB.ENTITA_PARAMETRO_H, DataBase.SP.ENTITA_PARAMETRO_H, new QryParams());
        }

        private void CaricaStagioni()
        {
            CaricaDati(DataBase.TAB.STAGIONE, DataBase.SP.STAGIONE, new QryParams());
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
                dt.Rows.Add(r);

            this[DataBase.TAB.MERCATI] = dt;
        }

        #endregion

        public DataRow CaricaApplicazione(object IdApplicazione)
        {
            CaricaApplicazioni();

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
